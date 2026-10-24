"""Backup management: list / restore / purge for the slot-based backup system
(core.safesave) plus DTG-stamped permanent snapshots.

There is no legacy ``*.bak-*`` handling: kitchensink4ppt never shipped a
per-mutation backup scheme, so there is nothing legacy to manage.

Restore discipline: the presentation's current content is staged for prev
first and lands there once the restore went through, so a restore is itself
undoable via prev, and a refused restore leaves prev as it was. A restore
that would change nothing says so instead of answering ok. Restores refuse
presentations that are open in PowerPoint (owner-file plus exclusive-lock
detection, the same probe the editing path uses) and validate the backup
payload before touching the target; the target is replaced in one swap.

The public ``create_snapshot`` behavior remains a visible permanent copy next
to the deck (or in ``dest_dir``).  Safety snapshots created automatically by
anonymization use its private ``_managed`` mode instead: they live inside the
deck's hidden managed-backup folder, appear in ``manage_backups`` listings,
and can be passed back as the ``source`` of a restore.
"""

from __future__ import annotations

import datetime as _dt
import os
import re
import shutil
import uuid
import zipfile
import zlib
from pathlib import Path

from lxml import etree

from ..core import safesave, xmlsafe, zipsafe
from ..core.errors import (
    DocumentCorrupt,
    DocumentLocked,
    DocumentNotFound,
    PptMcpError,
)
from ..core.package import (
    PRESENTATION_PART,
    PptxPackage,
    _is_ole_encrypted,
    has_read_only_attribute,
    kept_problems_note,
    known_problem_keys,
    read_only_error,
)
from ..core.safesave import (
    ANCHOR_SLOT,
    BACKUP_DIR_NAME,
    PREV_SLOT,
    SLOT_POLICY,
    write_lock,
)
from ..core.sandbox import check_path
from ..core.swapwait import asides, exists_after_swap, require_deck, through_swap


def _refuse_if_app_locked(path: Path) -> None:
    """Same detection PptxPackage uses: PowerPoint holds an exclusive lock."""
    owner_file = path.with_name("~$" + path.name[-153:])
    try:
        with open(path, "r+b"):
            pass
    except PermissionError:
        hint = " (PowerPoint owner file present)" if owner_file.exists() else ""
        raise DocumentLocked(
            f"{path.name} is open in PowerPoint or locked by another "
            f"process{hint}. Close it before restoring a backup over it."
        ) from None


def _stat_entry(p: Path) -> dict:
    st = p.stat()
    return {
        "path": str(p),
        "size_bytes": st.st_size,
        "modified": _dt.datetime.fromtimestamp(st.st_mtime).isoformat(
            timespec="seconds"
        ),
    }


_MANAGED_SNAPSHOT_PREFIX = "snapshot-"


def _managed_snapshots(folder: Path) -> list[dict]:
    """List permanent snapshots owned by one managed backup folder."""
    if not folder.is_dir():
        return []
    return [
        {"source": str(path), **_stat_entry(path)}
        for path in sorted(folder.glob(f"{_MANAGED_SNAPSHOT_PREFIX}*"))
        if path.is_file()
    ]


def _orphan_entry(folder: Path) -> dict:
    """One orphaned folder as the listing shows it. Its snapshots are named,
    because once the deck is gone they may be the only copy of it left, and
    a purge of orphans keeps them."""
    entry = {
        "folder": str(folder),
        "size_bytes": _dir_size(folder),
        "missing_presentation": str(safesave.source_doc_for(folder)),
    }
    snapshots = _managed_snapshots(folder)
    if snapshots:
        entry["snapshots"] = snapshots
    return entry


def _deck_is_gone(doc: Path) -> bool:
    """The deck is not at its path, is not being swapped in by a save, and
    no save that did not finish has it aside: its backups belong to no
    deck. A deck that is aside keeps them; they may be all there is."""
    return not exists_after_swap(doc) and not asides(doc)


def _orphan_folders(directory: Path) -> list[Path]:
    """Slot folders under directory/.ks4p-backups whose source file is gone."""
    root = directory / BACKUP_DIR_NAME
    if not root.is_dir():
        return []
    orphans = []
    for folder in sorted(p for p in root.iterdir() if p.is_dir()):
        if _deck_is_gone(safesave.source_doc_for(folder)):
            orphans.append(folder)
    return orphans


def _dir_size(folder: Path) -> int:
    return sum(p.stat().st_size for p in folder.rglob("*") if p.is_file())


# ---------------------------------------------------------------------- list


def _slot_entry(slot_file: str, path: Path) -> dict:
    """One slot as the listing shows it. `source` is the value a restore
    takes for it, the same key every listed snapshot carries."""
    name = slot_file.split(".")[0]
    return {"slot": name, "source": name, **_stat_entry(path)}


def list_backups(
    file_path: str | None = None, directory: str | None = None
) -> dict:
    if file_path is None and directory is None:
        raise PptMcpError("provide file_path (one presentation) or directory")
    if file_path is not None:
        check_path(file_path, "list backups")
        doc = Path(file_path).resolve()
        d = safesave.slot_dir(doc)
        slots = []
        for slot in SLOT_POLICY:
            sp = d / slot
            if sp.exists():
                slots.append(_slot_entry(slot, sp))
        return {
            "presentation": str(doc),
            "presentation_exists": doc.exists(),
            "slots": slots,
            "snapshots": _managed_snapshots(d),
            "orphaned_folders": [
                _orphan_entry(f) for f in _orphan_folders(doc.parent)
            ],
        }

    check_path(directory, "list backups")
    base = Path(directory).resolve()
    if not base.is_dir():
        raise DocumentNotFound(f"no directory at {base}")
    root = base / BACKUP_DIR_NAME
    presentations = []
    if root.is_dir():
        for folder in sorted(p for p in root.iterdir() if p.is_dir()):
            src = safesave.source_doc_for(folder)
            if _deck_is_gone(src):
                continue  # reported under orphaned_folders below
            slots = []
            for slot in SLOT_POLICY:
                sp = folder / slot
                if sp.exists():
                    slots.append(_slot_entry(slot, sp))
            presentations.append(
                {
                    "presentation": str(src),
                    "slots": slots,
                    "snapshots": _managed_snapshots(folder),
                }
            )
    return {
        "directory": str(base),
        "presentations": presentations,
        "orphaned_folders": [_orphan_entry(f) for f in _orphan_folders(base)],
    }


# ------------------------------------------------------------------- restore


def restore_backup(file_path: str, source: str) -> dict:
    """Replace the presentation's content with a backup. source: 'prev',
    'anchor', or a path to a .pptx file (e.g. a snapshot). The current
    content becomes prev once the restore lands, so the restore itself can
    be undone by restoring prev again."""
    check_path(file_path, "restore backup over presentation")
    doc = Path(file_path).resolve()
    if isinstance(source, str) and source.strip().lower() in ("prev", "anchor"):
        source = source.strip().lower()  # a slot's name, however it is cased
    # The restore target is an edit destination.  A target beneath the
    # hidden backup store would overwrite the very recovery data this
    # operation is meant to protect (and then create a nested store).
    safesave.refuse_protected_output(doc, purpose="restore the presentation")
    if source in ("prev", "anchor"):
        src = safesave.slot_dir(doc) / (
            PREV_SLOT if source == "prev" else ANCHOR_SLOT
        )
        label = source
    else:
        check_path(source, "read backup file")
        src = Path(source).resolve()
        label = f"file {src.name}"
        exists_after_swap(src)  # a deck given as the source may be mid-save
    if not src.is_file():
        raise DocumentNotFound(
            f"no backup to restore: {src} does not exist. "
            "Use manage_backups action='list' to see what is available."
        )

    # Refuse before write_lock creates the backup sidecar directory, and do
    # not misdiagnose a file attribute as another process's lock.
    if doc.exists() and has_read_only_attribute(doc):
        raise read_only_error(doc)

    with write_lock(doc):
        target_existed = doc.exists()
        if target_existed:
            _refuse_if_app_locked(doc)

        # Validate the backup payload BEFORE touching anything.
        payload = src.read_bytes()
        # A problem the deck being replaced already has is kept, as a save
        # keeps it; anything else in the backup is refused.
        kept = PptxPackage._validate_payload(
            payload, tolerate=known_problem_keys(doc),
            doc=f"the {source} backup of {doc.name}" if source in ("prev", "anchor") else src.name)
        if target_existed and _holds(doc, payload):
            raise _nothing_to_restore(doc, label)

        # Copy (not hardlink) the source to a temp beside the target first:
        # rotating prev below may clobber the very slot being restored from,
        # and the restored target must own its bytes outright.
        d = safesave.slot_dir(doc, create=True)
        tmp = d / f".restore-{uuid.uuid4().hex}.tmp"
        ticket = None
        try:
            shutil.copy2(src, tmp)
            safesave.claim_publish("the backup restore")
            if target_existed:
                ticket = safesave.prepare_rotation(doc, anchor=False)
            restore_warnings = safesave.promote(tmp, doc, ticket=ticket)
        except BaseException as exc:
            if ticket is not None:
                ticket.abort()
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
            if isinstance(exc, OSError):
                refusal = safesave.os_write_refusal(doc, exc)
                if refusal is not None:
                    raise refusal from exc
            raise

    rotated_prev = ticket is not None and ticket.rotated["prev"]
    result = {
        "restored": str(doc),
        "from": label,
        "bytes": len(payload),
        "prev_rotated": rotated_prev,
    }
    warnings = list(restore_warnings or [])
    if kept:
        warnings.append(kept_problems_note(kept))
    if warnings:
        result["warnings"] = warnings
    if rotated_prev:
        result["undo"] = (
            "restore source='prev' brings back the pre-restore content"
        )
    elif not target_existed:
        result["note"] = "presentation did not exist; nothing rotated into prev"
    return result


def _holds(doc: Path, payload: bytes) -> bool:
    """Whether the deck already holds exactly these bytes."""
    try:
        return doc.stat().st_size == len(payload) and doc.read_bytes() == payload
    except OSError:
        return False


def _nothing_to_restore(doc: Path, label: str) -> PptMcpError:
    exc = PptMcpError(
        f"Nothing to restore: {doc.name} already holds exactly what {label} "
        "holds, so restoring it would change nothing. Nothing was changed.")
    exc.code = "CONFLICT"
    exc.hint = ("manage_backups action='list' shows the other backups (prev, "
                "anchor and snapshots) and when each was made.")
    return exc


# --------------------------------------------------------------------- purge


def _collect_purge_targets(
    scope: str, file_path: str | None, directory: str | None
) -> tuple[list[Path], Path | None, list[dict]]:
    """Returns (targets, slot_folder_for_slots_scope, kept). An orphaned
    folder that holds snapshots is kept whole and listed in kept: a snapshot
    is permanent, and once its deck is gone it may be the only copy left."""
    if file_path is not None:
        check_path(file_path, "purge backups")
    if directory is not None:
        check_path(directory, "purge backups")
    if scope == "slots":
        if not file_path:
            raise PptMcpError(
                "scope='slots' needs file_path (whose slots to purge)"
            )
        doc = Path(file_path).resolve()
        d = safesave.slot_dir(doc)
        targets = [d / slot for slot in SLOT_POLICY if (d / slot).exists()]
        return targets, d, []
    if scope == "orphans":
        base = (
            Path(file_path).resolve().parent
            if file_path
            else Path(directory).resolve()
            if directory
            else None
        )
        if base is None:
            raise PptMcpError("scope='orphans' needs directory (or file_path)")
        targets, kept = [], []
        for folder in _orphan_folders(base):
            snapshots = _managed_snapshots(folder)
            if snapshots:
                kept.append({
                    "folder": str(folder),
                    "snapshots": [item["source"] for item in snapshots],
                    "reason": (
                        "it holds snapshots, which are permanent: they may be "
                        "the only copy of a deck that was moved, renamed or "
                        "deleted. Restore one with manage_backups "
                        "action='restore', or delete the folder yourself."),
                })
            else:
                targets.append(folder)
        return targets, None, kept
    raise PptMcpError(
        f"unknown purge scope {scope!r}; use 'orphans' or 'slots'"
    )


def purge_backups(
    scope: str,
    file_path: str | None = None,
    directory: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Delete backups. scope: 'orphans' (slot folders whose source
    presentation is gone) or 'slots' (one presentation's prev/anchor).
    dry_run=True (the default) only reports what WOULD be deleted; pass
    dry_run=False to actually delete."""
    targets, slot_folder, kept = _collect_purge_targets(
        scope, file_path, directory)
    report = []
    for t in targets:
        size = _dir_size(t) if t.is_dir() else t.stat().st_size
        report.append({"path": str(t), "size_bytes": size})
    total = sum(e["size_bytes"] for e in report)
    result = {
        "scope": scope,
        "dry_run": dry_run,
        ("would_delete" if dry_run else "deleted"): report,
        "total_bytes": total,
        "count": len(report),
    }
    if kept:
        result["kept"] = kept
    if dry_run:
        result["note"] = "nothing was deleted; pass dry_run=False to delete"
        return result

    cleans_slot_folder = (
        scope == "slots" and slot_folder is not None and slot_folder.is_dir()
    )
    if targets or cleans_slot_folder:
        safesave.claim_publish("the backup purge")
    for t in targets:
        if t.is_dir():
            shutil.rmtree(t, ignore_errors=False)
        else:
            t.unlink()
    # After purging a presentation's slots, drop its folder (and the
    # breadcrumb, if any) when nothing else is left in it; best-effort, a
    # leftover lockfile just stays. A folder that still holds snapshots
    # keeps its breadcrumb: without it a long-named deck's folder no longer
    # maps back to the deck and would read as an orphan.
    if scope == "slots" and slot_folder is not None and slot_folder.is_dir():
        crumb = slot_folder / safesave._SOURCE_NAME_FILE
        try:
            others = [entry for entry in slot_folder.iterdir() if entry != crumb]
        except OSError:
            others = [None]
        if not others:
            try:
                crumb.unlink(missing_ok=True)
                os.rmdir(slot_folder)
            except OSError:
                pass
    return result


# ----------------------------------------------------------------- snapshots

_DTG_PREFIX = re.compile(r"^\d{8}_\d{4}_")
_LABEL_BAD = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


#: The two parts every snapshot must be able to read, with the root element
#: each one has to carry.
_SNAPSHOT_XML = (("[Content_Types].xml", "Types"), (PRESENTATION_PART, "presentation"))
#: Most part data the snapshot check decompresses to verify checksums. The two
#: required parts are checked first and are never skipped: one that claims to
#: be larger than this is refused. In a package larger than this (far past any
#: ordinary deck) the remaining parts are copied unchecked rather than read.
SNAPSHOT_CHECK_BYTES = 1 << 30
_CUT_OFF = "is not a complete PowerPoint file (it may be cut off or damaged)"


def _ole_problem(doc: Path, size: int) -> str | None:
    """A password-protected deck is an OLE compound file, which this server
    cannot open, so only its frame is checked: a 512-byte header with the
    format's byte-order mark and sector size, and at least one sector."""
    with open(doc, "rb") as fh:
        header = fh.read(512)
    if len(header) < 512 or header[28:30] != b"\xfe\xff":
        return _CUT_OFF
    shift = int.from_bytes(header[30:32], "little")
    if shift not in (9, 12) or size < 512 + (1 << shift):
        return _CUT_OFF
    return None


def _snapshot_problem(doc: Path) -> str | None:
    """Why ``doc`` is not worth a permanent snapshot, or None when it is.

    A snapshot is a restore point, so it has to hold a deck: an empty file,
    a cut-off package or one whose contents are damaged would sit beside the
    real one looking like a safe copy. The checks are the ones a reader can
    make without opening the deck in PowerPoint: the zip directory, every
    part's checksum (up to SNAPSHOT_CHECK_BYTES of part data, the two
    required parts always), and that the content types and
    ppt/presentation.xml are readable XML of the right kind. A deck that
    passes can still hold a problem only PowerPoint would find. A
    password-protected deck (an OLE file) is kept as it is once its frame is
    whole; this server cannot look inside it.
    """
    size = doc.stat().st_size
    if size == 0:
        return "is empty (0 bytes)"
    with open(doc, "rb") as fh:
        head = fh.read(8)
    if _is_ole_encrypted(head):
        return _ole_problem(doc, size)
    try:
        with zipfile.ZipFile(doc) as package:
            names = set(package.namelist())
            if any(name not in names for name, _root in _SNAPSHOT_XML):
                return "holds no PowerPoint presentation"
            first = [name for name, _root in _SNAPSHOT_XML]
            infos = sorted(
                package.infolist(), key=lambda info: info.filename not in first
            )
            budget = SNAPSHOT_CHECK_BYTES
            for info in infos:
                if info.is_dir():
                    continue
                if info.file_size > budget or (
                        info.filename in first and zipsafe.too_large_to_read(info)):
                    if info.filename in first:
                        # A required part is never skipped, and one this large
                        # is no presentation part: the recorded size of a
                        # damaged central directory says so.
                        return (
                            f"has a damaged {info.filename} (its recorded size "
                            "is beyond any real deck)"
                        )
                    continue
                budget -= info.file_size
                with package.open(info) as part:  # the checksum is verified at EOF
                    data = part.read() if info.filename in first else None
                    while data is None and part.read(1 << 20):
                        pass
                if data is not None:
                    if xmlsafe.declares_dtd(data):
                        return (f"has a document type declaration (<!DOCTYPE ...>) in "
                                f"{info.filename}, which a PowerPoint presentation never has")
                    root = _root_name(data)
                    wanted = dict(_SNAPSHOT_XML)[info.filename]
                    if root is None:
                        return f"has a damaged {info.filename} (it is not readable XML)"
                    if root != wanted:
                        return "holds no PowerPoint presentation"
    except (zipfile.BadZipFile, zlib.error, EOFError, OSError, ValueError,
            NotImplementedError, RuntimeError):
        # RuntimeError: zipfile's answer to an encrypted part, which no
        # PowerPoint package carries.
        return _CUT_OFF
    return None


def _root_name(data: bytes) -> str | None:
    """The local name of an XML part's root element, or None when the part
    is not well-formed XML. Parsed with the hardened parser."""
    try:
        return etree.QName(xmlsafe.fromstring(data)).localname
    except (etree.XMLSyntaxError, ValueError, DocumentCorrupt):
        return None


def create_snapshot(
    file_path: str,
    *,
    label: str | None = None,
    dest_dir: str | None = None,
    _managed: bool = False,
) -> dict:
    """A permanent copy of the deck as it is now.

    By default the copy goes into the deck's HIDDEN backup folder
    (.ks4p-backups/<deck name>/ beside the deck), named
    snapshot-YYYYMMDD_HHMM_<name>[_<label>].pptx. It never sits beside the
    deck, so it can never be taken for a newer edition of it where the
    newest leading date stamp marks the current version (the old default
    put DTG-led copies there, and they outranked the deck). manage_backups
    lists it and restores from it. Nothing removes it: slot rotation
    touches only prev and anchor, a purge of slots only those two files,
    and a purge of orphaned folders keeps every folder holding a snapshot.

    With dest_dir the copy is a VISIBLE file in that folder, named after the
    deck: "<deck name> (snapshot[ <label>] YYYY-MM-DD HHMM).pptx". The deck's
    whole name comes first, a leading date stamp included, so the copy never
    carries a newer leading date stamp than the deck. (In an ordinal listing
    it sorts just before the deck, because a space sorts before a dot, so
    max() over the names still picks the deck; a name that sorted after the
    deck would make max() pick the copy.)

    The time in the name and the copy's modification time are both the time
    the copy was written (a plain copy kept the deck's last save time, so a
    file named 22:19 said 22:15). Never overwrites; a collision gets " (2)",
    " (3)". The answer gives the full path and says whether it is hidden.
    Copies earlier builds left beside a deck are not touched.

    ``_managed`` is the internal mode of the automatic safety snapshot taken
    before anonymization: always hidden, and refusing dest_dir.
    """
    check_path(file_path, "snapshot presentation")
    doc = Path(file_path).resolve()
    require_deck(doc, f"no presentation at {doc}")
    problem = through_swap(doc, lambda: _snapshot_problem(doc))
    if problem:
        raise DocumentCorrupt(
            f"{doc.name} {problem}, so a snapshot of it would keep nothing you "
            "could restore; no snapshot was written. copy_presentation copies "
            "the bytes as they are if you want them anyway."
        )
    if label is not None:
        label = label.strip()
        if not label:
            label = None
        elif _LABEL_BAD.search(label):
            raise PptMcpError(
                "label contains characters not allowed in filenames "
                '(< > : " / \\ | ? * or control characters)'
            )
        elif len(label) > 60:
            raise PptMcpError("label must be 60 characters or fewer")

    if _managed and dest_dir is not None:
        raise PptMcpError("managed snapshots choose their own backup location")
    hidden = _managed or not dest_dir
    if hidden:
        target_dir = safesave.slot_dir(doc, create=True)
    else:
        check_path(dest_dir, "write snapshot")
        target_dir = Path(dest_dir).resolve()
    if not target_dir.is_dir():
        raise DocumentNotFound(f"no directory at {target_dir}")

    # One reading of the clock names the copy and dates it.
    written = _dt.datetime.now()
    if hidden:
        stem = _DTG_PREFIX.sub("", doc.stem)
        base = (_MANAGED_SNAPSHOT_PREFIX + written.strftime("%Y%m%d_%H%M")
                + f"_{stem}" + (f"_{label}" if label else ""))
    else:
        part = f"snapshot {label}" if label else "snapshot"
        base = f"{doc.stem} ({part} {written.strftime('%Y-%m-%d %H%M')})"
    dest = target_dir / f"{base}{doc.suffix}"
    n = 2
    while dest.exists():
        dest = target_dir / f"{base} ({n}){doc.suffix}"
        n += 1

    if not hidden:
        safesave.refuse_protected_output(
            dest, doc, purpose="write the presentation snapshot"
        )
    safesave.claim_publish("the presentation snapshot")
    through_swap(doc, lambda: shutil.copyfile(doc, dest))
    stamp = written.timestamp()
    try:
        os.utime(dest, (stamp, stamp))
    except OSError:
        pass  # the copy is whole; only its date could not be set
    result = {
        "snapshot": str(dest),
        "source": str(doc),
        "label": label,
        "hidden": hidden,
        "location": (
            "In the deck's hidden backup folder (.ks4p-backups beside the "
            "deck), not in the deck's own folder. manage_backups "
            "action='list' lists it; action='restore' with this path as "
            "source brings it back."
            if hidden else
            "A visible file in dest_dir, named after the deck: the deck's "
            "whole name comes first, so the copy never carries a newer "
            "leading date stamp than the deck."
        ),
        "written_at": written.isoformat(timespec="seconds"),
    }
    if hidden:
        result["managed"] = True
    owner_file = doc.with_name("~$" + doc.name[-153:])
    if owner_file.exists():
        result["note"] = (
            "the presentation appears to be open in PowerPoint; unsaved "
            "changes are NOT in this snapshot (save there first for a "
            "current copy)"
        )
    return result


# ------------------------------------------------------------------ dispatch


def manage_backups(
    action: str,
    file_path: str | None = None,
    directory: str | None = None,
    source: str | None = None,
    scope: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Single entry point the server tool exposes. action: list|restore|purge."""
    if action == "list":
        return list_backups(file_path=file_path, directory=directory)
    if action == "restore":
        if not file_path:
            raise PptMcpError(
                "restore needs file_path (the presentation to restore)"
            )
        if not source:
            raise PptMcpError(
                "restore needs source: 'prev', 'anchor', or a snapshot's .pptx "
                "path, as manage_backups action='list' gives it in each "
                "backup's source"
            )
        return restore_backup(file_path, source)
    if action == "purge":
        if not scope:
            raise PptMcpError("purge needs scope: 'orphans' or 'slots'")
        return purge_backups(
            scope, file_path=file_path, directory=directory, dry_run=dry_run
        )
    raise PptMcpError(
        f"unknown action {action!r}; use 'list', 'restore', or 'purge'"
    )
