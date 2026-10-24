"""Slot-based backup rotation + per-file write serialization for the save path.

Backup model:

- A hidden ``.ks4p-backups/`` folder sits next to each mutated presentation
  (dot-prefix everywhere; FILE_ATTRIBUTE_HIDDEN is additionally set on
  Windows, non-fatal if that fails).
- Inside it, one subfolder per presentation (the file's own name,
  hash-suffixed when the name is very long; Korean/unicode names work).
- Exactly two stable slots per presentation (SLOT_POLICY below):
    ``prev.pptx``   - state before the most recent mutation, rotates every call
    ``anchor.pptx`` - session-start state, rotates only when the file has
                      been idle ANCHOR_IDLE_SECONDS or more, measured from the
                      prev slot's mtime (no state database).
- Rotation mechanism (crash-window-free): HARDLINK the current target file to
  a temp name inside the slot folder, then os.replace that link onto the slot.
  The presentation is NEVER absent from its own path during the rotation; the
  old content survives the caller's final swap under the slot link.
  Hardlink failure (non-NTFS, cloud placeholders, cross-volume) falls back to
  shutil.copy2. Link/replace calls are retried with backoff to ride out
  transient AV sharing violations.
- The final swap (``promote``) puts the new version in place with ReplaceFileW
  so the deck keeps its identity (Mark of the Web, permissions, attributes,
  creation time). ReplaceFileW moves the old file aside for an instant before
  the new one takes its place, so the path is missing for that instant; the
  readers in this server wait it out (core/swapwait.py).

Write serialization (fixes the parallel read-modify-save race):

- ``write_lock(path)`` context manager, held across the FULL
  read-modify-validate-save cycle of a mutation.
- In-process: one threading.RLock per file, keyed on
  normcase(realpath(path)).
- Cross-process: an advisory lockfile inside the presentation's slot folder
  carrying PID, a per-process-instance token, the holder's process creation
  time, and a timestamp. Locks with dead or recycled local PIDs are broken;
  live or named remote holders retain ownership. Acquisition waits up to
  LOCK_WAIT_SECONDS and then refuses with MutationLockTimeout naming the
  holder. Two server processes on one machine is the normal case.

  The lockfile is PUBLISHED ATOMICALLY (payload written to a private temp
  file, then hardlinked into place) and RELEASED ONLY WHILE STILL OURS.
  Both rules are load-bearing. The original two-syscall design failed
  under concurrency: creating the
  file with O_CREAT|O_EXCL and writing the payload as a second syscall left
  a window in which the lockfile existed but was EMPTY, and a concurrent
  acquirer reading it there parsed no pid, declared a LIVE lock stale,
  deleted it and took its own. Two holders then unlinked each other's files
  and the lock outlived every writer (reproduced 4/4 at four concurrent
  writers; see tests/unit/test_safesave_lock_contention.py). The LIVE route
  already carries this pattern in com/xproc.py; the routes now agree.

  A release that cannot delete the file never leaves it behind
  (core/lockrelease.py): a waiter polling the lockfile holds it open for a
  moment, and on Windows that makes the holder's unlink fail. The release
  retries, then hands a lockfile that is still ours to a sweeper thread, and
  the outermost hold in this process adopts a lockfile carrying our own
  token instead of mistaking it for a nested hold and never releasing it.
"""

from __future__ import annotations

import errno
import hashlib
import json
import os
import re
import socket
import shutil
import sys
import tempfile
import threading
import time
import uuid
from contextlib import contextmanager
from pathlib import Path

from . import lockrelease
from .errors import (
    DeckSetAside,
    DocumentHeld,
    DocumentLocked,
    DocumentReadOnly,
    PptMcpError,
)
from .sandbox import check_path
from .swapwait import aside_in_place as replace_in_flight
from .swapwait import aside_prefix as _swap_prefix
from .swapwait import aside_name, asides, deck_spelling, stored_name
from .swapwait import journal_dir as _journal_dir

# ---------------------------------------------------------------- constants

BACKUP_DIR_NAME = ".ks4p-backups"

#: Slot policy: stable slot file names, at most these per presentation ever.
#: (Two slots by design; a third prior-anchor slot was considered and
#: declined.)
SLOT_POLICY: tuple[str, ...] = ("prev.pptx", "anchor.pptx")
PREV_SLOT = "prev.pptx"
ANCHOR_SLOT = "anchor.pptx"

#: Idle gap (seconds) after which the next mutation is a "new session" and
#: the anchor slot rotates to the current pre-mutation state.
ANCHOR_IDLE_SECONDS = 60 * 60

#: Advisory lockfile name inside a presentation's slot folder.
LOCK_FILE_NAME = "write.lock"
#: How long an acquirer waits for a live holder before refusing.
LOCK_WAIT_SECONDS = 10.0
#: A lockfile older than this is broken regardless of PID liveness
#: (generous: no single mutation legitimately holds the lock this long).
LOCK_STALE_SECONDS = 10 * 60

#: Slot folder names longer than this are truncated + hash-suffixed
#: (keeps total path length sane for Windows even without long-path opt-in).
_MAX_FOLDER_NAME = 80

#: Breadcrumb written when the folder name had to be truncated, so orphan
#: detection can still map the folder back to its source file name.
_SOURCE_NAME_FILE = "source-name.txt"

_RETRY_DELAYS = (0.05, 0.1, 0.2, 0.4)  # seconds, between retries
_TRANSIENT_WINERRORS = {5, 32, 33}  # access denied / sharing violations
#: 32 and 33 are sharing and lock violations; 1175 is ReplaceFileW saying the
#: file being replaced could not be moved aside, which is the same story.
_SHARING_WINERRORS = {32, 33, 1175}
_PROTECTED_WINERRORS = {5, 19}  # access denied / write protected
_PROTECTED_ERRNOS = {errno.EACCES, errno.EPERM, errno.EROFS}


class MutationLockTimeout(PptMcpError):
    """Another process holds the write lock on this file and did not release
    it within the wait window; the mutation was refused, nothing changed."""


def refuse_protected_output(
    output: str | os.PathLike,
    *sources: str | os.PathLike,
    purpose: str = "write output",
) -> None:
    """Refuse output or edit paths that could destroy source or backup data."""
    out_key = canonical_key(output)
    for source in sources:
        same_file = False
        if source:
            source_key = canonical_key(source)
            if source_key == out_key:
                same_file = True
            else:
                # Resolved path strings do not expose hardlink aliases.  When
                # both names exist, samefile compares the underlying file
                # identity (volume + file index on Windows; device + inode on
                # POSIX).  A failed identity read is not treated as equality;
                # the ordinary path and backup-store guards still apply.
                try:
                    same_file = os.path.samefile(output, source)
                except (FileNotFoundError, OSError, ValueError):
                    same_file = False
        if same_file:
            raise PptMcpError(
                f"refusing to {purpose}: the output is the source presentation "
                "itself; choose a new output path"
            )
    parts = {os.path.normcase(part) for part in Path(out_key).parts}
    if os.path.normcase(BACKUP_DIR_NAME) in parts:
        raise PptMcpError(
            f"refusing to {purpose}: the target is inside a backup folder "
            f"({BACKUP_DIR_NAME}); backup data is never an output or edit target"
        )


# ------------------------------------------------------------ path plumbing


def canonical_key(path: str | os.PathLike) -> str:
    """Stable per-file identity: normcase(realpath). Lock keys use this."""
    return os.path.normcase(os.path.realpath(os.path.abspath(os.fspath(path))))


def _folder_name_for(doc_name: str) -> str:
    """Slot subfolder name for a presentation file name. The file's own name
    when it fits (trivial reverse mapping); truncated + 8-hex-hash when long."""
    if len(doc_name) <= _MAX_FOLDER_NAME:
        return doc_name
    digest = hashlib.sha1(
        os.path.normcase(doc_name).encode("utf-8", "surrogatepass")
    ).hexdigest()[:8]
    return f"{doc_name[: _MAX_FOLDER_NAME - 9]}-{digest}"


def slot_dir(doc_path: str | os.PathLike, *, create: bool = False) -> Path:
    """This file's slot folder inside .ks4p-backups (created on demand)."""
    p = Path(doc_path).resolve()
    root = p.parent / BACKUP_DIR_NAME
    folder = _folder_name_for(p.name)
    d = root / folder
    if create:
        d.mkdir(parents=True, exist_ok=True)
        _hide(root)  # dot-prefix everywhere + real hidden bit on Windows
        if folder != p.name:
            crumb = d / _SOURCE_NAME_FILE
            if not crumb.exists():
                try:
                    crumb.write_text(p.name, encoding="utf-8")
                except OSError:
                    pass  # breadcrumb is best-effort
    return d


def source_doc_for(folder: Path) -> Path:
    """Map a slot folder back to its source file path (may not exist)."""
    name = folder.name
    crumb = folder / _SOURCE_NAME_FILE
    if crumb.exists():
        try:
            name = crumb.read_text(encoding="utf-8").strip() or folder.name
        except OSError:
            pass
    return folder.parent.parent / name


def _hide(path: Path) -> None:
    """Set FILE_ATTRIBUTE_HIDDEN on Windows. Non-fatal on any failure."""
    if sys.platform != "win32":
        return
    try:
        import ctypes

        target = str(path)
        if len(target) > 240 and not target.startswith("\\\\?\\"):
            target = "\\\\?\\" + target
        FILE_ATTRIBUTE_HIDDEN = 0x2
        attrs = ctypes.windll.kernel32.GetFileAttributesW(target)
        if attrs != -1 and not (attrs & FILE_ATTRIBUTE_HIDDEN):
            ctypes.windll.kernel32.SetFileAttributesW(
                target, attrs | FILE_ATTRIBUTE_HIDDEN
            )
    except Exception:
        pass  # hidden bit is cosmetic; the dot-prefix stands on its own


# ------------------------------------------------------- retry / placement


def _is_transient(exc: OSError) -> bool:
    winerr = getattr(exc, "winerror", None)
    return isinstance(exc, PermissionError) or winerr in _TRANSIENT_WINERRORS


class SwapNotMade(OSError):
    """ReplaceFileW did not move the new version in, and the deck is exactly
    as it was: error 1176, or 1177 once the original was put back."""


class DeckLeftAside(OSError):
    """ReplaceFileW moved the deck aside, could not move the new version in
    (1177), and the deck could not be moved back: it is at ``aside``."""

    def __init__(self, deck: str, aside: str, text: str, winerror: int):
        super().__init__(0, text, deck, winerror)
        self.aside = aside


def os_write_refusal(
    path: str | os.PathLike, exc: OSError
) -> DocumentLocked | DocumentReadOnly | DeckSetAside | None:
    """Turn recognized OS sharing/protection failures into typed refusals.

    Keep this deliberately narrow.  Device failures, disk-full errors and
    programming defects must continue through the ordinary INTERNAL_ERROR
    boundary instead of being mislabeled as a lock or protection setting.
    """
    target = os.path.abspath(os.fspath(path))
    if isinstance(exc, DeckLeftAside):
        name = os.path.basename(target)
        refusal = DeckSetAside(
            f"The change to {name} could not finish, and the deck could not be "
            "put back in its place. It is kept, as it was before this change, "
            f"at {exc.aside}; nothing in it was lost. The next change to the "
            "deck through this server, or the next start of the server, puts "
            f"it back; you can also rename that file to {name}.")
        refusal.hint = ("Retry the call: it puts the deck back first. If it "
                        "cannot, rename the file named in the message back to "
                        "the deck's name.")
        return refusal
    if isinstance(exc, SwapNotMade):
        refusal = DocumentHeld(
            f"cannot write {target}: the operating system could not move the "
            "new version into place (another program, often an antivirus or "
            "backup scan, had the new copy open). The original file was NOT "
            "modified.")
        refusal.hint = "Retry the call in a moment; nothing was changed."
        return refusal
    winerr = getattr(exc, "winerror", None)
    permission_failure = (
        isinstance(exc, PermissionError)
        or winerr in _PROTECTED_WINERRORS
        or getattr(exc, "errno", None) in _PROTECTED_ERRNOS
    )
    # Windows can report ERROR_ACCESS_DENIED (rather than
    # ERROR_SHARING_VIOLATION) when an existing, otherwise writable target
    # is open without delete sharing.  Explicit read-only attributes are
    # rejected before this helper; at the atomic replace boundary this is
    # therefore the ordinary Explorer/AV/sync-client lock case.
    # DocumentHeld, not a plain DocumentLocked: PowerPoint's own hold was
    # excluded before the save (its owner file is checked first), so a
    # live='auto' call must not go looking for the deck in PowerPoint.
    if winerr in _SHARING_WINERRORS:
        refusal = DocumentHeld(
            f"cannot write {target}: another process is holding the file "
            "or the operating system reported a sharing violation. The "
            "original file was NOT modified."
        )
        refusal.hint = (
            "Close Explorer preview, sync, antivirus, or any other program "
            "reading this path, then retry."
        )
        return refusal
    if winerr == 5 and os.path.exists(target):
        refusal = DocumentHeld(
            f"cannot write {target}: the operating system denied access "
            "while replacing it. Another program may be holding the file, or "
            "this folder's permissions may not allow the change. The "
            "original file was NOT modified."
        )
        refusal.hint = (
            "Close any program reading this path (Explorer preview, sync, "
            "antivirus) and retry; if it repeats, check that you may change "
            "files in this folder."
        )
        return refusal
    if permission_failure:
        refusal = DocumentReadOnly(
            f"cannot write {target}: the operating system denied write "
            "access or the destination is write-protected. The original "
            "file was NOT modified."
        )
        refusal.hint = (
            "Clear the file's read-only attribute or write protection, "
            "grant write access to this path, and retry."
        )
        return refusal
    return None


def _with_retry(fn, *args):
    """Run an os-level file op, retrying briefly on AV sharing violations."""
    for delay in _RETRY_DELAYS:
        try:
            return fn(*args)
        except OSError as exc:
            if not _is_transient(exc):
                raise
            time.sleep(delay)
    return fn(*args)  # final attempt, exceptions propagate


def replace_with_retry(
    src: str | os.PathLike,
    dst: str | os.PathLike,
    *,
    publish: bool = False,
) -> None:
    """Atomically replace ``dst``, optionally claiming the call's write slot.

    Backup-slot maintenance uses the default because it does not publish the
    requested document result.  Final document placement passes
    ``publish=True`` so a timed-out dispatcher can prevent a late write at
    the last possible instant.
    """
    if publish:
        claim_publish()
    _with_retry(os.replace, os.fspath(src), os.fspath(dst))


def claim_publish(what: str = "the save") -> None:
    """Win the deadline race before the first durable publication step."""
    from .deadline import current

    current().commit_or_raise(what)


# --------------------------------------------- keeping the deck's identity
#
# An atomic save writes a new file and puts it at the deck's path. A plain
# os.replace hands the path a NEW file, so one ordinary edit dropped what
# Windows keeps on the file itself rather than in its bytes: the Mark of the
# Web (the Zone.Identifier stream, so a downloaded deck opened in PowerPoint
# without Protected View from then on), the deck's explicit permissions (the
# new file takes whatever its folder hands down), its attributes and its
# creation time. PowerPoint's own save keeps all of them, and so does
# ReplaceFileW, which ``promote`` uses for every save onto a deck that
# exists. Backup slots and other files this server owns keep
# ``replace_with_retry``.

_REPLACE_FILE = None

#: ReplaceFileW errors.
_ERROR_UNABLE_TO_REMOVE_REPLACED = 1175   # nothing changed (the file is held)
_ERROR_UNABLE_TO_MOVE_REPLACEMENT = 1176  # with a backup name: nothing changed
_ERROR_UNABLE_TO_MOVE_REPLACEMENT_2 = 1177  # the original is at the backup name
#: File systems without ReplaceFile (some network and WebDAV drives) answer
#: one of these.
_REPLACE_UNSUPPORTED = {1, 50, 87, 120}

#: ReplaceFileW flag: go on when the deck's permissions cannot be written onto
#: the new file. Used only after the call without it was denied.
_REPLACEFILE_IGNORE_ACL_ERRORS = 0x4

#: The attribute bits put back by hand where ReplaceFileW is not available:
#: hidden, system, archive, offline, not content indexed, and the OneDrive
#: pinned and unpinned states. The others cannot be set this way.
_CARRIED_ATTRIBUTES = 0x2 | 0x4 | 0x20 | 0x1000 | 0x2000 | 0x80000 | 0x100000


def _long_path(path: str) -> str:
    """``path`` absolute, in the extended-length form Windows needs for a
    path over 240 characters (a different prefix for a network path)."""
    target = os.path.abspath(path)
    if len(target) > 240 and not target.startswith("\\\\?\\"):
        if target.startswith("\\\\"):
            target = "\\\\?\\UNC\\" + target[2:]
        else:
            target = "\\\\?\\" + target
    return target


def _replace_file():
    """(ReplaceFileW, MoveFileExW, last error, error text), loaded once."""
    global _REPLACE_FILE
    if _REPLACE_FILE is None:
        import ctypes
        from ctypes import wintypes

        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        fn = k32.ReplaceFileW
        fn.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR,
                       wintypes.DWORD, ctypes.c_void_p, ctypes.c_void_p]
        fn.restype = wintypes.BOOL
        move = k32.MoveFileExW
        move.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD]
        move.restype = wintypes.BOOL
        _REPLACE_FILE = (fn, move, ctypes.get_last_error,
                         lambda e: ctypes.FormatError(e).strip())
    return _REPLACE_FILE


def _read_zone_identifier(path: str) -> bytes | None:
    try:
        with open(_long_path(path) + ":Zone.Identifier", "rb") as fh:
            return fh.read()
    except OSError:
        return None


def _write_zone_identifier(path: str, data: bytes) -> None:
    try:
        with open(_long_path(path) + ":Zone.Identifier", "wb") as fh:
            fh.write(data)
    except OSError:
        pass  # a lost Mark of the Web is a missing safety flag, not lost data


def _creation_time_and_attributes(path: str):
    """(creation FILETIME, attributes), either None when unreadable."""
    import ctypes
    from ctypes import wintypes

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateFileW.restype = wintypes.HANDLE
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
                                wintypes.HANDLE]
    k32.GetFileAttributesW.restype = wintypes.DWORD
    k32.GetFileAttributesW.argtypes = [wintypes.LPCWSTR]
    target = _long_path(path)
    ctime = None
    # FILE_READ_ATTRIBUTES, shared with everyone; BACKUP_SEMANTICS.
    handle = k32.CreateFileW(target, 0x80, 0x7, None, 3, 0x02000000, None)
    if handle and handle != wintypes.HANDLE(-1).value:
        try:
            created = wintypes.FILETIME()
            if k32.GetFileTime(handle, ctypes.byref(created), None, None):
                ctime = created
        finally:
            k32.CloseHandle(handle)
    attributes = k32.GetFileAttributesW(target)
    return ctime, (None if attributes == 0xFFFFFFFF else attributes)


def _put_back_creation_time_and_attributes(path: str, ctime, attributes) -> None:
    import ctypes
    from ctypes import wintypes

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateFileW.restype = wintypes.HANDLE
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
                                wintypes.HANDLE]
    target = _long_path(path)
    if attributes is not None:
        carried = attributes & _CARRIED_ATTRIBUTES
        k32.SetFileAttributesW(target, carried or 0x80)
    if ctime is not None:
        # FILE_WRITE_ATTRIBUTES, shared with everyone; BACKUP_SEMANTICS.
        handle = k32.CreateFileW(target, 0x100, 0x7, None, 3, 0x02000000, None)
        if handle and handle != wintypes.HANDLE(-1).value:
            try:
                k32.SetFileTime(handle, ctypes.byref(ctime), None, None)
            finally:
                k32.CloseHandle(handle)


def _dacl_text(path: str) -> str | None:
    """The file's permissions (its DACL) as SDDL text, or None when they
    cannot be read. Inherited entries carry the ID flag, so a deck whose
    permissions all come from its folder reads the same before and after a
    swap in that folder."""
    import ctypes
    from ctypes import wintypes

    adv = ctypes.WinDLL("advapi32", use_last_error=True)
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    get = adv.GetFileSecurityW
    get.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, ctypes.c_void_p,
                    wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]
    get.restype = wintypes.BOOL
    conv = adv.ConvertSecurityDescriptorToStringSecurityDescriptorW
    conv.argtypes = [ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
                     ctypes.POINTER(wintypes.LPWSTR), ctypes.c_void_p]
    conv.restype = wintypes.BOOL
    k32.LocalFree.argtypes = [ctypes.c_void_p]
    dacl_information = 0x4
    target = _long_path(path)
    need = wintypes.DWORD(0)
    get(target, dacl_information, None, 0, ctypes.byref(need))
    if not need.value:
        return None
    buf = ctypes.create_string_buffer(need.value)
    if not get(target, dacl_information, buf, need, ctypes.byref(need)):
        return None
    text = wintypes.LPWSTR()
    if not conv(buf, 1, dacl_information, ctypes.byref(text), None):
        return None
    try:
        return text.value
    finally:
        k32.LocalFree(ctypes.cast(text, ctypes.c_void_p))


def _replace_by_hand(src: str, dst: str) -> list[str]:
    """Where ReplaceFileW is not there (a file system without it), a plain
    replace, and then what the deck had is put back by hand: its Mark of the
    Web, creation time and attributes. Its permissions are not carried: the
    new file takes the folder's. A deck that had permissions of its own is
    named in a warning, which the answer carries."""
    import contextlib

    zone = _read_zone_identifier(dst)
    ctime = attributes = None
    with contextlib.suppress(Exception):
        ctime, attributes = _creation_time_and_attributes(dst)
    permissions_before = None
    with contextlib.suppress(Exception):
        permissions_before = _dacl_text(dst)
    replace_with_retry(src, dst)
    with contextlib.suppress(Exception):
        _put_back_creation_time_and_attributes(dst, ctime, attributes)
    if zone is not None:
        _write_zone_identifier(dst, zone)
    if permissions_before is None:
        return []
    try:
        kept = _dacl_text(dst) == permissions_before
    except Exception:
        kept = True  # unreadable now: nothing to report
    if kept:
        return []
    return [f"{os.path.basename(dst)} was saved, but its own permissions "
            "could not be kept: this drive cannot carry them onto a new "
            "file, so the deck now has the folder's permissions."]


def promote(src: str | os.PathLike, dst: str | os.PathLike, *,
            ticket: "RotationTicket | None" = None) -> list[str]:
    """Put the new version ``src`` in place of the deck ``dst``, keeping what
    Windows keeps on the file itself rather than in its bytes. Returns the
    warnings the caller passes on (usually none).

    ReplaceFileW (flags 0: an error while merging the deck's permissions,
    streams and attributes into the new file fails the call rather than
    being ignored) carries the deck's explicit permissions, its Mark of the
    Web, its attributes and its creation time onto the new file. In a folder
    where people may write files but not change their permissions (OWNER
    RIGHTS without WRITE_DAC, a common file-share hardening) that call is
    denied; it is made once more ignoring permission errors, which still
    carries everything else, and a deck whose own permissions did not come
    across is named in a warning. The
    original goes to a backup name first, beside it, so that a failure can
    put it back, and is removed after; the backup slots are hard links of
    the original and keep its bytes. The path is missing for the instant
    between the two moves, which a reader in this server waits out
    (core/swapwait.py). A reader that shares delete does not stop the
    swap; one that does not fails it with a sharing violation and changes
    nothing. Transient sharing violations (an antivirus scan) are retried.

    Where the destination is new, or off Windows, a plain replace as before.
    Where the file system has no ReplaceFile, a plain replace and the Mark of
    the Web, creation time and attributes put back by hand.

    ``ticket`` is the backup rotation staged for this swap
    (prepare_rotation). It lands on the slots only after the swap went
    through and is dropped when it did not, so a refused save or restore
    never costs the undo of the last change."""
    src_s = os.fspath(src)
    try:
        dst_s = deck_spelling(dst)
    except (OSError, ValueError):
        dst_s = os.path.abspath(os.fspath(dst))
    entry = backup = None
    try:
        if sys.platform != "win32" or not os.path.exists(dst_s):
            replace_with_retry(src_s, dst_s)
            warnings: list[str] = []
        else:
            try:
                fn, move, last_error, error_text = _replace_file()
            except (OSError, AttributeError):
                fn = None
            if fn is None:
                warnings = _replace_by_hand(src_s, dst_s)
            else:
                folder = os.path.dirname(os.path.abspath(dst_s)) or "."
                backup = os.path.join(folder, aside_name(
                    os.path.basename(dst_s), uuid.uuid4().hex[:12]))
                entry = _journal_open(dst_s, backup, src_s, ticket)
                warnings = _swap(fn, move, last_error, error_text,
                                 src_s, dst_s, backup)
    except BaseException as exc:
        if ticket is not None:
            ticket.abort()
        kept = _still_aside(backup)
        if kept and isinstance(exc, DeckLeftAside):
            _journal_mark(entry, swap="not-made")  # the deck is the old file
        _journal_close(entry, keep=kept)
        raise
    if ticket is not None:
        warnings = [*warnings, *ticket.land(os.path.basename(dst_s))]
    # The journal entry is kept while the old file is still aside (a copy
    # another program holds), so the deck's next write lock tidies it.
    kept = _still_aside(backup)
    if kept:
        _journal_mark(entry, swap="made")
    _journal_close(entry, keep=kept)
    return warnings


def _still_aside(backup: str | None) -> bool:
    return backup is not None and os.path.lexists(_long_path(backup))


def _swap(fn, move, last_error, error_text, src_s: str, dst_s: str,
          backup: str) -> list[str]:
    """ReplaceFileW and what its errors mean; see promote."""
    lsrc, ldst, lbak = _long_path(src_s), _long_path(dst_s), _long_path(backup)
    err = 0
    flags = 0
    permissions_before = None
    for delay in _RETRY_DELAYS + (None,):
        if fn(ldst, lsrc, lbak, flags, None, None):
            try:
                _with_retry(os.remove, lbak)
            except OSError:
                pass  # a program still holds the old copy; it is only a copy
            if flags and permissions_before is not None:
                try:
                    kept = _dacl_text(dst_s) == permissions_before
                except Exception:
                    kept = True  # unreadable now: nothing to report
                if not kept:
                    return [
                        f"{os.path.basename(dst_s)} was saved, but its own "
                        "permissions could not be kept: this folder does not "
                        "let file owners change permissions, so the deck now "
                        "has the folder's permissions."]
            return []
        err = last_error()
        if err == 5 and not flags:
            # Denied with flags 0. Where file owners may not change
            # permissions, ReplaceFileW cannot write the merged permissions
            # onto the new file; once more at once, ignoring that error.
            try:
                permissions_before = _dacl_text(dst_s)
            except Exception:
                permissions_before = None
            flags = _REPLACEFILE_IGNORE_ACL_ERRORS
            continue
        if err == _ERROR_UNABLE_TO_MOVE_REPLACEMENT_2:
            # The original was renamed to the backup name and the new
            # version could not take its place: put the original back.
            if not move(lbak, ldst, 0x1 | 0x8):  # REPLACE_EXISTING | WRITE_THROUGH
                raise DeckLeftAside(
                    dst_s, backup,
                    f"{error_text(err)} The presentation as it was is at "
                    f"{backup}; rename it back to {os.path.basename(dst_s)}.",
                    err)
            break
        if err in _REPLACE_UNSUPPORTED:
            return _replace_by_hand(src_s, dst_s)
        if err not in _TRANSIENT_WINERRORS | {
                _ERROR_UNABLE_TO_REMOVE_REPLACED,
                _ERROR_UNABLE_TO_MOVE_REPLACEMENT} or delay is None:
            break
        time.sleep(delay)
    if err in _TRANSIENT_WINERRORS | {_ERROR_UNABLE_TO_REMOVE_REPLACED}:
        # 32 sharing violation, 5 denied, 1175 the file could not be moved
        # aside: another program holds it, or the permissions refuse. The
        # deck is exactly as it was.
        raise PermissionError(13, error_text(err), dst_s, err)
    if err in (_ERROR_UNABLE_TO_MOVE_REPLACEMENT, _ERROR_UNABLE_TO_MOVE_REPLACEMENT_2):
        # 1176, or 1177 with the original put back: nothing changed.
        raise SwapNotMade(0, error_text(err), dst_s, err)
    raise OSError(0, error_text(err), dst_s, err)


# ------------------------------------------------ a save that did not finish
#
# ReplaceFileW moves the deck to its aside name and then moves the new version
# in. A server stopped between those two moves (a crash, an app closed while
# it saves, a client that ends its servers) leaves the deck's path empty and
# the deck under the aside name, whose words say what it holds. Each swap is
# written down first in a small journal in the server's own folder, so the
# next start of a server finds it; the deck's next write lock puts it back
# from the files beside it even without the journal. Reads never move it:
# they say where it is (core/swapwait.py).
#
# Recovery acts only on what an entry names when it is this deck's own: the
# old file under the deck's aside name in the deck's folder, the new version
# under one of this server's temporary names, and staged backup links in the
# deck's backup folder. It never moves an old copy over a deck that was saved
# again since, and never recreates a deck that was renamed, moved or deleted
# after its save had finished.

#: Journal entries of swaps running in this process right now.
_IN_FLIGHT: set[str] = set()
_IN_FLIGHT_GUARD = threading.Lock()

#: MoveFileExW: on disk before it returns. Without REPLACE_EXISTING it never
#: moves the old deck over anything that is at the deck's path.
_MOVEFILE_WRITE_THROUGH = 0x8

#: An entry whose deck folder has not been there for this long is dropped.
#: The deck's next write still puts it back from the files beside it; only
#: the search for it at every start ends.
_FOLDER_GONE_DROP_SECONDS = 7 * 24 * 3600

#: At start, the longest one entry for a deck on this computer may take to
#: show that its deck's folder is there. The server answers its client only
#: after this part of the recovery, which keeps to its budget (5 seconds).
#: An entry that runs out of time, and every entry for a deck on a network
#: share, is recovered in the background instead (_recover_later).
_FOLDER_PROBE_SECONDS = 1.5

#: In the background, the longest the recovery waits for one share to say
#: whether a deck's folder is there. A share that does not answer can take
#: 20 seconds to say no.
_LATER_PROBE_SECONDS = 60.0

#: GetDriveTypeW answers for drives on this computer: removable, fixed and
#: RAM disks. Network drives, and drive letters that name nothing now, are
#: recovered in the background.
_LOCAL_DRIVE_TYPES = {2, 3, 6}

#: The background recovery that is running or last ran (tests join it).
_LATER: threading.Thread | None = None

#: Staged backup links that no journal entry names are removed from a deck's
#: backup folder, under its write lock, once older than this.
_STALE_SLOT_TEMP_SECONDS = 3600

#: The name of a staged backup link: the time it was made, then a random id.
#: Links made before the time was in the name have the id only.
_SLOT_TEMP_NAME = re.compile(r"\.slot-(?:([0-9]{9,11})-)?[0-9a-f]{32}\.tmp")

#: What the next answer says about a recovery that left something for the
#: person to look at. Each is given once.
_NOTICES: list[str] = []
_NOTICED: set[str] = set()
_NOTICES_GUARD = threading.Lock()


def _slot_temp(folder: Path) -> Path:
    """A new staged backup link's path in a deck's backup folder."""
    return folder / f".slot-{int(time.time())}-{uuid.uuid4().hex}.tmp"


def _log(message: str) -> None:
    """One line on the server's error stream, never the protocol stream."""
    try:
        print(f"KitchenSink4PPT: {message}", file=sys.stderr, flush=True)
    except Exception:
        pass


def _notice(key: str, message: str) -> None:
    """Log ``message`` and keep it for the next answer, once per process for
    the same ``key``."""
    with _NOTICES_GUARD:
        if key in _NOTICED:
            return
        _NOTICED.add(key)
        _NOTICES.append(message)
    _log(message)


def take_notices() -> list[str]:
    """The recovery notices not yet given, for the next answer. Each notice
    is given once."""
    with _NOTICES_GUARD:
        taken = list(_NOTICES)
        _NOTICES.clear()
    return taken


def _file_identity(path) -> dict | None:
    """Size, modification time and file index: what tells one file from
    another without reading it. None when the file cannot be read.

    A backup slot and a staged backup link keep all three: nothing writes to
    them. The deck itself does not: when the old deck carries a named stream
    (a Mark of the Web), ReplaceFileW copies it onto the new version and that
    moves the new version's modification time to the instant of the swap. So
    recovery asks about the slots, never about the deck's own identity."""
    try:
        st = os.stat(_long_path(os.fspath(path)))
    except (OSError, ValueError):
        return None
    return {"size": st.st_size, "mtime_ns": st.st_mtime_ns, "ino": st.st_ino}


def _same_identity(want, have) -> bool:
    """Whether two _file_identity records name the same unchanged file."""
    if not (isinstance(want, dict) and isinstance(have, dict)):
        return False
    if want.get("size") != have.get("size") or want.get("mtime_ns") != have.get("mtime_ns"):
        return False
    if want.get("ino") and have.get("ino"):
        return want.get("ino") == have.get("ino")
    return True


def _journal_open(deck: str, aside: str, new: str,
                  ticket: "RotationTicket | None" = None) -> str | None:
    """Write a swap down before it starts. Best effort: without the entry the
    deck is still put back from the files beside it at its next write lock.
    The entry appears whole (written under a temporary name, then renamed),
    so a reader never takes a half-written one for a finished swap."""
    try:
        folder = _journal_dir()
        folder.mkdir(parents=True, exist_ok=True)
        name = uuid.uuid4().hex
        entry = str(folder / f"{name}.json")
        staging = folder / f"{name}.writing"
        stored = stored_name(deck)
        staging.write_text(json.dumps({
            "deck": os.path.abspath(deck), "aside": os.path.abspath(aside),
            "deck_short": stored[1] if stored else "",
            "new": os.path.abspath(new), "new_id": _file_identity(new),
            "staged": ticket.staged() if ticket is not None else [],
            "slot_ids": ticket.slot_ids() if ticket is not None else {},
            "pid": os.getpid(),
            "token": _OWNER_TOKEN, "pid_created": _OWNER_CREATED,
            "host": _local_host(), "time": time.time(),
        }), encoding="utf-8")
        os.replace(staging, entry)
    except Exception:
        return None
    with _IN_FLIGHT_GUARD:
        _IN_FLIGHT.add(entry)
    return entry


def _journal_rewrite(entry: str, record: dict) -> None:
    """Replace an entry's record, whole. Best effort."""
    try:
        staging = entry[: -len(".json")] + ".writing"
        Path(staging).write_text(json.dumps(record), encoding="utf-8")
        os.replace(staging, entry)
    except Exception:
        pass


def _journal_mark(entry: str | None, **fields) -> None:
    """Add what became of a swap to its entry (it stays in the journal)."""
    if entry is None:
        return
    try:
        with open(entry, encoding="utf-8") as fh:
            record = json.loads(fh.read())
    except (OSError, ValueError):
        return
    if isinstance(record, dict):
        _journal_rewrite(entry, {**record, **fields})


def _journal_close(entry: str | None, *, keep: bool = False) -> None:
    if entry is None:
        return
    with _IN_FLIGHT_GUARD:
        _IN_FLIGHT.discard(entry)
    if not keep:
        try:
            os.remove(entry)
        except OSError:
            pass


def _journal_entries() -> list[tuple[str, dict]]:
    folder = _journal_dir()
    try:
        names = sorted(n for n in os.listdir(folder) if n.endswith(".json"))
    except OSError:
        return []
    found = []
    for name in names:
        entry = str(folder / name)
        try:
            with open(entry, encoding="utf-8") as fh:
                record = json.loads(fh.read())
        except (OSError, ValueError):
            record = None
        found.append((entry, record if isinstance(record, dict) else {}))
    return found


def _settled(entry: str, record: dict) -> bool:
    """A journal entry whose swap is not running: its server is gone, or it
    is this server's and not running now."""
    with _IN_FLIGHT_GUARD:
        if entry in _IN_FLIGHT:
            return False
    if _foreign_host(record):
        return False
    return _is_ours(record) or _is_stale(record)


def _remove_quietly(path) -> None:
    if path is None:
        return
    try:
        os.remove(_long_path(os.fspath(path)))
    except OSError:
        pass


def _there(path) -> bool:
    return isinstance(path, str) and os.path.lexists(_long_path(path))


def _same_folder(a: str, b: str) -> bool:
    """Whether two folder paths name one folder: the same spelling, or the
    same folder on disk (8.3 short names, another case, a drive alias)."""
    if os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b)):
        return True
    try:
        return os.path.samefile(_long_path(a), _long_path(b))
    except (OSError, ValueError):
        return False


def _deck_names(deck: str, others: tuple[str, ...] = ()) -> set[str]:
    """The names the deck's saves may have used for it: its own, and those an
    entry that matched it recorded (the deck as that save spelled it, and its
    short name)."""
    names = {os.path.basename(os.path.abspath(deck))}
    names.update(other for other in others if isinstance(other, str) and other)
    return names


def _own_temporary(deck: str, path, others: tuple[str, ...] = ()) -> bool:
    """Whether ``path`` is one of the temporary files this server writes for
    ``deck``: the new version that a save, a copy or a restore moves into
    its place. Never the deck itself and never any other file. ``others``
    are other names the deck's saves may have used for it (_deck_names)."""
    if not isinstance(path, str) or not path:
        return False
    nc = os.path.normcase
    deck, path = os.path.abspath(deck), os.path.abspath(path)
    if nc(path) == nc(deck):
        return False
    folder = os.path.dirname(deck)
    where, base = os.path.split(path)
    names = _deck_names(deck, others)
    if _same_folder(where, folder):
        own_names = set()
        for name in names:
            digest = hashlib.sha1(name.encode("utf-8", "surrogatepass")).hexdigest()[:16]
            own_names |= {nc(name + ".ppt-mcp-tmp"), nc(f".{digest}.ppt-mcp-tmp"),
                          nc(name + ".ks4p-copy-tmp")}
        own = nc(base) in own_names
    else:
        try:
            backups = str(slot_dir(deck))
        except (OSError, ValueError):
            return False
        own = (_same_folder(where, backups)
               and re.fullmatch(r"\.restore-[0-9a-f]{32}\.tmp", base) is not None)
    if own and os.path.lexists(_long_path(path)) and os.path.lexists(_long_path(deck)):
        try:
            own = not os.path.samefile(_long_path(path), _long_path(deck))
        except OSError:
            pass
    return own


def _own_aside(deck: str, path, others: tuple[str, ...] = ()) -> bool:
    """Whether ``path`` has the name a save gives this deck's old file, in
    the deck's own folder. ``others`` are other names the deck's saves may
    have used for it (_deck_names)."""
    if not isinstance(path, str) or not path:
        return False
    folder = os.path.dirname(os.path.abspath(deck))
    where, base = os.path.split(os.path.abspath(path))
    return (_same_folder(where, folder)
            and any(base.startswith(_swap_prefix(name))
                    for name in _deck_names(deck, others)))


def _move_back(aside: str, deck: str) -> bool:
    """Move the old deck back to its path, never over anything there."""
    if sys.platform != "win32":
        return False  # promote leaves nothing aside off Windows
    try:
        _fn, move, _last_error, _error_text = _replace_file()
    except (OSError, AttributeError):
        return False
    return bool(move(_long_path(aside), _long_path(deck), _MOVEFILE_WRITE_THROUGH))


def _staged_moves(deck: str, record: dict) -> list[tuple[str, str]]:
    """The backup rotation a journal entry staged: (temp, slot) pairs, kept
    only when both sit in this deck's own backup folder."""
    folder = os.path.normcase(str(slot_dir(deck)))
    moves = []
    for pair in record.get("staged") or []:
        if not (isinstance(pair, list) and len(pair) == 2
                and all(isinstance(p, str) for p in pair)):
            continue
        tmp, slot = pair
        if (os.path.normcase(os.path.dirname(tmp)) == folder
                and os.path.basename(tmp).startswith(".slot-")
                and os.path.normcase(os.path.dirname(slot)) == folder
                and os.path.basename(slot) in SLOT_POLICY):
            moves.append((tmp, slot))
    return moves


def _pending_staged(deck: str, record: dict) -> list[tuple[str, str]]:
    """The staged moves of an entry that have not landed on the slots."""
    return [(tmp, slot) for tmp, slot in _staged_moves(deck, record)
            if os.path.lexists(_long_path(tmp))]


def _recorded_slot_ids(record: dict) -> dict:
    """The identity each slot had when the entry's save staged its rotation
    (None for a slot that was not there), keyed by normcase path. Entries
    from builds that did not record them have none."""
    found = record.get("slot_ids")
    if not isinstance(found, dict):
        return {}
    out = {}
    for slot, ident in found.items():
        if not isinstance(slot, str):
            continue
        if ident is None or (isinstance(ident, dict) and {"size", "mtime_ns"} <= set(ident)):
            out[os.path.normcase(slot)] = ident
    return out


def _slot_unmoved(record: dict, tmp: str, slot: str) -> bool:
    """Whether a staged move may still land: its slot is the very file it was
    when the save staged the move (or is not there at all), so landing does
    what the save would have done. A slot that a later save rotated holds a
    newer version than the staged one, and landing it would roll that back.

    An entry from a build that did not record its slots is judged by age: the
    move lands only over a version no newer than the staged one."""
    have = _file_identity(slot)
    if have is None:
        return not os.path.lexists(_long_path(slot))
    try:
        if os.path.samefile(_long_path(tmp), _long_path(slot)):
            return True
    except OSError:
        pass
    recorded = _recorded_slot_ids(record)
    key = os.path.normcase(slot)
    if key in recorded:
        want = recorded[key]
        return want is not None and _same_identity(want, have)
    staged = _file_identity(tmp)
    return staged is not None and have["mtime_ns"] <= staged["mtime_ns"]


def _land_staged(deck: str, record: dict) -> list[tuple[str, str]]:
    """Land the entry's staged moves whose slots have not moved since the
    save staged them; the others stay staged. A move that cannot land now
    (another program has the slot open) is retried briefly and then kept
    staged for the deck's next write lock, never deleted: it may be the only
    copy of the version before the save. Returns those kept moves."""
    held = []
    for tmp, slot in _pending_staged(deck, record):
        if not _slot_unmoved(record, tmp, slot):
            continue
        try:
            _with_retry(os.replace, _long_path(tmp), _long_path(slot))
        except OSError:
            held.append((tmp, slot))
    return held


def _drop_staged(deck: str, record: dict) -> None:
    for tmp, _slot in _staged_moves(deck, record):
        _remove_quietly(tmp)


def _swap_made(deck: str, record: dict) -> bool | None:
    """Whether the entry's swap moved the new version in: True, False, or
    None when the files cannot tell. A swap that went through but left the
    old file aside is marked so in its entry, and so is one whose deck was
    left aside; a stopped server marks nothing, but before the swap its new
    version still sits under its temporary name, and after it, until the
    backups land, the staged links are still there."""
    mark = record.get("swap")
    if mark == "made":
        return True
    if mark == "not-made":
        return False
    if _there(record.get("new")):
        return False
    if _pending_staged(deck, record):
        return True
    return None


def _moved_past(deck: str, record: dict, aside: str) -> bool:
    """Whether the prev backup has moved on past the old copy at ``aside``: a
    later save rotated it, so it holds a newer version than the old copy, as
    it would after any save that was not stopped. Judged by the identity the
    entry recorded for prev when its save staged the rotation, or by age for
    an entry from a build that did not record it."""
    prev = str(slot_dir(deck) / PREV_SLOT)
    have = _file_identity(prev)
    if have is None:
        return False
    recorded = _recorded_slot_ids(record)
    key = os.path.normcase(prev)
    if key in recorded:
        want = recorded[key]
        return want is None or not _same_identity(want, have)
    old = _file_identity(aside)
    return old is not None and have["mtime_ns"] > old["mtime_ns"]


def _same_content(a, b) -> bool:
    """Whether two files hold the same bytes (the same file counts)."""
    a, b = _long_path(os.fspath(a)), _long_path(os.fspath(b))
    try:
        if os.path.samefile(a, b):
            return True
        if os.path.getsize(a) != os.path.getsize(b):
            return False
        with open(a, "rb") as fa, open(b, "rb") as fb:
            while True:
                left, right = fa.read(1 << 20), fb.read(1 << 20)
                if left != right:
                    return False
                if not left:
                    return True
    except (OSError, ValueError):
        return False


def _in_a_slot(deck: str, path: str) -> bool:
    """Whether one of the deck's backup slots holds the bytes of ``path``."""
    folder = slot_dir(deck)
    return any(_same_content(path, folder / slot) for slot in SLOT_POLICY
               if os.path.lexists(_long_path(str(folder / slot))))


def _settle_present(deck: str, entry: str, record: dict) -> bool:
    """Finish an entry for a deck that is in its place. Returns True when the
    entry must stay: it names an old copy kept for the person, or a file
    that could not be removed yet."""
    new, aside = record.get("new"), record.get("aside")
    name = os.path.basename(deck)
    if _swap_made(deck, record) is False:
        # Never swapped in: the deck is as it was, and so are the slots.
        _remove_quietly(new)
        _drop_staged(deck, record)
        if _there(aside):
            # Another file was put at the deck's path while the deck was
            # aside. The deck as it was is the old copy: it never goes over
            # that file, and it is kept and named until it is gone. The entry
            # records that the swap was not made: with its temporary file and
            # staged links gone, the files no longer show it.
            if record.get("swap") != "not-made":
                _journal_mark(entry, swap="not-made")
            _notice(f"{entry}:replaced", (
                f"{name} was replaced by another file while a save of it was "
                f"stopped. The deck as it was before that save is kept at "
                f"{aside}; nothing was put over the file now at {deck}. Open "
                "it, or delete it if you do not need it."))
            return True
        return _there(new)
    held = []
    if _pending_staged(deck, record):
        # The swap ran to its end and only the clean-up was cut short: the
        # backups take the version from before it, as they would have had the
        # server not been stopped, wherever their slot is still the file it
        # was when the save staged them.
        held = _land_staged(deck, record) or []
    aside_there = _there(aside)
    if held:
        # Another program has a slot open: the staged version waits for the
        # deck's next write lock, with the old copy if it is still there.
        _notice(f"{entry}:held", (
            f"The backup of {name} could not be updated after a save of it was "
            "stopped, because another program has the backup open. The version "
            "from before that save is kept, and the deck's next change puts it "
            "in the backup."))
        return True
    blocked = _pending_staged(deck, record)
    if blocked:
        # A save that did not see this entry rotated these slots since, so
        # they hold newer versions than the staged ones, and landing those
        # would roll them back. What they hold may be in another slot now.
        if any(_in_a_slot(deck, tmp) for tmp, _slot in blocked):
            _drop_staged(deck, record)  # a slot holds the same version
        elif aside_there:
            _notice(f"{entry}:changed", (
                f"{name} was saved again after a save of it was stopped, so an "
                f"older copy from that save was kept at {aside} and the backups "
                "were left as they are. Nothing was changed; delete that file "
                "if you do not need it."))
            if not record.get("told"):
                _journal_mark(entry, told=True)
            return True
        else:
            if not record.get("told"):
                # Not the person's own deletion of a copy they were told
                # about: the save itself removed its old copy before it was
                # stopped. The version is older than both backups now, as
                # after any later save, and there is no slot left for it.
                _notice(f"{entry}:dropped", (
                    f"A backup of {name} from a save that was stopped was not "
                    f"kept: {name} has been saved again since, and its backups "
                    "now hold newer versions."))
            _drop_staged(deck, record)
            return False
    if not aside_there:
        return False
    # The old version is still aside. It goes when a backup slot holds its
    # bytes, or when its backups landed and prev has moved on past it since,
    # which is when a save that was not stopped would have let it go.
    landed = bool(record.get("staged"))
    if _in_a_slot(deck, aside) or (landed and _moved_past(deck, record, aside)):
        _remove_quietly(aside)
        return _there(aside)
    # Nowhere else holds that version (a save that kept no backup, or one
    # whose backup could not be kept): the old copy stays, and is named.
    _notice(f"{entry}:kept", (
        f"An older copy of {name}, from before a save of it that was stopped, "
        f"is kept at {aside}. The deck itself is up to date; delete that file "
        "if you do not need it."))
    return True


def _left_where_it_was(deck: str, entry: str, record: dict) -> None:
    """The deck is missing after a save of it had finished: someone renamed,
    moved or deleted it since. The old copy that save left is not made into
    the deck again; the next answer says where it is. The save's backups land
    as it would have landed them, in the backup folder of the deck's old
    name, and the entry records that the swap went through, which the files
    no longer show once the backups have landed."""
    if record.get("swap") != "made":
        _journal_mark(entry, swap="made")
    _land_staged(deck, record)
    _notice(f"{entry}:missing", (
        f"{os.path.basename(deck)} is not at {deck}: it was renamed, moved or "
        "deleted after a save that finished. An older copy from before that "
        f"save is at {record['aside']}. It was not put back, so it never "
        "replaces the deck; open it, or delete it if you do not need it."))


def _lexical_key(path: str) -> str:
    """A path's key from its spelling alone: absolute, separators and case
    normalized, trailing dots and spaces dropped as Windows drops them. No
    file-system call is made, so a share that does not answer cannot hold
    it up."""
    try:
        return os.path.normcase(os.path.abspath(path))
    except (OSError, ValueError):
        return os.path.normcase(path)


def _backups_of_a_finished_save(deck: str, entry: str, record: dict) -> bool:
    """An entry whose swap went through and whose old copy is gone, for a
    deck that is not in its place: its staged backups land as the save would
    have landed them, so the version before that save is kept in the backup
    folder of the deck's name. Returns True when the entry must stay (another
    program has a slot open)."""
    if _land_staged(deck, record):
        return True
    _drop_staged(deck, record)  # what is left is older than its slot holds
    return False


def _sweep_slot_temps(deck: str, named: set[str]) -> None:
    """Under the deck's write lock: remove staged backup links that no
    journal entry names and that are more than an hour old, left by a save
    that was stopped before it could tidy them."""
    try:
        folder = slot_dir(deck)
        names = os.listdir(_long_path(str(folder)))
    except (OSError, ValueError):
        return
    now = time.time()
    for name in names:
        match = _SLOT_TEMP_NAME.fullmatch(name)
        if match is None:
            continue
        path = str(folder / name)
        if _lexical_key(path) in named:
            continue
        if match.group(1):
            made = float(match.group(1))
        else:
            # A link made before the time was in its name: its file's own
            # modification time, the version it holds, is the best age known.
            try:
                made = os.stat(_long_path(path)).st_mtime
            except OSError:
                continue
        if now - made > _STALE_SLOT_TEMP_SECONDS:
            _remove_quietly(path)


def put_back_if_aside(doc_path: str | os.PathLike) -> bool:
    """Put back a deck that a save which did not finish left aside, and tidy
    up after swaps that did finish. Runs with the deck's write lock held
    (write_lock calls it at its outermost hold), so no save of this deck is
    running, here or in another server. True when it moved the deck back.

    Only the deck's own aside files are ever moved, only onto an empty path,
    and none is removed while the deck is missing. A file an entry names
    that is not this deck's own is never moved or removed: it is logged and
    left out of the entry."""
    given = os.path.abspath(os.fspath(doc_path))
    deck = deck_spelling(given)
    # The deck's own spellings: as given, as Windows resolves it, and its
    # real path. An entry is matched against them by its own spelling alone,
    # with no file-system call, so an entry for a deck on a share that does
    # not answer never holds up a change to another deck. Entries this build
    # writes record the deck as Windows resolves it; one an earlier build
    # wrote under another spelling of this deck is finished by the next start
    # of a server, which reaches each entry by its own spelling.
    keys = {_lexical_key(given), _lexical_key(deck), canonical_key(deck)}
    entries = _journal_entries()
    named = {_lexical_key(pair[0])
             for _entry, record in entries for pair in record.get("staged") or []
             if isinstance(pair, list) and pair and isinstance(pair[0], str)}
    ours = []
    for entry, record in entries:
        if not (isinstance(record.get("deck"), str)
                and _lexical_key(record["deck"]) in keys and _settled(entry, record)):
            continue
        record = dict(record)
        others = (os.path.basename(record["deck"]),
                  record.get("deck_short") if isinstance(record.get("deck_short"), str) else "")
        for field, own in (("new", _own_temporary), ("aside", _own_aside)):
            value = record.get(field)
            if value is not None and not own(deck, value, others):
                _log(f"a save journal entry for {deck} names {value!r} as its "
                     f"{field} file, which is not one of this deck's own files; "
                     "it was not touched.")
                record[field] = None
        ours.append((entry, record))
    _sweep_slot_temps(deck, named)
    if os.path.lexists(_long_path(deck)):
        for entry, record in ours:
            _journal_close(entry, keep=_settle_present(deck, entry, record))
        return False
    found = asides(deck)
    for _entry, record in ours:
        aside = record.get("aside")  # judged this deck's own above, or None
        if (isinstance(aside, str) and os.path.lexists(_long_path(aside))
                and os.path.normcase(aside) not in {os.path.normcase(a) for a in found}):
            found.append(aside)
    there = {os.path.normcase(a) for a in found}
    # The deck goes back under the name its save recorded, in its own case.
    deck = next((os.path.abspath(record["deck"]) for _entry, record in ours
                 if _lexical_key(record["deck"]) == _lexical_key(deck)), deck)

    def aside_found(record: dict) -> bool:
        aside = record.get("aside")
        return isinstance(aside, str) and os.path.normcase(aside) in there

    made = {entry for entry, record in ours
            if aside_found(record) and _swap_made(deck, record) is True}
    unmade = [record for entry, record in ours
              if aside_found(record) and entry not in made]
    if not unmade and made:
        # Every old copy here is from a save that finished, so the deck was
        # in its place after it: it is missing because someone renamed,
        # moved or deleted it. Never recreate an old copy under its name.
        for entry, record in ours:
            if entry in made:
                _left_where_it_was(deck, entry, record)
                continue
            swapped = _swap_made(deck, record)
            if swapped is True:
                _journal_close(entry, keep=_backups_of_a_finished_save(deck, entry, record))
                continue
            if swapped is False:
                _remove_quietly(record.get("new"))
            _drop_staged(deck, record)
            _journal_close(entry)
        return False
    chosen = None
    if found:
        recorded = {os.path.normcase(r["aside"]) for r in unmade}
        chosen = next((a for a in found if os.path.normcase(a) in recorded), found[0])
        if not _move_back(chosen, deck):
            return False
    for entry, record in ours:
        if entry in made:
            _journal_close(entry, keep=_settle_present(deck, entry, record))
            continue
        if _swap_made(deck, record) is True:
            _journal_close(entry, keep=_backups_of_a_finished_save(deck, entry, record))
            continue
        _remove_quietly(record.get("new"))  # the version whose save did not finish
        _drop_staged(deck, record)  # the deck is as it was; so are the slots
        _journal_close(entry)
    return chosen is not None


def _folder_there(folder: str, timeout: float) -> bool | None:
    """Whether a folder is there, asked on a helper thread so that a share
    that does not answer cannot hold the server's start: True, False, or
    None when no answer came within ``timeout`` seconds."""
    answer: list[bool] = []

    def ask() -> None:
        try:
            answer.append(os.path.isdir(folder))
        except Exception:
            answer.append(False)

    probe = threading.Thread(target=ask, name="ks4p-recovery-probe", daemon=True)
    probe.start()
    probe.join(max(0.0, timeout))
    return answer[0] if answer else None


def _folder_not_there(entry: str, record: dict) -> None:
    """An entry whose deck folder is not there now (gone, or out of reach):
    dropped once it has not been there for _FOLDER_GONE_DROP_SECONDS, marked
    with when it was first found missing otherwise."""
    since = record.get("folder_gone_since")
    if isinstance(since, (int, float)) and not isinstance(since, bool):
        if time.time() - since >= _FOLDER_GONE_DROP_SECONDS:
            _log(f"dropped the save journal entry for {record.get('deck')}: its "
                 "folder has not been there for 7 days. The deck's next change "
                 "still puts back anything a stopped save left beside it.")
            _journal_close(entry)
        return
    _journal_rewrite(entry, {**record, "folder_gone_since": time.time()})


def _on_this_computer(path: str) -> bool:
    """Whether a deck's path is on a drive of this computer rather than a
    network share. Told from the spelling and the drive's type, without
    asking the drive itself, so a share that does not answer cannot hold
    it up."""
    if sys.platform != "win32":
        return True
    target = _lexical_key(path)
    if target.startswith("\\\\?\\") and target[5:6] == ":":
        target = target[4:]
    if target.startswith("\\\\") or target[1:3] != ":\\":
        return False  # a network path, or a form not known here
    try:
        import ctypes

        kind = ctypes.windll.kernel32.GetDriveTypeW(target[:3])
    except Exception:
        return False
    return kind in _LOCAL_DRIVE_TYPES


def _recover_entry(entry: str, record: dict, *, wait_seconds: float,
                   probe_seconds: float) -> str | None:
    """Recover one settled entry: its deck's folder must answer within
    ``probe_seconds``. Returns the deck when this put it back, "" when the
    folder did not answer in time, and None otherwise."""
    deck = record["deck"]
    there = _folder_there(os.path.dirname(os.path.abspath(deck)), probe_seconds)
    if there is None:
        return ""
    if there is False:
        _folder_not_there(entry, record)
        return None
    if "folder_gone_since" in record:
        record = {k: v for k, v in record.items() if k != "folder_gone_since"}
        _journal_rewrite(entry, record)
    missing = not os.path.lexists(_long_path(deck))
    try:
        with write_lock(deck, wait_seconds=wait_seconds):
            pass  # the outermost hold puts the deck back
    except Exception:
        return None
    if missing and os.path.lexists(_long_path(deck)):
        return deck
    return None


def _recover_later(entries: list[tuple[str, dict]], wait_seconds: float) -> None:
    """Recover entries in the background, after the server has started
    answering: the entries for decks on network shares, and any the start
    had no time for. Each folder gets _LATER_PROBE_SECONDS to answer; one
    that does not is marked as not there (_folder_not_there). A daemon
    thread: it never holds up the server or its exit, and a run that is cut
    short is finished by the next start."""
    global _LATER

    def run() -> None:
        for entry, record in entries:
            try:
                if not os.path.exists(entry):
                    continue  # finished by a change to its deck meanwhile
                if _recover_entry(entry, record, wait_seconds=wait_seconds,
                                  probe_seconds=_LATER_PROBE_SECONDS) == "":
                    _folder_not_there(entry, record)
            except Exception:
                continue

    worker = threading.Thread(target=run, name="ks4p-recovery-later", daemon=True)
    _LATER = worker
    worker.start()


def recover_interrupted_saves(*, wait_seconds: float = 0.5,
                              budget_seconds: float = 5.0) -> list[str]:
    """Put back every deck that a save which did not finish left aside, for
    the journal entries whose server is gone. Run once when a server starts
    (never in a read-only one). A deck whose lock another server holds is
    left to that server: its next write lock does the same.

    The decks on this computer come first and keep to ``budget_seconds``,
    each folder answering within _FOLDER_PROBE_SECONDS. The decks on network
    shares, and any the budget did not reach or whose folder did not answer
    in time, are recovered in the background (_recover_later), so a share
    that does not answer never holds up the start. Returns the decks put
    back before the server answers."""
    restored: list[str] = []
    later: list[tuple[str, dict]] = []
    seen: set[str] = set()
    stop = time.monotonic() + budget_seconds
    for entry, record in _journal_entries():
        if not _settled(entry, record):
            continue
        deck = record.get("deck")
        if not isinstance(deck, str):
            try:
                if time.time() - os.path.getmtime(entry) > 3600:
                    _journal_close(entry)  # unreadable, and long settled
            except OSError:
                pass
            continue
        # A key that needs no file system call: the folder may not answer.
        quick = _lexical_key(deck)
        if quick in seen:
            continue
        seen.add(quick)
        left = stop - time.monotonic()
        if left <= 0 or not _on_this_computer(deck):
            later.append((entry, record))
            continue
        put_back = _recover_entry(entry, record, wait_seconds=wait_seconds,
                                  probe_seconds=min(_FOLDER_PROBE_SECONDS, left))
        if put_back == "":
            later.append((entry, record))
        elif put_back:
            restored.append(put_back)
    if later:
        _recover_later(later, wait_seconds)
    return restored


def _link_or_copy(src: Path, dst: Path) -> None:
    """Hardlink src as dst; copy2 fallback for non-NTFS / cloud placeholders /
    anything the link call rejects. copy2 preserves mtime, which the anchor
    idle measurement relies on."""
    try:
        _with_retry(os.link, str(src), str(dst))
    except OSError:
        _with_retry(shutil.copy2, str(src), str(dst))


def _place_onto_slot(current: Path, slot_path: Path) -> None:
    """Capture `current`'s present content into a slot without the file
    ever leaving its own path: link (or copy) to a temp name in the slot
    folder, then atomically replace the slot."""
    tmp = _slot_temp(slot_path.parent)
    try:
        _link_or_copy(current, tmp)
        replace_with_retry(tmp, slot_path)
    except BaseException:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        raise


# ------------------------------------------------------------- slot rotation


def rotate_slots(doc_path: str | os.PathLike) -> dict:
    """Rotate backup slots for `doc_path` at once, capturing its CURRENT
    content: anchor.pptx when absent or when the prev slot says the file has
    been idle 60+ minutes (a new session), prev.pptx on every call. Never
    unlinks or renames the presentation itself. The save, restore and copy
    paths stage the same rotation with prepare_rotation instead, and land it
    only after their swap went through."""
    return prepare_rotation(doc_path).commit()


class RotationTicket:
    """Two-phase slot rotation for a swap.

    Rotating the slots first and swapping second burned the prev slot when
    the swap was then refused: prev already held the current content, the
    version before the last change was gone, and a later restore prev
    answered ok and changed nothing. prepare_rotation() links the current
    content to temp files in the slot folder; land() (after the swap went
    through) replaces the slots with them; abort() removes them and leaves
    every slot exactly as it was. The links keep the old version alive
    across the swap, so land() still stores the bytes from before it."""

    def __init__(self) -> None:
        self._moves: list[tuple[Path, Path]] = []
        self._before: dict[str, dict | None] = {}
        self.rotated = {"prev": False, "anchor": False}

    def staged(self) -> list[list[str]]:
        return [[str(tmp), str(slot)] for tmp, slot in self._moves]

    def slot_ids(self) -> dict[str, dict | None]:
        """The identity each slot this rotation replaces had when it was
        staged, or None for a slot that was not there. Recovery from a save
        that was stopped lands a staged move only while its slot is still
        that file."""
        return dict(self._before)

    def commit(self) -> dict:
        for tmp, slot in self._moves:
            replace_with_retry(tmp, slot)
            self.rotated["prev" if slot.name == PREV_SLOT else "anchor"] = True
        self._moves = []
        return self.rotated

    def land(self, name: str) -> list[str]:
        """commit(), after a swap that went through; a failure is a warning,
        because the deck itself was saved."""
        try:
            self.commit()
        except OSError as exc:
            self.abort()
            return [f"{name} was saved, but its backup could not be updated "
                    f"({exc}); the prev backup still holds the version from "
                    "before the previous change."]
        return []

    def abort(self) -> None:
        for tmp, _slot in self._moves:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
        self._moves = []


def prepare_rotation(doc_path: str | os.PathLike, *,
                     anchor: bool = True) -> RotationTicket:
    """Stage a slot rotation for `doc_path` WITHOUT touching the slots: the
    anchor/prev policy of rotate_slots (prev only with anchor=False, as a
    restore does). The returned ticket's land() or commit() puts it on the
    slots; abort() drops it."""
    doc = Path(doc_path).resolve()
    d = slot_dir(doc, create=True)
    prev = d / PREV_SLOT
    anchor_slot = d / ANCHOR_SLOT
    ticket = RotationTicket()

    def stage(slot_path: Path) -> None:
        if not os.path.lexists(_long_path(str(slot_path))):
            ticket._before[str(slot_path)] = None
        else:
            seen = _file_identity(slot_path)
            if seen is not None:
                ticket._before[str(slot_path)] = seen
        tmp = _slot_temp(slot_path.parent)
        try:
            _link_or_copy(doc, tmp)
        except BaseException:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
            raise
        ticket._moves.append((tmp, slot_path))

    try:
        if anchor:
            if not anchor_slot.exists():
                stage(anchor_slot)
            elif prev.exists():
                try:
                    idle = time.time() - prev.stat().st_mtime
                except OSError:
                    idle = 0.0
                if idle >= ANCHOR_IDLE_SECONDS:
                    stage(anchor_slot)
        stage(prev)
    except BaseException:
        ticket.abort()
        raise
    return ticket


# ------------------------------------------------------------------ locking

_MUTEXES: dict[str, threading.RLock] = {}
_MUTEX_GUARD = threading.Lock()

#: Open write_lock holds per file in this process, nested ones included.
#: Read and changed only by the thread holding that file's mutex, so zero
#: means no hold in this process is open: a lockfile carrying our token is
#: then a leftover from a release that could not delete it, not a hold.
_DEPTH: dict[str, int] = {}

#: Deletes lockfiles a release could not, while this process lives.
_SWEEPER = lockrelease.Sweeper("ks4p-write-lock-sweeper")


def _mutex_for(key: str) -> threading.RLock:
    with _MUTEX_GUARD:
        lock = _MUTEXES.get(key)
        if lock is None:
            lock = _MUTEXES[key] = threading.RLock()
        return lock


def _pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    if sys.platform == "win32":
        # NEVER os.kill(pid, 0) on Windows - it terminates the process.
        try:
            import ctypes

            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            STILL_ACTIVE = 259
            kernel32 = ctypes.windll.kernel32
            handle = kernel32.OpenProcess(
                PROCESS_QUERY_LIMITED_INFORMATION, False, pid
            )
            if not handle:
                # Access denied means it exists (another user's process).
                return ctypes.get_last_error() == 5 or kernel32.GetLastError() == 5
            try:
                code = ctypes.c_ulong()
                if kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
                    return code.value == STILL_ACTIVE
                return True  # could not query; assume alive (conservative)
            finally:
                kernel32.CloseHandle(handle)
        except Exception:
            return True  # cannot tell; assume alive (never break a live lock)
    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except OSError:
        return True


def _process_create_time(pid: int) -> float | None:
    """Process creation time as a float, or None when it cannot be read.

    Windows only (ctypes GetProcessTimes, no psutil dependency); returns None
    elsewhere, which downgrades staleness detection to plain PID liveness
    rather than breaking it.
    """
    if pid <= 0 or sys.platform != "win32":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(
            PROCESS_QUERY_LIMITED_INFORMATION, False, pid
        )
        if not handle:
            return None
        try:
            creation = wintypes.FILETIME()
            exit_t = wintypes.FILETIME()
            kernel_t = wintypes.FILETIME()
            user_t = wintypes.FILETIME()
            ok = kernel32.GetProcessTimes(
                handle,
                ctypes.byref(creation),
                ctypes.byref(exit_t),
                ctypes.byref(kernel_t),
                ctypes.byref(user_t),
            )
            if not ok:
                return None
            return float(
                (creation.dwHighDateTime << 32) | creation.dwLowDateTime
            )
        finally:
            kernel32.CloseHandle(handle)
    except Exception:
        return None


#: A token unique to THIS process instance, minted once at import. A bare PID
#: is not an identity: PIDs are recycled, so a process that inherited a dead
#: writer's number read that writer's leaked lockfile as its OWN (the
#: re-entrancy amnesty below) and never cleaned it up, while a process that
#: merely shared the number with some unrelated python.exe was reported as the
#: holder in the refusal message. The token settles both: re-entrancy is token
#: equality, and a lock carrying our PID but a foreign token is a recycled-PID
#: leftover, which is stale by definition.
_OWNER_TOKEN = f"{os.getpid()}-{uuid.uuid4().hex}"

#: This process's own creation time, recorded in the payload so a later
#: acquirer can tell "the PID that wrote this is still running" from "some
#: other process now holds that number".
_OWNER_CREATED = _process_create_time(os.getpid())

#: A lockfile whose content we cannot parse is only assumed abandoned after
#: this long. Zero grace was the bug that made the lock leak under load: the
#: old acquire created the file with O_CREAT|O_EXCL and wrote the payload as a
#: SECOND syscall, so a concurrent acquirer that read it in between saw an
#: EMPTY file, parsed no pid, declared it stale, deleted a LIVE lock and took
#: its own. Two writers then believed they held it, and each one's release
#: unlinked the other's file. Writing the payload before the file is visible
#: (below) makes the window impossible; this grace covers a torn write from
#: any other source. Kept under LOCK_WAIT_SECONDS so a genuinely abandoned
#: unreadable lockfile is still cleared inside one wait window.
_UNREADABLE_GRACE_SECONDS = 5.0

#: Poll interval while waiting for a live holder to release.
_POLL_SECONDS = 0.1


def _read_lock_info(lock_path: Path) -> dict:
    try:
        info = json.loads(lock_path.read_text(encoding="utf-8"))
        if isinstance(info, dict):
            return info
    except (OSError, ValueError):
        pass
    return {}


def _publish_lockfile(lock_path: Path) -> bool:
    """Make a COMPLETE lockfile appear atomically, or report that one is
    already there. The payload is written to a private temp file first and
    linked into place, so the lockfile is never observable in a half-written
    state and its timestamp is minted at the instant it becomes visible (the
    old code computed the payload once BEFORE the retry loop, so a lock
    acquired after a wait was born already aged)."""
    payload = json.dumps({
        "pid": os.getpid(),
        "token": _OWNER_TOKEN,
        "pid_created": _OWNER_CREATED,
        "time": time.time(),
        "host": _local_host(),
    })
    tmp = lock_path.parent / f".lock-{uuid.uuid4().hex}.tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as fh:
            fh.write(payload)
        try:
            os.link(str(tmp), str(lock_path))
            return True
        except FileExistsError:
            return False
        except OSError:
            # No hardlink support (non-NTFS, some network shares): fall back
            # to the exclusive-create path, still writing the payload before
            # closing the handle.
            try:
                fd = os.open(str(lock_path),
                             os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            except FileExistsError:
                return False
            try:
                os.write(fd, payload.encode("utf-8"))
            finally:
                os.close(fd)
            return True
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass


def _is_ours(info: dict) -> bool:
    return not _foreign_host(info) and info.get("token") == _OWNER_TOKEN


def _break_lock(lock_path: Path, judged: dict) -> None:
    """Delete a lock nobody holds (stale or unreadable), with one winner
    (core/lockrelease.py break_stale): ``judged`` is the record the acquire
    loop read, or {} for a file it could not read, and the file goes only
    while it is still that one. Two waiters that had both read a dead
    holder's lockfile used to both delete "the lockfile", and the slower one
    deleted the live lock the faster one had just published. A failure is
    left to the acquire loop, which looks again on its next pass."""
    lockrelease.break_stale(lock_path, judged, _read_lock_info,
                            _UNREADABLE_GRACE_SECONDS)


def _deletable(lock_path: Path) -> bool:
    """The lockfile is there and this process may remove it: it carries our
    token, or it cannot be read at all (the old release's rule)."""
    if not lockrelease.present(lock_path):
        return False
    info = _read_lock_info(lock_path)
    return not info or _is_ours(info)


def _is_stale(info: dict) -> bool:
    """A lock nobody can still be holding. Callers must rule out _is_ours
    first, so our own PID reaching here means the number was recycled."""
    if _foreign_host(info):
        return False  # Wait/refuse; only the remote host can establish liveness.
    pid = info.get("pid", -1)
    if not isinstance(pid, int) or not _pid_alive(pid):
        return True
    # A slow live holder retains ownership regardless of elapsed time.
    # Timeout the waiter; never admit a second writer based on age alone.
    if pid == os.getpid():
        # Our own PID under a foreign token: the number was recycled and the
        # writer that held this lock is gone.
        return True
    born = info.get("pid_created")
    now_born = _process_create_time(pid)
    if born is not None and now_born is not None and born != now_born:
        # Same number, different process: Windows recycled the PID. Asking
        # only "is that PID alive" would treat a stranger's process as the
        # live holder and wait the full window for a lock nobody holds.
        return True
    return False


def _acquire_lockfile(
    lock_path: Path, doc_name: str, *, adopt_own: bool = False,
    wait_seconds: float | None = None,
) -> bool:
    """Create the advisory lockfile. Returns True when this call created it
    (and must therefore remove it); False for a re-entrant same-process hold.
    Raises MutationLockTimeout when a live holder does not release in time.

    ``adopt_own`` is set by the outermost hold in this process. A lockfile
    with our token is then a leftover that an earlier release could not
    delete, not a hold, so this call takes it over (returns True) and the
    hold releases it when it ends."""
    waits = LOCK_WAIT_SECONDS if wait_seconds is None else wait_seconds
    deadline = time.monotonic() + waits
    while True:
        if _publish_lockfile(lock_path):
            return True
        info = _read_lock_info(lock_path)
        if _is_ours(info):
            if adopt_own:
                return True
            # This process instance already holds it; the in-process mutex
            # (held by the caller) is the real serializer and a nested
            # acquisition must not deadlock.
            return False
        if not info:
            # Unparseable: give a torn write a moment to settle, then treat a
            # persistently unreadable lock as abandoned.
            try:
                born = lock_path.stat().st_mtime
            except OSError:
                continue                      # it vanished; try again
            if time.time() - born > _UNREADABLE_GRACE_SECONDS:
                _break_lock(lock_path, info)
            elif time.monotonic() > deadline:
                raise MutationLockTimeout(
                    f"{doc_name} is locked by a write.lock this process "
                    f"cannot read. Waited {int(waits)}s; retry, "
                    f"or delete {lock_path} if no other kitchensink4ppt "
                    "process is running."
                )
            else:
                time.sleep(_POLL_SECONDS)
            continue
        if _is_stale(info):
            _break_lock(lock_path, info)
            continue
        if time.monotonic() > deadline:
            pid = info.get("pid")
            stamp = info.get("time", 0.0)
            age = (
                time.time() - stamp
                if isinstance(stamp, (int, float))
                else None
            )
            holder = f"PID {pid}"
            if age is not None:
                holder += f", held for {int(age)}s"
            raise MutationLockTimeout(
                f"{doc_name} is being modified by another kitchensink4ppt "
                f"process ({holder}). Waited {int(waits)}s; "
                "retry once that operation finishes."
            )
        time.sleep(_POLL_SECONDS)


def _release_lockfile(lock_path: Path) -> bool:
    """Remove the lockfile ONLY while it is still ours, and make sure it went.

    Unconditional unlink is how a leaked lock outlived every writer: once two
    processes believed they held the lock, each release deleted whichever file
    happened to be there, including a live one, and the last writer to publish
    was left with nobody to clean up after it.

    A failed unlink is retried (a waiter reading the file blocks the delete
    for a moment on Windows). Returns True once nothing of ours is left, and
    False when the file is still ours after every retry; write_lock then
    hands it to the sweeper."""
    return lockrelease.release_lockfile(lock_path, _deletable)


@contextmanager
def write_lock(doc_path: str | os.PathLike, *, wait_seconds: float | None = None):
    """Serialize the full read-modify-validate-save cycle of one file.

    In-process mutex (per resolved path) + cross-process advisory lockfile in
    the file's slot folder. Hold this around PptxPackage(...) through
    pkg.save() so every mutation sees the previous one's result and response
    metadata (slide indices etc.) is computed against settled state.

    The outermost hold first puts back a deck that a save which did not
    finish left aside (put_back_if_aside), so every write starts from the
    deck in its place. ``wait_seconds`` shortens the wait for another
    holder (the start-up recovery uses it).
    """
    # Public entry that takes an arbitrary path; slot/backup paths all derive
    # from doc_path, so gating it here covers them too. No-op unless
    # KS4P_ALLOWED_ROOTS is set.
    check_path(doc_path, "modify presentation")
    given = doc_path
    try:
        doc_path = deck_spelling(doc_path)
    except (OSError, ValueError):
        pass
    key = canonical_key(doc_path)
    mutex = _mutex_for(key)
    mutex.acquire()
    depth = _DEPTH.get(key, 0)
    _DEPTH[key] = depth + 1
    owns_lockfile = False
    lock_path: Path | None = None
    try:
        try:
            d = slot_dir(doc_path, create=True)
            lock_path = d / LOCK_FILE_NAME
        except OSError:
            lock_path = None  # cannot host a lockfile; mutex still serializes
        if lock_path is not None:
            owns_lockfile = _acquire_lockfile(
                lock_path, Path(doc_path).name, adopt_own=depth == 0,
                wait_seconds=wait_seconds,
            )
            if owns_lockfile:
                _SWEEPER.forget(key)
            if depth == 0:
                try:
                    put_back_if_aside(given)
                except Exception:
                    pass  # the call's own read then says where the deck is
        yield
    finally:
        try:
            if (
                owns_lockfile and lock_path is not None
                and not _release_lockfile(lock_path)
            ):
                _SWEEPER.add(key, lock_path, mutex, _deletable)
        finally:
            if depth:
                _DEPTH[key] = depth
            else:
                _DEPTH.pop(key, None)
            mutex.release()


def _local_host() -> str:
    return os.environ.get("COMPUTERNAME") or socket.gethostname()


def _foreign_host(info: dict) -> bool:
    # Legacy payloads without host metadata retain same-machine behavior.
    # A named remote host cannot be assessed using this machine's PID table.
    host = info.get("host")
    return bool(host) and (not isinstance(host, str)
                           or host.casefold() != _local_host().casefold())
