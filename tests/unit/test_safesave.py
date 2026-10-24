"""Slot-based backups (.ks4p-backups/) and per-file write serialization.
Covers: slot rotation (exactly
2 slots ever), anchor idle-gap rotation, hardlink fallback, unicode/Korean
and very long file names, the never-absent-target invariant, lockfile
stale-break and live-holder refusal, re-entrancy, and backup=False semantics.

No ops layer exists yet, so mutations here edit docProps/core.xml's dc:title
directly through PptxPackage: a real dirty-part save cycle with none of the
later tool machinery.
"""

from __future__ import annotations

import os
import shutil
import sys
import time
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core import safesave
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.core.safesave import (
    ANCHOR_SLOT,
    BACKUP_DIR_NAME,
    PREV_SLOT,
    SLOT_POLICY,
    MutationLockTimeout,
)

_DC_TITLE = "{http://purl.org/dc/elements/1.1/}title"
_CORE = "docProps/core.xml"


def _mutate(doc: Path, title: str, *, backup: bool = True) -> None:
    """One full read-modify-validate-save cycle: set the core-properties
    title. Structurally identical to what any future mutating op does."""
    pkg = PptxPackage(doc)
    root = pkg.root(_CORE)
    t = root.find(_DC_TITLE)
    if t is None:
        t = etree.SubElement(root, _DC_TITLE)
    t.text = title
    pkg.mark_dirty(_CORE)
    pkg.save(do_backup=backup)


def _title(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        root = etree.fromstring(zf.read(_CORE))
    return root.findtext(_DC_TITLE) or ""


def _slot_files(doc: Path) -> list[str]:
    d = safesave.slot_dir(doc)
    if not d.is_dir():
        return []
    return sorted(
        p.name for p in d.iterdir() if p.is_file() and p.suffix == ".pptx"
    )


# ------------------------------------------------------------- slot rotation


def test_two_slots_ever_across_many_mutations(make_deck):
    doc = make_deck()
    for i in range(6):
        _mutate(doc, f"edit {i}")
    assert _slot_files(doc) == sorted(SLOT_POLICY)
    d = safesave.slot_dir(doc)
    # anchor = state before the FIRST mutation; prev = state before the LAST.
    assert _title(d / ANCHOR_SLOT) != "edit 0"
    assert _title(d / PREV_SLOT) == "edit 4"
    assert _title(doc) == "edit 5"


def test_backup_root_is_dot_prefixed_and_next_to_doc(make_deck):
    doc = make_deck()
    _mutate(doc, "x")
    root = doc.parent / BACKUP_DIR_NAME
    assert root.is_dir()
    assert (root / doc.name).is_dir()


def test_anchor_rotates_after_idle_gap(make_deck):
    doc = make_deck()
    _mutate(doc, "session one")
    _mutate(doc, "still session one")
    d = safesave.slot_dir(doc)
    # Simulate the idle gap by aging the prev slot's mtime (idle is measured
    # from slot mtimes only - no state database).
    old = time.time() - (safesave.ANCHOR_IDLE_SECONDS + 60)
    os.utime(d / PREV_SLOT, (old, old))
    _mutate(doc, "session two")
    # New-session anchor = content at the start of session two.
    assert _title(d / ANCHOR_SLOT) == "still session one"


def test_anchor_stable_within_session(make_deck):
    doc = make_deck()
    _mutate(doc, "first")
    d = safesave.slot_dir(doc)
    before = (d / ANCHOR_SLOT).read_bytes()
    _mutate(doc, "second")
    _mutate(doc, "third")
    assert (d / ANCHOR_SLOT).read_bytes() == before


def test_hardlink_failure_falls_back_to_copy(make_deck, monkeypatch):
    doc = make_deck()

    def no_link(*args, **kwargs):
        raise OSError(1, "cross-device / non-NTFS / cloud placeholder")

    monkeypatch.setattr(os, "link", no_link)
    _mutate(doc, "copied not linked")
    assert _slot_files(doc) == sorted(SLOT_POLICY)
    d = safesave.slot_dir(doc)
    # Slots are real independent copies and load as valid presentations.
    PptxPackage(d / PREV_SLOT)
    PptxPackage(d / ANCHOR_SLOT)


def test_korean_unicode_doc_name(make_deck):
    doc = make_deck("분기 보고 발표 (최종).pptx")
    _mutate(doc, "한글 제목")
    _mutate(doc, "더 많은 내용")
    d = safesave.slot_dir(doc)
    assert d.name == doc.name  # reverse mapping stays trivial
    assert _slot_files(doc) == sorted(SLOT_POLICY)


def test_very_long_doc_name_gets_hash_suffix_and_breadcrumb(make_deck):
    name = "아주_긴_한글_이름_" * 8 + "final_deck.pptx"
    assert len(name.encode("utf-8")) <= 255  # ext4 safety
    assert len(name) > 80  # long enough to trigger the folder-name truncation
    doc = make_deck(name)
    _mutate(doc, "long name content")
    d = safesave.slot_dir(doc)
    assert d.is_dir() and len(d.name) < len(name)
    # Breadcrumb maps the truncated folder back to the source file.
    assert safesave.source_doc_for(d) == doc


# ------------------------------------------- what touches the deck's path


def _norm(p) -> str:
    """A path as the trace compares it: absolute, case-folded where the file
    system folds case, and without the extended-length prefix the native
    Windows calls take."""
    s = os.fspath(p)
    unc, local = "\\\\?\\UNC\\", "\\\\?\\"
    if s.startswith(unc):
        s = "\\\\" + s[len(unc):]
    elif s.startswith(local):
        s = s[len(local):]
    return os.path.normcase(os.path.abspath(s))


#: The events that put a new version at the deck's path: a replace made
#: inside promote (off Windows, and where a file system has no ReplaceFile)
#: and a ReplaceFileW call that went through (Windows).
_SWAPS = ("swap", "native")


def _trace_one_mutation(doc: Path, monkeypatch, *, inner=None):
    """Run one mutation of ``doc`` with every operation that can touch a
    path traced. ``inner`` wraps the real promote (a mutant for the test
    that the check catches it). Returns (events, target, prev slot)."""
    target = _norm(doc.resolve())
    events: list[tuple[str, str, str]] = []
    real_link, real_replace = os.link, os.replace
    real_unlink, real_remove = os.unlink, os.remove

    monkeypatch.setattr(
        os, "link",
        lambda s, dst, **kw: (events.append(("link", _norm(s), _norm(dst))),
                              real_link(s, dst, **kw))[1],
    )
    in_promote = []

    def traced_replace(s, dst):
        swap = bool(in_promote) and _norm(dst) == target
        events.append(("swap" if swap else "replace", _norm(s), _norm(dst)))
        return real_replace(s, dst)

    monkeypatch.setattr(os, "replace", traced_replace)
    monkeypatch.setattr(
        os, "unlink",
        lambda p, *a, **kw: (events.append(("unlink", _norm(p), "")),
                             real_unlink(p, *a, **kw))[1],
    )
    monkeypatch.setattr(
        os, "remove",
        lambda p, *a, **kw: (events.append(("unlink", _norm(p), "")),
                             real_remove(p, *a, **kw))[1],
    )
    if sys.platform == "win32":
        real_native = safesave._replace_file

        def traced_native():
            fn, move, last_error, error_text = real_native()

            def replace_file(dst, src, *rest):
                made = fn(dst, src, *rest)
                if made:
                    events.append(("native", _norm(src), _norm(dst)))
                return made

            def move_file(src, dst, *rest):
                made = move(src, dst, *rest)
                if made:
                    events.append(("move", _norm(src), _norm(dst)))
                return made

            return replace_file, move_file, last_error, error_text

        monkeypatch.setattr(safesave, "_replace_file", traced_native)
    real_promote = safesave.promote if inner is None else inner(safesave.promote)

    def traced_promote(s, dst, **kw):
        events.append(("promote", _norm(s), _norm(dst)))
        in_promote.append(1)
        try:
            return real_promote(s, dst, **kw)
        finally:
            in_promote.pop()

    monkeypatch.setattr(safesave, "promote", traced_promote)
    _mutate(doc, "mutation")
    return events, target, _norm(safesave.slot_dir(doc) / PREV_SLOT)


def _target_swaps(events, target) -> list[tuple[str, str, str]]:
    return [e for e in events if e[0] in _SWAPS and e[2] == target]


def _assert_the_one_swap(events, target, prev_slot) -> None:
    # The presentation itself is never unlinked/removed at any point.
    assert all(not (op == "unlink" and src == target) for op, src, _ in events)
    touches = [e for e in events if e[1] == target or e[2] == target]
    kinds = [e[0] for e in touches]
    # One promote, and inside it exactly one swap: ReplaceFileW on Windows,
    # the plain replace elsewhere. No replace or move outside it lands on
    # the deck.
    assert kinds.count("promote") == 1, touches
    assert kinds.count("replace") == 0, touches
    assert kinds.count("move") == 0, touches
    swaps = _target_swaps(events, target)
    expected = "native" if sys.platform == "win32" else "swap"
    assert [e[0] for e in swaps] == [expected], touches
    # Required sequence: stage the current target (link), swap the validated
    # temp in for the target, and only THEN land the staged link on prev.
    i_link = next(
        i for i, (op, src, _) in enumerate(events)
        if op == "link" and src == target
    )
    i_prev = next(
        i for i, (op, _, dst) in enumerate(events)
        if op == "replace" and dst == prev_slot
    )
    i_promote = next(
        i for i, (op, _, dst) in enumerate(events)
        if op == "promote" and dst == target
    )
    i_swap = events.index(swaps[0])
    assert i_link < i_promote < i_swap < i_prev


def test_the_deck_is_only_touched_by_the_one_swap_after_the_slot_capture(make_deck, monkeypatch):
    """The deck is never unlinked or removed. The slot capture is staged
    first (a link of the deck), then one swap, safesave.promote, is the only
    operation that puts anything at the deck's path, and only after it went
    through does the staged link replace prev. The link keeps the old bytes
    alive across the swap, so they end up under prev. Inside promote the
    deck takes exactly one new version: ReplaceFileW on Windows, one plain
    replace elsewhere."""
    doc = make_deck()
    _mutate(doc, "seed")  # anchor set
    events, target, prev_slot = _trace_one_mutation(doc, monkeypatch)
    _assert_the_one_swap(events, target, prev_slot)


def test_the_swap_check_fails_a_promote_that_writes_the_deck_twice(make_deck, monkeypatch):
    """A promote that puts two versions at the deck's path inside the one
    call fails the check above, on every platform: here a plain replace of a
    copy onto the deck before the real swap."""
    doc = make_deck()
    _mutate(doc, "seed")

    def twice(real_promote):
        def promote(s, dst, **kw):
            twin = os.fspath(s) + ".twin"
            shutil.copyfile(s, twin)
            os.replace(twin, dst)
            return real_promote(s, dst, **kw)
        return promote

    events, target, prev_slot = _trace_one_mutation(doc, monkeypatch, inner=twice)
    assert len(_target_swaps(events, target)) == 2, events
    with pytest.raises(AssertionError):
        _assert_the_one_swap(events, target, prev_slot)


# ------------------------------------------------------------------ locking


def test_stale_lockfile_is_broken(make_deck):
    doc = make_deck()
    _mutate(doc, "seed")
    d = safesave.slot_dir(doc)
    lock = d / safesave.LOCK_FILE_NAME
    lock.write_text(
        '{"pid": 999999999, "time": 1.0}', encoding="utf-8"
    )  # dead pid, ancient timestamp
    with safesave.write_lock(doc):
        _mutate(doc, "after stale break")
    assert _title(doc) == "after stale break"
    assert not lock.exists()


def test_live_foreign_lock_refused_with_holder_named(make_deck, monkeypatch):
    doc = make_deck()
    _mutate(doc, "seed")
    d = safesave.slot_dir(doc)
    lock = d / safesave.LOCK_FILE_NAME
    foreign_pid = os.getpid() + 12345
    lock.write_text(
        f'{{"pid": {foreign_pid}, "time": {time.time()}}}', encoding="utf-8"
    )
    monkeypatch.setattr(safesave, "_pid_alive", lambda pid: True)
    monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 0.3)
    with pytest.raises(MutationLockTimeout, match=str(foreign_pid)):
        with safesave.write_lock(doc):
            pass
    lock.unlink()  # release for other tests' sake


def test_same_process_lockfile_is_reentrant(make_deck):
    doc = make_deck()
    _mutate(doc, "seed")
    with safesave.write_lock(doc):
        # A nested acquisition in the same process must not deadlock or refuse.
        with safesave.write_lock(doc):
            _mutate(doc, "nested")
    assert _title(doc) == "nested"


# ------------------------------------------------------------ backup=False


def test_backup_false_skips_rotation_never_the_atomic_save(make_deck, tmp_path):
    doc = make_deck()
    _mutate(doc, "no backup", backup=False)
    assert _title(doc) == "no backup"
    assert _slot_files(doc) == []  # no slots rotated
    # Atomic save still applied: no temp litter, presentation valid.
    assert not list(tmp_path.glob("*.ppt-mcp-tmp"))
    PptxPackage(doc)
