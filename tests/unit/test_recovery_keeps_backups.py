"""Recovery from a stopped save keeps every version it can, and says so when it
cannot.

A save that is stopped in its last instants leaves the version from before it
in a staged backup link (and, before its old copy is removed, in that old copy
beside the deck). Recovery puts it on the backup slots as the save would have.
When it cannot do that now, it keeps the version for later and the next
answer says so; when it cannot keep it at all, the next answer says that. An
old copy kept beside a deck is always named in the next answer.

Each case is a real kill of a saving process at one instant (see
test_recovery_marked_decks.py for the instants).
"""

from __future__ import annotations

import asyncio
import ctypes
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

import pytest
from pptx import Presentation

from kitchensink4ppt import server
from kitchensink4ppt.core import safesave, swapwait
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.plan import build

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="ReplaceFileW")

NO_WINDOW = 0x08000000

VICTIM = r'''
import asyncio, os, sys, time
from kitchensink4ppt import server  # noqa: F401
from kitchensink4ppt.core import safesave
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.plan import build
deck, marker, mode, backup = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4] == "1"
fn, move, last_error, text = safesave._replace_file()

def stop():
    open(marker, "w").close()
    time.sleep(600)

def replace(ldst, lsrc, lbak, flags, a, b):
    if mode == "mid":
        os.replace(ldst, lbak)
        stop()
    ok = fn(ldst, lsrc, lbak, flags, a, b)
    if mode == "late":
        stop()
    return ok

safesave._REPLACE_FILE = (replace, move, last_error, text)
if mode == "postremove":
    _land = safesave.RotationTicket.land
    def land(self, name):
        stop()
        return _land(self, name)
    safesave.RotationTicket.land = land
plan = build(Settings())
asyncio.run(dispatch(plan, "ppt_edit", {"op": "set_notes", "args": {
    "file_path": deck, "slide": 0, "text": "v2", "live": "off", "backup": backup}}))
'''

SAVE_IN_A_CHILD = (
    "import asyncio, sys\n"
    "from kitchensink4ppt import server\n"
    "from kitchensink4ppt.settings import Settings\n"
    "from kitchensink4ppt.surface.dispatch import dispatch\n"
    "from kitchensink4ppt.surface.plan import build\n"
    "plan = build(Settings())\n"
    "asyncio.run(dispatch(plan, 'ppt_edit', {'op': 'set_notes', 'args': {\n"
    "    'file_path': sys.argv[1], 'slide': 0, 'text': sys.argv[2], 'live': 'off'}}))\n")


@pytest.fixture(autouse=True)
def own_journal(tmp_path, monkeypatch):
    monkeypatch.setenv("KS4P_SAVE_JOURNAL_DIR", str(tmp_path / "journal"))
    safesave.take_notices()
    yield
    safesave.take_notices()


def _deck(path: Path, notes: str) -> Path:
    prs = Presentation()
    for title in ("Alpha", "Bravo"):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = title
        slide.notes_slide.notes_text_frame.text = notes
    prs.save(path)
    return path


def _notes(path) -> str | None:
    if not os.path.lexists(path):
        return None
    return Presentation(str(path)).slides[0].notes_slide.notes_text_frame.text


def _call(op, args):
    plan = build(Settings())
    return asyncio.run(dispatch(plan, plan.specs[op].group, {"op": op, "args": args}))


def _stopped(tmp_path: Path, mode: str, *, backup: bool = True) -> Path:
    """v0, then v1 through the product (prev = v0), then a save of v2 stopped
    at `mode`."""
    deck = _deck(tmp_path / "Deck.pptx", "v0")
    _payload, refused = _call("set_notes", {"file_path": str(deck), "slide": 0, "text": "v1",
                                            "live": "off"})
    assert not refused
    script = tmp_path / "victim.py"
    script.write_text(VICTIM, encoding="utf-8")
    marker = tmp_path / "stopped"
    victim = subprocess.Popen([sys.executable, "-X", "utf8", str(script), str(deck), str(marker),
                               mode, "1" if backup else "0"],
                              stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                              stderr=subprocess.DEVNULL, creationflags=NO_WINDOW)
    try:
        deadline = time.monotonic() + 120
        while not marker.exists():
            assert victim.poll() is None, "the saving process ended before the stop"
            assert time.monotonic() < deadline, "the saving process never reached the stop"
            time.sleep(0.02)
    finally:
        victim.kill()  # TerminateProcess on this test's own child
        victim.wait(timeout=60)
    time.sleep(0.3)
    return deck


def _prev(deck: Path) -> Path:
    return safesave.slot_dir(deck) / safesave.PREV_SLOT


def _staged_versions(deck: Path) -> list[str | None]:
    folder = safesave.slot_dir(deck)
    return [_notes(folder / n) for n in sorted(os.listdir(folder)) if n.startswith(".slot-")]


def _entries_for(deck: Path) -> list[str]:
    key = os.path.normcase(os.path.abspath(str(deck)))
    return [entry for entry, record in safesave._journal_entries()
            if isinstance(record.get("deck"), str)
            and os.path.normcase(os.path.abspath(record["deck"])) == key]


class _Held:
    """Another program holds a file open for reading and writing, without
    sharing delete, as a backup or sync tool does."""

    def __init__(self, path: Path):
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.CreateFileW.restype = ctypes.c_void_p
        k32.CreateFileW.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32, ctypes.c_uint32,
                                    ctypes.c_void_p, ctypes.c_uint32, ctypes.c_uint32,
                                    ctypes.c_void_p]
        k32.CloseHandle.argtypes = [ctypes.c_void_p]
        self._k32 = k32
        self.handle = k32.CreateFileW(str(path), 0x80000000, 0x3, None, 3, 0, None)
        assert self.handle not in (None, ctypes.c_void_p(-1).value), "could not hold the file"

    def release(self) -> None:
        if self.handle:
            self._k32.CloseHandle(self.handle)
            self.handle = None


@pytest.mark.parametrize("mode", ["late", "postremove"])
def test_a_backup_slot_held_by_another_program_keeps_the_staged_version(tmp_path, mode):
    """The prev slot is held while the start finishes a stopped save. The
    version before the last change must not be thrown away: it is kept, the
    next answer says the backup could not be updated yet, and the next change
    after the hold puts it in place."""
    deck = _stopped(tmp_path, mode)
    held = _Held(_prev(deck))
    try:
        server._recover_at_start(Settings())
        notices = safesave.take_notices()
        kept = "v1" in _staged_versions(deck) or "v1" in [_notes(a) for a in swapwait.asides(deck)]
        assert kept, "v1, the version before the last change, was thrown away"
        assert any("backup" in n and deck.name in n for n in notices), notices
        assert _entries_for(deck), "the record of the save was closed with its work undone"
    finally:
        held.release()
    payload, refused = _call("set_notes", {"file_path": str(deck), "slide": 0, "text": "v3",
                                           "live": "off"})
    assert not refused, payload
    assert _notes(deck) == "v3" and _notes(_prev(deck)) == "v2"
    assert swapwait.asides(deck) == [] and _staged_versions(deck) == []
    assert _entries_for(deck) == []


def test_a_staged_version_that_cannot_land_is_never_dropped_without_a_word(tmp_path):
    """The save removed its old copy and was stopped before its backups landed;
    then a server that cannot see this record (another build, another
    journal) saved the deck and rotated the backups. The staged version is now
    older than both backups and has no slot. Dropping it is what the rotation
    would have done, but the next answer must say so."""
    deck = _stopped(tmp_path, "postremove")
    env = dict(os.environ, KS4P_SAVE_JOURNAL_DIR=str(tmp_path / "another-journal"))
    subprocess.run([sys.executable, "-X", "utf8", "-c", SAVE_IN_A_CHILD, str(deck), "v3"],
                   check=True, timeout=300, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL, creationflags=NO_WINDOW, env=env)
    assert _notes(deck) == "v3" and _notes(_prev(deck)) == "v2"
    server._recover_at_start(Settings())
    notices = safesave.take_notices()
    assert _notes(_prev(deck)) == "v2", "an older backup was put over a newer one"
    if "v1" not in _staged_versions(deck):
        assert any(deck.name in n and "backup" in n for n in notices), (
            "v1 was dropped and nothing said so", notices)


def test_a_deck_replaced_by_hand_while_it_was_aside_keeps_its_old_copy_and_says_where(tmp_path):
    """A save was stopped between its two moves, so the deck was aside; then
    the person copied another file to the deck's path. That file is never
    replaced, the deck as it was stays where it is, and the next answer names
    where it is (once per run)."""
    deck = _stopped(tmp_path, "mid")
    other = _deck(tmp_path / "Other.pptx", "v9")
    shutil.copy2(other, deck)
    server._recover_at_start(Settings())
    notices = safesave.take_notices()
    aside = swapwait.asides(deck)
    assert _notes(deck) == "v9"
    assert len(aside) == 1 and _notes(aside[0]) == "v1"
    assert any(aside[0] in n for n in notices), ("the old copy was kept without a word", notices)
    payload, _ = _call("get_text", {"file_path": str(deck), "live": "off"})
    assert "notices" not in payload, "each notice is given once"
    # The next run of a server names it the same way: the deck was replaced.
    safesave._NOTICED.clear()
    with safesave.write_lock(deck):
        pass
    again = safesave.take_notices()
    assert any(aside[0] in n and "replaced by another file" in n for n in again), again
    assert _notes(deck) == "v9" and _notes(aside[0]) == "v1"


def test_a_save_without_a_backup_stopped_in_its_last_instant_names_the_old_copy(tmp_path):
    """backup=False keeps no backup; a stop in the save's last instant left the
    old copy beside the deck. It is kept, and the next answer says where."""
    deck = _stopped(tmp_path, "late", backup=False)
    server._recover_at_start(Settings())
    notices = safesave.take_notices()
    aside = swapwait.asides(deck)
    assert _notes(deck) == "v2"
    assert len(aside) == 1 and _notes(aside[0]) == "v1"
    assert any(aside[0] in n for n in notices), ("the old copy was kept without a word", notices)


def test_a_deck_renamed_after_a_stop_keeps_the_backup_of_its_last_change(tmp_path):
    """The save removed its old copy and was stopped before its backups landed;
    then the person renamed the deck. The staged version (the one before the
    last change) lands in the backups of the deck's old name, as the save would
    have put it there, instead of being thrown away."""
    deck = _stopped(tmp_path, "postremove")
    os.replace(deck, tmp_path / "Final.pptx")
    server._recover_at_start(Settings())
    assert not deck.exists(), "the old name must not be made again"
    assert _notes(_prev(deck)) == "v1" or "v1" in _staged_versions(deck), (
        "v1 was thrown away")
    assert _notes(tmp_path / "Final.pptx") == "v2"
