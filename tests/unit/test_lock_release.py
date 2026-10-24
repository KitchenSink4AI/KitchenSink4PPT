"""A cross-process lockfile never outlives its holder.

A server waiting for a lock reads the lockfile on every poll. On Windows an
ordinary read opens the file without FILE_SHARE_DELETE, so a holder whose
release lands inside that read cannot delete it. The old release swallowed
the failed unlink; the file kept the holder's token; the holder's next
acquisition took it for a nested hold of its own (owns_lockfile False) and
never released it again. Every other server was then refused for that deck,
or on the live route made to wait out the full window on every live edit,
until the holder process exited. Reproduced with two stdio servers: the
second server was refused "held for 12s", then "held for 22s", and got in
only after the first one exited.

The fix (core/lockrelease.py), on both locks:
- the release retries a failed unlink and reports whether the file went;
- a lockfile still ours after the retries goes to a sweeper thread;
- the outermost hold in the process adopts a lockfile with its own token.

Opening the lockfile for reading is exactly what a waiting server does, so
the tests below hold such a handle across a release.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import textwrap
import threading
import time
from pathlib import Path

import pytest

from kitchensink4ppt.com import xproc
from kitchensink4ppt.core import lockrelease, safesave

SRC = Path(__file__).resolve().parents[2] / "src"
IS_WIN = sys.platform == "win32"

needs_windows = pytest.mark.skipif(
    not IS_WIN, reason="only Windows refuses to delete a file another process has open"
)


def _deck(tmp_path: Path) -> Path:
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"not a deck; the lock never reads it")
    return deck


def _lock_of(deck: Path) -> Path:
    return safesave.slot_dir(deck, create=True) / safesave.LOCK_FILE_NAME


def _wait_until(predicate, seconds: float) -> bool:
    end = time.monotonic() + seconds
    while time.monotonic() < end:
        if predicate():
            return True
        time.sleep(0.02)
    return predicate()


# ------------------------------------------------------------- file route


@needs_windows
def test_a_release_blocked_by_a_reader_is_retried_until_it_goes(tmp_path):
    """The usual case: a waiter's read is brief, so a retry succeeds and
    nothing is left over, without the sweeper."""
    deck = _deck(tmp_path)
    lock = _lock_of(deck)
    with safesave.write_lock(deck):
        reader = open(lock, "rb")
        threading.Timer(0.03, reader.close).start()
    assert not lock.exists()
    assert safesave._SWEEPER.pending() == []


@needs_windows
def test_the_next_hold_adopts_a_lockfile_the_release_could_not_delete(tmp_path):
    """The reader stays open through the whole release,
    so every retry fails. The holder's next write must not take the leftover
    for a nested hold; it adopts it and releases it."""
    deck = _deck(tmp_path)
    lock = _lock_of(deck)
    with safesave.write_lock(deck):
        reader = open(lock, "rb")
    try:
        assert lock.exists(), "the reader should have blocked every retry"
        assert json.loads(lock.read_text(encoding="utf-8"))["token"] == safesave._OWNER_TOKEN
    finally:
        reader.close()
    with safesave.write_lock(deck):
        pass
    assert not lock.exists()
    assert safesave._SWEEPER.pending() == []


@needs_windows
def test_the_sweeper_removes_the_leftover_without_another_write(tmp_path):
    """A holder that never writes that deck again must not keep everyone
    else out until it exits."""
    deck = _deck(tmp_path)
    lock = _lock_of(deck)
    with safesave.write_lock(deck):
        reader = open(lock, "rb")
    assert lock.exists()
    assert str(lock) in safesave._SWEEPER.pending()
    reader.close()
    assert _wait_until(lambda: not lock.exists(), 10.0), "the sweeper never removed it"
    assert _wait_until(lambda: safesave._SWEEPER.pending() == [], 5.0)


_WAITER = textwrap.dedent(
    """
    import sys, time
    sys.path.insert(0, sys.argv[2])
    from kitchensink4ppt.core import safesave
    safesave.LOCK_WAIT_SECONDS = float(sys.argv[3])
    t0 = time.monotonic()
    try:
        with safesave.write_lock(sys.argv[1]):
            pass
        print("GOT", round(time.monotonic() - t0, 2), flush=True)
    except safesave.MutationLockTimeout as exc:
        print("REFUSED", exc, flush=True)
    """
)


@needs_windows
def test_another_process_gets_the_deck_after_a_blocked_release(tmp_path):
    """End to end, the way the defect showed itself: after a release that
    could not delete the lockfile, a second server process asks for the same
    deck while the first one stays alive and idle. It used to be refused
    until the first process exited; now it gets in."""
    deck = _deck(tmp_path)
    lock = _lock_of(deck)
    with safesave.write_lock(deck):
        reader = open(lock, "rb")
    reader.close()
    proc = subprocess.run(
        [sys.executable, "-c", _WAITER, str(deck), str(SRC), "6"],
        capture_output=True, text=True, timeout=60, stdin=subprocess.DEVNULL,
    )
    assert proc.stdout.startswith("GOT"), proc.stdout + proc.stderr
    assert not lock.exists()


def test_a_nested_hold_still_leaves_the_lockfile_to_the_outer_hold(tmp_path):
    """Adoption is for the outermost hold only; a nested hold in the same
    process must not take the file over and delete it under its parent."""
    deck = _deck(tmp_path)
    lock = _lock_of(deck)
    with safesave.write_lock(deck):
        with safesave.write_lock(deck):
            assert lock.exists()
        assert lock.exists(), "the nested hold deleted the outer hold's lockfile"
    assert not lock.exists()


def test_the_release_never_deletes_another_holders_lockfile(tmp_path):
    lock = tmp_path / safesave.LOCK_FILE_NAME
    lock.write_text(json.dumps({
        "pid": os.getpid(), "token": "someone-else", "time": time.time(),
    }), encoding="utf-8")
    assert safesave._release_lockfile(lock) is True  # nothing of ours there
    assert lock.exists()


# ------------------------------------------------------------- live route


@pytest.fixture()
def live_lock_dir(tmp_path, monkeypatch):
    """Redirect the live lock so the machine-wide one is never touched."""
    d = tmp_path / "live-locks"
    monkeypatch.setattr(xproc, "_lock_dir", lambda: d)
    return d


@needs_windows
def test_the_live_lock_is_adopted_after_a_blocked_release(live_lock_dir):
    """The live lock covers the whole PowerPoint
    application, so a leak here held every other server's live edits of
    any deck for LOCK_WAIT_SECONDS each, until this process exited."""
    lock = live_lock_dir / f"{xproc.APP_SCOPE}.lock"
    with xproc.cross_process_lock("first live call") as owns:
        assert owns is True
        reader = open(lock, "rb")
    reader.close()
    with xproc.cross_process_lock("a later live call") as owns:
        assert owns is True
    assert not lock.exists()
    assert xproc._SWEEPER.pending() == []


@needs_windows
def test_the_live_sweeper_removes_the_leftover_without_another_call(live_lock_dir):
    lock = live_lock_dir / f"{xproc.APP_SCOPE}.lock"
    with xproc.cross_process_lock("only live call"):
        reader = open(lock, "rb")
    assert lock.exists()
    reader.close()
    assert _wait_until(lambda: not lock.exists(), 10.0), "the sweeper never removed it"
    assert xproc.holder_info() is None


def test_a_nested_live_hold_still_leaves_the_lockfile_to_the_outer_hold(live_lock_dir):
    lock = live_lock_dir / f"{xproc.APP_SCOPE}.lock"
    with xproc.cross_process_lock("outer") as outer:
        with xproc.cross_process_lock("inner") as inner:
            assert (outer, inner) == (True, False)
        assert lock.exists()
    assert not lock.exists()


# ------------------------------------------------------- the sweeper itself


def test_the_sweeper_waits_while_a_hold_in_this_process_is_open(tmp_path):
    """It acts only with the lock's in-process mutex in hand, so it can
    never delete a lockfile that a hold in this process is using."""
    lock = tmp_path / "x.lock"
    lock.write_text("{}", encoding="utf-8")
    mutex = threading.RLock()
    sweeper = lockrelease.Sweeper("test-sweeper")
    mutex.acquire()
    try:
        sweeper.add("k", lock, mutex, lambda p: os.path.lexists(p))
        time.sleep(0.5)
        assert lock.exists(), "the sweeper deleted a lockfile while its mutex was held"
    finally:
        mutex.release()
    assert _wait_until(lambda: not lock.exists(), 10.0)
    assert _wait_until(lambda: sweeper.pending() == [], 5.0)


def test_the_sweeper_leaves_a_lockfile_that_is_no_longer_ours(tmp_path):
    lock = tmp_path / "x.lock"
    lock.write_text("{}", encoding="utf-8")
    sweeper = lockrelease.Sweeper("test-sweeper")
    sweeper.add("k", lock, threading.RLock(), lambda p: False)
    assert _wait_until(lambda: sweeper.pending() == [], 5.0)
    assert lock.exists()


# ---------------- a finished sweep must not retire the newer entry for the
# ---------------- same lockfile


class _StallingMutex:
    """A real RLock that, when the sweeper thread RELEASES it, stalls just
    after really letting go. That is the instant between "the sweep deleted
    the leftover and gave up the lock's mutex" and "the sweep retires its
    queue entry": a hold in this process can run in it, forget the old entry,
    and queue a new one for the same lockfile. Armed once per test."""

    def __init__(self, sweeper_thread_name: str):
        self._real = threading.RLock()
        self._thread_name = sweeper_thread_name
        self.armed = threading.Event()
        self.stalled = threading.Event()
        self.proceed = threading.Event()

    def acquire(self, blocking=True, timeout=-1):
        return self._real.acquire(blocking, timeout)

    def release(self):
        self._real.release()
        if self.armed.is_set() and threading.current_thread().name == self._thread_name:
            self.armed.clear()
            self.stalled.set()
            assert self.proceed.wait(20), "the test never let the sweeper go on"


def test_a_finished_sweep_leaves_a_newer_entry_for_the_same_lockfile_queued(tmp_path):
    """The unit form of the race: the old sweep must retire only the entry it
    worked on. Matching by path alone removed the newer entry, and the
    lockfile was then orphaned."""
    lock = tmp_path / "x.lock"
    lock.write_text("{}", encoding="utf-8")
    name = "test-sweeper-race"
    sweeper = lockrelease.Sweeper(name)
    mutex = _StallingMutex(name)
    ours = lambda p: os.path.lexists(p)  # noqa: E731
    mutex.armed.set()
    sweeper.add("k", lock, mutex, ours)                    # the first entry
    assert mutex.stalled.wait(10), "the sweeper never finished its first pass"
    assert not lock.exists()                               # it deleted the leftover
    mutex.acquire()                                        # a hold in this process comes back:
    try:
        lock.write_text("{}", encoding="utf-8")            # a new lockfile, a failed release,
        sweeper.add("k", lock, mutex, ours)                # a newer entry for the same path
    finally:
        mutex.release()
    assert str(lock) in sweeper.pending()
    mutex.proceed.set()                                    # the old sweep now finishes
    assert _wait_until(lambda: not lock.exists(), 10.0), \
        "the newer entry was retired by the older sweep and its lockfile was left"
    assert _wait_until(lambda: sweeper.pending() == [], 5.0)


_LIVE_WAITER = textwrap.dedent(
    """
    import sys, time
    sys.path.insert(0, sys.argv[1])
    from kitchensink4ppt.com import xproc
    t0 = time.monotonic()
    try:
        with xproc.cross_process_lock("waiter", wait=float(sys.argv[2])):
            pass
        print("GOT", round(time.monotonic() - t0, 2), flush=True)
    except Exception as exc:
        print("REFUSED", type(exc).__name__, exc, flush=True)
    """
)


@needs_windows
def test_the_file_lock_is_not_orphaned_when_a_newer_release_fails_during_an_older_sweep(
        tmp_path, monkeypatch):
    """The same race on the file route (write.lock), end to end: after the
    interleaving above the lockfile must stay queued until cleaned, and a
    second process must get the deck without any further hold by the owner."""
    deck = _deck(tmp_path)
    lock = _lock_of(deck)
    key = safesave.canonical_key(deck)
    mutex = _StallingMutex(safesave._SWEEPER._name)
    monkeypatch.setitem(safesave._MUTEXES, key, mutex)
    with safesave.write_lock(deck):                        # a release a reader blocks:
        reader1 = open(lock, "rb")
    assert lock.exists() and str(lock) in safesave._SWEEPER.pending()
    mutex.armed.set()
    reader1.close()                                        # the sweep can now delete it
    assert mutex.stalled.wait(10), "the sweeper never deleted the first leftover"
    assert not lock.exists()
    with safesave.write_lock(deck):                        # the next hold: a new lockfile,
        reader2 = open(lock, "rb")                         # whose release a reader blocks too
    assert lock.exists() and str(lock) in safesave._SWEEPER.pending()
    mutex.proceed.set()                                    # the old sweep finishes now
    reader2.close()
    # The owner is idle from here on. A second process must get the deck.
    proc = subprocess.run(
        [sys.executable, "-c", _WAITER, str(deck), str(SRC), "8"],
        capture_output=True, text=True, timeout=60, stdin=subprocess.DEVNULL,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    assert proc.stdout.startswith("GOT"), proc.stdout + proc.stderr
    assert _wait_until(lambda: not lock.exists(), 10.0)


@needs_windows
def test_the_live_lock_is_not_orphaned_when_a_newer_release_fails_during_an_older_sweep(
        live_lock_dir, monkeypatch):
    """The same on the app-wide live lock (com/xproc.py), which covers every
    other server's live edits of every deck."""
    lock = live_lock_dir / f"{xproc.APP_SCOPE}.lock"
    mutex = _StallingMutex(xproc._SWEEPER._name)
    monkeypatch.setitem(xproc._MUTEXES, xproc.APP_SCOPE, mutex)
    with xproc.cross_process_lock("first live call") as owns:
        assert owns is True
        reader1 = open(lock, "rb")
    assert lock.exists() and str(lock) in xproc._SWEEPER.pending()
    mutex.armed.set()
    reader1.close()
    assert mutex.stalled.wait(10), "the sweeper never deleted the first leftover"
    assert not lock.exists()
    with xproc.cross_process_lock("second live call") as owns:
        assert owns is True
        reader2 = open(lock, "rb")
    assert lock.exists() and str(lock) in xproc._SWEEPER.pending()
    mutex.proceed.set()
    reader2.close()
    env = {**os.environ, "KS4P_LIVE_LOCK_DIR": str(live_lock_dir)}
    proc = subprocess.run(
        [sys.executable, "-c", _LIVE_WAITER, str(SRC), "8"], env=env,
        capture_output=True, text=True, timeout=60, stdin=subprocess.DEVNULL,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    assert proc.stdout.startswith("GOT"), proc.stdout + proc.stderr
    assert _wait_until(lambda: not lock.exists(), 10.0)
