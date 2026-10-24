"""Cross-PROCESS serialization for live COM sessions.

The same design as the Word editor's com/xproc.py, added after a
concurrency test of the Word editor. That test ran against Word, but this
live layer shares Word's architecture, so it shared the gap:
``com/serial.py`` holds a process-wide RLock and says so honestly, and the
test measured what that scope leaves open. Two independent server
processes against one application overlapped in 49 of 50 operation pairs.
The application's own STA serializes individual COM CALLS, but a tool call
is a SEQUENCE of calls -- resolve a target, then write it -- and nothing
kept two processes from interleaving inside one sequence. In Word the
failure was a find-then-assign replace that double-applied and reported
success to both callers; the PowerPoint equivalents are
live_search_and_replace and every resolve-a-slide-then-mutate-it body in
live_ops.

PowerPoint being a singleton makes this MORE necessary, not less: two
server processes do not merely reach the same application, they cannot
reach anything else. Whatever isolation two Word instances might have
offered is not available here.

Design, the same as the Excel editor's ``core/safesave.write_lock``, which
learned these lessons the expensive way:

- **Atomic publish.** The payload is written to a private temp file and
  hard-linked into place, so the lockfile is never observable
  half-written. The old two-syscall create (O_CREAT|O_EXCL, then write)
  let a concurrent acquirer read an EMPTY file, parse no pid, call it
  stale, delete a LIVE lock and take its own; two holders then unlinked
  each other's files and the lock outlived every writer.
- **Per-process-instance token.** A bare PID is not an identity. PIDs are
  recycled, so a process inheriting a dead writer's number read that
  writer's leaked lockfile as its own forever. Re-entrancy is TOKEN
  equality; our own PID under a foreign token is stale by definition.
- **Release only if ours.** Unconditional unlink is how a leaked lock
  outlives every writer.
- **A release never leaves the file behind** (core/lockrelease.py). A
  waiting server reads the lockfile on every poll, and on Windows that read
  blocks the holder's unlink for a moment. The old release swallowed the
  failure, the file kept our token, and our next acquisition took it for a
  nested hold and never released it: every other server waited out the full
  window on every live edit until this process exited. The release now
  retries, a lockfile still ours goes to a sweeper thread, and the
  outermost hold in this process adopts a lockfile with its own token.
- **Stale detection via process create time.** The Excel editor's PidJournal
  fix proved that "is some process holding that number right now" is
  the wrong question: Windows recycles PIDs, and the answer can be the
  user's own application. The lock payload records the holder's process
  creation time (ctypes ``GetProcessTimes``, no psutil dependency), and a
  live PID whose creation time does not match is a recycled number, which
  makes the lock stale.

Scope: APPLICATION, matching the Word editor and matching what PowerPoint
is. The cost is that two servers driving two different open decks now
queue instead of overlapping, which costs nothing real: the STA already
admitted one COM call at a time.

Deliberate deviation from the Excel editor: the lockfile lives in a local
per-user directory, NOT beside the presentation. The live route never
writes the presentation's folder, the target may sit on OneDrive or a
network share where a stray lockfile would sync, and COM automation is
machine-local, so machine-local is exactly the right scope for the lock.
"""

from __future__ import annotations

import contextlib
import json
import os
import socket
import sys
import tempfile
import threading
import time
import uuid
from pathlib import Path

from ..core import lockrelease
from ..core.errors import LiveLockTimeout

#: Directory holding live-session lockfiles (per user, machine-local).
_LOCK_DIR_ENV = "KS4P_LIVE_LOCK_DIR"

#: One lock covers the PowerPoint application tier (see the scope note above).
APP_SCOPE = "powerpoint-app"

#: How long an acquirer waits for a live holder before refusing. Live COM
#: operations are SLOW -- measured at a 1191 ms median under load,
#: and a 20-mutation batch is one held session -- so the wait is generous.
#: A refusal here must mean "the other server is genuinely stuck", never
#: "the other server was doing ordinary work".
LOCK_WAIT_SECONDS = 120.0

#: Historical compatibility constant; live locks are never evicted by age.
LOCK_STALE_SECONDS = 10 * 60

#: A lockfile we cannot parse is only assumed abandoned after this long
#: (covers a torn write from any source; our own publish is atomic).
_UNREADABLE_GRACE_SECONDS = 30.0

#: Poll interval while waiting for a live holder.
_POLL_SECONDS = 0.05


def _lock_dir() -> Path:
    override = os.environ.get(_LOCK_DIR_ENV)
    if override:
        return Path(override)
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    return Path(base) / "pptx-mcp" / "live-locks"


def _process_create_time(pid: int) -> float | None:
    """Process creation time as a float, or None when it cannot be read.

    Windows only (ctypes GetProcessTimes); returns None elsewhere, which
    downgrades staleness detection to plain PID liveness rather than
    breaking it.
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
                return (
                    ctypes.get_last_error() == 5
                    or kernel32.GetLastError() == 5
                )
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


#: Identity of THIS process instance, minted once at import. PID + a random
#: token: re-entrancy is token equality, and a lock carrying our PID under a
#: foreign token is a recycled-PID leftover (stale by definition).
_OWNER_TOKEN = f"{os.getpid()}-{uuid.uuid4().hex}"
_OWNER_CREATED = _process_create_time(os.getpid())

_MUTEXES: dict[str, threading.RLock] = {}
_MUTEX_GUARD = threading.Lock()

#: Open holds per scope in this process, nested ones included. Read and
#: changed only by the thread holding that scope's mutex, so zero means no
#: hold in this process is open: a lockfile with our token is a leftover.
_DEPTH: dict[str, int] = {}

#: Deletes lockfiles a release could not, while this process lives.
_SWEEPER = lockrelease.Sweeper("ks4p-live-lock-sweeper")


def _mutex_for(key: str) -> threading.RLock:
    with _MUTEX_GUARD:
        lock = _MUTEXES.get(key)
        if lock is None:
            lock = _MUTEXES[key] = threading.RLock()
        return lock


def _read_lock_info(lock_path: Path) -> dict:
    try:
        info = json.loads(lock_path.read_text(encoding="utf-8"))
        if isinstance(info, dict):
            return info
    except (OSError, ValueError):
        pass
    return {}


def _publish_lockfile(lock_path: Path, holder: str) -> bool:
    """Make a COMPLETE lockfile appear atomically, or report that one is
    already there. The payload is written to a private temp file first and
    linked into place, so the lockfile is never observable half-written and
    its timestamp is minted at the instant it becomes visible."""
    payload = json.dumps({
        "pid": os.getpid(),
        "token": _OWNER_TOKEN,
        "pid_created": _OWNER_CREATED,
        "time": time.time(),
        "holder": holder,
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
            # to exclusive-create, still writing the payload before the
            # handle closes.
            try:
                fd = os.open(
                    str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY
                )
            except FileExistsError:
                return False
            try:
                os.write(fd, payload.encode("utf-8"))
            finally:
                os.close(fd)
            return True
    finally:
        with contextlib.suppress(OSError):
            tmp.unlink(missing_ok=True)


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
    """A lock nobody can still be holding."""
    if _foreign_host(info):
        return False  # Wait/refuse; only the remote host can establish liveness.
    pid = info.get("pid", -1)
    if not isinstance(pid, int) or not _pid_alive(pid):
        return True
    # A slow live holder retains ownership regardless of elapsed time.
    # Timeout the waiter; never admit a second writer based on age alone.
    if pid == os.getpid():
        # Our own PID under a foreign token (checked before this call): the
        # number was recycled and the writer is gone.
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
    lock_path: Path, holder: str, wait: float, *, adopt_own: bool = False
) -> bool:
    """Create the advisory lockfile. Returns True when this call created it
    (and must therefore remove it); False for a re-entrant same-process
    hold. Raises LiveLockTimeout when a live holder does not release.

    ``adopt_own`` is set by the outermost hold in this process. A lockfile
    with our token is then a leftover that an earlier release could not
    delete, not a hold, so this call takes it over (returns True) and the
    hold releases it when it ends."""
    deadline = time.monotonic() + wait
    while True:
        if _publish_lockfile(lock_path, holder):
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
            # Unparseable: give a torn write a moment to settle, then treat
            # a persistently unreadable lock as abandoned.
            try:
                born = lock_path.stat().st_mtime
            except OSError:
                continue                      # it vanished; try again
            if time.time() - born > _UNREADABLE_GRACE_SECONDS:
                _break_lock(lock_path, info)
            elif time.monotonic() > deadline:
                raise LiveLockTimeout(
                    "PowerPoint is held by a live-session lock this process "
                    f"cannot read. Waited {int(wait)}s; retry, or delete "
                    f"{lock_path} if no other kitchensink4ppt process is "
                    "running."
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
            detail = f"PID {pid}"
            if info.get("holder"):
                detail += f", running {info['holder']!r}"
            if age is not None:
                detail += f", held for {int(age)}s"
            raise LiveLockTimeout(
                "another kitchensink4ppt server process is driving PowerPoint "
                f"right now ({detail}). Live COM sessions are serialized "
                "across processes so two servers cannot interleave inside "
                f"one edit. Waited {int(wait)}s; nothing was changed. "
                "Retry once that operation finishes."
            )
        time.sleep(_POLL_SECONDS)


def _release_lockfile(lock_path: Path) -> bool:
    """Remove the lockfile ONLY while it is still ours, and make sure it
    went. Unconditional unlink is how a leaked lock outlives every writer.
    A failed unlink is retried; returns False when the file is still ours
    after every retry, and cross_process_lock hands it to the sweeper."""
    return lockrelease.release_lockfile(lock_path, _deletable)


def holder_info(scope: str = APP_SCOPE) -> dict | None:
    """Who holds the cross-process live lock right now, or None. Read-only;
    used by status reporting so a caller can see WHY it would queue."""
    lock_path = _lock_dir() / f"{scope}.lock"
    info = _read_lock_info(lock_path)
    if not info:
        return None
    # Rule out our own hold BEFORE the staleness test: _is_stale treats our
    # own PID as a recycled number (its callers have already checked
    # ownership), so probing it directly here reported this process's own
    # live hold as no-holder and the "ours" field below could never be True.
    if not _is_ours(info) and _is_stale(info):
        return None
    return {
        "pid": info.get("pid"),
        "holder": info.get("holder"),
        "held_for_s": round(time.time() - info.get("time", time.time()), 1),
        "ours": _is_ours(info),
    }


@contextlib.contextmanager
def cross_process_lock(
    holder: str,
    *,
    scope: str = APP_SCOPE,
    wait: float = LOCK_WAIT_SECONDS,
):
    """Serialize a whole live COM session against every other server process.

    Hold this around the FULL attach-resolve-probe-mutate-restore sequence.
    Yields True when this call published the lockfile, False when the hold
    is re-entrant within this process instance (the in-process mutex is the
    real serializer there).

    Degrades to the in-process mutex alone when the lock directory cannot
    be created; it never fails a live edit because a lockfile could not be
    hosted, and it never silently claims cross-process coverage it does not
    have (see ``lock_state``).
    """
    mutex = _mutex_for(scope)
    mutex.acquire()
    depth = _DEPTH.get(scope, 0)
    _DEPTH[scope] = depth + 1
    owns = False
    lock_path: Path | None = None
    try:
        try:
            d = _lock_dir()
            d.mkdir(parents=True, exist_ok=True)
            lock_path = d / f"{scope}.lock"
        except Exception:
            # Broad on purpose: a lock directory that cannot be hosted must
            # never be the reason a live edit fails. The in-process mutex
            # still serializes this process, and lock_state() reports the
            # reduced coverage rather than letting a caller assume it.
            lock_path = None
        if lock_path is not None:
            owns = _acquire_lockfile(lock_path, holder, wait, adopt_own=depth == 0)
            if owns:
                _SWEEPER.forget(scope)
        yield owns
    finally:
        try:
            if (
                owns and lock_path is not None
                and not _release_lockfile(lock_path)
            ):
                _SWEEPER.add(scope, lock_path, mutex, _deletable)
        finally:
            if depth:
                _DEPTH[scope] = depth
            else:
                _DEPTH.pop(scope, None)
            mutex.release()


def lock_state() -> dict:
    """Honest description of the coverage this module is providing, for
    status output: cross-process when a lockfile can be hosted, in-process
    only when it cannot."""
    try:
        d = _lock_dir()
        d.mkdir(parents=True, exist_ok=True)
        return {"cross_process": True, "lock_dir": str(d)}
    except Exception as exc:
        return {"cross_process": False, "reason": str(exc)}


def _local_host() -> str:
    return os.environ.get("COMPUTERNAME") or socket.gethostname()


def _foreign_host(info: dict) -> bool:
    # Legacy payloads without host metadata retain same-machine behavior.
    # A named remote host cannot be assessed using this machine's PID table.
    host = info.get("host")
    return bool(host) and (not isinstance(host, str)
                           or host.casefold() != _local_host().casefold())
