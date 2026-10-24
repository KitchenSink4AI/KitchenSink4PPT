"""Process-wide COM serialization (the same design as the Word editor).

Stress-testing the Word editor's live COM layer (concurrent callers, one
open document) proved that unserialized concurrent COM access produces
character-level interleaved writes, modal-dialog deadlocks on the save
temp-file swap, and unbounded hangs. PowerPoint inherits every one of
those failure modes and adds one
of its own: POWERPNT is a STRICT SINGLETON COM server. Word hands out a
private invisible instance per DispatchEx; PowerPoint does not. Every COM
path in this server therefore converges on ONE PowerPoint process, so
there is no per-instance isolation to fall back on and serialization is
not a mitigation here, it is the only thing standing between two
concurrent tool calls and a shared mutable application object.

Coverage contract: every COM entry point acquires the lock —
- the live layer (live.live_session wraps every live tool; live_status
  uses the bounded try-acquire form),
- the bridge layer (every public invisible-instance function, enforced by
  test_com_serialization's audit of the _com_serialized marker).

The lock is an RLock: nested acquisitions on one thread are legal (a
bridge helper called under an already-held lock must not deadlock).

Scope honesty: this lock serializes THIS server process only, and on a
singleton that was a real hole. A concurrency test measured the size of
it on the Word editor: two server processes against one application
overlapped in 49 of 50 operation pairs and corrupted a
resolve-then-write sequence while reporting success to both callers.
Cross-PROCESS serialization is a separate, file-based lock in
com/xproc.py, which every live session now holds on top of this one. Keep
the two distinct: this one is cheap, in-memory, and covers threads; that
one is a lockfile and covers the machine. Status reporting (lock_snapshot,
plus xproc.holder_info) lets powerpoint_status and live_status tell
callers honestly when a call would queue and on whom.
"""

from __future__ import annotations

import contextlib
import threading
import time

COM_LOCK = threading.RLock()

_state_lock = threading.Lock()
_current: dict | None = None      # {name, thread, started_wall, started_mono}
_last: dict | None = None         # completed operation record
_depth = 0                        # re-entrant depth on the owning thread
_serialized_total = 0             # ops that ran under the lock
_waited_total_ms = 0.0            # cumulative wait time across ops


@contextlib.contextmanager
def com_operation(name: str):
    """Hold the process-wide COM lock for the duration of one COM-touching
    operation. Records timing so the status tools can report contention."""
    global _current, _last, _depth, _serialized_total, _waited_total_ms
    t0 = time.monotonic()
    COM_LOCK.acquire()
    waited_ms = (time.monotonic() - t0) * 1000.0
    with _state_lock:
        _depth += 1
        outermost = _depth == 1
        if outermost:
            _current = {
                "name": name,
                "thread": threading.get_ident(),
                "started_wall": time.time(),
                "started_mono": time.monotonic(),
                "outcome": "running",
            }
            _serialized_total += 1
            _waited_total_ms += waited_ms
    started = time.monotonic()
    failed = False
    try:
        yield
    except BaseException:
        failed = True
        raise
    finally:
        with _state_lock:
            _depth -= 1
            if outermost:
                if _current["outcome"] == "running":
                    _current["outcome"] = "failed" if failed else "succeeded"
                _last = {
                    "name": name,
                    "thread": _current["thread"],
                    "duration_ms": round(
                        (time.monotonic() - started) * 1000.0, 1
                    ),
                    "waited_ms": round(waited_ms, 1),
                    "finished_wall": time.time(),
                    "outcome": _current["outcome"],
                }
                if "cancellation" in _current:
                    _last["cancellation"] = dict(_current["cancellation"])
                _current = None
        COM_LOCK.release()


def _mark_outcome(
    thread_id: int | None,
    outcome: str,
    *,
    cancellation: dict | None = None,
) -> bool:
    """Record an externally observed outcome for one serialized worker.

    A timeout is decided by the caller while the worker owns the COM lock,
    so the context manager cannot infer it from the worker's eventual return
    or exception.  Match the worker thread exactly; a stale or absent id is
    never allowed to annotate a different operation.  The completed-record
    leg closes the small race where ending the owned process releases the
    worker before the timeout caller can attach the cancellation details.
    """
    if thread_id is None:
        return False
    with _state_lock:
        target = None
        if _current is not None and _current.get("thread") == thread_id:
            target = _current
        elif _last is not None and _last.get("thread") == thread_id:
            target = _last
        if target is None:
            return False
        target["outcome"] = outcome
        if cancellation is not None:
            target["cancellation"] = dict(cancellation)
        return True


def serialized(name: str):
    """Decorator form of com_operation for whole-function COM operations.
    Marks the function so the coverage audit test can verify every COM
    entry point takes the lock."""

    def deco(fn):
        import functools

        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            with com_operation(name):
                return fn(*args, **kwargs)

        wrapper._com_serialized = name
        return wrapper

    return deco


def lock_snapshot() -> dict:
    """Contention report for the status tools: is a COM operation running
    right now, what is it, and how long has it held the lock. last_op
    carries the previous operation's duration, queue wait, and outcome."""
    with _state_lock:
        out: dict = {
            "held": _current is not None,
            "ops_serialized": _serialized_total,
        }
        if _current is not None:
            out["current_op"] = {
                "name": _current["name"],
                "running_ms": round(
                    (time.monotonic() - _current["started_mono"]) * 1000.0, 1
                ),
                "outcome": _current["outcome"],
            }
            if "cancellation" in _current:
                out["current_op"]["cancellation"] = dict(
                    _current["cancellation"]
                )
        if _last is not None:
            out["last_op"] = {
                "name": _last["name"],
                "duration_ms": _last["duration_ms"],
                "waited_ms": _last["waited_ms"],
                "outcome": _last["outcome"],
            }
            if "cancellation" in _last:
                out["last_op"]["cancellation"] = dict(_last["cancellation"])
        return out


def acquire(timeout: float) -> bool:
    """Bounded acquisition for callers that must stay responsive (the
    status tools). Pair with release() only when this returns True."""
    return COM_LOCK.acquire(timeout=timeout)


def release() -> None:
    COM_LOCK.release()
