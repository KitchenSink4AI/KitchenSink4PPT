"""One wall-clock deadline per operation, with a no-late-publish handshake."""

from __future__ import annotations

import contextvars
import math
import threading
import time
from contextlib import contextmanager
from typing import Iterator

from .errors import OpError
from .messages import TIMEOUT_MESSAGE


class Deadline:
    """A wall-clock budget whose cancel and publish decisions are exclusive."""

    def __init__(self, seconds: float | None) -> None:
        self.bound = seconds
        self._expires = (
            math.inf if seconds is None else time.monotonic() + float(seconds)
        )
        self._lock = threading.Lock()
        self._cancelled = False
        self._committed = False

    @property
    def bounded(self) -> bool:
        return self.bound is not None

    @property
    def cancelled(self) -> bool:
        return self._cancelled

    @property
    def committed(self) -> bool:
        return self._committed

    def remaining(self) -> float:
        if self._cancelled:
            return 0.0
        return max(0.0, self._expires - time.monotonic())

    def expired(self) -> bool:
        return self._cancelled or self.remaining() <= 0.0

    def wait_bound(self, cap: float) -> float:
        return max(0.0, min(float(cap), self.remaining()))

    def cancel(self) -> bool:
        """Return false when the publish point already committed."""
        with self._lock:
            if self._committed:
                return False
            self._cancelled = True
            return True

    def try_commit(self) -> bool:
        """Claim the publish point unless the caller already timed out."""
        with self._lock:
            if self._committed:
                return True
            if self._cancelled or time.monotonic() >= self._expires:
                return False
            self._committed = True
            return True

    def timeout_error(self, what: str = "the call") -> OpError:
        bound = f"{self.bound:g}s" if self.bounded else "configured"
        return OpError(
            "TIMEOUT",
            TIMEOUT_MESSAGE.format(what=what, bound=bound),
            details={"bound_s": self.bound, "saved": False},
        )

    def check(self, what: str = "the call") -> None:
        if self.expired():
            raise self.timeout_error(what)

    def commit_or_raise(self, what: str = "the call") -> None:
        if not self.try_commit():
            raise self.timeout_error(what)


_CURRENT: contextvars.ContextVar[Deadline | None] = contextvars.ContextVar(
    "ks4ppt_deadline", default=None
)


def current() -> Deadline:
    active_deadline = _CURRENT.get()
    return active_deadline if active_deadline is not None else Deadline(None)


@contextmanager
def active(value: Deadline) -> Iterator[Deadline]:
    token = _CURRENT.set(value)
    try:
        yield value
    finally:
        _CURRENT.reset(token)


def resolve(timeout: float | Deadline | None) -> Deadline:
    if isinstance(timeout, Deadline):
        return timeout
    if timeout is None:
        return current()
    return Deadline(float(timeout))
