"""Bounded, non-blocking delivery of the stdio server's stderr.

Some MCP clients pipe a server's stderr and never read it.  A synchronous
write then fills that pipe and can stop the JSON-RPC event loop.  Moving the
write to an ordinary ``sys.stderr`` writer thread is still unsafe: the stuck
thread owns the stream lock at interpreter shutdown, and an inherited handle
to the full client pipe can also stall child startup.

The installed transport therefore redirects descriptor 2 to an internal pipe
that is always drained.  Log records and raw stderr lines enter one bounded
queue without waiting.  A single daemon writer uses ``os.write`` on a private,
non-inheritable duplicate of the client's original descriptor.  Queue overflow
is counted.  Exit draining is bounded, and if the internal pump stops,
descriptor 2 and the Windows standard-error handle are handed back to the
client so writers cannot remain pointed at an undrained internal pipe.
"""

from __future__ import annotations

import atexit
import logging
import os
import queue
import re
import sys
import threading
import time
import traceback


PREFIXES = (
    "Failed to validate request: ",
    "Failed to validate notification: ",
)
MAX_LINE = 240
RAW_LINE = 1024
QUEUE_LINES = 256
STALL_S = 0.25
DRAIN_S = 1.0
DRAIN_MARK = b"\x1b[ks4ppt-drain]\x1b"

_METHOD = re.compile(r"input_value='([^']{1,80})'")


def _clip(text: str, limit: int) -> str:
    """Return at most ``limit`` UTF-8 bytes, on a character boundary."""
    data = text.encode("utf-8")
    if len(data) <= limit:
        return text
    return data[: limit - 3].decode("utf-8", "ignore") + "..."


class SdkValidationDump(logging.Filter):
    """Collapse the SDK's multi-line validation warning to one short line."""

    def filter(self, record: logging.LogRecord) -> bool:
        try:
            text = record.getMessage()
        except Exception:  # noqa: BLE001
            return True
        for prefix in PREFIXES:
            if text.startswith(prefix):
                break
        else:
            return True
        body = text[len(prefix) :]
        line = prefix + body.split("\n", 1)[0].strip()
        found = _METHOD.search(body)
        if found:
            line += f" (method {found.group(1)!r})"
        record.msg = line
        record.args = ()
        return True


class OneLine(logging.Formatter):
    """Format one clipped line without exception or stack trace text."""

    def __init__(self) -> None:
        super().__init__("%(levelname)s:%(name)s:%(message)s")

    def format(self, record: logging.LogRecord) -> str:
        try:
            record.message = record.getMessage()
            text = self.formatMessage(record)
        except Exception:  # noqa: BLE001
            text = f"{record.levelname}:{record.name}:(unformattable record)"
        text = " | ".join(
            part.strip() for part in text.splitlines() if part.strip()
        )
        return _clip(text, MAX_LINE - 1)


class NonBlockingStderr:
    """A bounded logging sink and its sole client-stderr writer."""

    def __init__(
        self,
        stream=None,
        maxsize: int = QUEUE_LINES,
        *,
        fd: int | None = None,
    ) -> None:
        self.queue: queue.Queue[str] = queue.Queue(maxsize)
        self.stream = stream
        self.fd = fd
        self.dropped = 0
        self._reported = 0
        self._drop_lock = threading.Lock()
        self.busy_since: float | None = None
        self.pump_alive = False
        self.drained = threading.Event()
        owner = self

        class _Handler(logging.Handler):
            def emit(self, record: logging.LogRecord) -> None:
                try:
                    line = self.format(record) + "\n"
                except Exception:  # noqa: BLE001
                    owner.drop()
                    return
                owner.offer(line)

            def handleError(self, record) -> None:  # noqa: N802
                owner.drop()

        self.handler = _Handler()
        self.handler.setFormatter(OneLine())
        self.handler.addFilter(SdkValidationDump())
        self.thread = threading.Thread(
            target=self._drain,
            daemon=True,
            name="ks4ppt-stderr-writer",
        )
        self.thread.start()

    def drop(self) -> None:
        with self._drop_lock:
            self.dropped += 1

    def offer(self, line: str) -> None:
        """Queue one line or drop and count it, without waiting."""
        try:
            self.queue.put_nowait(line)
        except queue.Full:
            self.drop()

    def _write(self, text: str) -> None:
        if self.fd is not None:
            data = text.replace("\n", os.linesep).encode("utf-8", "replace")
            while data:
                data = data[os.write(self.fd, data) :]
            return
        stream = self.stream if self.stream is not None else sys.stderr
        stream.write(text)
        stream.flush()

    def _drain(self) -> None:
        while True:
            line = self.queue.get()
            self.busy_since = time.monotonic()
            try:
                with self._drop_lock:
                    lost = self.dropped - self._reported
                    self._reported += lost
                if lost:
                    self._write(
                        f"WARNING:kitchensink4ppt:{lost} log lines dropped "
                        "while stderr was not being read\n"
                    )
                self._write(line)
            except Exception:  # noqa: BLE001
                pass
            finally:
                self.busy_since = None
                # unfinished_tasks includes this line until its write ends.
                self.queue.task_done()

    def flush(self, bound: float = DRAIN_S, stall: float = STALL_S) -> None:
        """Boundedly wait for queued and in-hand lines to finish writing."""
        end = time.monotonic() + bound
        while self.queue.unfinished_tasks and time.monotonic() < end:
            since = self.busy_since
            if since is not None and time.monotonic() - since > stall:
                return
            time.sleep(0.02)

    def pump(self, rfd: int, on_exit=None) -> None:
        """Drain redirected descriptor 2; hand it back if this pump stops."""
        try:
            self._pump_loop(rfd)
        finally:
            try:
                os.close(rfd)
            except OSError:
                pass
            self.pump_alive = False
            self.drained.set()
            if on_exit is not None:
                try:
                    on_exit(self)
                except Exception:  # noqa: BLE001
                    pass

    def _pump_loop(self, rfd: int) -> None:
        pending = b""
        while True:
            try:
                chunk = os.read(rfd, 65536)
            except OSError:
                if pending:
                    self._raw(pending)
                return
            if not chunk:
                if pending:
                    self._raw(pending)
                return
            try:
                pending = self._lines(pending + chunk)
            except Exception:  # noqa: BLE001
                pending = b""
                self.drop()

    def _lines(self, data: bytes) -> bytes:
        """Queue complete lines and return the unfinished tail."""
        *complete, tail = data.split(b"\n")
        for line in complete:
            if line.endswith(DRAIN_MARK):
                if line[: -len(DRAIN_MARK)]:
                    self._raw(line[: -len(DRAIN_MARK)])
                self.drained.set()
                continue
            self._raw(line)
        while len(tail) >= RAW_LINE:
            self._raw(tail[:RAW_LINE])
            tail = tail[RAW_LINE:]
        return tail

    def _raw(self, data: bytes) -> None:
        text = data.decode("utf-8", "replace").rstrip("\r")
        self.offer(_clip(text, RAW_LINE - 1) + "\n")


def _log_unraisable(unraisable) -> None:
    exc = unraisable.exc_value
    try:
        where = repr(unraisable.object)
    except Exception:  # noqa: BLE001
        where = "an object"
    logging.getLogger("kitchensink4ppt").warning(
        "%s %s: %s: %s",
        unraisable.err_msg or "Exception ignored in",
        where,
        type(exc).__name__ if exc is not None else "unknown",
        exc,
    )


def _log_thread_exception(args) -> None:
    if args.exc_type is SystemExit:
        return
    place = ""
    try:
        last = traceback.extract_tb(args.exc_traceback)[-1]
        place = (
            f" (at {os.path.basename(last.filename)}:{last.lineno} "
            f"in {last.name})"
        )
    except Exception:  # noqa: BLE001
        pass
    name = args.thread.name if args.thread is not None else "a thread"
    logging.getLogger("kitchensink4ppt").error(
        "Exception in thread %s: %s: %s%s",
        name,
        args.exc_type.__name__,
        args.exc_value,
        place,
    )


def _set_std_error_handle(fd: int) -> None:
    """Point the Windows standard-error handle at ``fd`` for children."""
    if os.name != "nt":
        return
    import ctypes
    import msvcrt
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.SetStdHandle.argtypes = (wintypes.DWORD, wintypes.HANDLE)
    kernel32.SetStdHandle.restype = wintypes.BOOL
    changed = kernel32.SetStdHandle(
        wintypes.DWORD(-12 & 0xFFFFFFFF),
        wintypes.HANDLE(msvcrt.get_osfhandle(fd)),
    )
    if not changed:
        raise ctypes.WinError(ctypes.get_last_error())


def _redirect_descriptor_2() -> tuple[int, int] | None:
    """Return ``(original stderr duplicate, internal pipe read end)``."""
    try:
        if sys.stderr is not None:
            sys.stderr.flush()
        original = os.dup(2)
    except (OSError, ValueError, AttributeError):
        return None
    read_end: int | None = None
    try:
        read_end, write_end = os.pipe()
        os.dup2(write_end, 2)
        os.close(write_end)
        _set_std_error_handle(2)
    except Exception:  # noqa: BLE001
        try:
            os.dup2(original, 2)
            _set_std_error_handle(2)
        except Exception:  # noqa: BLE001
            pass
        if read_end is not None:
            os.close(read_end)
        os.close(original)
        return None
    return original, read_end


_WRITER: NonBlockingStderr | None = None
_REPLACED: list[tuple[logging.Logger, logging.Handler]] = []
_HOOKS: dict = {}
_ORIGINAL_FD: int | None = None
_LOGGERS = ("", "fastmcp")
_LOCK = threading.RLock()
_ATEXIT = False


def _restore_descriptor_2(writer: NonBlockingStderr) -> None:
    """Restore the client's stderr when the installed pump terminates."""
    with _LOCK:
        if _WRITER is not writer or _ORIGINAL_FD is None:
            return
        try:
            os.dup2(_ORIGINAL_FD, 2)
            _set_std_error_handle(2)
        except Exception:  # noqa: BLE001
            return
    writer.offer(
        "WARNING:kitchensink4ppt:the stderr pump stopped; stderr is written "
        "directly again\n"
    )


def drain(bound: float = DRAIN_S) -> None:
    """Boundedly pass raw stderr through the pump and empty the queue."""
    writer = _WRITER
    if writer is None:
        return
    end = time.monotonic() + bound
    try:
        if sys.stderr is not None:
            sys.stderr.flush()
    except Exception:  # noqa: BLE001
        pass
    if writer.pump_alive:
        writer.drained.clear()
        try:
            os.write(2, DRAIN_MARK + b"\n")
            writer.drained.wait(max(0.0, end - time.monotonic()))
        except OSError:
            pass
    writer.flush(max(0.0, end - time.monotonic()))


def install(stream=None) -> NonBlockingStderr:
    """Install the bounded stderr path and return its writer; idempotent."""
    global _ATEXIT, _ORIGINAL_FD, _WRITER
    with _LOCK:
        if _WRITER is None:
            redirected = _redirect_descriptor_2() if stream is None else None
            if redirected is not None:
                _ORIGINAL_FD, read_end = redirected
                _WRITER = NonBlockingStderr(fd=_ORIGINAL_FD)
                _WRITER.pump_alive = True
                threading.Thread(
                    target=_WRITER.pump,
                    args=(read_end, _restore_descriptor_2),
                    daemon=True,
                    name="ks4ppt-stderr-pump",
                ).start()
                if not _ATEXIT:
                    atexit.register(drain)
                    _ATEXIT = True
            else:
                _WRITER = NonBlockingStderr(stream)
        writer = _WRITER
    for name in _LOGGERS:
        logger = logging.getLogger(name)
        for handler in list(logger.handlers):
            if handler is not writer.handler:
                logger.removeHandler(handler)
                _REPLACED.append((logger, handler))
        if writer.handler not in logger.handlers:
            logger.addHandler(writer.handler)
    logging.captureWarnings(True)
    if not _HOOKS:
        _HOOKS["unraisable"] = sys.unraisablehook
        _HOOKS["thread"] = threading.excepthook
        sys.unraisablehook = _log_unraisable
        threading.excepthook = _log_thread_exception
    return writer


def installed() -> NonBlockingStderr | None:
    return _WRITER


def uninstall() -> None:
    """Undo installation for tests.  The production server never calls it."""
    with _LOCK:
        _uninstall_locked()


def _uninstall_locked() -> None:
    global _ORIGINAL_FD, _WRITER
    if _WRITER is None:
        return
    for name in _LOGGERS:
        logging.getLogger(name).removeHandler(_WRITER.handler)
    for logger, handler in _REPLACED:
        if handler not in logger.handlers:
            logger.addHandler(handler)
    _REPLACED.clear()
    logging.captureWarnings(False)
    if _HOOKS:
        sys.unraisablehook = _HOOKS.pop("unraisable")
        threading.excepthook = _HOOKS.pop("thread")
    if _ORIGINAL_FD is not None:
        try:
            os.dup2(_ORIGINAL_FD, 2)
            _set_std_error_handle(2)
        except Exception:  # noqa: BLE001
            pass
        _ORIGINAL_FD = None
    _WRITER = None
