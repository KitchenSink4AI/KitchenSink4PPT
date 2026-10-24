"""Focused exit and pump-fallback checks for the bounded stderr path."""

from __future__ import annotations

import io
import json
import os
import subprocess
import sys
import textwrap
import threading
import time
from pathlib import Path

import pytest

from kitchensink4ppt import stderrlog


NOWIN = getattr(subprocess, "CREATE_NO_WINDOW", 0)
SRC = str(Path(stderrlog.__file__).resolve().parent.parent)


class _HeldWrite:
    def __init__(self) -> None:
        self.entered = threading.Event()
        self.release = threading.Event()
        self.text = ""

    def write(self, text: str) -> int:
        self.entered.set()
        self.release.wait()
        self.text += text
        return len(text)

    def flush(self) -> None:
        pass


def test_flush_counts_the_line_already_in_the_writers_hand():
    """An empty queue does not mean its last dequeued line was written."""
    stream = _HeldWrite()
    writer = stderrlog.NonBlockingStderr(stream=stream)
    writer.offer("last line\n")
    assert stream.entered.wait(5)
    assert writer.queue.empty()
    assert writer.queue.unfinished_tasks == 1

    finished = threading.Event()

    def flush() -> None:
        writer.flush(bound=5, stall=5)
        finished.set()

    thread = threading.Thread(target=flush, daemon=True)
    thread.start()
    assert not finished.wait(0.2), "flush ignored the line held by the writer"
    stream.release.set()
    assert finished.wait(5)
    assert stream.text == "last line\n"


def test_pump_forwards_the_final_unterminated_line_before_handback():
    """The raw stderr tail in hand at pipe EOF is part of the exit drain."""
    stream = io.StringIO()
    writer = stderrlog.NonBlockingStderr(stream=stream)
    read_end, write_end = os.pipe()
    pump = threading.Thread(target=writer.pump, args=(read_end,), daemon=True)
    pump.start()
    os.write(write_end, b"final line without newline")
    os.close(write_end)
    pump.join(5)
    assert not pump.is_alive()
    writer.flush(5)
    assert stream.getvalue() == "final line without newline\n"


@pytest.mark.timeout(60)
def test_atexit_delivers_the_last_traceback_line_to_a_reader():
    """A readable client receives the exception line printed at exit."""
    env = {**os.environ, "PYTHONPATH": SRC}
    code = (
        "from kitchensink4ppt import stderrlog\n"
        "stderrlog.install()\n"
        "raise RuntimeError('PPT-STDERR-LAST-LINE')\n"
    )
    for _ in range(3):
        result = subprocess.run(
            [sys.executable, "-X", "utf8", "-c", code],
            stdin=subprocess.DEVNULL,
            capture_output=True,
            env=env,
            timeout=30,
            creationflags=NOWIN,
        )
        assert result.returncode == 1
        assert b"RuntimeError: PPT-STDERR-LAST-LINE" in result.stderr


UNREAD_WITH_CHILD = textwrap.dedent(
    r"""
    import json, os, subprocess, sys, time
    from kitchensink4ppt import stderrlog

    stderrlog.install()
    started = time.monotonic()
    print("p" * 200000, file=sys.stderr)
    os.write(2, b"r" * 200000 + b"\n")
    child = subprocess.run(
        [sys.executable, "-c",
         "import sys; sys.stderr.write('c' * 300000); sys.stderr.flush()"],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        timeout=30,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    sys.stdout.write(json.dumps({
        "child": child.returncode,
        "seconds": round(time.monotonic() - started, 2),
    }) + "\n")
    sys.stdout.flush()
    """
)


@pytest.mark.timeout(120)
def test_raw_writes_and_child_startup_do_not_wait_on_unread_stderr():
    """An inherited handle is the drained pipe, not the stuck client."""
    env = {**os.environ, "PYTHONPATH": SRC}
    proc = subprocess.Popen(
        [sys.executable, "-X", "utf8", "-c", UNREAD_WITH_CHILD],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=env,
        creationflags=NOWIN,
    )
    try:
        output = proc.stdout.readline()
        assert output, "raw stderr or inherited child startup blocked"
        result = json.loads(output)
        assert result["child"] == 0
        assert result["seconds"] < 30, result
        assert proc.wait(timeout=30) == 0
    finally:
        if proc.poll() is None:
            proc.kill()
            proc.wait(timeout=30)


PUMP_STOPS = textwrap.dedent(
    r"""
    import json, os, sys, threading, time
    from kitchensink4ppt import stderrlog

    def stop(self, data):
        raise KeyboardInterrupt("injected pump stop")

    stderrlog.NonBlockingStderr._lines = stop
    stderrlog.install()
    os.write(2, b"trigger\n")
    end = time.monotonic() + 10
    while time.monotonic() < end and any(
            t.name == "ks4ppt-stderr-pump" for t in threading.enumerate()):
        time.sleep(0.05)
    box = {}
    def large_write():
        os.write(2, b"x" * 200000 + b"\n")
        box["done"] = True
    thread = threading.Thread(target=large_write, daemon=True)
    thread.start()
    thread.join(10)
    sys.stdout.write(json.dumps({"done": bool(box)}) + "\n")
    sys.stdout.flush()
    os._exit(0)
    """
)


@pytest.mark.timeout(120)
def test_a_stopped_pump_hands_descriptor_2_back_to_the_client():
    """A dead pump cannot leave writers on its undrained internal pipe."""
    env = {**os.environ, "PYTHONPATH": SRC}
    proc = subprocess.Popen(
        [sys.executable, "-X", "utf8", "-c", PUMP_STOPS],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=env,
        creationflags=NOWIN,
    )
    err = bytearray()

    def read_stderr() -> None:
        while True:
            chunk = proc.stderr.read1(65536)
            if not chunk:
                return
            err.extend(chunk)

    reader = threading.Thread(target=read_stderr, daemon=True)
    reader.start()
    try:
        output = proc.stdout.readline()
        assert output, "the descriptor-2 writer stayed blocked"
        assert json.loads(output) == {"done": True}
        assert proc.wait(timeout=30) == 0
        reader.join(5)
    finally:
        if proc.poll() is None:
            proc.kill()
            proc.wait(timeout=30)
    assert err.count(b"x") >= 200000, "the handed-back write was not delivered"
