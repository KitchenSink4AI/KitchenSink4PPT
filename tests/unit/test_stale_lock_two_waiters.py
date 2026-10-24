"""Two waiters on a dead holder's lockfile: one of them wins, never both.

Three servers wait on a deck whose ``write.lock`` was left by a server that
died. Two of them read the dead lock and judge it stale. B breaks it and
publishes its own lock; C, still holding its verdict from before, must find
B's live lock at the path and leave it alone. Before breaking went through
``lockrelease.break_stale``, C deleted whatever was at the path, B's live lock
included, and published too: B and C were both inside the lock and their
saves overwrote each other.

The interleaving is pinned with two real processes by holding C up between
its stale verdict and its break. The same is checked on the live route's
application lock (com/xproc.py), whose holder drives PowerPoint, with the
lock directory redirected so the machine-wide lock is never touched.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path

import pytest

NO_WINDOW = 0x08000000 if os.name == "nt" else 0

# C judges the dead lock stale, is held up before it breaks it until B is
# inside its own hold, then takes the lock and reports whether B was inside.
_C = r"""
import json, sys, time
from pathlib import Path
work = Path(sys.argv[1])
@IMPORT@
real = m._is_stale
def slow(info):
    r = real(info)
    if r and not (work / "c_judged").exists():
        (work / "c_judged").write_text("1")          # C has judged the dead lock stale ...
        end = time.time() + 10
        while not (work / "b_inside").exists() and time.time() < end:
            time.sleep(0.005)                         # ... and is held up before it breaks it
    return r
m._is_stale = slow
with @ENTER_C@:
    both = (work / "b_inside").exists() and not (work / "b_left").exists()
    (work / "c_result").write_text(json.dumps({"entered_while_b_inside": both}))
    (work / "c_inside").write_text("1")
    time.sleep(0.3)
"""

# B waits until C has judged, then takes the lock and stays inside until C
# has entered (the defect) or the wait runs out.
_B = r"""
import sys, time
from pathlib import Path
work = Path(sys.argv[1])
@IMPORT@
end = time.time() + 10
while not (work / "c_judged").exists() and time.time() < end:
    time.sleep(0.005)
with @ENTER_B@:
    (work / "b_inside").write_text("1")
    end = time.time() + 3
    while not (work / "c_inside").exists() and time.time() < end:
        time.sleep(0.005)
    (work / "b_left").write_text("1")
"""


def _children(module_import: str, enter_c: str, enter_b: str) -> tuple[str, str]:
    return (_C.replace("@IMPORT@", module_import).replace("@ENTER_C@", enter_c),
            _B.replace("@IMPORT@", module_import).replace("@ENTER_B@", enter_b))


def _created(pid: int):
    import ctypes
    import ctypes.wintypes as wt
    k32 = ctypes.WinDLL("kernel32")
    h = k32.OpenProcess(0x1000, False, pid)
    if not h:
        return None
    c, e, kt, ut = wt.FILETIME(), wt.FILETIME(), wt.FILETIME(), wt.FILETIME()
    k32.GetProcessTimes(h, ctypes.byref(c), ctypes.byref(e), ctypes.byref(kt), ctypes.byref(ut))
    k32.CloseHandle(h)
    return float(c.dwHighDateTime << 32 | c.dwLowDateTime)


def _race(tmp_path: Path, lock_path: Path, child_c: str, child_b: str, env: dict) -> dict:
    # A lock left by a process that has exited. Its handle stays open here, so
    # Windows cannot recycle the PID while the test runs.
    dead = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(0.3)"],
                            creationflags=NO_WINDOW, stdin=subprocess.DEVNULL)
    ctime = _created(dead.pid)
    dead.wait()
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    lock_path.write_text(json.dumps({
        "pid": dead.pid, "token": f"{dead.pid}-dead", "pid_created": ctime, "time": time.time(),
        "holder": "a server that died", "host": os.environ.get("COMPUTERNAME", "")}), encoding="utf-8")
    c = subprocess.Popen([sys.executable, "-c", child_c, str(tmp_path)], env=env,
                         creationflags=NO_WINDOW, stdin=subprocess.DEVNULL)
    b = subprocess.Popen([sys.executable, "-c", child_b, str(tmp_path)], env=env,
                         creationflags=NO_WINDOW, stdin=subprocess.DEVNULL)
    try:
        c.wait(timeout=60)
        b.wait(timeout=60)
    finally:
        for p in (b, c):
            if p.poll() is None:
                p.kill()
    return json.loads((tmp_path / "c_result").read_text())


@pytest.mark.skipif(os.name != "nt", reason="the PID checks are Windows-specific")
def test_breaking_a_dead_decks_lock_never_admits_a_second_writer(tmp_path):
    (tmp_path / "deck.pptx").write_bytes(b"")
    from kitchensink4ppt.core import safesave
    lock = safesave.slot_dir(tmp_path / "deck.pptx", create=True) / safesave.LOCK_FILE_NAME
    enter = 'm.write_lock(work / "deck.pptx")'
    child_c, child_b = _children("from kitchensink4ppt.core import safesave as m", enter, enter)
    result = _race(tmp_path, lock, child_c, child_b, dict(os.environ))
    assert result["entered_while_b_inside"] is False, "two writers inside write_lock at once"


@pytest.mark.skipif(os.name != "nt", reason="the PID checks are Windows-specific")
def test_breaking_a_dead_app_lock_never_admits_a_second_holder(tmp_path):
    live = tmp_path / "live"
    env = dict(os.environ, KS4P_LIVE_LOCK_DIR=str(live))
    from kitchensink4ppt.com import xproc
    child_c, child_b = _children("from kitchensink4ppt.com import xproc as m",
                                 'm.cross_process_lock("C")', 'm.cross_process_lock("B")')
    result = _race(tmp_path, live / f"{xproc.APP_SCOPE}.lock", child_c, child_b, env)
    assert result["entered_while_b_inside"] is False, "two holders inside the live app lock at once"
