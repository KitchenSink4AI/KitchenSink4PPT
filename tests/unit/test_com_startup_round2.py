"""Round 2: the second review's findings on the COM start-up path.

M3  the retry set was wider than the recovery action could justify, and a
    failed attempt that had already started a PowerPoint was not accounted
    for;
N1  an apartment that genuinely could not be armed was recorded and then
    ignored, so Dispatch spoke for it with a vaguer error;
N2  the caller-apartment evidence was claimed to be on the result and is
    not: it is an in-process diagnostic, reported through powerpoint_status.

No PowerPoint is started by anything in this file.
"""

from __future__ import annotations

import sys

import pytest

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="pywin32 com_error is Windows only"
)

from kitchensink4ppt.core.errors import PowerPointNotRunning  # noqa: E402


def _com_error(hresult: int):
    import pywintypes

    return pywintypes.com_error(hresult, "fabricated for the test", None, None)


@pytest.fixture
def bridge():
    from kitchensink4ppt.com import bridge as _b

    return _b


class _FakePythoncom:
    COINIT_APARTMENTTHREADED = 2

    def __init__(self, raises=None):
        self._raises = raises
        self.calls = 0

    def CoInitializeEx(self, flags):  # noqa: N802 - COM spelling
        self.calls += 1
        if self._raises is not None:
            raise self._raises


class _FakeWin32:
    """DispatchEx that can START a process before it fails, which is the
    sequence the old fakes could not represent."""

    def __init__(self, script, world):
        self.script = list(script)
        self.world = world
        self.calls = 0

    def DispatchEx(self, progid):  # noqa: N802 - COM spelling
        self.calls += 1
        step = self.script.pop(0) if self.script else None
        if step is None:
            return f"app:{progid}"
        exc, spawns = step
        if spawns:
            self.world["pids"] |= set(spawns)
        if exc is None:
            return f"app:{progid}"
        raise exc


# ------------------------------------------------- M3: the retry set


def test_only_the_recoverable_hresult_is_retried(bridge):
    """The only action between attempts is re-arming the apartment, so the
    retry set is exactly what that repairs."""
    assert bridge.CO_E_NOTINITIALIZED in bridge.STARTUP_RETRY_HRESULTS
    for name in ("RPC_E_CHANGED_MODE", "REGDB_E_CLASSNOTREG",
                 "CO_E_SERVER_EXEC_FAILURE"):
        assert getattr(bridge, name) not in bridge.STARTUP_RETRY_HRESULTS, (
            f"{name} is not repaired by re-running CoInitializeEx"
        )
    assert bridge.STARTUP_RETRY_HRESULTS <= bridge.STARTUP_HRESULTS


@pytest.mark.parametrize(
    "name", ["RPC_E_CHANGED_MODE", "CO_E_SERVER_EXEC_FAILURE"]
)
def test_the_unrecoverable_startup_faults_are_tried_once(bridge, name):
    world = {"pids": set()}
    win32 = _FakeWin32([(_com_error(getattr(bridge, name)), ())] * 2, world)
    with pytest.raises(PowerPointNotRunning):
        bridge._start_powerpoint(win32, _FakePythoncom(), set())
    assert win32.calls == 1


def test_a_failed_attempt_that_started_powerpoint_is_reported_not_ended(
    bridge, monkeypatch
):
    """M3's concrete failure, ruled on again in round 3 (G2).

    Attempt one starts POWERPNT.EXE and then fails before a proxy reaches
    Python. Round 2 ended that pid, reasoning that an empty entry snapshot
    made it ours. An empty snapshot is not ownership: the user can launch
    PowerPoint in that same window, and `tasklist` failing produces the
    same empty set. So the pid is reported and left alone, and the retry
    stops rather than launching a second process on top of it.
    """
    world = {"pids": set()}
    killed: list[int] = []

    monkeypatch.setattr(bridge, "powerpnt_pids", lambda: set(world["pids"]))

    def fake_run(cmd, **kwargs):
        if cmd[0] == "taskkill":
            pid = int(cmd[2])
            killed.append(pid)
            world["pids"].discard(pid)

        class R:
            returncode = 0

        return R()

    monkeypatch.setattr(bridge.subprocess, "run", fake_run)

    win32 = _FakeWin32(
        [
            # attempt 1: starts pid 4242, then fails with the one fault we
            # would otherwise retry
            (_com_error(bridge.CO_E_NOTINITIALIZED), (4242,)),
            (None, (4343,)),
        ],
        world,
    )
    with pytest.raises(PowerPointNotRunning) as exc_info:
        bridge._start_powerpoint(win32, _FakePythoncom(), set())
    assert killed == [], "nothing on this path may end a process"
    assert win32.calls == 1, "a process appeared, so the retry must stop"
    assert world["pids"] == {4242}, "the process that appeared survives"
    assert "4242" in str(exc_info.value)


def test_a_pre_existing_powerpoint_is_never_touched(bridge, monkeypatch):
    """The cleanup may only end what THIS call started."""
    world = {"pids": {999}}  # the user's PowerPoint
    killed: list[int] = []
    monkeypatch.setattr(bridge, "powerpnt_pids", lambda: set(world["pids"]))

    def fake_run(cmd, **kwargs):
        if cmd[0] == "taskkill":
            killed.append(int(cmd[2]))

        class R:
            returncode = 0

        return R()

    monkeypatch.setattr(bridge.subprocess, "run", fake_run)
    win32 = _FakeWin32(
        [(_com_error(bridge.CO_E_NOTINITIALIZED), ()), (None, ())], world
    )
    bridge._start_powerpoint(win32, _FakePythoncom(), {999})
    assert killed == [], "the user's PowerPoint must never be ended"
    assert world["pids"] == {999}


def test_a_stranded_process_is_reported_when_the_call_still_fails(
    bridge, monkeypatch
):
    """Round 3 (G2): the refusal names the pid it saw. It may not claim to
    have ended anything, because it does not, and since R7-2 it may not
    claim the process is still running either: the snapshot was taken
    while the start was failing and says nothing about now."""
    world = {"pids": set()}
    monkeypatch.setattr(bridge, "powerpnt_pids", lambda: set(world["pids"]))

    def fake_run(cmd, **kwargs):
        if cmd[0] == "taskkill":
            world["pids"].discard(int(cmd[2]))

        class R:
            returncode = 0

        return R()

    monkeypatch.setattr(bridge.subprocess, "run", fake_run)
    win32 = _FakeWin32(
        [(_com_error(bridge.CO_E_SERVER_EXEC_FAILURE), (7777,))], world
    )
    with pytest.raises(PowerPointNotRunning) as exc_info:
        bridge._start_powerpoint(win32, _FakePythoncom(), set())
    message = str(exc_info.value)
    assert "was ended" not in message
    assert "left running" not in message
    assert (
        "A PowerPoint process appeared during the failed start: pid 7777. "
        "This call did not force-end it; its current state was not "
        "re-checked."
    ) in message
    assert world["pids"] == {7777}, "the process is left exactly as found"


def test_server_exec_failure_does_not_claim_nothing_was_opened(bridge):
    """"Nothing was opened" is not knowable for a server that may have
    partially launched."""
    msg = str(bridge._classify(_com_error(bridge.CO_E_SERVER_EXEC_FAILURE)))
    assert "Nothing was opened and nothing was changed." not in msg
    assert "No presentation was opened" in msg
    # The apartment faults, where it IS knowable, keep the strong sentence.
    other = str(bridge._classify(_com_error(bridge.CO_E_NOTINITIALIZED)))
    assert "Nothing was opened and nothing was changed." in other


# ---------------------------------------- N1: refuse before Dispatch


def test_an_unarmable_apartment_refuses_before_any_dispatch(
    bridge, monkeypatch
):
    """A real CoInitializeEx failure is not something Dispatch can repair,
    and letting it try means a second, vaguer error speaks for the first."""
    dispatched: list[str] = []

    class _Win32:
        def DispatchEx(self, progid):  # noqa: N802
            dispatched.append(progid)
            return "app"

    monkeypatch.setattr(
        bridge, "_com_modules", lambda: (_FakePythoncom(), _Win32())
    )
    monkeypatch.setattr(
        bridge, "_ensure_apartment", lambda _p: "failed: E_OUTOFMEMORY"
    )
    gen = bridge._powerpoint_locked()
    with pytest.raises(PowerPointNotRunning) as exc_info:
        next(gen)
    assert dispatched == [], "no COM call may be attempted"
    assert "E_OUTOFMEMORY" in str(exc_info.value)
    assert "could not be initialized" in str(exc_info.value)


# ------------------------------- N2: the evidence is where it is claimed


def test_the_apartment_state_is_an_in_process_diagnostic(bridge):
    """It is recorded per thread and surfaced by powerpoint_status. It is
    deliberately NOT injected into operation results, and the comment that
    said otherwise was the defect."""
    import inspect
    import threading

    bridge._APARTMENT_STATE[threading.get_ident()] = "initialized"
    assert bridge.apartment_state() == "initialized"

    src = inspect.getsource(bridge._run_bounded)
    assert "carried now, on the result" not in src, (
        "the comment claimed caller-visible evidence that does not exist"
    )
    assert "IN-PROCESS diagnostic" in src

    # _run_bounded still returns only the operation's own value.
    out = bridge._run_bounded("t", 30.0, lambda: {"ok": True})
    assert out == {"ok": True}
    assert "caller_apartment" not in out


def test_powerpoint_status_is_where_the_apartment_is_reported(bridge):
    import inspect

    src = inspect.getsource(bridge.powerpoint_status)
    assert 'out["com_apartment"] = _ensure_apartment' in src


def test_cleanup_is_disarmed_when_powerpoint_was_already_running(
    bridge, monkeypatch
):
    """The rule the module has always promised: nothing here ends a process
    it did not start. If an instance was up at entry we ATTACHED, so a
    process appearing mid-call is not ours to reason about.

    This is not hypothetical. Writing the M3 test on a machine where another
    session had PowerPoint open, an earlier version of the fix read the real
    process table against an empty baseline and tried to end that session's
    instance.
    """
    world = {"pids": {111, 222}}  # somebody else's PowerPoint, plus a second
    killed: list[int] = []

    def table():
        # A third instance appears mid-call; it is NOT ours.
        world["pids"] = world["pids"] | {333}
        return set(world["pids"])

    monkeypatch.setattr(bridge, "powerpnt_pids", table)

    def fake_run(cmd, **kwargs):
        if cmd[0] == "taskkill":
            killed.append(int(cmd[2]))

        class R:
            returncode = 0

        return R()

    monkeypatch.setattr(bridge.subprocess, "run", fake_run)
    win32 = _FakeWin32(
        [(_com_error(bridge.CO_E_NOTINITIALIZED), ()), (None, ())],
        {"pids": set()},
    )
    app, ended = bridge._start_powerpoint(win32, _FakePythoncom(), {111, 222})
    assert app == "app:PowerPoint.Application"
    assert killed == [], "no process may be ended when we attached"
    assert ended == set()


def test_an_older_pythoncom_without_coinitializeex_still_works(bridge):
    """A MISSING API is not an apartment failure. Reporting it as one would
    turn an old pywin32 into "PowerPoint could not be started" on every
    call, which is the opposite of what N1 asked for."""

    class _Old:
        """pythoncom as it was before CoInitializeEx."""

        def __init__(self):
            self.calls = 0

        def CoInitialize(self):  # noqa: N802 - COM spelling
            self.calls += 1

    old = _Old()
    assert not hasattr(old, "CoInitializeEx")
    assert bridge._ensure_apartment(old) == "initialized"
    assert old.calls == 1
