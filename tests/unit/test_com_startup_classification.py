"""The COM start-up family: apartment, registration, launch.

Field report 2026-09-21 (P7): export_slide_image failed with "CoInitialize
has not been called" and the refusal carried code BAD_PARAMS, which sends a
caller to look at arguments that were never the problem. CO_E_NOTINITIALIZED
was in neither BUSY_HRESULTS nor GONE_HRESULTS, so _classify returned None,
the bridge raised a plain PptMcpError, and the server's code map turns that
into BAD_PARAMS.

These tests pin the classification, the apartment helper's two benign
outcomes, and the one bounded retry. They need no PowerPoint.
"""

from __future__ import annotations

import sys
import threading

import pytest

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="pywin32 com_error is Windows only"
)

from kitchensink4ppt.core.errors import (  # noqa: E402
    PowerPointNotRunning,
    PptMcpError,
)


def _com_error(hresult: int):
    import pywintypes

    return pywintypes.com_error(hresult, "fabricated for the test", None, None)


@pytest.fixture
def bridge():
    from kitchensink4ppt.com import bridge as _b

    return _b


# ------------------------------------------------------- classification


@pytest.mark.parametrize(
    "name",
    [
        "CO_E_NOTINITIALIZED",
        "CO_E_ALREADYINITIALIZED",
        "RPC_E_CHANGED_MODE",
        "REGDB_E_CLASSNOTREG",
        "CO_E_SERVER_EXEC_FAILURE",
    ],
)
def test_startup_hresults_are_classified(bridge, name):
    hr = getattr(bridge, name)
    typed = bridge._classify(_com_error(hr))
    assert isinstance(typed, PowerPointNotRunning), (
        f"{name} must be classified, not fall through to a plain "
        "PptMcpError (which the server maps to BAD_PARAMS)"
    )


def test_notinitialized_never_reaches_the_wire_as_bad_params(bridge):
    """The end-to-end shape of the field defect: the refusal code."""
    from kitchensink4ppt import server

    try:
        bridge._raise_classified(
            _com_error(bridge.CO_E_NOTINITIALIZED),
            "PowerPoint could not be started",
        )
    except PptMcpError as exc:
        code = server._classify(exc)
    else:  # pragma: no cover - the call above always raises
        pytest.fail("_raise_classified did not raise")
    assert code != "BAD_PARAMS"
    assert code in server.CLOSED_CODES
    assert code == "APP_NOT_RUNNING"


def test_startup_message_names_the_fault_and_the_apartment(bridge):
    bridge._APARTMENT_STATE[threading.get_ident()] = "changed-mode"
    try:
        msg = str(bridge._classify(_com_error(bridge.CO_E_NOTINITIALIZED)))
    finally:
        bridge._APARTMENT_STATE.pop(threading.get_ident(), None)
    assert "CO_E_NOTINITIALIZED" in msg
    assert "changed-mode" in msg
    assert "nothing was changed" in msg.lower()


def test_classnotreg_says_so(bridge):
    msg = str(bridge._classify(_com_error(bridge.REGDB_E_CLASSNOTREG)))
    assert "REGDB_E_CLASSNOTREG" in msg


def test_busy_and_gone_classification_is_unchanged(bridge):
    from kitchensink4ppt.core.errors import (
        PowerPointBusy,
        PowerPointDisconnected,
    )

    assert isinstance(
        bridge._classify(_com_error(bridge.RPC_E_CALL_REJECTED)),
        PowerPointBusy,
    )
    assert isinstance(
        bridge._classify(_com_error(bridge.RPC_E_DISCONNECTED)),
        PowerPointDisconnected,
    )
    assert bridge._classify(_com_error(-2147024891)) is None  # E_ACCESSDENIED


# ------------------------------------------------------- the apartment


class _FakePythoncom:
    COINIT_APARTMENTTHREADED = 2

    def __init__(self, raises=None):
        self._raises = raises
        self.calls = 0

    def CoInitializeEx(self, flags):  # noqa: N802 - COM spelling
        self.calls += 1
        if self._raises is not None:
            raise self._raises


def test_ensure_apartment_reports_initialized(bridge):
    fake = _FakePythoncom()
    assert bridge._ensure_apartment(fake) == "initialized"
    assert fake.calls == 1
    assert bridge.apartment_state() == "initialized"


def test_ensure_apartment_tolerates_already_initialized(bridge):
    """S_FALSE means the apartment was already up in this same model."""
    fake = _FakePythoncom(raises=_com_error(bridge.S_FALSE))
    assert bridge._ensure_apartment(fake) == "already"


def test_ensure_apartment_tolerates_changed_mode(bridge):
    """Another library already put this thread in the other threading
    model. Fighting it would break whatever set it."""
    fake = _FakePythoncom(raises=_com_error(bridge.RPC_E_CHANGED_MODE))
    assert bridge._ensure_apartment(fake) == "changed-mode"


def test_ensure_apartment_records_a_real_failure(bridge):
    fake = _FakePythoncom(raises=_com_error(-2147024882))  # E_OUTOFMEMORY
    state = bridge._ensure_apartment(fake)
    assert state.startswith("failed:")
    assert bridge.apartment_state().startswith("failed:")


# ------------------------------------------------------- the bounded retry


class _FakeWin32:
    def __init__(self, failures):
        self.failures = list(failures)
        self.calls = 0

    def DispatchEx(self, progid):  # noqa: N802 - COM spelling
        self.calls += 1
        if self.failures:
            raise self.failures.pop(0)
        return f"app:{progid}"


def test_start_powerpoint_retries_once_on_notinitialized(bridge, monkeypatch):
    # The real process table is off limits here: this asserts the retry, and
    # a PowerPoint another session owns must not be swept into it.
    monkeypatch.setattr(bridge, "powerpnt_pids", set)
    fake_py = _FakePythoncom()
    win32 = _FakeWin32([_com_error(bridge.CO_E_NOTINITIALIZED)])
    app, ended = bridge._start_powerpoint(win32, fake_py, set())
    assert app == "app:PowerPoint.Application"
    assert ended == set()
    assert win32.calls == 2, "the start must be retried exactly once"
    assert fake_py.calls == 1, "the apartment is re-armed before the retry"


def test_start_powerpoint_gives_up_after_the_one_retry(bridge):
    fake_py = _FakePythoncom()
    win32 = _FakeWin32(
        [_com_error(bridge.CO_E_NOTINITIALIZED)] * 2
    )
    with pytest.raises(PowerPointNotRunning):
        bridge._start_powerpoint(win32, fake_py, set())
    assert win32.calls == 2, "no unbounded retry loop"


def test_start_powerpoint_does_not_retry_an_unregistered_class(bridge):
    """A second attempt cannot register a class that is not registered."""
    fake_py = _FakePythoncom()
    win32 = _FakeWin32([_com_error(bridge.REGDB_E_CLASSNOTREG)] * 2)
    with pytest.raises(PowerPointNotRunning):
        bridge._start_powerpoint(win32, fake_py, set())
    assert win32.calls == 1


def test_start_powerpoint_does_not_retry_a_busy_instance(bridge):
    from kitchensink4ppt.core.errors import PowerPointBusy

    fake_py = _FakePythoncom()
    win32 = _FakeWin32([_com_error(bridge.RPC_E_CALL_REJECTED)] * 2)
    with pytest.raises(PowerPointBusy):
        bridge._start_powerpoint(win32, fake_py, set())
    assert win32.calls == 1


# ------------------------------------------- the caller-side apartment


def test_run_bounded_carries_a_caller_apartment_failure(bridge, monkeypatch):
    """The caller-side CoInitialize used to sit inside
    contextlib.suppress(Exception), so a thread whose apartment could not be
    armed produced no evidence at all."""
    boom = RuntimeError("apartment refused")

    def _fail(_pythoncom):
        state = f"failed: {boom}"
        bridge._APARTMENT_STATE[threading.get_ident()] = state
        return state

    monkeypatch.setattr(bridge, "_ensure_apartment", _fail)
    out = bridge._run_bounded("test_op", 30.0, lambda: "fine")
    assert out == "fine"
    # The state is recorded for the CALLING thread, which is this one.
    assert bridge.apartment_state().startswith("failed:")


def test_every_com_entry_point_uses_the_shared_apartment_helper():
    """(d) in the brief: not only export. A bare pythoncom.CoInitialize()
    left anywhere in the COM tier means one entry point can still drift."""
    from pathlib import Path

    import ast

    com_dir = Path(
        __file__
    ).resolve().parents[2] / "src" / "kitchensink4ppt" / "com"
    # _ensure_apartment IS the shared helper, and its own body is where the
    # fallback for a pywin32 without CoInitializeEx legitimately lives.
    # Everywhere else has to go through it.
    sanctioned = {"_ensure_apartment"}
    offenders = []
    for path in com_dir.glob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                continue
            if node.name in sanctioned:
                continue
            for inner in ast.walk(node):
                if (
                    isinstance(inner, ast.Call)
                    and isinstance(inner.func, ast.Attribute)
                    and inner.func.attr == "CoInitialize"
                ):
                    offenders.append(f"{path.name}:{inner.lineno}")
    assert not offenders, (
        "bare CoInitialize() calls left in the COM tier outside the shared "
        f"helper: {offenders}. Use _ensure_apartment so S_FALSE and "
        "RPC_E_CHANGED_MODE are handled in one place."
    )
