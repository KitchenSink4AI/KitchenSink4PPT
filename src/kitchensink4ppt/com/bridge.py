"""COM bridge: the operations only a running PowerPoint can perform.

PowerPoint is a SINGLETON COM server. Unlike Word, DispatchEx does NOT give
a private second instance when the user already has PowerPoint open; it
attaches to the user's running process. Every function here therefore obeys
the singleton contract:

- Record whether POWERPNT.EXE was running BEFORE the call (process table,
  never GetActiveObject). Quit() is issued ONLY when this call launched the
  process; an attached pre-existing instance is never quit, never has its
  windows touched, and only the presentations THIS call opened are closed.
- If the target file is already open in a pre-existing instance (FullName
  comparison across app.Presentations), the call REFUSES with DocumentLocked
  rather than operating on the user's open copy. The future live layer is
  the sanctioned path for open presentations.
- Invisibility is per-presentation: Presentations.Open(..., WithWindow=False).
  Application.Visible is never touched (setting it False raises on modern
  PowerPoint).
- CoInitialize is called freely and CoUninitialize is NEVER called: pywin32's
  CoUninitialize always decrements, and a paired call on a host-initialized
  thread destroys the host's apartment and disconnects every COM proxy it
  holds (verified empirically on the Word editor's live layer).
- After a self-launched instance is quit, a bounded 15s poll confirms the
  process actually exited (POWERPNT lingers a few seconds after Quit).
- Every exported artifact is verified to exist and be non-empty after the
  COM call reports success ("reported success but produced no output").

Nothing here kills processes it did not spawn, and nothing here attaches to
the user's instance beyond the refusal check above.

LIVE-SAFETY STACK (the same design as the Word editor's live layer):
- Serialization: every COM entry point runs under the process-wide COM lock
  (com/serial.py). _powerpoint() itself takes the lock, so the com_gates
  scripts that enter a session directly are covered too. On a singleton
  server this is not a mitigation, it is the only isolation there is.
- Bounded timeouts: hidden operations use a 60-second deadline.  After the
  existing false-timeout grace, cancellation reuses the server-exit reaper:
  it may end only a journalled process whose creation time, executable and
  /Automation command line are revalidated independently. COM normally also
  revalidates invisibility and the complete presentation set. If the owned
  instance's own modal dialog blocks that inspection, an identity-only route
  is allowed only for a journalled read-only open when the OS window layer
  sees that dialog and no visible application frame. Foreign, unsaved,
  untitled, ambiguous, and unowned instances are left running. The refusal
  records what happened.
- Queue abandonment: a call that gives up while its worker is still
  QUEUED on the serialization lock marks itself abandoned, atomically,
  against the worker's own attempt to start. The loser stands down, so a
  caller told the operation did not happen never has it happen anyway
  behind its back, and the retry it was offered is safe.
- DisplayAlerts: suppressed for the duration and RESTORED when the instance
  is the user's. The earlier code set ppAlertsNone on whatever instance it
  reached and never put it back, which on a singleton leaked into the
  user's own session.
- Dialogs: powerpoint_status reports modal dialogs seen at the OS WINDOW
  layer (com/dialogs.py), because COM cannot see the dialogs that are
  themselves blocking COM.
"""

from __future__ import annotations

import contextlib
import functools
import os
import re
import subprocess
import sys
import threading
import time
from pathlib import Path

from ..core.errors import (
    DocumentCorrupt,
    DocumentLocked,
    PowerPointBlocked,
    PowerPointBusy,
    PowerPointDisconnected,
    PowerPointNotRunning,
    PptMcpError,
)
from ..core import safesave
from ..core.sandbox import ServerScratch, check_path
from ..core.swapwait import require_deck
from . import serial as _serial

PROCESS_NAME = "POWERPNT.EXE"

PP_ALERTS_NONE = 1  # ppAlertsNone
PP_SAVE_AS_PDF = 32  # ppSaveAsPDF
PP_FIXED_FORMAT_TYPE_PDF = 2  # ppFixedFormatTypePDF
QUIT_POLL_SECONDS = 15.0
#: How much longer a timed-out operation is given before cancellation is
#: attempted. A worker that finishes in this window has finished and its
#: result is returned instead of a false timeout.
TIMEOUT_GRACE_SECONDS = 10.0

#: TEST SEAM. _run_bounded builds its internal completion signal
#: through this factory, so a test can substitute an event whose waits it
#: drives and pin WHICH stage observed completion without assuming
#: anything about thread scheduling. Production behaviour is unchanged:
#: this is threading.Event and nothing in the package reassigns it.
_COMPLETION_EVENT_FACTORY = threading.Event

DEFAULT_IMAGE_WIDTH = 1280

# HRESULTs as signed ints, the way pywin32 surfaces them.
RPC_E_CALL_REJECTED = -2147418111  # modal dialog / call rejected
RPC_E_SERVERCALL_RETRYLATER = -2147417846  # busy, retry later
RPC_E_DISCONNECTED = -2147417848  # server died under a live proxy
CO_E_OBJNOTCONNECTED = -2147220995  # proxy no longer connected
RPC_S_CALL_FAILED = -2147023170  # 0x800706BE, app killed mid-call
RPC_S_SERVER_UNAVAILABLE = -2147023174  # 0x800706BA, RPC server gone

# The COM START-UP family: the apartment, the registration and the launch.
# None of these is a caller mistake, and mapping them to PptMcpError sent
# them to the wire as BAD_PARAMS, which tells a caller to go and look at
# its own arguments.
CO_E_NOTINITIALIZED = -2147221008  # 0x800401F0, CoInitialize not called
CO_E_ALREADYINITIALIZED = -2147221007  # 0x800401F1, apartment already up
RPC_E_CHANGED_MODE = -2147417850  # 0x80010106, thread is in the other model
REGDB_E_CLASSNOTREG = -2147221164  # 0x80040154, PowerPoint not registered
CO_E_SERVER_EXEC_FAILURE = -2146959355  # 0x80080005, the server failed to run
S_FALSE = 1  # CoInitializeEx: already initialized in this same mode

BUSY_HRESULTS = {RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER}
GONE_HRESULTS = {
    RPC_E_DISCONNECTED,
    CO_E_OBJNOTCONNECTED,
    RPC_S_CALL_FAILED,
    RPC_S_SERVER_UNAVAILABLE,
}
#: Start-up faults: PowerPoint never came up, so nothing is running.
STARTUP_HRESULTS = {
    CO_E_NOTINITIALIZED,
    CO_E_ALREADYINITIALIZED,
    RPC_E_CHANGED_MODE,
    REGDB_E_CLASSNOTREG,
    CO_E_SERVER_EXEC_FAILURE,
}
#: The subset worth one more try. The ONLY action taken between the two
#: attempts is re-arming this thread's apartment, so the retry set is
#: exactly the errors that action can repair, and nothing else.
#: Deliberately NOT retried:
#:   REGDB_E_CLASSNOTREG       a second attempt cannot register a class;
#:   RPC_E_CHANGED_MODE        the thread's apartment model is already set,
#:                             and re-initializing it the same way cannot
#:                             change it, so a retry is theatre;
#:   CO_E_SERVER_EXEC_FAILURE  PowerPoint may have PARTIALLY LAUNCHED before
#:                             failing, and starting a second one without
#:                             accounting for the first is how a process
#:                             gets stranded and the zombie alarm fires on
#:                             an otherwise successful call.
STARTUP_RETRY_HRESULTS = {
    CO_E_NOTINITIALIZED,
    CO_E_ALREADYINITIALIZED,
}
STARTUP_RETRY_DELAY = 0.5  # seconds between the two start attempts

#: What the last CoInitializeEx on each thread did, so a start-up failure
#: can SAY what the apartment was instead of the caller guessing. Written
#: by _ensure_apartment, read by the start-up error message and by
#: powerpoint_status. Keyed by thread id.
_APARTMENT_STATE: dict[int, str] = {}


def _com_modules():
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:  # pragma: no cover
        raise PptMcpError(
            "pywin32 is not available; COM operations need it (Windows with "
            "PowerPoint installed). File-based tools work without it."
        ) from exc
    return pythoncom, win32com.client


def _hresults(exc) -> set:
    """Every HRESULT pywin32 might have stashed in a com_error: both
    exc.hresult and the scode buried in the EXCEPINFO tuple."""
    out = set()
    hr = getattr(exc, "hresult", None)
    if hr is not None:
        out.add(hr)
    args = getattr(exc, "args", ())
    if len(args) >= 3 and args[2]:
        with contextlib.suppress(Exception):
            scode = args[2][5]
            if scode is not None:
                out.add(scode)
    return out


def _classify(exc):
    """Map a com_error to our typed errors, or None if unrecognized."""
    hrs = _hresults(exc)
    if hrs & GONE_HRESULTS:
        return PowerPointDisconnected(
            "PowerPoint or the presentation closed while the operation was "
            "running; the operation may be incomplete."
        )
    if hrs & BUSY_HRESULTS:
        return PowerPointBusy(
            "PowerPoint is busy or has a dialog open (a dialog box, "
            "Backstage, or a running command). Close it and retry."
        )
    if hrs & STARTUP_HRESULTS:
        # "Nothing was opened" is only true when PowerPoint never ran. A
        # server-execution failure can leave a partially launched process,
        # so that case gets the honest sentence instead.
        aftermath = (
            " No presentation was opened and no file was changed; a "
            "PowerPoint process may have been started, and this call does "
            "not end PowerPoint processes."
            if CO_E_SERVER_EXEC_FAILURE in hrs
            else " Nothing was opened and nothing was changed."
        )
        return PowerPointNotRunning(
            "PowerPoint could not be started on this thread: "
            + _startup_detail(hrs)
            + aftermath
            + " The file-based tools do not need PowerPoint; retry the COM "
            "call, and if it keeps failing check that PowerPoint is "
            "installed and that no install repair is in progress."
        )
    return None


def _startup_detail(hrs: set) -> str:
    """The one sentence that names WHICH start-up fault this was, plus what
    this thread's COM apartment was doing when it happened."""
    if REGDB_E_CLASSNOTREG in hrs:
        what = (
            "PowerPoint.Application is not registered on this machine "
            "(REGDB_E_CLASSNOTREG)."
        )
    elif CO_E_SERVER_EXEC_FAILURE in hrs:
        what = (
            "the PowerPoint COM server failed to launch "
            "(CO_E_SERVER_EXEC_FAILURE)."
        )
    elif RPC_E_CHANGED_MODE in hrs:
        what = (
            "this thread's COM apartment is in the other threading model "
            "(RPC_E_CHANGED_MODE)."
        )
    else:
        what = (
            "this thread's COM apartment was not initialized "
            "(CO_E_NOTINITIALIZED)."
        )
    state = _APARTMENT_STATE.get(threading.get_ident())
    if state:
        what += f" The apartment on this thread reported: {state}."
    return what


def _ensure_apartment(pythoncom) -> str:
    """Initialize THIS thread's COM apartment and say what happened.

    CoInitializeEx rather than CoInitialize so the two benign outcomes are
    handled deliberately instead of by luck: S_FALSE means the apartment was
    already up in this same model (fine, and pywin32 does not raise for it),
    and RPC_E_CHANGED_MODE means some other code already put this thread in
    the other threading model. Fighting that would break whatever set it, so
    it is recorded and the call proceeds; COM still marshals, it just
    marshals differently.

    A pythoncom without CoInitializeEx falls back to CoInitialize, which is
    the same request in the older API. A MISSING API is not an apartment
    failure and must not be reported as one, or the refusal below would
    turn an old pywin32 into "PowerPoint could not be started".

    CoUninitialize is never paired here, for the reason in the module
    docstring. Returns one of "initialized", "already", "changed-mode", or
    "failed: ...", and records it per thread so a later start-up failure can
    quote it."""
    tid = threading.get_ident()
    init_ex = getattr(pythoncom, "CoInitializeEx", None)
    mode = getattr(pythoncom, "COINIT_APARTMENTTHREADED", 2)
    try:
        if init_ex is not None:
            init_ex(mode)
        else:
            pythoncom.CoInitialize()
        state = "initialized"
    except Exception as exc:
        hrs = _hresults(exc)
        if S_FALSE in hrs:
            state = "already"
        elif RPC_E_CHANGED_MODE in hrs or CO_E_ALREADYINITIALIZED in hrs:
            state = "changed-mode"
        else:
            state = f"failed: {exc}"
    _APARTMENT_STATE[tid] = state
    return state


def apartment_state(tid: int | None = None) -> str | None:
    """What the last CoInitializeEx on a thread did (diagnostic)."""
    return _APARTMENT_STATE.get(
        threading.get_ident() if tid is None else tid
    )


def _raise_classified(exc, fallback_message: str):
    typed = _classify(exc)
    if typed is not None:
        raise typed from exc
    raise PptMcpError(f"{fallback_message}: {exc}") from exc


def powerpnt_pids() -> set | None:
    """POWERPNT.EXE process ids via the process table (never COM), or None
    when the table could not be READ. PID precision is what lets the
    server name exactly the instance it started and nothing else.

    NONE MEANS UNKNOWN AND NEVER MEANS "NOTHING WAS RUNNING". This used to
    return an empty set for both, and every ownership decision in the
    module then read a failed tasklist as an empty machine: a session
    concluded it had LAUNCHED the PowerPoint it had merely attached to and
    quit the user's instance on the way out, unsaved work and all. Nothing
    that can end, quit or kill a process may act on None.

    A genuine no-match is NOT a failure: tasklist prints its INFO line and
    the answer is an empty set. Unknown is the subprocess raising or timing
    out, a non-zero exit with nothing on stdout, no output at all, or a row
    naming POWERPNT.EXE whose pid column will not parse.
    """
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {PROCESS_NAME}", "/FO", "CSV",
             "/NH"],
            capture_output=True,
            text=True,
            timeout=30,
            stdin=subprocess.DEVNULL,
        )
    except Exception:
        return None
    if result.returncode != 0:
        # EVERY nonzero return is unknown, text on stdout or not: tasklist
        # prints "ERROR: Access is denied." to stdout and exits 1, which
        # parsed as a clean empty table.
        # A genuine no-match exits 0 with its INFO line, so nothing that
        # really means "nothing is running" is lost here.
        return None
    stdout = result.stdout or ""
    if not stdout.strip():
        return None
    pids = set()
    for ln in stdout.splitlines():
        if PROCESS_NAME not in ln.upper():
            continue
        parts = ln.split('","')
        if len(parts) < 2:
            return None
        try:
            pids.add(int(parts[1].strip('"')))
        except ValueError:
            return None
    return pids


def powerpnt_count() -> int:
    """POWERPNT.EXE process count via the process table (never COM), or -1
    when the table could not be read. A diagnostic: every decision that can
    end a process reads powerpnt_pids() itself and stops on None."""
    pids = powerpnt_pids()
    return -1 if pids is None else len(pids)


# POWERPNT.EXE pids this server LAUNCHED, keyed by spawning thread.
# EVIDENCE for diagnostics, not a termination target list. Timeout
# cancellation uses the stronger _OWNED_JOURNAL below, including creation
# time, and then reruns every exit-reaper identity and presentation check. An
# entry exists only when _powerpoint() READ the process table, found no
# PowerPoint running, started one itself, read the table again to name the
# one new pid, and found that instance in the state of a just-created
# automation server. When we attached, when the table could not be read at
# either end, or when the application did not look freshly created, this
# dict stays empty for that thread.
_SELF_LAUNCHED_PIDS: dict[int, set] = {}
_OWNED_JOURNAL: dict[int, dict[int, int]] = {}
_OWNED_PRESENTATION_JOURNAL: dict[int, set[str]] = {}
_OWNED_WRITE_INTENT_TIDS: set[int] = set()
_COM_WORKERS: set[threading.Thread] = set()
_COM_WORKERS_LOCK = threading.Lock()
SHUTDOWN_JOIN_S = 60.0
_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _self_launched_pids_for_thread(tid) -> set:
    """The POWERPNT pids the given worker thread is known, ON POSITIVE
    EVIDENCE, to have launched. Empty whenever we attached to a PowerPoint
    we did not start, which is the normal case.

    This token records only what was observed at entry. It is used to make
    an unowned-timeout refusal informative, never by itself as permission to
    end a process."""
    return set(_SELF_LAUNCHED_PIDS.get(tid) or set())


_PROCESS_TERMINATE = 0x0001
_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000


def _kernel32():
    """Return kernel32 with pointer-safe process API prototypes."""
    import ctypes
    from ctypes import wintypes

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.OpenProcess.restype = wintypes.HANDLE
    k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    k32.CloseHandle.argtypes = [wintypes.HANDLE]
    k32.GetProcessTimes.argtypes = [wintypes.HANDLE] + [
        ctypes.POINTER(wintypes.FILETIME)
    ] * 4
    k32.QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE,
        wintypes.DWORD,
        wintypes.LPWSTR,
        ctypes.POINTER(wintypes.DWORD),
    ]
    return k32


def _creation_time(pid: int, handle=None) -> int | None:
    """Read one process creation FILETIME, or return None on any doubt."""
    if sys.platform != "win32":
        return None
    import ctypes
    from ctypes import wintypes

    k32 = _kernel32()
    own = handle is None
    process = handle or k32.OpenProcess(
        _PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid)
    )
    if not process:
        return None
    try:
        times = [wintypes.FILETIME() for _ in range(4)]
        if not k32.GetProcessTimes(process, *[ctypes.byref(value) for value in times]):
            return None
        created = times[0]
        return (created.dwHighDateTime << 32) | created.dwLowDateTime
    finally:
        if own:
            k32.CloseHandle(process)


def _is_automation_instance(pid: int) -> bool | None:
    """Return True only when the live process command line names /Automation."""
    if sys.platform != "win32":
        return None
    try:
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                "(Get-CimInstance Win32_Process -Filter "
                f"\"ProcessId={int(pid)}\").CommandLine",
            ],
            capture_output=True,
            text=True,
            timeout=20,
            stdin=subprocess.DEVNULL,
            creationflags=_NO_WINDOW,
        )
    except Exception:
        return None
    line = (result.stdout or "").strip()
    if result.returncode != 0 or not line:
        return None
    return "/automation" in line.lower()


def _powerpoint_app_pid(app) -> int | None:
    """Return the singleton PowerPoint pid, or ``None`` on any ambiguity.

    A hidden automation instance has no usable ``Application.HWND``.  Process
    ownership is established separately from the journaled pid, creation
    time, executable image and ``/Automation`` command line.  This helper is
    only the conservative singleton-to-proxy check: if the process table does
    not contain exactly one PowerPoint, it refuses to guess which proxy COM
    returned.
    """
    del app  # the proxy has no process identity API when it is hidden
    pids = powerpnt_pids()
    if pids is None or len(pids) != 1:
        return None
    return next(iter(pids))


def _inspect_exit_ownership(
    pid: int,
    owned_paths: set[str],
    *,
    timeout: float = 10.0,
) -> tuple[bool, str]:
    """Re-read the singleton and close only presentations this server opened.

    The exit reaper may end the process only when it is still invisible and
    every open presentation belongs to this server.  COM inspection itself
    is bounded: a blocked or ambiguous read leaves the process alone.
    """
    outcome: list[tuple[bool, str]] = []

    def inspect() -> None:
        try:
            import pythoncom
            import win32com.client

            _ensure_apartment(pythoncom)
            app = win32com.client.GetActiveObject("PowerPoint.Application")
            actual_pid = _powerpoint_app_pid(app)
            if actual_pid != int(pid):
                outcome.append((False, "the active PowerPoint PID could not be matched"))
                return
            visible = int(app.Visible) != 0
            presentations = []
            foreign: list[str] = []
            unsaved_owned: list[str] = []
            count = int(app.Presentations.Count)
            for index in range(1, count + 1):
                pres = app.Presentations.Item(index)
                try:
                    full_name = str(pres.FullName or "")
                except Exception:
                    full_name = ""
                key = safesave.canonical_key(full_name) if full_name else ""
                try:
                    saved = bool(pres.Saved)
                except Exception:
                    # An unreadable saved state is not permission to discard
                    # work.  Treat it as unsaved and leave the process alive.
                    saved = False
                presentations.append((pres, key, full_name, saved))
                if not key or key not in owned_paths:
                    foreign.append(Path(full_name).name if full_name else "an unsaved presentation")
                elif not saved:
                    unsaved_owned.append(Path(full_name).name or "an unsaved presentation")

            # Release only saved presentations that this server opened,
            # without Save or SaveAs.  An unsaved owned presentation is left
            # open and prevents process termination; ambiguity never grants
            # permission to discard work.
            for pres, key, _full_name, saved in reversed(presentations):
                if key and key in owned_paths and saved:
                    with contextlib.suppress(Exception):
                        pres.Close()

            if visible:
                outcome.append((False, "the application is visible"))
            elif foreign:
                names = ", ".join(foreign[:3])
                extra = "" if len(foreign) <= 3 else f" (+{len(foreign) - 3} more)"
                outcome.append((False, f"it holds a presentation this server did not open: {names}{extra}"))
            elif unsaved_owned:
                names = ", ".join(unsaved_owned[:3])
                extra = "" if len(unsaved_owned) <= 3 else f" (+{len(unsaved_owned) - 3} more)"
                outcome.append((False, f"a presentation this server opened has unsaved changes: {names}{extra}"))
            else:
                outcome.append((True, "invisible with no foreign or unsaved presentations"))
        except Exception as exc:
            outcome.append((False, f"ownership inspection failed: {type(exc).__name__}"))

    worker = threading.Thread(
        target=inspect,
        name=f"ks4p-exit-inspect-{int(pid)}",
        daemon=True,
    )
    worker.start()
    worker.join(max(0.0, float(timeout)))
    if worker.is_alive():
        return False, "ownership inspection timed out"
    if not outcome:
        return False, "ownership inspection produced no result"
    return outcome[0]


def _identity_only_modal_cleanup_safe(
    pid: int,
    owned_paths: set[str],
    inspection_why: str,
) -> tuple[bool, str]:
    """Allow the narrow no-COM reaper route for this server's modal only.

    The normal inspection remains authoritative whenever COM answers.  This
    fallback exists for the circular failure where a modal raised by a
    journalled read-only ``Presentations.Open`` prevents COM from answering
    the ownership questions.  Process identity is checked by the caller;
    here the independent Win32 window layer must show at least one dialog
    and no visible non-dialog window for that exact pid.
    """
    unavailable = (
        inspection_why in {
            "ownership inspection timed out",
            "ownership inspection produced no result",
        }
        or inspection_why.startswith("ownership inspection failed:")
    )
    if not unavailable:
        return False, "COM ownership inspection answered and refused cleanup"
    if not owned_paths:
        return False, "no journalled read-only presentation open"

    try:
        from . import dialogs as _dialogs

        windows = _dialogs.window_classes({int(pid)})
    except Exception as exc:
        return False, f"window inspection failed: {type(exc).__name__}"
    if not windows:
        return False, "no visible modal window was attributed to the owned pid"
    modal = [
        item for item in windows
        if item.get("class") in _dialogs.DIALOG_CLASSES
    ]
    non_dialog = [
        item for item in windows
        if item.get("class") not in _dialogs.DIALOG_CLASSES
    ]
    if not modal:
        return False, "no modal dialog was attributed to the owned pid"
    if non_dialog:
        classes = ", ".join(
            sorted({str(item.get("class") or "unknown") for item in non_dialog})
        )
        return False, f"the owned pid has a visible non-dialog window: {classes}"
    return True, "identity matched and only the owned pid's modal dialog is visible"


def _end_owned_powerpoint_at_exit(
    journal: dict[int, int],
    owned_presentations: dict[int, set[str]] | None = None,
) -> dict:
    """End only journaled PowerPoint automation processes at server exit.

    Every target must still match the recorded creation time, executable
    image and /Automation command line.  An unreadable or mismatched fact
    leaves the process alone.
    """
    result: dict = {"ended": [], "left": []}
    if sys.platform != "win32":
        return result
    import ctypes
    from ctypes import wintypes

    k32 = _kernel32()
    k32.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
    owned_presentations = owned_presentations or {}
    for pid, created in sorted(journal.items()):
        process = k32.OpenProcess(
            _PROCESS_TERMINATE | _PROCESS_QUERY_LIMITED_INFORMATION,
            False,
            int(pid),
        )
        if not process:
            result["left"].append({"pid": pid, "why": "gone or not openable"})
            continue
        try:
            if _creation_time(pid, handle=process) != created:
                result["left"].append(
                    {"pid": pid, "why": "creation time does not match"}
                )
                continue
            image = ctypes.create_unicode_buffer(1024)
            size = wintypes.DWORD(len(image))
            if (
                not k32.QueryFullProcessImageNameW(
                    process, 0, image, ctypes.byref(size)
                )
                or not image.value.lower().endswith("\\powerpnt.exe")
            ):
                result["left"].append({"pid": pid, "why": "not POWERPNT.EXE"})
                continue
            if _is_automation_instance(pid) is not True:
                result["left"].append(
                    {"pid": pid, "why": "not read as an /Automation instance"}
                )
                continue
            safe_to_end, why = _inspect_exit_ownership(
                pid, set(owned_presentations.get(pid) or set())
            )
            if not safe_to_end:
                identity_safe, identity_why = _identity_only_modal_cleanup_safe(
                    pid,
                    set(owned_presentations.get(pid) or set()),
                    why,
                )
                if not identity_safe:
                    result["left"].append(
                        {"pid": pid, "why": f"{why}; {identity_why}"}
                    )
                    continue
            if k32.TerminateProcess(process, 1):
                result["ended"].append(pid)
            else:
                result["left"].append({"pid": pid, "why": "could not be ended"})
        finally:
            k32.CloseHandle(process)
    return result


def _cancel_hung_call(worker_tid: int | None) -> dict:
    """Attempt to end only the hidden PowerPoint owned by one hung worker.

    The journal is necessary but not sufficient authority.  The shared
    reaper revalidates process identity and inspects current presentations;
    any foreign, unsaved, untitled, visible, unreadable, or otherwise
    ambiguous state leaves the process alone.  No journal means no attempt.
    """
    if worker_tid is None:
        return {"attempted": False, "ended": [], "left": []}
    journal = dict(_OWNED_JOURNAL.get(worker_tid) or {})
    if not journal:
        return {"attempted": False, "ended": [], "left": []}
    paths = set(_OWNED_PRESENTATION_JOURNAL.get(worker_tid) or set())
    if worker_tid in _OWNED_WRITE_INTENT_TIDS:
        # Identity-only cleanup is allowed only for read-only opens.  Passing
        # no owned paths makes the fallback refuse while the normal COM
        # inspection can still prove a safe cleanup when PowerPoint answers.
        paths = set()
    presentations = {pid: set(paths) for pid in journal}
    try:
        result = _end_owned_powerpoint_at_exit(journal, presentations)
    except Exception as exc:  # cancellation failure must not mask timeout
        result = {
            "ended": [],
            "left": [
                {
                    "pid": pid,
                    "why": f"cancellation check failed: {type(exc).__name__}",
                }
                for pid in sorted(journal)
            ],
        }
    return {"attempted": True, **result}


def _drain_on_exit(timeout: float = SHUTDOWN_JOIN_S) -> dict:
    """Wait once for COM workers, then reap only positively owned instances."""
    end = time.monotonic() + max(0.0, float(timeout))
    with _COM_WORKERS_LOCK:
        workers = list(_COM_WORKERS)
    for worker in workers:
        worker.join(max(0.0, end - time.monotonic()))
    still = [worker for worker in workers if worker.is_alive()]
    journal = {
        pid: created
        for worker in still
        for pid, created in (_OWNED_JOURNAL.get(worker.ident) or {}).items()
    }
    owned_presentations = {
        pid: (
            set()
            if worker.ident in _OWNED_WRITE_INTENT_TIDS
            else set(_OWNED_PRESENTATION_JOURNAL.get(worker.ident) or set())
        )
        for worker in still
        for pid in (_OWNED_JOURNAL.get(worker.ident) or {})
    }
    ended = (
        _end_owned_powerpoint_at_exit(journal, owned_presentations)
        if journal
        else {"ended": [], "left": []}
    )
    if ended["ended"]:
        sys.stderr.write(
            "[kitchensink4ppt] ended the hidden PowerPoint instance(s) this "
            "server started after the bounded exit wait: "
            + ", ".join(map(str, ended["ended"]))
            + "\n"
        )
    for item in ended["left"]:
        sys.stderr.write(
            f"[kitchensink4ppt] left PowerPoint pid {item['pid']} running: "
            f"{item['why']}\n"
        )
    if still and not ended["ended"] and not ended["left"]:
        sys.stderr.write(
            f"[kitchensink4ppt] {len(still)} PowerPoint call(s) remain after "
            f"the {max(0.0, float(timeout)):g}s exit wait; no unverified "
            "process was touched\n"
        )
    return {
        "waited_for": len(workers),
        "still_running": len(still),
        "ended": ended["ended"],
        "left": ended["left"],
    }


@contextlib.contextmanager
def _alerts_suppressed(app):
    """DisplayAlerts off around COM calls, restored after. With alerts on,
    contention and repair prompts freeze PowerPoint behind a modal dialog
    that COM itself cannot see or dismiss; with alerts off the same
    condition raises a catchable com_error. The restore matters on the
    singleton: an unrestored ppAlertsNone would leak into the user's own
    PowerPoint session and silently swallow their prompts too."""
    prev = None
    with contextlib.suppress(Exception):
        prev = app.DisplayAlerts
        app.DisplayAlerts = PP_ALERTS_NONE
    try:
        yield
    finally:
        if prev is not None:
            with contextlib.suppress(Exception):
                app.DisplayAlerts = prev


def _run_bounded(name: str, timeout: float, fn):
    """Run fn on a worker thread under the COM lock, bounding how long the
    CALLER waits.

    On expiry: if the worker never got the lock, that is queue contention
    and PowerPointBusy names the operation actually running. If it got the
    lock and then stalled, a grace wait avoids false timeouts before an
    identity-checked cancellation attempt. Cancellation can end only the
    hidden process positively journalled for this worker and currently
    confirmed to contain no foreign, unsaved, or untitled presentation.

    `timeout` is the PowerPoint response deadline rather than the maximum
    caller elapsed time: TIMEOUT_GRACE_SECONDS, dialog inspection, and the
    bounded identity/presentation checks follow. A result that arrives
    before cancellation begins is still returned.

    The completion check is TWO-STAGE: once the instant the grace wait
    ends, before any process or dialog inspection, so a completed success
    is never delayed behind a process-table call; and once more after that
    inspection, which can itself take long enough for the worker to finish
    inside it. Either way the worker's own answer wins, its result
    returned or its exception raised, because reporting a timeout for an
    operation that succeeded is a false failure the caller acts on. Only if
    it is still not done does cancellation begin. The
    resulting refusal distinguishes an ended owned instance from a process
    left alone because ownership or document safety could not be proved.
    """
    # Preserve the bridge's documented side effect: every public
    # operation once ran ON THE CALLING THREAD and left that thread's COM
    # apartment initialized (this module initializes freely and never
    # uninitializes, by design). Moving the body to a worker thread moved
    # the CoInitialize with it, so callers that went on to make their own
    # COM calls afterwards started failing with "CoInitialize has not been
    # called" (caught by test_av's media scenario). Initialize
    # the caller's apartment too, so the contract survives the port.
    # The failure used to be swallowed by contextlib.suppress, which meant a
    # caller thread whose apartment could not be armed produced no evidence
    # at all: the next COM call two frames away failed with "CoInitialize
    # has not been called" and the report said nothing about why.
    #
    # It is recorded now, in _APARTMENT_STATE, keyed by thread. That is an
    # IN-PROCESS diagnostic and nothing more: it is read by _startup_detail
    # when a start-up fault is classified, and reported by powerpoint_status
    # as com_apartment. It is deliberately NOT injected into this operation's
    # return value, because every public result here has an established shape
    # and diagnostics do not belong conditionally inside them.
    caller_apartment = "no pythoncom"
    try:
        import pythoncom

        caller_apartment = _ensure_apartment(pythoncom)
    except Exception as exc:  # noqa: BLE001 - recorded, never raised here
        caller_apartment = f"failed: {exc}"

    result: dict = {"caller_apartment": caller_apartment}
    done = _COMPLETION_EVENT_FACTORY()
    worker_tid: list = []

    # ONE decision per call, taken atomically under this lock, with exactly
    # two outcomes: the caller ABANDONED the queued call, or the worker
    # STARTED it. Whichever side reaches the decision first wins and the
    # other is told it lost, so the two can never both be true.
    decision_lock = threading.Lock()
    decision: list = []

    def _elect(outcome: str) -> bool:
        """Take this call's one decision, or lose it to the other side."""
        with decision_lock:
            if decision:
                return False
            decision.append(outcome)
            return True

    def worker():
        worker_tid.append(threading.get_ident())
        try:
            with _serial.com_operation(name):
                # THE GHOST WRITE STOPS HERE. Winning the serialization
                # lock is not permission to run. The caller may have given
                # up while this thread sat in the queue, and it was told
                # the operation had not happened; running fn() now would
                # perform that write anyway, so a retried non-idempotent
                # edit would execute twice (two slide deletions removing
                # two different slides). An abandoned call releases the
                # lock and exits WITHOUT touching PowerPoint.
                if not _elect("started"):
                    _serial._mark_outcome(
                        threading.get_ident(), "abandoned"
                    )
                    return
                result["value"] = fn()
        except BaseException as exc:  # noqa: BLE001 - re-raised in caller
            result["error"] = exc
        finally:
            done.set()
            with _COM_WORKERS_LOCK:
                _COM_WORKERS.discard(threading.current_thread())

    t = threading.Thread(target=worker, daemon=True, name=f"ks4p-{name}")
    with _COM_WORKERS_LOCK:
        _COM_WORKERS.add(t)
    t.start()
    if done.wait(timeout):
        if "error" in result:
            raise result["error"]
        return result["value"]
    if _elect("abandoned"):
        # The worker had NOT started fn when this decision was taken, so it
        # never will: it is still queued on the serialization lock, and
        # when it reaches the front it loses this same decision, releases
        # the lock and exits. Nothing ran and nothing was changed.
        if "error" in result:
            # It failed before it could reach the decision at all, so its
            # own error is the honest answer, not a queue-contention one.
            raise result["error"]
        snap = _serial.lock_snapshot()
        running = (snap.get("current_op") or {}).get(
            "name", "another COM operation"
        )
        raise PowerPointBusy(
            f"{name} waited {timeout:.0f}s for the COM serialization lock "
            f"({running} is still running); retry when it finishes. "
            "powerpoint_status reports the running operation."
            " The queued call was abandoned before its operation ran; it "
            "made no document changes. It is safe to retry after "
            "powerpoint_status reports no COM operation in progress."
        )
    # The worker won the decision, so fn IS running, or has already
    # finished. This is no longer queue contention and the caller must not
    # be told to simply retry: the grace path decides from here, and an
    # operation still running reaches the cancellation decision.
    # The deadline has passed, but cancellation has not been attempted yet,
    # so the grace wait below is a real second chance rather than a formality.
    # A worker that finishes during it finished, and reporting a timeout for
    # an operation that succeeded is a false failure the caller would act
    # on.
    done.wait(TIMEOUT_GRACE_SECONDS)

    def _completed():
        """The worker's own answer, when it has one. TWO-STAGE by design:
        asked once the moment the grace wait ends, BEFORE any process or
        dialog inspection, so a completed success is never delayed behind
        a process-table call, and once more after that inspection, which
        can itself take long enough for the worker to finish inside it."""
        if not done.is_set():
            return False
        if "error" in result:
            raise result["error"]
        return "value" in result

    if _completed():
        return result["value"]

    dialogs_seen = []
    with contextlib.suppress(Exception):
        from . import dialogs as _dialogs

        dialogs_seen = _dialogs.pending_dialogs()

    if _completed():
        return result["value"]
    # Evidence for the refusal is read HERE, at the moment the refusal is
    # built, not before the wait: a token sampled earlier can have been
    # cleared by a worker that finished and quit its instance meanwhile.
    ours = _self_launched_pids_for_thread(
        worker_tid[0] if worker_tid else None
    )
    timed_out_tid = worker_tid[0] if worker_tid else None
    _serial._mark_outcome(timed_out_tid, "timed_out")
    cancellation = _cancel_hung_call(timed_out_tid)
    _serial._mark_outcome(
        timed_out_tid, "timed_out", cancellation=cancellation
    )

    ended = cancellation["ended"]
    left = cancellation["left"]
    if ended:
        pids = ", ".join(str(pid) for pid in ended)
        detail = ""
        action = (
            "This server ended the hidden PowerPoint it had started for this call "
            f"(pid {pids}), so the output file may be incomplete or missing; check "
            "it before you retry."
        )
    elif left:
        reasons = "; ".join(
            f"pid {item['pid']}: {item['why']}" for item in left[:3]
        )
        detail = (
            " PowerPoint was not ended, because ending it was not safe or not possible "
            f"({reasons})."
        )
        action = (
            "The operation was NOT cancelled: it may still be running and "
            "may still finish and save its output. Do not retry yet: call "
            "powerpoint_status and wait until it reports no COM operation "
            "in progress, then check the output file before repeating the "
            "call."
        )
    elif ours:
        pids = ", ".join(str(p) for p in sorted(ours))
        detail = (
            f" (PowerPoint process pid {pids} was not running when this "
            "call began; this call did not force-end it, and its current "
            "state was not re-checked)"
        )
        action = (
            "The operation was NOT cancelled: it may still be running and "
            "may still finish and save its output. Do not retry yet: call "
            "powerpoint_status and wait until it reports no COM operation "
            "in progress, then check the output file before repeating the "
            "call."
        )
    else:
        detail = (
            " (no newly started PowerPoint process could be identified; no "
            "process was force-ended)"
        )
        action = (
            "The operation was NOT cancelled: it may still be running and "
            "may still finish and save its output. Do not retry yet: call "
            "powerpoint_status and wait until it reports no COM operation "
            "in progress, then check the output file before repeating the "
            "call."
        )
    message = action + detail
    if dialogs_seen:
        titles = ", ".join(
            d.get("title") or d.get("class", "?") for d in dialogs_seen[:3]
        )
        # The note follows text that may already end in a full stop; join
        # it without doubling the stop, and end the note with one.
        message = (
            message.rstrip(".") + f". PowerPoint has a dialog open: {titles}."
        )
    blocked = PowerPointBlocked(
        f"PowerPoint did not answer within {timeout:.0f} s. " + message
    )
    # What happened, on the exception itself, so a caller tells a time
    # limit from every other refusal without reading the words.
    blocked.timed_out = True
    blocked.time_limit_s = timeout
    blocked.ended = list(ended)
    blocked.left = list(left)
    raise blocked


def _bounded_op(name: str, default: float):
    """Public-function wrapper: adds a tunable timeout parameter and runs
    the body via _run_bounded (which serializes). Marks the function for
    the entry-point coverage audit."""

    def deco(fn):
        @functools.wraps(fn)
        def wrapper(*args, timeout: float = default, **kwargs):
            try:
                timeout = float(timeout)
            except (TypeError, ValueError):
                raise PptMcpError(
                    f"timeout must be a number of seconds, got {timeout!r}"
                ) from None
            if not 5 <= timeout <= 3600:
                raise PptMcpError(
                    "timeout must be between 5 and 3600 seconds"
                )
            return _run_bounded(name, timeout, lambda: fn(*args, **kwargs))

        wrapper._com_serialized = name
        wrapper._com_bounded = default
        return wrapper

    return deco


class _PowerPointSession:
    """What _powerpoint() yields: the app plus the bookkeeping the cleanup
    contract needs (did WE launch the process; which presentations WE opened)."""

    __slots__ = ("app", "launched", "opened")

    def __init__(self, app, launched: bool):
        self.app = app
        self.launched = launched
        self.opened: list = []


def _acquisition_token(app, before_pids) -> set | None:
    """The pids this call may claim to own, or None when it may not.

    ALL of these must hold, and a read that fails is a no:

    - the entry snapshot was READ and was empty, so nothing was up;
    - the table read now names EXACTLY ONE new pid, so there is one
      candidate rather than a crowd;
    - the application itself looks like a just-created automation
      instance: not visible, no presentations, no windows.

    The last part is the only one a concurrent launch cannot fake. A
    PowerPoint the user started is visible and, as soon as they open
    anything, has a presentation and a window; a fresh DispatchEx instance
    reports Visible 0, Presentations.Count 0 and Windows.Count 0 (measured
    on PowerPoint 16). Doubt means NOT ours, and not ours
    means this session behaves as an attach: nothing quit, nothing killed.
    """
    if before_pids is None or before_pids:
        return None
    after = powerpnt_pids()
    if after is None:
        return None
    created = after - before_pids
    if len(created) != 1:
        return None
    try:
        if int(app.Visible) != 0:
            return None
        if int(app.Presentations.Count) != 0:
            return None
        if int(app.Windows.Count) != 0:
            return None
    except Exception:
        return None
    return created


def _still_ours(app) -> bool:
    """Whether the instance may still be quit, re-read AT RELEASE and after
    this call's own presentations have been closed.

    The user can arrive during the call: they open a deck in the instance
    we started, or make it visible. Either way it is no longer ours to end,
    so this call closes what it opened and leaves the application running.
    """
    try:
        return int(app.Visible) == 0 and int(app.Presentations.Count) == 0
    except Exception:
        return False


@contextlib.contextmanager
def _powerpoint():
    """Singleton-safe PowerPoint context.

    Records the pre-call POWERPNT running state, dispatches (attaching to the
    user's instance if one exists, launching otherwise), disables alerts, and
    yields a _PowerPointSession. Cleanup closes ONLY presentations this call
    opened; Quit() fires ONLY if this call launched the process, followed by
    a bounded poll confirming the process exited. CoUninitialize is never
    called (see module docstring).

    The whole session runs under the process-wide COM lock, so callers
    that enter here directly (the com_gates scripts) are serialized without
    having to remember to be. When this call LAUNCHES PowerPoint its pid is
    recorded as evidence for a timeout refusal; when it attaches to the
    user's instance nothing is recorded, and the user's DisplayAlerts
    setting is restored on the way out instead of being left suppressed.
    """
    with _serial.com_operation("powerpoint_session"):
        yield from _powerpoint_locked()


def _start_powerpoint(win32client, pythoncom, before_pids):
    """DispatchEx with ONE bounded retry. NOTHING HERE ENDS A PROCESS.

    The only action taken between the two attempts is re-arming this
    thread's apartment, so only the faults that action repairs are retried
    (STARTUP_RETRY_HRESULTS).

    The previous version read the process table after a failed attempt and
    force-ended every POWERPNT.EXE that had appeared since entry, on the
    reasoning that an empty entry snapshot made any new process ours. It is
    not evidence of ownership: the user or another session can launch
    PowerPoint between that snapshot and this call's failure, and that
    instance was then killed with no save prompt. The snapshot can also be
    empty because `tasklist` itself failed, which turns "unknown" into
    "nothing was running".

    So the difference is only LOOKED at, never acted on: a pid that appears
    across a failed attempt is reported in the refusal as something left
    running, and its appearance stops the retry, because launching again on
    top of a state we cannot explain is how one unexplained process becomes
    two.

    `before_pids` is the entry snapshot, used only to tell an old process
    from a new one in that report. When it is NON-empty an instance was
    already up, we were attaching to it rather than starting one, and a
    process appearing mid-call is not ours to reason about at all: the look
    is skipped and the retry runs as before. When it is None the process
    table could not be read, so "new" has no meaning and the look is
    skipped for that reason instead.

    Returns (app, pids_that_appeared)."""
    watch = before_pids is not None and not before_pids
    appeared: set = set()
    last = None
    for attempt in (0, 1):
        try:
            return win32client.DispatchEx("PowerPoint.Application"), appeared
        except Exception as exc:  # noqa: BLE001 - re-raised classified below
            last = exc
            if watch:
                with contextlib.suppress(Exception):
                    now = powerpnt_pids()
                    if now is not None:
                        appeared |= now - before_pids
            if (
                attempt == 1
                or appeared
                or not (_hresults(exc) & STARTUP_RETRY_HRESULTS)
            ):
                break
            with contextlib.suppress(Exception):
                _ensure_apartment(pythoncom)
            time.sleep(STARTUP_RETRY_DELAY)
    _raise_startup(last, appeared)


def _raise_startup(exc, appeared: set) -> None:
    """Classify a start-up failure and report, without acting on it, any
    PowerPoint process that appeared while it was failing."""
    note = ""
    if appeared:
        pids = ", ".join(str(p) for p in sorted(appeared))
        note = (
            " A PowerPoint process appeared during the failed start: "
            f"pid {pids}. This call did not force-end it; its current "
            "state was not re-checked."
        )
    typed = _classify(exc)
    if typed is not None:
        if note:
            typed = type(typed)(str(typed) + note)
        raise typed from exc
    raise PptMcpError(
        f"PowerPoint could not be started on this thread: {exc}{note}"
    ) from exc


def _powerpoint_locked():
    pythoncom, win32client = _com_modules()
    # CoInitializeEx on the thread that is about to make the COM calls.
    # Tolerant of "already initialized" (S_FALSE) and of a thread another
    # library put in the other threading model (RPC_E_CHANGED_MODE), both
    # handled in _ensure_apartment; never paired with CoUninitialize.
    state = _ensure_apartment(pythoncom)
    if state.startswith("failed:"):
        # An apartment that genuinely could not be armed is not
        # something Dispatch can recover from. Refusing here names the real
        # cause instead of letting a second, vaguer COM failure speak for it.
        raise PowerPointNotRunning(
            "PowerPoint could not be started on this thread: the COM "
            f"apartment could not be initialized ({state[len('failed: '):]}). "
            "Nothing was opened and nothing was changed."
        )
    tid = threading.get_ident()
    before_pids = powerpnt_pids()
    # OWNERSHIP NEEDS POSITIVE EVIDENCE, and an unreadable process table is
    # not evidence of anything. before_pids is None when tasklist failed;
    # treating that as "nothing was running" is what made a session quit
    # the instance it had only attached to.
    started_one = before_pids is not None and not before_pids
    pre_count = len(before_pids) if before_pids is not None else 0
    app, _appeared = _start_powerpoint(win32client, pythoncom, before_pids)
    # OWNERSHIP IS A TOKEN READ OFF THE APPLICATION, not a difference of
    # two process-table readings. PowerPoint is a single-instance
    # automation server, so a user who starts it between the snapshot and
    # DispatchEx is handed to us as if we had started it ourselves; the
    # before/after difference then names THEIR pid and the cleanup quits
    # THEIR PowerPoint.
    created = _acquisition_token(app, before_pids)
    if created:
        # Record exactly the process we can show is ours, as evidence.
        _SELF_LAUNCHED_PIDS[tid] = created
        stamps = {pid: _creation_time(pid) for pid in created}
        stamps = {pid: stamp for pid, stamp in stamps.items() if stamp is not None}
        if stamps:
            _OWNED_JOURNAL[tid] = stamps
            _OWNED_PRESENTATION_JOURNAL[tid] = set()
    launched = created is not None
    session = _PowerPointSession(app, launched)
    completed = False
    try:
        if launched:
            # We own this process and will Quit it, so suppression needs no
            # restore. Deliberately no restoring guard here: the guard's
            # closure would hold a COM reference to app past the Quit and
            # the exit poll would then read a zombie.
            with contextlib.suppress(Exception):
                app.DisplayAlerts = PP_ALERTS_NONE
            yield session
        else:
            # The user's instance: suppress for our calls, put their
            # setting back afterwards.
            with _alerts_suppressed(app):
                yield session
        completed = True
    finally:
        _SELF_LAUNCHED_PIDS.pop(tid, None)
        _OWNED_JOURNAL.pop(tid, None)
        _OWNED_PRESENTATION_JOURNAL.pop(tid, None)
        _OWNED_WRITE_INTENT_TIDS.discard(tid)
        for pres in reversed(session.opened):
            with contextlib.suppress(Exception):
                pres.Close()
        session.opened.clear()
        with contextlib.suppress(NameError):
            del pres  # the loop variable is itself a COM reference
        zombie = False
        # Ownership is re-checked HERE, with this call's presentations
        # already closed. If the user opened something in the instance we
        # started, or made it visible, it is theirs now: we leave it
        # running rather than quitting it out from under them.
        if launched and _still_ours(app):
            # PowerPoint will not exit while external COM references are
            # outstanding: drop ours (ops del their locals before this
            # cleanup runs; see _release note in open_presentation) and
            # collect cycles BEFORE Quit, or the exit poll reads a zombie.
            import gc

            with contextlib.suppress(Exception):
                app.Quit()
            session.app = None
            del app
            gc.collect()
            deadline = time.monotonic() + QUIT_POLL_SECONDS
            zombie = True
            while time.monotonic() < deadline:
                with contextlib.suppress(Exception):
                    now = powerpnt_pids()
                    if now is None:
                        # The table went unreadable mid-poll. Unknown is not
                        # evidence of a zombie, and polling it again cannot
                        # become evidence either.
                        zombie = False
                        break
                    if len(now) <= pre_count:
                        zombie = False
                        break
                time.sleep(1.0)
        # Raise the zombie alarm only on an otherwise-clean run; never mask
        # the original exception with a cleanup complaint.
        if zombie and completed:
            raise PptMcpError(
                "the PowerPoint instance this call launched did not exit "
                f"within {QUIT_POLL_SECONDS:.0f}s of Quit(); a POWERPNT.EXE "
                "process may be lingering (zombie_check() to confirm). It was "
                "launched by this tool and is safe to end via Task Manager."
            )


def open_presentation(session: _PowerPointSession, path: Path, *, read_only: bool = True):
    """Open a presentation invisibly inside a _powerpoint() session.

    REFUSAL RULE: when attached to a pre-existing user instance, a target
    already open there is refused with DocumentLocked; this bridge never
    operates on the user's open copy. Opened presentations are registered on
    the session so cleanup closes exactly what this call opened.
    """
    target = str(path.resolve())
    if not session.launched:
        open_names: list[str] = []
        with contextlib.suppress(Exception):
            for pres in session.app.Presentations:
                open_names.append(str(pres.FullName))
        from .live import same_deck

        if any(same_deck(n, target) for n in open_names):
            raise DocumentLocked(
                f"{path.name} is open in the user's running PowerPoint; this "
                "tool never operates on the user's open copy. Close it in "
                "PowerPoint and retry, or use the live operations (live_save, "
                "live_scroll_to, live_status; dual-mode editing tools route "
                "there via live='auto') to work on the open copy."
            )
    if session.launched:
        tid = threading.get_ident()
        if read_only:
            # Record the attempted open before COM: an encrypted or corrupt
            # file can block inside Presentations.Open and never return a
            # proxy to journal afterwards.  This records intent only; the
            # process identity and modal-only window checks remain required.
            _OWNED_PRESENTATION_JOURNAL.setdefault(tid, set()).add(
                safesave.canonical_key(target)
            )
        else:
            _OWNED_WRITE_INTENT_TIDS.add(tid)
    try:
        pres = session.app.Presentations.Open(
            target,
            ReadOnly=read_only,
            Untitled=False,
            WithWindow=False,
        )
    except Exception as exc:
        typed = _classify(exc)
        if typed is not None:
            raise typed from exc
        raise _open_failure(path, exc) from exc
    session.opened.append(pres)
    return pres


#: Names Windows keeps for devices, with or without an extension; PowerPoint
#: opens nothing at a path with a folder or file named like one.
#: COM and LPT also take a superscript one, two or three (Windows lists them).
_SUPERSCRIPT_DIGITS = "".join(map(chr, (0xB9, 0xB2, 0xB3)))
_DEVICE_NAME = re.compile(
    r"(?i)^(con|prn|aux|nul|com[0-9" + _SUPERSCRIPT_DIGITS + r"]|lpt[0-9"
    + _SUPERSCRIPT_DIGITS + r"])(\..*)?$")


def _reserved_component(path: Path) -> str | None:
    """The first folder or file name in `path` that Windows keeps for a
    device, or None."""
    for part in Path(path).parts[1:]:
        if _DEVICE_NAME.match(part.rstrip(" .")):
            return part
    return None


def _open_failure(path: Path, exc: BaseException) -> DocumentCorrupt:
    """The refusal for a deck PowerPoint would not open. A path through a
    device-named folder is named as the cause: the deck itself may be fine,
    and "damaged" would send the user after the wrong problem."""
    from ..core.messages import RESERVED_NAME_HINT, RESERVED_NAME_OPEN

    component = _reserved_component(path)
    if component is not None:
        refusal = DocumentCorrupt(
            RESERVED_NAME_OPEN.format(name=Path(path).name, component=component))
        refusal.hint = RESERVED_NAME_HINT
        return refusal
    return DocumentCorrupt(
        f"PowerPoint could not open {Path(path).name} (with alerts disabled a "
        f"repair/recovery path surfaces as this error): {exc}"
    )


def _require_file(path: str, purpose: str) -> Path:
    path = check_path(path, purpose)
    require_deck(path, f"no file at {path}")
    return Path(path)


def _verify_output(out: Path, what: str) -> None:
    if not out.exists() or out.stat().st_size == 0:
        raise PptMcpError(
            f"PowerPoint reported success but produced no {what} at {out}"
        )


def _settled_pdf_output(
    source: Path, output: str | None, default: Path, overwrite: bool = False
) -> Path:
    """Settle and guard the exact PDF path before any COM method sees it.
    An existing file passes only with overwrite=True; the deck and its
    backups never do."""
    raw = Path(output) if output else default
    # A trailing dot is an empty extension. Python 3.14 reports it as the
    # suffix "."; 3.12 and 3.13 report no suffix and would keep the dot, so
    # strip it first to behave the same on every supported version.
    if raw.name.rstrip(".") and raw.name.endswith("."):
        raw = raw.with_name(raw.name.rstrip("."))
    out = raw.with_suffix(".pdf") if raw.suffix in ("", ".") else raw
    out = Path(check_path(str(out), "PowerPoint PDF output"))
    safesave.refuse_protected_output(
        out, source, purpose="write the PowerPoint PDF export"
    )
    if out.is_dir():
        raise PptMcpError(
            f"refusing to write the PowerPoint PDF export: {out} is a folder; "
            "name a .pdf file"
        )
    if out.exists() and not overwrite:
        raise PptMcpError(
            f"refusing to write the PowerPoint PDF export: {out} already "
            "exists; send overwrite=true to replace it, or choose a new output path"
        )
    return out


def _export_target(out: Path) -> Path:
    """Where PowerPoint writes: the output itself when it is new, or a
    temporary name beside it when the call replaces a file, so the old file
    stays whole until the new one is complete (_publish_export)."""
    if not out.exists():
        return out
    return out.with_name(f".~ks4p-{os.getpid()}-{out.name}")


def _publish_export(target: Path, out: Path, what: str) -> None:
    """Put a finished export in place: verify it, then replace the old file
    in one step. A failed export leaves the old file as it was."""
    if target == out:
        _verify_output(out, what)
        return
    try:
        _verify_output(target, what)
        os.replace(target, out)
    finally:
        with contextlib.suppress(OSError):
            if target.exists():
                target.unlink()


@_bounded_op("com_export_pdf", default=60.0)
def com_export_pdf(
    path: str, output: str | None = None, overwrite: bool = False
) -> dict:
    """Export a .pptx to PDF via PowerPoint (highest fidelity on this machine).

    The source is opened ReadOnly and exported with SaveCopyAs (which never
    touches the open presentation's identity or saved state), falling back to
    ExportAsFixedFormat; the source file is never modified.
    """
    p = _require_file(path, "PDF export source")
    out = _settled_pdf_output(p, output, p.with_suffix(".pdf"), overwrite)
    out.parent.mkdir(parents=True, exist_ok=True)
    target = _export_target(out)
    with _powerpoint() as session:
        pres = open_presentation(session, p, read_only=True)
        try:
            try:
                pres.SaveCopyAs(str(target.resolve()), PP_SAVE_AS_PDF)
            except Exception as first_exc:
                typed = _classify(first_exc)
                if typed is not None:
                    raise typed from first_exc
                try:
                    pres.ExportAsFixedFormat(
                        str(target.resolve()), PP_FIXED_FORMAT_TYPE_PDF
                    )
                except Exception as exc:
                    _raise_classified(exc, "PDF export failed")
        finally:
            del pres  # release the proxy so the launched instance can exit
    _publish_export(target, out, "PDF")
    return {"pdf": str(out), "bytes": out.stat().st_size, "engine": "powerpoint-com"}


#: ExportAsFixedFormat OutputType constants (ppPrintOutputType) by handout
#: slides-per-page. Non-contiguous on purpose: 4-up and 9-up were added later
#: than 2/3/6-up, so the enum order is historical, not numeric.
PP_HANDOUT_OUTPUT = {
    1: 10,  # ppPrintOutputOneSlideHandouts
    2: 2,   # ppPrintOutputTwoSlideHandouts
    3: 3,   # ppPrintOutputThreeSlideHandouts
    4: 8,   # ppPrintOutputFourSlideHandouts
    6: 4,   # ppPrintOutputSixSlideHandouts
    9: 9,   # ppPrintOutputNineSlideHandouts
}
PP_PRINT_OUTPUT_NOTES_PAGES = 5  # ppPrintOutputNotesPages
PP_FIXED_FORMAT_INTENT_PRINT = 2  # ppFixedFormatIntentPrint
PP_PRINT_HANDOUT_HORIZONTAL_FIRST = 2  # ppPrintHandoutHorizontalFirst


@_bounded_op("com_export_handout", default=60.0)
def com_export_handout(
    path: str,
    output: str | None = None,
    slides_per_page: int = 3,
    include_notes: bool = False,
    overwrite: bool = False,
) -> dict:
    """Export a .pptx to a handout-layout PDF via PowerPoint COM.

    Route (researched + verified empirically on PowerPoint 365):
    Presentation.ExportAsFixedFormat with ppFixedFormatTypePDF and the
    handout ppPrintOutputType constant. This is the parameter route that
    works headless (WithWindow=False); the PrintOptions route drives the
    PRINTER pipeline (ActivePrinter/Print-to-PDF) and needs print-driver
    dialogs, so it is not used. HandoutOrder applies left-to-right reading
    order (horizontal first, PowerPoint's dialog default for grids).

    slides_per_page: 1 | 2 | 3 | 4 | 6 | 9 (3-up is the classic
    lines-beside-slides handout). include_notes=True exports notes pages
    instead (one slide + its speaker notes per page); slides_per_page does
    not apply there and a non-default value refuses rather than being
    silently ignored. The source file is never modified (opened ReadOnly).
    """
    p = _require_file(path, "handout export source")
    if include_notes:
        if slides_per_page != 3:
            raise PptMcpError(
                "include_notes=True exports notes pages (one slide per "
                "page); slides_per_page does not apply; drop it or use "
                "include_notes=False"
            )
        output_type = PP_PRINT_OUTPUT_NOTES_PAGES
    else:
        if slides_per_page not in PP_HANDOUT_OUTPUT:
            raise PptMcpError(
                f"slides_per_page must be one of "
                f"{sorted(PP_HANDOUT_OUTPUT)}, got {slides_per_page!r}"
            )
        output_type = PP_HANDOUT_OUTPUT[slides_per_page]
    out = _settled_pdf_output(
        p,
        output,
        p.with_name(p.stem + ("_notes.pdf" if include_notes else "_handout.pdf")),
        overwrite,
    )
    out.parent.mkdir(parents=True, exist_ok=True)
    target = _export_target(out)
    with _powerpoint() as session:
        pres = open_presentation(session, p, read_only=True)
        try:
            try:
                # PrintRange=None must be passed EXPLICITLY: the parameter is
                # a VT_DISPATCH pointer and pywin32 marshals its omitted-value
                # placeholder into "The Python instance can not be converted
                # to a COM object" (verified empirically).
                pres.ExportAsFixedFormat(
                    str(target.resolve()),
                    PP_FIXED_FORMAT_TYPE_PDF,
                    Intent=PP_FIXED_FORMAT_INTENT_PRINT,
                    FrameSlides=-1 if not include_notes else 0,  # msoTriState
                    HandoutOrder=PP_PRINT_HANDOUT_HORIZONTAL_FIRST,
                    OutputType=output_type,
                    PrintRange=None,
                )
            except Exception as exc:
                _raise_classified(exc, "handout PDF export failed")
        finally:
            del pres  # release the proxy so the launched instance can exit
    _publish_export(target, out, "handout PDF")
    return {
        "pdf": str(out),
        "bytes": out.stat().st_size,
        "engine": "powerpoint-com",
        "layout": "notes_pages" if include_notes else f"{slides_per_page}_per_page",
    }


@_bounded_op("com_export_slide_images", default=60.0)
def com_export_slide_images(
    path: str,
    output_dir: str | None = None,
    slides: list[int] | None = None,
    width: int = DEFAULT_IMAGE_WIDTH,
    height: int | None = None,
) -> dict:
    """Export slides as PNG images via Slide.Export (the render-to-verify
    primitive).

    `slides` takes 0-based presentation-order indices (matching the file
    layer); None means every slide. Default width 1280px; height derives
    from the deck's slide aspect ratio unless given. Files are named
    slideN.png with N the 1-based slide position.

    `output_dir` is a folder a caller named, checked against
    KS4P_ALLOWED_ROOTS, or a ServerScratch the server made for its own
    intermediate render (export_shape_image), used as it is.
    """
    p = _require_file(path, "slide image export source")
    if isinstance(output_dir, ServerScratch):
        out_dir = Path(output_dir.path)
    else:
        if output_dir:
            output_dir = check_path(output_dir, "slide image output directory")
        out_dir = Path(output_dir) if output_dir else p.with_name(p.stem + "_slides")
    out_dir.mkdir(parents=True, exist_ok=True)
    if width <= 0:
        raise PptMcpError(f"width must be positive, got {width}")
    # Validate `slides` BEFORE launching PowerPoint. A string is iterable,
    # so list("C:/deck.pptx") used to become a per-character slide list and
    # the caller got "slide index C out of range" after a full launch-open
    # cycle (seen in use: 8.7 seconds to produce a nonsense
    # message). Refuse in microseconds with a message that names the actual
    # mistake.
    if slides is not None:
        if isinstance(slides, (str, bytes)):
            raise PptMcpError(
                "slides must be a list of 0-based integers, got a string "
                f"({slides!r}); a string would be read one character per "
                "slide. Pass [0, 1, 2], or None for every slide."
            )
        if not isinstance(slides, (list, tuple, set, frozenset, range)):
            raise PptMcpError(
                "slides must be a list of 0-based integers, got "
                f"{type(slides).__name__}. Pass [0, 1, 2], or None for "
                "every slide."
            )
        for idx in slides:
            if isinstance(idx, bool) or not isinstance(idx, int):
                raise PptMcpError(
                    "slides must contain 0-based integers, found "
                    f"{idx!r} ({type(idx).__name__})"
                )
    images: list[dict] = []
    with _powerpoint() as session:
        pres = open_presentation(session, p, read_only=True)
        try:
            try:
                slide_count = int(pres.Slides.Count)
                if height is None:
                    sw = float(pres.PageSetup.SlideWidth)
                    sh = float(pres.PageSetup.SlideHeight)
                    height = max(1, round(width * sh / sw))
                wanted = (
                    list(range(slide_count)) if slides is None else list(slides)
                )
                for idx in wanted:
                    if not isinstance(idx, int) or idx < 0 or idx >= slide_count:
                        raise PptMcpError(
                            f"slide index {idx} out of range, presentation "
                            f"has {slide_count} (indices are 0-based)"
                        )
                for idx in wanted:
                    out_file = out_dir / f"slide{idx + 1}.png"
                    pres.Slides.Item(idx + 1).Export(
                        str(out_file.resolve()), "PNG", width, height
                    )
                    _verify_output(out_file, f"PNG for slide {idx + 1}")
                    images.append({"slide_number": idx + 1, "slide": idx,
                                   "file": str(out_file)})
            except PptMcpError:
                raise
            except Exception as exc:
                _raise_classified(exc, "slide image export failed")
        finally:
            del pres  # release the proxy so the launched instance can exit
    return {
        "images": images,
        "width": width,
        "height": height,
        "engine": "powerpoint-com",
    }


@_bounded_op("com_validate_opens_clean", default=60.0)
def com_validate_opens_clean(path: str) -> dict:
    """Open in invisible PowerPoint (alerts disabled) and force a FULL content
    load: every slide, every top-level shape, every group member, every table
    cell, and a text read of every text frame among them. Corruption surfaces
    on access, not on open; with DisplayAlerts off it raises instead of
    hanging on a modal repair dialog. A busy or disconnected PowerPoint is
    reported as itself, never as a verdict on the file."""
    p = _require_file(path, "validate opens clean")
    with _powerpoint() as session:
        try:
            pres = open_presentation(session, p, read_only=True)
        except (DocumentLocked, PowerPointBusy, PowerPointDisconnected):
            raise  # environment problems, not a verdict on the file
        except PptMcpError as exc:
            return {"opens_clean": False, "error": str(exc)}
        try:
            try:
                # _full_load in its own frame so slide/shape proxies are
                # released on return (outstanding proxies block app exit).
                stats = _full_load(pres)
            except (DocumentLocked, PowerPointBusy, PowerPointDisconnected):
                # The APPLICATION went sideways mid-walk. Reporting that as
                # a corrupt file would be an authoritative-looking lie about
                # a deck that is fine.
                raise
            except FullLoadFailed as exc:
                return _opens_clean_failure(exc)
            except Exception as exc:  # full-load failure = not clean
                return {"opens_clean": False, "error": str(exc)}
        finally:
            del pres  # release the proxy so the launched instance can exit
    return {"opens_clean": True, **stats}


#: The opens-clean check's time limit grows with what it has to do. Most of
#: it is the walk: PowerPoint loads each shape the first time it is touched,
#: about 23 ms a shape on PowerPoint 16 (1,636 shapes over 50 slides in 33 to
#: 39 s; reading fewer properties a shape did not change that), while start,
#: open and close took 5 to 8 s and notes pages added nothing measurable.
#: A fixed 60 s stopped a healthy deck of 47 slides and about 1,400 shapes,
#: so the limit is the old 60 s plus about two and a half times the measured
#: cost per shape, an allowance per slide, notes page and megabyte of media
#: (a busy computer is slower), and never more than the cap.
OPENS_CLEAN_BASE_S = 60.0
OPENS_CLEAN_PER_SLIDE_S = 0.25
OPENS_CLEAN_PER_SHAPE_S = 0.06
OPENS_CLEAN_PER_NOTES_S = 0.1
OPENS_CLEAN_PER_MEDIA_MB_S = 0.2
#: The most the check ever waits.
OPENS_CLEAN_CAP_S = 300.0


def opens_clean_time_limit(
    slides: int, shapes: int, notes: int = 0, media_bytes: int = 0
) -> tuple[float, bool]:
    """(seconds, capped) for the opens-clean check of a deck this size;
    capped is True when the cap, not the deck, set the limit."""
    raw = (
        OPENS_CLEAN_BASE_S
        + max(0, slides) * OPENS_CLEAN_PER_SLIDE_S
        + max(0, shapes) * OPENS_CLEAN_PER_SHAPE_S
        + max(0, notes) * OPENS_CLEAN_PER_NOTES_S
        + max(0, media_bytes) / 1_000_000 * OPENS_CLEAN_PER_MEDIA_MB_S
    )
    return round(min(OPENS_CLEAN_CAP_S, raw), 1), raw >= OPENS_CLEAN_CAP_S


class FullLoadFailed(Exception):
    """A full content load that failed on a KNOWN slide, and where the XML
    allowed it, a known shape.

    Before this, a refusal came back as one string with no coordinates, and
    finding the bad shape in a 45-slide deck was about 40 minutes of manual
    bisection.
    """

    def __init__(
        self,
        message: str,
        *,
        slide_index: int | None = None,
        slide_id: int | None = None,
        shape_index: int | None = None,
        shape_name: str | None = None,
    ) -> None:
        super().__init__(message)
        self.slide_index = slide_index
        self.slide_id = slide_id
        self.shape_index = shape_index
        self.shape_name = shape_name


def _opens_clean_failure(exc: FullLoadFailed) -> dict:
    """The not-clean verdict, carrying whatever identity the load knew.
    The validate tool returns this dict as its `powerpoint` block."""
    out: dict = {"opens_clean": False, "error": str(exc)}
    for key, value in (
        ("failed_slide_index", exc.slide_index),
        ("failed_slide_id", exc.slide_id),
        ("failed_shape_index", exc.shape_index),
        ("failed_shape_name", exc.shape_name),
    ):
        if value is not None:
            out[key] = value
    return out


def _slide_id_of(slide) -> int | None:
    """A slide's SlideID, or None when even that cannot be read (which is
    itself a symptom, and must not mask the real failure).

    A COM fault meaning the APPLICATION is busy or gone is classified and
    re-raised first. Swallowing every exception here let a modal dialog or
    a dead proxy pass as "this slide has no readable id", and the walk then
    carried on and produced a verdict about the FILE. Only an ordinary read
    failure returns None."""
    try:
        return int(slide.SlideID)
    except Exception as exc:
        _reraise_environment(exc)
        return None


#: MsoTriState. msoTrue is -1; msoTriStateMixed is -2 and is NOT a yes.
#: Truthiness treated mixed as true, which probed an unsupported TextFrame
#: on an unusual shape and turned it into a false failure.
MSO_TRUE = -1

#: MsoShapeType msoGroup.
MSO_GROUP = 6

#: Ceiling on how many shapes one walk visits. A deck past this is walked
#: as far as the cap and the result says so, rather than sitting inside
#: the bounded operation until it times out and reports nothing at all.
FULL_LOAD_MAX_SHAPES = 20000


def _mso_true(value) -> bool:
    """An MsoTriState that really is msoTrue."""
    try:
        return int(value) == MSO_TRUE
    except (TypeError, ValueError):
        return False


def _optional(obj, name):
    """A property that may not exist on this object at all. A COM fault
    (busy, disconnected) still propagates; only a missing member is
    swallowed, because that is the object saying 'not applicable'."""
    try:
        return getattr(obj, name)
    except AttributeError:
        return None


class _WalkState:
    """Running totals for one full-load walk."""

    def __init__(self) -> None:
        self.shapes_walked = 0
        self.text_reads = 0
        self.truncated = False

    @property
    def room(self) -> bool:
        if self.shapes_walked >= FULL_LOAD_MAX_SHAPES:
            self.truncated = True
            return False
        return True


def _full_load(pres) -> dict:
    """Force a full content load. Corruption surfaces on access, not on
    open, so every slide and every shape is TOUCHED and every text frame
    is READ.

    Boundary, stated exactly because the old docstring promised more than
    it did: top-level shapes on every slide, the
    members of every group recursively, and the cells of every table.
    Notes pages, layouts, masters, chart internals and SmartArt data
    models are NOT walked; the layer 1 package check covers their text
    bodies. The old walk read ONE text frame for the whole deck, so a
    shape that raised on access after the first successful read was never
    touched and the deck came back clean.

    Every read sits inside its own try/except, so a refusal names WHERE it
    happened: 1-based slide index, SlideID, 1-based top-level shape index,
    shape name (and the nested path when the failure is inside a group or
    a table). A COM fault that means the APPLICATION is busy or gone is
    re-raised as itself rather than blamed on the file.
    """
    state = _WalkState()
    try:
        slide_count = int(pres.Slides.Count)
    except Exception as exc:
        # The FIRST access is a COM access like every other one in this
        # walk, and it used to sit outside the classification: a busy or
        # disconnected PowerPoint faulting here reached the generic
        # validate catch and came back as opens_clean false on a deck
        # nobody had looked at yet. No
        # slide has been reached, so the refusal carries no coordinates.
        _reraise_environment(exc)
        raise FullLoadFailed(
            f"the slide collection could not be read: {exc}"
        ) from exc
    shapes_total = 0
    for i in range(1, slide_count + 1):
        slide_id = None
        try:
            slide = pres.Slides.Item(i)
            slide_id = _slide_id_of(slide)
            shape_count = int(slide.Shapes.Count)
        except Exception as exc:
            _reraise_environment(exc)
            raise FullLoadFailed(
                f"slide {i} (id {slide_id}) failed to load: {exc}",
                slide_index=i, slide_id=slide_id,
            ) from exc
        shapes_total += shape_count
        for j in range(1, shape_count + 1):
            if not state.room:
                break
            _walk_shape(slide.Shapes, j, i, slide_id, j, "", state)
        if state.truncated:
            break
    out = {
        "slides": slide_count,
        "shapes": shapes_total,
        "shapes_walked": state.shapes_walked,
        "text_reads": state.text_reads,
    }
    if state.truncated:
        out["walk_truncated"] = (
            f"the walk stopped after {state.shapes_walked} shapes "
            f"(cap {FULL_LOAD_MAX_SHAPES}); shapes past that point were "
            "not checked"
        )
    return out


def _reraise_environment(exc: BaseException) -> None:
    """Re-raise a COM fault that means the APPLICATION is busy or gone.

    Without this, PowerPoint showing a dialog halfway through the walk
    came back as an authoritative corruption verdict naming an innocent
    shape. The classifier and its HRESULT tables
    already existed; the walk simply never consulted them.
    """
    if isinstance(
        exc, (DocumentLocked, PowerPointBusy, PowerPointDisconnected)
    ):
        raise exc
    typed = _classify(exc)
    if typed is not None:
        raise typed from exc


def _walk_shape(
    collection, index: int, slide_index: int, slide_id, top_index: int,
    path: str, state: _WalkState,
) -> None:
    """Touch one shape, read its text, and descend into groups and tables."""
    shape_name = None
    where = f"slide {slide_index} (id {slide_id}) shape {top_index}"
    try:
        shp = collection.Item(index)
        state.shapes_walked += 1
        shape_name = _optional(shp, "Name")
        shape_name = str(shape_name) if shape_name is not None else None
        here = f"{path}/{shape_name}" if path else (shape_name or "")
        if _mso_true(_optional(shp, "HasTextFrame")):
            frame = shp.TextFrame
            if _mso_true(_optional(frame, "HasText")):
                _ = frame.TextRange.Text
                state.text_reads += 1
        if _mso_true(_optional(shp, "HasTable")):
            _walk_table(shp.Table, slide_index, slide_id, top_index, here,
                        state)
        if _optional(shp, "Type") == MSO_GROUP:
            members = shp.GroupItems
            count = int(members.Count)
            for k in range(1, count + 1):
                if not state.room:
                    return
                _walk_shape(members, k, slide_index, slide_id, top_index,
                            here, state)
    except FullLoadFailed:
        raise
    except Exception as exc:
        _reraise_environment(exc)
        named = f" {shape_name!r}" if shape_name else ""
        inside = f" inside {path!r}" if path else ""
        raise FullLoadFailed(
            f"{where}{named}{inside} failed to load: {exc}",
            slide_index=slide_index, slide_id=slide_id,
            shape_index=top_index, shape_name=shape_name,
        ) from exc


def _walk_table(
    table, slide_index: int, slide_id, top_index: int, path: str,
    state: _WalkState,
) -> None:
    """Read every cell of a table. A corrupt cell body is exactly the
    shape of the failure this check exists to find."""
    rows = int(table.Rows.Count)
    cols = int(table.Columns.Count)
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            if not state.room:
                return
            try:
                cell_shape = table.Cell(r, c).Shape
                state.shapes_walked += 1
                if _mso_true(_optional(cell_shape, "HasTextFrame")):
                    frame = cell_shape.TextFrame
                    if _mso_true(_optional(frame, "HasText")):
                        _ = frame.TextRange.Text
                        state.text_reads += 1
            except Exception as exc:
                _reraise_environment(exc)
                raise FullLoadFailed(
                    f"slide {slide_index} (id {slide_id}) shape {top_index} "
                    f"{path!r} table cell (row {r}, column {c}) failed to "
                    f"load: {exc}",
                    slide_index=slide_index, slide_id=slide_id,
                    shape_index=top_index, shape_name=path or None,
                ) from exc


def _installed_version() -> str | None:
    """PowerPoint's registered version from the registry; no COM attach."""
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CurVer"
        ) as key:
            curver, _ = winreg.QueryValueEx(key, "")
        # "PowerPoint.Application.16" -> "16"
        return str(curver).rsplit(".", 1)[-1] if curver else None
    except OSError:
        return None
    except ImportError:  # pragma: no cover - non-Windows
        return None


def powerpoint_installed() -> bool:
    return _installed_version() is not None


def powerpoint_status() -> dict:
    """Diagnostic: installed? running? blocked by a modal dialog? is a COM
    operation running right now? which presentations are open in the user's
    instance (paths only, no content)?

    Open presentations are read from the Running Object Table by DISPLAY NAME
    only (open Office documents register a file moniker); nothing is bound
    and the user's instance is never attached to (no GetActiveObject).

    Three additions, all answerable without touching COM:
    - pending_dialogs / blocked: modal dialogs read at the OS WINDOW layer.
      COM cannot report these, because a modal dialog is exactly what stops
      COM from answering. A status tool that says "ready" while the user
      stares at an alert cascade is the bug this fixes.
    - com_serialization: whether a COM operation currently holds the lock,
      so a caller learns that its next call would QUEUE rather than run.
    - The ROT probe itself runs under a bounded try-acquire, so a long
      operation cannot hang the status call; when the lock is busy the
      probe is skipped and the result says so.
    """
    out: dict = {
        "installed": powerpoint_installed(),
        "version": _installed_version(),
        "powerpoint_running": False,
        "open_presentations": [],
    }
    try:
        pids = powerpnt_pids()
    except Exception as exc:
        out["error"] = f"process table check failed: {exc}"
        out["com_serialization"] = _serial.lock_snapshot()
        return out
    if pids is None:
        # Unknown, which is not the same answer as "not running" and must
        # not be dressed up as one.
        out["error"] = "process table check failed: tasklist did not answer"
        out["com_serialization"] = _serial.lock_snapshot()
        return out
    out["powerpoint_running"] = bool(pids)

    # Window layer first: it works even when COM is wedged, so it is what
    # makes a "blocked" verdict trustworthy.
    pending: list = []
    with contextlib.suppress(Exception):
        from . import dialogs as _dialogs

        pending = _dialogs.pending_dialogs(pids=pids)
    out["pending_dialogs"] = pending
    out["blocked"] = bool(pending)
    if pending:
        out["blocked_note"] = (
            "PowerPoint has a modal dialog open; COM calls will be rejected "
            "or will hang until it is dismissed. Dismiss it in PowerPoint "
            "(this server never clicks a dialog for you), then retry."
        )
    out["com_serialization"] = _serial.lock_snapshot()

    if not pids:
        return out
    if not _serial.acquire(timeout=2.0):
        out["note"] = (
            "COM serialization lock held by a running operation; open "
            "presentations were not probed"
        )
        return out
    try:
        exts = (".pptx", ".pptm", ".ppt", ".ppsx", ".potx", ".potm")
        try:
            pythoncom, _win32com = _com_modules()
        except PptMcpError as exc:
            out["error"] = str(exc)
            return out
        out["com_apartment"] = _ensure_apartment(pythoncom)
        with contextlib.suppress(Exception):
            rot = pythoncom.GetRunningObjectTable()
            for moniker in rot.EnumRunning():
                ctx = pythoncom.CreateBindCtx(0)
                try:
                    name = moniker.GetDisplayName(ctx, None)
                except Exception:
                    continue
                # Names only; NEVER GetObject/bind (binding the wrong
                # moniker kind can be a native crash, and binding at all
                # would attach to the user's instance).
                if isinstance(name, str) and name.lower().endswith(exts):
                    out["open_presentations"].append(name)
        out["open_presentation_count"] = len(out["open_presentations"])
        return out
    finally:
        _serial.release()


# The bounded try-acquire above is the serialization contract for this
# function; the marker lets the entry-point audit see it.
powerpoint_status._com_serialized = "powerpoint_status"


def zombie_check() -> dict:
    """Count POWERPNT.EXE processes (leak detection diagnostic); -1 when
    the process table could not be read."""
    return {"powerpnt_processes": powerpnt_count()}
