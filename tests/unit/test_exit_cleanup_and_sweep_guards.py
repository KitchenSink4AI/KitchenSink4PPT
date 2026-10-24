"""Regression tests for exit-time PowerPoint handling and the release path.

Each test's docstring gives the behaviour it pins and why. The COM-gated
test skips unless KS4_VERIFY_PPT_COM=1 is set in a window where no
PowerPoint is running.
"""

from __future__ import annotations

import ast
import os
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]


@pytest.mark.skipif(os.name != "nt", reason="PowerPoint COM is Windows-only")
@pytest.mark.skipif(os.environ.get("KS4_VERIFY_PPT_COM") != "1",
                    reason="needs PowerPoint in a clean window: set KS4_VERIFY_PPT_COM=1")
def test_s1_the_exit_reaper_can_identify_its_own_hidden_powerpoint():
    """The exit reaper can identify its own hidden PowerPoint.
    _inspect_exit_ownership must not match the process through
    `Application.HWND`, which a hidden PowerPoint with no window does not
    have ('Member not found'; late-bound pywin32 hands back a bound method,
    int() of it raises, the helper returns None). Every exit inspection
    would then answer 'the active PowerPoint PID could not be matched' and
    the instance this server started would never be ended, so the
    disconnect-during-a-COM-call test would fail again."""
    import gc
    import time

    import pythoncom
    import win32com.client

    from kitchensink4ppt.com import bridge

    if bridge.powerpnt_pids():
        pytest.skip("a PowerPoint is already running")
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        started = bridge.powerpnt_pids() or set()
        assert len(started) == 1
        assert bridge._powerpoint_app_pid(app) == next(iter(started))
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        app = None
        gc.collect()
        time.sleep(3)


def test_s2_the_non_live_op_sweep_never_reaches_powerpoint():
    """The non-live op sweep never reaches PowerPoint. The unmarked sweep in
    test_family_op_sweep_gate.py must not dispatch validate,
    export_slide_image (engine='com'), export_handout and the live ops with
    no live marker or COM gate: with a PowerPoint already running it
    attaches to that instance (in every run it opened validate.pptx /
    export_slide_image.pptx inside a held stand-in instance and switched
    its DisplayAlerts, once leaving it at ppAlertsNone). A `-m "not live"`
    run on a user's desk must not touch their PowerPoint."""
    text = (ROOT / "tests" / "unit" / "test_family_op_sweep_gate.py").read_text(encoding="utf-8")
    tree = ast.parse(text)
    gated = "pytest.mark.live" in text or "_com_gate" in text or "needs_closed_powerpoint" in text \
        or "KS4_SWEEP_PPT_COM" in text
    assert gated, "the non-live sweep runs PowerPoint COM ops without a gate"
    assert tree is not None


def test_s3_the_release_workflow_builds_owner_test_from_a_git_archive():
    """The release workflow builds owner-test artifacts from a git archive.
    build_editions.py --owner-test refuses unless
    scripts/ARCHIVE_REVISION.txt was substituted by `git archive`, so
    publish.yml must not run it in an actions/checkout tree: the release
    build could never pass once the copy is filled."""
    wf = (ROOT / ".github" / "workflows" / "publish.yml").read_text(encoding="utf-8")
    build_at = wf.index("scripts/build_editions.py")
    assert "git archive" in wf[:build_at], "publish.yml builds owner-test artifacts from a checkout"


def test_s4_exit_cleanup_honours_unsaved_server_opened_presentations():
    """Exit cleanup honours unsaved server-opened presentations. The reaper
    ends PowerPoint 'only when no foreign or unsaved presentation remains',
    so _inspect_exit_ownership must consult `Saved` before closing a
    server-owned presentation and ending the process."""
    src = (ROOT / "src" / "kitchensink4ppt" / "com" / "bridge.py").read_text(encoding="utf-8")
    tree = ast.parse(src)
    fn = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "_inspect_exit_ownership")
    assert ".Saved" in ast.get_source_segment(src, fn)
