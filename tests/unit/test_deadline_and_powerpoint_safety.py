"""Regression checks for the deadline and PowerPoint safety fixes."""

from __future__ import annotations

import ast
import asyncio
import subprocess
import sys
import threading
import time
from pathlib import Path

import pytest

from kitchensink4ppt import registry
from kitchensink4ppt.core.errors import OpError
from kitchensink4ppt.ops import backups
from kitchensink4ppt.surface.dispatch import run_bounded


def test_purge_claims_before_its_first_delete(tmp_path, monkeypatch):
    slot = tmp_path / "slots"
    slot.mkdir()
    (slot / "anchor.pptx").write_bytes(b"anchor")
    (slot / "prev.pptx").write_bytes(b"prev")
    collected = threading.Event()

    def slow_collect(*_args, **_kwargs):
        time.sleep(0.05)
        collected.set()
        return [slot], None

    monkeypatch.setattr(backups, "_collect_purge_targets", slow_collect)
    with pytest.raises(OpError) as caught:
        asyncio.run(run_bounded(
            lambda: backups.purge_backups("orphans", directory=str(tmp_path), dry_run=False),
            0.01,
        ))
    assert caught.value.code == "TIMEOUT"
    assert collected.wait(1)
    assert (slot / "anchor.pptx").exists()
    assert (slot / "prev.pptx").exists()


def test_every_live_deck_mutator_claims_before_its_first_known_com_mutation():
    path = Path(__file__).resolve().parents[2] / "src/kitchensink4ppt/com/live_ops.py"
    tree = ast.parse(path.read_text(encoding="utf-8"))
    mutators = {}
    for node in tree.body:
        if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        run_live = [
            child
            for child in ast.walk(node)
            if isinstance(child, ast.Call)
            and isinstance(child.func, ast.Attribute)
            and isinstance(child.func.value, ast.Name)
            and child.func.value.id == "live"
            and child.func.attr == "run_live"
        ]
        if not run_live:
            continue
        is_read_only = any(
            keyword.arg == "mutating"
            and isinstance(keyword.value, ast.Constant)
            and keyword.value.value is False
            for call in run_live
            for keyword in call.keywords
        )
        if not is_read_only:
            mutators[node.name] = node

    assert mutators, "no mutating live.run_live entry points were discovered"
    known_mutating_methods = {"AddShape", "AddTextbox", "AddSlide", "Delete", "Save"}
    for name, node in sorted(mutators.items()):
        calls = [
            child
            for child in ast.walk(node)
            if isinstance(child, ast.Call)
            and isinstance(child.func, ast.Attribute)
            and isinstance(child.func.value, ast.Name)
            and child.func.value.id == "safesave"
            and child.func.attr == "claim_publish"
        ]
        mutations = []
        for child in ast.walk(node):
            if isinstance(child, (ast.Assign, ast.AnnAssign, ast.AugAssign)):
                targets = child.targets if isinstance(child, ast.Assign) else [child.target]
                if any(isinstance(target, ast.Attribute) for target in targets):
                    mutations.append(child)
            if (
                isinstance(child, ast.Call)
                and isinstance(child.func, ast.Attribute)
                and (
                    child.func.attr in known_mutating_methods
                    or (
                        isinstance(child.func.value, ast.Name)
                        and child.func.value.id == "live"
                        and child.func.attr == "set_text_chunked"
                    )
                )
            ):
                mutations.append(child)
        assert calls, f"{name} has no publication claim"
        assert mutations, f"{name} has no recognized COM mutation"
        assert min(call.lineno for call in calls) < min(
            mutation.lineno for mutation in mutations
        ), f"{name} claims only after its first COM mutation"


def test_com_bound_operations_use_bridge_default_timeouts():
    # Importing the server populates the fixed registry through @_tool.
    import kitchensink4ppt.server  # noqa: F401

    specs = registry.public_ops()
    expected = {
        "export_pdf": 120.0,
        "export_handout": 120.0,
        "export_slide_image": 120.0,
        "validate": 360.0,  # the opens-clean cap (300 s) plus the same room
    }
    assert {name: specs[name].timeout_s for name in expected} == expected


def test_many_stuck_calls_do_not_starve_a_fast_call():
    release = threading.Event()
    all_entered = threading.Event()
    entered_lock = threading.Lock()
    entered = 0

    def stuck():
        nonlocal entered
        with entered_lock:
            entered += 1
            if entered == 40:
                all_entered.set()
        release.wait()
        return {"late": True}

    async def exercise():
        calls = [
            asyncio.create_task(run_bounded(stuck, 1.0))
            for _ in range(40)
        ]
        assert await asyncio.to_thread(all_entered.wait, 30)
        outcomes = await asyncio.gather(*calls, return_exceptions=True)
        assert all(isinstance(item, OpError) for item in outcomes)
        return await run_bounded(lambda: {"fast": True}, 10.0)

    try:
        assert asyncio.run(exercise()) == {"fast": True}
    finally:
        release.set()


def test_timed_out_worker_does_not_hold_process_exit():
    script = """
import asyncio
import threading
import kitchensink4ppt.surface.dispatch as dispatch
from kitchensink4ppt.core.errors import OpError
from kitchensink4ppt.surface.dispatch import run_bounded

dispatch.EXIT_DRAIN_S = 0.05
gate = threading.Event()

async def main():
    try:
        await run_bounded(gate.wait, 0.01)
    except OpError:
        pass

asyncio.run(main())
"""
    completed = subprocess.run(
        [sys.executable, "-c", script],
        check=False,
        capture_output=True,
        text=True,
        timeout=5,
    )
    assert completed.returncode == 0, completed.stderr
    assert "still running after 0.05s at exit" in completed.stderr


def test_shutdown_cancels_uncommitted_calls_and_waits_for_cleanup():
    import kitchensink4ppt.surface.dispatch as dispatch

    release = threading.Event()
    entered = threading.Event()
    outcome = {}

    def late_publish():
        entered.set()
        release.wait(2)
        try:
            from kitchensink4ppt.core import safesave

            safesave.claim_publish("the late test write")
            outcome["state"] = "published"
        except OpError as exc:
            outcome["state"] = exc.code
        return {"done": True}

    async def exercise():
        deadline = dispatch.call_deadline.Deadline(30)

        def invoke():
            with dispatch.call_deadline.active(deadline):
                return late_publish()

        future = dispatch._start_worker(invoke, deadline)
        assert entered.wait(30)
        drain = asyncio.create_task(asyncio.to_thread(dispatch.drain_inflight, 5))
        end = time.monotonic() + 5
        while not deadline.cancelled and time.monotonic() < end:
            await asyncio.sleep(0.001)
        assert deadline.cancelled
        release.set()
        report = await drain
        await future
        return report

    report = asyncio.run(exercise())
    assert report == {"cancelled": 1, "waited": 1, "unfinished": 0}
    assert outcome == {"state": "TIMEOUT"}


def test_shutdown_preserves_a_committed_call_and_waits_for_its_result():
    import kitchensink4ppt.surface.dispatch as dispatch
    from kitchensink4ppt.core import safesave

    release = threading.Event()
    committed = threading.Event()

    def finishing_save():
        safesave.claim_publish("the test save")
        committed.set()
        release.wait(2)
        return {"saved": True}

    async def exercise():
        deadline = dispatch.call_deadline.Deadline(30)

        def invoke():
            with dispatch.call_deadline.active(deadline):
                return finishing_save()

        future = dispatch._start_worker(invoke, deadline)
        assert committed.wait(30)
        drain_started = threading.Event()

        def drain_now():
            drain_started.set()
            return dispatch.drain_inflight(5)

        drain = asyncio.create_task(asyncio.to_thread(drain_now))
        assert await asyncio.to_thread(drain_started.wait, 5)
        release.set()
        report = await drain
        return report, await future

    report, result = asyncio.run(exercise())
    assert report == {"cancelled": 0, "waited": 1, "unfinished": 0}
    assert result == {"saved": True}


def test_shutdown_wait_is_one_total_bound_and_ends_nothing(capsys):
    import kitchensink4ppt.surface.dispatch as dispatch

    release = threading.Event()

    async def exercise():
        deadline = dispatch.call_deadline.Deadline(30)

        def invoke():
            with dispatch.call_deadline.active(deadline):
                return release.wait(2)

        future = dispatch._start_worker(invoke, deadline)
        await asyncio.sleep(0.02)
        started = time.monotonic()
        report = await asyncio.to_thread(dispatch.drain_inflight, 0.05)
        elapsed = time.monotonic() - started
        alive = [
            thread
            for thread in threading.enumerate()
            if thread.name.startswith(dispatch.WORKER_PREFIX) and thread.is_alive()
        ]
        release.set()
        await future
        return report, elapsed, alive

    report, elapsed, alive = asyncio.run(exercise())
    assert report == {"cancelled": 1, "waited": 1, "unfinished": 1}
    assert elapsed < 0.5
    assert alive, "the bounded drain must not end a worker thread"
    stderr = capsys.readouterr().err
    assert stderr.count("\n") == 1
    assert "no process was force-ended" in stderr


def test_server_main_drains_inflight_calls_when_stdio_stops(monkeypatch):
    import kitchensink4ppt.server as server

    events = []

    class StubServer:
        def run(self, **kwargs):
            events.append(("run", kwargs))

    monkeypatch.setattr(server.Settings, "from_env", lambda: server.Settings())
    monkeypatch.setattr(server, "create_server", lambda _settings: StubServer())
    monkeypatch.setattr(
        server._surface_dispatch,
        "drain_inflight",
        lambda: events.append(("drain", {})),
    )

    server.main()

    assert events == [
        (
            "run",
            {
                "transport": "stdio",
                "show_banner": False,
                "log_level": "CRITICAL",
            },
        ),
        ("drain", {}),
    ]
