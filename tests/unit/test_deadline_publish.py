"""The surface timeout and final publication point make one decision."""

from __future__ import annotations

import asyncio
import concurrent.futures
import threading
import time

import pytest

from kitchensink4ppt.core import deadline, safesave
from kitchensink4ppt.core.errors import OpError
from kitchensink4ppt.surface.dispatch import run_bounded


def test_timeout_before_publish_leaves_the_target_unchanged(tmp_path):
    target = tmp_path / "target.txt"
    target.write_text("before", encoding="utf-8")

    def late_worker():
        time.sleep(0.05)
        safesave.claim_publish("the test write")
        target.write_text("after", encoding="utf-8")
        return {"file": str(target)}

    with pytest.raises(OpError) as caught:
        asyncio.run(run_bounded(late_worker, 0.01))
    assert caught.value.code == "TIMEOUT"
    assert caught.value.details == {"bound_s": 0.01, "saved": False}
    assert target.read_text(encoding="utf-8") == "before"


def test_publish_before_timeout_returns_the_real_result(tmp_path, monkeypatch):
    target = tmp_path / "target.txt"
    target.write_text("before", encoding="utf-8")
    committed = threading.Event()

    def committed_worker():
        safesave.claim_publish("the test write")
        committed.set()
        target.write_text("after", encoding="utf-8")
        return {"file": str(target), "saved": True}

    async def timeout_after_commit(awaitable, *, timeout):
        while not committed.is_set():
            await asyncio.sleep(0)
        raise TimeoutError

    monkeypatch.setattr(asyncio, "wait_for", timeout_after_commit)
    result = asyncio.run(run_bounded(committed_worker, 1.0))
    assert result["saved"] is True
    assert target.read_text(encoding="utf-8") == "after"


def test_committed_deadline_stays_committed_after_the_clock_expires():
    call = deadline.Deadline(0.01)
    assert call.try_commit() is True
    time.sleep(0.02)
    assert call.try_commit() is True
    assert call.cancel() is False


def test_call_does_not_wait_for_the_shared_default_executor(tmp_path):
    target = tmp_path / "queued.txt"
    gate = threading.Event()
    occupied = threading.Event()

    def hold_only_worker():
        occupied.set()
        gate.wait(2)

    def queued_worker():
        safesave.claim_publish("the queued write")
        target.write_text("late", encoding="utf-8")
        return {"file": str(target)}

    async def exercise():
        loop = asyncio.get_running_loop()
        pool = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        loop.set_default_executor(pool)
        blocker = asyncio.create_task(asyncio.to_thread(hold_only_worker))
        while not occupied.is_set():
            await asyncio.sleep(0)
        try:
            result = await run_bounded(queued_worker, 0.5)
            assert result == {"file": str(target)}
        finally:
            gate.set()
            await blocker
            await asyncio.sleep(0.01)
            pool.shutdown(wait=True)

    asyncio.run(exercise())
    assert target.read_text(encoding="utf-8") == "late"


def test_post_publish_wait_is_bounded_and_reports_uncertain_save(monkeypatch):
    import kitchensink4ppt.surface.dispatch as dispatch

    release = threading.Event()

    def stuck_after_publish():
        safesave.claim_publish("the test write")
        release.wait(2)
        return {"saved": True}

    monkeypatch.setattr(dispatch, "MIN_POST_COMMIT_WAIT_S", 0.01)
    try:
        with pytest.raises(OpError) as caught:
            # Leave enough scheduling margin for the worker to reach the
            # publication point; this test is about the post-commit wait,
            # not a race over whether the commit happened before expiry.
            asyncio.run(run_bounded(stuck_after_publish, 0.5))
        assert caught.value.code == "TIMEOUT"
        assert caught.value.details == {"bound_s": 0.5, "saved": None}
    finally:
        release.set()
