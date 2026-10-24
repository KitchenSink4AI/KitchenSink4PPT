"""A read that lands inside another server's save is a typed answer (2.0.0rc3).

PptxPackage._load checks the lock, then opens the zip. A save by another
server that replaces the file between the two makes the open raise a bare
PermissionError, which nothing typed, so the reader was told the server had
failed (INTERNAL_ERROR, 1 read in 565 under contention). The replace takes
milliseconds: the load now looks again, and a file still being written is
DOCUMENT_LOCKED with a message that says so. It is not "open in PowerPoint",
so a live='auto' call does not go looking for the deck there.
"""

from __future__ import annotations

import asyncio
import sys

import pytest

from kitchensink4ppt import server  # noqa: F401 - populates the registry
from kitchensink4ppt.core import package
from kitchensink4ppt.registry import public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.plan import build


def _call(tool: str, op: str, args: dict):
    plan = build(Settings(), public_ops())
    return asyncio.run(dispatch(plan, tool, {"op": op, "args": args}))


def _flaky_zip(monkeypatch, deck: str, failures: int | None):
    """ZipFile on the deck raises the save race's PermissionError, for the
    first ``failures`` opens (None: every open)."""
    real = package.zipfile.ZipFile
    seen = {"n": 0}

    def flaky(file, *args, **kwargs):
        if str(file) == deck:
            seen["n"] += 1
            if failures is None or seen["n"] <= failures:
                raise PermissionError(13, "Permission denied", deck)
        return real(file, *args, **kwargs)

    monkeypatch.setattr(package.zipfile, "ZipFile", flaky)
    return seen


def test_a_read_that_meets_a_save_once_looks_again_and_succeeds(make_deck, monkeypatch):
    deck = str(make_deck())
    seen = _flaky_zip(monkeypatch, deck, failures=1)
    payload, is_error = _call("ppt_read", "get_presentation_view", {"file_path": deck})
    assert is_error is False, payload
    assert seen["n"] == 2
    assert payload["slide_count"] >= 1


def test_a_read_that_keeps_meeting_a_save_is_document_locked(make_deck, monkeypatch):
    deck = str(make_deck())
    _flaky_zip(monkeypatch, deck, failures=None)
    payload, is_error = _call("ppt_read", "get_presentation_view", {"file_path": deck})
    assert is_error is True
    assert payload["error"]["code"] == "DOCUMENT_LOCKED", payload
    assert "being written by another program" in payload["error"]["message"]
    assert "PowerPoint" not in payload["error"]["message"]
    assert "Retry" in payload["error"]["hint"]


def test_a_live_auto_read_does_not_go_looking_for_powerpoint(make_deck, monkeypatch):
    """live='auto' falls through to the open copy only on DocumentLocked;
    a save race is not that, so the live layer is never called."""
    deck = str(make_deck())
    _flaky_zip(monkeypatch, deck, failures=None)
    called = []
    monkeypatch.setattr(
        server._lo, "live_get_text",
        lambda *a, **k: called.append(a) or {"slides": []},
    )
    payload, is_error = _call(
        "ppt_read", "get_text", {"file_path": deck, "live": "auto"}
    )
    assert is_error is True
    assert payload["error"]["code"] == "DOCUMENT_LOCKED"
    assert called == []


# ---------------------------------------------------- the lock probe is retried too


def _flaky_probe(monkeypatch, failures: int | None, mode: str = "r+b"):
    """The open that PptxPackage._check_lock uses to probe the deck meets the
    save race's PermissionError, for the first ``failures`` probes (None:
    every probe). Returns the probe counter."""
    import builtins

    seen = {"n": 0}

    def fake_open(file, open_mode="r", *args, **kwargs):
        if open_mode == mode:
            seen["n"] += 1
            if failures is None or seen["n"] <= failures:
                raise PermissionError(13, "Permission denied", str(file))
        return builtins.open(file, open_mode, *args, **kwargs)

    monkeypatch.setattr(package, "open", fake_open, raising=False)
    return seen


def test_a_lock_probe_that_meets_a_save_looks_again_and_succeeds(make_deck, monkeypatch):
    """The probe just before the zip open can land in the same replace. It
    used to refuse at once with "open in PowerPoint or locked by another
    process", though no PowerPoint was involved."""
    deck = str(make_deck())
    seen = _flaky_probe(monkeypatch, failures=2)
    payload, is_error = _call("ppt_read", "get_text", {"file_path": deck, "live": "off"})
    assert is_error is False, payload
    assert seen["n"] == 3


def test_a_lock_probe_that_keeps_meeting_a_save_is_document_locked_and_says_so(
        make_deck, monkeypatch):
    deck = str(make_deck())
    seen = _flaky_probe(monkeypatch, failures=None)
    payload, is_error = _call("ppt_read", "get_text", {"file_path": deck, "live": "off"})
    assert is_error is True
    assert payload["error"]["code"] == "DOCUMENT_LOCKED", payload
    assert "being written by another program" in payload["error"]["message"]
    assert "PowerPoint" not in payload["error"]["message"]
    assert "Retry" in payload["error"]["hint"]
    assert seen["n"] == 1 + len(package._READ_RACE_DELAYS)


def test_a_read_only_decks_probe_looks_again_too(make_deck, monkeypatch):
    import os
    import stat

    deck = make_deck()
    os.chmod(deck, stat.S_IREAD)
    try:
        seen = _flaky_probe(monkeypatch, failures=2, mode="rb")
        payload, is_error = _call("ppt_read", "get_text", {"file_path": str(deck), "live": "off"})
        assert is_error is False, payload
        assert seen["n"] >= 3
    finally:
        os.chmod(deck, stat.S_IWRITE | stat.S_IREAD)


def test_a_live_auto_read_in_a_save_race_does_not_answer_app_not_running(make_deck, monkeypatch):
    """With live='auto' a lock refusal sends the read looking for the deck in
    PowerPoint, which answers APP_NOT_RUNNING when it is not running: the
    opposite of what looking again is for."""
    deck = str(make_deck())
    seen = _flaky_probe(monkeypatch, failures=2)
    payload, is_error = _call("ppt_read", "get_text", {"file_path": deck})
    assert is_error is False, payload
    assert seen["n"] == 3

    called = []
    monkeypatch.setattr(
        server._lo, "live_get_text",
        lambda *a, **k: called.append(a) or {"slides": []},
    )
    _flaky_probe(monkeypatch, failures=None)
    payload, is_error = _call("ppt_read", "get_text", {"file_path": deck, "live": "auto"})
    assert is_error is True
    assert payload["error"]["code"] == "DOCUMENT_LOCKED"
    assert called == []


def test_a_probe_refusal_with_powerpoints_owner_file_is_final(make_deck, monkeypatch):
    """PowerPoint writes ~$name beside a deck it has open. With that file
    there the refused probe is PowerPoint's hold, known at once (no read-race
    pauses): a live='off' read reads the saved file and says so (it used to
    refuse DOCUMENT_LOCKED), and a live='auto' read goes to the open copy."""
    deck = make_deck()
    owner = deck.with_name("~$" + deck.name)
    owner.write_bytes(b"")
    seen = _flaky_probe(monkeypatch, failures=None)
    payload, is_error = _call("ppt_read", "get_text", {"file_path": str(deck), "live": "off"})
    assert is_error is False, payload
    assert "open in PowerPoint" in payload["open_deck_note"]
    assert seen["n"] == 1

    called = []
    monkeypatch.setattr(
        server._lo, "live_get_text",
        lambda *a, **k: called.append(a) or {"slides": []},
    )
    payload, is_error = _call("ppt_read", "get_text", {"file_path": str(deck), "live": "auto"})
    assert is_error is False, payload
    assert len(called) == 1


# ------------------------------- a save has the old deck aside for an instant


def _aside(deck, how_long: float | None):
    """The state ReplaceFileW leaves for an instant: the deck moved to its
    swap name, the path missing. After ``how_long`` seconds it goes back
    (None: never)."""
    import threading

    from kitchensink4ppt.core import safesave

    marker = deck.with_name(safesave._swap_prefix(deck.name) + "feedface.tmp")
    deck.rename(marker)

    def put_back():
        marker.rename(deck)

    timer = None
    if how_long is not None:
        timer = threading.Timer(how_long, put_back)
        timer.start()
    return marker, timer


def test_a_deck_a_save_has_aside_is_waited_for(make_deck):
    deck = make_deck()
    marker, timer = _aside(deck, 0.12)
    try:
        pkg = package.PptxPackage(deck)
    finally:
        timer.join()
    assert pkg.slide_parts()
    assert deck.exists() and not marker.exists()


def test_a_deck_a_save_never_puts_back_is_not_found_after_the_wait(make_deck):
    from kitchensink4ppt.core.errors import DocumentNotFound

    deck = make_deck()
    marker, _ = _aside(deck, None)
    with pytest.raises(DocumentNotFound, match="No file at"):
        package.PptxPackage(deck)
    assert marker.exists()


def test_a_deck_that_is_not_there_is_reported_at_once(tmp_path):
    import time

    from kitchensink4ppt.core.errors import DocumentNotFound

    started = time.monotonic()
    with pytest.raises(DocumentNotFound, match="No file at"):
        package.PptxPackage(tmp_path / "nothing.pptx")
    assert time.monotonic() - started < 0.2, "a deck that is not there must not be waited for"


def test_a_deck_that_goes_missing_mid_read_and_comes_back_is_read_again(make_deck, monkeypatch):
    import builtins

    deck = make_deck()
    seen = {"n": 0}

    def vanishes_once(file, mode="r", *args, **kwargs):
        if mode == "r+b":
            seen["n"] += 1
            if seen["n"] == 1:
                raise FileNotFoundError(2, "gone for an instant", str(file))
        return builtins.open(file, mode, *args, **kwargs)

    monkeypatch.setattr(package, "open", vanishes_once, raising=False)
    pkg = package.PptxPackage(deck)
    assert pkg.slide_parts() and seen["n"] == 2


def test_a_deck_that_goes_missing_mid_read_and_stays_missing_is_not_found(make_deck, monkeypatch):
    from kitchensink4ppt.core.errors import DocumentNotFound

    deck = make_deck()

    def gone(file, mode="r", *args, **kwargs):
        if mode == "r+b":
            deck.unlink()
            raise FileNotFoundError(2, "gone", str(file))
        raise AssertionError("not reached")

    monkeypatch.setattr(package, "open", gone, raising=False)
    with pytest.raises(DocumentNotFound, match="No file at"):
        package.PptxPackage(deck)


# --------------------- every route that reads a deck waits out the same swap


def _route(name, deck, tmp_path, make_deck):
    """(tool, op, args) for a route that reads ``deck``; any other deck it
    needs is made here, before the deck is put aside."""
    if name == "copy_presentation":
        return "ppt_deck", "copy_presentation", {"file_path": str(deck),
                                                 "dest_path": str(tmp_path / "copy.pptx")}
    if name == "create_snapshot":
        return "ppt_deck", "create_snapshot", {"file_path": str(deck), "dest_dir": str(tmp_path)}
    if name == "split_deck":
        out = tmp_path / "split"
        out.mkdir()
        return "ppt_deck", "split_deck", {"file_path": str(deck), "output_dir": str(out),
                                          "by": "ranges", "ranges": [{"start": 0, "end": 0}]}
    if name == "merge_decks":
        dest = make_deck("merge_dest.pptx", seed=3)
        return "ppt_deck", "merge_decks", {"file_path": str(dest), "sources": [str(deck)]}
    if name == "copy_slide_between":
        dest = make_deck("csb_dest.pptx", seed=4)
        return "ppt_deck", "copy_slide_between", {"file_path": str(dest),
                                                  "source_path": str(deck), "slide": 0}
    if name == "validate":
        return "ppt_review", "validate", {"file_path": str(deck)}
    if name == "create_presentation":
        return "ppt_deck", "create_presentation", {"path": str(tmp_path / "from_template.pptx"),
                                                   "template": str(deck)}
    if name == "restore_from_a_file":
        target = make_deck("restore_target.pptx", seed=5)
        return "ppt_deck", "manage_backups", {"action": "restore", "file_path": str(target),
                                              "source": str(deck)}
    if name == "diagnose":
        return "ppt_inspect", "diagnose", {"file_path": str(deck)}
    raise AssertionError(name)


ROUTES = ["copy_presentation", "create_snapshot", "split_deck", "merge_decks",
          "copy_slide_between", "validate", "create_presentation",
          "restore_from_a_file", "diagnose"]


@pytest.mark.parametrize("name", ROUTES)
def test_a_route_that_reads_a_deck_waits_out_another_saves_swap(name, make_deck, tmp_path, monkeypatch):
    """ReplaceFileW leaves the deck's path empty for a few milliseconds
    (once 182 ms, measured from another process). Each of these routes
    checked the path first and answered NOT_FOUND at once; they now wait,
    as the package reader does."""
    from kitchensink4ppt.com import bridge

    monkeypatch.setattr(bridge, "com_validate_opens_clean",
                        lambda fp, **_kw: {"opens_clean": True})
    deck = make_deck("source.pptx")
    tool, op, args = _route(name, deck, tmp_path, make_deck)
    tool = public_ops()[op].group
    _marker, timer = _aside(deck, 0.15)
    try:
        payload, is_error = _call(tool, op, args)
    finally:
        timer.join()
    assert is_error is False, payload
    if name == "diagnose":
        assert "file does not exist" not in str(payload), payload


@pytest.mark.parametrize("module", ["ops.export", "com.bridge"])
def test_the_export_check_waits_out_another_saves_swap(module, make_deck):
    """export_pdf, export_handout, export_slide_image (both engines) and the
    opens-clean check all start with _require_file."""
    import importlib
    from pathlib import Path

    check = importlib.import_module(f"kitchensink4ppt.{module}")._require_file
    deck = make_deck("source.pptx")
    _marker, timer = _aside(deck, 0.15)
    try:
        got = check(str(deck), "export source")
    finally:
        timer.join()
    assert Path(got) == deck


def test_a_route_still_says_at_once_that_a_deck_is_not_there(tmp_path):
    import time

    started = time.monotonic()
    payload, is_error = _call("ppt_deck", "create_snapshot",
                              {"file_path": str(tmp_path / "nothing.pptx")})
    assert is_error and payload["error"]["code"] == "NOT_FOUND", payload
    assert time.monotonic() - started < 0.5


# ------------------- a write refused by another program stays a file refusal


@pytest.mark.skipif(sys.platform != "win32", reason="Windows sharing modes")
def test_a_live_auto_write_refused_by_another_program_does_not_go_looking_for_powerpoint(
        make_deck, monkeypatch):
    """A reader that does not share delete (an antivirus, indexer, sync
    client or viewer) refuses the save's swap. live='auto' used to read
    that DocumentLocked as "open in PowerPoint" and send the write to the
    live layer: APP_NOT_RUNNING with no PowerPoint involved, or an attach to
    the user's running PowerPoint. The owner-file check had already
    excluded PowerPoint."""
    deck = make_deck()
    called = []
    monkeypatch.setattr(server._lo, "live_set_notes", lambda *a, **k: called.append(a) or {})
    with open(deck, "rb"):  # shares read and write, not delete
        payload, is_error = _call("ppt_edit", "set_notes", {"file_path": str(deck), "slide": 0,
                                                            "text": "x", "live": "auto"})
    assert called == []
    assert is_error is True and payload["error"]["code"] == "DOCUMENT_LOCKED", payload
    assert "PowerPoint" not in payload["error"]["hint"]


@pytest.mark.skipif(sys.platform != "win32", reason="Windows sharing modes")
def test_a_live_auto_write_on_a_deck_powerpoint_has_open_still_goes_to_powerpoint(
        make_deck, monkeypatch):
    """PowerPoint's owner file beside a deck that is held: that hold is
    PowerPoint's, and the edit goes to the open copy as before."""
    import ctypes
    from ctypes import wintypes

    deck = make_deck()
    deck.with_name("~$" + deck.name).write_bytes(bytes(165))
    called = []
    monkeypatch.setattr(server._lo, "live_set_notes",
                        lambda *a, **k: called.append(a) or {"slide": 0})
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateFileW.restype = wintypes.HANDLE
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
                                wintypes.HANDLE]
    held = k32.CreateFileW(str(deck), 0x80000000, 0, None, 3, 0x80, None)  # no sharing
    assert held and held != wintypes.HANDLE(-1).value
    try:
        payload, is_error = _call("ppt_edit", "set_notes", {"file_path": str(deck), "slide": 0,
                                                            "text": "x", "live": "auto"})
    finally:
        k32.CloseHandle(held)
    assert len(called) == 1, payload
