"""validate gives PowerPoint the time its deck needs, and a timeout reads as
a timeout.

Field report: validate's opens-clean check stopped at 60 s on a healthy
47-slide deck of about 1,400 shapes with speaker notes on every slide, and
the answer read "opens-clean check unavailable here (PowerPoint did not
answer within 60 s ... the output file may be incomplete or missing)": an
alarming line for a read-only check, about a file that opened fine.

Now the time limit grows with the deck (its slides, shapes, notes pages and
media) up to a cap, and a timeout is answered as one: opens_clean null,
timed_out true, and a note that says it is a time limit and not a verdict
on the deck.
"""

from __future__ import annotations

import sys

import pytest

from kitchensink4ppt import server
from kitchensink4ppt.com import bridge
from kitchensink4ppt.core.errors import PowerPointBlocked
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import notes as nt
from kitchensink4ppt.ops import slides as sl
from kitchensink4ppt.ops import text as tx


# ------------------------------------------------------------ the limit


def test_the_time_limit_grows_with_the_deck_and_stops_at_a_cap():
    small, small_capped = bridge.opens_clean_time_limit(slides=5, shapes=60)
    field, field_capped = bridge.opens_clean_time_limit(
        slides=47, shapes=1400, notes=47)
    huge, huge_capped = bridge.opens_clean_time_limit(
        slides=900, shapes=40000, notes=900, media_bytes=2_000_000_000)
    assert 60.0 <= small < field < huge
    assert not small_capped and not field_capped
    assert huge == bridge.OPENS_CLEAN_CAP_S and huge_capped
    # The field deck stopped at 60 s and again at 72 s; its limit is well
    # past both.
    assert field >= 120.0


def test_notes_and_media_count_too():
    base, _ = bridge.opens_clean_time_limit(slides=47, shapes=1400)
    with_notes, _ = bridge.opens_clean_time_limit(slides=47, shapes=1400, notes=47)
    with_media, _ = bridge.opens_clean_time_limit(
        slides=47, shapes=1400, media_bytes=200_000_000)
    assert with_notes > base and with_media > base


def test_the_call_budget_leaves_room_for_the_longest_limit():
    from kitchensink4ppt.registry import public_ops

    budget = public_ops()["validate"].timeout_s
    assert budget >= bridge.OPENS_CLEAN_CAP_S + 60.0


def _deck(make_deck, slides=12, boxes=8, notes=True):
    path = make_deck("big.pptx", extra_slides=0)
    pkg = PptxPackage(path)
    while len(pkg.slide_parts()) < slides:
        sl.insert_slide(pkg, "Blank")
    for s in range(slides):
        for k in range(boxes):
            tx.insert_textbox(pkg, s, f"box {k}", 0.2 + k, 0.2, 0.9, 0.5)
        if notes:
            nt.set_notes(pkg, s, f"Speaker notes for slide {s + 1}.")
    pkg.save()
    return path


def test_validate_gives_powerpoint_the_limit_for_its_deck(make_deck, monkeypatch):
    seen = {}

    def fake(path, timeout=None):
        seen["timeout"] = timeout
        return {"opens_clean": True}

    monkeypatch.setattr(bridge, "com_validate_opens_clean", fake)
    path = _deck(make_deck)
    out = server.validate(str(path))
    assert out["powerpoint"]["opens_clean"] is True
    counts = out["opens_clean_time_limit"]
    expected, _capped = bridge.opens_clean_time_limit(
        counts["slides"], counts["shapes"], counts["notes"], counts["media_bytes"])
    assert seen["timeout"] == expected
    assert counts["slides"] == 12 and counts["notes"] == 12
    assert counts["shapes"] >= 12 * 8


# ------------------------------------------------------- a timeout reads as one


def _timed_out(limit, ended=True):
    exc = PowerPointBlocked(
        f"PowerPoint did not answer within {limit:.0f} s. This server ended the "
        "hidden PowerPoint it had started for this call (pid 4242), so the "
        "output file may be incomplete or missing; check it before you retry.")
    exc.timed_out = True
    exc.time_limit_s = limit
    exc.ended = [4242] if ended else []
    exc.left = [] if ended else [{"pid": 4242, "why": "not proven ours"}]
    return exc


def test_a_timeout_is_answered_as_a_timeout_not_as_a_broken_deck(make_deck, monkeypatch):
    def slow(path, timeout=None):
        raise _timed_out(timeout)

    monkeypatch.setattr(bridge, "com_validate_opens_clean", slow)
    path = _deck(make_deck, slides=6, boxes=3)
    out = server.validate(str(path))
    pp = out["powerpoint"]
    assert pp["opens_clean"] is None
    assert pp["timed_out"] is True
    assert pp["time_limit_s"] == out["opens_clean_time_limit"]["seconds"]
    note = out["powerpoint_note"]
    assert "time limit" in note and "not a verdict" in note
    assert "output file" not in note and "incomplete" not in note
    assert out["payload_valid"] is True
    assert out["ok"] is True  # the payload passed; PowerPoint gave no verdict


def test_a_powerpoint_left_running_after_a_timeout_is_named(make_deck, monkeypatch):
    def slow(path, timeout=None):
        raise _timed_out(timeout, ended=False)

    monkeypatch.setattr(bridge, "com_validate_opens_clean", slow)
    path = _deck(make_deck, slides=4, boxes=2, notes=False)
    out = server.validate(str(path))
    assert "powerpoint_status" in out["powerpoint_note"]


def test_the_cap_is_named_when_it_is_the_limit(make_deck, monkeypatch):
    monkeypatch.setattr(bridge, "opens_clean_time_limit",
                        lambda *a, **k: (bridge.OPENS_CLEAN_CAP_S, True))

    def slow(path, timeout=None):
        raise _timed_out(timeout)

    monkeypatch.setattr(bridge, "com_validate_opens_clean", slow)
    path = _deck(make_deck, slides=4, boxes=2, notes=False)
    out = server.validate(str(path))
    assert "the most this check waits" in out["powerpoint_note"]
    assert out["opens_clean_time_limit"]["capped"] is True


# ------------------------------------- what the note says the limit was set from


def _note(**counts):
    full = {"slides": 0, "shapes": 0, "notes": 0, "media_bytes": 0, **counts}
    return server._timed_out_note(5.0, False, full, True, False)


def test_the_timeout_note_names_only_what_the_deck_has():
    """In real PowerPoint the note read "the time this check gives a deck of
    30 slides, 750 shapes and 0 notes pages": a count of nothing."""
    note = _note(slides=30, shapes=750)
    assert "the time this check gives a deck with 30 slides and 750 shapes." in note
    assert "0 notes" not in note
    assert "within 5 s" in note
    assert "not a verdict on the deck: the package check passed" in note


def test_the_timeout_note_counts_as_a_reader_says_them():
    assert "a deck with 1 slide, 1 shape and 1 notes page." in _note(
        slides=1, shapes=1, notes=1)
    assert "a deck with 47 slides, 1,400 shapes and 47 notes pages." in _note(
        slides=47, shapes=1400, notes=47)
    assert "a deck with no slides." in _note()


def test_the_timeout_note_names_the_media_that_grew_the_limit():
    """Media grows the limit too (OPENS_CLEAN_PER_MEDIA_MB_S), and the note
    never said so."""
    assert "a deck with 3 slides, 5 shapes and 123.5 MB of media." in _note(
        slides=3, shapes=5, media_bytes=123_456_789)


def test_validate_names_the_deck_it_timed_out_on(make_deck, monkeypatch):
    def slow(path, timeout=None):
        raise _timed_out(timeout)

    monkeypatch.setattr(bridge, "com_validate_opens_clean", slow)
    path = _deck(make_deck, slides=4, boxes=2, notes=False)
    out = server.validate(str(path))
    counts = out["opens_clean_time_limit"]
    note = out["powerpoint_note"]
    # The base deck brings notes pages of its own; the note names the deck's
    # own counts, and none of them is a zero.
    assert counts["slides"] == 4 and counts["notes"] >= 1
    assert f"a deck with 4 slides, {counts['shapes']:,} shapes" in note
    assert f"{counts['notes']} notes page" in note
    assert " 0 " not in note


def test_the_bridge_marks_its_own_timeouts(monkeypatch):
    """The bounded runner says it timed out, and after how long, on the
    exception itself, so a caller never has to read it from the words."""
    import threading

    monkeypatch.setattr(bridge, "TIMEOUT_GRACE_SECONDS", 0.05)
    release = threading.Event()
    try:
        with pytest.raises(PowerPointBlocked) as caught:
            bridge._run_bounded("probe_timeout", 0.2, lambda: release.wait(5))
    finally:
        release.set()
    assert getattr(caught.value, "timed_out", False) is True
    assert caught.value.time_limit_s == 0.2


# ------------------------------------------------------------- drawn for real


@pytest.mark.live
@pytest.mark.timeout(1200)
def test_a_deck_like_the_field_deck_opens_clean_in_time(make_deck):
    """About 50 slides, 30 shapes each and notes on every slide, through
    validate itself, in a hidden PowerPoint of its own."""
    if sys.platform != "win32":
        pytest.skip("PowerPoint is Windows-only")
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        pytest.skip("pywin32 not installed")
    if not bridge.powerpoint_installed():
        pytest.skip("PowerPoint is not installed on this machine")
    import time

    deadline = time.monotonic() + 30  # an instance that is quitting
    running = bridge.powerpnt_pids()
    while (running is None or running) and time.monotonic() < deadline:
        time.sleep(1.0)
        running = bridge.powerpnt_pids()
    if running is None or running:
        pytest.skip("SKIPPED-POWERPOINT-BUSY: another PowerPoint is running")
    path = _deck(make_deck, slides=50, boxes=30, notes=True)
    out = server.validate(str(path))
    assert out["powerpoint"] and out["powerpoint"].get("opens_clean") is True, out
