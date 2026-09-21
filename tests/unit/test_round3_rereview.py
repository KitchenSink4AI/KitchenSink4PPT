"""Round 3: the second review's five release gates on the 1.3.1 candidate.

G1  the measured text path ignored a:rPr/@spc and resolved inherited run
    size, level and list-style margins incompletely, so an overflow could
    come back as `font-metrics` + `likely_overflow: false`;
G2  the failed-DispatchEx cleanup inferred ownership from an empty process
    snapshot and force-ended every POWERPNT pid that appeared, so a
    PowerPoint launched concurrently by the user could be killed;
G3  the text carry compared the wrong hover element (a:hlinkHover instead
    of a:hlinkMouseOver) and let one paragraph's disagreement strip links
    that were unanimous inside other paragraphs;
G4  a text= replacement could not state outline level 0, so an unindented
    line kept the old paragraph's nesting;
G5  the first Slides.Count access in the layer 2 walk sat outside the
    environment classification, so a busy PowerPoint became opens_clean
    false.

The probes here are the reviewer's, adapted. No PowerPoint is started by
anything in this file.
"""

from __future__ import annotations

import inspect
import sys

import pytest
from lxml import etree

from kitchensink4ppt.core.package import NSMAP, qn
from kitchensink4ppt.ops import fontmetrics as _fm
from kitchensink4ppt.ops import text as tx


# ------------------------------------------------------------------ G1


def _metrics_shape(runs, *, spacing=None, lst_size=None, wrap="none",
                   cx=1828800, cy=548640):
    """A p:sp whose txBody holds one paragraph of `runs`.

    Each run is (text, typeface, sz-in-hundredths or None). `spacing` sets
    a:rPr/@spc on every run; `lst_size` puts a size on the shape's own
    lvl1pPr so the run inherits it instead of stating one.
    """
    sp = etree.Element(qn("p:sp"))
    sppr = etree.SubElement(sp, qn("p:spPr"))
    xfrm = etree.SubElement(sppr, qn("a:xfrm"))
    etree.SubElement(xfrm, qn("a:off"), x="0", y="0")
    etree.SubElement(xfrm, qn("a:ext"), cx=str(cx), cy=str(cy))
    body = etree.SubElement(sp, qn("p:txBody"))
    etree.SubElement(body, qn("a:bodyPr"), wrap=wrap)
    lst = etree.SubElement(body, qn("a:lstStyle"))
    if lst_size is not None:
        lvl = etree.SubElement(lst, qn("a:lvl1pPr"))
        default = etree.SubElement(lvl, qn("a:defRPr"), sz=str(lst_size))
        etree.SubElement(default, qn("a:latin"), typeface="Calibri")
    p = etree.SubElement(body, qn("a:p"))
    etree.SubElement(p, qn("a:pPr"), lvl="0")
    for text, family, size in runs:
        r = etree.SubElement(p, qn("a:r"))
        attrs = {"sz": str(size)} if size is not None else {}
        if spacing is not None:
            attrs["spc"] = str(spacing)
        rpr = etree.SubElement(r, qn("a:rPr"), **attrs)
        etree.SubElement(rpr, qn("a:latin"), typeface=family)
        etree.SubElement(r, qn("a:t")).text = text
    return sp, body, p


def _overflow(sp, body):
    return tx._overflow_heuristic(
        sp, body, body.find(qn("a:bodyPr")), 100.0, 0.0
    )


metrics = pytest.mark.skipif(
    not _fm.available() or _fm.find_font_face("Calibri") is None,
    reason="needs the metrics extra and Calibri installed",
)


@metrics
def test_character_spacing_widens_the_measured_line():
    """G1a: PowerPoint adds a:rPr/@spc between characters. Measuring the
    same string with and without it used to give the identical ratio."""
    plain_sp, plain_body, _ = _metrics_shape([("EXPANDED", "Calibri", 1200)])
    spaced_sp, spaced_body, _ = _metrics_shape(
        [("EXPANDED", "Calibri", 1200)], spacing=1000
    )
    plain = _overflow(plain_sp, plain_body)
    spaced = _overflow(spaced_sp, spaced_body)
    assert plain["method"] == "font-metrics"
    assert spaced["method"] == "font-metrics"
    assert spaced["width_ratio"] > plain["width_ratio"]


@metrics
def test_character_spacing_can_flip_the_overflow_verdict():
    """G1a, the customer-visible half: 10pt of tracking per character on a
    label that fits without it must not come back as a fit."""
    plain_sp, plain_body, _ = _metrics_shape([("EXPANDED", "Calibri", 1200)])
    spaced_sp, spaced_body, _ = _metrics_shape(
        [("EXPANDED", "Calibri", 1200)], spacing=1000
    )
    assert _overflow(plain_sp, plain_body)["likely_overflow"] is False
    spaced = _overflow(spaced_sp, spaced_body)
    assert spaced["likely_overflow"] is True


@metrics
def test_an_inherited_list_style_size_is_measured_at_that_size():
    """G1b: a run with no sz of its own, under a 28pt lvl1pPr, was measured
    at the body default (12pt), which can reverse an overflow verdict."""
    _sp, body, p = _metrics_shape(
        [("Inherited", "Calibri", None)], lst_size=2800
    )
    runs = tx._paragraph_runs(p, body, "Calibri", 12.0)
    assert runs is not None
    assert runs[0].size_pt == 28.0


@metrics
def test_a_size_that_does_not_resolve_forces_an_honest_estimate():
    """G1b: no sz on the run, nothing in the chain to resolve it, and no
    placeholder to inherit from. The measured path must refuse rather than
    measure the body default."""
    sp, body, _p = _metrics_shape([("Unsized", "Calibri", None)])
    rec = _overflow(sp, body)
    assert rec["method"] == "estimate"
    assert "size" in rec["method_reason"]


@metrics
def test_the_run_that_forced_the_fallback_is_the_one_reported():
    """G1/N3: a mixed body used to blame the FIRST run's face for a later
    run's missing font."""
    sp, body, _p = _metrics_shape(
        [("fine ", "Calibri", 1200),
         ("missing", "NoSuchFontFamilyAnywhere", 1200)]
    )
    rec = _overflow(sp, body)
    assert rec["method"] == "estimate"
    assert "NoSuchFontFamilyAnywhere" in rec["method_reason"]


@metrics
def test_an_inherited_list_style_indent_narrows_the_usable_width():
    """G1b: marL/indent that come from the shape's own list style rather
    than from the paragraph used to be read as zero, so the measurement
    thought it had the whole frame to lay text into."""
    sp, body, _p = _metrics_shape(
        [("Indented text here", "Calibri", 1200)], wrap="square"
    )
    flat = _overflow(sp, body)
    lvl = body.find(qn("a:lstStyle")).makeelement(qn("a:lvl1pPr"), {})
    lvl.set("marL", "914400")  # one inch of the frame is gone
    body.find(qn("a:lstStyle")).append(lvl)
    indented = _overflow(sp, body)
    assert flat["method"] == indented["method"] == "font-metrics"
    assert indented["height_ratio"] > flat["height_ratio"]


# ------------------------------------------------------------------ G2

pytestmark_win = pytest.mark.skipif(
    sys.platform != "win32", reason="the COM bridge is Windows only"
)


class _Pythoncom:
    COINIT_APARTMENTTHREADED = 2

    def CoInitializeEx(self, _flags):  # noqa: N802 - COM spelling
        return None


@pytestmark_win
def test_a_powerpoint_that_appears_during_a_failed_start_is_not_killed(
    monkeypatch,
):
    """G2: the entry snapshot was empty, so every pid that appeared after
    it was force-ended. A user launching PowerPoint in that window lost
    unsaved work."""
    from kitchensink4ppt.com import bridge

    world = {"pids": set()}
    killed: list[int] = []

    class Win32:
        def DispatchEx(self, _progid):  # noqa: N802 - COM spelling
            # The user launched PowerPoint after the empty entry snapshot
            # and before this call reported its fault.
            world["pids"].add(4242)
            raise RuntimeError("dispatch failed")

    def fake_run(cmd, **_kwargs):
        killed.append(int(cmd[2]))

        class Result:
            returncode = 0

        return Result()

    monkeypatch.setattr(bridge, "powerpnt_pids", lambda: set(world["pids"]))
    monkeypatch.setattr(bridge.subprocess, "run", fake_run)
    with pytest.raises(Exception):
        bridge._start_powerpoint(Win32(), _Pythoncom(), set())
    assert killed == [], "no process may be ended on a failed start"
    assert world["pids"] == {4242}, "the user's PowerPoint must survive"


@pytestmark_win
def test_the_failed_start_reports_the_process_it_left_running(monkeypatch):
    """G2: the refusal used to claim a process 'was ended'. It may now
    only say what it saw."""
    from kitchensink4ppt.com import bridge

    world = {"pids": set()}

    class Win32:
        def DispatchEx(self, _progid):  # noqa: N802 - COM spelling
            world["pids"].add(7777)
            raise _com_error(bridge.CO_E_SERVER_EXEC_FAILURE)

    monkeypatch.setattr(bridge, "powerpnt_pids", lambda: set(world["pids"]))
    monkeypatch.setattr(
        bridge.subprocess, "run", lambda *a, **k: pytest.fail("no subprocess")
    )
    with pytest.raises(Exception) as exc_info:
        bridge._start_powerpoint(Win32(), _Pythoncom(), set())
    message = str(exc_info.value)
    assert "was ended" not in message
    assert "7777" in message
    assert "left running" in message


@pytestmark_win
def test_nothing_in_the_failed_start_path_can_end_a_process():
    """G2: a source guard, because the danger is a call that is easy to
    reintroduce and expensive to notice."""
    from kitchensink4ppt.com import bridge

    for fn in (bridge._start_powerpoint, bridge._raise_startup):
        source = inspect.getsource(fn)
        for forbidden in ("taskkill", "Terminate", "TerminateProcess"):
            assert forbidden not in source, (
                f"{fn.__name__} must not be able to end a process"
            )


@pytestmark_win
def test_the_retry_stops_once_a_powerpoint_has_appeared(monkeypatch):
    """G2: retrying DispatchEx after something appeared would launch a
    second process on top of a state we cannot explain."""
    from kitchensink4ppt.com import bridge

    world = {"pids": set()}
    calls = {"n": 0}

    class Win32:
        def DispatchEx(self, _progid):  # noqa: N802 - COM spelling
            calls["n"] += 1
            world["pids"].add(5150)
            raise _com_error(bridge.CO_E_NOTINITIALIZED)

    monkeypatch.setattr(bridge, "powerpnt_pids", lambda: set(world["pids"]))
    with pytest.raises(Exception):
        bridge._start_powerpoint(Win32(), _Pythoncom(), set())
    assert calls["n"] == 1, "a process appeared, so the retry must stop"


def _com_error(hresult: int):
    import pywintypes

    return pywintypes.com_error(hresult, "fabricated for the test", None, None)


# ------------------------------------------------------------------ G3


def _body_with_paragraphs(specs):
    """specs: list of lists of rId-or-None, one list per paragraph."""
    body = etree.Element(qn("p:txBody"))
    etree.SubElement(body, qn("a:bodyPr"))
    for runs in specs:
        p = etree.SubElement(body, qn("a:p"))
        for i, rid in enumerate(runs):
            r = etree.SubElement(p, qn("a:r"))
            rpr = etree.SubElement(r, qn("a:rPr"))
            rpr.set("sz", "1800")
            if rid is not None:
                tag, value = rid
                etree.SubElement(rpr, qn(tag)).set(
                    f"{{{NSMAP['r']}}}id", value
                )
            etree.SubElement(r, qn("a:t")).text = f"run{i}"
    return body


def _carry(old_body, paragraph_count):
    from kitchensink4ppt.ops import shapes as sh

    new = etree.Element(qn("p:txBody"))
    etree.SubElement(new, qn("a:bodyPr"))
    for i in range(paragraph_count):
        p = etree.SubElement(new, qn("a:p"))
        etree.SubElement(p, qn("a:pPr"))
        r = etree.SubElement(p, qn("a:r"))
        etree.SubElement(r, qn("a:rPr"))
        etree.SubElement(r, qn("a:t")).text = f"new{i}"
    carried, facts = sh._carry_text_properties(old_body, new, set())
    return new, carried, facts


def test_a_paragraph_keeps_the_link_all_of_its_runs_share():
    """G3: paragraph 1 mixes linked and unlinked runs; paragraph 2 shares
    one link. The body-global flag stripped paragraph 2's link too."""
    old = _body_with_paragraphs([
        [("a:hlinkClick", "rId1"), None],
        [("a:hlinkClick", "rId2"), ("a:hlinkClick", "rId2")],
    ])
    new, _carried, facts = _carry(old, 2)
    paras = new.findall(qn("a:p"))
    first = paras[0].find(f"{qn('a:r')}/{qn('a:rPr')}")
    second = paras[1].find(f"{qn('a:r')}/{qn('a:rPr')}")
    assert first.find(qn("a:hlinkClick")) is None, (
        "the disagreeing paragraph must not be relinked"
    )
    assert second.find(qn("a:hlinkClick")) is not None, (
        "a link every run of THIS paragraph shared must survive"
    )
    assert "hyperlinks_dropped" in facts
    assert "0" in facts["hyperlinks_dropped"]


def test_a_disagreeing_hover_link_is_seen():
    """G3: the code compared a:hlinkHover, which is not the run-level
    element, so hover links compared equal whatever they pointed at."""
    old = _body_with_paragraphs([
        [("a:hlinkMouseOver", "rId1"), None],
    ])
    new, _carried, facts = _carry(old, 1)
    rpr = new.find(f"{qn('a:p')}/{qn('a:r')}/{qn('a:rPr')}")
    assert rpr.find(qn("a:hlinkMouseOver")) is None, (
        "the first run's hover link was spread over the replacement"
    )
    assert "hyperlinks_dropped" in facts


def test_the_run_level_hover_element_is_the_schema_one():
    from kitchensink4ppt.ops import shapes as sh

    assert "a:hlinkMouseOver" in sh._HLINK_TAGS
    assert "a:hlinkHover" not in sh._HLINK_TAGS


# ------------------------------------------------------------------ G4


def test_an_unindented_line_states_level_zero():
    """G4: a line with no leading tab said nothing about its level, so a
    flat replacement kept the nesting of the paragraph it replaced."""
    specs = tx._parse_text_paragraphs("Top level\n\tSecond level\nTop again")
    assert [s["level"] for s in specs] == [0, 1, 0]
    assert all(s["level_explicit"] for s in specs)


def test_a_flat_replacement_unnests_a_nested_paragraph():
    """G4 end to end through the carry: the old paragraph is at level 1 and
    the replacement line has no tab, so the result is level 0."""
    from kitchensink4ppt.ops import shapes as sh

    old = etree.Element(qn("p:txBody"))
    etree.SubElement(old, qn("a:bodyPr"))
    p = etree.SubElement(old, qn("a:p"))
    etree.SubElement(p, qn("a:pPr")).set("lvl", "1")
    r = etree.SubElement(p, qn("a:r"))
    etree.SubElement(r, qn("a:rPr")).set("sz", "1800")
    etree.SubElement(r, qn("a:t")).text = "nested"

    new = etree.Element(qn("p:txBody"))
    etree.SubElement(new, qn("a:bodyPr"))
    for spec in tx._parse_text_paragraphs("flat"):
        new.append(tx._build_paragraph(spec))
    tx._landing_spots(new)
    sh._carry_text_properties(old, new, set())
    ppr = new.find(f"{qn('a:p')}/{qn('a:pPr')}")
    assert ppr is not None and ppr.get("lvl") == "0", (
        "an unindented replacement line must not stay nested"
    )


# ------------------------------------------------------------------ G5


@pytestmark_win
def test_a_busy_powerpoint_at_the_first_slide_count_is_not_a_file_verdict():
    """G5: pres.Slides.Count sat outside the classification, so a busy
    application became opens_clean false on an innocent deck."""
    from kitchensink4ppt.com import bridge
    from kitchensink4ppt.core.errors import PowerPointBusy

    class _Slides:
        @property
        def Count(self):  # noqa: N802 - COM spelling
            raise _com_error(bridge.RPC_E_CALL_REJECTED)

    class Pres:
        Slides = _Slides()

    with pytest.raises(PowerPointBusy):
        bridge._full_load(Pres())


@pytestmark_win
def test_a_failure_at_the_first_slide_count_reports_no_coordinates():
    """G5: no slide was reached, so the four failed_* keys are absent and
    the verdict says only that the load failed."""
    from kitchensink4ppt.com import bridge

    class _Slides:
        @property
        def Count(self):  # noqa: N802 - COM spelling
            raise ValueError("the slide collection could not be read")

    class Pres:
        Slides = _Slides()

    with pytest.raises(bridge.FullLoadFailed) as exc_info:
        bridge._full_load(Pres())
    out = bridge._opens_clean_failure(exc_info.value)
    assert out["opens_clean"] is False
    for key in ("failed_slide_index", "failed_slide_id",
                "failed_shape_index", "failed_shape_name"):
        assert key not in out


def test_validate_does_not_promise_coordinates_it_may_not_have():
    from kitchensink4ppt import server

    doc = " ".join((server.validate.__doc__ or "").split())
    assert "failed_slide_index" in doc
    assert "when it can be located" in doc, (
        "the coordinate sentence must say the coordinates may be absent"
    )


# ----------------------------------------------------------------- G3b
#
# The run-level hover element again, this time on the READING side. Every
# whole-tree sweep and every container scan named a:hlinkClick and
# a:hlinkHover, which is the SHAPE-level pair. A run's a:hlinkMouseOver
# (what PowerPoint writes for Insert > Action > Mouse Over on selected
# text) was invisible to all of them: list_hyperlinks did not report it,
# set/remove left it in place, and a sweep that removed its relationship
# left its r:id pointing at nothing.


def _run_hover_link(pkg, slide, shape_id, rid):
    """Turn the run's existing a:hlinkClick into an a:hlinkMouseOver, which
    is what a mouse-over action on text looks like in a real deck."""
    from kitchensink4ppt.ops import read as _read
    from kitchensink4ppt.ops import shapes as _shapes

    part = _read.slide_table(pkg)[slide]["part"]
    elem, _chain = _shapes._find_shape(pkg, part, shape_id)
    swapped = 0
    for rpr in elem.iter(qn("a:rPr")):
        for el in list(rpr):
            if el.tag == qn("a:hlinkClick"):
                el.tag = qn("a:hlinkMouseOver")
                swapped += 1
    pkg.mark_dirty(part)
    assert swapped == 1
    return part


def test_the_hyperlink_element_list_is_shared_and_complete():
    from kitchensink4ppt.ops._runmap import HLINK_ELEMENTS, HLINK_HOVER_ELEMENTS

    assert set(HLINK_ELEMENTS) == {
        "a:hlinkClick", "a:hlinkHover", "a:hlinkMouseOver"
    }
    assert set(HLINK_HOVER_ELEMENTS) == {"a:hlinkHover", "a:hlinkMouseOver"}


def test_a_run_level_mouse_over_link_is_listed(make_deck):
    from kitchensink4ppt.core.package import PptxPackage
    from kitchensink4ppt.ops import links, text as _text

    pkg = PptxPackage(make_deck("hover_list.pptx"))
    box = _text.insert_textbox(pkg, 0, "Click here for details", 1, 1, 4, 1)
    links.set_hyperlink(
        pkg, 0, {"shape_id": box["shape_id"], "paragraph": 0},
        url="https://example.com/",
    )
    _run_hover_link(pkg, 0, box["shape_id"], "rId2")
    found = [
        rec for rec in links.list_hyperlinks(pkg)["hyperlinks"]
        if rec.get("where") == "text"
    ]
    assert found, "a run's mouse-over link was not listed at all"
    assert found[0]["trigger"] == "hover"
    assert found[0].get("url") == "https://example.com/"


def test_removing_a_run_level_mouse_over_link_works(make_deck):
    from kitchensink4ppt.core.package import PptxPackage
    from kitchensink4ppt.ops import links, read as _read
    from kitchensink4ppt.ops import shapes as _shapes, text as _text

    pkg = PptxPackage(make_deck("hover_remove.pptx"))
    box = _text.insert_textbox(pkg, 0, "Click here for details", 1, 1, 4, 1)
    links.set_hyperlink(
        pkg, 0, {"shape_id": box["shape_id"], "paragraph": 0},
        url="https://example.com/",
    )
    _run_hover_link(pkg, 0, box["shape_id"], "rId2")
    result = links.remove_hyperlink(pkg, 0, box["shape_id"])
    assert result["removed"] >= 1
    part = _read.slide_table(pkg)[0]["part"]
    elem, _chain = _shapes._find_shape(pkg, part, box["shape_id"])
    assert not list(elem.iter(qn("a:hlinkMouseOver"))), (
        "the mouse-over link survived a removal that reported success"
    )


def test_deleting_a_slide_neuters_a_run_level_mouse_over_jump(make_deck):
    """The dangerous one: the rel goes, and an unswept r:id is left behind
    pointing at a relationship that no longer exists."""
    from kitchensink4ppt.core.package import PptxPackage
    from kitchensink4ppt.ops import links, read as _read
    from kitchensink4ppt.ops import shapes as _shapes, slides as _slides
    from kitchensink4ppt.ops import text as _text

    pkg = PptxPackage(make_deck("hover_jump.pptx", extra_slides=2))
    box = _text.insert_textbox(pkg, 0, "Jump to the last slide", 1, 1, 4, 1)
    links.set_hyperlink(
        pkg, 0, {"shape_id": box["shape_id"], "paragraph": 0}, to_slide=2
    )
    _run_hover_link(pkg, 0, box["shape_id"], "rId2")
    _slides.delete_slide(pkg, 2)
    part = _read.slide_table(pkg)[0]["part"]
    elem, _chain = _shapes._find_shape(pkg, part, box["shape_id"])
    left = list(elem.iter(qn("a:hlinkMouseOver")))
    assert not left, (
        "the jump's rel was removed but its mouse-over element was left "
        "behind with a dangling r:id"
    )


# ----------------------------------------------------------------- G2b
#
# The same empty-snapshot flaw one layer up. powerpnt_pids() returned an
# empty set whenever tasklist errored, timed out or printed something it
# could not parse, so "unknown" and "nothing was running" were the same
# value. _powerpoint_locked then computed launched = not before_pids, and
# a session that had merely ATTACHED to the owner's PowerPoint believed it
# had launched it and called Quit() on the way out.


@pytestmark_win
def test_an_unreadable_process_table_is_unknown_not_empty(monkeypatch):
    from kitchensink4ppt.com import bridge

    def boom(*_a, **_k):
        raise OSError("tasklist is not available")

    monkeypatch.setattr(bridge.subprocess, "run", boom)
    assert bridge.powerpnt_pids() is None
    assert bridge.powerpnt_count() == -1


@pytestmark_win
@pytest.mark.parametrize(
    "stdout, code",
    [
        ("", 0),                                    # no output at all
        ("", 1),                                    # errored, said nothing
        ('"POWERPNT.EXE"\n', 0),                    # row we cannot parse
        ('"POWERPNT.EXE","not-a-pid","Console"\n', 0),
    ],
)
def test_an_unparsable_process_table_is_unknown(monkeypatch, stdout, code):
    from kitchensink4ppt.com import bridge

    class Result:
        pass

    def fake_run(*_a, **_k):
        r = Result()
        r.stdout, r.stderr, r.returncode = stdout, "", code
        return r

    monkeypatch.setattr(bridge.subprocess, "run", fake_run)
    assert bridge.powerpnt_pids() is None


@pytestmark_win
def test_a_genuine_no_match_is_still_an_empty_set(monkeypatch):
    """The regression this fix must not cause: tasklist's INFO line means
    nothing is running, and the server must still be able to own what it
    starts next."""
    from kitchensink4ppt.com import bridge

    class Result:
        stdout = "INFO: No tasks are running which match the criteria.\n"
        stderr = ""
        returncode = 0

    monkeypatch.setattr(bridge.subprocess, "run", lambda *a, **k: Result())
    assert bridge.powerpnt_pids() == set()
    assert bridge.powerpnt_count() == 0


class _FakeApp:
    """Enough of PowerPoint.Application to run a session to its cleanup."""

    def __init__(self):
        self.quit_calls = 0
        self.DisplayAlerts = None
        self.Presentations = []

    def Quit(self):  # noqa: N802 - COM spelling
        self.quit_calls += 1


def _run_session(monkeypatch, bridge, table):
    """Enter and leave one _powerpoint() session against a fake table and a
    fake DispatchEx, and hand back (app, session.launched)."""
    app = _FakeApp()
    monkeypatch.setattr(bridge, "powerpnt_pids", table)
    monkeypatch.setattr(
        bridge, "_com_modules", lambda: (_Pythoncom(), _Win32(app))
    )
    monkeypatch.setattr(bridge, "_ensure_apartment", lambda _p: "initialized")
    monkeypatch.setattr(bridge, "QUIT_POLL_SECONDS", 5.0)
    seen = {}
    with bridge._powerpoint() as session:
        seen["launched"] = session.launched
    return app, seen["launched"]


class _Win32:
    def __init__(self, app):
        self._app = app

    def DispatchEx(self, _progid):  # noqa: N802 - COM spelling
        return self._app


@pytestmark_win
def test_an_unreadable_table_never_quits_the_instance_it_attached_to(
    monkeypatch,
):
    """The harm: tasklist fails, the owner has PowerPoint open, the server
    attaches to it and then quits it with the owner's unsaved work in it."""
    from kitchensink4ppt.com import bridge

    app, launched = _run_session(monkeypatch, bridge, lambda: None)
    assert launched is False, "an unknown table is not ownership"
    assert app.quit_calls == 0, "the owner's PowerPoint was quit"
    assert not bridge._SELF_LAUNCHED_PIDS, "the kill-switch was armed"


@pytestmark_win
def test_an_unreadable_table_leaves_the_kill_switch_disarmed(monkeypatch):
    from kitchensink4ppt.com import bridge
    import threading

    _app, _launched = _run_session(monkeypatch, bridge, lambda: None)
    assert bridge._kill_self_launched_for_thread(
        threading.get_ident()
    ) is False


@pytestmark_win
def test_a_known_empty_table_still_owns_and_quits_what_it_started(
    monkeypatch,
):
    """The normal path must behave exactly as before: nothing running, we
    start one, we quit it."""
    from kitchensink4ppt.com import bridge

    calls = {"n": 0}

    def table():
        calls["n"] += 1
        if calls["n"] == 1:
            return set()          # entry: nothing was running
        if calls["n"] == 2:
            return {9191}         # our DispatchEx started this one
        return set()              # after Quit: it exited

    app, launched = _run_session(monkeypatch, bridge, table)
    assert launched is True
    assert app.quit_calls == 1


@pytestmark_win
def test_a_known_busy_table_still_attaches_and_does_not_quit(monkeypatch):
    from kitchensink4ppt.com import bridge

    app, launched = _run_session(monkeypatch, bridge, lambda: {321})
    assert launched is False
    assert app.quit_calls == 0


@pytestmark_win
def test_powerpoint_status_does_not_call_an_unknown_table_not_running(
    monkeypatch,
):
    from kitchensink4ppt.com import bridge

    monkeypatch.setattr(bridge, "powerpnt_pids", lambda: None)
    out = bridge.powerpoint_status()
    assert out["powerpoint_running"] is not True
    assert "error" in out and "process table" in out["error"]


# ------------------------------------------------------------------ S1


@metrics
def test_the_metrics_note_states_a_bounded_allowance_not_a_direction():
    """The note used to claim the model "errs wide rather than narrow",
    which is an absolute the calibration cannot support. It states what the
    allowance IS and what that costs the reader instead."""
    sp, body, _p = _metrics_shape([("Some words", "Calibri", 1200)])
    note = _overflow(sp, body)["note"]
    assert "errs wide" not in note
    assert "3.5 pt allowance" in note
    assert "calibrated on tested PowerPoint layouts" in note
    assert "may be reported as wrapping" in note
    assert note.endswith("PowerPoint's rendering is still the final authority")
