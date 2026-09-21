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

import contextlib
import inspect
import sys
import threading

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
def test_the_failed_start_reports_what_it_saw_and_what_it_did_not_do(
    monkeypatch,
):
    """G2 and R7-2: the refusal used to claim a process 'was ended', then
    that it 'was left running'. It may say only what it observed and what
    this call did not do; liveness at read time is not knowable here."""
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
    assert "left running" not in message
    assert (
        "A PowerPoint process appeared during the failed start: pid 7777. "
        "This call did not force-end it; its current state was not "
        "re-checked."
    ) in message


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


class _Collection:
    """A COM collection stub: what Presentations and Windows need to be."""

    def __init__(self, count=0):
        self.Count = count


class _FakeApp:
    """Enough of PowerPoint.Application to run a session to its cleanup.

    Since round 4 that includes the acquisition token: a just-created
    automation instance is not visible and holds no presentations and no
    windows, and that is the only part of ownership a concurrent launch
    cannot fake.
    """

    def __init__(self, visible=0, presentations=0, windows=0):
        self.quit_calls = 0
        self.DisplayAlerts = None
        self.Visible = visible
        self.Presentations = _Collection(presentations)
        self.Windows = _Collection(windows)

    def Quit(self):  # noqa: N802 - COM spelling
        self.quit_calls += 1


def _run_session(monkeypatch, bridge, table, app=None, during=None):
    """Enter and leave one _powerpoint() session against a fake table and a
    fake DispatchEx, and hand back (app, session.launched).

    `during` runs INSIDE the session, which is how a test represents the
    user or another client arriving while the operation is in flight.
    """
    app = app if app is not None else _FakeApp()
    monkeypatch.setattr(bridge, "powerpnt_pids", table)
    monkeypatch.setattr(
        bridge, "_com_modules", lambda: (_Pythoncom(), _Win32(app))
    )
    monkeypatch.setattr(bridge, "_ensure_apartment", lambda _p: "initialized")
    monkeypatch.setattr(bridge, "QUIT_POLL_SECONDS", 5.0)
    seen = {}
    with bridge._powerpoint() as session:
        seen["launched"] = session.launched
        if during is not None:
            during()
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
    assert bridge._self_launched_pids_for_thread(
        threading.get_ident()
    ) == set()


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


# ================================================================ ROUND 4


# ------------------------------------------------------------ R4-1


@pytestmark_win
def test_a_nonzero_tasklist_with_stdout_text_is_unknown(monkeypatch):
    """tasklist prints 'ERROR: Access is denied.' on STDOUT and exits 1.
    Only the empty-stdout case was treated as unknown, so an access error
    parsed as a clean empty table."""
    from kitchensink4ppt.com import bridge

    class Result:
        stdout = "ERROR: Access is denied.\n"
        stderr = ""
        returncode = 1

    monkeypatch.setattr(bridge.subprocess, "run", lambda *a, **k: Result())
    assert bridge.powerpnt_pids() is None


@pytestmark_win
def test_the_users_instance_handed_to_us_by_dispatch_is_never_ours(
    monkeypatch,
):
    """R4-1, the blocker. PowerPoint is a single-instance automation
    server: a user who starts it between the empty snapshot and DispatchEx
    is handed to US, and the before/after pid difference names THEIR
    process. The reviewer's probe, adapted."""
    import threading

    from kitchensink4ppt.com import bridge

    snapshots = iter((set(), {4242}, set(), set()))
    # Their PowerPoint: visible, with the deck they were working on.
    theirs = _FakeApp(visible=-1, presentations=1, windows=1)
    app, launched = _run_session(
        monkeypatch, bridge, lambda: next(snapshots), app=theirs
    )
    assert launched is False, "a user's instance was classified as ours"
    assert app.quit_calls == 0, "the user's PowerPoint was quit"
    assert threading.get_ident() not in bridge._SELF_LAUNCHED_PIDS


@pytestmark_win
def test_an_instance_that_is_not_freshly_created_is_never_ours(monkeypatch):
    """The same race one step subtler: the user launched it but has not
    opened anything yet, so only Visible gives them away."""
    from kitchensink4ppt.com import bridge

    snapshots = iter((set(), {4242}, set(), set()))
    app, launched = _run_session(
        monkeypatch, bridge, lambda: next(snapshots),
        app=_FakeApp(visible=-1),
    )
    assert launched is False
    assert app.quit_calls == 0


@pytestmark_win
def test_more_than_one_new_pid_is_not_a_single_candidate(monkeypatch):
    from kitchensink4ppt.com import bridge

    snapshots = iter((set(), {1111, 2222}, set(), set()))
    app, launched = _run_session(monkeypatch, bridge, lambda: next(snapshots))
    assert launched is False
    assert app.quit_calls == 0


@pytestmark_win
def test_a_presentation_opened_by_someone_else_stops_the_quit(monkeypatch):
    """Another client attaches to our hidden instance and opens work in it
    while we run. At release its Presentations.Count is not zero, so the
    application is left running and only what we opened was closed."""
    from kitchensink4ppt.com import bridge

    calls = {"n": 0}

    def table():
        calls["n"] += 1
        return set() if calls["n"] == 1 else {9191}

    app = _FakeApp()

    def arrive():
        app.Presentations.Count = 1  # their deck, opened mid-call

    _app, launched = _run_session(
        monkeypatch, bridge, table, app=app, during=arrive
    )
    assert launched is True, "acquisition was legitimate"
    assert app.quit_calls == 0, "someone else's work was in that instance"


@pytestmark_win
def test_the_user_making_it_visible_stops_the_quit(monkeypatch):
    from kitchensink4ppt.com import bridge

    calls = {"n": 0}

    def table():
        calls["n"] += 1
        return set() if calls["n"] == 1 else {9191}

    app = _FakeApp()
    _app, launched = _run_session(
        monkeypatch, bridge, table, app=app,
        during=lambda: setattr(app, "Visible", -1),
    )
    assert launched is True
    assert app.quit_calls == 0


@pytestmark_win
def test_nothing_in_the_com_package_can_end_a_powerpoint_process():
    """1.3.1 removes force-kill from the timeout path, and there is no
    other place in the package that ends a process. A source guard,
    because this is easy to reintroduce and expensive to notice."""
    import pathlib

    from kitchensink4ppt.com import bridge

    root = pathlib.Path(bridge.__file__).parent
    offenders = []
    for path in sorted(root.glob("*.py")):
        body = path.read_text(encoding="utf-8")
        for forbidden in ("taskkill", "TerminateProcess", ".Terminate("):
            if forbidden in body:
                offenders.append(path.name + " contains " + forbidden)
    assert not offenders, "; ".join(offenders)


@pytestmark_win
def test_a_timed_out_operation_names_the_pid_and_ends_nothing(monkeypatch):
    """It used to terminate that process. A timeout cannot revalidate a
    hung apartment, so the refusal reports and nothing is ended.

    Deterministic: the worker blocks on an event the test owns, so nothing
    here depends on how fast this machine is (R6-3)."""
    from kitchensink4ppt.com import bridge
    from kitchensink4ppt.core.errors import PowerPointBlocked

    real_run = bridge.subprocess.run

    def no_kill(cmd, **kwargs):
        # Reading the process table is fine; ENDING one is what is gone.
        if cmd and cmd[0] == "taskkill":
            pytest.fail("the timeout path tried to end a process")
        return real_run(cmd, **kwargs)

    monkeypatch.setattr(bridge.subprocess, "run", no_kill)
    with _timeout_harness(monkeypatch, bridge, token={5150}) as h:
        with pytest.raises(PowerPointBlocked) as exc_info:
            h.run()
    message = str(exc_info.value)
    assert "5150" in message
    assert "terminated" not in message
    assert "was left running" not in message


# ------------------------------------------------------------ R4-2


@metrics
def test_the_inherited_face_comes_from_the_runs_own_level():
    """R4-2: _body_typeface took the first a:defRPr anywhere in the shape's
    lstStyle, so a level-two run was measured in level one's face."""
    sp, body, p = _metrics_shape([("Inherited face", "Calibri", 1200)])
    lst = body.find(qn("a:lstStyle"))
    # Level one names a face; level two, where this paragraph actually
    # sits, names none. A level-one face must not reach a level-two run.
    lvl1 = etree.SubElement(lst, qn("a:lvl1pPr"))
    defrpr = etree.SubElement(lvl1, qn("a:defRPr"))
    etree.SubElement(defrpr, qn("a:latin")).set(
        "typeface", "NoSuchFontFamilyAnywhere"
    )
    etree.SubElement(
        etree.SubElement(lst, qn("a:lvl2pPr")), qn("a:defRPr")
    )
    p.find(qn("a:pPr")).set("lvl", "1")  # the SECOND outline level
    rpr = p.find(qn("a:r") + "/" + qn("a:rPr"))
    rpr.remove(rpr.find(qn("a:latin")))  # it has to inherit a face
    rec = _overflow(sp, body)
    reason = rec.get("method_reason") or ""
    assert rec.get("typeface") != "NoSuchFontFamilyAnywhere"
    assert "NoSuchFontFamilyAnywhere" not in reason, (
        "the run inherited a face from another outline level"
    )


@metrics
@pytest.mark.parametrize("name, value", [
    ("marL", "not-a-number"),
    ("indent", "12.5"),
    ("marL", "99999999999"),
])
def test_a_malformed_indent_is_unmeasurable_not_zero(name, value):
    """R4-2: a malformed explicit marL/indent became zero, which measures
    a frame nobody has."""
    sp, body, p = _metrics_shape([("Some text", "Calibri", 1200)])
    p.find(qn("a:pPr")).set(name, value)
    rec = _overflow(sp, body)
    assert rec["method"] == "estimate"
    assert name in rec["method_reason"]


# ------------------------------------------------------------ R4-3


@pytest.mark.parametrize("hover_tag", ["a:hlinkMouseOver", "a:hlinkHover"])
@pytest.mark.parametrize("kind", ["url", "slide"])
def test_setting_a_link_replaces_an_existing_run_hover_link(
    make_deck, hover_tag, kind
):
    """R4-3: _set_run_hlink removed only a:hlinkClick, so the old hover
    element and its relationship survived beside the new link."""
    from kitchensink4ppt.core.package import PptxPackage
    from kitchensink4ppt.ops import links
    from kitchensink4ppt.ops import read as _read
    from kitchensink4ppt.ops import shapes as _shapes
    from kitchensink4ppt.ops import text as _text

    rid_attr = "{" + NSMAP["r"] + "}id"
    pkg = PptxPackage(make_deck("hover_replace.pptx", extra_slides=2))
    box = _text.insert_textbox(pkg, 0, "Click here for details", 1, 1, 4, 1)
    target = {"shape_id": box["shape_id"], "paragraph": 0}
    links.set_hyperlink(pkg, 0, target, url="https://old.example.com/")

    part = _read.slide_table(pkg)[0]["part"]
    elem, _chain = _shapes._find_shape(pkg, part, box["shape_id"])
    old_rid = None
    for rpr in elem.iter(qn("a:rPr")):
        for el in list(rpr):
            if el.tag == qn("a:hlinkClick"):
                old_rid = el.get(rid_attr)
                el.tag = qn(hover_tag)
    pkg.mark_dirty(part)
    assert old_rid

    if kind == "url":
        links.set_hyperlink(pkg, 0, target, url="https://new.example.com/")
    else:
        links.set_hyperlink(pkg, 0, target, to_slide=2)

    elem, _chain = _shapes._find_shape(pkg, part, box["shape_id"])
    assert not list(elem.iter(qn(hover_tag))), (
        "the old hover link survived the replacement"
    )
    rids = {rel.get("Id") for rel in pkg.rels_for(part).getroot()}
    assert old_rid not in rids, "the replaced link's relationship was kept"


@pytest.mark.parametrize("kind", ["url", "slide"])
def test_setting_a_link_replaces_an_existing_shape_hover_link(make_deck, kind):
    """The shape-level equivalent: a:hlinkHover inside p:cNvPr is the
    correct element there, and _set_cnvpr_hlink left it behind."""
    from kitchensink4ppt.core.package import PptxPackage
    from kitchensink4ppt.ops import links
    from kitchensink4ppt.ops import read as _read
    from kitchensink4ppt.ops import shapes as _shapes
    from kitchensink4ppt.ops import text as _text

    rid_attr = "{" + NSMAP["r"] + "}id"
    pkg = PptxPackage(make_deck("hover_shape.pptx", extra_slides=2))
    box = _text.insert_textbox(pkg, 0, "A box", 1, 1, 4, 1)
    links.set_hyperlink(pkg, 0, box["shape_id"], url="https://old.example/")

    part = _read.slide_table(pkg)[0]["part"]
    elem, _chain = _shapes._find_shape(pkg, part, box["shape_id"])
    cnvpr = elem.find(qn("p:nvSpPr") + "/" + qn("p:cNvPr"))
    old_rid = None
    for el in list(cnvpr):
        if el.tag == qn("a:hlinkClick"):
            old_rid = el.get(rid_attr)
            el.tag = qn("a:hlinkHover")
    pkg.mark_dirty(part)
    assert old_rid

    if kind == "url":
        links.set_hyperlink(pkg, 0, box["shape_id"], url="https://new.example/")
    else:
        links.set_hyperlink(pkg, 0, box["shape_id"], to_slide=2)

    elem, _chain = _shapes._find_shape(pkg, part, box["shape_id"])
    cnvpr = elem.find(qn("p:nvSpPr") + "/" + qn("p:cNvPr"))
    assert cnvpr.find(qn("a:hlinkHover")) is None, (
        "the old shape-level hover link survived the replacement"
    )
    rids = {rel.get("Id") for rel in pkg.rels_for(part).getroot()}
    assert old_rid not in rids


# ------------------------------------------------------------ R4-4


@pytestmark_win
@pytest.mark.parametrize("hresult_name, error_name", [
    ("RPC_E_CALL_REJECTED", "PowerPointBusy"),
    ("RPC_E_DISCONNECTED", "PowerPointDisconnected"),
])
def test_a_busy_or_gone_powerpoint_at_slide_id_is_not_swallowed(
    hresult_name, error_name
):
    """R4-4: _slide_id_of caught EVERY exception and returned None, so a
    busy or disconnected application never reached the classifier and the
    walk carried on to a verdict about the file."""
    from kitchensink4ppt.com import bridge
    from kitchensink4ppt.core import errors

    class _Slide:
        @property
        def SlideID(self):  # noqa: N802 - COM spelling
            raise _com_error(getattr(bridge, hresult_name))

    with pytest.raises(getattr(errors, error_name)):
        bridge._slide_id_of(_Slide())


@pytestmark_win
def test_an_ordinary_slide_id_read_failure_is_still_just_none():
    from kitchensink4ppt.com import bridge

    class _Slide:
        @property
        def SlideID(self):  # noqa: N802 - COM spelling
            raise ValueError("not a number")

    assert bridge._slide_id_of(_Slide()) is None


# ================================================================ ROUND 5


# ------------------------------------------------------------ R5-1, R6-1
#
# Everything below is EVENT-DRIVEN. Round 5 proved late completion with a
# 10.8 second sleep racing _run_bounded's grace wait, and on the
# reviewer's machine the read-only dialog inspection took long enough that
# the worker won the race and the test failed. No sleep decides an outcome
# here (R6-3): the worker blocks on an event the test owns, and the test
# releases it from inside the stubbed dialog probe, which runs at a known
# point in the grace window.


class _TimeoutHarness:
    """One _run_bounded call whose worker the test drives by hand.

    `release` unblocks the worker. `during_diagnostics` runs inside the
    stubbed dialog probe, which _run_bounded calls after its grace wait
    and before it decides anything: that is the deterministic hook for
    "something happened while the deadline was passing".
    """

    def __init__(self, bridge, token=None, outcome=None, during=None):
        import threading

        self.bridge = bridge
        self.token = token
        self.outcome = outcome if outcome is not None else (lambda: {"ok": 1})
        self.during = during
        self.release = threading.Event()
        self.entered = threading.Event()
        self.finished = threading.Event()
        self.tid = None

    def body(self):
        import threading

        self.tid = threading.get_ident()
        if self.token:
            self.bridge._SELF_LAUNCHED_PIDS[self.tid] = set(self.token)
        self.entered.set()
        self.release.wait(30)
        try:
            return self.outcome()
        finally:
            self.finished.set()

    def diagnostics(self):
        if self.during is not None:
            self.during(self)
        return []

    def run(self, timeout=0.05):
        return self.bridge._run_bounded("harness-op", timeout, self.body)

    def clear_token(self):
        self.bridge._SELF_LAUNCHED_PIDS.pop(self.tid, None)

    def finish_now(self):
        """Let the worker run to completion and wait for it, from inside
        the grace window."""
        self.release.set()
        assert self.finished.wait(30), "the worker never finished"


@contextlib.contextmanager
def _timeout_harness(monkeypatch, bridge, token=None, outcome=None,
                     during=None):
    from kitchensink4ppt.com import dialogs as _dialogs

    harness = _TimeoutHarness(bridge, token, outcome, during)
    # A short grace keeps the test fast; the hook below, not the clock, is
    # what decides every outcome.
    monkeypatch.setattr(bridge, "TIMEOUT_GRACE_SECONDS", 2.0)
    monkeypatch.setattr(
        _dialogs, "pending_dialogs", lambda *a, **k: harness.diagnostics()
    )
    try:
        yield harness
    finally:
        harness.release.set()
        harness.finished.wait(30)
        harness.clear_token()


@pytestmark_win
def test_a_worker_still_running_when_the_refusal_is_built_gets_the_refusal(
    monkeypatch,
):
    """R6-3(a). The worker is held, so the refusal is what comes back, and
    it carries the whole uncancelled-operation warning."""
    from kitchensink4ppt.com import bridge
    from kitchensink4ppt.core.errors import PowerPointBlocked

    with _timeout_harness(monkeypatch, bridge) as h:
        with pytest.raises(PowerPointBlocked) as exc_info:
            h.run()
        assert not h.finished.is_set(), "the worker finished before the refusal"
    message = str(exc_info.value)
    assert "was aborted" not in message
    assert "The operation was NOT cancelled" in message
    assert "may still finish and save its output" in message
    assert "Do not retry yet" in message
    assert (
        "call powerpoint_status and wait until it reports no COM operation "
        "in progress"
    ) in message
    assert "check the output file before repeating the call" in message


# ------------------------------------------------- R7-3 stage control
#
# Round 6 proved the two-stage ordering with tests that still left the
# scheduler a vote: the stage-one case used timeout=0 and ASSUMED the
# first wait could not observe a brand-new thread (200 instrumented runs
# say it always did, so the first post-grace check was never reached),
# and the stage-two harness set its own event before _run_bounded set its
# internal one. Both are replaced here.
#
# _COMPLETION_EVENT_FACTORY is the seam. The event below is real, and the
# worker sets it exactly as in production; what the test decides is WHICH
# WAIT observes it, by running an action at a chosen wait. _run_bounded
# waits on it exactly twice, the caller's deadline and then the grace, so
# the stage is named by index and nothing is left to timing.

_STAGE_JOIN = 30.0


class _StagedCompletion:
    """_run_bounded's internal completion event, with the observing stage
    chosen by the test rather than by the scheduler (R7-3)."""

    INITIAL_WAIT = 0
    GRACE_WAIT = 1

    def __init__(self, release, actions=None):
        self._inner = threading.Event()
        self._release = release
        self._actions = dict(actions or {})
        self.waits = []

    # the event interface _run_bounded uses
    def set(self):
        self._inner.set()

    def is_set(self):
        return self._inner.is_set()

    def wait(self, timeout=None):
        stage = len(self.waits)
        self.waits.append(timeout)
        action = self._actions.get(stage)
        if action is not None:
            action(self)
        return self._inner.is_set()

    # what the test drives
    def finish_worker(self, *_):
        """Let the worker run to completion and block until its OWN
        completion signal is set, so no later step can outrun it."""
        self._release.set()
        assert self._inner.wait(_STAGE_JOIN), "the worker never signalled"


@contextlib.contextmanager
def _staged_run(monkeypatch, actions=None, during_diagnostics=None):
    """One _run_bounded call with the completion seam installed and the
    dialog probe counted. Yields (run, staged, probes)."""
    from kitchensink4ppt.com import bridge
    from kitchensink4ppt.com import dialogs as _dialogs

    release = threading.Event()
    staged = _StagedCompletion(release, actions)
    probes = []

    def probe(*_a, **_k):
        probes.append(1)
        if during_diagnostics is not None:
            during_diagnostics(staged)
        return []

    monkeypatch.setattr(bridge, "_COMPLETION_EVENT_FACTORY", lambda: staged)
    monkeypatch.setattr(_dialogs, "pending_dialogs", probe)

    def body():
        release.wait(_STAGE_JOIN)
        return {"slides": 7}

    try:
        yield (lambda: bridge._run_bounded("staged-op", 30.0, body),
               staged, probes)
    finally:
        release.set()


@pytestmark_win
def test_completion_during_the_initial_wait_returns_without_any_grace(
    monkeypatch,
):
    """R7-3 stage ZERO: the ordinary success. The deadline wait itself
    observes completion, so no grace and no inspection happen at all."""
    with _staged_run(
        monkeypatch,
        actions={_StagedCompletion.INITIAL_WAIT:
                 _StagedCompletion.finish_worker},
    ) as (run, staged, probes):
        assert run() == {"slides": 7}
    assert len(staged.waits) == 1, "the grace wait ran on a finished worker"
    assert probes == [], "a finished worker was inspected anyway"


@pytestmark_win
def test_completion_during_the_grace_wait_skips_the_diagnostics(
    monkeypatch,
):
    """R7-3 stage ONE, the first post-grace check. Completion becomes
    visible during the GRACE wait, and the check that runs the instant
    that wait ends returns it, so a finished success is never delayed
    behind a process-table call that can itself block."""
    with _staged_run(
        monkeypatch,
        actions={_StagedCompletion.GRACE_WAIT:
                 _StagedCompletion.finish_worker},
    ) as (run, staged, probes):
        assert run() == {"slides": 7}
    assert len(staged.waits) == 2, "the grace wait was not reached"
    assert probes == [], (
        "a completed result waited behind the process and dialog inspection"
    )


@pytestmark_win
def test_completion_during_the_diagnostics_returns_on_the_second_check(
    monkeypatch,
):
    """R7-3 stage TWO, the second check. The inspection can itself take
    long enough for the worker to finish inside it, and the answer it
    gives then is still the worker's own."""
    with _staged_run(
        monkeypatch,
        during_diagnostics=_StagedCompletion.finish_worker,
    ) as (run, staged, probes):
        assert run() == {"slides": 7}
    assert len(staged.waits) == 2
    assert probes == [1], "the second check ran without any inspection"


@pytestmark_win
def test_a_worker_that_never_completes_gets_the_refusal(monkeypatch):
    """R7-3 stage THREE. Neither check ever sees completion, the
    inspection did run, and the refusal is what comes back."""
    from kitchensink4ppt.core.errors import PowerPointBlocked

    with _staged_run(monkeypatch) as (run, staged, probes):
        with pytest.raises(PowerPointBlocked) as exc_info:
            run()
        assert len(staged.waits) == 2
        assert probes == [1]
    assert "The operation was NOT cancelled" in str(exc_info.value)


# ----------------------------------------------- R7-1 queued ghost write
#
# A caller that gives up while its worker is still QUEUED on the COM
# serialization lock used to get PowerPointBusy and an invitation to
# retry, while the worker stayed in the queue and ran fn() anyway once the
# holder released. A retried non-idempotent edit therefore ran twice: a
# slide deletion by index deleted one slide when the abandoned worker
# finally started and another when the retry ran. One atomic decision per
# call now settles it, and these tests prove both sides of it.

_ABANDONED_SENTENCE = (
    " The queued call was abandoned before its operation ran; it made no "
    "document changes. It is safe to retry after powerpoint_status reports "
    "no COM operation in progress."
)


@contextlib.contextmanager
def _lock_holder(name="existing-write"):
    """Hold the process-wide COM serialization lock until released."""
    from kitchensink4ppt.com import serial as _serial

    held = threading.Event()
    release = threading.Event()

    def hold():
        with _serial.com_operation(name):
            held.set()
            release.wait(_STAGE_JOIN)

    t = threading.Thread(target=hold, daemon=True, name="ks4p-test-holder")
    t.start()
    assert held.wait(_STAGE_JOIN), "the holder never took the lock"
    try:
        yield release
    finally:
        release.set()
        t.join(_STAGE_JOIN)


def _queued_worker(name):
    for t in threading.enumerate():
        if t.name == f"ks4p-{name}":
            return t
    return None


@pytestmark_win
def test_a_call_abandoned_while_queued_never_runs_its_operation():
    """R7-1(a). The holder keeps the lock past the caller's wait, so the
    caller is told the call was abandoned. The proof is what happens
    AFTER the holder releases: the queued worker reaches the front of the
    queue, loses the decision, and exits without ever entering fn."""
    from kitchensink4ppt.com import bridge
    from kitchensink4ppt.core.errors import PowerPointBusy

    calls = []
    with _lock_holder() as release_holder:
        with pytest.raises(PowerPointBusy) as exc_info:
            bridge._run_bounded(
                "queued-write", 0.05, lambda: calls.append("ran")
            )
        assert calls == [], "the operation ran before the caller gave up"
        worker = _queued_worker("queued-write")
        assert worker is not None, "the queued worker was not found"
        release_holder.set()
        worker.join(_STAGE_JOIN)
        assert not worker.is_alive(), "the queued worker never finished"
    assert calls == [], (
        "the abandoned worker executed the operation after the holder "
        "released the lock, so a caller told nothing had happened had it "
        "happen behind its back"
    )
    assert _ABANDONED_SENTENCE in str(exc_info.value)


@pytestmark_win
def test_a_retry_after_an_abandoned_call_runs_the_operation_once():
    """R7-1(c). The refusal invites a retry, so the retry must be the
    only execution there ever is."""
    from kitchensink4ppt.com import bridge
    from kitchensink4ppt.core.errors import PowerPointBusy

    calls = []

    def delete_a_slide():
        calls.append("ran")
        return {"deleted": 1}

    with _lock_holder() as release_holder:
        with pytest.raises(PowerPointBusy):
            bridge._run_bounded("queued-write", 0.05, delete_a_slide)
        worker = _queued_worker("queued-write")
        release_holder.set()
        if worker is not None:
            worker.join(_STAGE_JOIN)
    assert bridge._run_bounded(
        "queued-write", 30.0, delete_a_slide
    ) == {"deleted": 1}
    assert calls == ["ran"], (
        "the non-idempotent operation ran twice, once from the abandoned "
        "queued worker and once from the retry"
    )


@pytestmark_win
def test_a_worker_that_wins_the_start_decision_is_never_told_to_retry(
    monkeypatch,
):
    """R7-1(b). The other side of the same decision: the worker takes the
    lock and starts fn in the instant before the caller would have
    abandoned it. The caller must then follow the grace path and get the
    worker's own answer, never a queue-contention retry.

    The completion seam makes the order explicit rather than hoped for:
    the lock is handed over during the caller's deadline wait, and that
    wait does not return until fn has provably been entered."""
    from kitchensink4ppt.com import bridge
    from kitchensink4ppt.com import dialogs as _dialogs
    from kitchensink4ppt.core.errors import PowerPointBusy

    calls = []
    entered = threading.Event()
    finish = threading.Event()

    def op():
        calls.append("ran")
        entered.set()
        finish.wait(_STAGE_JOIN)
        return {"slides": 4}

    with _lock_holder() as release_holder:

        def start_the_worker(_staged):
            release_holder.set()
            assert entered.wait(_STAGE_JOIN), "the worker never started fn"

        staged = _StagedCompletion(
            finish,
            {
                _StagedCompletion.INITIAL_WAIT: start_the_worker,
                _StagedCompletion.GRACE_WAIT: _StagedCompletion.finish_worker,
            },
        )
        monkeypatch.setattr(
            bridge, "_COMPLETION_EVENT_FACTORY", lambda: staged
        )
        monkeypatch.setattr(_dialogs, "pending_dialogs", lambda *a, **k: [])
        try:
            assert bridge._run_bounded("racing-write", 30.0, op) == {
                "slides": 4
            }
        except PowerPointBusy as exc:  # pragma: no cover - the defect
            pytest.fail(f"a started operation was reported as queued: {exc}")
    assert calls == ["ran"], "the operation did not run exactly once"


@pytestmark_win
def test_a_worker_that_raises_during_the_grace_surfaces_that_error(
    monkeypatch,
):
    """R6-3(c). Its own error, not a timeout that hides it."""
    from kitchensink4ppt.com import bridge
    from kitchensink4ppt.core.errors import DocumentLocked

    def boom():
        raise DocumentLocked("the deck is open in PowerPoint")

    with _timeout_harness(
        monkeypatch, bridge, outcome=boom, during=lambda h: h.finish_now(),
    ) as h:
        with pytest.raises(DocumentLocked, match="open in PowerPoint"):
            h.run()


@pytestmark_win
def test_a_token_cleared_during_the_grace_is_not_reported_as_a_pid(
    monkeypatch,
):
    """R6-3(d). The worker recorded pid 5150, then its session ended and
    cleared the token while the deadline was passing. The evidence is read
    when the refusal is BUILT, so the refusal cannot name a pid that is no
    longer recorded, and the new wording cannot say it is still running."""
    from kitchensink4ppt.com import bridge
    from kitchensink4ppt.core.errors import PowerPointBlocked

    with _timeout_harness(
        monkeypatch, bridge, token={5150},
        during=lambda h: h.clear_token(),
    ) as h:
        with pytest.raises(PowerPointBlocked) as exc_info:
            h.run()
    message = str(exc_info.value)
    assert "5150" not in message
    assert "was left running" not in message
    assert (
        "no newly started PowerPoint process could be identified; no "
        "process was force-ended"
    ) in message


@pytestmark_win
def test_a_recorded_pid_is_reported_state_neutrally(monkeypatch):
    """R6-2. What was observed and what this path did not do; no claim
    about the process's current state and none about ownership."""
    from kitchensink4ppt.com import bridge
    from kitchensink4ppt.core.errors import PowerPointBlocked

    with _timeout_harness(monkeypatch, bridge, token={5150}) as h:
        with pytest.raises(PowerPointBlocked) as exc_info:
            h.run()
    message = str(exc_info.value)
    assert (
        "(PowerPoint process pid 5150 was not running when this call began; "
        "this call did not force-end it, and its current state was not "
        "re-checked)"
    ) in message
    assert "this call started" not in message
    assert "was left running" not in message


@pytestmark_win
@pytest.mark.parametrize("table", [
    pytest.param(lambda: None, id="unknown-process-table"),
    pytest.param(lambda: {321}, id="ownership-gate-refused"),
    pytest.param(set, id="worker-never-touched-powerpoint"),
])
def test_every_no_token_case_gets_the_same_neutral_clause(monkeypatch, table):
    """R6-3(e). An empty token is NOT evidence that PowerPoint was already
    running: the table may be unreadable, the ownership gate may have
    refused, no token may have been recorded yet, or the worker may never
    have touched PowerPoint at all. One honest sentence covers all of it."""
    from kitchensink4ppt.com import bridge
    from kitchensink4ppt.core.errors import PowerPointBlocked

    monkeypatch.setattr(bridge, "powerpnt_pids", table)
    with _timeout_harness(monkeypatch, bridge) as h:
        with pytest.raises(PowerPointBlocked) as exc_info:
            h.run()
    message = str(exc_info.value)
    assert (
        "no newly started PowerPoint process could be identified; no "
        "process was force-ended"
    ) in message
    assert "already running" not in message


def test_no_timeout_test_decides_an_outcome_with_a_sleep():
    """R6-3, pinned: the round-5 proof failed on the reviewer's machine
    because a sleep raced the grace wait. Nothing in this file may go back
    to that."""
    import pathlib
    import re

    source = pathlib.Path(__file__).read_text(encoding="utf-8")
    offenders = [
        line.strip()
        for line in source.splitlines()
        if re.search(r"\btime\.sleep\(", line)
    ]
    assert not offenders, "\n".join(offenders)


@pytestmark_win
def test_a_partial_startup_process_is_not_said_to_be_ended():
    """R5-1(c) and R6-2: _classify told the caller a half-started
    PowerPoint "is ended by this call". Nothing is ended, and the
    replacement claims no current state either."""
    from kitchensink4ppt.com import bridge

    message = str(bridge._classify(_com_error(bridge.CO_E_SERVER_EXEC_FAILURE)))
    assert "is ended by this call" not in message
    assert "was left running" not in message
    assert (
        "a PowerPoint process may have been started, and this call does "
        "not end PowerPoint processes"
    ) in message


@pytestmark_win
def test_no_string_in_the_com_package_claims_a_process_was_ended():
    """R5-1(d): a source guard over the runtime strings, alongside the
    round-4 guard that no code can end a process."""
    import pathlib
    import re

    from kitchensink4ppt.com import bridge

    banned = re.compile(
        r"(was|is|were|are)\s+(aborted|ended|killed|terminated)"
    )
    negated = re.compile(r"\b(no|not|never|nothing|used to)\b", re.I)
    offenders = []
    root = pathlib.Path(bridge.__file__).parent
    for path in sorted(root.glob("*.py")):
        for n, line in enumerate(
            path.read_text(encoding="utf-8").splitlines(), 1
        ):
            stripped = line.strip()
            if stripped.startswith("#"):
                continue  # comments explaining the history are fine
            hit = banned.search(line)
            if hit and not negated.search(line[: hit.start()]):
                offenders.append(f"{path.name}:{n}: {stripped}")
    assert not offenders, "\n".join(offenders)


# ------------------------------------------------------------ R5-2


@metrics
def test_a_negative_marl_is_unmeasurable():
    """R5-2, the reviewer's probe. marL is ST_TextMargin, which is
    NONNEGATIVE; the shared signed bound accepted marL="-1"."""
    sp, body, p = _metrics_shape([("Some text", "Calibri", 1200)])
    p.find(qn("a:pPr")).set("marL", "-1")
    rec = _overflow(sp, body)
    assert rec["method"] == "estimate"
    assert "marL" in rec["method_reason"]
    assert "outside the range DrawingML allows" in rec["method_reason"]


@metrics
def test_a_negative_marl_inherited_through_the_chain_is_unmeasurable():
    """Stated or inherited makes no difference: it is invalid either way."""
    sp, body, p = _metrics_shape([("Some text", "Calibri", 1200)])
    lvl = etree.SubElement(body.find(qn("a:lstStyle")), qn("a:lvl1pPr"))
    lvl.set("marL", "-342900")
    rec = _overflow(sp, body)
    assert rec["method"] == "estimate"
    assert "marL" in rec["method_reason"]


@metrics
def test_a_hanging_indent_stays_measurable():
    """The guard for the other half: indent is ST_TextIndent, signed,
    because a negative indent is how every bulleted list is written."""
    sp, body, p = _metrics_shape([("Some text", "Calibri", 1200)])
    ppr = p.find(qn("a:pPr"))
    ppr.set("marL", "342900")
    ppr.set("indent", "-342900")
    rec = _overflow(sp, body)
    assert rec["method"] == "font-metrics", rec.get("method_reason")


def test_para_margins_rejects_a_negative_marl_directly():
    """The reviewer's probe as written: _para_margins itself must refuse."""
    p = etree.Element(qn("a:p"))
    etree.SubElement(p, qn("a:pPr")).set("marL", "-1")
    with pytest.raises(tx._Unmeasurable):
        tx._para_margins(p)


def test_para_margins_keeps_a_negative_indent():
    p = etree.Element(qn("a:p"))
    ppr = etree.SubElement(p, qn("a:pPr"))
    ppr.set("marL", "342900")
    ppr.set("indent", "-342900")
    assert tx._para_margins(p) == (342900, -342900)
