"""Every operation that takes a `shape` takes the same addresses.

Field report: fit_text {"shape": "a:cccb22"} answered NOT_FOUND "no shape
'a:cccb22'", while the same anchor from get_presentation_view worked in
apply_edits; the caller had to look up the numeric id. The text operations
resolved shapes on their own, by id or exact name only. They now share the
shape operations' resolver (shapes._find_shape), which takes:
- the shape id (a number, or digits as text when no shape has that name);
- the shape's name, searched through groups, a repeated name refused with
  every candidate;
- "Group/Shape" for a member picked through its group's name;
- the view anchor get_presentation_view prints, "a:3f9c2a" or "[a:3f9c2a]".
"""

from __future__ import annotations

import re
import types

import pytest
from lxml import etree

from kitchensink4ppt import server
from kitchensink4ppt.core.errors import (
    AmbiguousTarget,
    PptMcpError,
    TargetNotFound,
)
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import equations as eqn
from kitchensink4ppt.ops import shapes as shp
from kitchensink4ppt.ops import slides as sl
from kitchensink4ppt.ops import text as tx
from kitchensink4ppt.ops import view as vw
from kitchensink4ppt.ops.read import get_slide_info

_LONG = " ".join(["overflowing words"] * 60)


@pytest.fixture()
def deck(make_deck):
    """(pkg, slide) with a long-text box "Body box" and a group "Panel"
    holding a box "Label" on a new Blank slide."""
    pkg = PptxPackage(make_deck("addresses.pptx", extra_slides=0))
    slide = sl.insert_slide(pkg, "Blank")["index"]
    body = tx.insert_textbox(pkg, slide, _LONG, 0.5, 0.5, 4.0, 0.8,
                             name="Body box", size_pt=24)["shape_id"]
    left = tx.insert_textbox(pkg, slide, "Left label", 0.5, 3.0, 2.0, 0.6,
                             name="Label", size_pt=14)["shape_id"]
    right = tx.insert_textbox(pkg, slide, "Right label", 3.0, 3.0, 2.0, 0.6,
                              name="Other", size_pt=14)["shape_id"]
    shp.group_shapes(pkg, slide, [left, right], name="Panel")
    return types.SimpleNamespace(pkg=pkg, slide=slide, body=body, label=left)


def _anchor(pkg, slide, text: str) -> str:
    """The anchor get_presentation_view prints on the line with `text`."""
    view = vw.get_presentation_view(pkg, scope=slide)["view"]
    for line in view.splitlines():
        if text in line:
            found = re.match(r"\[(a:[0-9a-f]+)\]", line.strip())
            if found:
                return found.group(1)
    raise AssertionError(f"no anchored line holds {text!r}:\n{view}")


def _font_sizes(pkg, slide, sid):
    part = get_slide_info(pkg, slide)["part"]
    elem, _chain = shp._find_shape(pkg, part, sid)
    return [r.get("sz") for r in elem.iter(qn("a:rPr"))]


# ------------------------------------------------------------------ fit_text


def test_fit_text_takes_the_view_anchor(deck):
    anchor = _anchor(deck.pkg, deck.slide, "overflowing words")
    out = tx.fit_text(deck.pkg, deck.slide, anchor, min_size=8)
    assert [r["shape_id"] for r in out["fitted"]] == [deck.body]


def test_fit_text_takes_the_anchor_as_the_view_prints_it(deck):
    anchor = _anchor(deck.pkg, deck.slide, "overflowing words")
    out = tx.fit_text(deck.pkg, deck.slide, f"[{anchor}]", min_size=8)
    assert [r["shape_id"] for r in out["fitted"]] == [deck.body]


@pytest.mark.parametrize("form", ["digits", "name"])
def test_fit_text_takes_digits_and_names_like_the_shape_operations(deck, form):
    shape = str(deck.body) if form == "digits" else "Body box"
    out = tx.fit_text(deck.pkg, deck.slide, shape, min_size=8)
    assert [r["shape_id"] for r in out["fitted"]] == [deck.body]


def test_fit_text_reaches_into_a_group_by_path(deck):
    out = tx.fit_text(deck.pkg, deck.slide, "Panel/Label", min_size=8)
    assert out["skipped"][0]["shape_id"] == deck.label


def test_an_anchor_on_another_slide_names_that_slide(deck):
    other = sl.insert_slide(deck.pkg, "Blank")["index"]
    tx.insert_textbox(deck.pkg, other, "Elsewhere entirely", 1, 1, 3, 1)
    anchor = _anchor(deck.pkg, other, "Elsewhere entirely")
    with pytest.raises(TargetNotFound) as caught:
        tx.fit_text(deck.pkg, deck.slide, anchor)
    message = str(caught.value)
    assert f"slide_number {other + 1}" in message
    assert anchor in message


def test_a_stale_anchor_says_to_read_the_view_again(deck):
    with pytest.raises(TargetNotFound, match="get_presentation_view"):
        tx.fit_text(deck.pkg, deck.slide, "a:000000")


@pytest.mark.parametrize("anchor", ["s:256", "t:abc123:r1c1"])
def test_a_slide_or_cell_anchor_is_not_a_shape(deck, anchor):
    with pytest.raises(PptMcpError, match="shape anchor"):
        tx.fit_text(deck.pkg, deck.slide, anchor)


def test_a_repeated_name_is_still_refused_with_its_candidates(deck):
    tx.insert_textbox(deck.pkg, deck.slide, "Twin", 6, 5, 2, 0.5, name="Body box")
    with pytest.raises(AmbiguousTarget, match="Body box"):
        tx.fit_text(deck.pkg, deck.slide, "Body box")


# --------------------------------------------------------- other text ops


def test_format_text_takes_the_anchor(deck):
    anchor = _anchor(deck.pkg, deck.slide, "overflowing words")
    tx.format_text(deck.pkg, deck.slide, anchor, bold=True)
    part = get_slide_info(deck.pkg, deck.slide)["part"]
    elem, _chain = shp._find_shape(deck.pkg, part, deck.body)
    assert all(r.get("b") == "1" for r in elem.iter(qn("a:rPr")))


def test_set_bullets_takes_digits_and_the_anchor(deck):
    tx.set_bullets(deck.pkg, deck.slide, str(deck.body), style="none")
    anchor = _anchor(deck.pkg, deck.slide, "overflowing words")
    tx.set_bullets(deck.pkg, deck.slide, anchor, style="none")


def test_get_autofit_state_takes_the_anchor(deck):
    anchor = _anchor(deck.pkg, deck.slide, "overflowing words")
    state = tx.get_autofit_state(deck.pkg, deck.slide, anchor)
    assert [s["shape_id"] for s in state["shapes"]] == [deck.body]


def test_add_equation_to_shape_takes_the_anchor(deck):
    anchor = _anchor(deck.pkg, deck.slide, "Right label")
    out = eqn.add_equation_to_shape(deck.pkg, deck.slide, anchor, "x^2")
    assert out["shape_id"] != deck.body


# --------------------------------------------------- shape ops and the wire


def test_set_shape_takes_the_anchor_too(deck):
    anchor = _anchor(deck.pkg, deck.slide, "overflowing words")
    shp.set_shape(deck.pkg, deck.slide, anchor, x=1.0)
    part = get_slide_info(deck.pkg, deck.slide)["part"]
    elem, _chain = shp._find_shape(deck.pkg, part, deck.body)
    off = elem.find(f"{qn('p:spPr')}/{qn('a:xfrm')}/{qn('a:off')}")
    assert int(off.get("x")) == 914400


def test_fit_text_over_the_server_takes_the_anchor(make_deck):
    path = make_deck("wire.pptx", extra_slides=0)
    pkg = PptxPackage(path)
    slide = sl.insert_slide(pkg, "Blank")["index"]
    sid = tx.insert_textbox(pkg, slide, _LONG, 0.5, 0.5, 4.0, 0.8,
                            size_pt=24)["shape_id"]
    pkg.save()
    anchor = _anchor(PptxPackage(path), slide, "overflowing words")
    out = server.fit_text(str(path), slide, shape=anchor, min_size=8)
    assert [r["shape_id"] for r in out["changed"]["fitted"]] == [sid]


# ------------------------------------------------------------- live route


class _FakeShape:
    def __init__(self, sid, name):
        self.Id, self.Name, self.Type = sid, name, 1


class _FakeShapes:
    def __init__(self, items):
        self._items = items
        self.Count = len(items)

    def Item(self, i):
        return self._items[i - 1]


class _FakeSlide:
    def __init__(self, slide_id, shapes):
        self.SlideID, self.SlideIndex = slide_id, 1
        self.Shapes = _FakeShapes(shapes)


def test_the_open_deck_route_takes_the_anchor_too():
    """A deck open in PowerPoint is edited through COM; the same anchor the
    file's view printed finds the same shape there (the anchor is the
    slide id and the shape id, which PowerPoint reports as they are in the
    file)."""
    from kitchensink4ppt.com import live_ops

    slide = _FakeSlide(258, [_FakeShape(2, "Title 1"), _FakeShape(6, "Body")])
    anchor = "a:" + vw._digest(258, 6)[:6]
    assert live_ops._find_shape_live(slide, anchor).Id == 6
    assert live_ops._find_shape_live(slide, f"[{anchor}]").Id == 6
    with pytest.raises(TargetNotFound):
        live_ops._find_shape_live(slide, "a:000000")


# ------------------------------------------------------------------- help


@pytest.mark.parametrize("op", [
    "fit_text", "format_text", "set_bullets", "get_autofit_state",
    "add_equation_to_shape", "set_shape", "set_shape_adjustments",
])
def test_help_describes_every_shape_address(op):
    from kitchensink4ppt.registry import public_ops
    from kitchensink4ppt.settings import Settings
    from kitchensink4ppt.surface.help import help_payload
    from kitchensink4ppt.surface.plan import build

    plan = build(Settings(), public_ops())
    shape = help_payload(plan, op)["args_schema"]["properties"]["shape"]
    text = shape.get("description", "")
    assert "a:3f9c2a" in text and "Group/Shape" in text, (op, text)
