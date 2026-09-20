"""Regressions for the 2026-09-20 defense-deck field report.

One test (or more) per punchlist issue in
`Developer Feedback/General Feedback/Beta Test Report - 2026-09-20 -
Final Defense Deck Build.md`:

#869  extract_text paged a 43-slide deck but put the `page` block last,
      behind the whole joined "text", where a client-side truncation ate
      it; and the tool took no offset, so the block's own continuation
      hint named a call that could not be made.
#870  check_layout's tiny_text check skipped every run with no explicit
      sz, which is every ordinary body run on a templated deck.
#871  no contrast audit against the backdrop beneath a shape, and the
      overlap audit saw neither in-group siblings nor placeholders whose
      geometry is inherited.
#872  create_presentation with no template built a 4:3 deck while its
      docstring promised 16:9, and never stated the size it built.
#873  line={"type":"none"} left a themed outline painted, and
      insert_connector dropped end_arrow/start_arrow in silence.
#874  set_shape refused to seed geometry on a layout-inheriting
      placeholder, and vertical anchor had no route that did not replace
      the shape's whole text.
#875  a text replacement dropped explicit run colour, reset paragraph
      alignment and bullets, and wrote only <a:latin> for a font.
#876  delete_table_cols documented a widths value it rejects, and
      get_workflows had no recipe for inheriting a deck and fixing it.
"""

from __future__ import annotations

import json

import pytest
from lxml import etree
from pptx import Presentation
from pptx.util import Inches

from kitchensink4ppt.core import budget
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import read as rd


# --------------------------------------------------------------- #869


def _wordy_deck(path, slides: int = 12):
    """A deck with enough text per slide to overflow a small budget."""
    prs = Presentation()
    blank = prs.slide_layouts[6]
    for i in range(slides):
        slide = prs.slides.add_slide(blank)
        box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.5), Inches(8), Inches(5)
        )
        frame = box.text_frame
        frame.text = f"Slide {i} heading"
        for j in range(6):
            para = frame.add_paragraph()
            para.text = f"Slide {i} paragraph {j}: " + ("lorem ipsum " * 12)
    prs.save(str(path))
    return path


def test_paged_text_leads_with_the_page_block_not_trails_it(
    tmp_path, monkeypatch
):
    """#869: the honesty signal has to survive a truncated payload.

    All the weight of this answer is in "text". A `page` block written
    after it is the first thing any downstream cut removes, which is how
    a 43-slide deck came back as 23 slides that looked complete.
    """
    monkeypatch.setenv(budget.ENV_MAX_CHARS, "6000")
    pkg = PptxPackage(str(_wordy_deck(tmp_path / "wordy.pptx")))
    result = rd.get_text(pkg, include_notes=True)

    keys = list(result)
    assert keys[0] == "page", f"page block is not first: {keys}"
    assert "returned" in keys
    assert keys.index("returned") < keys.index("text")
    assert result["returned"] == len(result["slides"])
    assert result["returned"] < result["slide_count"]
    assert result["page"]["truncated"] is True

    # And it survives a cut that keeps only the first fifth of the wire form.
    wire = json.dumps(result, ensure_ascii=False, default=str)
    assert "next_offset" in wire[: len(wire) // 5]


def test_extract_text_can_act_on_its_own_continuation_hint(
    tmp_path, monkeypatch
):
    """#869: the page block said "call again with offset=N" at a tool
    that had no offset parameter."""
    from kitchensink4ppt import server

    monkeypatch.setenv(budget.ENV_MAX_CHARS, "6000")
    deck = str(_wordy_deck(tmp_path / "wordy2.pptx"))

    seen: list[int] = []
    offset = 0
    for _ in range(50):
        result = server.extract_text(deck, offset=offset)
        seen.extend(s["index"] for s in result["slides"])
        page = result.get("page")
        if page is None or page["next_offset"] is None:
            break
        offset = page["next_offset"]
    assert seen == list(range(12)), seen


def test_a_deck_that_fits_keeps_its_historical_get_text_shape(tmp_path):
    """#869 guard: the reorder is invisible below the ceiling."""
    pkg = PptxPackage(str(_wordy_deck(tmp_path / "small.pptx", slides=2)))
    result = rd.get_text(pkg)
    assert list(result) == ["slide_count", "slides", "text"]


# --------------------------------------------------------------- #873


@pytest.fixture()
def drawable(make_deck):
    """(pkg, slide index) with an empty slide to draw on."""
    from kitchensink4ppt.ops import slides as sl

    pkg = PptxPackage(make_deck("field.pptx", extra_slides=0))
    return pkg, sl.insert_slide(pkg, 0)["index"]


def _ln_of(pkg, slide, shape_id):
    from kitchensink4ppt.core.package import qn
    from kitchensink4ppt.ops import shapes as shp
    from kitchensink4ppt.ops.read import get_slide_info

    part = get_slide_info(pkg, slide)["part"]
    elem, _chain = shp._find_shape(pkg, part, shape_id)
    return elem.find(f"{qn('p:spPr')}/{qn('a:ln')}")


def test_line_type_none_suppresses_the_outline_like_fill_type_none(drawable):
    """#873A: line={"type":"none"} emitted a bare <a:ln/>, which suppresses
    nothing; the p:style lnRef still painted a themed border. Its sibling
    fill={"type":"none"} got this right, so the two disagreed."""
    from kitchensink4ppt.core.package import qn
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    res = shp.insert_shape(
        pkg, slide, "rect", 1, 1, 4, 1,
        fill={"type": "solid", "color": "E0F3FB"},
        line={"type": "none"},
    )
    ln = _ln_of(pkg, slide, res["shape_id"])
    assert ln is not None
    assert ln.find(qn("a:noFill")) is not None, etree.tostring(ln)


def test_connector_writes_the_arrowhead_the_caller_named(drawable):
    """#873B: end_arrow / start_arrow were dropped in silence, and an
    unarrowed line looks like a designed line, so a render review passes."""
    from kitchensink4ppt.core.package import qn
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    a = shp.insert_shape(pkg, slide, "rect", 1, 1, 1, 1)["shape_id"]
    b = shp.insert_shape(pkg, slide, "rect", 5, 3, 1, 1)["shape_id"]
    res = shp.insert_connector(
        pkg, slide, "straight", start_shape=a, end_shape=b,
        line={"color": "00405C", "width": 2.25, "dash": "dash",
              "end_arrow": "triangle"},
    )
    ln = _ln_of(pkg, slide, res["shape_id"])
    assert ln.get("w") == "28575"
    assert ln.find(qn("a:prstDash")).get("val") == "dash"
    tail = ln.find(qn("a:tailEnd"))
    assert tail is not None, etree.tostring(ln)
    assert tail.get("type") == "triangle"


@pytest.mark.parametrize(
    "alias,tag", [
        ("head", "a:headEnd"), ("head_end", "a:headEnd"),
        ("start_arrow", "a:headEnd"), ("tail", "a:tailEnd"),
        ("tail_end", "a:tailEnd"), ("end_arrow", "a:tailEnd"),
    ],
)
def test_every_arrowhead_alias_reaches_the_same_element(alias, tag):
    """#873B: the schema words and the words a caller reaches for both
    land, so neither spelling is a silent no-op."""
    from kitchensink4ppt.core.package import qn
    from kitchensink4ppt.ops import geometry as geo

    ln = geo.line_element({"width": 1, alias: {"type": "stealth", "w": "lg"}})
    el = ln.find(qn(tag))
    assert el is not None, f"{alias} wrote nothing"
    assert el.get("type") == "stealth"
    assert el.get("w") == "lg"


def test_two_aliases_for_one_end_refuse_rather_than_one_winning():
    """#873B: silently picking a winner is the same defect one layer up."""
    from kitchensink4ppt.core.errors import PptMcpError
    from kitchensink4ppt.ops import geometry as geo

    with pytest.raises(PptMcpError) as exc:
        geo.line_element({"tail": "triangle", "end_arrow": "stealth"})
    assert "twice" in str(exc.value)


@pytest.mark.parametrize(
    "spec,what", [
        ({"width": 2, "endarrow": "triangle"}, "line"),
        ({"width": 2, "colour": "FF0000"}, "line"),
    ],
)
def test_unknown_line_keys_refuse_instead_of_dropping(spec, what):
    """#873: silent key-dropping is the shared root cause of both halves."""
    from kitchensink4ppt.core.errors import PptMcpError
    from kitchensink4ppt.ops import geometry as geo

    with pytest.raises(PptMcpError) as exc:
        geo.line_element(spec)
    assert "unknown" in str(exc.value).lower()


def test_unknown_fill_keys_refuse_instead_of_dropping():
    """#873: the same contract on the sibling parameter."""
    from kitchensink4ppt.core.errors import PptMcpError
    from kitchensink4ppt.ops import geometry as geo

    with pytest.raises(PptMcpError):
        geo.fill_element({"type": "solid", "color": "FF0000", "opacity": 0.5})
    # the spelled-right keys still work
    assert geo.fill_element({"type": "solid", "color": "FF0000",
                             "alpha": 0.5}) is not None


# --------------------------------------------------------------- #875


def _styled_body(pkg, slide, shape_id):
    """Give a shape a body worth preserving: left-aligned, bullets off,
    dark-navy runs in Arial across three paragraphs."""
    from kitchensink4ppt.ops import text as txt

    txt.set_placeholder_text  # noqa: B018  (import guard, see set_text below)
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops.read import get_slide_info
    from kitchensink4ppt.ops import shapes as shp

    part = get_slide_info(pkg, slide)["part"]
    elem, _chain = shp._find_shape(pkg, part, shape_id)
    body = elem.find(_qn("p:txBody"))
    for child in list(body.findall(_qn("a:p"))):
        body.remove(child)
    bodypr = body.find(_qn("a:bodyPr"))
    bodypr.set("anchor", "t")
    for line in ("First", "Second", "Third"):
        p = etree.SubElement(body, _qn("a:p"))
        ppr = etree.SubElement(p, _qn("a:pPr"))
        ppr.set("algn", "l")
        ppr.set("lvl", "1")
        etree.SubElement(ppr, _qn("a:buNone"))
        r = etree.SubElement(p, _qn("a:r"))
        rpr = etree.SubElement(r, _qn("a:rPr"))
        rpr.set("lang", "en-US")
        rpr.set("sz", "1800")
        fill = etree.SubElement(rpr, _qn("a:solidFill"))
        clr = etree.SubElement(fill, _qn("a:srgbClr"))
        clr.set("val", "1F3864")
        for tag in ("a:latin", "a:ea", "a:cs"):
            el = etree.SubElement(rpr, _qn(tag))
            el.set("typeface", "Arial")
        t = etree.SubElement(r, _qn("a:t"))
        t.text = line
    return body


def _first_rpr(pkg, slide, shape_id, para: int = 0):
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops.read import get_slide_info
    from kitchensink4ppt.ops import shapes as shp

    part = get_slide_info(pkg, slide)["part"]
    elem, _chain = shp._find_shape(pkg, part, shape_id)
    body = elem.find(_qn("p:txBody"))
    p = body.findall(_qn("a:p"))[para]
    return p.find(f"{_qn('a:r')}/{_qn('a:rPr')}"), p.find(_qn("a:pPr")), body


def test_retext_keeps_run_colour_alignment_bullets_and_anchor(drawable):
    """#875 items 1 and 2: a plain text replacement re-centred paragraphs,
    put bullets back where set_bullets(style="none") had removed them,
    reset the vertical anchor, and dropped explicit run colour, all with
    nothing but changed: ["text"] to show for it."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    sid = shp.insert_shape(pkg, slide, "rect", 1, 1, 6, 3, text="x")["shape_id"]
    _styled_body(pkg, slide, sid)

    res = shp.set_shape(pkg, slide, sid, text="One\nTwo\nThree")
    rpr, ppr, body = _first_rpr(pkg, slide, sid)

    assert ppr.get("algn") == "l", "paragraph was re-aligned"
    assert ppr.get("lvl") == "1", "outline level was dropped"
    assert ppr.find(_qn("a:buNone")) is not None, "bullets came back"
    assert body.find(_qn("a:bodyPr")).get("anchor") == "t"
    clr = rpr.find(f"{_qn('a:solidFill')}/{_qn('a:srgbClr')}")
    assert clr is not None and clr.get("val") == "1F3864"
    assert rpr.get("sz") == "1800"
    assert rpr.find(_qn("a:latin")).get("typeface") == "Arial"
    # and the result says what it carried rather than staying silent
    assert set(res["preserved"]) >= {"alignment", "bullets", "run_color"}


def test_retext_still_applies_what_the_caller_did_name(drawable):
    """#875 guard: preservation is not stickiness. A named key wins."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    sid = shp.insert_shape(pkg, slide, "rect", 1, 1, 6, 3, text="x")["shape_id"]
    _styled_body(pkg, slide, sid)

    shp.set_shape(
        pkg, slide, sid, text="One",
        text_style={"align": "center", "color": "C00000", "size": 32,
                    "anchor": "middle"},
    )
    rpr, ppr, body = _first_rpr(pkg, slide, sid)
    assert ppr.get("algn") == "ctr"
    assert body.find(_qn("a:bodyPr")).get("anchor") == "ctr"
    assert rpr.get("sz") == "3200"
    clr = rpr.find(f"{_qn('a:solidFill')}/{_qn('a:srgbClr')}")
    assert clr.get("val") == "C00000"
    # untouched properties still ride through
    assert ppr.find(_qn("a:buNone")) is not None
    assert rpr.find(_qn("a:latin")).get("typeface") == "Arial"


def test_apply_edits_set_text_preserves_the_same_properties(drawable, tmp_path):
    """#875 item 2 reached the deck through apply_edits' set_text op, which
    is the route that put bullet glyphs back on a defense deck's closing
    slide after they had been explicitly removed."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import batch, shapes as shp

    pkg, slide = drawable
    sid = shp.insert_shape(pkg, slide, "rect", 1, 1, 6, 3, text="x")["shape_id"]
    _styled_body(pkg, slide, sid)

    batch.apply_edits(pkg, [{
        "op": "set_text", "slide": slide, "shape": sid,
        "text": "Thank you",
    }])
    rpr, ppr, _body = _first_rpr(pkg, slide, sid)
    assert ppr.get("algn") == "l"
    assert ppr.find(_qn("a:buNone")) is not None
    assert rpr.find(f"{_qn('a:solidFill')}/{_qn('a:srgbClr')}").get(
        "val") == "1F3864"


def test_text_style_font_writes_ea_and_cs_not_only_latin(drawable):
    """#875 item 3: a latin-only write hands CJK and complex-script runs to
    the theme's minor font while reporting the typeface as set."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    sid = shp.insert_shape(
        pkg, slide, "rect", 1, 1, 4, 1, text="Delta",
        text_style={"font": "Georgia", "size": 18},
    )["shape_id"]
    rpr, _ppr, _body = _first_rpr(pkg, slide, sid)
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rpr.find(_qn(tag))
        assert el is not None, f"{tag} was not written"
        assert el.get("typeface") == "Georgia"
