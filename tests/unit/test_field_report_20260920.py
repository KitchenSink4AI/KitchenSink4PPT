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
        pkg, slide, "rect", 1, 1, 4, 1, text="Sample",
        text_style={"font": "Georgia", "size": 18},
    )["shape_id"]
    rpr, _ppr, _body = _first_rpr(pkg, slide, sid)
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rpr.find(_qn(tag))
        assert el is not None, f"{tag} was not written"
        assert el.get("typeface") == "Georgia"


# --------------------------------------------------------------- #874


@pytest.fixture()
def inheriting_placeholder(make_deck):
    """(pkg, slide, shape_id) for a title placeholder that carries an empty
    p:spPr and takes its whole box from the layout."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import slides as sl
    from kitchensink4ppt.ops.read import get_slide_info
    from kitchensink4ppt.ops import shapes as shp

    pkg = PptxPackage(make_deck("inherit.pptx", extra_slides=0))
    slide = sl.insert_slide(pkg, 0)["index"]
    part = get_slide_info(pkg, slide)["part"]
    tree = pkg.root(part).find(f"{_qn('p:cSld')}/{_qn('p:spTree')}")
    for sp in tree.iter(_qn("p:sp")):
        nvpr = sp.find(f"{_qn('p:nvSpPr')}/{_qn('p:nvPr')}")
        if nvpr is None or nvpr.find(_qn("p:ph")) is None:
            continue
        sppr = sp.find(_qn("p:spPr"))
        for child in list(sppr):
            sppr.remove(child)  # back to a pure inheriting placeholder
        elem, _chain = shp._find_shape(pkg, part, shp._shape_id(sp))
        return pkg, slide, shp._shape_id(sp)
    pytest.skip("layout 0 carries no placeholder to test with")


def test_set_shape_seeds_geometry_on_an_inheriting_placeholder(
    inheriting_placeholder
):
    """#874A: the refusal was circular. It said to set an absolute position
    and size first, and set_shape is the tool that does that, so a full
    x/y/w/h call could never get past it."""
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide, sid = inheriting_placeholder
    res = shp.set_shape(pkg, slide, sid, x=1.167, y=0.95, w=11.0, h=2.10)

    assert "geometry" in res["changed"]
    from kitchensink4ppt.ops.read import get_slide_info

    rec = next(
        s for s in get_slide_info(pkg, slide)["shapes"] if s["id"] == sid
    )
    geo = rec["geometry"]
    assert geo is not None, "the placeholder still has no explicit box"
    assert round(geo["x"] / 914400, 3) == 1.167
    assert round(geo["cx"] / 914400, 2) == 11.0
    assert round(geo["cy"] / 914400, 2) == 2.10
    assert any("seeded" in w for w in res.get("warnings", []))


def test_seeding_a_partial_move_keeps_the_inherited_dimensions(
    inheriting_placeholder
):
    """#874A: a partial call seeds the rest from the layout box rather than
    inventing zeros."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import inherit as inh
    from kitchensink4ppt.ops.read import get_slide_info
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide, sid = inheriting_placeholder
    part = get_slide_info(pkg, slide)["part"]
    elem, _chain = shp._find_shape(pkg, part, sid)
    inherited = inh.inherited_box(pkg, part, elem)
    assert inherited is not None, "no layout box to inherit (bad fixture)"

    shp.set_shape(pkg, slide, sid, y=4.0)
    rec = next(
        s for s in get_slide_info(pkg, slide)["shapes"] if s["id"] == sid
    )
    assert rec["geometry"]["x"] == inherited[0]
    assert rec["geometry"]["cx"] == inherited[2]
    assert rec["geometry"]["cy"] == inherited[3]
    assert round(rec["geometry"]["y"] / 914400, 2) == 4.0


def test_vertical_anchor_changes_without_replacing_the_text(drawable):
    """#874B: the only route to bodyPr anchor was set_shape's text_style,
    which replaces the whole body with one flat string and so flattens
    bullet levels, per-paragraph alignment and per-run formatting."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import shapes as shp
    from kitchensink4ppt.ops import text as txt

    pkg, slide = drawable
    sid = shp.insert_shape(pkg, slide, "rect", 1, 1, 6, 3, text="x")["shape_id"]
    _styled_body(pkg, slide, sid)

    res = txt.format_text(pkg, slide, sid, anchor="top")
    rpr, ppr, body = _first_rpr(pkg, slide, sid)
    assert body.find(_qn("a:bodyPr")).get("anchor") == "t"
    assert res["frame_changed"] == ["anchor"]
    # the rich text is untouched: three paragraphs, bullets off, navy runs
    assert len(body.findall(_qn("a:p"))) == 3
    assert ppr.find(_qn("a:buNone")) is not None
    assert rpr.find(f"{_qn('a:solidFill')}/{_qn('a:srgbClr')}").get(
        "val") == "1F3864"


def test_format_text_refuses_an_unknown_anchor(drawable):
    """#874B: a typo must refuse, not write a broken attribute."""
    from kitchensink4ppt.core.errors import PptMcpError
    from kitchensink4ppt.ops import shapes as shp
    from kitchensink4ppt.ops import text as txt

    pkg, slide = drawable
    sid = shp.insert_shape(pkg, slide, "rect", 1, 1, 6, 3, text="hi")["shape_id"]
    with pytest.raises(PptMcpError):
        txt.format_text(pkg, slide, sid, anchor="up")


def test_apply_edits_reaches_the_anchor_as_text_anchor(drawable):
    """#874B: a bare "anchor" key in an edit already names the view anchor
    that ADDRESSES the shape, so the property rides as text_anchor."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import batch, shapes as shp

    pkg, slide = drawable
    sid = shp.insert_shape(pkg, slide, "rect", 1, 1, 6, 3, text="a\nb")["shape_id"]
    batch.apply_edits(pkg, [{
        "op": "format_text", "slide": slide, "shape": sid,
        "text_anchor": "bottom",
    }])
    _rpr, _ppr, body = _first_rpr(pkg, slide, sid)
    assert body.find(_qn("a:bodyPr")).get("anchor") == "b"


# --------------------------------------------------------------- #870


def _master_body_size(pkg, slide_part, pt: float) -> None:
    """Pin the master's bodyStyle level-1 size, the size an ordinary body
    run on a templated deck actually renders at."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import inherit as inh

    layout = inh.layout_part_of(pkg, slide_part)
    master = inh.master_part_of(pkg, layout)
    styles = pkg.root(master).find(_qn("p:txStyles"))
    body = styles.find(_qn("p:bodyStyle"))
    lvl = body.find(_qn("a:lvl1pPr"))
    defrpr = lvl.find(_qn("a:defRPr"))
    if defrpr is None:
        defrpr = etree.SubElement(lvl, _qn("a:defRPr"))
    defrpr.set("sz", str(int(pt * 100)))
    pkg.mark_dirty(master)


def test_tiny_text_resolves_a_size_inherited_from_the_master(
    inheriting_placeholder
):
    """#870: the check read a:rPr @sz and skipped every run without one,
    which is every ordinary body run on a templated deck. Those runs are
    exactly the ones rendering at the master's size."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import design_check as dc
    from kitchensink4ppt.ops.read import get_slide_info
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide, _sid = inheriting_placeholder
    part = get_slide_info(pkg, slide)["part"]
    _master_body_size(pkg, part, 8.0)

    # Give the body placeholder text carrying NO explicit size.
    tree = pkg.root(part).find(f"{_qn('p:cSld')}/{_qn('p:spTree')}")
    target = None
    for sp in tree.iter(_qn("p:sp")):
        ph = sp.find(f"{_qn('p:nvSpPr')}/{_qn('p:nvPr')}/{_qn('p:ph')}")
        if ph is None or (ph.get("type") or "body") in ("title", "ctrTitle"):
            continue
        body = sp.find(_qn("p:txBody"))
        for p in list(body.findall(_qn("a:p"))):
            body.remove(p)
        p = etree.SubElement(body, _qn("a:p"))
        r = etree.SubElement(p, _qn("a:r"))
        etree.SubElement(r, _qn("a:rPr")).set("lang", "en-US")
        etree.SubElement(r, _qn("a:t")).text = (
            "A body line long enough not to read as a short label."
        )
        target = shp._shape_id(sp)
        break
    if target is None:
        pytest.skip("the generated layout has no body placeholder")

    findings = dc.check_layout(pkg, slide=slide, checks=["tiny_text"])[
        "findings"]
    hit = [f for f in findings if target in f["shape_ids"]]
    assert hit, f"8pt inherited body text was not flagged: {findings}"
    assert hit[0]["inherited_size"] is True
    assert 8.0 in hit[0]["sizes_pt"]
    assert "master" in " ".join(hit[0]["size_sources"])


def test_an_explicit_size_still_reports_as_its_own(drawable):
    """#870 guard: resolution does not relabel what the slide states."""
    from kitchensink4ppt.ops import design_check as dc
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    sid = shp.insert_shape(
        pkg, slide, "rect", 1, 1, 6, 2,
        text="A body line long enough not to read as a short label.",
        text_style={"size": 6},
    )["shape_id"]
    findings = dc.check_layout(pkg, slide=slide, checks=["tiny_text"])[
        "findings"]
    hit = [f for f in findings if sid in f["shape_ids"]]
    assert hit
    assert hit[0]["inherited_size"] is False
    assert hit[0]["size_sources"] == ["run"]


def test_overflow_can_estimate_a_placeholder_that_inherits_its_box(
    inheriting_placeholder
):
    """#870, overflow-adjacent: the estimate dead-ended twice on the same
    shape, once for the missing xfrm and once on a flat 18pt guess. The
    v1 build got zero findings on slides whose body ran off the bottom."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import design_check as dc
    from kitchensink4ppt.ops.read import get_slide_info
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide, _sid = inheriting_placeholder
    part = get_slide_info(pkg, slide)["part"]
    _master_body_size(pkg, part, 28.0)

    tree = pkg.root(part).find(f"{_qn('p:cSld')}/{_qn('p:spTree')}")
    target = None
    for sp in tree.iter(_qn("p:sp")):
        ph = sp.find(f"{_qn('p:nvSpPr')}/{_qn('p:nvPr')}/{_qn('p:ph')}")
        if ph is None or (ph.get("type") or "body") in ("title", "ctrTitle"):
            continue
        body = sp.find(_qn("p:txBody"))
        for p in list(body.findall(_qn("a:p"))):
            body.remove(p)
        for i in range(24):
            p = etree.SubElement(body, _qn("a:p"))
            r = etree.SubElement(p, _qn("a:r"))
            etree.SubElement(r, _qn("a:rPr")).set("lang", "en-US")
            etree.SubElement(r, _qn("a:t")).text = (
                f"Line {i}: far more body copy than this frame can hold, "
                "set at the master's twenty-eight point body size."
            )
        target = shp._shape_id(sp)
        break
    if target is None:
        pytest.skip("the generated layout has no body placeholder")

    findings = dc.check_layout(pkg, slide=slide, checks=["overflow"])[
        "findings"]
    assert [f for f in findings if target in f["shape_ids"]], findings


# --------------------------------------------------------------- #872


def test_a_deck_created_with_no_template_is_actually_16_9(tmp_path):
    """#872: the docstring promised 16:9 and the file that landed was the
    stock python-pptx 4:3 default. This is the first call in any
    from-scratch build, so every coordinate after it was computed against
    the wrong canvas."""
    from kitchensink4ppt.ops import slides as sl

    out = tmp_path / "fresh.pptx"
    res = sl.create_presentation(out)

    assert res["slide_size"]["cx"] == 12192000
    assert res["slide_size"]["cy"] == 6858000
    assert res["slide_size"]["w_in"] == 13.333
    assert res["slide_size"]["type"] is None

    pkg = PptxPackage(out)
    from kitchensink4ppt.core.package import qn as _qn

    sldsz = pkg.presentation().find(_qn("p:sldSz"))
    assert sldsz.get("cx") == "12192000"
    assert sldsz.get("type") is None


def test_widening_moves_the_layouts_with_the_canvas(tmp_path):
    """#872: a canvas stretch alone would leave the whole design sitting
    in the left ten inches. The two canvases share a height, so the
    conversion is an exact horizontal scale."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import slides as sl

    out = tmp_path / "fresh2.pptx"
    sl.create_presentation(out)
    pkg = PptxPackage(out)

    widest = 0
    tallest = 0
    for part, _name in sl._layouts(pkg):
        for ext in pkg.root(part).iter(_qn("a:ext")):
            widest = max(widest, int(ext.get("cx")))
            tallest = max(tallest, int(ext.get("cy")))
    assert widest > 9144000, "layout geometry never left the 4:3 canvas"
    assert widest <= 12192000, "layout geometry overshot the canvas"
    assert tallest <= 6858000, "vertical geometry was scaled; it must not be"


def test_a_templated_deck_keeps_its_own_canvas(tmp_path):
    """#872 guard: only the no-template path changes size."""
    from pptx import Presentation as _Prs
    from kitchensink4ppt.ops import slides as sl

    template = tmp_path / "t.pptx"
    prs = _Prs()
    prs.slide_width = 9144000
    prs.slide_height = 6858000
    prs.save(str(template))

    res = sl.create_presentation(tmp_path / "from_t.pptx", template=template)
    assert res["slide_size"]["cx"] == 9144000


# --------------------------------------------------------------- #876


@pytest.fixture()
def table_deck(make_deck):
    """(pkg, slide, table selector) for a five-column table."""
    from kitchensink4ppt.ops import slides as sl
    from kitchensink4ppt.ops import tables as tb

    pkg = PptxPackage(make_deck("tables.pptx", extra_slides=0))
    slide = sl.insert_slide(pkg, 0)["index"]
    res = tb.create_table(pkg, slide, rows=3, cols=5, x=1, y=1, w=8, h=2)
    return pkg, slide, {"shape_id": res["shape_id"]}


def test_delete_table_cols_accepts_the_widths_value_it_documents(table_deck):
    """#876: the docstring said widths='fit' and the code refused it, so a
    caller following the documentation got BAD_PARAMS."""
    from kitchensink4ppt.ops import tables as tb

    pkg, slide, sel = table_deck
    before = tb.get_table(pkg, slide, sel)
    total = sum(before["column_widths_in"])

    res = tb.delete_table_cols(pkg, slide, sel, 1, 1, widths="fit")
    assert res["widths"] == "fit"
    after = tb.get_table(pkg, slide, sel)
    assert len(after["column_widths_in"]) == 4
    assert sum(after["column_widths_in"]) == pytest.approx(total, abs=0.02)


def test_rescale_still_works_as_an_alias(table_deck):
    """#876: 'rescale' is the value that shipped; renaming it away would
    break the callers who read the code instead of the docstring."""
    from kitchensink4ppt.ops import tables as tb

    pkg, slide, sel = table_deck
    res = tb.delete_table_cols(pkg, slide, sel, 1, 1, widths="rescale")
    assert res["widths"] == "fit"


def test_both_column_tools_take_the_same_vocabulary(table_deck):
    """#876: the two structural column tools used different words for the
    same concept, and one documented the other's."""
    from kitchensink4ppt.core.errors import PptMcpError
    from kitchensink4ppt.ops import tables as tb

    pkg, slide, sel = table_deck
    assert tb.insert_table_cols(
        pkg, slide, sel, 1, 1, widths="fit")["widths"] == "fit"
    assert tb.delete_table_cols(
        pkg, slide, sel, 1, 1, widths="fit")["widths"] == "fit"
    for call in (tb.insert_table_cols, tb.delete_table_cols):
        with pytest.raises(PptMcpError) as exc:
            call(pkg, slide, sel, 1, 1, widths="stretch")
        assert "fit" in str(exc.value) and "shift" in str(exc.value)


def test_there_is_a_recipe_for_inheriting_a_deck_and_fixing_it():
    """#876: every recipe assumed building new, and the most common real
    job is repairing a deck that already exists."""
    from kitchensink4ppt.ops import workflows as wf

    recipe = wf.get_workflows(task="refresh-an-existing-deck")
    tools = [s["tool"] for s in recipe["steps"]]
    for expected in (
        "copy_presentation", "get_presentation_view", "check_layout",
        "export_slide_image", "apply_edits", "validate",
    ):
        assert expected in tools, f"{expected} missing from {tools}"


def test_render_and_review_names_the_packs_its_fixes_need():
    """#876: the loop was listed as needing assembly-export and design, and
    fixing what you saw needs the graphics and table tools."""
    from kitchensink4ppt.ops import workflows as wf

    packs = wf.get_workflows(task="render-and-review")["packs"]
    assert "graphics" in packs
    assert "tables-charts" in packs


# --------------------------------------------------------------- #871


def _contrast(pkg, slide, **opts):
    from kitchensink4ppt.ops import design_check as dc

    checks = [{"check": "contrast", **opts}] if opts else ["contrast"]
    return dc.check_layout(pkg, slide=slide, checks=checks)["findings"]


def test_contrast_sees_the_band_a_label_is_floating_on(drawable):
    """#871 defect 1: white labels on a pale band measured 1.17:1 and the
    check never saw them, because it compared each run against its OWN
    shape's fill and the labels are unfilled text boxes sitting on the
    band beneath them."""
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    # A pale band, then an unfilled label on top of it in white. The label
    # OVERHANGS the band's bottom edge, which is how real ones sit: the
    # old rule needed the band to contain the whole label box.
    shp.insert_shape(pkg, slide, "rect", 0.5, 2.0, 10.0, 1.2,
                     fill="F7EAE7", line={"type": "none"},
                     name="lane A band")
    label = shp.insert_shape(
        pkg, slide, "rect", 0.8, 2.9, 1.2, 0.5,
        fill={"type": "none"}, line={"type": "none"}, text="AA",
        text_style={"size": 20, "color": "FFFFFF"},
    )["shape_id"]

    hits = [f for f in _contrast(pkg, slide) if label in f["shape_ids"]]
    assert hits, "white on a pale band was not flagged"
    assert hits[0]["severity"] == "error"
    assert hits[0]["ratio"] < 2.0
    # The band, not the white slide the old rule fell through to.
    assert hits[0]["fill_color"] == "F7EAE7"


def test_dark_on_dark_over_an_overhanging_band_is_not_passed(drawable):
    """#871: the sharper half of the same miss. Falling through to the
    white slide background made dark text on a dark band look like good
    contrast, so the check reported nothing at all."""
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    shp.insert_shape(pkg, slide, "rect", 0.5, 2.0, 10.0, 1.2,
                     fill="1F3864", line={"type": "none"}, name="lane band")
    label = shp.insert_shape(
        pkg, slide, "rect", 0.8, 2.9, 1.6, 0.5,
        fill={"type": "none"}, line={"type": "none"}, text="BB",
        text_style={"size": 20, "color": "203A60"},
    )["shape_id"]

    hits = [f for f in _contrast(pkg, slide) if label in f["shape_ids"]]
    assert hits, "navy on navy was passed as if it sat on the white slide"
    assert hits[0]["fill_color"] == "1F3864"


def test_a_translucent_band_composites_before_it_is_judged(drawable):
    """#871: the band in the field case was accent1 at 12% alpha, and the
    colour the eye sees is the composite, not the full-strength accent."""
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    shp.insert_shape(
        pkg, slide, "rect", 0.5, 2.0, 10.0, 1.2, name="wash",
        fill={"type": "solid", "color": "C00000", "alpha": 0.12},
        line={"type": "none"},
    )
    label = shp.insert_shape(
        pkg, slide, "rect", 0.8, 2.3, 1.2, 0.5,
        fill={"type": "none"}, line={"type": "none"}, text="AA",
        text_style={"size": 20, "color": "FFFFFF"},
    )["shape_id"]
    hits = [f for f in _contrast(pkg, slide) if label in f["shape_ids"]]
    assert hits, "white over a 12% wash on white was not flagged"
    # Composited, not judged against full-strength C00000 (which white
    # would pass against).
    assert hits[0]["fill_color"] != "C00000"


def test_no_fill_in_spPr_beats_the_style_fillRef(drawable):
    """#871: the opposite half of the same gap produced seven false
    positives, judging text against an orange invented from p:style when
    the shape carries a:noFill and paints nothing."""
    from kitchensink4ppt.ops import design_check as dc
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    sid = shp.insert_shape(
        pkg, slide, "rect", 1, 1, 3, 1, fill={"type": "none"},
        line={"type": "none"}, text="milestone",
        text_style={"size": 12, "color": "1F3864"},
    )["shape_id"]
    part = __import__(
        "kitchensink4ppt.ops.read", fromlist=["get_slide_info"]
    ).get_slide_info(pkg, slide)["part"]
    elem, _chain = shp._find_shape(pkg, part, sid)
    ctx = dc._SlideCtx(pkg, {"part": part, "index": slide, "slide_id": 0})
    srec = next(s for s in ctx.all if s["id"] == sid)
    hexval, skip = dc._own_fill_hex(elem, ctx.resolver())
    assert hexval == dc._TRANSPARENT and skip is None
    # And the backdrop it falls through to is the white slide, not accent1.
    backdrop, reason = dc._backdrop_hex(srec, ctx, ctx.resolver())
    assert reason is None
    assert backdrop == "FFFFFF"
    assert not [f for f in _contrast(pkg, slide) if sid in f["shape_ids"]]


def test_an_unresolvable_backdrop_is_reported_not_passed(drawable, tmp_path):
    """#871: a picture under the text means the check cannot know, and
    'cannot know' is a result, not a pass."""
    from PIL import Image

    from kitchensink4ppt.ops import media as md
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    img = tmp_path / "bg.png"
    Image.new("RGB", (64, 64), (200, 200, 200)).save(img)
    md.insert_image(pkg, slide, str(img), 0.5, 0.5, 6, 4)
    label = shp.insert_shape(
        pkg, slide, "rect", 1.0, 1.0, 2.0, 0.6,
        fill={"type": "none"}, line={"type": "none"}, text="over a photo",
        text_style={"size": 14, "color": "FFFFFF"},
    )["shape_id"]

    hits = [f for f in _contrast(pkg, slide) if label in f["shape_ids"]]
    assert hits, "text over a picture was passed in silence"
    assert hits[0]["severity"] == "info"
    assert "unresolvable_backdrop" in hits[0]


def test_an_inherited_run_colour_is_judged_and_its_source_named(
    inheriting_placeholder
):
    """#871: the old caveat exempted every ordinary bulleted body slide
    from the contrast check, since those runs carry no colour of their
    own. That is exactly why they are worth checking."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops.read import get_slide_info
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide, _sid = inheriting_placeholder
    part = get_slide_info(pkg, slide)["part"]

    tree = pkg.root(part).find(f"{_qn('p:cSld')}/{_qn('p:spTree')}")
    target = None
    for sp in tree.iter(_qn("p:sp")):
        ph = sp.find(f"{_qn('p:nvSpPr')}/{_qn('p:nvPr')}/{_qn('p:ph')}")
        if ph is None or (ph.get("type") or "body") in ("title", "ctrTitle"):
            continue
        body = sp.find(_qn("p:txBody"))
        for p in list(body.findall(_qn("a:p"))):
            body.remove(p)
        p = etree.SubElement(body, _qn("a:p"))
        r = etree.SubElement(p, _qn("a:r"))
        etree.SubElement(r, _qn("a:rPr")).set("lang", "en-US")
        etree.SubElement(r, _qn("a:t")).text = "White on white, inherited."
        target, target_elem = shp._shape_id(sp), sp
        break
    if target is None:
        pytest.skip("the generated layout has no body placeholder")

    # Paint the LAYOUT placeholder's text white; the slide background is
    # white and the run itself says nothing about colour.
    from kitchensink4ppt.ops import inherit as inh

    layout = inh.layout_part_of(pkg, part)
    key = inh.placeholder_key(inh.placeholder_of(target_elem))
    twin = inh.layout_twin(pkg, layout, key)
    assert twin is not None
    lst = twin.find(f"{_qn('p:txBody')}/{_qn('a:lstStyle')}")
    if lst is None:
        body = twin.find(_qn("p:txBody"))
        lst = etree.Element(_qn("a:lstStyle"))
        body.insert(1, lst)
    lvl = lst.find(_qn("a:lvl1pPr"))
    if lvl is None:
        lvl = etree.SubElement(lst, _qn("a:lvl1pPr"))
    defrpr = lvl.find(_qn("a:defRPr"))
    if defrpr is None:
        defrpr = etree.SubElement(lvl, _qn("a:defRPr"))
    for child in list(defrpr):
        if etree.QName(child).localname.endswith("Fill"):
            defrpr.remove(child)
    fill = etree.Element(_qn("a:solidFill"))
    etree.SubElement(fill, _qn("a:srgbClr")).set("val", "FFFFFF")
    defrpr.insert(0, fill)
    pkg.mark_dirty(layout)

    hits = [f for f in _contrast(pkg, slide) if target in f["shape_ids"]]
    assert hits, "an inherited white-on-white run was not judged"
    assert hits[0]["color_source"] == "layout"
    assert hits[0]["text_color"] == "FFFFFF"


def test_overlap_recurses_into_groups(drawable):
    """#871 defect 2: every real collision in the field deck's timeline
    was between two boxes inside a group, and the check compared only
    top-level shapes. A deck built with this server's own diagram tooling
    is grouped, so it was structurally invisible."""
    from kitchensink4ppt.ops import design_check as dc
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    a = shp.insert_shape(pkg, slide, "rect", 1.0, 1.0, 2.0, 1.0,
                         text="2001", name="Milestone A")["shape_id"]
    b = shp.insert_shape(pkg, slide, "rect", 2.925, 1.0, 2.0, 1.0,
                         text="2002", name="Milestone B")["shape_id"]
    grp = shp.group_shapes(pkg, slide, [a, b], name="Timeline")["group_id"]

    hits = [
        f for f in dc.check_layout(pkg, slide=slide, checks=["overlap"])[
            "findings"]
        if f.get("group_id") == grp
    ]
    assert hits, "a 0.075in corner clip inside a group was not flagged"
    assert set(hits[0]["shape_ids"]) == {a, b}
    assert hits[0]["overlap_in"][0] == pytest.approx(0.075, abs=0.01)


def test_families_meant_to_touch_are_left_out_of_the_group_pass(drawable):
    """#871: the exclusion list is what keeps the in-group pass usable as
    a gate. A timeline's tick sits ON its label by design."""
    from kitchensink4ppt.ops import design_check as dc
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    label = shp.insert_shape(pkg, slide, "rect", 1.0, 1.0, 2.0, 1.0,
                             text="2001", name="Milestone A")["shape_id"]
    tick = shp.insert_shape(pkg, slide, "rect", 1.4, 1.4, 0.2, 0.2,
                            name="tick 1")["shape_id"]
    grp = shp.group_shapes(pkg, slide, [label, tick], name="Timeline")[
        "group_id"]

    hits = [
        f for f in dc.check_layout(pkg, slide=slide, checks=["overlap"])[
            "findings"]
        if f.get("group_id") == grp
    ]
    assert hits == [], f"the tick was flagged against its label: {hits}"


# ===================================================================
# ROUND 2: every finding from the adversarial verification of PR #29.
# MAJOR-A and MAJOR-B are REGRESSIONS round 1 introduced; main did not
# have them. The rest are defects round 1 left standing or created.
# ===================================================================


def _mk(pkg, slide, **kw):
    from kitchensink4ppt.ops import shapes as shp
    return shp.insert_shape(pkg, slide, "rect", 1, 1, 4, 1, **kw)["shape_id"]


def _rpr(pkg, slide, sid, para=0):
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops.read import get_slide_info
    from kitchensink4ppt.ops import shapes as shp
    part = get_slide_info(pkg, slide)["part"]
    elem, _c = shp._find_shape(pkg, part, sid)
    body = elem.find(_qn("p:txBody"))
    p0 = body.findall(_qn("a:p"))[para]
    return p0.find(f"{_qn('a:r')}/{_qn('a:rPr')}"), body


# ------------------------------------------------------------ MAJOR-A


@pytest.mark.parametrize("key,attr", [("bold", "b"), ("italic", "i")])
def test_an_explicit_false_toggle_turns_the_attribute_off(drawable, key, attr):
    """MAJOR-A: _apply_rpr wrote b/i only when the value was TRUTHY, so a
    False produced no attribute and the carry kept the old run's b="1".
    The named key lost, which is the exact opposite of the contract round
    1 claimed. main did not have this: it rebuilt the body from scratch,
    so False at least meant absent."""
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    sid = _mk(pkg, slide, text="X", text_style={"bold": True, "italic": True})
    before, _b = _rpr(pkg, slide, sid)
    assert before.get(attr) == "1"

    shp.set_shape(pkg, slide, sid, text="Y", text_style={key: False})
    after, _b = _rpr(pkg, slide, sid)
    assert after.get(attr) == "0", (
        f"text_style {key}=False left {attr}={after.get(attr)!r}"
    )


def test_an_explicit_false_toggle_also_turns_off_an_inherited_one(drawable):
    """MAJOR-A: writing the explicit off attribute is what defeats a bold
    inherited from the placeholder, which absence never could."""
    pkg, slide = drawable
    sid = _mk(pkg, slide, text="X", text_style={"bold": False})
    rpr, _b = _rpr(pkg, slide, sid)
    assert rpr.get("b") == "0"


def test_an_absent_toggle_still_writes_nothing(drawable):
    """MAJOR-A guard: absent must stay absent. "Present and false" and
    "not mentioned" are different instructions."""
    pkg, slide = drawable
    sid = _mk(pkg, slide, text="X", text_style={"size": 20})
    rpr, _b = _rpr(pkg, slide, sid)
    assert rpr.get("b") is None and rpr.get("i") is None


def test_underline_false_turns_underline_off_across_a_retext(drawable):
    """MAJOR-A, the same audit one key over."""
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    sid = _mk(pkg, slide, text="X", text_style={"underline": True})
    shp.set_shape(pkg, slide, sid, text="Y", text_style={"underline": False})
    rpr, _b = _rpr(pkg, slide, sid)
    assert rpr.get("u") == "none"


def test_wrap_true_states_itself_rather_than_relying_on_absence(drawable):
    """MAJOR-A, same family: wrap=True wrote nothing, so it could only
    mean "on" by luck of what the old body happened to carry."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    sid = _mk(pkg, slide, text="X", text_style={"wrap": False})
    _r, body = _rpr(pkg, slide, sid)
    assert body.find(_qn("a:bodyPr")).get("wrap") == "none"

    shp.set_shape(pkg, slide, sid, text="Y", text_style={"wrap": True})
    _r, body = _rpr(pkg, slide, sid)
    assert body.find(_qn("a:bodyPr")).get("wrap") == "square"


# ------------------------------------------------------------ MAJOR-B


def _crowd(pkg, slide, sid, scale="40000", red="20000"):
    """The cached autofit numbers PowerPoint writes on a crowded frame."""
    from kitchensink4ppt.core.package import qn as _qn
    _r, body = _rpr(pkg, slide, sid)
    bp = body.find(_qn("a:bodyPr"))
    na = etree.SubElement(bp, _qn("a:normAutofit"))
    na.set("fontScale", scale)
    na.set("lnSpcReduction", red)
    return bp


def test_a_stale_autofit_scale_does_not_ride_onto_the_new_text(drawable):
    """MAJOR-B: the old bodyPr was deep-copied wholesale, cached
    normAutofit numbers and all. That scale belonged to the OLD text, and
    PowerPoint renders from it until the frame is edited in the app, so a
    box that once overflowed kept shrinking text that now fits."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    sid = _mk(pkg, slide, text="long " * 20)
    _crowd(pkg, slide, sid)

    res = shp.set_shape(pkg, slide, sid, text="Short now.")
    _r, body = _rpr(pkg, slide, sid)
    na = body.find(f"{_qn('a:bodyPr')}/{_qn('a:normAutofit')}")

    assert na is not None, "the autofit MODE was dropped; only the numbers go"
    assert na.get("fontScale") is None, "a stale fontScale rode onto new text"
    assert na.get("lnSpcReduction") is None
    assert res.get("autofit_scale_reset") is True


def test_the_autofit_mode_itself_survives_a_retext(drawable):
    """MAJOR-B: dropping the mode would be a different defect. spAutoFit
    means the frame grows with the text and must stay."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    sid = _mk(pkg, slide, text="x")
    _r, body = _rpr(pkg, slide, sid)
    etree.SubElement(body.find(_qn("a:bodyPr")), _qn("a:spAutoFit"))

    res = shp.set_shape(pkg, slide, sid, text="y")
    _r, body = _rpr(pkg, slide, sid)
    assert body.find(f"{_qn('a:bodyPr')}/{_qn('a:spAutoFit')}") is not None
    assert res.get("autofit_scale_reset") is None


def test_a_frame_with_no_cached_scale_reports_no_reset(drawable):
    """MAJOR-B guard: the flag is a fact, not decoration."""
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    sid = _mk(pkg, slide, text="x")
    res = shp.set_shape(pkg, slide, sid, text="y")
    assert "autofit_scale_reset" not in res


# ------------------------------------------------------------ MINOR-J


def test_collapsing_mixed_runs_is_reported_not_claimed_as_preserved(drawable):
    """MINOR-J.3: a paragraph of differently formatted runs collapses to
    the FIRST run's formatting. main collapsed them too, so the loss is
    not new, but reporting preserved over the wreckage is."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    sid = _mk(pkg, slide, text="A", text_style={"color": "FF0000"})
    _r, body = _rpr(pkg, slide, sid)
    para = body.findall(_qn("a:p"))[0]
    run = etree.SubElement(para, _qn("a:r"))
    rpr = etree.SubElement(run, _qn("a:rPr"))
    rpr.set("b", "1")
    fill = etree.SubElement(rpr, _qn("a:solidFill"))
    etree.SubElement(fill, _qn("a:srgbClr")).set("val", "0000FF")
    etree.SubElement(run, _qn("a:t")).text = "B"

    res = shp.set_shape(pkg, slide, sid, text="merged")
    assert res.get("runs_collapsed_to_first") is True


def test_a_single_run_paragraph_reports_no_collapse(drawable):
    """MINOR-J.3 guard: the flag must not fire on every edit."""
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    sid = _mk(pkg, slide, text="A", text_style={"color": "FF0000"})
    res = shp.set_shape(pkg, slide, sid, text="B")
    assert "runs_collapsed_to_first" not in res


# ------------------------------------------------------------ MAJOR-C


def _dark_header_table(pkg, slide):
    """A table built with this server's own tools: explicit dark fill on
    row 0, text colour left to the table style part."""
    from kitchensink4ppt.ops import tables as tb

    res = tb.create_table(pkg, slide, rows=3, cols=3, x=1, y=1, w=8, h=2)
    sel = {"shape_id": res["shape_id"]}
    tb.set_table_cells(pkg, slide, sel, [
        {"row": 0, "col": c, "text": f"Header {c}"} for c in range(3)
    ])
    tb.format_table_cells(pkg, slide, sel,
                          range={"r1": 0, "c1": 0, "r2": 0, "c2": 2},
                          fill="1F3864")
    return res["shape_id"]


def test_a_dark_header_table_produces_no_fabricated_error(drawable):
    """MAJOR-C: the fill resolves (it is explicit), the header TEXT colour
    lives in the table style part, which this check cannot reach, so the
    colour fell through to the theme default (black) and the check
    reported an ERROR at gate severity on output this server itself
    produced. The header text is white."""
    from kitchensink4ppt.ops import design_check as dc

    pkg, slide = drawable
    sid = _dark_header_table(pkg, slide)
    findings = dc.check_layout(pkg, slide=slide, checks=["contrast"])[
        "findings"]
    bad = [
        f for f in findings
        if sid in f["shape_ids"] and f["severity"] in ("error", "warning")
    ]
    assert bad == [], f"fabricated a contrast failure from a guess: {bad}"


def test_a_guessed_colour_is_reported_as_unresolved_not_judged(drawable):
    """MAJOR-C: "cannot know" is a result. It comes back as info with the
    reason named, never at gate severity."""
    from kitchensink4ppt.ops import design_check as dc

    pkg, slide = drawable
    sid = _dark_header_table(pkg, slide)
    findings = dc.check_layout(pkg, slide=slide, checks=["contrast"])[
        "findings"]
    mine = [f for f in findings if sid in f["shape_ids"]]
    assert mine, "the table was passed in total silence"
    assert all(f["severity"] == "info" for f in mine)
    assert any(f.get("reason") == "unresolved_colour_source" for f in mine), mine


def test_a_real_low_contrast_run_is_still_an_error(drawable):
    """MAJOR-C guard: the downgrade must not blunt the check. A colour the
    slide STATES is still judged."""
    from kitchensink4ppt.ops import design_check as dc
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    sid = shp.insert_shape(
        pkg, slide, "rect", 1, 4, 4, 1, fill="1F3864", text="hard to read",
        text_style={"size": 12, "color": "203A60"},
    )["shape_id"]
    findings = dc.check_layout(pkg, slide=slide, checks=["contrast"])[
        "findings"]
    hit = [f for f in findings if sid in f["shape_ids"]]
    assert hit and hit[0]["severity"] in ("error", "warning")


def test_an_unresolved_placeholder_colour_is_downgraded_too(drawable):
    """MAJOR-C: the rule is about the SOURCE, not about tables. A shape
    whose colour only resolves to the theme default is a guess wherever it
    sits."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import design_check as dc
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    # Fill near-black, and say nothing at all about the text colour.
    sid = shp.insert_shape(pkg, slide, "rect", 1, 4, 4, 1, fill="111111",
                           text="unstated colour")["shape_id"]
    part = __import__("kitchensink4ppt.ops.read",
                      fromlist=["get_slide_info"]).get_slide_info(
        pkg, slide)["part"]
    elem, _c = shp._find_shape(pkg, part, sid)
    style = elem.find(_qn("p:style"))
    if style is not None:
        elem.remove(style)  # strip the fontRef so nothing but tx1 is left

    findings = dc.check_layout(pkg, slide=slide, checks=["contrast"])[
        "findings"]
    mine = [f for f in findings if sid in f["shape_ids"]]
    assert mine
    assert all(f["severity"] == "info" for f in mine), mine


def test_a_deck_local_table_style_resolves_instead_of_guessing(drawable):
    """MAJOR-C, the real fix where it is reachable: a table style DEFINED
    in tableStyles.xml is readable, so its cell text colour and fill are
    resolved rather than guessed. PowerPoint's built-in styles are only
    REFERENCED by GUID and their definitions are not in the file at all,
    which is why the downgrade above still has to exist."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import design_check as dc
    from kitchensink4ppt.ops import tables as tb

    pkg, slide = drawable
    res = tb.create_table(pkg, slide, rows=2, cols=2, x=1, y=1, w=6, h=1.5)
    sid = res["shape_id"]
    tb.set_table_cells(pkg, slide, {"shape_id": sid}, [
        {"row": 0, "col": 0, "text": "Header"},
        {"row": 1, "col": 0, "text": "Body"},
    ])
    tbl = tb.resolve_table(pkg, slide, {"shape_id": sid})["tbl"]
    style_id = tbl.find(f"{_qn('a:tblPr')}/{_qn('a:tableStyleId')}").text

    # Define that style locally: white text on a dark first row.
    lst = pkg.root("ppt/tableStyles.xml")
    style = etree.SubElement(lst, _qn("a:tblStyle"))
    style.set("styleId", style_id)
    style.set("styleName", "local test style")
    first = etree.SubElement(style, _qn("a:firstRow"))
    txs = etree.SubElement(first, _qn("a:tcTxStyle"))
    etree.SubElement(txs, _qn("a:srgbClr")).set("val", "FFFFFF")
    tcs = etree.SubElement(first, _qn("a:tcStyle"))
    fill = etree.SubElement(tcs, _qn("a:fill"))
    solid = etree.SubElement(fill, _qn("a:solidFill"))
    etree.SubElement(solid, _qn("a:srgbClr")).set("val", "1F3864")
    pkg.mark_dirty("ppt/tableStyles.xml")

    findings = dc.check_layout(pkg, slide=slide, checks=["contrast"])[
        "findings"]
    mine = [f for f in findings if sid in f["shape_ids"]]
    # White on dark navy resolves and passes: no gate finding, and no
    # "cannot resolve" info for the header either.
    assert not [f for f in mine if f["severity"] != "info"], mine

    from kitchensink4ppt.ops.read import get_slide_info
    ctx = dc._SlideCtx(pkg, {"part": get_slide_info(pkg, slide)["part"],
                             "index": slide, "slide_id": 0})
    resolved = dc._table_style_for(ctx, tbl)
    assert resolved is not None
    assert resolved.text_color(0, 0, 2, 2) == "FFFFFF"
    assert resolved.fill(0, 0, 2, 2) == "1F3864"


# ------------------------------------------------------------ MINOR-D


def test_an_other_family_placeholder_does_not_inherit_the_date_box(
    inheriting_placeholder
):
    """MINOR-D: layout_twin guards its family fallback with
    family != "other"; master_twin did not, and family_of() lumps pic,
    tbl, chart, ftr, dt and sldNum into one "other" bucket. Measured
    against a REAL master that has a date placeholder: a pic placeholder
    was resolving to it and inheriting the footer strip at the bottom of
    the slide."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import inherit as inh
    from kitchensink4ppt.ops.read import get_slide_info

    pkg, slide, _sid = inheriting_placeholder
    part = get_slide_info(pkg, slide)["part"]
    layout = inh.layout_part_of(pkg, part)
    master = inh.master_part_of(pkg, layout)

    assert inh.family_of("pic") == "other"
    assert inh.family_of("dt") == "other"
    # The master really does carry an "other"-family placeholder to be
    # wrongly matched, so this is a live trap, not a vacuous assertion.
    others = [
        sp for sp, ph in inh._placeholders(pkg, master)
        if inh.family_of(ph.get("type") or "body") == "other"
    ]
    assert others, "fixture master has no other-family placeholder to trip on"
    assert inh.master_twin(pkg, master, ("pic", None)) is None


def test_master_twin_still_matches_a_real_title_or_body(inheriting_placeholder):
    """MINOR-D guard: the fix must not blind the families that DO have one
    master placeholder each."""
    from kitchensink4ppt.ops import inherit as inh
    from kitchensink4ppt.ops.read import get_slide_info

    pkg, slide, _sid = inheriting_placeholder
    part = get_slide_info(pkg, slide)["part"]
    layout = inh.layout_part_of(pkg, part)
    master = inh.master_part_of(pkg, layout)
    assert inh.master_twin(pkg, master, ("title", None)) is not None
    assert inh.master_twin(pkg, master, ("body", "1")) is not None


def test_a_pic_placeholder_with_no_twin_gets_no_inherited_box(drawable):
    """MINOR-D, end to end: no box beats the wrong box, so set_shape
    refuses rather than seeding the footer strip."""
    from kitchensink4ppt.core.errors import UnsupportedStructure
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import inherit as inh
    from kitchensink4ppt.ops.read import get_slide_info
    from kitchensink4ppt.ops import shapes as shp

    pkg, slide = drawable
    part = get_slide_info(pkg, slide)["part"]
    tree = pkg.root(part).find(f"{_qn('p:cSld')}/{_qn('p:spTree')}")
    sp = etree.SubElement(tree, _qn("p:sp"))
    nv = etree.SubElement(sp, _qn("p:nvSpPr"))
    cnv = etree.SubElement(nv, _qn("p:cNvPr"))
    cnv.set("id", "900")
    cnv.set("name", "Picture Placeholder 900")
    etree.SubElement(nv, _qn("p:cNvSpPr"))
    nvpr = etree.SubElement(nv, _qn("p:nvPr"))
    ph = etree.SubElement(nvpr, _qn("p:ph"))
    ph.set("type", "pic")
    ph.set("idx", "77")
    etree.SubElement(sp, _qn("p:spPr"))
    body = etree.SubElement(sp, _qn("p:txBody"))
    etree.SubElement(body, _qn("a:bodyPr"))
    etree.SubElement(body, _qn("a:lstStyle"))
    etree.SubElement(body, _qn("a:p"))

    assert inh.inherited_box(pkg, part, sp) is None
    with pytest.raises(UnsupportedStructure):
        shp.set_shape(pkg, slide, 900, x=1, y=1, w=2, h=2)


# ------------------------------------------------------------ MINOR-H


def test_line_type_none_refuses_siblings_exactly_as_fill_does():
    """MINOR-H: the type=="none" branch returned before reading any other
    key, so line={"type":"none","width":4} dropped the width in silence.
    That is the same asymmetry #873 was raised to remove, left in place
    one branch over."""
    from kitchensink4ppt.core.errors import PptMcpError
    from kitchensink4ppt.ops import geometry as geo

    with pytest.raises(PptMcpError) as exc:
        geo.line_element({"type": "none", "width": 4, "color": "FF0000"})
    message = str(exc.value).lower()
    assert "width" in message and "color" in message

    with pytest.raises(PptMcpError):
        geo.fill_element({"type": "none", "color": "FF0000"})
    assert geo.line_element({"type": "none"}) is not None
    assert geo.line_element("none") is not None


# ------------------------------------------------------------ MINOR-F/G


def _grouped_pair(pkg, slide, name_a, name_b):
    from kitchensink4ppt.ops import shapes as shp
    a = shp.insert_shape(pkg, slide, "rect", 1.0, 1.0, 2.0, 1.0,
                         text="one", name=name_a)["shape_id"]
    b = shp.insert_shape(pkg, slide, "rect", 2.9, 1.0, 2.0, 1.0,
                         text="two", name=name_b)["shape_id"]
    return shp.group_shapes(pkg, slide, [a, b], name="Group")["group_id"]


def test_an_empty_exclude_list_means_exclude_nothing(drawable):
    """MINOR-F: opts.get("exclude_names") or DEFAULT made an empty list
    falsy, so there was no way to ask for an unfiltered in-group pass."""
    from kitchensink4ppt.ops import design_check as dc

    pkg, slide = drawable
    grp = _grouped_pair(pkg, slide, "tick one", "tick two")

    default = dc.check_layout(pkg, slide=slide, checks=["overlap"])["findings"]
    assert not [f for f in default if f.get("group_id") == grp]

    unfiltered = dc.check_layout(pkg, slide=slide, checks=[
        {"check": "overlap", "exclude_names": []}])["findings"]
    assert [f for f in unfiltered if f.get("group_id") == grp], (
        "exclude_names=[] still fell back to the default list"
    )


def test_the_exclusion_list_matches_whole_words_not_substrings(drawable):
    """MINOR-G: "band" in name.lower() silently exempted "Bandwidth chart"
    and "Brand box". Any name containing rule, tick, arrow or axis had the
    same hole."""
    from kitchensink4ppt.ops import design_check as dc

    pkg, slide = drawable
    grp = _grouped_pair(pkg, slide, "Bandwidth chart", "Brand box")
    hits = [
        f for f in dc.check_layout(pkg, slide=slide, checks=["overlap"])[
            "findings"]
        if f.get("group_id") == grp
    ]
    assert hits, "a real collision was dropped because a name contained band"


def test_a_genuine_band_is_still_exempt(drawable):
    """MINOR-G guard: word matching must not break the exclusion itself."""
    from kitchensink4ppt.ops import design_check as dc

    pkg, slide = drawable
    grp = _grouped_pair(pkg, slide, "lane band", "tick 3")
    hits = [
        f for f in dc.check_layout(pkg, slide=slide, checks=["overlap"])[
            "findings"]
        if f.get("group_id") == grp
    ]
    assert hits == []


def test_the_overlap_check_says_how_many_pairs_it_suppressed(drawable):
    """MINOR-G: filtering nobody can see is filtering nobody can check."""
    from kitchensink4ppt.ops import design_check as dc

    pkg, slide = drawable
    _grouped_pair(pkg, slide, "lane band", "tick 3")
    result = dc.check_layout(pkg, slide=slide, checks=["overlap"])
    caveats = result.get("checks") or []
    overlap = next(c for c in caveats if c.get("check") == "overlap")
    assert overlap.get("suppressed_by_exclude_names") >= 1


# ------------------------------------------------------------ MINOR-I


def test_tiny_text_honours_a_cached_font_scale(drawable):
    """MINOR-I: a shape at an explicit 28pt with normAutofit
    fontScale="25000" renders at 7pt and produced no finding, while
    _check_overflow in the same battery DID apply the scale. Two checks in
    one battery disagreeing about the size of the same text is worse than
    either being wrong alone."""
    from kitchensink4ppt.core.package import qn as _qn
    from kitchensink4ppt.ops import design_check as dc

    pkg, slide = drawable
    sid = _mk(pkg, slide, text="A body line long enough to count as body.",
              text_style={"size": 28})
    _r, body = _rpr(pkg, slide, sid)
    na = etree.SubElement(body.find(_qn("a:bodyPr")), _qn("a:normAutofit"))
    na.set("fontScale", "25000")

    findings = dc.check_layout(pkg, slide=slide, checks=["tiny_text"])[
        "findings"]
    hit = [f for f in findings if sid in f["shape_ids"]]
    assert hit, "28pt shrunk to 7pt by a cached scale was not flagged"
    assert hit[0]["sizes_pt"] == [7.0]


def test_an_unscaled_shape_is_unaffected(drawable):
    """MINOR-I guard."""
    from kitchensink4ppt.ops import design_check as dc

    pkg, slide = drawable
    sid = _mk(pkg, slide, text="A body line long enough to count as body.",
              text_style={"size": 28})
    findings = dc.check_layout(pkg, slide=slide, checks=["tiny_text"])[
        "findings"]
    assert not [f for f in findings if sid in f["shape_ids"]]


# ------------------------------------------------------------ MINOR-E


def test_a_format_text_edit_using_anchor_is_told_the_key_it_wants(drawable):
    """MINOR-E: an agent reads format_text(anchor=...) and writes
    {"op": "format_text", "anchor": "middle"}. It refused, correctly and
    atomically, with a message about slide-id anchor grammar that named no
    way forward."""
    from kitchensink4ppt.core.errors import PptMcpError
    from kitchensink4ppt.ops import batch

    pkg, slide = drawable
    sid = _mk(pkg, slide, text="hi")
    with pytest.raises(PptMcpError) as exc:
        batch.apply_edits(pkg, [{
            "op": "format_text", "slide": slide, "shape": sid,
            "anchor": "middle",
        }])
    assert "text_anchor" in str(exc.value)


def test_apply_edits_documents_text_anchor():
    """MINOR-E: the word appeared only in a source comment and in _OPS."""
    from kitchensink4ppt import server

    assert "text_anchor" in (server.apply_edits.__doc__ or "")
