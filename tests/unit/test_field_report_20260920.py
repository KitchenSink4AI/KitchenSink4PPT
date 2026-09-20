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
