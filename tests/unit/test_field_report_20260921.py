"""Regressions for the 2026-09-21 defense-deck field report (45 slides
built end to end through the server).

P3  copy_presentation to a folder that does not exist refused with the
    raw WinError text and a NOT_FOUND hint telling the caller to re-read
    the deck, which is not the problem and not the fix.
P5  set_placeholder_text rebuilt the body as bare a:p / a:r, so a text
    box whose formatting lives on paragraph-level pPr/defRPr came back
    with no indents, no bullet definition, no size, no colour and no
    typeface. (The set_shape / apply_edits set_text route was fixed for
    the 2026-09-20 report; the end-to-end test here pins all four
    text-writing routes against the same body.)
P6  a p:txBody with zero a:p passed layer 1 and made PowerPoint refuse
    the whole deck, and layer 2 reported "it failed" without naming the
    slide or the shape: about 40 minutes of manual bisection.
P8  create_table split the box evenly and offered no way to state row
    heights or column widths at creation.
887 the enable_tools note promised a tool list refresh that a worker or
    subagent never gets.
"""

from __future__ import annotations

import pytest
from lxml import etree
from pptx import Presentation
from pptx.util import Inches

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import shapes as shp
from kitchensink4ppt.ops.read import get_slide_info

# ------------------------------------------------------------------ P5

#: The body the field run actually had: every visual property on a:pPr
#: (indent, spacing, bullet) and a:defRPr (size, colour, typeface), and
#: runs carrying nothing but text. Nothing states an alignment or an
#: anchor, so a write that introduces one has invented it.
_MARL = "285750"
_INDENT = "-285750"


def _paragraph_level_body(body, lines):
    """Rebuild `body` as paragraphs whose formatting lives entirely on
    a:pPr / a:defRPr, with bare runs. `lines` is [(text, level), ...]."""
    for p in list(body.findall(qn("a:p"))):
        body.remove(p)
    bodypr = body.find(qn("a:bodyPr"))
    bodypr.attrib.pop("anchor", None)
    for text, level in lines:
        p = etree.SubElement(body, qn("a:p"))
        ppr = etree.SubElement(p, qn("a:pPr"))
        if level:
            ppr.set("lvl", str(level))
        ppr.set("marL", _MARL)
        ppr.set("indent", _INDENT)
        spc = etree.SubElement(ppr, qn("a:spcAft"))
        pts = etree.SubElement(spc, qn("a:spcPts"))
        pts.set("val", "600")
        bufont = etree.SubElement(ppr, qn("a:buFont"))
        bufont.set("typeface", "Arial")
        buchar = etree.SubElement(ppr, qn("a:buChar"))
        buchar.set("char", "•")
        defrpr = etree.SubElement(ppr, qn("a:defRPr"))
        defrpr.set("sz", "1400")
        fill = etree.SubElement(defrpr, qn("a:solidFill"))
        clr = etree.SubElement(fill, qn("a:srgbClr"))
        clr.set("val", "1F3864")
        latin = etree.SubElement(defrpr, qn("a:latin"))
        latin.set("typeface", "Calibri")
        r = etree.SubElement(p, qn("a:r"))  # bare run: no a:rPr at all
        t = etree.SubElement(r, qn("a:t"))
        t.text = text
    return body


@pytest.fixture()
def field_deck(tmp_path):
    """(pkg, slide, textbox_id, placeholder_id): one slide carrying a free
    text box and a body placeholder, both formatted the field-run way."""
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[1])  # title + content
    path = tmp_path / "field.pptx"
    prs.save(str(path))
    pkg = PptxPackage(str(path))
    sid = shp.insert_shape(
        pkg, 0, "rect", 1, 1, 6, 3, text="placeholder"
    )["shape_id"]
    part = get_slide_info(pkg, 0)["part"]
    elem, _chain = shp._find_shape(pkg, part, sid)
    _paragraph_level_body(
        elem.find(qn("p:txBody")),
        [("Alpha one", 0), ("Beta two", 1), ("Gamma three", 0)],
    )
    ph_id = None
    tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    for sp in tree.findall(qn("p:sp")):
        ph = sp.find(
            f"{qn('p:nvSpPr')}/{qn('p:nvPr')}/{qn('p:ph')}"
        )
        if ph is None or ph.get("type") not in (None, "body", "obj"):
            continue
        ph_id = int(
            sp.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}").get("id")
        )
        _paragraph_level_body(
            sp.find(qn("p:txBody")),
            [("Alpha one", 0), ("Beta two", 1), ("Gamma three", 0)],
        )
    assert ph_id is not None, "layout 1 gave no body placeholder"
    pkg.mark_dirty(part)
    return pkg, 0, sid, ph_id


def _paragraphs_of(pkg, slide, shape_id):
    part = get_slide_info(pkg, slide)["part"]
    elem, _chain = shp._find_shape(pkg, part, shape_id)
    body = elem.find(qn("p:txBody"))
    return body, body.findall(qn("a:p"))


def _assert_field_formatting_intact(pkg, slide, shape_id, route):
    """Every property the field report watched die, on every paragraph,
    plus the two attributes the old rebuild invented."""
    body, paras = _paragraphs_of(pkg, slide, shape_id)
    bodypr = body.find(qn("a:bodyPr"))
    assert bodypr.get("anchor") is None, (
        f"{route} invented a vertical anchor of {bodypr.get('anchor')!r}"
    )
    assert paras, f"{route} left the body with no paragraphs"
    for i, p in enumerate(paras):
        where = f"{route} paragraph {i}"
        ppr = p.find(qn("a:pPr"))
        assert ppr is not None, f"{where}: a:pPr is gone"
        assert ppr.get("algn") is None, f"{where}: alignment was invented"
        assert ppr.get("marL") == _MARL, f"{where}: marL lost"
        assert ppr.get("indent") == _INDENT, f"{where}: indent lost"
        spc = ppr.find(f"{qn('a:spcAft')}/{qn('a:spcPts')}")
        assert spc is not None and spc.get("val") == "600", (
            f"{where}: spcAft lost"
        )
        assert ppr.find(qn("a:buFont")).get("typeface") == "Arial", (
            f"{where}: buFont lost"
        )
        assert ppr.find(qn("a:buChar")).get("char") == "•", (
            f"{where}: buChar lost"
        )
        defrpr = ppr.find(qn("a:defRPr"))
        assert defrpr is not None, f"{where}: defRPr is gone"
        assert defrpr.get("sz") == "1400", f"{where}: defRPr size lost"
        clr = defrpr.find(f"{qn('a:solidFill')}/{qn('a:srgbClr')}")
        assert clr is not None and clr.get("val") == "1F3864", (
            f"{where}: defRPr colour lost"
        )
        assert defrpr.find(qn("a:latin")).get("typeface") == "Calibri", (
            f"{where}: defRPr typeface lost"
        )
    # the two outline levels survive as themselves
    assert [p.find(qn("a:pPr")).get("lvl") for p in paras[:3]] == [
        None, "1", None
    ], f"{route}: outline levels moved"


def test_set_placeholder_text_keeps_paragraph_level_formatting(field_deck):
    """P5: the route the 2026-09-20 fix did not cover. The rebuild dropped
    every a:pPr child and every run property; a 45-slide deck had to be
    written through python-pptx because of it."""
    from kitchensink4ppt.ops import text as txt

    pkg, slide, _sid, ph_id = field_deck
    out = txt.set_placeholder_text(
        pkg, slide, "body", paragraphs=[
            {"text": "Alpha one edited"},
            {"text": "Beta two edited", "level": 1},
            {"text": "Gamma three edited"},
        ],
    )
    assert out["paragraphs"] == 3
    _assert_field_formatting_intact(
        pkg, slide, ph_id, "set_placeholder_text"
    )
    _body, paras = _paragraphs_of(pkg, slide, ph_id)
    assert [
        p.find(f"{qn('a:r')}/{qn('a:t')}").text for p in paras
    ] == ["Alpha one edited", "Beta two edited", "Gamma three edited"]
    assert "preserved" in out, "the carry stayed silent about what it kept"
    assert set(out["preserved"]) >= {"bullets"}


def test_set_placeholder_text_carries_run_properties_too(field_deck):
    """P5: runs that DO carry rPr keep their size, colour and typeface."""
    from kitchensink4ppt.ops import text as txt

    pkg, slide, _sid, ph_id = field_deck
    part = get_slide_info(pkg, slide)["part"]
    elem, _chain = shp._find_shape(pkg, part, ph_id)
    for r in elem.find(qn("p:txBody")).iter(qn("a:r")):
        rpr = etree.SubElement(r, qn("a:rPr"))
        r.remove(rpr)
        r.insert(0, rpr)
        rpr.set("sz", "1100")
        fill = etree.SubElement(rpr, qn("a:solidFill"))
        etree.SubElement(fill, qn("a:srgbClr")).set("val", "C00000")
        etree.SubElement(rpr, qn("a:latin")).set("typeface", "Georgia")

    txt.set_placeholder_text(pkg, slide, "body", text="One\n\tTwo\nThree")
    _body, paras = _paragraphs_of(pkg, slide, ph_id)
    for i, p in enumerate(paras):
        rpr = p.find(f"{qn('a:r')}/{qn('a:rPr')}")
        assert rpr is not None, f"paragraph {i}: run properties are gone"
        assert rpr.get("sz") == "1100", f"paragraph {i}: run size lost"
        assert rpr.find(
            f"{qn('a:solidFill')}/{qn('a:srgbClr')}"
        ).get("val") == "C00000", f"paragraph {i}: run colour lost"
        assert rpr.find(qn("a:latin")).get("typeface") == "Georgia"


def test_set_placeholder_text_more_paragraphs_than_before(field_deck):
    """P5: extra paragraphs take the last old one's shape, the way
    pressing Enter in PowerPoint does, instead of arriving unformatted."""
    from kitchensink4ppt.ops import text as txt

    pkg, slide, _sid, ph_id = field_deck
    txt.set_placeholder_text(
        pkg, slide, "body",
        paragraphs=[{"text": f"Line {i}"} for i in range(6)],
    )
    _body, paras = _paragraphs_of(pkg, slide, ph_id)
    assert len(paras) == 6
    for i, p in enumerate(paras):
        ppr = p.find(qn("a:pPr"))
        assert ppr is not None and ppr.get("marL") == _MARL, (
            f"paragraph {i} arrived with no formatting"
        )


def test_set_placeholder_text_stated_level_beats_the_carried_one(field_deck):
    """P5 guard: preservation is not stickiness. A level the caller states
    wins over the level of the paragraph it replaces."""
    from kitchensink4ppt.ops import text as txt

    pkg, slide, _sid, ph_id = field_deck
    txt.set_placeholder_text(
        pkg, slide, "body",
        paragraphs=[
            {"text": "A", "level": 3},
            {"text": "B", "level": 0},
            {"text": "C", "level": 2},
        ],
    )
    _body, paras = _paragraphs_of(pkg, slide, ph_id)
    assert [p.find(qn("a:pPr")).get("lvl") for p in paras] == ["3", "0", "2"]
    # and the rest of the paragraph still rode through
    assert paras[1].find(qn("a:pPr")).get("marL") == _MARL


def test_every_text_writing_route_keeps_the_field_formatting(field_deck):
    """P5 end to end: the original repro, run against all four routes that
    write text into a shape. Any one of them failing is the field bug."""
    from kitchensink4ppt.ops import batch, text as txt

    pkg, slide, sid, ph_id = field_deck

    batch.apply_edits(pkg, [{
        "op": "set_text", "slide": slide, "shape": sid,
        "text": "Alpha one\nBeta two\nGamma three",
    }])
    _assert_field_formatting_intact(pkg, slide, sid, "apply_edits set_text")

    shp.set_shape(pkg, slide, sid, text="Alpha one\nBeta two\nGamma three")
    _assert_field_formatting_intact(pkg, slide, sid, "set_shape(text=)")

    txt.set_placeholder_text(
        pkg, slide, "body", text="Alpha one\n\tBeta two\nGamma three"
    )
    _assert_field_formatting_intact(
        pkg, slide, ph_id, "set_placeholder_text"
    )

    txt.search_and_replace(pkg, "Alpha", "Omega")
    _assert_field_formatting_intact(pkg, slide, sid, "search_and_replace")
    _assert_field_formatting_intact(
        pkg, slide, ph_id, "search_and_replace"
    )


def test_set_placeholder_text_with_no_paragraphs_leaves_one(field_deck):
    """P6 producer 1: paragraphs=[] used to leave a p:txBody with zero
    a:p, which is the corruption PowerPoint refuses to open."""
    from kitchensink4ppt.ops import text as txt

    pkg, slide, _sid, ph_id = field_deck
    txt.set_placeholder_text(pkg, slide, "body", paragraphs=[])
    _body, paras = _paragraphs_of(pkg, slide, ph_id)
    assert len(paras) == 1, "an empty body is a deck PowerPoint refuses"
    assert paras[0].find(qn("a:r")) is None


# ------------------------------------------------------------------ P6


def _empty_the_first_body(pkg, slide=0):
    """Strip every a:p from the first text body on a slide: the exact
    corruption the field run produced and then had to bisect by hand."""
    part = get_slide_info(pkg, slide)["part"]
    tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    for sp in tree.findall(qn("p:sp")):
        body = sp.find(qn("p:txBody"))
        if body is None:
            continue
        for p in list(body.findall(qn("a:p"))):
            body.remove(p)
        pkg.mark_dirty(part)
        return int(sp.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}").get("id"))
    raise AssertionError("no text body on the slide")


def test_save_refuses_a_text_body_with_no_paragraphs(tmp_path):
    """P6 layer 1: payload_valid said true on a deck PowerPoint would not
    open. The content model wants at least one a:p per body."""
    from kitchensink4ppt.core.errors import ValidationFailed

    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[1])
    path = tmp_path / "empty_body.pptx"
    prs.save(str(path))
    pkg = PptxPackage(str(path))
    shape_id = _empty_the_first_body(pkg)

    with pytest.raises(ValidationFailed) as excinfo:
        pkg.save(str(tmp_path / "out.pptx"))
    message = str(excinfo.value)
    assert "ppt/slides/slide1.xml" in message, message
    assert "a:p" in message or "paragraph" in message, message
    assert str(shape_id) in message, (
        f"the failure does not name the shape: {message}"
    )


def test_empty_table_cell_body_is_refused_too(tmp_path):
    """P6 layer 1: a:txBody in a table cell has the same content model."""
    from kitchensink4ppt.core.errors import ValidationFailed
    from kitchensink4ppt.ops import tables as tb

    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    path = tmp_path / "table.pptx"
    prs.save(str(path))
    pkg = PptxPackage(str(path))
    tb.create_table(pkg, 0, 2, 2, 1, 1, 6, 2, data=[["a", "b"], ["c", "d"]])
    part = get_slide_info(pkg, 0)["part"]
    tbl = pkg.root(part).find(f".//{qn('a:tbl')}")
    cell_body = tbl.find(f".//{qn('a:tc')}/{qn('a:txBody')}")
    for p in list(cell_body.findall(qn("a:p"))):
        cell_body.remove(p)
    pkg.mark_dirty(part)

    with pytest.raises(ValidationFailed) as excinfo:
        pkg.save(str(tmp_path / "out.pptx"))
    assert "a:p" in str(excinfo.value) or "paragraph" in str(excinfo.value)


def test_the_whole_corpus_still_validates(tmp_path):
    """P6 guard: _validate_payload runs on every save, so a new assertion
    that refuses a legitimate deck would brick the server. Every corpus
    deck (real or generated) must still round-trip."""
    from pathlib import Path

    corpus = Path(__file__).resolve().parents[1] / "corpus"
    decks = sorted(corpus.glob("*.pptx")) + sorted(corpus.glob("*.potx"))
    assert decks, "no corpus to check"
    for deck in decks:
        pkg = PptxPackage(str(deck))
        out = tmp_path / f"rt_{deck.stem}.pptx"
        pkg.save(str(out))  # raises ValidationFailed if the check is wrong
        assert out.is_file()


def test_full_load_names_the_slide_and_shape_that_failed():
    """P6 layer 2: 'it failed' with no slide and no shape cost about 40
    minutes of manual bisection on a 45-slide deck."""
    from kitchensink4ppt.com import bridge

    class _Shape:
        def __init__(self, name, ok=True):
            self.Name = name
            self._ok = ok

        @property
        def HasTextFrame(self):
            if not self._ok:
                raise OSError("the shape refused to load")
            return False

    class _Shapes:
        def __init__(self, shapes):
            self._shapes = shapes
            self.Count = len(shapes)

        def Item(self, i):
            return self._shapes[i - 1]

    class _Slide:
        def __init__(self, sid, shapes):
            self.SlideID = sid
            self.Shapes = _Shapes(shapes)

    class _Slides:
        def __init__(self, slides):
            self._slides = slides
            self.Count = len(slides)

        def Item(self, i):
            return self._slides[i - 1]

    class _Pres:
        def __init__(self, slides):
            self.Slides = _Slides(slides)

    pres = _Pres([
        _Slide(256, [_Shape("Title 1")]),
        _Slide(257, [_Shape("Body 2"), _Shape("Broken 3", ok=False)]),
    ])
    with pytest.raises(bridge.FullLoadFailed) as excinfo:
        bridge._full_load(pres)
    exc = excinfo.value
    assert exc.slide_index == 2
    assert exc.slide_id == 257
    assert exc.shape_index == 2
    assert exc.shape_name == "Broken 3"
    assert "slide 2" in str(exc) and "Broken 3" in str(exc)


def test_opens_clean_surfaces_the_failing_slide_and_shape(monkeypatch):
    """P6 layer 2: and the identity reaches the validate tool's result."""
    from kitchensink4ppt.com import bridge

    exc = bridge.FullLoadFailed(
        "slide 4 (id 260) shape 7 'Agenda Body': boom",
        slide_index=4, slide_id=260, shape_index=7, shape_name="Agenda Body",
    )
    out = bridge._opens_clean_failure(exc)
    assert out["opens_clean"] is False
    assert out["failed_slide_index"] == 4
    assert out["failed_slide_id"] == 260
    assert out["failed_shape_index"] == 7
    assert out["failed_shape_name"] == "Agenda Body"
    assert "boom" in out["error"]


def test_agenda_body_with_no_entries_still_has_a_paragraph(tmp_path):
    """P6 producer 2: the agenda body emptied itself and appended one
    paragraph per entry, so an empty entry list left zero a:p. Both public
    callers refuse an empty list today; the body builder is the layer that
    must not be able to produce the corruption at all."""
    from kitchensink4ppt.ops import assembly as asm

    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[1])
    path = tmp_path / "agenda.pptx"
    prs.save(str(path))
    pkg = PptxPackage(str(path))
    body_elem = asm._body_shape(pkg, pkg.slide_parts()[0], tagged=False)
    asm._fill_agenda_body(pkg, 0, body_elem, [], False, 0, [])
    body = body_elem.find(qn("p:txBody"))
    assert body.findall(qn("a:p")), "the agenda body has no paragraphs"
    pkg.save(str(tmp_path / "agenda_out.pptx"))


def test_notes_rebuilt_from_nothing_still_has_a_paragraph(tmp_path):
    """P6 producer 3: ops/notes.py _rebuild_body with an empty spec list."""
    from kitchensink4ppt.ops import notes as nt

    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    path = tmp_path / "notes.pptx"
    prs.save(str(path))
    pkg = PptxPackage(str(path))
    nt.set_notes(pkg, 0, "first draft of the note")
    nt.set_notes(pkg, 0, paragraphs=[])
    out = tmp_path / "notes_out.pptx"
    pkg.save(str(out))  # layer 1 refuses a body with no paragraphs
    assert out.is_file()


# ------------------------------------------------------------------ P3


def _copy_presentation(**kwargs):
    from kitchensink4ppt import packs, server  # noqa: F401  (registers tools)

    return packs.tool_objects()["copy_presentation"].fn(**kwargs)


def _deck_at(path):
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.save(str(path))
    return path


def test_copy_presentation_happy_path(tmp_path):
    """P3: copy_presentation had no behavioural test at all."""
    src = _deck_at(tmp_path / "src.pptx")
    dest = tmp_path / "copy.pptx"
    out = _copy_presentation(file_path=str(src), dest_path=str(dest))
    assert out["ok"] is True
    assert out["overwrote_existing"] is False
    assert dest.is_file()
    assert dest.read_bytes() == src.read_bytes()


def test_copy_presentation_refuses_an_existing_destination(tmp_path):
    src = _deck_at(tmp_path / "src.pptx")
    dest = _deck_at(tmp_path / "dest.pptx")
    out = _copy_presentation(file_path=str(src), dest_path=str(dest))
    assert out["ok"] is False
    assert out["error"]["code"] == "CONFLICT"
    assert "overwrite=True" in out["error"]["message"]


def test_copy_presentation_overwrite_rotates_the_previous_content(tmp_path):
    src = _deck_at(tmp_path / "src.pptx")
    dest = tmp_path / "dest.pptx"
    _deck_at(dest)
    before = dest.read_bytes()
    out = _copy_presentation(
        file_path=str(src), dest_path=str(dest), overwrite=True
    )
    assert out["ok"] is True
    assert out["overwrote_existing"] is True
    assert dest.read_bytes() == src.read_bytes()
    from kitchensink4ppt.core import safesave

    prev = safesave.slot_dir(dest) / safesave.PREV_SLOT
    anchor = safesave.slot_dir(dest) / safesave.ANCHOR_SLOT
    kept = [p for p in (prev, anchor) if p.is_file()]
    assert kept, "the overwritten content was not kept in any backup slot"
    assert any(p.read_bytes() == before for p in kept)


def test_copy_presentation_missing_parent_says_so(tmp_path):
    """P3: the refusal said '[WinError 3] The system cannot find the path
    specified' and hinted at re-reading the deck. Neither is true: the
    deck is fine, the destination folder does not exist."""
    src = _deck_at(tmp_path / "src.pptx")
    dest = tmp_path / "no_such_folder" / "copy.pptx"
    out = _copy_presentation(file_path=str(src), dest_path=str(dest))
    assert out["ok"] is False
    error = out["error"]
    assert "destination folder does not exist" in error["message"]
    assert str(dest.parent) in error["message"]
    assert "get_presentation_view" not in error["hint"], (
        f"the misleading hint is still attached: {error['hint']!r}"
    )
    assert error["hint"], "the refusal gives the caller nothing to do"
    assert not dest.parent.exists(), "a typo in a path created a folder"


def test_a_declared_hint_overrides_the_blanket_one():
    """P3 mechanism: the per-code _HINTS entry is a fallback, not a
    verdict, so a refusal that knows better can say so."""
    from kitchensink4ppt import server
    from kitchensink4ppt.core.errors import TargetNotFound

    plain = TargetNotFound("no shape 9 on slide 1")
    assert "get_presentation_view" in server._refusal(plain)["error"]["hint"]

    exc = TargetNotFound("no shape 9 on slide 1")
    exc.hint = "do this other thing instead"
    assert (
        server._refusal(exc)["error"]["hint"] == "do this other thing instead"
    )


# ------------------------------------------------------------------ P8


def test_create_table_accepts_row_heights_and_col_widths(tmp_path):
    """P8: a 27-row reference table at 10 pt needed per-row heights, and
    the only route was a second call after the fact."""
    from kitchensink4ppt.ops import tables as tb
    from kitchensink4ppt.ops import geometry as g

    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    path = tmp_path / "t.pptx"
    prs.save(str(path))
    pkg = PptxPackage(str(path))
    tb.create_table(
        pkg, 0, 3, 2, 1, 1, 6, 3,
        row_heights=[0.5, 1.0, 1.5], col_widths=[2.0, 4.0],
    )
    part = get_slide_info(pkg, 0)["part"]
    tbl = pkg.root(part).find(f".//{qn('a:tbl')}")
    widths = [int(c.get("w")) for c in tbl.iter(qn("a:gridCol"))]
    heights = [int(r.get("h")) for r in tbl.findall(qn("a:tr"))]
    assert widths == [g.in_to_emu(2.0), g.in_to_emu(4.0)]
    assert heights == [g.in_to_emu(h) for h in (0.5, 1.0, 1.5)]


def test_create_table_accepts_an_index_dict(tmp_path):
    from kitchensink4ppt.ops import tables as tb
    from kitchensink4ppt.ops import geometry as g

    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    path = tmp_path / "t.pptx"
    prs.save(str(path))
    pkg = PptxPackage(str(path))
    tb.create_table(pkg, 0, 3, 2, 1, 1, 6, 3, row_heights={1: 2.0})
    part = get_slide_info(pkg, 0)["part"]
    tbl = pkg.root(part).find(f".//{qn('a:tbl')}")
    heights = [int(r.get("h")) for r in tbl.findall(qn("a:tr"))]
    assert heights[1] == g.in_to_emu(2.0)
    # rows the dict did not name split what is left of the box, so the
    # table still fits the h it was given
    assert heights[0] == heights[2]
    assert sum(heights) == g.in_to_emu(3)


def test_create_table_refuses_a_wrong_length_list(tmp_path):
    from kitchensink4ppt.core.errors import PptMcpError
    from kitchensink4ppt.ops import tables as tb

    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    path = tmp_path / "t.pptx"
    prs.save(str(path))
    pkg = PptxPackage(str(path))
    with pytest.raises(PptMcpError) as excinfo:
        tb.create_table(pkg, 0, 3, 2, 1, 1, 6, 3, row_heights=[1.0, 1.0])
    assert "3 row" in str(excinfo.value)


def test_create_table_refuses_sizes_that_overflow_the_box(tmp_path):
    from kitchensink4ppt.core.errors import PptMcpError
    from kitchensink4ppt.ops import tables as tb

    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    path = tmp_path / "t.pptx"
    prs.save(str(path))
    pkg = PptxPackage(str(path))
    with pytest.raises(PptMcpError) as excinfo:
        tb.create_table(
            pkg, 0, 2, 2, 1, 1, 6, 3, col_widths=[5.0, 5.0]
        )
    message = str(excinfo.value)
    assert "col_widths" in message and "6" in message


def test_create_table_description_stays_in_budget():
    from kitchensink4ppt import packs, server  # noqa: F401  (registers tools)

    tool = packs.tool_objects()["create_table"]
    desc = tool.description or ""
    assert "row_heights" in desc and "col_widths" in desc
    assert round(len(desc) / 4) <= 130, (
        f"create_table description is ~{round(len(desc) / 4)} tokens"
    )
