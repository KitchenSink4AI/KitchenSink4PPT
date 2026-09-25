"""Punch-list #928 and #929, found by the Anthropic-directory prep run of
2026-09-24, which sent the submission's example prompts to the released
1.3.1 through headless Claude Code.

- P1 (accessibility) missed a deleted slide title, because any text near
  the top of a slide counted as a de-facto title, and missed a title read
  after its body, because the reading-order check only spoke up on three or
  more shapes with two inverted pairs (#928).
- P2 (fonts and colors) missed the stray fonts in 2 of 4 runs: the
  assistant enabled the design pack, whose extract_brand listed the literal
  fill colors but only the THEME fonts, so it reported one typeface on a
  deck carrying nine (#929).

The deck both prompts ran on ships here byte-for-byte as
tests/fixtures/q3_program_update_p1.pptx (MD5 below). Slide indexes are
0-based, as in every finding: slide 1 lost its title placeholder, slide 2
has a picture with no alt text, slide 3's title is read after its body,
slide 4 has a chart with no alt text."""

from __future__ import annotations

import hashlib
import shutil
import struct
import zlib
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import access, media, read, shapes, slides, text, themes
from kitchensink4ppt.ops import design_check as dc

FIXTURE = Path(__file__).resolve().parents[1] / "fixtures" / "q3_program_update_p1.pptx"
FIXTURE_MD5 = "bc2d83887bbea530b1157bc4e872a6d4"

#: Layouts of create_presentation's default template.
TITLE_AND_CONTENT, TITLE_ONLY, BLANK = 1, 5, 6


def png_bytes(w: int = 40, h: int = 20, rgb=(200, 30, 30)) -> bytes:
    def chunk(tag: bytes, payload: bytes) -> bytes:
        c = tag + payload
        return struct.pack(">I", len(payload)) + c + struct.pack(
            ">I", zlib.crc32(c) & 0xFFFFFFFF
        )

    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


@pytest.fixture()
def p1_deck(tmp_path):
    """A private copy of the P1/P2 input deck, so no test can touch the
    fixture itself."""
    dst = tmp_path / "Q3_Program_Update.pptx"
    shutil.copy2(FIXTURE, dst)
    return PptxPackage(dst)


def _deck(tmp_path, layout: int) -> PptxPackage:
    path = tmp_path / "deck.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, layout)
    return pkg


def _audit(pkg, check, slide=None):
    res = access.audit_accessibility(pkg, slide)
    return [f for f in res["findings"] if f["check"] == check]


def _layout(pkg, check, slide=0, **opts):
    entry = {"check": check, **opts} if opts else check
    return dc.check_layout(pkg, slide, [entry])["findings"]


def _order(pkg, slide=0) -> list[int]:
    part = pkg.slide_parts()[slide]
    tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    return [
        int(read._cnvpr(el).get("id"))
        for el in tree
        if read._cnvpr(el) is not None and el.tag != qn("p:nvGrpSpPr")
    ]


def _title_id(pkg, slide=0) -> int:
    part = pkg.slide_parts()[slide]
    tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    for el, kind, _z, _p in read.iter_shapes(tree):
        ph = read._ph(el) if kind == "placeholder" else None
        if ph is not None and ph.get("type") in ("title", "ctrTitle"):
            return int(read._cnvpr(el).get("id"))
    raise AssertionError("no title placeholder")


def _titled_slide(tmp_path, title="Downtime causes"):
    pkg = _deck(tmp_path, TITLE_ONLY)
    text.set_placeholder_text(pkg, 0, "title", title)
    return pkg, _title_id(pkg)


# ===================================================== the P1 input deck


def test_fixture_is_the_deck_the_prompts_ran_on():
    assert hashlib.md5(FIXTURE.read_bytes()).hexdigest() == FIXTURE_MD5


def test_p1_deck_audit_finds_all_three_problems_and_nothing_else(p1_deck):
    """The whole answer P1 asks for, from one call: the deleted title, the
    two graphics without alt text, the title read after its body. The
    table's unresolved-style contrast note is the only other finding."""
    res = access.audit_accessibility(p1_deck)
    got = sorted((f["check"], f["slide_index"]) for f in res["findings"])
    assert got == [
        ("alt_text", 2),
        ("alt_text", 4),
        ("contrast", 4),
        ("missing_title", 1),
        ("reading_order", 3),
    ]


def test_p1_deleted_title_is_flagged_and_its_lookalike_named(p1_deck):
    [f] = _audit(p1_deck, "missing_title")
    assert f["slide_index"] == 1
    # The content placeholder whose first bullet sits near the top used to
    # pass as the title; it is now named as the lookalike instead.
    assert f["candidate_shape_ids"] == [3]
    assert "Content Placeholder 2" in f["message"]
    assert "apply_layout never adds one" in f["fix"]


def test_p1_title_after_body_is_flagged_on_a_two_shape_slide(p1_deck):
    [f] = _audit(p1_deck, "reading_order")
    assert f["slide_index"] == 3
    assert f["rule"] == "title_first"
    assert f["severity"] == "warning"
    assert f["heuristic"] is False
    assert f["title_id"] == 2
    assert f["read_before_title"] == [3]
    assert f["document_order"] == [3, 2]
    assert f["suggested_order"] == [2, 3]
    assert "set_reading_order(slide=3, order=[2, 3])" in f["fix"]


def test_p1_suggested_order_clears_the_finding(p1_deck):
    [f] = _audit(p1_deck, "reading_order")
    access.set_reading_order(p1_deck, 3, f["suggested_order"])
    assert _audit(p1_deck, "reading_order") == []


def test_p1_audit_leaves_the_file_unchanged(p1_deck):
    before = hashlib.md5(Path(p1_deck.path).read_bytes()).hexdigest()
    access.audit_accessibility(p1_deck)
    assert not p1_deck._dirty
    assert hashlib.md5(Path(p1_deck.path).read_bytes()).hexdigest() == before


# ============================================== missing_title (#928, part 1)


def test_deleting_the_title_placeholder_is_flagged(tmp_path):
    """The P1 defect built through the real ops: a Title and Content slide
    whose title shape was deleted. Its content placeholder starts in the
    top quarter, which used to hide the finding."""
    pkg = _deck(tmp_path, TITLE_AND_CONTENT)
    text.set_placeholder_text(pkg, 0, "body", "Where we are\n\tBudget")
    shapes.delete_shape(pkg, 0, _title_id(pkg))
    [f] = _audit(pkg, "missing_title")
    assert f["candidate_shape_ids"] == [3]
    assert len(_layout(pkg, "missing_title")) == 1


def test_a_text_box_near_the_top_is_not_a_title(tmp_path):
    pkg = _deck(tmp_path, BLANK)
    box = text.insert_textbox(pkg, 0, "Looks Like A Title", 1, 0.4, 8, 1)
    [f] = _layout(pkg, "missing_title")
    assert f["candidate_shape_ids"] == [box["shape_id"]]
    assert f["shape_ids"] == []


def test_a_title_placeholder_with_text_is_a_title(tmp_path):
    pkg, _tid = _titled_slide(tmp_path)
    assert _audit(pkg, "missing_title") == []
    assert _layout(pkg, "missing_title") == []


def test_a_hidden_title_placeholder_still_titles_the_slide(tmp_path):
    """PowerPoint hides a slide title by keeping the placeholder and taking
    it out of view; screen readers and Outline view still read it."""
    pkg, tid = _titled_slide(tmp_path)
    part = pkg.slide_parts()[0]
    for el in pkg.root(part).iter(qn("p:cNvPr")):
        if el.get("id") == str(tid):
            el.set("hidden", "1")
    assert _audit(pkg, "missing_title") == []


def test_an_empty_title_placeholder_is_flagged_with_its_id(tmp_path):
    pkg = _deck(tmp_path, TITLE_ONLY)
    [f] = _layout(pkg, "missing_title")
    assert f["shape_ids"] == [_title_id(pkg)]
    assert "fills the existing empty title" in f["fix"]


def test_a_declared_title_name_counts_only_when_declared(tmp_path):
    pkg = _deck(tmp_path, BLANK)
    box = text.insert_textbox(pkg, 0, "Quarter in review", 1, 0.4, 8, 1)
    shapes.set_shape(pkg, 0, box["shape_id"], name="Headline")
    assert len(_layout(pkg, "missing_title")) == 1
    assert _layout(pkg, "missing_title", title_shape_names=["headline"]) == []
    # A bare string is one name, not a list of letters.
    assert _layout(pkg, "missing_title", title_shape_names="Headline") == []
    assert len(_layout(pkg, "missing_title", title_shape_names=["Other"])) == 1
    # The audit reports what a screen reader meets, so it takes no
    # declarations: a text box is not a title to it.
    assert len(_audit(pkg, "missing_title")) == 1


def test_a_hidden_or_empty_declared_shape_does_not_count(tmp_path):
    pkg = _deck(tmp_path, BLANK)
    box = shapes.insert_shape(pkg, 0, "rectangle", 1, 0.4, 8, 1)
    shapes.set_shape(pkg, 0, box["shape_id"], name="Headline")
    assert len(_layout(pkg, "missing_title", title_shape_names=["Headline"])) == 1


def test_title_shape_names_rejects_non_strings(tmp_path):
    from kitchensink4ppt.core.errors import PptMcpError

    pkg = _deck(tmp_path, BLANK)
    with pytest.raises(PptMcpError, match="must be str"):
        _layout(pkg, "missing_title", title_shape_names=[3])


# ============================================= reading order (#928, part 2)


def test_two_shape_slide_title_after_body_is_flagged(tmp_path):
    pkg, tid = _titled_slide(tmp_path)
    body = text.insert_textbox(pkg, 0, "Filter shortages: 62%", 1, 2, 10, 3)
    bid = body["shape_id"]
    access.set_reading_order(pkg, 0, [bid, tid])
    [f] = _audit(pkg, "reading_order", 0)
    assert (f["rule"], f["severity"]) == ("title_first", "warning")
    assert f["read_before_title"] == [bid]
    assert f["suggested_order"] == [tid, bid]


def test_two_shape_slide_title_first_is_clean(tmp_path):
    pkg, _tid = _titled_slide(tmp_path)
    text.insert_textbox(pkg, 0, "Filter shortages: 62%", 1, 2, 10, 3)
    assert _audit(pkg, "reading_order", 0) == []


def test_a_kicker_above_the_title_may_be_read_first(tmp_path):
    """Text ABOVE the title reads in the order the eye takes; that is not
    the defect."""
    pkg, tid = _titled_slide(tmp_path)
    kicker = text.insert_textbox(pkg, 0, "SECTION 2", 0.67, 0.0, 4, 0.25)
    access.set_reading_order(pkg, 0, [kicker["shape_id"], tid])
    assert _audit(pkg, "reading_order", 0) == []


def test_a_background_the_title_sits_on_is_exempt(tmp_path, tmp_path_factory):
    """A full-slide picture has to stay BEHIND the title, which in spTree
    means before it; moving it after would cover the title."""
    pkg, tid = _titled_slide(tmp_path)
    img = tmp_path_factory.mktemp("img") / "bg.png"
    img.write_bytes(png_bytes(160, 90))
    pic = media.insert_image(pkg, 0, str(img), 0, 0, 13.333, 7.5)
    access.set_reading_order(pkg, 0, [pic["shape_id"], tid])
    assert _audit(pkg, "reading_order", 0) == []


def test_a_decorative_picture_read_first_is_exempt(tmp_path, tmp_path_factory):
    pkg, tid = _titled_slide(tmp_path)
    img = tmp_path_factory.mktemp("img") / "logo.png"
    img.write_bytes(png_bytes())
    pic = media.insert_image(pkg, 0, str(img), 10, 6, 2, 1)
    access.set_reading_order(pkg, 0, [pic["shape_id"], tid])
    # Not decorative yet: a screen reader announces it before the title.
    [f] = _audit(pkg, "reading_order", 0)
    assert f["read_before_title"] == [pic["shape_id"]]
    # Office's "Mark as decorative": screen readers skip the shape.
    part = pkg.slide_parts()[0]
    for cnvpr in pkg.root(part).iter(qn("p:cNvPr")):
        if cnvpr.get("id") == str(pic["shape_id"]):
            ext_lst = etree.SubElement(cnvpr, qn("a:extLst"))
            ext = etree.SubElement(
                ext_lst, qn("a:ext"),
                uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}",
            )
            etree.SubElement(
                ext,
                "{http://schemas.microsoft.com/office/drawing/2017/decorative}"
                "decorative",
                val="1",
            )
    assert _audit(pkg, "reading_order", 0) == []


def test_untexted_and_hidden_shapes_before_the_title_are_ignored(tmp_path):
    pkg, tid = _titled_slide(tmp_path)
    band = shapes.insert_shape(pkg, 0, "rectangle", 0, 6.5, 13.333, 1)
    ghost = text.insert_textbox(pkg, 0, "draft note", 1, 5, 4, 1)
    part = pkg.slide_parts()[0]
    for cnvpr in pkg.root(part).iter(qn("p:cNvPr")):
        if cnvpr.get("id") == str(ghost["shape_id"]):
            cnvpr.set("hidden", "1")
    access.set_reading_order(pkg, 0, [band["shape_id"], ghost["shape_id"], tid])
    assert _audit(pkg, "reading_order", 0) == []


def test_gross_mismatch_with_a_late_title_is_one_warning(tmp_path):
    """Both rules on one slide yield ONE finding: the vote's full suggested
    order, raised to a warning, with the title named."""
    pkg, tid = _titled_slide(tmp_path)
    a = text.insert_textbox(pkg, 0, "Bottom", 1, 6, 3, 0.8)["shape_id"]
    b = text.insert_textbox(pkg, 0, "Middle", 1, 4, 3, 0.8)["shape_id"]
    c = text.insert_textbox(pkg, 0, "Upper", 1, 2, 3, 0.8)["shape_id"]
    access.set_reading_order(pkg, 0, [a, b, c, tid])
    [f] = _audit(pkg, "reading_order", 0)
    assert f["rule"] == "visual_order"
    assert f["severity"] == "warning"
    assert f["title_id"] == tid
    assert f["read_before_title"] == [a, b, c]
    assert f["suggested_order"][0] == tid
    access.set_reading_order(pkg, 0, f["suggested_order"])
    assert _audit(pkg, "reading_order", 0) == []


def test_reading_order_caveat_names_the_title_rule(p1_deck):
    res = access.audit_accessibility(p1_deck)
    assert "title_first" in res["caveats"]["reading_order"]
    assert "near the top" in res["caveats"]["missing_title"]


# ============================================ used fonts in design (#929)


def test_extract_brand_lists_the_fonts_the_slides_use(p1_deck):
    brand = themes.extract_brand(p1_deck)
    fonts = {f["typeface"]: f for f in brand["explicit_fonts"]}
    assert set(fonts) == {
        "Arial", "Arial Black", "Calibri", "Calibri Light", "Comic Sans MS",
        "Georgia", "Segoe UI", "Times New Roman", "Verdana",
    }
    assert fonts["Comic Sans MS"]["slides"] == [3]
    assert fonts["Comic Sans MS"]["count"] == 2
    assert fonts["Georgia"]["slides"] == [4]  # every table cell
    assert fonts["Arial"]["slides"] == [0, 3]
    # The theme font set explicitly is still the theme font; the rest are
    # the strays an audit is looking for.
    assert fonts["Calibri"]["theme_slots"] == ["major", "minor"]
    assert all(
        f["theme_slots"] == [] for name, f in fonts.items() if name != "Calibri"
    )
    counts = [f["count"] for f in brand["explicit_fonts"]]
    assert counts == sorted(counts, reverse=True)
    assert brand["explicit_font_total"] == 27
    assert brand["runs_without_explicit_font"] == 2
    assert "font_inventory" in brand["scope"]
    # The colors the prompt also asks about were already right; still are.
    assert {"1F4E79", "1F4E7A", "C00000", "FF0000"} <= {
        f["hex"] for f in brand["explicit_fills"]
    }
    assert not p1_deck._dirty


def test_extract_brand_skips_theme_references(tmp_path):
    pkg, _tid = _titled_slide(tmp_path)
    part = pkg.slide_parts()[0]
    for rpr in pkg.root(part).iter(qn("a:rPr")):
        etree.SubElement(rpr, qn("a:latin"), typeface="+mj-lt")
    brand = themes.extract_brand(pkg)
    assert brand["explicit_fonts"] == []
    assert brand["explicit_font_total"] == 0
    assert brand["runs_without_explicit_font"] >= 1


def test_apply_brand_takes_a_result_carrying_fonts(p1_deck, tmp_path):
    brand = themes.extract_brand(p1_deck)
    pkg = _deck(tmp_path, BLANK)
    res = themes.apply_brand(pkg, brand)
    assert res["themes_updated"]


def test_the_design_pack_reaches_the_fonts_in_use():
    """#929's invariant: an assistant that enables only the design pack can
    still see which typefaces the slides use."""
    from kitchensink4ppt import packs, server  # noqa: F401

    assert packs.pack_of("extract_brand") == "design"
    tool = packs.tool_objects()["extract_brand"]
    assert "typefaces" in (tool.description or "")
