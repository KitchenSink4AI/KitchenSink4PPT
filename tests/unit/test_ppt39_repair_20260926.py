"""Repairs to PR #39 (punch-list #928/#929) from the independent verify of
2026-09-24 (VERIFY #53, PPT section).

- P-B1: the title_first repair moved every "read too early" shape past the
  title, checking only that the moved shape did not overlap the TITLE.
  spTree order is also z-order, so a picture moved past an annotation
  drawn on it covered the annotation, and a re-audit reported nothing.
  No suggested reorder may change the relative z-order of two shapes whose
  boxes overlap; when the title cannot come first without that, no reorder
  is suggested and the overlapping shapes are named instead.
- P-M1: audit_accessibility merged only the first budget page of its
  delegated check_layout call, so on a large deck its by_check counts for
  missing_title and contrast were short (military_brief: 106 of 479).
- P-m1: a group of untexted autoshapes counted as content, while the same
  shapes ungrouped did not.
- P-m3: the whole-slide vote ignored Office's decorative flag and could
  suggest putting a decorative picture in front."""

from __future__ import annotations

from itertools import combinations

from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import access, media, shapes, slides, text
from kitchensink4ppt.ops import design_check as dc

from test_directory_fixes_20260924 import (
    BLANK,
    _audit,
    _deck,
    _order,
    _titled_slide,
    png_bytes,
)


def _boxes(pkg, slide=0) -> dict[int, tuple]:
    part = pkg.slide_parts()[slide]
    return {r["id"]: r["box"] for r in access._top_level_records(pkg, part)}


def _flipped_overlaps(pkg, before: list[int], after: list[int], slide=0):
    """Pairs of overlapping shapes whose relative spTree (= z) order differs
    between `before` and `after`."""
    boxes = _boxes(pkg, slide)
    rb = {s: i for i, s in enumerate(before)}
    ra = {s: i for i, s in enumerate(after)}
    out = []
    for a, b in combinations(before, 2):
        if not access._boxes_intersect(boxes[a], boxes[b]):
            continue
        if (rb[a] < rb[b]) != (ra[a] < ra[b]):
            out.append((a, b))
    return out


def _picture(pkg, tmp_path_factory, x, y, w, h, slide=0):
    img = tmp_path_factory.mktemp("img") / "pic.png"
    img.write_bytes(png_bytes(160, 90))
    return media.insert_image(pkg, slide, str(img), x, y, w, h)["shape_id"]


def _mark_decorative(pkg, sid, slide=0):
    part = pkg.slide_parts()[slide]
    for cnvpr in pkg.root(part).iter(qn("p:cNvPr")):
        if cnvpr.get("id") == str(sid):
            ext_lst = etree.SubElement(cnvpr, qn("a:extLst"))
            ext = etree.SubElement(
                ext_lst, qn("a:ext"), uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"
            )
            etree.SubElement(
                ext,
                "{http://schemas.microsoft.com/office/drawing/2017/decorative}"
                "decorative",
                val="1",
            )


# ======================================================= P-B1 (blocker)


def test_title_first_fix_keeps_an_annotation_on_its_picture(tmp_path, tmp_path_factory):
    """The verify's S8: a picture, a highlight circle drawn ON it, the title
    last in spTree. The old fix sent the picture above the circle."""
    pkg, tid = _titled_slide(tmp_path)
    pic = _picture(pkg, tmp_path_factory, 1, 2, 6, 4)
    circle = shapes.insert_shape(pkg, 0, "rectangle", 3, 3, 1, 1)["shape_id"]
    access.set_reading_order(pkg, 0, [pic, circle, tid])
    [f] = _audit(pkg, "reading_order", 0)
    assert f["rule"] == "title_first"
    assert f["read_before_title"] == [pic]
    before = _order(pkg)
    assert _flipped_overlaps(pkg, before, f["suggested_order"]) == []
    # The circle travels with the picture it is drawn on.
    assert f["suggested_order"] == [tid, pic, circle]
    access.set_reading_order(pkg, 0, f["suggested_order"])
    assert _audit(pkg, "reading_order", 0) == []


def test_title_first_fix_with_no_safe_order_names_the_overlap(tmp_path, tmp_path_factory):
    """A band drawn over the picture that also runs under the title: the
    picture cannot move past the title without leaving the band behind it,
    and the band cannot move past the title without covering it. No
    reorder is suggested; the shapes in the way are named."""
    pkg, tid = _titled_slide(tmp_path)
    tbox = _boxes(pkg)[tid]
    pic = _picture(pkg, tmp_path_factory, 1, 3, 6, 3)
    emu = 914400
    band_top = (tbox[1] + tbox[3] / 2) / emu
    band = shapes.insert_shape(
        pkg, 0, "rectangle", 0.5, band_top, 3, 4 - band_top
    )["shape_id"]
    access.set_reading_order(pkg, 0, [pic, band, tid])
    boxes = _boxes(pkg)
    assert access._boxes_intersect(boxes[band], boxes[tid])
    assert access._boxes_intersect(boxes[band], boxes[pic])
    [f] = _audit(pkg, "reading_order", 0)
    assert f["rule"] == "title_first"
    assert f["severity"] == "warning"
    assert f["read_before_title"] == [pic]
    assert f.get("suggested_order") is None
    assert "set_reading_order(" not in f["fix"]
    assert f["z_order_blocked"] == [{"shape_id": pic, "chain": [pic, band, tid]}]
    assert f"shape {band}" in f["fix"]


def test_combined_vote_and_title_fix_flips_no_overlapping_pair(tmp_path):
    """The vote's suggestion (raised to a warning when the title is late)
    used to put untexted shapes first, sending an annotation behind the
    box it is drawn on."""
    pkg, tid = _titled_slide(tmp_path)
    a = text.insert_textbox(pkg, 0, "Bottom", 1, 6, 3, 0.8)["shape_id"]
    b = text.insert_textbox(pkg, 0, "Middle", 1, 4, 3, 0.8)["shape_id"]
    c = text.insert_textbox(pkg, 0, "Upper", 1, 2, 3, 0.8)["shape_id"]
    mark = shapes.insert_shape(pkg, 0, "rectangle", 1.2, 6.1, 0.5, 0.5)["shape_id"]
    access.set_reading_order(pkg, 0, [a, mark, b, c, tid])
    [f] = _audit(pkg, "reading_order", 0)
    assert f["rule"] == "visual_order"
    before = _order(pkg)
    assert _flipped_overlaps(pkg, before, f["suggested_order"]) == []
    assert f["suggested_order"][0] == tid
    access.set_reading_order(pkg, 0, f["suggested_order"])
    assert _audit(pkg, "reading_order", 0) == []


def test_vote_alone_flips_no_overlapping_pair(tmp_path):
    pkg = _deck(tmp_path, BLANK)
    a = text.insert_textbox(pkg, 0, "Bottom", 1, 6, 3, 0.8)["shape_id"]
    b = text.insert_textbox(pkg, 0, "Middle", 1, 4, 3, 0.8)["shape_id"]
    c = text.insert_textbox(pkg, 0, "Upper", 1, 2, 3, 0.8)["shape_id"]
    mark = shapes.insert_shape(pkg, 0, "rectangle", 1.2, 6.1, 0.5, 0.5)["shape_id"]
    access.set_reading_order(pkg, 0, [a, mark, b, c])
    [f] = _audit(pkg, "reading_order", 0)
    assert f["rule"] == "visual_order"
    assert _flipped_overlaps(pkg, _order(pkg), f["suggested_order"]) == []


# ======================================================== P-M1 (major)


def test_audit_counts_every_delegated_finding_on_a_paged_deck(tmp_path, monkeypatch):
    """Enough untitled slides that check_layout's first page cannot hold
    them all: the audit must still count every one."""
    monkeypatch.setenv("KS4P_MAX_OUTPUT_CHARS", "4000")
    path = tmp_path / "deck.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    n = 40
    for i in range(n):
        slides.insert_slide(pkg, BLANK)
        text.insert_textbox(pkg, i, f"Looks like a title {i}", 1, 0.4, 8, 1)
    direct = dc.check_layout(pkg, None, ["missing_title"])
    assert direct["finding_count"] == n
    assert "page" in direct  # the delegated call really pages here
    res = access.audit_accessibility(pkg)
    assert res["by_check"]["missing_title"] == n
    assert res["finding_count"] == sum(res["by_check"].values())
    # Paging the audit itself walks every delegated finding.
    seen, offset = [], 0
    while True:
        page = access.audit_accessibility(pkg, offset=offset)
        seen.extend(
            f["slide_index"] for f in page["findings"] if f["check"] == "missing_title"
        )
        if "page" not in page or not page["page"].get("next_offset"):
            break
        offset = page["page"]["next_offset"]
    assert sorted(seen) == list(range(n))


# ======================================================= minors P-m1, P-m3


def test_untexted_group_before_the_title_is_not_content(tmp_path):
    """The verify's S4d: grouping two untexted accents must not turn them
    into content the title has to precede."""
    pkg, tid = _titled_slide(tmp_path)
    r1 = shapes.insert_shape(pkg, 0, "rectangle", 0, 7.0, 5, 0.2)["shape_id"]
    r2 = shapes.insert_shape(pkg, 0, "rectangle", 5, 7.0, 5, 0.2)["shape_id"]
    g = shapes.group_shapes(pkg, 0, [r1, r2])["group_id"]
    access.set_reading_order(pkg, 0, [g, tid])
    assert _audit(pkg, "reading_order", 0) == []


def test_group_holding_text_before_the_title_is_content(tmp_path):
    pkg, tid = _titled_slide(tmp_path)
    t1 = text.insert_textbox(pkg, 0, "Caption", 1, 5, 3, 0.6)["shape_id"]
    r1 = shapes.insert_shape(pkg, 0, "rectangle", 1, 5.7, 3, 0.2)["shape_id"]
    g = shapes.group_shapes(pkg, 0, [t1, r1])["group_id"]
    access.set_reading_order(pkg, 0, [g, tid])
    [f] = _audit(pkg, "reading_order", 0)
    assert f["read_before_title"] == [g]


def test_vote_ignores_a_decorative_picture(tmp_path, tmp_path_factory):
    """The verify's S2: a decorative picture read first must not drive the
    vote, whose text promises decorative shapes are kept behind."""
    pkg, tid = _titled_slide(tmp_path)
    pic = _picture(pkg, tmp_path_factory, 1, 4, 3, 2)
    _mark_decorative(pkg, pic)
    body = text.insert_textbox(pkg, 0, "Body", 5, 3, 4, 1)["shape_id"]
    access.set_reading_order(pkg, 0, [pic, tid, body])
    assert _audit(pkg, "reading_order", 0) == []


def test_combined_fix_reads_the_title_before_every_shape_it_can(tmp_path, tmp_path_factory):
    """military_brief slide 20 in miniature: the title must wait for a band
    it sits on, and the band for the picture under it. The picture cannot
    be read after the title without a z-order change, so it is named; the
    text boxes owe nothing to the band and must still follow the title."""
    pkg, tid = _titled_slide(tmp_path)
    tbox = _boxes(pkg)[tid]
    emu = 914400
    pic = _picture(pkg, tmp_path_factory, 7, 3, 5, 3)
    band_top = (tbox[1] + tbox[3] / 2) / emu
    band = shapes.insert_shape(
        pkg, 0, "rectangle", 7.5, band_top, 2, 4 - band_top
    )["shape_id"]
    a = text.insert_textbox(pkg, 0, "Bottom", 1, 6, 3, 0.8)["shape_id"]
    b = text.insert_textbox(pkg, 0, "Middle", 1, 4, 3, 0.8)["shape_id"]
    c = text.insert_textbox(pkg, 0, "Upper", 1, 2.2, 3, 0.8)["shape_id"]
    access.set_reading_order(pkg, 0, [pic, band, a, b, c, tid])
    [f] = _audit(pkg, "reading_order", 0)
    assert f["rule"] == "visual_order"
    assert f["read_before_title"] == [pic, a, b, c]
    assert f["z_order_blocked"] == [{"shape_id": pic, "chain": [pic, band, tid]}]
    s = f["suggested_order"]
    assert _flipped_overlaps(pkg, _order(pkg), s) == []
    assert all(s.index(tid) < s.index(x) for x in (a, b, c))
