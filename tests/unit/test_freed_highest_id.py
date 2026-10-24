"""Deleting the shape (or slide) with the highest id says its anchor will be reused.

A view anchor names a shape by its slide id and shape id. PowerPoint gives
a new shape one more than the highest id left on the slide, and so does
this server (measured: after a reopen PowerPoint reused a deleted shape's
id, and it reused a deleted last slide's id even in the same session). So
after the highest-numbered shape is deleted, the next new shape takes its
id, and an anchor read before the delete names the new shape: an edit sent
with it lands there, though the view module promised that deleting a shape
makes its anchor stale. The anchor scheme stays (the live route can see
only these two ids); the delete now says which anchors the next new shapes
will answer to, and the promise is corrected.

The deck is made up and built by code.
"""

from __future__ import annotations

from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import shapes as shp
from kitchensink4ppt.ops import slides as sl
from kitchensink4ppt.ops import view as vw


def _slide_with_shapes(make_deck):
    pkg = PptxPackage(make_deck("ids.pptx", extra_slides=0))
    slide = sl.insert_slide(pkg, 0)["index"]
    keep = shp.insert_shape(pkg, slide, "rect", 1, 1, 2, 1, text="KEEP")["shape_id"]
    old = shp.insert_shape(pkg, slide, "rect", 4, 1, 2, 1, text="OLD-UMBER")["shape_id"]
    return pkg, slide, keep, old


def test_deleting_the_highest_id_names_the_anchor_a_new_shape_will_take(make_deck):
    pkg, slide, _keep, old = _slide_with_shapes(make_deck)
    slide_id = sl._resolve_slide(pkg, slide)[3]
    anchor = "a:" + vw.anchor_digest(slide_id, old)[:6]
    res = shp.delete_shape(pkg, slide, old)
    notes = " ".join(res.get("warnings", []))
    assert anchor in notes and str(old) in notes  # before the fix: nothing said
    new = shp.insert_shape(pkg, slide, "ellipse", 4, 3, 2, 1, text="NEW")["shape_id"]
    assert new == old  # the id really is reused, as PowerPoint does


def test_deleting_a_lower_id_says_nothing(make_deck):
    pkg, slide, keep, _old = _slide_with_shapes(make_deck)
    res = shp.delete_shape(pkg, slide, keep)
    assert "warnings" not in res


def test_ungrouping_a_group_with_the_highest_id_names_its_anchor(make_deck):
    pkg, slide, keep, old = _slide_with_shapes(make_deck)
    gid = shp.group_shapes(pkg, slide, [keep, old])["group_id"]
    slide_id = sl._resolve_slide(pkg, slide)[3]
    res = shp.ungroup_shapes(pkg, slide, gid)
    assert "a:" + vw.anchor_digest(slide_id, gid)[:6] in " ".join(res.get("warnings", []))


def test_deleting_the_last_slide_id_names_its_anchor(make_deck):
    pkg = PptxPackage(make_deck("slides.pptx"))
    entries = sl._slide_entries(pkg)
    top = max(range(len(entries)), key=lambda i: int(entries[i].get("id")))
    top_id = int(entries[top].get("id"))
    res = sl.delete_slide(pkg, top)
    assert f"s:{top_id}" in " ".join(res.get("warnings", []))
    assert sl.insert_slide(pkg, 0)["slide_id"] == top_id  # reused, as PowerPoint does
    other = sl.delete_slide(pkg, 0)
    assert "warnings" not in other
