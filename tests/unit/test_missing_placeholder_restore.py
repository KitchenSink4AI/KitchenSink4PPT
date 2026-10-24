"""Regression for restoring a placeholder deleted from an existing slide."""

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import read, slides, text


def test_set_placeholder_text_restores_unique_layout_title(tmp_path):
    path = tmp_path / "deleted-title.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 1)  # Title and Content
    part = pkg.slide_parts()[0]
    tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    title = next(
        shape
        for shape in tree.findall(qn("p:sp"))
        if text._ph_of(shape).get("type", "obj") in ("title", "ctrTitle")
    )
    tree.remove(title)
    pkg.mark_dirty(part)

    out = text.set_placeholder_text(pkg, 0, "title", "Restored title")

    assert out["restored_from_layout"] is True
    restored = next(
        shape
        for shape in tree.findall(qn("p:sp"))
        if text._shape_id(shape) == out["shape_id"]
    )
    assert text._ph_of(restored).get("type") in ("title", "ctrTitle")
    assert restored.find(f"{qn('p:spPr')}/{qn('a:xfrm')}") is None
    assert text.shape_text(restored) == "Restored title"

    saved = pkg.save(tmp_path / "restored.pptx")
    reopened = PptxPackage(saved)
    assert "Restored title" in read.get_text(reopened)["text"]


def test_existing_placeholder_reports_that_it_was_not_restored(tmp_path):
    path = tmp_path / "normal-title.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 1)

    out = text.set_placeholder_text(pkg, 0, "title", "Existing title")

    assert out["restored_from_layout"] is False
