"""A deck whose slide ids reached the format's highest id.

New slide ids are given out above the highest in use and never reused, so a
view anchor or comment of a deleted slide never names a new one. At the
format's highest id (2147483647) a new slide is refused; the refusal now says
what happened and the next step, and that next step (merge_decks into a new
deck) works. Every deck here is made up by code.
"""

from __future__ import annotations

import pytest

from kitchensink4ppt.core.errors import PptMcpError
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import assembly, slides


def _deck_at_ceiling(tmp_path):
    path = tmp_path / "ceiling.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)
    slides.insert_slide(pkg, 6)
    ids = pkg.presentation().find(qn("p:sldIdLst"))
    ids[-1].set("id", "2147483647")
    pkg.mark_dirty("ppt/presentation.xml")
    pkg.save(do_backup=False)
    return path


def test_a_new_slide_at_the_ceiling_is_refused_with_the_next_step(tmp_path):
    path = _deck_at_ceiling(tmp_path)
    before = path.read_bytes()
    pkg = PptxPackage(path)
    with pytest.raises(PptMcpError) as refused:
        slides.insert_slide(pkg, 6)
    message = str(refused.value)
    assert "2147483647" in message
    assert "nothing was changed" in message
    assert "merge_decks" in message
    assert path.read_bytes() == before


def test_the_next_step_works(tmp_path):
    old = _deck_at_ceiling(tmp_path)
    new = tmp_path / "fresh.pptx"
    slides.create_presentation(new)
    pkg = PptxPackage(new)
    assembly.merge_decks(pkg, [str(old)])
    pkg.save(do_backup=False)
    pkg = PptxPackage(new)
    ids = [int(e.get("id")) for e in pkg.presentation().find(qn("p:sldIdLst"))]
    assert len(ids) >= 2 and max(ids) < 2147483647
    slides.insert_slide(pkg, 6)


def test_the_refusal_hint_is_about_the_deck_not_the_request(tmp_path):
    """The refusal ended with the generic BAD_PARAMS hint ("The request did
    not fit this operation ... ppt_help with this op gives its arguments"),
    which sends a caller to change arguments that were fine. Its hint names
    the step that works, as its message does."""
    from kitchensink4ppt.surface.envelope import refusal

    pkg = PptxPackage(_deck_at_ceiling(tmp_path))
    with pytest.raises(PptMcpError) as refused:
        slides.insert_slide(pkg, 6)
    error = refusal(refused.value)["error"]
    assert error["code"] == "BAD_PARAMS"
    assert "did not fit this operation" not in error["hint"], error["hint"]
    assert "merge_decks" in error["hint"] and "new deck" in error["hint"], error["hint"]
