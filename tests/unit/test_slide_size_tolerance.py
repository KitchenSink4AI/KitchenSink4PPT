"""A slide size a few hundred EMU off is the same slide size.

A deck built with python-pptx at 13.333 inches is 12192305 or 12191695 EMU
wide, against 12192000 for a 16:9 deck made in PowerPoint. Every
copy_slide_between from one into the other warned of a slide size mismatch
while the copies were correct. Sizes now count as equal when they differ by
less than 0.01% or less than 1,000 EMU (about 0.001 inch); a real mismatch
still warns, in copy_slide_between and in merge_decks.
"""

from __future__ import annotations

import asyncio

import pytest
from pptx import Presentation
from pptx.util import Emu, Inches

from kitchensink4ppt import server  # noqa: F401  (the import registers ops)
from kitchensink4ppt.registry import public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.plan import build

WIDE, TALL = 12192000, 6858000


def _call(op: str, args: dict):
    plan = build(Settings(), public_ops())
    return asyncio.run(dispatch(plan, "ppt_deck", {"op": op, "args": args}))


def _deck(path, cx: int, cy: int, text: str) -> str:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(cx), Emu(cy)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1)).text_frame.text = text
    prs.save(str(path))
    return str(path)


def _size_warnings(payload) -> list[str]:
    return [w for w in payload.get("warnings", []) if "slide size" in w]


@pytest.mark.parametrize("source_cx", [12192305, 12191695])
def test_a_python_pptx_wide_deck_copies_without_a_size_warning(tmp_path, source_cx):
    dest = _deck(tmp_path / "dest.pptx", WIDE, TALL, "Destination")
    src = _deck(tmp_path / "src.pptx", source_cx, TALL, "Source")
    payload, is_error = _call("copy_slide_between", {"file_path": dest, "source_path": src,
                                                     "slide": 0})
    assert is_error is False, payload
    assert _size_warnings(payload) == []


def test_a_real_mismatch_still_warns(tmp_path):
    dest = _deck(tmp_path / "dest.pptx", WIDE, TALL, "Destination")
    src = _deck(tmp_path / "src.pptx", 9144000, TALL, "Four by three")
    payload, is_error = _call("copy_slide_between", {"file_path": dest, "source_path": src,
                                                     "slide": 0})
    assert is_error is False, payload
    assert len(_size_warnings(payload)) == 1


def test_merge_decks_uses_the_same_tolerance(tmp_path):
    dest = _deck(tmp_path / "dest.pptx", WIDE, TALL, "Destination")
    near = _deck(tmp_path / "near.pptx", 12192305, TALL, "Near")
    far = _deck(tmp_path / "far.pptx", 9144000, TALL, "Far")
    payload, is_error = _call("merge_decks", {"file_path": dest, "sources": [near, far]})
    assert is_error is False, payload
    text = str(payload)
    assert "near.pptx is" not in text
    assert "far.pptx is" in text


@pytest.mark.parametrize("a,b,equal", [
    ((WIDE, TALL), (WIDE + 305, TALL), True),
    ((WIDE, TALL), (WIDE - 305, TALL), True),
    ((WIDE, TALL), (WIDE + 999, TALL - 999), True),
    ((WIDE, TALL), (WIDE + 1200, TALL), True),     # under 0.01% of 12192000
    ((WIDE, TALL), (WIDE + 1300, TALL), False),    # over both
    ((WIDE, TALL), (WIDE, TALL + 1001), False),    # 0.01% of the height is 686
    ((WIDE, TALL), (9144000, TALL), False),
])
def test_the_rule(a, b, equal):
    from kitchensink4ppt.ops.interdeck import same_slide_size

    assert same_slide_size(a, b) is equal
