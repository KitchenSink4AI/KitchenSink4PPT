"""A long answer says how much it holds before the bulk of it.

A client may show only the start of a long tool answer: Claude Code keeps an
answer over a few thousand characters in a file and shows the model its first
2,000 characters. The paging reads put their page block (the true total, how
many came back, whether the answer was cut and the offset that continues it)
after the items, near character 45,000 of a cut answer, so a model reading the
start saw a count and some items and never learned the answer was cut.

Every long success answer now leads with ok and file, then the page block,
then its short values (counts, totals, notes), then its long lists and texts.
A short answer keeps its order exactly. Every deck here is made up by code.
"""

from __future__ import annotations

import asyncio
import json

import pytest
from pptx import Presentation
from pptx.util import Inches

from kitchensink4ppt import server  # noqa: F401  (the import registers ops)
from kitchensink4ppt.core import budget
from kitchensink4ppt.registry import public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.envelope import compact_json
from kitchensink4ppt.surface.plan import build

PREVIEW = 2_000


def _call(tool: str, op: str, args: dict) -> dict:
    payload, is_error = asyncio.run(
        dispatch(build(Settings(), public_ops()), tool, {"op": op, "args": args}))
    assert is_error is False, payload
    return payload


@pytest.fixture(scope="module")
def long_deck(tmp_path_factory):
    """150 made-up slides of ten overlapping text boxes each: every paging
    read of it is cut at the output limit."""
    path = tmp_path_factory.mktemp("long") / "long_made_up.pptx"
    prs = Presentation()
    for s in range(150):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for b in range(10):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
            box.text_frame.text = f"Made-up line {b} on slide {s + 1} about the plan"
    prs.save(str(path))
    return str(path)


@pytest.mark.parametrize("tool, op, args", [
    ("ppt_read", "list_elements", {"kind": "shapes"}),
    ("ppt_read", "find_text", {"query": "plan"}),
    ("ppt_read", "get_presentation_view", {}),
    ("ppt_read", "get_text", {"live": "off"}),
    ("ppt_inspect", "check_layout", {}),
    ("ppt_inspect", "get_autofit_state", {}),
])
def test_a_cut_answer_says_its_total_in_the_first_2000_characters(long_deck, tool, op, args):
    payload = _call(tool, op, {"file_path": long_deck, **args})
    wire = compact_json(payload)
    assert len(wire) > PREVIEW
    page = payload["page"]
    assert page["truncated"] is True and page["next_offset"] is not None
    head = wire[:PREVIEW]
    assert '"page":{' in head
    assert f'"total":{page["total"]}' in head
    assert f'"next_offset":{page["next_offset"]}' in head
    assert list(payload)[0] == "ok"


def test_a_long_answer_about_one_slide_still_starts_with_the_slide(tmp_path):
    path = tmp_path / "crowded_made_up.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for b in range(400):
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
        box.text_frame.text = f"Made-up crowded line {b} about the plan"
    prs.save(str(path))
    payload = _call("ppt_read", "get_slide_info", {"file_path": str(path), "slide_number": 1})
    wire = compact_json(payload)
    assert len(wire) > PREVIEW
    keys = list(payload)
    assert keys[:2] == ["slide_number", "slide_count"]
    assert keys.index("ok") < keys.index("shapes")
    assert '"page":{' in wire[:PREVIEW] and payload["page"]["truncated"] is True


def test_short_values_come_before_the_long_lists(long_deck):
    payload = _call("ppt_read", "find_text", {"file_path": long_deck, "query": "plan"})
    keys = list(payload)
    assert keys.index("count") < keys.index("matches")
    assert keys.index("page") < keys.index("count")


def test_a_short_answer_keeps_its_order(tmp_path):
    path = tmp_path / "short_made_up.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1)).text_frame.text = "Made-up plan"
    prs.save(str(path))
    payload = _call("ppt_read", "find_text", {"file_path": str(path), "query": "plan"})
    assert len(compact_json(payload)) <= PREVIEW
    assert list(payload) == ["query", "regex", "count", "matches", "ok"]


def test_lead_with_totals_moves_nothing_in_or_out():
    long_list = [{"n": i, "text": "made-up " * 20} for i in range(40)]
    payload = {"items": long_list, "count": 40, "note": "made-up note",
               "page": {"total": 90, "truncated": True}, "ok": True, "file": "made_up.pptx"}
    led = budget.lead_with_totals(payload)
    assert led == payload
    assert list(led) == ["ok", "file", "page", "count", "note", "items"]
    small = {"count": 1, "ok": True}
    assert budget.lead_with_totals(small) is small


def test_llms_txt_says_long_answers_lead_with_their_totals():
    """The model-facing summary says what a long answer now does, so a client
    or a model reading only the start of one knows the totals are there."""
    from pathlib import Path

    text = (Path(__file__).resolve().parents[2] / "docs" / "llms.txt").read_text(encoding="utf-8")
    assert ("Long answers start with their totals and paging details, so a client that "
            "shows only the start of an answer still sees how much there is.") in text
