"""Annotations for the fixed nine-tool surface and its operation registry."""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from kitchensink4ppt.core import tool_annotations as ann
from kitchensink4ppt.registry import public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface.plan import build
from kitchensink4ppt.surface.tools import make_tool

DOCS = Path(__file__).resolve().parents[2] / "docs"

MARKETING = {
    "best", "powerful", "easy", "easiest", "fast", "fastest", "pro",
    "premium", "ultimate", "smart", "magic", "magical", "free", "awesome",
    "amazing", "seamless", "effortless", "revolutionary", "unlimited",
    "kitchensink4ai", "kitchensink4word", "kitchensink4xl",
    "kitchensink4ppt", "kitchensink4web",
}

NON_DESTRUCTIVE_OPS = {
    "add_comment", "add_entrance_animation", "add_equation_to_shape",
    "add_layout_placeholder", "copy_presentation", "copy_shape_between",
    "copy_slide_between",
    "create_chart", "create_layout", "create_presentation",
    "create_snapshot", "create_table", "duplicate_slide", "export_handout",
    "export_pdf", "export_shape_image", "export_slide_image", "export_table",
    "generate_agenda_slide", "generate_diagram", "insert_audio",
    "insert_connector", "insert_equation", "insert_image",
    "insert_master_shape", "insert_shape", "insert_slide",
    "insert_table_cols", "insert_table_rows", "insert_textbox",
    "insert_video", "live_scroll_to", "merge_decks", "split_deck",
    "svg_to_shapes",
    # the update ops change the server program or its notice setting,
    # never a deck (core/update_flow.py)
    "update_notices", "update_now",
}


def _fixed_tools() -> dict:
    plan = build(Settings())
    return {name: make_tool(spec, plan) for name, spec in plan.tools.items()}


def _wire(tool) -> dict:
    return tool.to_mcp_tool().model_dump(exclude_none=True)


def _doc_rows(filename: str) -> list:
    text = (DOCS / filename).read_text(encoding="utf-8")
    rows = []
    for line in text.splitlines():
        if line.startswith("| `"):
            cells = [c.strip() for c in line.strip("|").split("|")]
            rows.append([c.strip("`") for c in cells])
    return rows


def test_every_fixed_tool_carries_a_non_empty_unique_title():
    tools = _fixed_tools()
    titles = {name: _wire(tool).get("title") for name, tool in tools.items()}
    assert all(titles.values())
    assert len(set(titles.values())) == len(titles)


def test_fixed_titles_are_short_and_carry_no_marketing_language():
    offenders = {}
    for name, tool in _fixed_tools().items():
        title = _wire(tool)["title"]
        if title.startswith("COPY_SLOT:"):
            assert "FACTS TO CONVEY:" in title, (name, title)
            continue
        assert len(title) <= 40, (name, title)
        hit = set(re.findall(r"[a-z0-9]+", title.lower())) & MARKETING
        if hit:
            offenders[name] = sorted(hit)
    assert not offenders


def test_operation_title_table_covers_the_registry_exactly():
    registered = set(public_ops())
    assert set(ann.TITLES) == registered


def test_an_untitled_name_raises_rather_than_getting_a_generated_title():
    with pytest.raises(RuntimeError):
        ann.title("a_tool_that_does_not_exist")


def test_the_published_title_table_matches_the_module():
    rows = _doc_rows("TOOL_TITLES.md")
    published = {name: title for name, title, *_ in rows}
    assert {name: published[name] for name in ann.TITLES} == ann.TITLES
    assert set(published) - set(ann.TITLES) <= {"enable_tools", "disable_tools"}


def test_every_operation_has_explicit_read_only_and_destructive_classes():
    missing = [
        name for name, spec in public_ops().items()
        if not isinstance(spec.read_only, bool)
        or not isinstance(spec.destructive, bool)
    ]
    assert not missing


def test_no_read_only_operation_is_destructive():
    noisy = [
        name for name, spec in public_ops().items()
        if spec.read_only and spec.destructive
    ]
    assert not noisy


def test_destructive_classification_covers_mutating_operations():
    mutating = {name for name, spec in public_ops().items() if not spec.read_only}
    classified = ann.DESTRUCTIVE | ann.NON_DESTRUCTIVE
    assert classified == mutating
    assert not (ann.DESTRUCTIVE & ann.NON_DESTRUCTIVE)


def test_non_destructive_set_matches_hand_audited_allowlist():
    actual = {
        name for name, spec in public_ops().items()
        if not spec.read_only and not spec.destructive
    }
    assert actual == NON_DESTRUCTIVE_OPS


def test_an_unclassified_mutating_operation_raises():
    with pytest.raises(RuntimeError):
        ann.annotations("a_tool_that_does_not_exist", read_only=False)


def test_fixed_tool_annotations_are_complete_and_closed_world():
    for name, tool in _fixed_tools().items():
        annotations = tool.annotations
        assert annotations is not None
        assert annotations.readOnlyHint is not None
        assert annotations.destructiveHint is not None
        assert annotations.idempotentHint is not None
        # only ppt_review reaches a remote service (update_now: PyPI, GitHub)
        assert annotations.openWorldHint is (name == "ppt_review")
        if annotations.readOnlyHint:
            assert annotations.destructiveHint is False
            assert annotations.idempotentHint is True


def test_published_annotation_tables_match_the_module():
    text = (DOCS / "TOOL_ANNOTATIONS.md").read_text(encoding="utf-8")
    destructive, non_destructive = set(), set()
    bucket = None
    for line in text.splitlines():
        if line.startswith("## Tools that may perform destructive"):
            bucket = destructive
        elif line.startswith("## Tools that may not"):
            bucket = non_destructive
        elif line.startswith("## "):
            bucket = None
        elif bucket is not None and line.startswith("| `"):
            bucket.add(line.strip("|").split("|")[0].strip().strip("`"))
    assert destructive == ann.DESTRUCTIVE
    assert non_destructive & set(ann.NON_DESTRUCTIVE) == set(ann.NON_DESTRUCTIVE)
    assert non_destructive - set(ann.NON_DESTRUCTIVE) <= {
        "enable_tools", "disable_tools"
    }


def test_every_destructive_row_states_a_reason():
    text = (DOCS / "TOOL_ANNOTATIONS.md").read_text(encoding="utf-8")
    thin = []
    for line in text.splitlines():
        if not line.startswith("| `"):
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        if len(cells) == 2 and len(cells[1]) < 20:
            thin.append(cells[0])
    assert not thin
