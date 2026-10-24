"""Every argument that counts slides from 0 has a twin that counts from 1.

slide had slide_number and slides had slide_numbers; the other arguments
that count slides still took only the 0-based form, so a caller who learned
"slide_number counts from 1" on get_notes was refused, or silently off by
one, on copy_position, reorder_slides, split_deck, the scope of the read
operations, set_hyperlink, move_slide, insert_slide, copy_shape_between, and
the position of duplicate_slide, copy_slide_between and generate_agenda_slide.
Each now takes a twin counted from 1. Sending both forms is refused with a
hint rather than guessed between; a null twin reads as not sent; ppt_help
lists each twin right after its field; and a refusal's valid_fields names
the twins.
"""

from __future__ import annotations

import asyncio
import shutil

import jsonschema
import pytest
from pptx import Presentation

from kitchensink4ppt import server  # noqa: F401  (the import registers ops)
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.registry import public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface import normalize
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.help import help_payload
from kitchensink4ppt.surface.plan import build

TITLES = ["One", "Two", "Three", "Four"]


def _deck(path) -> str:
    prs = Presentation()
    for title in TITLES:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = title
    prs.save(str(path))
    return str(path)


def _call(tool: str, op: str, args: dict):
    plan = build(Settings(), public_ops())
    return asyncio.run(dispatch(plan, tool, {"op": op, "args": args}))


def _ok(tool: str, op: str, args: dict) -> dict:
    payload, is_error = _call(tool, op, args)
    assert is_error is False, payload
    return payload


def _titles(deck: str) -> list[str]:
    return [slide.shapes.title.text for slide in Presentation(deck).slides]


def _pair(tmp_path, name: str) -> tuple[str, str]:
    a = _deck(tmp_path / f"{name}_a.pptx")
    b = str(tmp_path / f"{name}_b.pptx")
    shutil.copy2(a, b)
    return a, b


# ------------------------------------------------- each twin means the same


def test_reorder_slides_order_numbers(tmp_path):
    a, b = _pair(tmp_path, "reorder")
    _ok("ppt_edit", "reorder_slides", {"file_path": a, "order": [2, 0, 1, 3]})
    _ok("ppt_edit", "reorder_slides", {"file_path": b, "order_numbers": [3, 1, 2, 4]})
    assert _titles(a) == _titles(b) == ["Three", "One", "Two", "Four"]


def test_move_slide_to_number(tmp_path):
    a, b = _pair(tmp_path, "move")
    _ok("ppt_edit", "move_slide", {"file_path": a, "slide": 3, "to": 0})
    _ok("ppt_edit", "move_slide", {"file_path": b, "slide_number": 4, "to_number": 1})
    assert _titles(a) == _titles(b) == ["Four", "One", "Two", "Three"]


def test_insert_slide_position_number(tmp_path):
    a, b = _pair(tmp_path, "insert")
    one = _ok("ppt_edit", "insert_slide", {"file_path": a, "layout": 6, "position": 1})
    two = _ok("ppt_edit", "insert_slide", {"file_path": b, "layout": 6, "position_number": 2})
    assert one["changed"]["index"] == two["changed"]["index"] == 1


def _op(op: str, args: dict) -> dict:
    """One call to op through the tool that serves it."""
    plan = build(Settings(), public_ops())
    payload, is_error = asyncio.run(
        dispatch(plan, plan.specs[op].group, {"op": op, "args": args}))
    assert is_error is False, payload
    return payload


def test_duplicate_slide_position_number(tmp_path):
    """duplicate_slide, copy_slide_between and generate_agenda_slide
    took position counted from 0 and refused position_number, although the
    changelog said every such argument had a twin counted from 1."""
    a, b = _pair(tmp_path, "duplicate")
    _op("duplicate_slide", {"file_path": a, "slide": 0, "position": 2})
    _op("duplicate_slide", {"file_path": b, "slide_number": 1, "position_number": 3})
    assert _titles(a) == _titles(b) == ["One", "Two", "One", "Three", "Four"]


def test_copy_slide_between_position_number(tmp_path):
    a, b = _pair(tmp_path, "between")
    source = _deck(tmp_path / "source.pptx")
    _op("copy_slide_between", {"file_path": a, "source_path": source, "slide": 3,
                               "position": 1})
    _op("copy_slide_between", {"file_path": b, "source_path": source, "slide": 3,
                               "position_number": 2})
    assert _titles(a) == _titles(b) == ["One", "Four", "Two", "Three", "Four"]


def test_generate_agenda_slide_position_number(tmp_path):
    a, b = _pair(tmp_path, "agenda")
    one = _op("generate_agenda_slide", {"file_path": a, "position": 2})
    two = _op("generate_agenda_slide", {"file_path": b, "position_number": 3})
    assert len(_titles(a)) == len(_titles(b)) == 5
    assert Presentation(a).slides[2].shapes.title.text ==         Presentation(b).slides[2].shapes.title.text
    assert one["changed"].get("index", one["changed"].get("position")) ==         two["changed"].get("index", two["changed"].get("position"))


def test_set_hyperlink_to_slide_number(tmp_path):
    a, b = _pair(tmp_path, "link")
    shape = server.get_slide_info(a, 0, live="off")["shapes"][0]["id"]
    one = _ok("ppt_edit", "set_hyperlink", {
        "file_path": a, "slide": 0, "target": shape, "to_slide": 2})
    two = _ok("ppt_edit", "set_hyperlink", {
        "file_path": b, "slide": 0, "target": shape, "to_slide_number": 3})
    assert one["changed"]["target_slide_number"] == two["changed"]["target_slide_number"] == 3


def test_copy_position_from_slide_number_and_to_slide_numbers(tmp_path):
    a, b = _pair(tmp_path, "position")
    for deck in (a, b):
        for index in range(4):
            _ok("ppt_graphics", "insert_shape", {
                "file_path": deck, "slide": index, "shape_type": "rect", "name": "Logo",
                "x": 1 + index, "y": 1, "w": 1, "h": 1})
    one = _ok("ppt_graphics", "copy_position", {
        "file_path": a, "from_slide": 0, "shape": "Logo", "to_slides": [1, 2]})
    two = _ok("ppt_graphics", "copy_position", {
        "file_path": b, "from_slide_number": 1, "shape": "Logo",
        "to_slide_numbers": [2, 3]})
    strip = lambda p: {k: v for k, v in p["changed"].items() if k not in ("file",)}  # noqa: E731
    assert strip(one) == strip(two)


def test_split_deck_number_ranges(tmp_path):
    a, b = _pair(tmp_path, "split")
    out_a, out_b = tmp_path / "out_a", tmp_path / "out_b"
    one = _ok("ppt_deck", "split_deck", {
        "file_path": a, "output_dir": str(out_a), "by": "ranges",
        "ranges": [{"start": 0, "end": 1, "name": "first"}, {"start": 2, "end": 3}]})
    two = _ok("ppt_deck", "split_deck", {
        "file_path": b, "output_dir": str(out_b), "by": "ranges",
        "number_ranges": [{"start": 1, "end": 2, "name": "first"}, {"start": 3, "end": 4}]})
    pieces = lambda p: [len(Presentation(f["path"]).slides) for f in p["outputs"]]  # noqa: E731
    assert pieces(one) == pieces(two) == [2, 2]


@pytest.mark.parametrize("op", ["get_text", "find_text", "list_hyperlinks"])
def test_a_scope_takes_slide_number_and_slide_numbers(tmp_path, op):
    deck = _deck(tmp_path / "scope.pptx")
    extra = {"query": "T"} if op == "find_text" else {}
    one = _ok("ppt_read", op, {"file_path": deck, "scope": 1, **extra})
    two = _ok("ppt_read", op, {"file_path": deck, "slide_number": 2, **extra})
    assert two["slide_number"] == one["slide_number"] == 2
    assert "slide_note" in one and "slide_note" not in two  # 1-based needs no note
    many = _ok("ppt_read", op, {"file_path": deck, "slide_numbers": [1, 3], **extra})
    zero = _ok("ppt_read", op, {"file_path": deck, "scope": [0, 2], **extra})
    drop = lambda p: {k: v for k, v in p.items() if k not in ("normalized_from",)}  # noqa: E731
    assert drop(many) == drop(zero)


def test_copy_shape_between_source_slide_number(tmp_path):
    src = _deck(tmp_path / "src.pptx")
    dst = _deck(tmp_path / "dst.pptx")
    _ok("ppt_graphics", "insert_shape", {
        "file_path": src, "slide": 1, "shape_type": "rect", "name": "Mark",
        "x": 1, "y": 1, "w": 1, "h": 1})
    payload = _ok("ppt_graphics", "copy_shape_between", {
        "file_path": dst, "source_path": src, "source_slide_number": 2,
        "shape": "Mark", "slide_number": 4})
    assert payload["changed"]["source"]["slide_number"] == 2
    assert payload["slide_number"] == 4
    payload, is_error = _call("ppt_graphics", "copy_shape_between", {
        "file_path": dst, "source_path": src, "source_slide": 9,
        "shape": "Mark", "slide": 0})
    assert is_error and "source_slide_number 1 to 4" in payload["error"]["message"]


# --------------------------------------------------------------- the rules

BOTH = [
    ("ppt_edit", "reorder_slides", {"order": [0, 1, 2, 3], "order_numbers": [1, 2, 3, 4]}),
    ("ppt_edit", "move_slide", {"slide": 0, "to": 1, "to_number": 2}),
    ("ppt_edit", "insert_slide", {"layout": 6, "position": 1, "position_number": 2}),
    (None, "duplicate_slide", {"slide": 0, "position": 1, "position_number": 2}),
    (None, "generate_agenda_slide", {"position": 1, "position_number": 2}),
    ("ppt_graphics", "copy_position", {
        "from_slide": 0, "from_slide_number": 1, "shape": "Title 1"}),
    ("ppt_graphics", "copy_position", {
        "from_slide": 0, "shape": "Title 1", "to_slides": [1], "to_slide_numbers": [2]}),
    ("ppt_read", "get_text", {"scope": 1, "slide_number": 2}),
    ("ppt_read", "get_text", {"scope": [1], "slide_numbers": [2]}),
    ("ppt_read", "get_text", {"slide_number": 1, "slide_numbers": [2]}),
]


@pytest.mark.parametrize("tool, op, args", BOTH)
def test_both_forms_are_refused_with_a_hint(tmp_path, tool, op, args):
    deck = _deck(tmp_path / "both.pptx")
    before = (tmp_path / "both.pptx").read_bytes()
    tool = tool or build(Settings(), public_ops()).specs[op].group
    payload, is_error = _call(tool, op, {"file_path": deck, **args})
    assert is_error is True, payload
    error = payload["error"]
    assert error["code"] == "BAD_ARGS"
    assert "both sent" in error["message"]
    assert "counted from 1" in error.get("hint", "")
    assert (tmp_path / "both.pptx").read_bytes() == before


def test_a_null_twin_reads_as_not_sent(tmp_path):
    deck = _deck(tmp_path / "null.pptx")
    _ok("ppt_edit", "move_slide", {"file_path": deck, "slide": 3, "to": 0, "to_number": None})
    _ok("ppt_read", "get_text", {"file_path": deck, "scope": 1, "slide_numbers": None})
    assert _titles(deck)[0] == "Four"


@pytest.mark.parametrize("op, twin, value", [
    ("reorder_slides", "order_numbers", [0, 1, 2, 3]),
    ("move_slide", "to_number", 0),
    ("get_text", "slide_numbers", ["first"]),
])
def test_a_twin_counts_from_one(tmp_path, op, twin, value):
    deck = _deck(tmp_path / "zero.pptx")
    tool = {"reorder_slides": "ppt_edit", "move_slide": "ppt_edit"}.get(op, "ppt_read")
    args = {"file_path": deck, twin: value}
    if op == "move_slide":
        args["slide"] = 0
    payload, is_error = _call(tool, op, args)
    assert is_error is True
    assert "counted from 1" in payload["error"]["message"]


TWINS = {
    "copy_position": [("from_slide", "from_slide_number"), ("to_slides", "to_slide_numbers")],
    "reorder_slides": [("order", "order_numbers")],
    "split_deck": [("ranges", "number_ranges")],
    "set_hyperlink": [("to_slide", "to_slide_number")],
    "move_slide": [("to", "to_number")],
    "insert_slide": [("position", "position_number")],
    "duplicate_slide": [("position", "position_number")],
    "copy_slide_between": [("position", "position_number")],
    "generate_agenda_slide": [("position", "position_number")],
    "copy_shape_between": [("source_slide", "source_slide_number")],
    "get_text": [("scope", "slide_number"), ("scope", "slide_numbers")],
    "list_comments": [("scope", "slide_number"), ("scope", "slide_numbers")],
    "export_slide_image": [("slides", "slide_numbers")],
}


def test_help_lists_each_twin_right_after_its_field():
    plan = build(Settings(), public_ops())
    for op, pairs in TWINS.items():
        names = list(help_payload(plan, op)["args_schema"]["properties"])
        for field, twin in pairs:
            assert twin in names, (op, twin)
            after = names[names.index(field) + 1:]
            assert twin in after[:2], (op, field, twin, names)
            description = help_payload(plan, op)["args_schema"]["properties"][twin]["description"]
            assert "from 1" in description, (op, twin)


def test_the_help_schema_takes_a_required_field_by_its_twin(tmp_path):
    plan = build(Settings(), public_ops())
    move = help_payload(plan, "move_slide")["args_schema"]
    jsonschema.validate({"file_path": "d.pptx", "slide": 0, "to_number": 2}, move)
    jsonschema.validate({"file_path": "d.pptx", "slide_number": 1, "to": 2}, move)
    with pytest.raises(jsonschema.ValidationError):
        jsonschema.validate({"file_path": "d.pptx", "slide": 0}, move)
    position = help_payload(plan, "copy_position")["args_schema"]
    jsonschema.validate({"file_path": "d.pptx", "from_slide_number": 1, "shape": "A"}, position)


def test_valid_fields_name_the_twins(tmp_path):
    deck = _deck(tmp_path / "fields.pptx")
    payload, is_error = _call("ppt_edit", "move_slide", {
        "file_path": deck, "slide": 0, "to": 1, "nonsense": 1})
    assert is_error is True
    fields = payload["error"]["details"]["valid_fields"]
    assert fields.index("to_number") == fields.index("to") + 1
    assert normalize.accepted_fields(
        public_ops()["get_text"].args.model_fields, op="get_text"
    )[1:4] == ["scope", "slide_number", "slide_numbers"]


def test_set_footer_keeps_its_own_slide_number(tmp_path):
    """set_footer's slide_number means "show the slide number"; its scope
    takes no twin, so that argument keeps its meaning."""
    fields = public_ops()["set_footer"].args.model_fields
    assert "slide_number" in fields
    assert all(field != "scope" for field, _t, _k in normalize.twins_for("set_footer", fields))
    PptxPackage  # noqa: B018  (imported for symmetry with the other twin tests)


#: Every argument that counts slides from 0, by operation (the read
#: operations' scope included), checked over all 145 operations' arguments
#: for 2.0.0rc8. set_footer's scope has no twin: its own slide_number turns
#: slide numbers on, and the changelog names that exception.
COUNTS_FROM_0 = {
    "copy_position": ("from_slide", "to_slides"),
    "copy_shape_between": ("source_slide",),
    "copy_slide_between": ("position",),
    "duplicate_slide": ("position",),
    "export_slide_image": ("slides",),
    "generate_agenda_slide": ("position",),
    "insert_slide": ("position",),
    "manage_custom_show": ("slides",),
    "move_slide": ("to",),
    "reorder_slides": ("order",),
    "set_hyperlink": ("to_slide",),
    "split_deck": ("ranges",),
    **{op: ("scope",) for op in (
        "get_presentation_view", "get_text", "find_text", "list_elements",
        "list_hyperlinks", "list_equations", "list_comments", "audit_accessibility",
        "search_and_replace")},
}


def test_every_argument_that_counts_slides_from_0_has_its_twin():
    """The changelog's "Every argument that counts slides from 0 has
    a twin counted from 1" was false for position on duplicate_slide,
    copy_slide_between and generate_agenda_slide. Each argument above has
    its twin, and the one exception is named where the claim is made."""
    from pathlib import Path

    ops = public_ops()
    for op, fields in COUNTS_FROM_0.items():
        have = {field for field, _twin, _kind in
                normalize.twins_for(op, list(ops[op].args.model_fields))}
        for field in fields:
            assert field in have, (op, field)
    assert normalize.twins_for("set_footer", list(ops["set_footer"].args.model_fields)) == []
    root = Path(__file__).resolve().parents[2]
    for surface in ("CHANGELOG.md", "docs/llms.txt"):
        text = (root / surface).read_text(encoding="utf-8")
        assert "set_footer" in text and "turns slide numbers on" in text, surface
