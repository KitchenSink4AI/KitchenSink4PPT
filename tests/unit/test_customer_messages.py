"""Customer messages stay copy-reachable and name real operations."""

from __future__ import annotations

import ast
import re
from pathlib import Path

import pytest

from kitchensink4ppt import server  # noqa: F401 - populate the registry
from kitchensink4ppt.core import budget, deadline
from kitchensink4ppt.core.errors import OpError
from kitchensink4ppt.core import messages
from kitchensink4ppt.registry import public_ops
from kitchensink4ppt.surface import copy
from kitchensink4ppt.surface import envelope
from kitchensink4ppt.surface.dispatch import resolve, validate
from kitchensink4ppt.surface.plan import build
from kitchensink4ppt.settings import Settings


ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / "src" / "kitchensink4ppt"

_OP_PREFIXES = (
    "add_", "align_", "anonymize_", "apply_", "audit_", "check_",
    "clear_", "comment_", "compare_", "compress_", "copy_", "create_",
    "deck_", "delete_", "diagnose_", "distribute_", "duplicate_",
    "export_", "extract_", "find_", "fit_", "font_", "format_",
    "generate_", "get_", "group_", "import_", "insert_", "list_",
    "live_", "manage_", "merge_", "move_", "powerpoint_", "refresh_",
    "remove_", "reorder_", "replace_", "reply_", "resolve_", "search_",
    "set_", "split_", "svg_", "ungroup_", "unmerge_", "update_",
    "validate_", "zombie_",
)
_TOKEN = re.compile(r"\b[a-z][a-z0-9]*(?:_[a-z0-9]+)+\b")
_RESULT_FIELDS = frozenset({
    "caveat", "caveats", "fix", "hint", "message", "note", "usage",
    "warning", "warnings",
})
_SERVED_NON_OPERATIONS = frozenset({
    "check_interval_hours",  # update-policy field
    "comment_id",            # result/argument field
    "move_slide_into",       # manage_section action
})


def _strings(node: ast.AST):
    for child in ast.walk(node):
        if isinstance(child, ast.Constant) and isinstance(child.value, str):
            yield child.value


def _served_expressions(tree: ast.AST):
    """Yield expressions that can reach an MCP error or result message."""
    for node in ast.walk(tree):
        if isinstance(node, ast.Raise) and node.exc is not None:
            yield node.exc
        elif isinstance(node, ast.Dict):
            for key, value in zip(node.keys, node.values):
                if (
                    isinstance(key, ast.Constant)
                    and key.value in _RESULT_FIELDS
                ):
                    yield value
        elif isinstance(node, ast.Call):
            name = (
                node.func.attr if isinstance(node.func, ast.Attribute)
                else node.func.id if isinstance(node.func, ast.Name)
                else ""
            )
            if name == "finding":
                yield from node.args[2:4]


def test_served_operation_references_exist_in_registry():
    from kitchensink4ppt.ops import batch

    # The edits apply_edits batches (set_text, set_paragraph, ...) are named
    # in served strings too, and ppt_help answers for each of them.
    known = set(public_ops()) | set(batch._OPS)
    bad: list[str] = []
    for path in SOURCE.rglob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for expression in _served_expressions(tree):
            for value in _strings(expression):
                for token in _TOKEN.findall(value):
                    if (
                        token.startswith(_OP_PREFIXES)
                        and token not in known
                        and token not in _SERVED_NON_OPERATIONS
                    ):
                        bad.append(
                            f"{path.relative_to(ROOT)}:{getattr(expression, 'lineno', '?')}: {token}"
                        )
    assert not bad, "unknown operation references in served strings:\n" + "\n".join(bad)


def test_ops_layer_does_not_import_surface_copy():
    offenders = []
    for name in ("access.py", "design_check.py"):
        path = SOURCE / "ops" / name
        text = path.read_text(encoding="utf-8")
        if "surface.copy" in text or "surface import copy" in text:
            offenders.append(str(path.relative_to(ROOT)))
    assert offenders == []


def test_engine_copy_is_reexported_and_ship_gate_reachable():
    assert copy.MISSING_TITLE is messages.MISSING_TITLE
    assert copy.TITLE_READ_AFTER_BODY is messages.TITLE_READ_AFTER_BODY
    inventory = copy.all_strings()
    for value in messages.ENGINE_COPY:
        assert value in inventory


def test_named_surface_refusals_use_the_writer_reachable_templates():
    plan = build(Settings(), public_ops())

    with pytest.raises(OpError) as missing:
        resolve(plan, "ppt_read", "not_a_real_operation")
    assert missing.value.hint == messages.OP_NOT_FOUND_HINT.format(
        op="not_a_real_operation"
    )

    with pytest.raises(OpError) as misplaced:
        resolve(plan, "ppt_read", "insert_slide")
    assert misplaced.value.message == messages.WRONG_TOOL_MESSAGE.format(
        op="insert_slide", owner="ppt_edit", tool="ppt_read"
    )

    spec = public_ops()["get_text"]
    with pytest.raises(OpError) as invalid:
        validate(spec, {"unknown": True})
    assert invalid.value.message == messages.BAD_ARGS_MESSAGE.format(
        op="get_text"
    )
    assert invalid.value.hint == messages.BAD_ARGS_HINT.format(op="get_text")


def test_deadline_and_limit_refusals_use_engine_copy_templates():
    timed_out = deadline.Deadline(1).timeout_error("the operation")
    assert timed_out.message == messages.TIMEOUT_MESSAGE.format(
        what="the operation", bound="1s"
    )

    with pytest.raises(OpError) as too_large:
        budget.enforce(
            {"nested": {"number": 1}},
            1,
            wrap=envelope.success,
            supports_offset=False,
        )
    assert too_large.value.message == messages.LIMIT_EXCEEDED_MESSAGE
    assert too_large.value.hint == messages.LIMIT_EXCEEDED_HINT.format(
        paging="This operation does not support offset paging."
    )
