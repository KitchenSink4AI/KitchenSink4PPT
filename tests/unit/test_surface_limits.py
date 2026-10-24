"""B2z description and compact-schema ceilings for PowerPoint."""

from __future__ import annotations

import asyncio

import pytest

from kitchensink4ppt import server  # noqa: F401 - populates the registry
from kitchensink4ppt.registry import GROUP_OPS, OpSpec, StrictArgs, public_ops
from kitchensink4ppt.core import budget
from kitchensink4ppt.settings import DEFAULT_MAX_OUTPUT_CHARS, Settings
from kitchensink4ppt.surface import copy, envelope
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.help import example
from kitchensink4ppt.surface.plan import (
    MAX_DESCRIPTION_CHARS,
    MAX_SCHEMA_BYTES,
    PlanError,
    build,
    schema_bytes,
)


class EmptyArgs(StrictArgs):
    pass


def _catalog(summary_size: int = 18) -> dict[str, OpSpec]:
    result: dict[str, OpSpec] = {}
    summary = "S" * summary_size
    for group, names in GROUP_OPS.items():
        read_only = group in {"ppt_read", "ppt_inspect"}
        for name in names:
            result[name] = OpSpec(
                name, group, summary, EmptyArgs, lambda **_kwargs: {},
                not read_only, read_only, False, False,
            )
    return result


@pytest.mark.parametrize(
    "settings",
    [
        Settings(),
        Settings(read_only=True),
        Settings(forbid=("ppt_graphics", "get_text")),
    ],
)
def test_every_plan_stays_under_both_limits(settings):
    plan = build(settings, _catalog())
    for tool in plan.tools.values():
        assert len(tool.description) <= MAX_DESCRIPTION_CHARS
        assert schema_bytes(tool.input_schema) <= MAX_SCHEMA_BYTES


def test_overflow_moves_only_whole_lines_and_each_line_appears_once():
    catalog = _catalog(summary_size=180)
    plan = build(Settings(), catalog)
    overflowed = 0
    for group, tool in plan.tools.items():
        if group == "ppt_help":
            continue
        op_description = tool.input_schema["properties"]["op"]["description"]
        combined = tool.description + "\n" + op_description
        for op in tool.ops:
            line = copy.OP_LINES.get(op) or catalog[op].summary
            assert combined.count(f"- {op}: {line}") == 1
        if copy.OPS_CONTINUED in tool.description:
            overflowed += 1
            assert op_description.startswith(copy.OP_PARAM_CONTINUED)
        else:
            assert op_description == copy.OP_PARAM
    assert overflowed > 0


def test_schema_over_limit_refuses_plan_instead_of_silently_truncating(monkeypatch):
    monkeypatch.setattr(copy, "ARGS_PARAM", "X" * MAX_SCHEMA_BYTES)
    with pytest.raises(PlanError, match="schema exceeds"):
        build(Settings(), _catalog())


def test_every_operation_passes_through_the_final_output_cap():
    cap = 2_000
    specs = public_ops()
    plan = build(Settings(max_output_chars=cap), specs)

    async def sweep():
        for name, spec in specs.items():
            payload, is_error = await dispatch(
                plan,
                plan.op_tool[name],
                {"op": name, "args": example(spec)},
                invoke=lambda _spec, _args: {"blob": "x" * (cap * 3)},
            )
            assert is_error is False, name
            assert len(envelope.compact_json(payload)) <= cap, name
            assert payload["page"]["truncated"] is True, name

    asyncio.run(sweep())


def test_dispatch_preserves_a_negative_handler_verdict():
    specs = public_ops()
    plan = build(Settings(), specs)

    async def call_validate():
        return await dispatch(
            plan,
            plan.op_tool["validate"],
            {"op": "validate", "args": {"file_path": "unused.pptx"}},
            invoke=lambda _spec, _args: {
                "ok": False,
                "payload_valid": False,
                "powerpoint": {"opens_clean": False},
            },
        )

    payload, is_error = asyncio.run(call_validate())
    assert is_error is False
    assert payload["ok"] is False


def test_output_budget_has_one_default():
    assert budget.DEFAULT_MAX_CHARS == DEFAULT_MAX_OUTPUT_CHARS


def test_final_cap_repaging_keeps_the_engine_page_truth():
    result = {
        "slides": [f"slide-{index}:" + ("x" * 1_000) for index in range(29)],
        "page": {
            "unit": "slides",
            "total": 70,
            "offset": 0,
            "returned": 29,
            "omitted": 41,
            "truncated": True,
            "reason": "budget",
            "next_offset": 29,
            "hint": "41 more slides were not returned; call again with offset=29.",
        },
    }
    capped = budget.enforce(
        result,
        5_000,
        wrap=envelope.success,
        supports_offset=True,
        requested_offset=0,
    )

    returned = len(capped["slides"])
    assert 0 < returned < 29
    assert capped["page"]["total"] == 70
    assert capped["page"]["offset"] == 0
    assert capped["page"]["returned"] == returned
    assert capped["page"]["omitted"] == 70 - returned
    assert capped["page"]["next_offset"] == returned
    assert f"offset={returned}" in capped["page"]["hint"]
    assert "offset=29" not in capped["page"]["hint"]
