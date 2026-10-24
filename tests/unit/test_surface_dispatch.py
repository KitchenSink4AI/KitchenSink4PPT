"""Strict dispatch and JSON-text compatibility tests."""

from __future__ import annotations

import asyncio
from typing import Any

import pytest

from kitchensink4ppt.core import errors as core_errors
from kitchensink4ppt.registry import OpSpec, StrictArgs
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface import copy, envelope
from kitchensink4ppt.surface.dispatch import (
    _admits, approx_active_tokens, decode_json_text, dispatch, parse, validate,
)
from kitchensink4ppt.surface.plan import build


class DecodeArgs(StrictArgs):
    items: list[int]
    config: dict[str, int]
    text: str
    either: str | list[int]


class EmptyArgs(StrictArgs):
    pass


def _spec(handler=lambda **kwargs: kwargs) -> OpSpec:
    return OpSpec(
        "get_text", "ppt_read", "Read text.", DecodeArgs, handler,
        False, True, False, False,
    )


def test_parse_requires_exact_top_level_shape():
    assert parse({"op": "get_text"}) == ("get_text", {})
    with pytest.raises(Exception):
        parse({"op": "get_text", "file_path": "a.pptx"})
    with pytest.raises(Exception):
        parse({"op": "get_text", "args": []})


def test_json_text_decodes_only_required_object_and_array_fields():
    output, decoded = decode_json_text(_spec(), {
        "items": "[1,2]",
        "config": '{"x":3}',
        "text": "[1,2]",
        "either": "[1,2]",
    })
    assert output == {
        "items": [1, 2],
        "config": {"x": 3},
        "text": "[1,2]",
        "either": "[1,2]",
    }
    assert decoded == ["items", "config"]


def test_malformed_or_wrong_kind_json_is_left_for_normal_validation():
    output, decoded = decode_json_text(_spec(), {
        "items": "[broken",
        "config": "[1,2]",
        "text": "1e3",
        "either": "text",
    })
    assert output["items"] == "[broken"
    assert output["config"] == "[1,2]"
    assert output["text"] == "1e3"
    assert decoded == []


def test_schema_kind_detection_follows_one_of_and_all_of():
    root = {
        "$defs": {"items": {"type": "array", "items": {"type": "integer"}}},
        "one": {"oneOf": [{"$ref": "#/$defs/items"}, {"type": "object"}]},
        "all": {"allOf": [{"type": "array"}, {"$ref": "#/$defs/items"}]},
    }
    assert _admits(root["one"], root) == {"array", "object"}
    assert _admits(root["all"], root) == {"array"}


def test_schema_kind_walk_covers_refs_allof_and_oneof():
    root = {"$defs": {"ObjectArgs": {"type": "object"}}}
    assert _admits({"allOf": [{"$ref": "#/$defs/ObjectArgs"}]}, root) == {"object"}
    assert _admits(
        {"oneOf": [{"type": "array"}, {"type": "string"}]}, root
    ) == {"array", "string"}


def test_validation_is_strict_and_does_not_echo_bad_values():
    with pytest.raises(Exception) as caught:
        validate(_spec(), {
            "items": [], "config": {}, "text": "safe", "either": "safe",
            "secret": "do-not-echo",
        })
    assert "do-not-echo" not in str(caught.value)


def test_dispatch_returns_flat_success_and_records_json_text_decoding():
    spec = _spec()
    plan = build(Settings(), {"get_text": spec})
    payload, is_error = asyncio.run(dispatch(plan, "ppt_read", {
        "op": "get_text",
        "args": {
            "items": "[1,2]", "config": '{"x":3}',
            "text": "literal", "either": "literal",
        },
    }))
    assert is_error is False
    assert payload["ok"] is True
    assert payload["items"] == [1, 2]
    assert payload["json_text_decoded"] == ["items", "config"]
    assert "result" not in payload


def test_forbidden_and_wrong_tool_calls_refuse_with_actionable_codes():
    spec = _spec()
    forbidden = build(Settings(forbid=("get_text",)), {"get_text": spec})
    payload, is_error = asyncio.run(dispatch(
        forbidden, "ppt_read", {"op": "get_text", "args": {}}
    ))
    assert is_error is True
    assert payload["error"]["code"] == "NOT_AVAILABLE"
    assert "decided_by" in payload["error"]["details"]

    normal = build(Settings(), {"get_text": spec})
    payload, is_error = asyncio.run(dispatch(
        normal, "ppt_inspect", {"op": "get_text", "args": {}}
    ))
    assert is_error is True
    assert payload["error"]["code"] == "BAD_ARGS"


def test_existing_engine_exceptions_keep_their_closed_refusal_class():
    payload = envelope.refusal(core_errors.DocumentNotFound("missing deck"))
    assert payload["ok"] is False
    assert payload["error"]["code"] == "NOT_FOUND"
    assert payload["error"]["message"] == "missing deck"
    assert payload["error"]["hint"] == copy.HINTS["NOT_FOUND"]
    assert set(copy.HINTS) == set(core_errors.CODES)


def test_status_result_reports_the_plan_that_handled_the_call():
    spec = OpSpec(
        "get_server_info", "ppt_read", "Server facts.", EmptyArgs,
        lambda: {
            "tools_registered": 999,
            "operations_registered": 999,
            "surface": {"fixed": False},
        },
        False, True, False, False,
    )
    settings = Settings(read_only=True, ignored_legacy=("KS4P_MODE",))
    plan = build(settings, {"get_server_info": spec})

    payload, is_error = asyncio.run(dispatch(
        plan, "ppt_read", {"op": "get_server_info", "args": {}}
    ))

    assert is_error is False
    assert payload["tools_registered"] == 2  # ppt_read + ppt_help
    assert payload["operations_registered"] == 1
    assert payload["surface"] == {
        "fixed": True,
        "tools": {"ppt_read": 1},
        "operations": 1,
        "read_only": True,
        "forbid": [],
        "legacy_settings_ignored": ["KS4P_MODE"],
        "approx_active_tokens": approx_active_tokens(plan),
    }
