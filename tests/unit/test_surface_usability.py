"""Focused regressions for surface usability."""

from __future__ import annotations

import asyncio
import json
from pathlib import Path
from typing import Literal

import pytest

from kitchensink4ppt import server
from kitchensink4ppt.ops import comments, workflows
from kitchensink4ppt.registry import OpSpec, StrictArgs, public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface import copy, dispatch as surface_dispatch, envelope
from kitchensink4ppt.surface.help import help_payload
from kitchensink4ppt.surface.plan import build


ROOT = Path(__file__).resolve().parents[2]


class ReadArgs(StrictArgs):
    file_path: str
    detail: Literal["outline", "text", "full"] = "text"


def _read_spec(handler) -> OpSpec:
    return OpSpec(
        "get_text", "ppt_read", "Read presentation text.", ReadArgs, handler,
        False, True, False, False,
    )


def _call(spec: OpSpec, arguments: dict):
    plan = build(Settings(), {spec.name: spec})
    return asyncio.run(surface_dispatch.dispatch(plan, "ppt_read", arguments))


@pytest.mark.parametrize("exc", [
    AttributeError("missing member"),
    KeyError("missing key"),
    IndexError("missing item"),
])
def test_e3_unexpected_engine_shape_failures_are_internal(exc):
    payload = envelope.refusal(exc)
    assert payload["error"]["code"] == "INTERNAL_ERROR"
    assert payload["error"]["message"] == copy.INTERNAL_ERROR_MESSAGE
    assert str(exc) not in json.dumps(payload)


def test_c1_folded_operation_name_routes_and_echoes_original():
    spec = _read_spec(lambda **values: values)
    for sent in ("GetText", "GET_TEXT", "get-text", "get text"):
        payload, is_error = _call(spec, {
            "op": sent,
            "args": {"file_path": "deck.pptx"},
        })
        assert is_error is False, payload
        assert payload["normalized_from"]["op"] == sent


def test_c1_folded_operation_keys_are_collision_free():
    from kitchensink4ppt.surface import normalize

    normalize.op_keys(public_ops())


def test_c2_flat_args_and_file_synonyms_are_safe_and_echoed():
    spec = _read_spec(lambda **values: values)
    payload, is_error = _call(spec, {
        "op": "get_text", "file": "deck.pptx", "detail": "text",
    })
    assert is_error is False, payload
    assert set(payload["normalized_from"]["flat_args"]) == {"file", "detail"}
    assert payload["normalized_from"]["args"]["file"] == "deck.pptx"

    ran = []
    guarded = _read_spec(lambda **values: ran.append(values) or values)
    payload, is_error = _call(guarded, {
        "op": "get_text", "file_path": "deck.pptx", "unknown": True,
    })
    assert is_error is True
    assert payload["error"]["code"] == "BAD_ARGS"
    assert ran == []


def test_e2_enum_values_ignore_case_and_separators_but_unknown_values_refuse():
    spec = _read_spec(lambda **values: values)
    payload, is_error = _call(spec, {
        "op": "get_text",
        "args": {"file_path": "deck.pptx", "detail": "FULL"},
    })
    assert is_error is False, payload
    assert payload["detail"] == "full"
    assert payload["normalized_from"]["args"]["detail"] == "FULL"

    payload, is_error = _call(spec, {
        "op": "get_text",
        "args": {"file_path": "deck.pptx", "detail": "complete"},
    })
    assert is_error is True
    assert payload["error"]["code"] == "BAD_ARGS"


def test_bad_args_names_valid_and_required_fields():
    spec = _read_spec(lambda **values: values)
    payload, is_error = _call(spec, {
        "op": "get_text", "args": {"unknown": True},
    })
    assert is_error is True
    details = payload["error"]["details"]
    assert details["valid_fields"] == ["file_path", "detail"]
    assert details["required_fields"] == ["file_path"]


def test_customer_argument_aliases_normalize_without_expanding_schema():
    from kitchensink4ppt.surface import normalize

    fields = public_ops()["set_theme_fonts"].args.model_fields
    args, sent = normalize.normalize_args(
        {"file_path": "deck.pptx", "heading": "Georgia", "body": "Arial"},
        fields,
        op="set_theme_fonts",
    )
    assert args["major"] == "Georgia"
    assert args["minor"] == "Arial"
    assert sent == {"heading": "Georgia", "body": "Arial"}

    create_fields = public_ops()["create_presentation"].args.model_fields
    args, sent = normalize.normalize_args(
        {"file_path": "new.potx"}, create_fields, op="create_presentation"
    )
    assert args == {"path": "new.potx"}
    assert sent == {"file_path": "new.potx"}


def test_c3_unknown_operation_suggests_without_running_any_handler():
    ran = []
    spec = _read_spec(lambda **values: ran.append(values) or values)
    payload, is_error = _call(spec, {
        "op": "get_txt", "args": {"file_path": "deck.pptx"},
    })
    assert is_error is True
    assert payload["error"]["code"] == "NOT_FOUND"
    # The suggestion carries the served operation line: the writer's OP_LINES
    # entry when one is filled, else the spec summary.
    assert payload["error"]["details"]["did_you_mean"] == [{
        "op": "get_text", "tool": "ppt_read",
        "purpose": copy.OP_LINES.get("get_text") or "Read presentation text.",
    }]
    assert payload["error"]["message"] == copy.OP_NOT_FOUND_MESSAGE.format(
        op="get_txt"
    )
    assert ran == []


def test_help_uses_the_same_folded_routing_and_nonexecuting_suggestions():
    spec = _read_spec(lambda **values: values)
    plan = build(Settings(), {spec.name: spec})
    result = help_payload(plan, "Get-Text")
    assert result["op"] == "get_text"
    assert result["normalized_from"] == {"op": "Get-Text"}
    with pytest.raises(Exception) as caught:
        help_payload(plan, "get_txt")
    assert caught.value.code == "NOT_FOUND"
    assert caught.value.details["did_you_mean"][0]["op"] == "get_text"


def test_b7_migrate_from_v1_covers_every_released_tool():
    out = workflows.get_workflows("migrate-from-v1")
    released = json.loads(
        (ROOT / "migration" / "released_tools_1.3.1.json").read_text("utf-8")
    )["tools"]
    assert {row["old"] for row in out["map"]} == set(released)
    for row in out["map"]:
        if row["op"] is not None:
            assert row["op"] in public_ops()
            assert row["tool"] == public_ops()[row["op"]].group
    assert "v1 has no partial-apply mode" not in (server.apply_edits.__doc__ or "")
    from kitchensink4ppt.ops import batch
    with pytest.raises(Exception) as caught:
        batch.apply_edits(None, [], atomic=False)
    assert "v1" not in str(caught.value).casefold()


def test_b8_comment_initials_use_letters_only():
    assert comments._initials("Claude (Reviewer)") == "CR"
    assert comments._initials("Jane Q. Reviewer") == "JQR"


def test_e5_view_names_what_non_full_details_hide(make_deck):
    text = server.get_presentation_view(make_deck(), detail="text")["view"]
    assert copy.VIEW_TEXT_OMISSIONS in text
    outline = server.get_presentation_view(make_deck(), detail="outline")["view"]
    assert copy.VIEW_OUTLINE_OMISSIONS in outline
    full = server.get_presentation_view(make_deck(), detail="full")["view"]
    assert copy.VIEW_TEXT_OMISSIONS not in full
    assert copy.VIEW_OUTLINE_OMISSIONS not in full


def test_e7_server_info_reports_the_active_surface_token_bill():
    spec = OpSpec(
        "get_server_info", "ppt_read", "Server facts.",
        type("NoArgs", (StrictArgs,), {}), lambda: {},
        False, True, False, False,
    )
    plan = build(Settings(), {spec.name: spec})
    payload, is_error = asyncio.run(surface_dispatch.dispatch(
        plan, "ppt_read", {"op": "get_server_info", "args": {}}
    ))
    assert is_error is False, payload
    assert payload["surface"]["approx_active_tokens"] == (
        surface_dispatch.approx_active_tokens(plan)
    )
    assert payload["surface"]["approx_active_tokens"] > 0
    direct = server.get_server_info()["surface"]["approx_active_tokens"]
    default_plan = build(Settings(), public_ops())
    assert direct == surface_dispatch.approx_active_tokens(default_plan)


def test_list_elements_reads_type_as_kind_and_names_the_argument(make_deck):
    """Callers reach for `type`; the argument is `kind`. The call goes
    through, and the answer names `kind` so the next call uses it."""
    deck = str(make_deck())
    plan = build(Settings(), public_ops())
    payload, is_error = asyncio.run(surface_dispatch.dispatch(plan, "ppt_read", {
        "op": "list_elements", "args": {"file_path": deck, "type": "slides"}}))
    assert is_error is False, payload
    assert payload["count"] >= 1
    assert payload["normalized_from"]["args"] == {"type": "slides"}
    assert "`kind`" in payload["normalized_from"]["hint"]

    payload, is_error = asyncio.run(surface_dispatch.dispatch(plan, "ppt_read", {
        "op": "list_elements", "args": {"file_path": deck, "kind": "slides"}}))
    assert is_error is False, payload
    assert "normalized_from" not in payload

    payload, is_error = asyncio.run(surface_dispatch.dispatch(plan, "ppt_read", {
        "op": "list_elements",
        "args": {"file_path": deck, "kind": "slides", "type": "shapes"}}))
    assert is_error is True, "two spellings of one argument are not guessed between"
