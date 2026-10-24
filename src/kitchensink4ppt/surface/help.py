"""Registry-generated PowerPoint operation help and validated examples."""

from __future__ import annotations

import copy as _copy
from typing import Any, Mapping

from kitchensink4ppt.core.errors import OpError

from ..registry import OpSpec
from . import copy, normalize
from .plan import SurfacePlan


def args_schema(spec: OpSpec) -> dict[str, Any]:
    schema = _copy.deepcopy(spec.args.model_json_schema())
    schema.pop("examples", None)
    schema["additionalProperties"] = False
    properties = schema.get("properties") or {}
    if "slide" in properties and normalize.SLIDE_NUMBER not in properties:
        # slide_number is accepted wherever slide is; list it right after.
        ordered: dict[str, Any] = {}
        for key, prop in properties.items():
            ordered[key] = prop
            if key == "slide":
                prop["description"] = copy.SLIDE_PARAM
                ordered[normalize.SLIDE_NUMBER] = _copy.deepcopy(copy.SLIDE_NUMBER_PARAM)
        schema["properties"] = ordered
        required = schema.get("required") or []
        if "slide" in required:
            # A required slide may be named either way, so the schema requires
            # one of the two rather than slide itself: a client that checks
            # calls against this schema must not block the slide_number call
            # the server accepts. Both together still reach the server, which
            # refuses them when they name different slides. A slide_number
            # sent alone must be a number: null on its own names no slide.
            schema["required"] = [name for name in required if name != "slide"]
            choice = [
                {"required": ["slide"]},
                {
                    "required": [normalize.SLIDE_NUMBER],
                    "properties": {normalize.SLIDE_NUMBER: {"type": "integer"}},
                },
            ]
            if "anyOf" in schema:
                schema["allOf"] = [{"anyOf": schema.pop("anyOf")}, {"anyOf": choice}]
            else:
                schema["anyOf"] = choice
    twins = normalize.twins_for(spec.name, properties)
    if twins:
        # Each 1-based twin is accepted wherever its argument is and listed
        # right after it: slide_numbers after slides, to_number after to.
        ordered = {}
        for key, prop in schema["properties"].items():
            ordered[key] = prop
            for field, twin, kind in twins:
                if field == key:
                    ordered[twin] = twin_param(field, twin, kind)
        schema["properties"] = ordered
        for field in dict.fromkeys(field for field, _t, _k in twins):
            if field not in (schema.get("required") or []):
                continue
            # A required argument may be sent as its twin instead, so the
            # schema requires one of them, as it does for slide.
            schema["required"] = [name for name in schema["required"] if name != field]
            _require_one_of(schema, [{"required": [field]}, *(
                {"required": [twin],
                 "properties": {twin: {"type": "integer" if kind == "one" else "array"}}}
                for f, twin, kind in twins if f == field)])
    # Other arguments that count slides from 0, described where they are
    # chosen: an earlier version described only slide, and a model sent the
    # user's slide 3 as to_slide 3 (the fourth slide) with nothing in the
    # answer to show it.
    described = schema.get("properties") or {}
    for notes in (copy.ARG_NOTES, copy.ARG_HELP):
        for name, text in notes.get(spec.name, {}).items():
            if name in described:
                described[name]["description"] = text
    return schema


def twin_param(field: str, twin: str, kind: str) -> dict[str, Any]:
    """The schema ppt_help shows for one 1-based twin. Null is allowed, as
    for slide_number: the server reads a null twin as not sent."""
    if (field, twin) == ("slides", normalize.SLIDE_NUMBERS):
        return _copy.deepcopy(copy.SLIDE_NUMBERS_PARAM)
    prop: dict[str, Any] = {"title": twin.replace("_", " ").title()}
    if kind == "one":
        prop.update({"type": ["integer", "null"], "minimum": 1})
    elif kind == "list":
        prop.update({"type": ["array", "null"], "items": {"type": "integer", "minimum": 1}})
    elif kind == "scope":
        span = {"type": "object",
                "properties": {"start": {"type": "integer", "minimum": 1},
                               "end": {"type": "integer", "minimum": 1}},
                "required": ["start", "end"]}
        prop["anyOf"] = [
            {"type": "array", "items": {"anyOf": [{"type": "integer", "minimum": 1}, span]}},
            span,
            {"type": "null"},
        ]
    else:
        prop.update({"type": ["array", "null"], "items": {
            "type": "object",
            "properties": {"start": {"type": "integer", "minimum": 1},
                           "end": {"type": "integer", "minimum": 1},
                           "name": {"type": "string"}},
            "required": ["start", "end"],
        }})
    prop["description"] = copy.TWIN_HELP[(field, twin)]
    return prop


def _require_one_of(schema: dict[str, Any], choice: list) -> None:
    """Require one of the alternatives in `choice`, beside whatever the
    schema already requires (a top-level anyOf stays the slide's choice)."""
    schema.setdefault("allOf", []).append({"anyOf": choice})


def _placeholder(prop: dict[str, Any], defs: dict[str, Any]) -> Any:
    if "$ref" in prop:
        return _placeholder(defs.get(prop["$ref"].rsplit("/", 1)[-1], {}), defs)
    if "enum" in prop:
        return prop["enum"][0]
    if "const" in prop:
        return prop["const"]
    for key in ("anyOf", "oneOf", "allOf"):
        if key in prop:
            options = [item for item in prop[key] if item.get("type") != "null"] or prop[key]
            return _placeholder(options[0], defs)
    kind = prop.get("type")
    if kind == "string":
        return "presentation.pptx" if "path" in str(prop).lower() else "text"
    if kind == "integer":
        return max(1, int(prop.get("minimum", 1)))
    if kind == "number":
        return max(1.0, float(prop.get("minimum", 1.0)))
    if kind == "boolean":
        return False
    if kind == "array":
        count = max(1, int(prop.get("minItems", 0)))
        return [_placeholder(prop.get("items", {}), defs) for _ in range(count)]
    if kind == "object" or "properties" in prop:
        properties = prop.get("properties", {})
        return {
            name: _placeholder(properties[name], defs)
            for name in prop.get("required", [])
        }
    return None


def example(spec: OpSpec) -> dict[str, Any]:
    from .examples import EXAMPLES

    extra = spec.args.model_config.get("json_schema_extra")
    if isinstance(extra, dict) and extra.get("examples"):
        candidate = dict(extra["examples"][0])
    else:
        schema = spec.args.model_json_schema()
        candidate = _placeholder(schema, schema.get("$defs", {}))
    # Tests and downstream integrations may construct a private OpSpec with
    # the same operation name but a deliberately smaller schema.  Apply only
    # example fields that this exact frozen model exposes.
    candidate.update({
        name: value
        for name, value in EXAMPLES.get(spec.name, {}).items()
        if name in spec.args.model_fields
    })
    spec.args.model_validate(candidate)
    return candidate


def validate_examples(specs: Mapping[str, OpSpec]) -> dict[str, dict[str, Any]]:
    """Synthesize and validate one example for every visible operation."""
    return {name: example(spec) for name, spec in specs.items() if not spec.internal}


def help_payload(plan: SurfacePlan, op: str | None) -> dict[str, Any]:
    if op is None or op == "":
        return {"usage": copy.HELP_USAGE, "tools": plan.visible_ops()}
    if not isinstance(op, str):
        raise OpError("BAD_ARGS", "op must be an operation name")
    sent_op = op
    canonical = normalize.route_op(op, plan.specs)
    if canonical is not None:
        op = canonical
    if (
        op in plan.forbidden_ops
        or op in plan.hidden_by_read_only
        or plan.op_tool.get(op) in plan.forbidden_tools
    ):
        raise OpError(
            "NOT_AVAILABLE", f"{op} is turned off by this server's launch settings",
            details={"op": op, "decided_by": copy.SETTINGS_OWNER},
        )
    spec = plan.specs.get(op)
    batch_only = _batch_only_help(plan, op) if spec is None else None
    if batch_only is not None:
        return batch_only
    if spec is None or not plan.is_visible(op):
        visible = {
            name: candidate
            for name, candidate in plan.specs.items()
            if plan.is_visible(name)
        }
        near = [
            {
                "op": name,
                "tool": plan.op_tool[name],
                "purpose": copy.OP_LINES.get(name) or visible[name].summary,
            }
            for name in normalize.suggestions(sent_op, visible)
        ]
        details: dict[str, Any] = {"op": sent_op}
        if near:
            details["did_you_mean"] = near
        raise OpError(
            "NOT_FOUND", copy.OP_NOT_FOUND_MESSAGE.format(op=sent_op),
            hint=copy.OP_NOT_FOUND_HINT.format(op=sent_op),
            details=details,
        )
    tool = plan.op_tool[op]
    sample = example(spec)
    batches = _batch_rows()
    result = {
        "op": op,
        "tool": tool,
        "summary": copy.OP_LINES.get(op) or spec.summary,
        "writes": spec.writes,
        "read_only": spec.read_only,
        "destructive": spec.destructive,
        "runs_live": spec.runs_live,
        "args_schema": args_schema(spec),
        "example": {"op": op, "args": sample},
        "call": f"{tool}(op={op!r}, args={{...}})",
    }
    if sent_op != op:
        result["normalized_from"] = {"op": sent_op}
    if op == "apply_edits":
        result["batchable_ops"] = batches
        edits = result["args_schema"].get("properties", {}).get("edits")
        if isinstance(edits, dict):
            edits["description"] = copy.BATCH_EDITS_PARAM.format(
                ops=", ".join(row["op"] for row in batches))
    elif any(row["op"] == op for row in batches):
        result["batchable_in"] = "apply_edits"
    return result


def _batch_rows() -> list[dict[str, Any]]:
    """apply_edits' batchable ops, from the batch's own table."""
    from ..ops import batch

    return batch.batchable_ops()


def _batch_only_help(plan: SurfacePlan, op: str) -> dict[str, Any] | None:
    """Help for an edit that only apply_edits runs (set_text), when
    apply_edits itself is available: what it is, its fields and an example
    batch. None for any other name."""
    if op in plan.specs or not plan.is_visible("apply_edits"):
        return None
    row = next((item for item in _batch_rows() if item["op"] == op), None)
    if row is None:
        return None
    tool = plan.op_tool["apply_edits"]
    return {
        "op": op,
        "tool": tool,
        "batch_only": True,
        "use": "apply_edits",
        "summary": copy.BATCH_ONLY_SUMMARY.format(op=op),
        "fields": row["fields"],
        "required": row["required"],
        "address": row["address"],
        "example": {"op": "apply_edits", "args": {
            "file_path": "presentation.pptx", "edits": [row["example"]]}},
        "call": f"{tool}(op='apply_edits', args={{...}})",
    }


__all__ = ["args_schema", "example", "help_payload", "validate_examples"]
