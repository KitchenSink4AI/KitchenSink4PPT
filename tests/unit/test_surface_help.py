"""Generated help schema and example validation."""

from __future__ import annotations

import pytest

from kitchensink4ppt import server  # noqa: F401 - populates the operation registry
from kitchensink4ppt.registry import OpSpec, StrictArgs, public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface.help import args_schema, example, help_payload, validate_examples
from kitchensink4ppt.surface.plan import build


class ShapeArgs(StrictArgs):
    file_path: str
    slide: int
    labels: list[str]
    options: dict[str, int] | None = None


def _spec() -> OpSpec:
    return OpSpec(
        "get_text", "ppt_read", "Read slide text.", ShapeArgs,
        lambda **kwargs: kwargs, False, True, False, False,
    )


def test_help_schema_is_strict_and_example_validates():
    spec = _spec()
    schema = args_schema(spec)
    assert schema["additionalProperties"] is False
    # A required slide may also be sent as slide_number, so the schema
    # requires one of the two rather than slide itself. Sent on its own,
    # slide_number must be a number: a null names no slide.
    assert schema["required"] == ["file_path", "labels"]
    assert schema["anyOf"] == [
        {"required": ["slide"]},
        {"required": ["slide_number"],
         "properties": {"slide_number": {"type": "integer"}}},
    ]
    sample = example(spec)
    assert ShapeArgs.model_validate(sample)
    assert sample["file_path"]
    assert isinstance(sample["labels"], list)


def test_all_example_validation_scaffolding_returns_one_per_op():
    spec = _spec()
    assert validate_examples({"get_text": spec}) == {"get_text": example(spec)}


def test_every_released_operation_has_a_schema_valid_help_example():
    samples = validate_examples(public_ops())
    assert len(samples) == 145
    for name, sample in samples.items():
        public_ops()[name].args.model_validate(sample)


def test_help_without_op_lists_visible_surface_and_with_op_returns_contract():
    spec = _spec()
    plan = build(Settings(), {"get_text": spec})
    listing = help_payload(plan, None)
    assert listing["tools"] == {"ppt_read": ["get_text"]}

    detail = help_payload(plan, "get_text")
    assert detail["tool"] == "ppt_read"
    # The served summary is the writer's OP_LINES line when one exists
    # (get_text has one), else the spec's own summary.
    from kitchensink4ppt.surface import copy as _copy
    assert detail["summary"] == (_copy.OP_LINES.get("get_text") or "Read slide text.")
    assert detail["read_only"] is True
    assert detail["writes"] is False
    assert detail["example"]["op"] == "get_text"
    ShapeArgs.model_validate(detail["example"]["args"])


def test_help_hides_forbidden_ops():
    spec = _spec()
    plan = build(Settings(forbid=("get_text",)), {"get_text": spec})
    with pytest.raises(Exception) as caught:
        help_payload(plan, "get_text")
    assert getattr(caught.value, "code", None) == "NOT_AVAILABLE"
