"""Foundation tests for registry generation and fixed surface planning."""

from __future__ import annotations

from typing import Any

import pytest
from pydantic import ValidationError

from kitchensink4ppt import server  # noqa: F401 - populate the released registry
from kitchensink4ppt.registry import (
    GROUP_OPS,
    OP_TO_GROUP,
    OpSpec,
    StrictArgs,
    args_model,
    first_docstring_sentence,
    public_ops,
)
from kitchensink4ppt.settings import Settings, SettingsError
from kitchensink4ppt.surface.plan import (
    ALL_TOOLS,
    POWERPOINT_APP_CAPABILITY,
    PlanError,
    build,
)
from kitchensink4ppt.surface.tools import install


class EmptyArgs(StrictArgs):
    pass


def _catalog() -> dict[str, OpSpec]:
    catalog: dict[str, OpSpec] = {}
    for group, names in GROUP_OPS.items():
        read_only = group in {"ppt_read", "ppt_inspect"}
        for name in names:
            catalog[name] = OpSpec(
                name=name,
                group=group,
                summary=f"{name} shipped summary.",
                args=EmptyArgs,
                handler=lambda **_kwargs: {},
                writes=not read_only,
                read_only=read_only,
                destructive=name == "delete_slide",
                runs_live=name.startswith("live_"),
            )
    return catalog


def test_adopted_map_has_exact_group_counts_and_145_unique_ops():
    assert {name: len(ops) for name, ops in GROUP_OPS.items()} == {
        "ppt_read": 18,
        "ppt_inspect": 17,
        "ppt_edit": 19,
        "ppt_graphics": 27,
        "ppt_tables": 16,
        "ppt_design": 19,
        "ppt_review": 7,
        "ppt_deck": 22,
    }
    assert len(OP_TO_GROUP) == 145
    assert set(OP_TO_GROUP) == {op for names in GROUP_OPS.values() for op in names}


def test_default_plan_has_eight_groups_plus_help_and_each_op_once():
    plan = build(Settings(), _catalog())
    assert tuple(plan.tools) == ALL_TOOLS
    listed = [op for tool, ops in plan.visible_ops().items() for op in ops]
    assert len(listed) == 145
    assert len(set(listed)) == 145
    assert set(listed) == set(OP_TO_GROUP)
    assert plan.tools["ppt_read"].annotations["readOnlyHint"] is True
    assert plan.tools["ppt_edit"].annotations["readOnlyHint"] is False
    assert plan.tools["ppt_edit"].annotations["destructiveHint"] is True


def test_read_only_and_forbid_are_frozen_into_the_plan():
    catalog = _catalog()
    read_only = build(Settings(read_only=True), catalog)
    assert tuple(read_only.tools) == ("ppt_read", "ppt_inspect", "ppt_help")
    assert len(read_only.hidden_by_read_only) == 110

    narrowed = build(
        Settings(forbid=("ppt_graphics", "get_text", "ppt_help")), catalog
    )
    assert "ppt_graphics" not in narrowed.tools
    assert "ppt_help" not in narrowed.tools
    assert "get_text" not in narrowed.tools["ppt_read"].ops
    assert narrowed.forbidden_ops == frozenset({"get_text"})


def test_unknown_forbid_name_and_bad_group_classification_refuse_startup():
    with pytest.raises(SettingsError, match="not a tool or op"):
        build(Settings(forbid=("not_real",)), _catalog())

    catalog = _catalog()
    original = catalog["get_text"]
    catalog["get_text"] = OpSpec(
        **{**original.__dict__, "group": "ppt_edit"}
    )
    with pytest.raises(PlanError, match="belongs to ppt_read"):
        build(Settings(), catalog)


def test_ppt_app_capability_forbids_every_live_or_com_route():
    catalog = public_ops()
    plan = build(Settings(forbid=(POWERPOINT_APP_CAPABILITY,)), catalog)
    expected = {
        name
        for name, spec in catalog.items()
        if spec.runs_live or spec.timeout_s is not None
    }
    assert plan.forbidden_ops == frozenset(expected)
    assert POWERPOINT_APP_CAPABILITY not in plan.forbidden_tools
    assert expected
    for name in expected:
        assert not plan.is_visible(name), name

    # Pin both inferred-live and explicit COM routes, plus file-only survival.
    for name in (
        "get_text", "insert_shape", "live_save", "powerpoint_status",
        "zombie_check", "export_pdf", "export_handout",
        "export_slide_image", "validate",
    ):
        assert name in expected
    assert plan.is_visible("extract_text")
    assert plan.is_visible("create_presentation")


def test_signature_model_is_strict_drops_context_and_preserves_defaults():
    class Context:
        pass

    def sample(file_path: str, values: list[int], enabled: bool = True, ctx: Context | None = None):
        """First shipped sentence. A later sentence is not the summary."""
        return {"file_path": file_path, "values": values, "enabled": enabled, "ctx": ctx}

    model = args_model(sample)
    assert set(model.model_fields) == {"file_path", "values", "enabled"}
    assert model.model_validate({"file_path": "a.pptx", "values": [1]}).enabled is True
    with pytest.raises(Exception, match="Extra inputs are not permitted"):
        model.model_validate({"file_path": "a.pptx", "values": [], "surprise": 1})
    assert first_docstring_sentence(sample) == "First shipped sentence."


@pytest.mark.parametrize("op_name", sorted(public_ops()))
def test_every_released_operation_refuses_an_unknown_argument(op_name):
    with pytest.raises(ValidationError, match="Extra inputs are not permitted"):
        public_ops()[op_name].args.model_validate({"__unknown_argument__": True})


def test_a_legacy_lock_refuses_startup_instead_of_widening_surface():
    env = {
        "KS4P_MODE": "full",
        "KS4P_PACK_POLICY": "locked",
        "KS4P_ALL_TOOLS": "1",
        "KS4P_LOCK_TOOLS": "1",
    }
    with pytest.raises(SettingsError, match="KS4P_PACK_POLICY.*KS4P_READ_ONLY"):
        Settings.from_env(env)


def test_install_is_the_one_explicit_registration_step_and_keeps_plan_identity():
    class StubMcp:
        def __init__(self):
            self.added: list[Any] = []

        def add_tool(self, tool):
            self.added.append(tool)

    plan = build(Settings(), _catalog())
    mcp = StubMcp()
    installed = install(mcp, plan)
    assert tuple(tool.name for tool in installed) == ALL_TOOLS
    assert tuple(mcp.added) == installed
    assert all(tool._plan is plan for tool in installed)
