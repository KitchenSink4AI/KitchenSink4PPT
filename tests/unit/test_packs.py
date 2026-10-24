"""Registry and planning guarantees for the fixed PowerPoint surface.
This replaces the retired pack registry and runtime toggle suite. The
operation catalog is still tested exhaustively, but its public MCP projection
is now immutable for the life of a server process.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from pydantic import ValidationError

from kitchensink4ppt import server
from kitchensink4ppt.registry import GROUP_OPS, OP_TO_GROUP, public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface.plan import ALL_TOOLS, build


ROOT = Path(__file__).resolve().parents[2]


def test_retired_pack_module_and_toggle_operations_are_absent():
    assert not (ROOT / "src" / "kitchensink4ppt" / "packs.py").exists()
    assert not hasattr(server, "enable_tools")
    assert not hasattr(server, "disable_tools")


def test_registry_contains_each_released_operation_once():
    registry = public_ops()
    expected = {name for names in GROUP_OPS.values() for name in names}
    assert set(registry) == expected
    assert len(registry) == len(OP_TO_GROUP) == 145
    assert sum(len(names) for names in GROUP_OPS.values()) == 145


def test_registry_handlers_are_the_exported_server_functions():
    for name, spec in public_ops().items():
        assert spec.handler is getattr(server, name)
        assert spec.name == name
        assert spec.group == OP_TO_GROUP[name]
        assert spec.summary
        assert "\n" not in spec.summary


def test_every_argument_model_forbids_unknown_fields():
    for name, spec in public_ops().items():
        with pytest.raises(ValidationError):
            spec.args.model_validate({"__unknown_surface_argument__": True})


def test_default_plan_is_exactly_eight_groups_plus_help():
    plan = build(Settings())
    assert tuple(plan.tools) == ALL_TOOLS
    assert len(plan.tools) == 9
    listed = [op for tool in plan.tools.values() if tool.name != "ppt_help"
              for op in tool.ops]
    assert len(listed) == 145
    assert len(set(listed)) == 145
    assert set(listed) == set(public_ops())


def test_group_membership_matches_the_adopted_map():
    plan = build(Settings())
    for group, expected in GROUP_OPS.items():
        assert plan.tools[group].ops == expected
        for op in expected:
            assert plan.op_tool[op] == group


def test_font_inventory_is_a_read_only_inspection_operation():
    """Font discovery belongs beside audits, not design mutations."""
    plan = build(Settings())
    assert OP_TO_GROUP["font_inventory"] == "ppt_inspect"
    assert "font_inventory" in plan.tools["ppt_inspect"].ops
    assert "font_inventory" not in plan.tools["ppt_design"].ops
    assert plan.tools["ppt_inspect"].annotations["readOnlyHint"] is True


def test_fixed_plan_has_no_runtime_mutation_api():
    plan = build(Settings())
    assert not hasattr(plan, "enable")
    assert not hasattr(plan, "disable")
    assert not hasattr(plan, "toggle")
    before = tuple(plan.tools)
    assert tuple(plan.tools) == before == ALL_TOOLS


def test_plans_are_independent_immutable_startup_views():
    default = build(Settings())
    restricted = build(Settings(forbid=("ppt_graphics", "get_text")))
    assert tuple(default.tools) == ALL_TOOLS
    assert "ppt_graphics" not in restricted.tools
    assert "get_text" not in restricted.tools["ppt_read"].ops
    assert "ppt_graphics" in default.tools
    assert "get_text" in default.tools["ppt_read"].ops


def test_group_annotations_match_read_write_boundary():
    plan = build(Settings())
    for name in ("ppt_read", "ppt_inspect", "ppt_help"):
        assert plan.tools[name].annotations["readOnlyHint"] is True
        assert plan.tools[name].annotations["destructiveHint"] is False
    for name in (
        "ppt_edit", "ppt_graphics", "ppt_tables", "ppt_design",
        "ppt_review", "ppt_deck",
    ):
        assert plan.tools[name].annotations["readOnlyHint"] is False
