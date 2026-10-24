"""The released-tool to family-surface map is complete and reproducible."""

from __future__ import annotations

import importlib.util
import json
from pathlib import Path

from kitchensink4ppt import registry, server  # noqa: F401
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface import plan


ROOT = Path(__file__).resolve().parents[2]
RELEASED = ROOT / "migration" / "released_tools_1.3.1.json"
MAP = ROOT / "migration" / "v1_to_v2.json"
TABLE = ROOT / "docs" / "surface_map.md"
REMOVED_REASON = "removed: every operation is available from the start"


def _generator():
    spec = importlib.util.spec_from_file_location(
        "ppt_generate_surface_map", ROOT / "scripts" / "generate_surface_map.py"
    )
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def test_every_public_op_is_visible_on_exactly_one_fixed_tool():
    surface = plan.build(Settings(), registry.public_ops())
    seen: dict[str, str] = {}
    for tool_name, tool in surface.tools.items():
        if tool_name == plan.HELP_TOOL:
            continue
        for op in tool.ops:
            assert op not in seen, (op, seen.get(op), tool_name)
            seen[op] = tool_name
    assert set(seen) == set(registry.public_ops())
    assert len(seen) == 145


def test_every_released_tool_maps_to_a_real_destination_or_exact_removal():
    released = json.loads(RELEASED.read_text(encoding="utf-8"))["tools"]
    mapping = json.loads(MAP.read_text(encoding="utf-8"))["tools"]
    assert len(released) == 142
    assert set(mapping) == set(released)

    surface = plan.build(Settings(), registry.public_ops())
    for old, entry in mapping.items():
        assert entry["args_changed"] is None
        if old in {"enable_tools", "disable_tools"}:
            assert entry == {
                "tool": None,
                "op": None,
                "args_changed": None,
                "reason": REMOVED_REASON,
            }
        else:
            assert set(entry) == {"tool", "op", "args_changed"}
            assert entry["op"] == old
            assert old in surface.tools[entry["tool"]].ops


def test_checked_in_json_and_markdown_are_fresh():
    table, data = _generator().render()
    assert TABLE.read_text(encoding="utf-8") == table, (
        "docs/surface_map.md is stale; rerun scripts/generate_surface_map.py"
    )
    assert json.loads(MAP.read_text(encoding="utf-8")) == data, (
        "migration/v1_to_v2.json is stale; rerun scripts/generate_surface_map.py"
    )


def test_measured_surface_stays_under_the_phase_b_ceiling():
    spec = importlib.util.spec_from_file_location(
        "ppt_measure_surface", ROOT / "scripts" / "measure_surface.py"
    )
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    report = module.measure()
    assert report["approx_tokens"] <= 9_022
    assert report["approx_token_ceiling"] == 9_022
