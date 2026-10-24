"""Acceptance pins for the fixed PowerPoint family surface."""

from __future__ import annotations

import ast
import asyncio
import json
import re
from pathlib import Path

from fastmcp import Client

from kitchensink4ppt import server


ROOT = Path(__file__).resolve().parents[2]
EXPECTED = {
    "ppt_read",
    "ppt_inspect",
    "ppt_edit",
    "ppt_graphics",
    "ppt_tables",
    "ppt_design",
    "ppt_review",
    "ppt_deck",
    "ppt_help",
}


async def _listed() -> tuple[set[str], object]:
    async with Client(server.create_server()) as client:
        names = {tool.name for tool in await client.list_tools()}
        return names, client.initialize_result


def test_wire_lists_exactly_the_nine_fixed_tools():
    names, _ = asyncio.run(_listed())
    assert names == EXPECTED


def test_the_server_does_not_advertise_list_changed():
    _, initialized = asyncio.run(_listed())
    tools = initialized.capabilities.tools
    assert not bool(getattr(tools, "listChanged", False))


def test_pack_surface_and_runtime_toggles_are_gone():
    assert not (ROOT / "src" / "kitchensink4ppt" / "packs.py").exists()
    assert not hasattr(server, "enable_tools")
    assert not hasattr(server, "disable_tools")
    lowered = server.create_server().instructions.lower()
    assert "enable_tools" not in lowered
    assert not re.search(r"\bpacks?\b", lowered)


def test_only_surface_tools_registers_mcp_tools():
    offenders: list[str] = []
    source_root = ROOT / "src" / "kitchensink4ppt"
    for path in source_root.rglob("*.py"):
        relative = path.relative_to(source_root).as_posix()
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            target = node.func
            if (
                isinstance(target, ast.Attribute)
                and target.attr in {"add_tool", "tool"}
            ):
                if relative != "surface/tools.py":
                    offenders.append(f"{relative}:{node.lineno}:{target.attr}")
    assert not offenders, offenders


def test_one_fixed_tool_dispatches_a_released_operation():
    async def call():
        async with Client(server.create_server()) as client:
            result = await client.call_tool(
                "ppt_read", {"op": "get_server_info", "args": {}}
            )
            return result, {tool.name for tool in await client.list_tools()}

    result, after = asyncio.run(call())
    assert not result.is_error
    assert result.structured_content is None
    payload = json.loads("".join(
        block.text for block in result.content if hasattr(block, "text")
    ))
    assert payload["ok"] is True
    assert payload["edition"] in {
        "community",
        "business",
    }
    assert after == EXPECTED
