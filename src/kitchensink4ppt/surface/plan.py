"""Build the fixed nine-tool PowerPoint surface once at startup."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any, Mapping

from ..registry import GROUP_OPS, GROUPS, OP_TO_GROUP, OpSpec, public_ops
from ..settings import ENV_FORBID, Settings, SettingsError
from . import copy


HELP_TOOL = "ppt_help"
POWERPOINT_APP_CAPABILITY = "ppt_app"
READ_ONLY_GROUPS = frozenset({"ppt_read", "ppt_inspect"})
ALL_TOOLS = GROUPS + (HELP_TOOL,)
MAX_DESCRIPTION_CHARS = 2048
MAX_SCHEMA_BYTES = 5000
MAX_TOOL_NAME_CHARS = 64

#: Ops that reach a remote service (update_now asks PyPI and GitHub), so the
#: tool holding them says openWorldHint true. Every other tool says false.
OPEN_WORLD_OPS = frozenset({"update_now"})


class PlanError(RuntimeError):
    """The registry violates the adopted PowerPoint surface contract."""


@dataclass(frozen=True)
class ToolPlan:
    name: str
    ops: tuple[str, ...]
    read_only: bool
    destructive: bool
    title: str
    description: str
    input_schema: dict[str, Any]
    annotations: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class SurfacePlan:
    settings: Settings
    tools: dict[str, ToolPlan]
    op_tool: dict[str, str]
    forbidden_ops: frozenset[str]
    forbidden_tools: frozenset[str]
    hidden_by_read_only: frozenset[str]
    specs: Mapping[str, OpSpec]

    def visible_ops(self) -> dict[str, list[str]]:
        return {
            tool.name: list(tool.ops)
            for tool in self.tools.values()
            if tool.name != HELP_TOOL
        }

    def is_visible(self, op: str) -> bool:
        tool = self.op_tool.get(op)
        return tool in self.tools and op in self.tools[tool].ops


def schema_bytes(schema: dict[str, Any]) -> int:
    return len(json.dumps(schema, separators=(",", ":"), ensure_ascii=False).encode("utf-8"))


def _op_line(spec: OpSpec) -> str:
    return f"- {spec.name}: {copy.OP_LINES.get(spec.name) or spec.summary}"


def _layout(head: str, lines: list[str]) -> tuple[str, str | None]:
    """Keep whole op lines within the description limit, moving overflow."""
    full = "\n".join([head, copy.OPS_HEADING, *lines])
    if len(full) <= MAX_DESCRIPTION_CHARS:
        return full, None
    kept: list[str] = []
    for line in lines:
        trial = "\n".join([head, copy.OPS_HEADING, *kept, line, copy.OPS_CONTINUED])
        if len(trial) > MAX_DESCRIPTION_CHARS:
            break
        kept.append(line)
    rest = lines[len(kept):]
    description = "\n".join([head, copy.OPS_HEADING, *kept, copy.OPS_CONTINUED])
    return description, "\n".join([copy.OP_PARAM_CONTINUED, *rest])


def _group_schema(ops: list[str], op_description: str | None = None) -> dict[str, Any]:
    return {
        "type": "object",
        "properties": {
            "op": {
                "type": "string",
                "enum": ops,
                "description": op_description or copy.OP_PARAM,
            },
            "args": {
                "type": "object",
                "description": copy.ARGS_PARAM,
                "additionalProperties": True,
            },
        },
        "required": ["op"],
        "additionalProperties": False,
    }


def _help_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "properties": {
            "op": {"type": "string", "description": copy.HELP_OP_PARAM},
        },
        "additionalProperties": False,
    }


def check_registry(ops: Mapping[str, OpSpec]) -> None:
    for name, spec in ops.items():
        expected = OP_TO_GROUP.get(name)
        if expected is None:
            raise PlanError(f"op {name} is not assigned to the adopted surface")
        if spec.group != expected:
            raise PlanError(f"op {name} belongs to {expected}, not {spec.group}")
        group_read_only = spec.group in READ_ONLY_GROUPS
        if spec.read_only != group_read_only:
            kind = "read-only" if spec.read_only else "writing"
            raise PlanError(f"{kind} op {name} is in incompatible tool {spec.group}")
        if spec.read_only and (spec.writes or spec.destructive):
            raise PlanError(f"read-only op {name} carries a writing annotation")
    for group in GROUPS:
        if group not in copy.TOOLS:
            raise PlanError(f"no copy slot for tool {group}")
    if HELP_TOOL not in copy.TOOLS:
        raise PlanError(f"no copy slot for tool {HELP_TOOL}")


def build(settings: Settings, ops: Mapping[str, OpSpec] | None = None) -> SurfacePlan:
    selected = public_ops() if ops is None else {
        name: spec for name, spec in ops.items() if not spec.internal
    }
    from ..core import update_flow

    if not update_flow.community():
        # Business: the update ops are not on its tool list (its tools stay
        # byte-identical to the release before them).
        selected = {n: s for n, s in selected.items() if n not in update_flow.OPS}
    check_registry(selected)
    known = set(ALL_TOOLS) | set(selected) | {POWERPOINT_APP_CAPABILITY}
    unknown = sorted(set(settings.forbid) - known)
    if unknown:
        raise SettingsError(
            f"{ENV_FORBID} names {', '.join(unknown)}, which is not a tool or op of this server"
        )

    forbidden_tools = frozenset(name for name in settings.forbid if name in ALL_TOOLS)
    forbidden_op_names = {name for name in settings.forbid if name in selected}
    if POWERPOINT_APP_CAPABILITY in settings.forbid:
        # ``runs_live`` covers dual-mode and live-only routes. A bounded
        # operation timeout marks the explicit COM bridge exports/validator.
        # Fail closed: the virtual capability is unavailable if an operation
        # can attach to or start PowerPoint through either route.
        forbidden_op_names.update(
            name
            for name, spec in selected.items()
            if spec.runs_live or spec.timeout_s is not None
        )
    forbidden_ops = frozenset(forbidden_op_names)
    hidden_read_only: set[str] = set()
    op_tool = {name: spec.group for name, spec in selected.items()}
    tools: dict[str, ToolPlan] = {}

    for group in GROUPS:
        read_only = group in READ_ONLY_GROUPS
        members = [name for name in GROUP_OPS[group] if name in selected]
        if settings.read_only and not read_only:
            hidden_read_only.update(members)
            continue
        if group in forbidden_tools:
            continue
        visible = [name for name in members if name not in forbidden_ops]
        if not visible:
            continue
        slot = copy.TOOLS[group]
        description, op_description = _layout(
            slot["description"], [_op_line(selected[name]) for name in visible]
        )
        destructive = any(selected[name].destructive for name in visible)
        annotations = {
            "title": slot["title"],
            "readOnlyHint": read_only,
            "destructiveHint": destructive,
            "idempotentHint": read_only,
            "openWorldHint": any(name in OPEN_WORLD_OPS for name in visible),
        }
        tools[group] = ToolPlan(
            group, tuple(visible), read_only, destructive, slot["title"],
            description, _group_schema(visible, op_description), annotations,
        )

    if HELP_TOOL not in forbidden_tools and tools:
        slot = copy.TOOLS[HELP_TOOL]
        listed = tuple(op for group in GROUPS if group in tools for op in tools[group].ops)
        tools[HELP_TOOL] = ToolPlan(
            HELP_TOOL, listed, True, False, slot["title"], slot["description"],
            _help_schema(),
            {
                "title": slot["title"],
                "readOnlyHint": True,
                "destructiveHint": False,
                "idempotentHint": True,
                "openWorldHint": False,
            },
        )

    for name, tool in tools.items():
        if len(name) > MAX_TOOL_NAME_CHARS:
            raise PlanError(f"tool name {name!r} exceeds {MAX_TOOL_NAME_CHARS} characters")
        if len(tool.description) > MAX_DESCRIPTION_CHARS:
            raise PlanError(f"tool {name} description exceeds {MAX_DESCRIPTION_CHARS} characters")
        if schema_bytes(tool.input_schema) > MAX_SCHEMA_BYTES:
            raise PlanError(f"tool {name} schema exceeds {MAX_SCHEMA_BYTES} bytes")
        if not tool.title:
            raise PlanError(f"tool {name} has no title")

    return SurfacePlan(
        settings=settings,
        tools=tools,
        op_tool=op_tool,
        forbidden_ops=forbidden_ops,
        forbidden_tools=forbidden_tools,
        hidden_by_read_only=frozenset(hidden_read_only),
        specs=selected,
    )

__all__ = [
    "ALL_TOOLS", "HELP_TOOL", "MAX_DESCRIPTION_CHARS", "MAX_SCHEMA_BYTES",
    "OPEN_WORLD_OPS", "POWERPOINT_APP_CAPABILITY", "PlanError", "READ_ONLY_GROUPS", "SurfacePlan", "ToolPlan", "build",
    "check_registry", "schema_bytes",
]
