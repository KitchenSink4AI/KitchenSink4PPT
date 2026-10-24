"""Install the fixed surface on one FastMCP server instance.

This is deliberately the only surface module that calls ``mcp.add_tool``.
The plan is captured by each installed tool instance, so repeated
``create_server(settings=...)`` calls do not share mutable runtime state.
"""

from __future__ import annotations

from typing import Any

from fastmcp.tools.tool import Tool, ToolResult
from mcp.types import ToolAnnotations
from pydantic import PrivateAttr

from kitchensink4ppt.core.errors import OpError
from kitchensink4ppt.core import budget
from kitchensink4ppt.core import update_flow

from . import envelope
from .dispatch import dispatch, shape
from .help import help_payload
from .plan import HELP_TOOL, SurfacePlan, ToolPlan


class PptTool(Tool):
    """A group or help tool bound to one immutable startup plan."""

    _plan: SurfacePlan = PrivateAttr()

    def __init__(self, *, plan: SurfacePlan, **data: Any) -> None:
        super().__init__(**data)
        self._plan = plan

    async def run(self, arguments: dict[str, Any]) -> ToolResult:
        if self.name == HELP_TOOL:
            try:
                if not isinstance(arguments, dict) or set(arguments) - {"op"}:
                    raise OpError("BAD_ARGS", "ppt_help takes one optional argument, op")
                result = help_payload(self._plan, arguments.get("op"))
                result = budget.enforce(
                    result,
                    self._plan.settings.max_output_chars,
                    wrap=envelope.success,
                )
                payload = envelope.success(result)
                is_error = False
            except BaseException as exc:
                if isinstance(exc, (KeyboardInterrupt, SystemExit)):
                    raise
                payload = budget.cap_refusal(
                    envelope.refusal(exc), self._plan.settings.max_output_chars
                )
                is_error = True
        else:
            # The update hook (core/update_flow.py), step 1: update_now's own
            # Yes/No dialog, before the op runs. It sees the op name exactly
            # as the dispatcher routes it.
            token = await update_flow.before_call(_routed_op(self._plan, arguments))
            try:
                payload, is_error = await dispatch(self._plan, self.name, arguments)
            finally:
                update_flow.end_call(token)
        # Step 2: the weekly background check and, after a read-only call
        # only, the update notice (a dialog, or one `update` key).
        tp = self._plan.tools.get(self.name)
        payload = await update_flow.after_call(
            payload, is_error,
            read_only=self.name == HELP_TOOL or bool(tp and tp.read_only),
            cap=self._plan.settings.max_output_chars,
            can_update=self._plan.is_visible("update_now"))
        return envelope.tool_result(payload, is_error=is_error)


def _routed_op(plan: SurfacePlan, arguments: Any) -> str | None:
    """The canonical op a call routes to, or None when it does not parse."""
    try:
        return shape(plan, arguments)[0]
    except Exception:  # noqa: BLE001 (dispatch reports the refusal itself)
        return None


def make_tool(tool: ToolPlan, plan: SurfacePlan) -> PptTool:
    annotations = tool.annotations
    return PptTool(
        plan=plan,
        name=tool.name,
        title=tool.title,
        description=tool.description,
        parameters=tool.input_schema,
        output_schema=None,
        annotations=ToolAnnotations(
            title=annotations["title"],
            readOnlyHint=annotations["readOnlyHint"],
            destructiveHint=annotations["destructiveHint"],
            idempotentHint=annotations["idempotentHint"],
            openWorldHint=annotations["openWorldHint"],
        ),
    )


def install(mcp: Any, plan: SurfacePlan) -> tuple[PptTool, ...]:
    """Install every planned tool on ``mcp`` in stable plan order."""
    installed: list[PptTool] = []
    for tool_plan in plan.tools.values():
        tool = make_tool(tool_plan, plan)
        mcp.add_tool(tool)
        installed.append(tool)
    return tuple(installed)


__all__ = ["PptTool", "install", "make_tool"]
