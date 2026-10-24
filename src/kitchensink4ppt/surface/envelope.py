"""Small, single-emission response helpers for the family surface."""

from __future__ import annotations

import json
from typing import Any
from xml.etree.ElementTree import ParseError as XmlParseError

from lxml import etree as lxml_etree

from kitchensink4ppt.core import errors as _errors
from kitchensink4ppt.core.errors import OpError
from kitchensink4ppt.core.safesave import MutationLockTimeout
from kitchensink4ppt.core.sandbox import SandboxViolation

from . import copy


_CODE_MAP: tuple[tuple[type[BaseException], str], ...] = (
    (_errors.AmbiguousTarget, "AMBIGUOUS_LOCATION"),
    (_errors.TargetNotFound, "NOT_FOUND"),
    (_errors.DocumentNotFound, "NOT_FOUND"),
    (MutationLockTimeout, "DOCUMENT_LOCKED"),
    (_errors.DocumentReadOnly, "PROTECTED"),
    (_errors.DocumentLocked, "DOCUMENT_LOCKED"),
    (_errors.DocumentBusy, "DOCUMENT_LOCKED"),
    (_errors.DocumentCorrupt, "UNSUPPORTED_CONTENT"),
    (_errors.DocumentEditProtected, "PROTECTED"),
    (_errors.DocumentProtected, "UNSUPPORTED_CONTENT"),
    (_errors.UnsupportedStructure, "UNSUPPORTED_CONTENT"),
    (_errors.ValidationFailed, "VALIDATION_FAILED"),
    (_errors.ProtectedViewRefused, "PROTECTED_VIEW"),
    (_errors.PowerPointNotRunning, "APP_NOT_RUNNING"),
    (_errors.DocumentNotOpenInPowerPoint, "APP_NOT_RUNNING"),
    (_errors.LiveLockTimeout, "APP_BUSY"),
    (_errors.PowerPointBusy, "APP_BUSY"),
    (_errors.PowerPointBlocked, "APP_BLOCKED"),
    (_errors.PowerPointDisconnected, "CONFLICT"),
    (SandboxViolation, "BAD_PARAMS"),
    (_errors.PptMcpError, "BAD_PARAMS"),
    (FileExistsError, "CONFLICT"),
    (FileNotFoundError, "NOT_FOUND"),
    (ValueError, "BAD_PARAMS"),
    (TypeError, "BAD_PARAMS"),
    (XmlParseError, "BAD_PARAMS"),
    (lxml_etree.LxmlError, "BAD_PARAMS"),
    (RecursionError, "UNSUPPORTED_CONTENT"),
    (OverflowError, "BAD_PARAMS"),
)


def _classify(exc: BaseException) -> str:
    declared = getattr(exc, "code", None)
    if isinstance(declared, str) and declared in _errors.CODES:
        return declared
    for exception_type, code in _CODE_MAP:
        if isinstance(exc, exception_type):
            return code
    return "INTERNAL_ERROR"


def success(payload: dict[str, Any]) -> dict[str, Any]:
    """Keep the product's flat ``ok`` and ``file`` success fields."""
    result = dict(payload)
    # Some successful operations return a negative domain verdict.  For
    # example, ``validate`` can complete normally and report that a deck is
    # not clean.  Preserve that verdict while adding the transport success
    # flag for handlers that do not already carry one.
    result.setdefault("ok", True)
    return result


def refusal(exc: BaseException) -> dict[str, Any]:
    if isinstance(exc, OpError):
        code = exc.code
        error = {
            "code": code,
            "message": exc.message,
            "hint": exc.hint or copy.HINTS[code],
            "details": exc.details or {},
        }
    else:
        code = _classify(exc)
        declared_hint = getattr(exc, "hint", None)
        declared_detail = getattr(exc, "details", None)
        if declared_detail is None:
            declared_detail = getattr(exc, "detail", None)
        details = declared_detail if isinstance(declared_detail, dict) else {}
        message = str(exc)
        error = {
            "code": code,
            "message": message if code != "INTERNAL_ERROR" else copy.INTERNAL_ERROR_MESSAGE,
            "hint": declared_hint if isinstance(declared_hint, str) and declared_hint.strip()
                    else copy.HINTS[code],
            "details": details,
        }
    return {"ok": False, "error": error}


def compact_json(payload: dict[str, Any]) -> str:
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def tool_result(payload: dict[str, Any], *, is_error: bool = False):
    """Put the payload on the wire once, as one compact text block."""
    from fastmcp.tools.tool import ToolResult
    from mcp.types import TextContent

    return ToolResult(
        content=[TextContent(type="text", text=compact_json(payload))],
        is_error=is_error,
    )


__all__ = ["compact_json", "refusal", "success", "tool_result"]
