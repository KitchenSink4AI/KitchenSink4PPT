"""Operation registry for the fixed KitchenSink4PPT family surface.

The editing functions still live in :mod:`kitchensink4ppt.server`.  During
the surface migration the server's private ``_tool`` decorator registers each
function here instead of registering 140 independent MCP tools.  Argument
models are generated once, at registration time, from the shipped function
signatures; calls then validate against those frozen strict models.
"""

from __future__ import annotations

import inspect
import re
from dataclasses import dataclass
from typing import Any, Callable, Mapping, get_args, get_type_hints

from pydantic import BaseModel, ConfigDict, ValidationError, create_model


GROUP_OPS: dict[str, tuple[str, ...]] = {
    "ppt_read": (
        "get_presentation_view", "get_presentation_info", "get_slide_info",
        "get_text", "find_text", "list_elements", "list_hyperlinks",
        "get_table", "get_chart_data", "get_notes", "extract_text",
        "list_comments", "list_equations", "get_media_playback",
        "get_transitions", "list_animations", "get_server_info",
        "get_workflows",
    ),
    "ppt_inspect": (
        "audit_accessibility", "check_layout", "font_inventory",
        "comment_report", "compare_decks", "deck_statistics", "diagnose",
        "extract_brand", "zombie_check", "get_theme", "get_autofit_state",
        "get_document_properties", "get_footer_support",
        "list_master_elements", "get_export_engines", "live_status",
        "powerpoint_status",
    ),
    "ppt_edit": (
        "apply_edits", "search_and_replace", "set_placeholder_text",
        "fit_text", "format_text", "set_bullets", "insert_textbox",
        "set_notes", "delete_notes", "set_hyperlink", "remove_hyperlink",
        "insert_slide", "delete_slide", "duplicate_slide", "reorder_slides",
        "move_slide", "set_slide_hidden", "manage_section", "set_language",
    ),
    "ppt_graphics": (
        "insert_shape", "insert_connector", "set_shape", "set_shape_adjustments",
        "delete_shape",
        "group_shapes", "ungroup_shapes", "align_shapes", "distribute_shapes",
        "set_z_order", "svg_to_shapes", "set_diagram_text",
        "generate_diagram", "copy_format", "copy_position",
        "copy_shape_between", "insert_image",
        "replace_image", "set_image", "replace_image_everywhere",
        "insert_video", "insert_audio", "set_media_playback",
        "insert_equation", "add_equation_to_shape", "set_alt_text",
        "set_reading_order",
    ),
    "ppt_tables": (
        "create_table", "set_table_cells", "merge_cells", "unmerge_cells",
        "insert_table_rows", "delete_table_rows", "insert_table_cols",
        "delete_table_cols", "format_table_cells", "set_column_widths",
        "set_row_heights", "apply_table_style", "import_table",
        "create_chart", "update_chart_data", "format_chart",
    ),
    "ppt_design": (
        "set_slide_size", "apply_layout", "set_theme_colors",
        "set_theme_fonts", "apply_brand", "set_slide_background",
        "set_master_placeholder", "set_layout_placeholder",
        "add_layout_placeholder", "remove_layout_placeholder",
        "insert_master_shape", "delete_master_shape", "create_layout",
        "set_master_background", "replace_fonts", "replace_colors",
        "set_transition", "add_entrance_animation", "clear_animations",
    ),
    "ppt_review": (
        "add_comment", "reply_to_comment", "resolve_comment",
        "delete_comment", "anonymize_deck", "update_now", "update_notices",
    ),
    "ppt_deck": (
        "create_presentation", "copy_presentation", "create_snapshot",
        "manage_backups", "copy_slide_between", "merge_decks", "split_deck",
        "export_pdf", "export_slide_image", "export_shape_image", "export_handout",
        "export_table",
        "set_document_properties", "generate_agenda_slide",
        "refresh_agenda_slide", "compress_deck", "set_show_properties",
        "manage_custom_show", "set_footer", "validate", "live_save",
        "live_scroll_to",
    ),
}

GROUPS = tuple(GROUP_OPS)
OP_TO_GROUP = {name: group for group, names in GROUP_OPS.items() for name in names}
if len(OP_TO_GROUP) != sum(len(names) for names in GROUP_OPS.values()):  # pragma: no cover
    raise RuntimeError("an operation appears in more than one PowerPoint surface group")

RESERVED_ARGS = frozenset()

# These operations can enter the serialized PowerPoint bridge.  The bridge
# starts identity-checked cancellation at 60 seconds; the outer 120-second
# budget leaves room for its false-timeout grace and bounded process and
# presentation revalidation to return an honest cancellation result.
# validate gives PowerPoint more time on a bigger deck, up to the bridge's
# OPENS_CLEAN_CAP_S (300 s), so its budget is that cap plus the same room.
_COM_DEFAULT_TIMEOUT_S = {
    "export_pdf": 120.0,
    "export_handout": 120.0,
    "export_slide_image": 120.0,
    "export_shape_image": 120.0,
    "validate": 360.0,
}


class RegistryError(RuntimeError):
    """The operation catalog is inconsistent with the adopted surface."""


class StrictArgs(BaseModel):
    """Base for generated operation arguments: unknown keys always refuse."""

    model_config = ConfigDict(extra="forbid", arbitrary_types_allowed=True)


@dataclass(frozen=True)
class OpSpec:
    name: str
    group: str
    summary: str
    args: type[BaseModel]
    handler: Callable[..., Any]
    writes: bool
    read_only: bool
    destructive: bool
    runs_live: bool
    timeout_s: float | None = None

    @property
    def internal(self) -> bool:
        return self.name.startswith("_")

    def invoke(self, payload: Mapping[str, Any]) -> dict[str, Any]:
        """Validate a payload and invoke the existing keyword-argument handler."""
        try:
            model = self.args.model_validate(dict(payload))
        except ValidationError as exc:
            raise RegistryError(
                f"arguments for {self.name} did not match its schema: "
                f"{validation_problems(exc)}"
            ) from None
        result = self.handler(**model.model_dump())
        if not isinstance(result, dict):
            raise RegistryError(f"{self.name} returned a non-object result")
        return result


REGISTRY: dict[str, OpSpec] = {}


def validation_problems(exc: ValidationError) -> list[dict[str, str]]:
    """Describe validation errors without echoing caller-supplied values."""
    return [
        {
            "field": ".".join(str(part) for part in item["loc"]) or "(root)",
            "problem": item["msg"],
        }
        for item in exc.errors(include_url=False, include_input=False)
    ]


def first_docstring_sentence(fn: Callable[..., Any]) -> str:
    """Return the first shipped docstring sentence, normalized to one line."""
    doc = inspect.getdoc(fn) or ""
    text = " ".join(doc.split())
    if not text:
        raise RegistryError(f"op {fn.__name__!r} needs a docstring summary")
    match = re.search(r"(?<=[.!?])(?:\s|$)", text)
    return text[: match.start()].strip() if match else text


def _is_context(name: str, annotation: Any) -> bool:
    if name in {"ctx", "context"}:
        return True
    candidates = (annotation, *get_args(annotation))
    return any(
        (getattr(candidate, "__name__", "") or getattr(candidate, "_name", "")) == "Context"
        for candidate in candidates
    )


def args_model(fn: Callable[..., Any], *, name: str | None = None) -> type[BaseModel]:
    """Generate the strict Pydantic model for ``fn`` exactly once."""
    signature = inspect.signature(fn, follow_wrapped=True)
    try:
        hints = get_type_hints(fn)
    except (NameError, TypeError):
        hints = {}
    fields: dict[str, tuple[Any, Any]] = {}
    for param in signature.parameters.values():
        annotation = hints.get(param.name, param.annotation)
        if _is_context(param.name, annotation):
            continue
        if param.kind in (inspect.Parameter.VAR_POSITIONAL, inspect.Parameter.VAR_KEYWORD):
            raise RegistryError(f"op {name or fn.__name__!r} cannot expose variadic arguments")
        if param.kind is inspect.Parameter.POSITIONAL_ONLY:
            raise RegistryError(f"op {name or fn.__name__!r} cannot expose positional-only arguments")
        if annotation is inspect.Parameter.empty:
            annotation = Any
        default = ... if param.default is inspect.Parameter.empty else param.default
        fields[param.name] = (annotation, default)
    model_name = "".join(part.capitalize() for part in (name or fn.__name__).split("_")) + "Args"
    return create_model(model_name, __base__=StrictArgs, **fields)


def _classifications(name: str) -> tuple[bool, bool]:
    from .core.readonly import read_only_hint
    from .core.tool_annotations import destructive_hint

    read_only = read_only_hint(name)
    destructive = False if read_only else destructive_hint(name)
    if not read_only and destructive is None:
        raise RegistryError(f"op {name!r} has no destructive classification")
    return read_only, bool(destructive)


def register(
    handler: Callable[..., Any],
    *,
    name: str | None = None,
    group: str | None = None,
    summary: str | None = None,
    writes: bool | None = None,
    read_only: bool | None = None,
    destructive: bool | None = None,
    runs_live: bool | None = None,
    timeout_s: float | None = None,
) -> Callable[..., Any]:
    """Register an existing boundary wrapper and return it unchanged."""
    op_name = name or handler.__name__
    expected_group = OP_TO_GROUP.get(op_name)
    if expected_group is None and not op_name.startswith("_"):
        raise RegistryError(f"op {op_name!r} is not assigned to a surface group")
    chosen_group = group or expected_group or ""
    if expected_group is not None and chosen_group != expected_group:
        raise RegistryError(
            f"op {op_name!r} belongs to {expected_group}, not {chosen_group}"
        )
    if op_name in REGISTRY:
        raise RegistryError(f"op {op_name!r} registered twice")
    classified_ro, classified_destructive = _classifications(op_name)
    if read_only is not None and bool(read_only) != classified_ro:
        raise RegistryError(f"op {op_name!r} conflicts with its read-only classification")
    if destructive is not None and bool(destructive) != classified_destructive:
        raise RegistryError(f"op {op_name!r} conflicts with its destructive classification")
    effective_timeout = (
        _COM_DEFAULT_TIMEOUT_S.get(op_name) if timeout_s is None else timeout_s
    )
    if effective_timeout is not None and not 0 < float(effective_timeout) <= 3600:
        raise RegistryError(f"op {op_name!r} timeout_s must be above 0 and at most 3600")
    signature = inspect.signature(handler, follow_wrapped=True)
    inferred_live = "live" in signature.parameters or op_name.startswith("live_") or op_name in {
        "powerpoint_status", "zombie_check"
    }
    spec = OpSpec(
        name=op_name,
        group=chosen_group,
        summary=(summary or first_docstring_sentence(handler)).strip(),
        args=args_model(handler, name=op_name),
        handler=handler,
        writes=(not classified_ro) if writes is None else bool(writes),
        read_only=classified_ro,
        destructive=classified_destructive,
        runs_live=inferred_live if runs_live is None else bool(runs_live),
        timeout_s=(
            None if effective_timeout is None else float(effective_timeout)
        ),
    )
    if "\n" in spec.summary or not spec.summary:
        raise RegistryError(f"op {op_name!r} needs a one-line summary")
    REGISTRY[op_name] = spec
    return handler


def op(**metadata: Any) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Decorator form of :func:`register`, useful for isolated operations."""
    return lambda fn: register(fn, **metadata)


def public_ops() -> dict[str, OpSpec]:
    return {name: spec for name, spec in REGISTRY.items() if not spec.internal}


def clear_for_tests() -> None:
    """Reset the process-local catalog. Unit tests only; production never calls this."""
    REGISTRY.clear()


__all__ = [
    "GROUP_OPS", "GROUPS", "OP_TO_GROUP", "REGISTRY", "RESERVED_ARGS",
    "OpSpec", "RegistryError", "StrictArgs", "args_model", "clear_for_tests",
    "first_docstring_sentence", "op", "public_ops", "register",
    "validation_problems",
]
