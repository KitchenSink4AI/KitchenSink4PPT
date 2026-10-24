"""Parse, resolve, strictly validate and invoke one family-surface operation."""

from __future__ import annotations

import asyncio
import atexit
import contextvars
import functools
import itertools
import json
import sys
import threading
import time
from typing import Any, Callable, Mapping

from pydantic import ValidationError

from kitchensink4ppt.core import budget
from kitchensink4ppt.core import deadline as call_deadline
from kitchensink4ppt.core import messages as copy_messages
from kitchensink4ppt.core import slide_echo
from kitchensink4ppt.core import package as _package
from kitchensink4ppt.core.errors import OpError
from kitchensink4ppt.core.safesave import take_notices as _recovery_notices

from ..registry import OpSpec, validation_problems
from . import copy, envelope, normalize
from .plan import HELP_TOOL, SurfacePlan


_TOOL_KEYS = frozenset({"op", "args"})
#: Operations whose answer names files after slides counted from 1: when
#: the call sent their 0-based list itself (not its 1-based twin), the
#: answer says how the numbers were read (core/messages.SLIDES_NOTE).
_ZERO_BASED_LISTS = {"export_slide_image": "slides"}

# A publication that wins the deadline race gets one last finite wait for its
# real result. Long COM operations use their own larger per-op timeout here.
MIN_POST_COMMIT_WAIT_S = 60.0

WORKER_PREFIX = "asyncio-ks4p-call"
_worker_seq = itertools.count(1)

# Timed-out calls keep running because Python cannot stop a thread safely.
# Track them so process shutdown can stop every call that has not crossed its
# publication point, then give all workers a bounded chance to release files
# and COM objects before daemon-thread teardown.
EXIT_DRAIN_S = 60.0
_INFLIGHT: dict[threading.Thread, call_deadline.Deadline] = {}
_INFLIGHT_LOCK = threading.Lock()
_EXIT_DRAIN_LOCK = threading.Lock()
_EXIT_DRAINED = False
_EXIT_DRAIN_DONE = threading.Event()


def has_inflight() -> bool:
    """Return whether a dispatched operation worker is still running."""
    current = threading.current_thread()
    with _INFLIGHT_LOCK:
        return any(
            worker is not current and worker.is_alive()
            for worker in _INFLIGHT
        )


def _consume_late_result(task: asyncio.Future) -> None:
    """Retrieve a cancelled-deadline worker's eventual exception quietly."""
    if task.cancelled():
        return
    try:
        task.exception()
    except (asyncio.CancelledError, Exception):
        pass


def drain_inflight(bound: float | None = None) -> dict[str, int]:
    """Cancel unpublished calls and wait a bounded total time for workers.

    This never ends a thread or a PowerPoint process.  A worker that remains
    blocked after the bound is reported once and process exit continues.
    """
    if bound is None:
        # Resolve at call time so a host (and the shutdown regression tests)
        # can tune the grace period after importing the module.
        bound = EXIT_DRAIN_S
    end = time.monotonic() + max(0.0, float(bound))
    current = threading.current_thread()
    with _INFLIGHT_LOCK:
        items = [
            (worker, deadline)
            for worker, deadline in _INFLIGHT.items()
            if worker is not current and worker.is_alive()
        ]
    cancelled = 0
    for _worker, deadline in items:
        was_cancelled = deadline.cancelled
        if deadline.cancel() and not was_cancelled:
            cancelled += 1
    for worker, _deadline in items:
        worker.join(max(0.0, end - time.monotonic()))
    unfinished = sum(1 for worker, _deadline in items if worker.is_alive())
    if unfinished:
        sys.stderr.write(
            f"[kitchensink4ppt] {unfinished} call(s) still running after "
            f"{max(0.0, float(bound)):g}s at exit; no process was force-ended\n"
        )
        sys.stderr.flush()
    return {"cancelled": cancelled, "waited": len(items), "unfinished": unfinished}


def drain_at_exit_once() -> dict[str, int]:
    """Run the server-exit drain exactly once, including owned COM cleanup."""
    global _EXIT_DRAINED
    with _EXIT_DRAIN_LOCK:
        if _EXIT_DRAINED:
            _EXIT_DRAIN_DONE.wait(EXIT_DRAIN_S + 5.0)
            return {"cancelled": 0, "waited": 0, "unfinished": 0}
        _EXIT_DRAINED = True
    try:
        result = drain_inflight()
        try:
            from kitchensink4ppt.com import bridge

            # The outer call workers already received the one shared
            # 60-second exit grace. Reap only positively identified COM
            # processes now; do not impose a second wait on the same call.
            bridge._drain_on_exit(timeout=0.0)
        except Exception as exc:
            sys.stderr.write(f"[kitchensink4ppt] COM exit cleanup failed: {exc}\n")
            sys.stderr.flush()
        return result
    finally:
        _EXIT_DRAIN_DONE.set()


def _drain_at_exit() -> None:
    drain_at_exit_once()


atexit.register(_drain_at_exit)


def _start_worker(
    fn: Callable[[], dict[str, Any]], deadline: call_deadline.Deadline
) -> asyncio.Future:
    """Run one call in its own daemon thread and return its result future.

    A timed-out worker cannot be stopped safely. Giving every call its own
    worker prevents those late threads from filling asyncio's shared default
    executor. Daemon workers also preserve process exit if a third-party call
    never returns; no PowerPoint process is force-closed here.
    """
    loop = asyncio.get_running_loop()
    future = loop.create_future()
    context = contextvars.copy_context()

    def settle(ok: bool, value: Any) -> None:
        if future.cancelled():
            return
        if ok:
            future.set_result(value)
        else:
            future.set_exception(value)

    def worker() -> None:
        try:
            outcome = (True, context.run(fn))
        except BaseException as exc:
            outcome = (False, exc)
        finally:
            with _INFLIGHT_LOCK:
                _INFLIGHT.pop(threading.current_thread(), None)
        try:
            loop.call_soon_threadsafe(settle, *outcome)
        except RuntimeError:
            pass

    future.add_done_callback(_consume_late_result)
    thread = threading.Thread(
        target=worker,
        daemon=True,
        name=f"{WORKER_PREFIX}-{next(_worker_seq)}",
    )
    with _INFLIGHT_LOCK:
        _INFLIGHT[thread] = deadline
    try:
        thread.start()
    except BaseException:
        with _INFLIGHT_LOCK:
            _INFLIGHT.pop(thread, None)
        raise
    return future


def parse(arguments: Any) -> tuple[str, dict[str, Any]]:
    if not isinstance(arguments, dict):
        raise OpError("BAD_ARGS", "tool arguments must be an object")
    extra = sorted(set(arguments) - _TOOL_KEYS)
    if extra:
        raise OpError(
            "BAD_ARGS",
            "unknown top-level keys; operation arguments belong inside args",
            details={"unknown": extra},
        )
    op = arguments.get("op")
    if not isinstance(op, str) or not op:
        raise OpError("BAD_ARGS", "op is required and must be an operation name")
    args = arguments.get("args")
    if args is None:
        args = {}
    if not isinstance(args, dict):
        raise OpError("BAD_ARGS", "args must be an object")
    return op, dict(args)


def unknown_op(plan: SurfacePlan, op: str, tool: str | None) -> OpError:
    """Build a nonexecuting NOT_FOUND refusal with bounded suggestions."""
    visible = {
        name: spec for name, spec in plan.specs.items() if plan.is_visible(name)
    }
    near = [
        {
            "op": name,
            "tool": plan.op_tool[name],
            "purpose": copy.OP_LINES.get(name) or visible[name].summary,
        }
        for name in normalize.suggestions(op, visible)
    ]
    details: dict[str, Any] = {"op": op}
    if tool is not None:
        details["tool"] = tool
    if near:
        details["did_you_mean"] = near
    return OpError(
        "NOT_FOUND",
        copy.OP_NOT_FOUND_MESSAGE.format(op=op),
        hint=copy_messages.OP_NOT_FOUND_HINT.format(op=op),
        details=details,
    )


def shape(
    plan: SurfacePlan, arguments: Any
) -> tuple[str, dict[str, Any], dict[str, Any]]:
    """Normalize the operation spelling and safe flat call shape."""
    sent: dict[str, Any] = {}
    if isinstance(arguments, dict):
        attempted = arguments.get("op")
        if isinstance(attempted, str) and attempted:
            canonical = normalize.route_op(attempted, plan.specs)
            if canonical is not None and canonical != attempted:
                sent["op"] = attempted
                arguments = {**arguments, "op": canonical}
            if canonical is not None:
                arguments, moved = normalize.flatten(
                    arguments, plan.specs[canonical].args.model_fields, op=canonical
                )
                if moved:
                    sent["flat_args"] = moved
    op, args = parse(arguments)
    return op, args, sent


def resolve(
    plan: SurfacePlan,
    tool: str,
    op: str,
    specs: Mapping[str, OpSpec] | None = None,
) -> OpSpec:
    catalog = plan.specs if specs is None else specs
    if op in plan.forbidden_ops or plan.op_tool.get(op) in plan.forbidden_tools:
        raise OpError(
            "NOT_AVAILABLE",
            f"{op} is turned off by this server's launch settings",
            details={"op": op, "decided_by": copy.SETTINGS_OWNER},
        )
    if op in plan.hidden_by_read_only:
        raise OpError(
            "NOT_AVAILABLE",
            f"{op} changes state and this server runs read-only",
            details={"op": op, "decided_by": copy.SETTINGS_OWNER},
        )
    if op not in catalog:
        raise unknown_op(plan, op, tool)
    owner = plan.op_tool.get(op)
    if owner != tool:
        raise OpError(
            "BAD_ARGS",
            copy_messages.WRONG_TOOL_MESSAGE.format(
                op=op, owner=owner, tool=tool
            ),
            details={"op": op, "tool": owner},
        )
    return catalog[op]


@functools.lru_cache(maxsize=512)
def _schema(model: type) -> dict[str, Any]:
    return model.model_json_schema()


def _admits(schema: Any, root: dict[str, Any], seen: frozenset[str] = frozenset()) -> set[str]:
    if not isinstance(schema, dict) or not schema:
        return {"any"}
    kinds: set[str] = set()
    ref = schema.get("$ref")
    if isinstance(ref, str):
        if not ref.startswith("#/") or ref in seen:
            return {"any"}
        node: Any = root
        for part in ref[2:].split("/"):
            node = node.get(part) if isinstance(node, dict) else None
        kinds |= _admits(node, root, seen | {ref})
    for key in ("anyOf", "oneOf"):
        for child in schema.get(key) or []:
            kinds |= _admits(child, root, seen)
    if schema.get("allOf"):
        parts = [_admits(child, root, seen) for child in schema["allOf"]]
        intersection = set.intersection(*parts) if parts else set()
        if parts and all("any" in part for part in parts):
            intersection = {"any"}
        kinds |= intersection
    raw_type = schema.get("type")
    if raw_type is not None:
        for kind in raw_type if isinstance(raw_type, list) else [raw_type]:
            kinds.add(kind if kind in {"object", "array", "string"} else "other")
    elif "properties" in schema or "additionalProperties" in schema:
        kinds.add("object")
    elif "items" in schema or "prefixItems" in schema:
        kinds.add("array")
    for key in ("const", "enum"):
        if key in schema:
            values = schema[key] if key == "enum" else [schema[key]]
            for value in values:
                kinds.add(
                    "string" if isinstance(value, str)
                    else "array" if isinstance(value, list)
                    else "object" if isinstance(value, dict)
                    else "other"
                )
    if not kinds and not ({"$ref", "anyOf", "oneOf", "allOf"} & set(schema)):
        return {"any"}
    return kinds


def decode_json_text(spec: OpSpec, args: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
    """Decode JSON text only for fields that require an object or array."""
    root = _schema(spec.args)
    properties = root.get("properties") or {}
    decoded: list[str] = []
    output = dict(args)
    for name, value in args.items():
        if not isinstance(value, str) or name not in properties:
            continue
        text = value.strip()
        if not text or text[0] not in "[{":
            continue
        kinds = _admits(properties[name], root)
        if "any" in kinds or "string" in kinds or not ({"array", "object"} & kinds):
            continue
        try:
            parsed = json.loads(text)
        except (ValueError, RecursionError):
            continue
        if (
            isinstance(parsed, list) and "array" in kinds
            or isinstance(parsed, dict) and "object" in kinds
        ):
            output[name] = parsed
            decoded.append(name)
    return output, decoded


def validate(spec: OpSpec, args: dict[str, Any]) -> dict[str, Any]:
    try:
        model = spec.args.model_validate(args)
    except ValidationError as exc:
        raise OpError(
            "BAD_ARGS",
            copy_messages.BAD_ARGS_MESSAGE.format(op=spec.name),
            hint=copy_messages.BAD_ARGS_HINT.format(op=spec.name),
            details={
                "op": spec.name,
                "problems": validation_problems(exc),
                "valid_fields": normalize.accepted_fields(
                    spec.args.model_fields, op=spec.name
                ),
                "required_fields": [
                    name for name, field in spec.args.model_fields.items()
                    if field.is_required()
                ],
            },
        ) from None
    return model.model_dump()


def _active_surface(result: dict[str, Any], plan: SurfacePlan, op: str) -> dict[str, Any]:
    """Bind status operations to the same frozen plan that handled the call."""
    if op not in {"get_server_info", "diagnose"}:
        return result
    counts = {
        name: len(tool.ops)
        for name, tool in plan.tools.items()
        if name != HELP_TOOL
    }
    surface = {
        "fixed": True,
        "tools": counts,
        "operations": sum(counts.values()),
        "read_only": plan.settings.read_only,
        "forbid": list(plan.settings.forbid),
        "legacy_settings_ignored": list(plan.settings.ignored_legacy),
        "approx_active_tokens": approx_active_tokens(plan),
    }
    result = {**result, "surface": surface}
    if op == "get_server_info":
        result["tools_registered"] = len(plan.tools)
        result["operations_registered"] = surface["operations"]
    return result


def approx_active_tokens(plan: SurfacePlan) -> int:
    """Approximate the compact tools/list bill by the measurement method."""
    total = 0
    for tool in plan.tools.values():
        wire = {
            "name": tool.name,
            "title": tool.title,
            "description": tool.description,
            "inputSchema": tool.input_schema,
            "annotations": tool.annotations,
        }
        total += round(len(json.dumps(
            wire, ensure_ascii=False, separators=(",", ":")
        )) / 4)
    return total


async def run_bounded(fn: Callable[[], dict[str, Any]], timeout_s: float) -> dict[str, Any]:
    """Bound one call and prevent a timed-out worker from publishing later."""
    deadline = call_deadline.Deadline(timeout_s)

    def invoke() -> dict[str, Any]:
        with call_deadline.active(deadline):
            return fn()

    task = _start_worker(invoke, deadline)
    try:
        done, _pending = await asyncio.wait({task}, timeout=timeout_s)
        if done:
            return task.result()
        if deadline.cancel():
            raise deadline.timeout_error("the operation") from None
        # The operation crossed its final publication point before the timer.
        # Give it one finite second chance to report the real result. If that
        # also expires, the answer must acknowledge that the change started.
        done, _pending = await asyncio.wait(
            {task}, timeout=max(MIN_POST_COMMIT_WAIT_S, float(timeout_s))
        )
        if done:
            return task.result()
        raise OpError(
            "TIMEOUT",
            "the operation ran past its limit after its change started; "
            "completion is unknown, so check the target before retrying",
            details={"bound_s": timeout_s, "saved": None},
        ) from None
    except asyncio.CancelledError:
        deadline.cancel()
        task.cancel()
        raise


def _with_notices(payload: dict[str, Any]) -> dict[str, Any]:
    """Add to an answer what a recovery left for the person to look at
    (an older copy of a deck kept beside it). Each notice is given once."""
    notices = _recovery_notices()
    if notices:
        payload = {**payload, "notices": notices}
    return payload


async def dispatch(
    plan: SurfacePlan,
    tool: str,
    arguments: Any,
    *,
    invoke: Callable[[OpSpec, dict[str, Any]], dict[str, Any]] | None = None,
) -> tuple[dict[str, Any], bool]:
    """Run the foundation pipeline and return one flat payload plus error flag."""
    probe: slide_echo.Probe | None = None
    armed = None
    try:
        op, args, normalized_from = shape(plan, arguments)
        if tool == HELP_TOOL:
            raise OpError("INTERNAL_ERROR", "ppt_help is not a group tool")
        spec = resolve(plan, tool, op)
        args, decoded = decode_json_text(spec, args)
        args, sent = normalize.normalize_args(
            args, spec.args.model_fields, op=spec.name
        )
        args, sent_slide, slide_form = normalize.slide_args(
            args, spec.args.model_fields
        )
        sent_under = {
            canonical: synonym
            for synonym, canonical in normalize.synonyms_for(spec.name).items()
            if synonym in sent
        }
        args, twin_forms = normalize.twin_args(
            args, spec.args.model_fields, op=spec.name, renamed=sent_under
        )
        args, sent_enum = normalize.enum_fix(_schema(spec.args), args)
        if sent or sent_slide or sent_enum:
            normalized_from["args"] = {**sent, **sent_slide, **sent_enum}
        renamed = normalize.renamed_hint(spec.name, sent)
        if renamed:
            normalized_from["hint"] = renamed
        try:
            payload = validate(spec, args)
        except OpError as exc:
            if decoded:
                if not isinstance(exc.details, dict):
                    exc.details = {}
                exc.details.setdefault("json_text_decoded", decoded)
            raise
        runner = invoke or (lambda chosen, values: chosen.handler(**values))
        if "slide" in spec.args.model_fields:
            probe, armed = slide_echo.arm(payload.get("slide"), slide_form, spec.name)
        elif spec.name in slide_echo.SCOPE_OPS:
            probe, armed = slide_echo.arm(
                payload.get("scope"),
                twin_forms.get("scope", sent_under.get("scope", "scope")),
                spec.name,
            )
        try:
            # A read-only operation opens decks to read: one PowerPoint has
            # open is read from its saved file, never refused for it.
            with _package.reading_call(spec.read_only) as held_reads:
                result = await run_bounded(
                    lambda: runner(spec, payload),
                    spec.timeout_s or plan.settings.call_timeout_s,
                )
        finally:
            slide_echo.disarm(armed)
            armed = None
        if not isinstance(result, dict):
            raise OpError("INTERNAL_ERROR", f"{op} returned a non-object result")
        result = slide_echo.lead(result, probe)
        held_note = _package.open_deck_note(held_reads)
        if held_note and isinstance(result, dict):
            result = {**result, "open_deck_note": held_note}
        listed = _ZERO_BASED_LISTS.get(spec.name)
        sent_list = payload.get(listed) if listed else None
        if (
            isinstance(sent_list, list) and sent_list and listed not in twin_forms
            and all(isinstance(n, int) and not isinstance(n, bool) for n in sent_list)
        ):
            result = {**result, f"{listed}_note": copy_messages.SLIDES_NOTE.format(
                sent=sent_list, numbers=[n + 1 for n in sent_list])}
        result = _active_surface(result, plan, op)
        if decoded:
            result = {**result, "json_text_decoded": decoded}
        if normalized_from:
            result = {**result, "normalized_from": normalized_from}
        try:
            result = budget.enforce(
                result,
                plan.settings.max_output_chars,
                wrap=envelope.success,
                supports_offset="offset" in spec.args.model_fields,
                requested_offset=int(payload.get("offset") or 0),
            )
        except OpError as exc:
            # A write has already happened by now. Answering LIMIT_EXCEEDED
            # would tell the client a saved change failed, so a writing
            # operation's answer is shortened instead, never refused.
            if exc.code != "LIMIT_EXCEEDED" or spec.read_only:
                raise
            result = budget.committed(
                result,
                plan.settings.max_output_chars,
                wrap=envelope.success,
                note=copy_messages.COMMITTED_OUTPUT_NOTE,
            )
        # A long answer leads with its totals and notes: a client may show
        # only its first 2,000 characters.
        return budget.lead_with_totals(_with_notices(envelope.success(result))), False
    except asyncio.CancelledError:
        raise
    except BaseException as exc:  # every ordinary failure becomes one refusal
        if isinstance(exc, (KeyboardInterrupt, SystemExit)):
            raise
        payload = envelope.refusal(exc)
        if "normalized_from" in locals() and normalized_from:
            payload["error"].setdefault("details", {}).setdefault(
                "normalized_from", normalized_from
            )
        if probe is not None and probe.hit is not None:
            # A refusal after the slide was found still says which slide.
            details = payload["error"].setdefault("details", {})
            for key, value in probe.fields(done=False).items():
                details.setdefault(key, value)
        return _with_notices(
            budget.cap_refusal(payload, plan.settings.max_output_chars)), True


__all__ = [
    "approx_active_tokens", "decode_json_text", "dispatch", "drain_at_exit_once",
    "drain_inflight", "has_inflight", "parse", "resolve", "run_bounded", "shape",
    "unknown_op", "validate",
]
