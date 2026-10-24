"""Unambiguous call-shape normalization for the fixed PowerPoint surface.

Only exact folded spellings are routed.  Every accepted change is returned
to dispatch so it can be echoed as ``normalized_from``.  Unknown keys and
unknown values remain untouched for strict validation to refuse.
"""

from __future__ import annotations

import json
import re
from typing import Any

from kitchensink4ppt.core.errors import OpError


FILE_SYNONYMS = ("path", "file", "file_path")
ARG_SYNONYMS = {
    "set_theme_fonts": {
        "heading": "major",
        "headings": "major",
        "body": "minor",
    },
    "list_elements": {
        "type": "kind",
    },
    # The list answer groups prev and anchor under "slots"; the restore
    # argument is source, because it also takes a snapshot's path.
    "manage_backups": {
        "slot": "source",
    },
}
#: Names a model reaches for first, read for every operation they fit:
#: shape_id where the operation takes shape (get_slide_info answers with id,
#: other answers say shape_id), and font_size for the text operations' font
#: size, whatever that argument is called there. An operation that has an
#: argument of the synonym's own name keeps it (delete_master_shape).
ARG_SYNONYMS["*"] = {"shape_id": "shape"}
for _op in ("format_text", "insert_textbox"):
    ARG_SYNONYMS[_op] = {**ARG_SYNONYMS.get(_op, {}), "font_size": "size_pt"}
for _op in ("set_master_placeholder", "set_layout_placeholder"):
    ARG_SYNONYMS[_op] = {**ARG_SYNONYMS.get(_op, {}), "font_size": "size", "size_pt": "size"}
del _op

#: The operations whose scope selects slides also read slides as another
#: name for scope, counted from 0 like it: get_presentation_view {slides:
#: [28]} was refused when scope was meant.
ARG_SYNONYMS.update({
    op: {**ARG_SYNONYMS.get(op, {}), "slides": "scope"}
    for op in (
        "get_presentation_view", "get_text", "find_text", "list_elements",
        "list_hyperlinks", "list_equations", "list_comments",
        "audit_accessibility", "search_and_replace",
    )
})
DESTRUCTIVE_VERBS = ("delete", "clear", "redact", "anonymiz", "purge", "overwrite")


def fold(value: str) -> str:
    """Fold case plus spaces, underscores, and hyphens."""
    return re.sub(r"[\s_\-]+", "", value).casefold()


def op_keys(ops) -> dict[str, str]:
    """Return folded-name -> canonical-name, rejecting any collision."""
    keys: dict[str, str] = {}
    for name in ops:
        key = fold(name)
        if key in keys and keys[key] != name:
            raise ValueError(
                f"operation names {keys[key]} and {name} fold to the same key {key}"
            )
        keys[key] = name
    return keys


def route_op(op: str, ops) -> str | None:
    """Resolve only an exact canonical or exact folded operation spelling."""
    if op in ops:
        return op
    return op_keys(ops).get(fold(op))


def _distance(a: str, b: str) -> int:
    previous = list(range(len(b) + 1))
    for i, left in enumerate(a, 1):
        current = [i]
        for j, right in enumerate(b, 1):
            current.append(min(
                previous[j] + 1,
                current[j - 1] + 1,
                previous[j - 1] + (left != right),
            ))
        previous = current
    return previous[-1]


def suggestions(op: str, specs, limit: int = 3) -> list[str]:
    """Return close visible names without inventing a route.

    A destructive operation is not suggested unless the attempted spelling
    itself contains its destructive verb.
    """
    query = fold(op)
    scored: list[tuple[int, str]] = []
    for name, spec in specs.items():
        folded = fold(name)
        if spec.destructive and not any(
            verb in query and verb in folded for verb in DESTRUCTIVE_VERBS
        ):
            continue
        distance = _distance(query, folded)
        if query and (query in folded or folded in query):
            distance = min(distance, abs(len(folded) - len(query)) // 2)
        scored.append((distance, name))
    scored.sort()
    cutoff = max(3, len(query) // 2)
    return [name for distance, name in scored if distance <= cutoff][:limit]


def flatten(
    arguments: dict[str, Any], fields, *, op: str | None = None
) -> tuple[dict[str, Any], list[str]]:
    """Move only known flat fields and file aliases into ``args``.

    Unknown top-level keys remain in place so the strict outer schema refuses
    them.  A duplicate name also remains untouched rather than guessing which
    value the caller intended.
    """
    extra = [name for name in arguments if name not in ("op", "args")]
    movable = [
        name for name in extra
        if name in fields or name in FILE_SYNONYMS
        or ("slide" in fields and name in SLIDE_ALIASES)
        or any(name == twin for _f, twin, _k in twins_for(op, fields))
    ]
    if not movable:
        return arguments, []
    existing = arguments.get("args")
    if existing is not None and not isinstance(existing, dict):
        return arguments, []
    args = dict(existing or {})
    if any(name in args for name in movable):
        return arguments, []
    output = {name: value for name, value in arguments.items() if name not in movable}
    for name in movable:
        args[name] = arguments[name]
    output["args"] = args
    return output, movable


def _file_arg(fields) -> str | None:
    for name in ("file_path", "target_path", "template_path", "path"):
        if name in fields:
            return name
    return None


def synonyms_for(op: str | None) -> dict[str, str]:
    """The other names one operation reads, synonym -> its own argument."""
    return {**ARG_SYNONYMS.get("*", {}), **ARG_SYNONYMS.get(op or "", {})}


def normalize_args(
    args: dict[str, Any], fields, *, op: str | None = None
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Normalize file aliases before schema validation."""
    output = dict(args)
    sent: dict[str, Any] = {}
    target = _file_arg(fields)
    for synonym in FILE_SYNONYMS:
        if synonym in output and synonym not in fields and target and target not in output:
            sent[synonym] = output[synonym]
            output[target] = output.pop(synonym)
    for synonym, canonical in synonyms_for(op).items():
        if synonym not in output or synonym in fields or canonical not in fields:
            continue
        theirs, ours = output[synonym], output.get(canonical)
        if theirs is not None and ours is not None and theirs != ours:
            # Two names for one argument, with two values: refuse rather
            # than pick one.
            raise OpError(
                "BAD_ARGS",
                f"{synonym} and {canonical} were both sent, with different "
                f"values; {op} reads {synonym} as another name for "
                f"{canonical}, so send one of them",
                hint=(f"Send {canonical} alone; {op} names this argument "
                      f"{canonical}."),
                details={"fields": [synonym, canonical]},
            )
        sent[synonym] = output.pop(synonym)
        if ours is None:
            output[canonical] = theirs
    return output, sent


def renamed_hint(op: str | None, sent: dict[str, Any]) -> str | None:
    """For each argument read under another name, a sentence naming the
    argument the operation takes, so the next call can use it."""
    pairs = [(synonym, canonical)
             for synonym, canonical in synonyms_for(op).items()
             if synonym in sent]
    if not pairs:
        return None
    return " ".join(
        f"`{synonym}` was read as `{canonical}`; {op} names this argument "
        f"`{canonical}`." for synonym, canonical in pairs)


#: The 1-based alternative to a ``slides`` list, as people count slides.
SLIDE_NUMBERS = "slide_numbers"
_SLIDES_HINT = (
    "Send slide_numbers counted from 1 ([1, 2, 4] is the first, second and "
    "fourth slides) or slides counted from 0 ([0, 1, 3] names the same "
    "slides), not both."
)

#: Operations whose ``scope`` selects slides (core/slide_echo.SCOPE_OPS).
#: set_footer is not one: its own slide_number argument means "show the
#: slide number".
_SCOPE_OPS = (
    "get_presentation_view", "get_text", "find_text", "list_elements",
    "list_hyperlinks", "list_equations", "list_comments", "audit_accessibility",
    "search_and_replace",
)

#: The 1-based twin of every argument that counts slides from 0, by
#: operation ("*" wherever the field is): (field, twin, kind). kind "one" is
#: a number, "list" a list of numbers, "ranges" split_deck's
#: {"start", "end", "name"} with start and end counted from 1.
TWINS: dict[str, tuple[tuple[str, str, str], ...]] = {
    "*": (("slides", SLIDE_NUMBERS, "list"),),
    "copy_position": (
        ("from_slide", "from_slide_number", "one"),
        ("to_slides", "to_slide_numbers", "list"),
    ),
    "reorder_slides": (("order", "order_numbers", "list"),),
    "split_deck": (("ranges", "number_ranges", "ranges"),),
    "set_hyperlink": (("to_slide", "to_slide_number", "one"),),
    "move_slide": (("to", "to_number", "one"),),
    "insert_slide": (("position", "position_number", "one"),),
    "duplicate_slide": (("position", "position_number", "one"),),
    "copy_slide_between": (("position", "position_number", "one"),),
    "generate_agenda_slide": (("position", "position_number", "one"),),
    "copy_shape_between": (("source_slide", "source_slide_number", "one"),),
    **{op: (("scope", "slide_number", "one"), ("scope", SLIDE_NUMBERS, "scope"))
       for op in _SCOPE_OPS},
}


def twins_for(op: str | None, fields) -> list[tuple[str, str, str]]:
    """The (field, twin, kind) rows that apply to one operation: the field
    is one of its arguments and the twin is not."""
    rows = [*TWINS.get(op or "", ()), *TWINS["*"]]
    out: list[tuple[str, str, str]] = []
    for field, twin, kind in rows:
        if field in fields and twin not in fields and all(t != twin for _f, t, _k in out):
            out.append((field, twin, kind))
    return out


def _twin_hint(field: str, twin: str, kind: str) -> str:
    if (field, twin) == ("slides", SLIDE_NUMBERS):
        return _SLIDES_HINT
    if kind in ("list", "scope"):
        return (f"Send {twin} counted from 1 ([1, 2] is the first and second slides) "
                f"or {field} counted from 0 ([0, 1] names the same slides), not both.")
    if kind == "ranges":
        return (f"Send {twin} with start and end counted from 1 or {field} with them "
                "counted from 0, not both.")
    return (f"Send {twin} counted from 1 (2 is the second slide) or {field} counted "
            "from 0 (1 names the same slide), not both.")


def _from_one(raw: Any, twin: str, kind: str, hint: str) -> Any:
    """A twin's value counted from 0, or a refusal naming the rule."""
    if isinstance(raw, str) and raw.strip()[:1] in ("[", "{"):
        try:
            raw = json.loads(raw)
        except ValueError:
            pass
    if kind == "one":
        number = _whole(raw)
        if number is None or number < 1:
            raise OpError(
                "BAD_ARGS",
                f"{twin} is a whole number counted from 1: the first slide is 1",
                hint=hint, details={"field": twin},
            )
        return number - 1
    if kind == "scope":
        # Numbers counted from 1, ranges counted from 1, or a list of both.
        if isinstance(raw, dict):
            return _range_from_one(raw, twin, hint)
        if not isinstance(raw, (list, tuple)):
            raw = [None]
        out = []
        for item in raw:
            if isinstance(item, dict):
                out.append(_range_from_one(item, twin, hint))
                continue
            number = _whole(item)
            if number is None or number < 1:
                raise OpError(
                    "BAD_ARGS",
                    f"{twin} is a list of whole numbers counted from 1, or a range "
                    '{"start", "end"} counted from 1: [1, 2, 4] is the first, second and '
                    'fourth slides, and {"start": 1, "end": 3} the first three',
                    hint=hint, details={"field": twin},
                )
            out.append(number - 1)
        return out
    if kind == "list":
        numbers = ([_whole(item) for item in raw]
                   if isinstance(raw, (list, tuple)) else [None])
        if any(number is None or number < 1 for number in numbers):
            message = (
                "slide_numbers is a list of whole numbers counted from 1: the first "
                "slide is 1, so [1, 2, 4] is the first, second and fourth slides"
                if twin == SLIDE_NUMBERS else
                f"{twin} is a list of whole numbers counted from 1: the first slide is 1"
            )
            raise OpError("BAD_ARGS", message, hint=hint, details={"field": twin})
        return [number - 1 for number in numbers]
    ranges = raw if isinstance(raw, (list, tuple)) else None
    out = []
    for item in ranges or [None]:
        bounds = ([_whole(item.get(key)) for key in ("start", "end")]
                  if isinstance(item, dict) else [None])
        if any(number is None or number < 1 for number in bounds):
            raise OpError(
                "BAD_ARGS",
                f'{twin} is a list of {{"start", "end", "name"}} with start and end '
                'counted from 1, both included: {"start": 1, "end": 3} is the first '
                "three slides",
                hint=hint, details={"field": twin},
            )
        out.append({**item, "start": bounds[0] - 1, "end": bounds[1] - 1})
    return out


def _range_from_one(item: dict, twin: str, hint: str) -> dict:
    """A {"start", "end"} range counted from 1, as one counted from 0."""
    bounds = [_whole(item.get(key)) for key in ("start", "end")] if set(item) == {
        "start", "end"} else [None]
    if any(number is None or number < 1 for number in bounds):
        raise OpError(
            "BAD_ARGS",
            f'{twin} takes a range as {{"start", "end"}} counted from 1, both '
            'included: {"start": 1, "end": 3} is the first three slides',
            hint=hint, details={"field": twin},
        )
    return {"start": bounds[0] - 1, "end": bounds[1] - 1}


def twin_args(
    args: dict[str, Any], fields, *, op: str | None = None,
    renamed: dict[str, str] | None = None,
) -> tuple[dict[str, Any], dict[str, str]]:
    """Resolve each 1-based twin into the argument it stands for, counted
    from 0. Returns the new args and, by field, the twin it was sent as.
    Sending a field and its twin (or two twins of one field) is refused
    rather than guessed between; a null twin is read as not sent, as every
    optional argument is. `renamed` maps an argument to the other name the
    caller sent it under (slides for scope), so a refusal names what was
    sent."""
    renamed = renamed or {}
    rows = twins_for(op, fields)
    if not any(twin in args for _f, twin, _k in rows):
        return args, {}
    output = dict(args)
    forms: dict[str, str] = {}
    for field, twin, kind in rows:
        if twin not in output:
            continue
        raw = output.pop(twin)
        if raw is None:
            continue
        hint = _twin_hint(renamed.get(field, field), twin, kind)
        if field in forms:
            raise OpError(
                "BAD_ARGS",
                f"{forms[field]} and {twin} were both sent; send one of them",
                hint=hint, details={"fields": [forms[field], twin]},
            )
        if output.get(field) is not None:
            sent_as = renamed.get(field, field)
            raise OpError(
                "BAD_ARGS",
                f"{sent_as} and {twin} were both sent; send one of them",
                hint=hint, details={"fields": [sent_as, twin]},
            )
        output[field] = _from_one(raw, twin, kind, hint)
        forms[field] = twin
    return output, forms


def slides_args(args: dict[str, Any], fields) -> dict[str, Any]:
    """Resolve ``slide_numbers`` into ``slides`` wherever an operation takes
    a ``slides`` list (twin_args with no operation named)."""
    return twin_args(args, fields)[0]


#: The 1-based alternative to ``slide``, as people count slides.
SLIDE_NUMBER = "slide_number"
#: Other names a caller reaches for, each with exactly one reading.
_SLIDE_SPELLINGS = ("slide_index", "slide_id")
SLIDE_ALIASES = (SLIDE_NUMBER, *_SLIDE_SPELLINGS)
_DIGITS = re.compile(r"\s*\d+\s*")


def _whole(value: Any) -> int | None:
    """An int, or a string of digits, as an int; anything else is None."""
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, str) and _DIGITS.fullmatch(value):
        return int(value)
    return None


def _slide_text(value: Any) -> Any:
    """Decode a slide selector a caller sent as text: "1" or '{"slide_id": 257}'."""
    if not isinstance(value, str):
        return value
    number = _whole(value)
    if number is not None:
        return number
    text = value.strip()
    if text.startswith("{"):
        try:
            parsed = json.loads(text)
        except (ValueError, RecursionError):
            return value
        if (
            isinstance(parsed, dict) and set(parsed) == {"slide_id"}
            and _whole(parsed["slide_id"]) is not None
        ):
            return {"slide_id": _whole(parsed["slide_id"])}
    return value


def slide_args(
    args: dict[str, Any], fields
) -> tuple[dict[str, Any], dict[str, Any], str | None]:
    """Resolve the ways a caller can name one slide into ``slide``.

    ``slide`` counts from 0. ``slide_number`` counts from 1, the way people
    count, so slide_number 2 is slide 1. ``slide_index`` is the 0-based
    name the results use, and ``slide_id`` is the durable id. Two names for
    the same slide are fine; two that disagree refuse rather than guess.
    Returns the new args, what was reinterpreted (for ``normalized_from``;
    ``slide_number`` is a documented argument, so it is not listed there)
    and the name the slide was sent under, slide_number first when given.
    A name sent as null is treated as not sent, as every optional argument
    is: clients that fill unused fields with null are common.
    """
    if "slide" not in fields:
        return args, {}, None
    output = dict(args)
    sent: dict[str, Any] = {}
    named: list[tuple[str, Any]] = []
    if output.get("slide") is not None:
        decoded = _slide_text(output["slide"])
        if decoded is not output["slide"]:
            sent["slide"] = output["slide"]
            output["slide"] = decoded
        named.append(("slide", output["slide"]))
    if SLIDE_NUMBER in output and SLIDE_NUMBER not in fields:
        raw = output.pop(SLIDE_NUMBER)
        if raw is not None:
            number = _whole(raw)
            if number is None or number < 1:
                raise OpError(
                    "BAD_ARGS",
                    "slide_number counts from 1: the first slide is slide_number 1",
                    hint=_SLIDE_HINT,
                    details={"field": SLIDE_NUMBER},
                )
            named.append((SLIDE_NUMBER, number - 1))
    for name in _SLIDE_SPELLINGS:
        if name in output and name not in fields:
            raw = output.pop(name)
            if raw is None:
                continue
            number = _whole(raw)
            if number is None or number < 0:
                raise OpError(
                    "BAD_ARGS",
                    f"{name} must be a whole number; {_SLIDE_RULE}",
                    hint=_SLIDE_HINT,
                    details={"field": name},
                )
            sent[name] = raw
            named.append((name, number if name == "slide_index" else {"slide_id": number}))
    if not named:
        return output, sent, None
    first, chosen = named[0]
    for name, value in named[1:]:
        if value == chosen:
            continue
        both_indexes = all(
            isinstance(item, int) and not isinstance(item, bool)
            for item in (value, chosen)
        )
        reason = (
            f"{first} and {name} name different slides; {_SLIDE_RULE}"
            if both_indexes
            else f"{first} and {name} each name a slide"
        )
        raise OpError(
            "BAD_ARGS",
            f"{reason}. Send one of them.",
            hint=_SLIDE_HINT,
            details={"fields": [item[0] for item in named]},
        )
    output["slide"] = chosen
    names = [item[0] for item in named]
    form = SLIDE_NUMBER if SLIDE_NUMBER in names else names[0]
    return output, sent, form


def accepted_fields(fields, op: str | None = None) -> list[str]:
    """The argument names an operation accepts, slide_number and every
    1-based twin included, each right after the argument it stands for."""
    names = list(fields)
    if "slide" in names and SLIDE_NUMBER not in names:
        names.insert(names.index("slide") + 1, SLIDE_NUMBER)
    rows = twins_for(op, fields)
    for field, twin, _kind in rows:
        if twin in names:
            continue
        at = names.index(field) + 1
        while at < len(names) and any(names[at] == t for f, t, _k in rows if f == field):
            at += 1
        names.insert(at, twin)
    return names


_SLIDE_RULE = (
    "slide counts from 0 and slide_number from 1, so the second slide is "
    "slide 1 or slide_number 2"
)
_SLIDE_HINT = (
    "Name one slide with slide_number (counting from 1), or slide (counting "
    'from 0), or {"slide_id": N}.'
)


def enum_fix(schema: dict[str, Any], args: dict[str, Any]) -> tuple[dict[str, Any], dict]:
    """Match string enum values without regard to case or separators."""
    properties = schema.get("properties") or {}
    output = dict(args)
    sent: dict[str, Any] = {}
    for name, value in args.items():
        if not isinstance(value, str) or name not in properties:
            continue
        values = _enum_values(properties[name], schema)
        if not values or value in values:
            continue
        hit = {fold(candidate): candidate for candidate in values
               if isinstance(candidate, str)}.get(fold(value))
        if hit is not None:
            sent[name] = value
            output[name] = hit
    return output, sent


def _enum_values(prop: dict, root: dict) -> list:
    if "$ref" in prop:
        prop = (root.get("$defs") or {}).get(prop["$ref"].rsplit("/", 1)[-1], {})
    if "enum" in prop:
        return list(prop["enum"])
    if "const" in prop:
        return [prop["const"]]
    for key in ("anyOf", "oneOf"):
        for child in prop.get(key) or []:
            values = _enum_values(child, root)
            if values:
                return values
    return []


__all__ = [
    "SLIDE_NUMBER", "TWINS", "enum_fix", "flatten", "fold", "normalize_args", "op_keys",
    "route_op", "slide_args", "suggestions", "twin_args", "twins_for",
]
