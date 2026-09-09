"""Output budget: a read tool never answers with more than it is allowed to.

Why this module exists. Before it, nothing in the read surface carried a
limit, a page size, or a ceiling of any kind. `list_elements(kind="shapes")`
on a 131-slide deck returned 743,972 characters in one response, roughly
186,000 tokens, which is larger than most models' entire context window. One
innocent call ended the session. `get_text` and `get_presentation_view` on
the same deck returned 108,745 and 163,207 characters. The `scope` parameter
let a caller narrow the read, but only if the caller already knew the deck
was large, and finding that out was not a signposted first step.

The contract, in both directions:

1. NEVER EXCEEDED. The serialized payload of a budgeted list or block
   sequence stays at or under `max_chars()`. This is measured on the same
   JSON text the MCP layer sends, not estimated.

2. NOTHING SILENTLY DROPPED. When a read is cut short, the answer says so
   in a `page` block that names the true total, how many items came back,
   how many were omitted, and the exact `offset` that continues the read.
   A caller who ignores `page` still cannot be misled: `count` remains the
   true total, and `returned` never claims more than it delivered.

3. SMALL READS UNCHANGED. When the whole answer fits and the caller asked
   for no paging, no `page` block is added and the payload is exactly what
   the tool returned before this module existed. Paying attention to the
   budget costs nothing on a deck that does not need it.

The ceiling is `KS4P_MAX_OUTPUT_CHARS`, default 60,000 characters (about
15,000 tokens). It is a character count rather than a token count on
purpose: characters are exact and free to measure, tokens are neither, and
a ratio of four is close enough for a ceiling whose job is to prevent a
catastrophe rather than to fill a window precisely. An unparseable value
refuses at startup rather than silently reverting to the default, on the
same reasoning as packs.parse_toggle: a typo that quietly restores an
unbounded read is the failure this module was written to prevent.

Compaction is separate and opt-in (`compact=True` on the list tools). It
changes the shape of the answer, from a list of objects to a `fields` header
plus one array per item, which drops the repeated JSON key names that were
65% of `list_elements`' payload and 48% of `find_text`'s. It is lossless:
every field of every item is still there, in a named column.
"""

from __future__ import annotations

import json
import os

from .errors import PptMcpError

#: Default per-call output ceiling in characters (~15,000 tokens at the
#: 4-chars-per-token approximation the server uses elsewhere).
DEFAULT_MAX_CHARS = 60_000

#: Floor for the env override. Below this a budget cannot hold even one
#: substantial item plus its continuation note, so it would page forever.
MIN_MAX_CHARS = 2_000

ENV_MAX_CHARS = "KS4P_MAX_OUTPUT_CHARS"


def max_chars() -> int:
    """The per-call output ceiling. `KS4P_MAX_OUTPUT_CHARS` overrides the
    default; anything that is not a positive integer of at least
    MIN_MAX_CHARS is refused rather than shrugged off, because a typo that
    silently restored an unbounded read is exactly the failure this budget
    exists to prevent."""
    raw = os.environ.get(ENV_MAX_CHARS)
    if raw is None or not str(raw).strip():
        return DEFAULT_MAX_CHARS
    text = str(raw).strip()
    try:
        value = int(text)
    except ValueError:
        raise PptMcpError(
            f"unknown {ENV_MAX_CHARS} value {raw!r}: use a whole number of "
            f"characters (at least {MIN_MAX_CHARS}), or unset it for the "
            f"{DEFAULT_MAX_CHARS} default. Refusing to start rather than "
            "guessing, because guessing here would restore the unbounded "
            "read this setting bounds."
        ) from None
    if value < MIN_MAX_CHARS:
        raise PptMcpError(
            f"{ENV_MAX_CHARS}={value} is below the {MIN_MAX_CHARS}-character "
            "floor; a budget that small cannot return one item plus its "
            "continuation note."
        )
    return value


def _size(obj) -> int:
    """Serialized size of one item, measured the way the wire measures it."""
    return len(json.dumps(obj, ensure_ascii=False, default=str))


#: Characters reserved for the `page` block, its key, and the container key
#: holding the items, so an answer that spends its whole budget on items
#: still has room to explain what it left out. The largest page block
#: observed is about 320 characters (eight fields plus the continuation
#: hint); the rest is headroom, because a ceiling that is occasionally
#: generous is right and a ceiling that is once exceeded is not.
PAGE_BLOCK_RESERVE = 640


def overhead_of(header: dict) -> int:
    """Budget already spent before the first item: the answer's own header
    fields plus room for the continuation note."""
    return _size(header) + PAGE_BLOCK_RESERVE


def normalize_window(limit, offset, *, what: str = "items") -> tuple[int | None, int]:
    """Validate caller-supplied paging. Refuses loudly: a negative offset
    silently clamped to zero would return the wrong page and say nothing."""
    if offset is None:
        offset = 0
    if not isinstance(offset, int) or isinstance(offset, bool) or offset < 0:
        raise PptMcpError(
            f"offset must be a whole number of {what} at or above 0, got "
            f"{offset!r}"
        )
    if limit is not None:
        if not isinstance(limit, int) or isinstance(limit, bool) or limit < 1:
            raise PptMcpError(
                f"limit must be a whole number of {what} at or above 1, or "
                f"null for as many as the budget allows, got {limit!r}"
            )
    return limit, offset


#: What a truncated string carries in its own tail, so the reader learns the
#: cut from the text itself and not only from the `page` block.
TRUNCATION_MARK = "\n[... {n} more characters not returned; budget reached]"


def truncate_text(text: str, allowance: int) -> tuple[str, int]:
    """Cut one oversized string to fit, and say in the string how much went.
    Returns (cut_text, characters_omitted); allowance is a JSON-serialized
    character allowance, so the cut is conservative."""
    mark_len = len(TRUNCATION_MARK.format(n=len(text)))
    keep = max(0, allowance - mark_len - 2)  # -2 for the string's own quotes
    if keep >= len(text):
        return text, 0
    dropped = len(text) - keep
    return text[:keep] + TRUNCATION_MARK.format(n=dropped), dropped


def shrink_record(item, allowance: int):
    """Default shrinker for a record: cut its longest text field down to
    fit, marking the cut in the text. A record whose bulk is not text (a
    SmartArt frame carrying a hundred nodes, say) comes back whole and the
    `page` block reports it as oversized, because silently dropping half a
    structure would be worse than one honest overrun."""
    if not isinstance(item, dict):
        return item, 0
    longest, size = None, 0
    for key, value in item.items():
        if isinstance(value, str) and len(value) > size:
            longest, size = key, len(value)
    if longest is None:
        return item, 0
    fixed = _size({k: v for k, v in item.items() if k != longest})
    text, dropped = truncate_text(item[longest], allowance - fixed - len(longest) - 6)
    if not dropped:
        return item, 0
    new = dict(item)
    new[longest] = text
    return new, dropped


def page_items(
    items: list,
    *,
    limit=None,
    offset=0,
    budget: int | None = None,
    overhead: int = 0,
    unit: str = "items",
    narrow_hint: str = "",
    shrink=None,
) -> tuple[list, dict | None]:
    """Cut `items` down to what fits, and say exactly what was left out.

    Returns (kept, page). `page` is None when the caller asked for no window
    and the whole list fit, so an answer that never needed budgeting is
    byte-identical to what it was before the budget existed.

    One item always comes back, because a page of nothing tells a caller
    less than a page of one. When that one item is by itself larger than
    the whole budget, `shrink(item, allowance)` gets a chance to cut it
    down and report how much it dropped; readers whose items are text pass
    one. Without a shrinker the single oversized item is returned whole and
    the `page` block says so under `oversized`, so the overrun is named
    rather than discovered.
    """
    limit, offset = normalize_window(limit, offset, what=unit)
    total = len(items)
    ceiling = max_chars() if budget is None else budget

    window = items[offset:] if offset else items
    if limit is not None:
        window = window[:limit]

    spent = overhead
    kept: list = []
    stopped_on_budget = False
    oversized = 0
    shrunk = 0
    for item in window:
        # +2 for the ", " that json.dumps puts between array elements. It
        # was +1 first, which undercounted by one character per item and
        # let a 5,714-shape answer land 500 characters over its ceiling.
        cost = _size(item) + 2
        if not kept and spent + cost > ceiling:
            if shrink is not None:
                item, shrunk = shrink(item, ceiling - spent - 2)
                cost = _size(item) + 2
            if spent + cost > ceiling:
                oversized = cost
            kept.append(item)
            spent += cost
            stopped_on_budget = True
            break
        if kept and spent + cost > ceiling:
            stopped_on_budget = True
            break
        kept.append(item)
        spent += cost

    returned = len(kept)
    omitted = total - offset - returned
    if omitted < 0:
        omitted = 0
    windowed = bool(offset) or limit is not None
    if not windowed and not stopped_on_budget and returned == total:
        return kept, None

    reason = "budget" if stopped_on_budget else ("limit" if limit is not None else "offset")
    page = {
        "unit": unit,
        "total": total,
        "offset": offset,
        "returned": returned,
        "omitted": omitted,
        "truncated": omitted > 0,
        "reason": reason,
        "next_offset": (offset + returned) if omitted > 0 else None,
    }
    if shrunk:
        page["chars_cut_from_last"] = shrunk
    if oversized:
        page["oversized"] = oversized
    if omitted > 0:
        page["hint"] = _continue_hint(offset + returned, omitted, unit, narrow_hint)
    return kept, page


def page_blocks(
    blocks: list[str],
    *,
    limit=None,
    offset=0,
    budget: int | None = None,
    overhead: int = 0,
    unit: str = "slides",
    narrow_hint: str = "",
) -> tuple[list[str], dict | None]:
    """The same contract for a sequence of text blocks (the markdown view).
    Blocks are still measured as JSON, because a markdown block full of
    newlines and quotes costs more on the wire than its raw length: every
    newline is two characters once escaped, and undercounting that is how a
    budget quietly gets exceeded."""
    limit, offset = normalize_window(limit, offset, what=unit)
    total = len(blocks)
    ceiling = max_chars() if budget is None else budget

    window = blocks[offset:] if offset else blocks
    if limit is not None:
        window = window[:limit]

    spent = overhead
    kept: list[str] = []
    stopped_on_budget = False
    shrunk = 0
    for block in window:
        cost = _size(block) + 1
        if not kept and spent + cost > ceiling:
            block, shrunk = truncate_text(block, ceiling - spent - 1)
            kept.append(block)
            stopped_on_budget = True
            break
        if kept and spent + cost > ceiling:
            stopped_on_budget = True
            break
        kept.append(block)
        spent += cost

    returned = len(kept)
    omitted = max(0, total - offset - returned)
    windowed = bool(offset) or limit is not None
    if not windowed and not stopped_on_budget and returned == total:
        return kept, None

    reason = "budget" if stopped_on_budget else ("limit" if limit is not None else "offset")
    page = {
        "unit": unit,
        "total": total,
        "offset": offset,
        "returned": returned,
        "omitted": omitted,
        "truncated": omitted > 0,
        "reason": reason,
        "next_offset": (offset + returned) if omitted > 0 else None,
    }
    if shrunk:
        page["chars_cut_from_last"] = shrunk
    if omitted > 0:
        page["hint"] = _continue_hint(offset + returned, omitted, unit, narrow_hint)
    return kept, page


def _continue_hint(next_offset: int, omitted: int, unit: str, narrow_hint: str) -> str:
    base = (
        f"{omitted} more {unit} were not returned; call again with "
        f"offset={next_offset} for the next page"
    )
    if narrow_hint:
        return f"{base}, or {narrow_hint}"
    return base + "."


def columnar(items: list[dict]) -> dict:
    """Lossless compaction: a list of objects becomes a `fields` header and
    one array per item, which is what removes the repeated key names that
    dominate list-shaped payloads. Column order is first-seen order across
    the whole list, so a field that only some items carry still gets a
    column; a null in that column means the item did not carry the field,
    which is the same thing the object form said by omitting the key."""
    fields: list[str] = []
    seen: set[str] = set()
    for item in items:
        for key in item:
            if key not in seen:
                seen.add(key)
                fields.append(key)
    rows = [[item.get(f) for f in fields] for item in items]
    return {"fields": fields, "rows": rows}
