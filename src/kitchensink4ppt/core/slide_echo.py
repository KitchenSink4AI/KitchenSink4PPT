"""Name the slide a call addressed, counted the way people count.

Slide arguments count from 0 and people count from 1. A model that sends
slide=2 for "slide 2" gets the third slide, and a read that succeeds on the
wrong slide looks exactly like a right answer: "slide 2 has no notes" when
the notes are on the slide before. So every response to a call that names
one slide leads with that slide's 1-based ``slide_number``, the deck's
``slide_count`` and the slide's ``slide_title`` (when it has one), and the
caller can see which slide it got before it reports anything.

The slide resolvers report what they resolved through a context variable
that dispatch arms for one call, so nothing is read twice and the live
PowerPoint route reports through the same hook. Only a resolution of the
call's own ``slide`` selector counts; a second selector the operation
resolves on the way (a jump-link target, say) cannot take its place. The
number and title describe the slide as the call found it, before any
change the call made; slide_count is the deck after the call. A call that
removed or copied the slide names it accordingly (STYLES), so a deleted
slide is never read as one still in the deck.
"""

from __future__ import annotations

import contextlib
import contextvars
import re
from typing import Any, Callable

from .messages import SLIDE_NOTE, SLIDE_NOTE_CONTRAST, SLIDE_NOTE_OF, SLIDE_NOTE_TITLED

#: Longest title echoed; a title is an identifier here, not content.
TITLE_CHARS = 80

_ACTIVE: contextvars.ContextVar["Probe | None"] = contextvars.ContextVar(
    "ks4p_slide_probe", default=None
)
_SPACE = re.compile(r"\s+")


def single(selector: Any) -> bool:
    """True for a selector naming exactly one slide: an index or a slide id."""
    if isinstance(selector, bool):
        return False
    if isinstance(selector, int):
        return True
    return (
        isinstance(selector, dict)
        and set(selector) == {"slide_id"}
        and isinstance(selector["slide_id"], int)
        and not isinstance(selector["slide_id"], bool)
    )


def _same(a: Any, b: Any) -> bool:
    if isinstance(a, bool) or isinstance(b, bool):
        return False
    if isinstance(a, int) and isinstance(b, int):
        return a == b
    if isinstance(a, dict) and isinstance(b, dict):
        return a == b
    return False


def clean_title(text: Any) -> str | None:
    """One line, trimmed, bounded; None when there is no title text."""
    if not isinstance(text, str):
        return None
    line = _SPACE.sub(" ", text).strip()
    if not line:
        return None
    if len(line) > TITLE_CHARS:
        line = line[: TITLE_CHARS - 3].rstrip() + "..."
    return line


#: The argument names that count slides from 0. A call that used one of
#: them gets a plain-words note, because the same number means the next
#: slide to a person.
ZERO_BASED = ("slide", "slide_index", "scope", "slides")

#: Operations whose ``scope`` selects slides, armed like ``slide`` when the
#: scope names exactly one slide: get_text(scope=2) leads with the slide it
#: read, the way get_notes(slide=2) does. set_footer is left out on purpose,
#: since its own slide_number argument means "show the slide number".
SCOPE_OPS = frozenset({
    "get_presentation_view", "get_text", "find_text", "list_elements",
    "list_hyperlinks", "list_equations", "list_comments", "audit_accessibility",
    "search_and_replace",
})

#: Operations whose answer names the slide differently, so that slide_count
#: is always the deck as it is after the call. delete_slide's slide is gone
#: (deleted_slide_number, one slide fewer); duplicate_slide's is the source
#: of a copy (source_slide_number, one slide more); copy_slide_between's is
#: in another deck (source_slide_number and source_slide_count describe
#: that deck). A refusal changed nothing and uses the plain names.
STYLES = {
    "delete_slide": "deleted",
    "duplicate_slide": "copied",
    "copy_slide_between": "source_deck",
}

_ORDINALS = (
    "first", "second", "third", "fourth", "fifth", "sixth", "seventh",
    "eighth", "ninth", "tenth", "eleventh", "twelfth",
)


def ordinal(number: int) -> str:
    """1 -> 'first', 13 -> '13th', 22 -> '22nd'."""
    if 1 <= number <= len(_ORDINALS):
        return _ORDINALS[number - 1]
    if 10 <= number % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(number % 10, "th")
    return f"{number}{suffix}"


class Probe:
    """What one call's ``slide`` selector resolved to, once."""

    __slots__ = ("selector", "sent_as", "style", "hit")

    def __init__(
        self, selector: Any, sent_as: str | None = None, style: str | None = None
    ) -> None:
        self.selector = selector
        self.sent_as = sent_as
        self.style = style
        self.hit: tuple[int, int, str | None] | None = None

    def offer(
        self,
        selector: Any,
        index: int,
        count: int,
        title: Callable[[], Any] | str | None,
    ) -> None:
        if self.hit is not None or not _same(selector, self.selector):
            return
        try:
            text = title() if callable(title) else title
        except Exception:  # the echo must never be why a call fails
            text = None
        self.hit = (int(index), int(count), clean_title(text))

    def fields(self, done: bool = True) -> dict[str, Any]:
        """The echo; ``done`` is False for a refusal, which changed nothing."""
        if self.hit is None:
            return {}
        index, count, title = self.hit
        style = self.style if done else None
        out: dict[str, Any] = {}
        if self.sent_as in ZERO_BASED and isinstance(self.selector, int):
            out["slide_note"] = SLIDE_NOTE.format(
                name=self.sent_as,
                index=self.selector,
                verb="was" if style == "deleted" else "is",
                ordinal=ordinal(index + 1),
                # The slide a person means by the same number, which is
                # the usual mistake; slide 0 has no such reading.
                contrast=(
                    SLIDE_NOTE_CONTRAST.format(ordinal=ordinal(self.selector))
                    if self.selector >= 1 else ""
                ),
                number=index + 1,
                of_count="" if style else SLIDE_NOTE_OF.format(count=count),
                titled=SLIDE_NOTE_TITLED.format(title=title) if title else "",
            )
        if style == "deleted":
            out["deleted_slide_number"] = index + 1
            if title:
                out["deleted_slide_title"] = title
            out["slide_count"] = count - 1
        elif style in ("copied", "source_deck"):
            out["source_slide_number"] = index + 1
            if title:
                out["source_slide_title"] = title
            if style == "copied":
                out["slide_count"] = count + 1
            else:
                out["source_slide_count"] = count
        else:
            out["slide_number"] = index + 1
            out["slide_count"] = count
            if title:
                out["slide_title"] = title
        return out


def arm(
    selector: Any, sent_as: str | None = None, op: str | None = None
) -> tuple[Probe | None, contextvars.Token | None]:
    """Start listening for ``selector`` in the current context.

    ``sent_as`` is the argument name the caller used for it, which decides
    whether the answer also carries the 0-based note; ``op`` picks the
    field names (STYLES).
    """
    if not single(selector):
        return None, None
    probe = Probe(selector, sent_as, STYLES.get(op or ""))
    return probe, _ACTIVE.set(probe)


def disarm(token: contextvars.Token | None) -> None:
    if token is not None:
        _ACTIVE.reset(token)


def report(
    selector: Any,
    index: int,
    count: int,
    title: Callable[[], Any] | str | None = None,
) -> None:
    """Called by a slide resolver after it resolved ``selector``."""
    probe = _ACTIVE.get()
    if probe is not None:
        probe.offer(selector, index, count, title)


def restart() -> None:
    """Forget a file-mode resolution when the call moves to PowerPoint."""
    probe = _ACTIVE.get()
    if probe is not None:
        probe.hit = None


#: True while a slide the caller named by slide_number (counted from 1) is
#: resolved inside a batch edit, where no call-level probe knows the name.
_NUMBERED: contextvars.ContextVar[bool] = contextvars.ContextVar(
    "ks4p_slide_numbered", default=False)


@contextlib.contextmanager
def numbered(flag: bool):
    """Resolve with out-of-range refusals worded in slide_number when flag."""
    token = _NUMBERED.set(bool(flag))
    try:
        yield
    finally:
        _NUMBERED.reset(token)


def _sent_as_number(index) -> bool:
    if _NUMBERED.get():
        return True
    probe = _ACTIVE.get()
    return (
        probe is not None and probe.sent_as == "slide_number"
        and _same(index, probe.selector)
    )


def range_message(index: int, count: int) -> str:
    """The out-of-range refusal every slide resolver gives, numbering included.

    A caller that sent the user's 1-based number for the last slide lands one
    past the end; saying both ranges lets it correct itself in one step. A
    slide sent as slide_number is named by that number, the one the caller
    sent, not by the 0-based index it was turned into.
    """
    plural = "" if count == 1 else "s"
    if (
        count >= 1 and isinstance(index, int) and not isinstance(index, bool)
        and _sent_as_number(index)
    ):
        from .messages import SLIDE_NUMBER_RANGE, SLIDE_NUMBER_RANGE_ONE

        if count == 1:
            return SLIDE_NUMBER_RANGE_ONE.format(number=index + 1)
        return SLIDE_NUMBER_RANGE.format(
            number=index + 1, count=count, plural=plural, last=count - 1)
    head = f"slide index {index} out of range, presentation has {count} slide{plural}"
    if count <= 0:
        return head
    if count == 1:
        return f"{head}: slide counts from 0, so the only slide is 0 (slide_number 1)"
    return (
        f"{head}: slide counts from 0, so it runs 0 to {count - 1} "
        f"(slide_number 1 to {count})"
    )


def position_message(position: Any, count: int, deck: str = "the deck") -> str:
    """The refusal for a new slide's position outside 0..count.

    ``count`` is the deck as it is now, before the slide is added, so the
    message never describes a slide that was not made.
    """
    plural = "" if count == 1 else "s"
    return (
        f"position {position} out of range; {deck} has {count} slide{plural}, "
        f"so position counts from 0 and runs 0 to {count} ({count} adds it last)"
    )


def lead(result: dict[str, Any], probe: Probe | None) -> dict[str, Any]:
    """Put the echo first, without replacing a key the operation set."""
    if probe is None:
        return result
    fields = {
        key: value for key, value in probe.fields().items() if key not in result
    }
    if not fields:
        return result
    return {**fields, **result}


__all__ = [
    "Probe", "SCOPE_OPS", "STYLES", "TITLE_CHARS", "ZERO_BASED", "arm", "clean_title",
    "disarm", "lead", "ordinal", "position_message", "range_message", "report",
    "restart", "single",
]
