"""The case_note of find_text and search_and_replace.

An exact-case call (match_case true, the default) answers only for text in
exactly the case it was sent. When that leaves out matches that differ only
in case, the answer says so in one sentence (core/messages CASE_NOTE_*): how
many, on which slides counted from 1, and how to include them
(match_case=false). The tally runs over the same paragraphs, scope and notes
setting as the call itself, so the note never points outside what was asked.
"""

from __future__ import annotations

import re as _stdlib_re

from ..core import messages as _msg
from . import _regex

#: How many spellings the note quotes, and how many slides it lists before
#: "and N more".
_MAX_SPELLINGS = 3
_MAX_SLIDES = 8
_MAX_SPELLING_CHARS = 40


def any_case_spans(text: str, query: str, regex: bool) -> list[tuple[int, int]]:
    """Non-overlapping, non-empty matches of `query` in `text`, in any case.
    Case is folded on the text itself (IGNORECASE), never on a lower-cased
    copy, whose length can differ from the text's."""
    if regex:
        found = _regex.finditer(query, text, ignore_case=True)
    else:
        found = _stdlib_re.finditer(
            _stdlib_re.escape(query), text, _stdlib_re.IGNORECASE
        )
    return [m.span() for m in found if m.end() > m.start()]


def _shown(word: str) -> str:
    word = word.replace("\r", " ").replace("\n", " ").replace("\v", " ")
    if len(word) > _MAX_SPELLING_CHARS:
        word = word[:_MAX_SPELLING_CHARS] + "..."
    return f'"{word}"'


def _slides(indexes) -> str:
    numbers = [i + 1 for i in sorted(indexes)]
    if len(numbers) == 1:
        return f"slide_number {numbers[0]}"
    shown = [str(n) for n in numbers[:_MAX_SLIDES]]
    rest = len(numbers) - len(shown)
    if rest:
        return f"slide_numbers {', '.join(shown)} and {rest} more"
    return f"slide_numbers {', '.join(shown[:-1])} and {shown[-1]}"


class CaseTally:
    """Matches in another case, per slide, gathered while a call scans."""

    def __init__(self, query: str, regex: bool):
        self.query = query
        self.regex = regex
        self.extra: dict[int, int] = {}
        self.spellings: list[str] = []

    def add(self, slide_index: int, text: str, exact_spans, skip=()) -> None:
        """Count, for one paragraph's text, the matches an any-case search
        finds beyond the exact-case ones (exact_spans: (start, end, ...)).
        A match that touches a `skip` range (start, end) is not counted:
        search_and_replace passes the equations and other wrapped content a
        replace never rewrites, so the note never offers to replace them."""
        if not text:
            return
        others = any_case_spans(text, self.query, self.regex)
        if skip:
            def outside(span) -> bool:
                return not any(lo < span[1] and hi > span[0] for lo, hi in skip)

            others = [span for span in others if outside(span)]
            exact_spans = [span for span in exact_spans if outside(span)]
        more = len(others) - len(exact_spans)
        if more <= 0:
            return
        self.extra[slide_index] = self.extra.get(slide_index, 0) + more
        exact = {text[span[0]:span[1]] for span in exact_spans}
        if not self.regex:
            exact.add(self.query)
        for start, end in others:
            word = text[start:end]
            if word not in exact and word not in self.spellings:
                self.spellings.append(word)

    @property
    def count(self) -> int:
        return sum(self.extra.values())

    def note(self, kind: str, exact_count: int) -> str | None:
        """The sentence for kind 'find' or 'replace', or None when no match
        was left out."""
        count = self.count
        if not count:
            return None
        one = count == 1
        spellings = ""
        if self.spellings:
            listed = ", ".join(_shown(w) for w in self.spellings[:_MAX_SPELLINGS])
            if len(self.spellings) > _MAX_SPELLINGS:
                listed += " and others"
            spellings = f" ({listed})"
        words = {
            "count": count,
            "match_word": "match" if one else "matches",
            "is_are": "is" if one else "are",
            "was_were": "was" if one else "were",
            "them": "it" if one else "them",
            "spellings": spellings,
            "slides": _slides(self.extra),
        }
        if kind == "replace":
            return _msg.CASE_NOTE_REPLACE.format(**words)
        if exact_count:
            return _msg.CASE_NOTE_FIND_MORE.format(**words)
        return _msg.CASE_NOTE_FIND_NONE.format(**words)
