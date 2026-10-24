"""Customer-copy slots shared by the engine and the family surface.

Engine modules must not import :mod:`kitchensink4ppt.surface`.  Keeping the
small set of engine-owned messages here lets the product writer reach them
without reversing that dependency.  ``surface.copy`` re-exports every slot
and includes it in its ship-gate inventory.
"""

from __future__ import annotations


SLOT = "COPY" "_SLOT:"

#: Served both as the missing_title check's description and as its finding
#: message (ops/design_check.py). Facts from _check_missing_title: only a
#: non-empty title or ctrTitle placeholder counts.
MISSING_TITLE = (
    "The slide has no title: no title placeholder holds text. A text box styled like a "
    "title does not count, because screen readers and the slide list use only the title "
    "placeholder."
)
#: Finding message for ops/access.py _check_title_before_body.
TITLE_READ_AFTER_BODY = (
    "Screen readers reach other content on this slide before its title, because the title "
    "comes later in the slide's shape order."
)
BATCH_ATOMIC_ONLY = (
    "atomic must be true: every edit in the batch is saved together, or "
    "none is. Send edits that should succeed or fail on their own as separate "
    "calls."
)
#: Added to the answer of a read that read the saved file of a deck open in
#: PowerPoint (core/package.py; PowerPoint shares read access to an open
#: deck, and the file holds what was last saved).
OPEN_DECK_READ_NOTE = (
    "{name} is open in PowerPoint, so this read is of the file as last saved: changes made in "
    "PowerPoint since it was last saved are not included. get_text and get_slide_info read "
    "the open copy, unsaved changes included; other reads see those changes once the deck is "
    "saved in PowerPoint."
)
#: Follows OPEN_DECK_READ_NOTE when a read with a live route tried the open
#: copy first and could not reach it.
OPEN_DECK_LIVE_FAILED = " The open copy could not be read live: {reason}"
#: Added to an answer whose call sent a 0-based slides list (the user's
#: slide 44 sent as slides [44] renders the 45th slide, slide45.png).
SLIDES_NOTE = (
    "slides counts from 0: slides {sent} rendered slide_numbers {numbers}. To count from "
    "1, the way people do, send slide_numbers instead."
)
#: What set_text (and set_shape with text) does to a shape whose paragraphs
#: are formatted differently. Facts from ops/shapes.py set_shape and
#: _carry_text_properties; tests/unit/test_paragraph_edits.py holds them.
SET_TEXT_FORMATTING = (
    "set_text replaces every paragraph of the shape. The text splits into paragraphs at "
    "'\\n' ('\\v' is a line break inside one). New paragraph N takes old paragraph N's "
    "paragraph properties (alignment, bullets, level, spacing) and the character formatting "
    "of its first run; new paragraphs past the old count take the last old paragraph's, and "
    "old paragraphs past the new count are dropped with their formatting. A paragraph whose "
    "runs were formatted differently takes its first run's look, and the answer says "
    "runs_collapsed_to_first. The shape's frame (anchor, wrap, insets, autofit) is kept. To "
    "change one paragraph and keep the others as they are, use set_paragraph, "
    "insert_paragraph or delete_paragraph in apply_edits."
)
VIEW_TEXT_OMISSIONS = (
    "detail=\"text\" leaves out shapes with no text and all positions and "
    "sizes (detail=\"full\" includes them); headings number slides from 1, "
    "while slide arguments count from 0."
)
VIEW_OUTLINE_OMISSIONS = (
    "detail=\"outline\" shows slide titles only; use detail=\"text\" for text, "
    "tables and notes, or detail=\"full\" for every shape with its position "
    "and size."
)

# These templates cover customer-visible refusals raised below the surface
# package.  Format fields are deliberately named so a copy edit cannot lose
# the runtime facts the message must carry.
OP_NOT_FOUND_HINT = (
    "{op} is not an operation of this server; ppt_help lists every tool and operation."
)
OP_NOT_FOUND_MESSAGE = (
    "{op} is not an operation, so nothing ran; close matches, if any, are "
    "listed under did_you_mean."
)
WRONG_TOOL_MESSAGE = (
    "{op} belongs to the {owner} tool, not {tool}; call {owner} with op {op}."
)
BAD_ARGS_MESSAGE = (
    "The arguments for {op} do not match its schema; details.problems lists each one."
)
BAD_ARGS_HINT = (
    "Call ppt_help with op {op} for its arguments and an example."
)
TIMEOUT_MESSAGE = (
    "{what} ran past its {bound} limit and was stopped before anything was saved."
)
LIMIT_EXCEEDED_MESSAGE = (
    "The answer is larger than the server's output limit and cannot be shortened without "
    "dropping content."
)
LIMIT_EXCEEDED_HINT = (
    "Narrow the request, for example to fewer slides or one scope. {paging}"
)
#: Leads every answer to a call that named one slide by a 0-based index
#: (slide or slide_index), because a model that sent the user's own slide
#: number there got the next slide and cannot tell from a plausible answer.
SLIDE_NOTE = (
    "You sent {name}={index}, which counts from 0, so this {verb} the {ordinal} slide"
    "{contrast} (slide_number {number}{of_count}){titled}."
)
#: The deck size inside the note, left out when the call changed it.
SLIDE_NOTE_OF = " of {count}"
#: Names the slide a person would have meant by the same number.
SLIDE_NOTE_CONTRAST = ", not the {ordinal}"
SLIDE_NOTE_TITLED = ', titled "{title}"'
#: The page note on a finished write whose answer had to be cut to fit the
#: output limit. Whatever the call changed stands; only its answer is shorter.
COMMITTED_OUTPUT_NOTE = (
    "The call finished, and any change it made stands. Its answer passed the output "
    "limit, so long values were shortened or left out; read the deck to see the result."
)

#: case_note on find_text and search_and_replace: an exact-case call (match_case
#: true, the default) left out matches that differ only in case. One sentence:
#: how many, which slides (slide_number, counted from 1), and how to include them.
#: {spellings} is empty or ' ("Budget", "BUDGET")'. Built by ops/_casenote.py;
#: tests/unit/test_match_case.py holds the facts.
CASE_NOTE_FIND_NONE = (
    "No match in this exact case, but {count} {match_word} in another case{spellings} "
    "{is_are} on {slides}; send match_case=false to include {them}."
)
CASE_NOTE_FIND_MORE = (
    "{count} more {match_word} in another case{spellings} {is_are} on {slides} and "
    "{is_are} not listed; send match_case=false to include {them}."
)
CASE_NOTE_REPLACE = (
    "{count} {match_word} in another case{spellings} on {slides} {was_were} left "
    "unchanged; send match_case=false to replace {them} too."
)

#: A save or a restore that kept problems the deck already had (it refuses
#: only what it would add): core/package.py kept_problems_note. {first} is the
#: first problem as the check words it.
KEPT_PROBLEM_ONE = (
    "The deck already had a problem this change did not cause, and the save kept it "
    "as it was: {first}."
)
KEPT_PROBLEMS_MANY = (
    "The deck already had {count} problems this change did not cause, and the save "
    "kept them as they were; the first: {first}."
)

#: Decks refused when opened (core/errors.UnsupportedPackage carries the hint).
STRICT_DECK_REFUSAL = (
    "{name} is saved as Strict Open XML, which this server does not read; read as it "
    "is, it would look like a deck with no slides. Nothing was changed."
)
MAIN_PART_REFUSAL = (
    "{name} keeps its presentation in {part}, not in ppt/presentation.xml where "
    "PowerPoint saves it, so this server does not read it. Nothing was changed."
)

#: The hint of a refusal for an argument that needs a closed deck, sent for a
#: deck open in PowerPoint (server._live_refuse). The generic
#: UNSUPPORTED_CONTENT hint said the deck might be protected or damaged.
LIVE_FILE_ONLY_HINT = (
    "The deck is fine: this argument works on a closed deck only. Save and close the "
    "deck in PowerPoint and send the call again, or leave the argument out to change "
    "the open deck."
)
#: An out-of-range slide sent as slide_number (counted from 1): the refusal
#: names the number the caller sent (core/slide_echo.range_message).
SLIDE_NUMBER_RANGE = (
    "slide_number {number} out of range, presentation has {count} slide{plural}: "
    "slide_number counts from 1, so it runs 1 to {count} (slide 0 to {last})"
)
SLIDE_NUMBER_RANGE_ONE = (
    "slide_number {number} out of range, presentation has 1 slide: slide_number counts "
    "from 1, so the only slide is slide_number 1 (slide 0)"
)
#: PowerPoint opens nothing under a folder or file named like a device.
RESERVED_NAME_OPEN = (
    "PowerPoint could not open {name}: the path has a folder or file named "
    "{component!r}, a name Windows keeps for a device (con, prn, aux, nul, com1 to "
    "com9, lpt1 to lpt9), and PowerPoint opens nothing under such a name. The deck "
    "itself may be fine."
)
RESERVED_NAME_HINT = "Move the deck to a folder with another name, then send the call again."

ENGINE_COPY = (
    BATCH_ATOMIC_ONLY,
    OPEN_DECK_READ_NOTE,
    OPEN_DECK_LIVE_FAILED,
    SLIDES_NOTE,
    SET_TEXT_FORMATTING,
    MISSING_TITLE,
    TITLE_READ_AFTER_BODY,
    OP_NOT_FOUND_HINT,
    OP_NOT_FOUND_MESSAGE,
    WRONG_TOOL_MESSAGE,
    BAD_ARGS_MESSAGE,
    BAD_ARGS_HINT,
    TIMEOUT_MESSAGE,
    LIMIT_EXCEEDED_MESSAGE,
    LIMIT_EXCEEDED_HINT,
    SLIDE_NOTE,
    SLIDE_NOTE_CONTRAST,
    SLIDE_NOTE_OF,
    SLIDE_NOTE_TITLED,
    COMMITTED_OUTPUT_NOTE,
    VIEW_TEXT_OMISSIONS,
    VIEW_OUTLINE_OMISSIONS,
    CASE_NOTE_FIND_NONE,
    CASE_NOTE_FIND_MORE,
    CASE_NOTE_REPLACE,
    KEPT_PROBLEM_ONE,
    KEPT_PROBLEMS_MANY,
    STRICT_DECK_REFUSAL,
    MAIN_PART_REFUSAL,
    LIVE_FILE_ONLY_HINT,
    SLIDE_NUMBER_RANGE,
    SLIDE_NUMBER_RANGE_ONE,
    RESERVED_NAME_OPEN,
    RESERVED_NAME_HINT,
)


__all__ = [
    "BATCH_ATOMIC_ONLY",
    "CASE_NOTE_FIND_MORE",
    "CASE_NOTE_FIND_NONE",
    "CASE_NOTE_REPLACE",
    "OPEN_DECK_LIVE_FAILED",
    "OPEN_DECK_READ_NOTE",
    "SLIDES_NOTE",
    "SET_TEXT_FORMATTING",
    "BAD_ARGS_HINT",
    "BAD_ARGS_MESSAGE",
    "COMMITTED_OUTPUT_NOTE",
    "ENGINE_COPY",
    "KEPT_PROBLEM_ONE",
    "KEPT_PROBLEMS_MANY",
    "LIMIT_EXCEEDED_HINT",
    "LIMIT_EXCEEDED_MESSAGE",
    "LIVE_FILE_ONLY_HINT",
    "MAIN_PART_REFUSAL",
    "MISSING_TITLE",
    "OP_NOT_FOUND_HINT",
    "OP_NOT_FOUND_MESSAGE",
    "RESERVED_NAME_HINT",
    "RESERVED_NAME_OPEN",
    "SLIDE_NOTE",
    "SLIDE_NOTE_CONTRAST",
    "SLIDE_NOTE_OF",
    "SLIDE_NOTE_TITLED",
    "SLIDE_NUMBER_RANGE",
    "SLIDE_NUMBER_RANGE_ONE",
    "SLOT",
    "STRICT_DECK_REFUSAL",
    "TIMEOUT_MESSAGE",
    "TITLE_READ_AFTER_BODY",
    "VIEW_OUTLINE_OMISSIONS",
    "VIEW_TEXT_OMISSIONS",
    "WRONG_TOOL_MESSAGE",
]
