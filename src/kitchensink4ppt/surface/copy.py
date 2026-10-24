"""Customer copy for the PowerPoint family surface.

Written from the code and the op docstrings, in the register of the Word
and Excel editors' copy. TEMPLATE_NOTICE is legal text: one factual
sentence followed by the python-pptx 1.0.2 copyright line and permission
notice copied verbatim from that release's LICENSE file.
"""

from __future__ import annotations

from ..core.messages import (
    BATCH_ATOMIC_ONLY,
    BAD_ARGS_HINT,
    BAD_ARGS_MESSAGE,
    ENGINE_COPY,
    LIMIT_EXCEEDED_HINT,
    LIMIT_EXCEEDED_MESSAGE,
    MISSING_TITLE,
    OP_NOT_FOUND_HINT,
    OP_NOT_FOUND_MESSAGE,
    SET_TEXT_FORMATTING,
    SLOT,
    TIMEOUT_MESSAGE,
    TITLE_READ_AFTER_BODY,
    VIEW_OUTLINE_OMISSIONS,
    VIEW_TEXT_OMISSIONS,
    WRONG_TOOL_MESSAGE,
)

#: Appended after the product description server.py already carries (file
#: based, two-slot backups, live='auto'). The whole assembled text stays
#: under Claude Code's 2,048-character cap on server instructions; the word
#: "update" may not appear (test_update_check), nor the retired toggles.
INSTRUCTIONS = (
    "Each tool takes {op, args}: op names an operation listed in the tool's description, "
    "and args holds its arguments, including file_path (the .pptx) unless the operation "
    "names other paths. ppt_help with an op returns its arguments and an example. ppt_read "
    "reads content and ppt_inspect runs audits; neither changes a file. Edits to a deck open "
    "in PowerPoint stay unsaved for the user to review and save; call ppt_deck live_save only "
    "when the user asks. Launch settings may make this server read-only or remove tools and "
    "operations; no tool call changes that."
)

#: ppt_deck head addition (edits to an open deck stay unsaved for
#: the user; live_save only when asked). Also served through ppt_edit's head.
OPEN_DECK_SENTENCE = (
    "Edits to a deck open in PowerPoint stay unsaved for the user to review and save; call "
    "live_save only when the user asks.")

TOOLS: dict[str, dict[str, str]] = {
    "ppt_read": {
        "title": "Read PowerPoint Deck",
        "description": (
            "Read slide content: the anchored deck view, deck and slide facts, text, find, "
            "elements, links, tables, chart data, notes, comments, equations, media, "
            "transitions and animations. Never changes a file. Audits and checks are "
            "ppt_inspect; to change anything use ppt_edit, ppt_graphics, ppt_tables, "
            "ppt_design, ppt_review or ppt_deck."),
    },
    "ppt_inspect": {
        "title": "Inspect and Audit Deck",
        "description": (
            "Audits and deck-level facts: accessibility, layout checks, font inventory, "
            "comment report, deck comparison, statistics, theme, brand, autofit, properties, "
            "footer support, masters, export engines, and PowerPoint status. Never changes a "
            "file. Slide content reads are ppt_read."),
    },
    "ppt_edit": {
        "title": "Edit PowerPoint Slides and Text",
        "description": (
            "Slide text and slides: the batch editor (apply_edits), find and replace, "
            "placeholders, fit, text formatting, bullets, text boxes, notes, hyperlinks, and "
            "inserting, deleting, duplicating, ordering, hiding and sectioning slides. Shapes "
            "and pictures are ppt_graphics; tables and charts are ppt_tables. Every edit backs "
            "the file up first and saves atomically. " + OPEN_DECK_SENTENCE),
    },
    "ppt_graphics": {
        "title": "PowerPoint Shapes and Media",
        "description": (
            "Shapes and media: insert, edit, group, align, distribute and restack shapes; "
            "connectors; SVG to native shapes; diagrams; format and position painters; "
            "pictures, video and audio; equations; alt text and reading order. Slide text is "
            "ppt_edit. Every edit backs the file up first and saves atomically."),
    },
    "ppt_tables": {
        "title": "PowerPoint Tables and Charts",
        "description": (
            "Tables and charts: create, fill, merge, resize, style and import tables, and "
            "create, refill and format charts. Table and chart reads are ppt_read. Row and "
            "column addresses are 0-based. Every edit backs the file up first and saves "
            "atomically."),
    },
    "ppt_design": {
        "title": "PowerPoint Design and Masters",
        "description": (
            "Deck design: slide size, layouts, theme colors and fonts, brand transfer, "
            "backgrounds, masters and layout placeholders, deck-wide font and color "
            "replacement, transitions and entrance animations. Theme and master reads are "
            "ppt_inspect. Every edit backs the file up first and saves atomically."),
    },
    "ppt_review": {
        "title": "PowerPoint Comments and Review",
        "description": (
            "Review: add, reply to, resolve and delete modern threaded comments, and anonymize "
            "authors. Reading comments is ppt_read list_comments or ppt_inspect "
            "comment_report. anonymize_deck saves a snapshot automatically before changing "
            "anything, because the change cannot be undone in the file."),
    },
    "ppt_deck": {
        "title": "PowerPoint Decks and Exports",
        "description": (
            "Whole decks and PowerPoint itself: create, copy, snapshot, backups, copy slides "
            "between decks, merge, split, export to PDF, images, handouts or tables, document "
            "properties, agenda slides, compression, show settings, custom shows, footers, the "
            "opens-clean check (validate), and saving or scrolling a deck open in PowerPoint. "
            + OPEN_DECK_SENTENCE),
    },
    "ppt_help": {
        "title": "PowerPoint Tool Help",
        "description": (
            "Read-only help. With no arguments it lists every tool and its operations; with op "
            "it returns that operation's full arguments, an example call and its tool. Call it "
            "when arguments are unclear or after a BAD_ARGS refusal. Reads no deck."),
    },
}

OP_PARAM = "The operation to run: one of the operations listed in this tool's description."
#: Slide arguments count from 0 and people count from 1; a model that mixes
#: the two reads or edits the wrong slide.
#: The rule rides in every tool's args description, where an argument is
#: chosen, and slide_number lets a caller count the way the user does.
SLIDE_RULE = (
    "Slides count from 1 in slide_number and from 0 in slide: the user's slide 2 is "
    "slide_number 2 or slide 1.")
ARGS_PARAM = (
    "The operation's arguments as an object, including file_path unless the operation names "
    "other paths or reads no deck; unknown keys are refused. " + SLIDE_RULE + " ppt_help "
    "lists every argument.")
#: ppt_help's description of ``slide`` on every operation that takes one.
SLIDE_PARAM = (
    "The slide, counted from 0: the first slide is 0, so the slide the user calls slide 2 is "
    'slide 1. {"slide_id": N} also works. Or send slide_number instead.')
#: ppt_help lists slide_number after slide wherever slide is accepted. Null is
#: allowed because the server reads a null slide_number as not sent, as it
#: does every optional argument: clients that fill unused fields with null are
#: common, and a schema check must not block the call the server accepts.
SLIDE_NUMBER_PARAM = {
    "type": ["integer", "null"],
    "minimum": 1,
    "title": "Slide Number",
    "description": (
        "The same slide counted from 1, the way people count: slide_number 2 is the second "
        "slide. Send it instead of slide."),
}
#: ppt_help lists slide_numbers after slides wherever a list of slides is
#: accepted; null reads as not sent, as for slide_number.
SLIDE_NUMBERS_PARAM = {
    "type": ["array", "null"],
    "items": {"type": "integer", "minimum": 1},
    "title": "Slide Numbers",
    "description": (
        "The same slides counted from 1, the way people count: [1, 2, 4] is the first, "
        "second and fourth slides. Send it instead of slides, not with it."),
}
#: ppt_help's description of a shape addressed by id or by name.
SHAPE_PARAM = (
    "The shape's id, its name, or its anchor from get_presentation_view (such as "
    "a:3f9c2a). Shapes inside a group are found too; when a name repeats, send "
    "\"Group/Shape\" (the group's name, a slash, the shape's name). shape_id is read as "
    "another name for shape.")
#: ppt_help's description of a scope that selects slides.
SCOPE_PARAM = (
    "Which slides, counted from 0 like slide: leave it out for every slide, or send one "
    'slide (the user\'s slide 2 is 1, or {"slide_id": N}), a range such as {"start": 1, '
    '"end": 3} (both included), or a list of them.')
#: The scope of the read operations, which also takes slide_number and
#: slide_numbers (set_footer's scope does not: its own slide_number argument
#: means "show the slide number").
SCOPE_TWIN_PARAM = SCOPE_PARAM + (
    " Or send slide_number or slide_numbers, counted from 1. slides is read as another "
    "name for scope.")
#: ppt_help's descriptions of the 1-based twins, by (argument, twin).
TWIN_HELP: dict[tuple[str, str], str] = {
    ("slides", "slide_numbers"): SLIDE_NUMBERS_PARAM["description"],
    ("scope", "slide_number"): (
        "One slide counted from 1, the way people count: slide_number 2 is the second "
        "slide. Send it instead of scope."),
    ("scope", "slide_numbers"): (
        "Slides counted from 1, the way people count: [1, 3] is the first and third "
        'slides, and a range such as {"start": 2, "end": 4} is the second to fourth, both '
        "included. Send it instead of scope, not with it."),
    ("from_slide", "from_slide_number"): (
        "The same slide counted from 1, the way people count: from_slide_number 2 is the "
        "second slide. Send it instead of from_slide."),
    ("to_slides", "to_slide_numbers"): (
        "The same slides counted from 1, the way people count: [2, 3] is the second and "
        "third slides. Send it instead of to_slides, not with it."),
    ("order", "order_numbers"): (
        "The new order with every slide counted from 1, the way people count: [3, 1, 2] "
        "puts the third slide first. Send it instead of order, not with it."),
    ("ranges", "number_ranges"): (
        'For by="ranges": the same ranges with start and end counted from 1, both '
        'included: {"start": 1, "end": 3} is the first three slides. Send it instead of '
        "ranges, not with it."),
    ("to_slide", "to_slide_number"): (
        "The slide the link jumps to, counted from 1, the way people count: "
        "to_slide_number 3 is the third slide. Send it instead of to_slide."),
    ("to", "to_number"): (
        "Where the slide goes, counted from 1, the way people count: to_number 1 makes it "
        "the first slide. Send it instead of to."),
    ("position", "position_number"): (
        "Where the new slide goes, counted from 1, the way people count: position_number 3 "
        "makes it the third slide. Send it instead of position."),
    ("source_slide", "source_slide_number"): (
        "The source slide counted from 1, the way people count: source_slide_number 2 is "
        "the second slide of the source deck. Send it instead of source_slide."),
}
#: ppt_help's descriptions of the other arguments that name slides or slide
#: positions, by operation. Most count from 0, like slide; the show range
#: counts from 1, and says so.
ARG_NOTES: dict[str, dict[str, str]] = {
    **{op: {"scope": SCOPE_TWIN_PARAM} for op in (
        "get_presentation_view", "get_text", "find_text",
        "list_equations", "list_comments", "audit_accessibility", "search_and_replace")},
    "list_hyperlinks": {"scope": SCOPE_TWIN_PARAM + (
        " Links in those slides' notes and on the layouts and masters they use are listed "
        "too. set_hyperlink and remove_hyperlink reach the links on shapes and their text; "
        "the others are listed for reading. A click that only plays a sound is listed as kind "
        "sound; the play controls of a video or audio shape are not links and are not listed.")},
    "set_footer": {"scope": SCOPE_PARAM},
    "list_elements": {"scope": SCOPE_TWIN_PARAM + " Only the kinds that sit on slides use it."},
    "set_hyperlink": {"to_slide": (
        "The slide the link jumps to, counted from 0 like slide (the user's slide 3 is 2), "
        'or {"slide_id": N}. The answer names it as target_slide_number and '
        "target_slide_title. Or send to_slide_number, counted from 1.")},
    "move_slide": {"to": (
        "Where the slide goes, counted from 0 like slide: 0 makes it the first slide, so "
        "the user's slide 2 is to 1. Or send to_number, counted from 1.")},
    "insert_slide": {"position": (
        "Where the new slide goes, counted from 0 like slide: 0 makes it the first slide, "
        "so the user's slide 3 is position 2. Leave it out to add it last, or send "
        "position_number, counted from 1.")},
    "duplicate_slide": {"position": (
        "Where the copy goes, counted from 0 like slide: 0 makes it the first slide, so the "
        "user's slide 3 is position 2. Leave it out to put it right after the original, or "
        "send position_number, counted from 1.")},
    "copy_shape_between": {"source_slide": (
        "The slide that holds the shape in the source deck, counted from 0 like slide (the "
        "user's slide 2 is 1), or {\"slide_id\": N}. Or send source_slide_number, counted "
        "from 1.")},
    "copy_slide_between": {"position": (
        "Where the copied slide goes in this deck, counted from 0 like slide: 0 makes it the "
        "first slide. Leave it out to add it last, or send position_number, counted from 1.")},
    "generate_agenda_slide": {"position": (
        "Where the agenda goes, counted from 0 like slide: the default 1 makes it the "
        "second slide, after the title slide. Or send position_number, counted from 1.")},
    "reorder_slides": {"order": (
        "Every slide once, each counted from 0 like slide, in the new order: [2, 0, 1] "
        "puts the third slide first. Or send order_numbers, counted from 1.")},
    "export_slide_image": {"slides": (
        "The slides to render, counted from 0 like slide (the user's slide 2 is 1). Leave it "
        "out for every slide, or send slide_numbers instead.")},
    "manage_custom_show": {"slides": (
        "For create: the show's slides in order, each counted from 0 like slide (the user's "
        'slide 2 is 1) or {"slide_id": N}. A slide can appear more than once. Or send '
        "slide_numbers instead.")},
    "copy_position": {
        "from_slide": (
            "The slide that holds the source shape, counted from 0 like slide (the user's "
            'slide 2 is 1), or {"slide_id": N}. Or send from_slide_number, counted from 1.'),
        "to_slides": (
            'The slides to update, each counted from 0 like slide or {"slide_id": N}. Leave '
            "it out for every other slide, or send to_slide_numbers, counted from 1."),
    },
    "split_deck": {"ranges": (
        'For by="ranges": a list of {"start", "end", "name"}, with start and end counted '
        'from 0 like slide and both included: {"start": 0, "end": 2} is the first three '
        "slides. Or send number_ranges, counted from 1.")},
    "set_show_properties": {"range": (
        '"all", {"start", "end"} counted from 1 the way people count, both included '
        '({"start": 2, "end": 4} shows slides 2 to 4), or {"custom_show": name or id}.')},
    "manage_section": {"section": "A section's name, or its position counted from 0."},
}
#: Who a comment or reply is signed by when the call names no author.
COMMENT_AUTHOR_PARAM = (
    "The name to sign with. Leave author out to sign with the user's own name: their Office "
    "user name, or their Windows sign-in name when Office has none. Pass author=\"assistant\" "
    "to sign as \"AI Assistant via KS4PPT\", which a server started with "
    "KS4P_AUTHOR=assistant does by default; any other name is used as given. An email "
    "address is never written as a name: one sent here is set aside, the default signs "
    "instead, and the answer says so. The answer's author_source says where the name "
    "came from.")
#: ppt_help's descriptions of other arguments a model has to get right the
#: first time, by operation.
ARG_HELP: dict[str, dict[str, str]] = {
    "set_notes": {"notes": (
        "Many slides in one call, with one backup and one save: a list of {slide_id or "
        "slide_number, text or paragraphs}, such as [{\"slide_id\": 257, \"text\": \"...\"}, "
        "{\"slide_number\": 4, \"text\": \"...\"}]. Each entry names its own slide (a bare "
        "0-based slide is refused), every slide is checked before anything is written, and "
        "an entry may carry its own flatten. Send it instead of slide; file mode only.")},
    "add_comment": {"author": COMMENT_AUTHOR_PARAM},
    "reply_to_comment": {"author": COMMENT_AUTHOR_PARAM},
    "set_shape": {
        "shape": SHAPE_PARAM,
        "text": "Replaces all of the shape's text. " + SET_TEXT_FORMATTING,
        "scale_text": (
            "With w or h: scale every font size, text inset and point spacing inside the "
            "shape or group by the same ratio (the smaller of the width and height ratios), "
            "so labels shrink with a figure instead of wrapping. Off by default."),
        "scale_lines": (
            "With w or h: scale every outline width inside by the same ratio, theme lines "
            "included. Off by default."),
        "no_wrap": (
            "Keep each paragraph of the text boxes inside on one line (no wrapping), so "
            "small labels do not break. Other shapes keep their wrapping. Off by default."),
    },
    "format_text": {
        "shape": SHAPE_PARAM,
        "size_pt": "The font size in points. font_size is read as another name for it."},
    "insert_textbox": {
        "size_pt": "The font size in points. font_size is read as another name for it."},
    **{op: {"size": (
        "The font size in points. font_size and size_pt are read as other names for it.")}
       for op in ("set_master_placeholder", "set_layout_placeholder")},
    "find_text": {"match_case": (
        "True by default: only text in exactly this case matches, so \"budget\" does not "
        "find \"Budget\". Send match_case=false to match any case. When matches in another "
        "case exist, the answer's case_note says how many and on which slides.")},
    "search_and_replace": {
        "match_case": (
            "True by default: only text in exactly this case is replaced, so \"budget\" "
            "leaves \"Budget\" as it is. Send match_case=false to replace any case. When "
            "matches in another case are left, the answer's case_note says how many and on "
            "which slides."),
        "remove_empty_paragraphs": (
            "Remove each paragraph a replacement leaves with no text (or only spaces and "
            "line breaks), in slide text and table cells, and in notes when include_notes "
            "is true, so no blank line is left behind. SmartArt keeps its paragraphs. "
            "Paragraphs that were empty before are kept, and a text body keeps at least one "
            "paragraph. Off by default; file mode only."),
    },
    "export_shape_image": {
        "shape": SHAPE_PARAM,
        "output": (
            "The PNG file to write; leave it out for one beside the deck named after the "
            "slide and the shape. An existing file is replaced only with overwrite=true."),
        "width": (
            "The image width in pixels (default 1280); the height follows the shape's "
            "proportions. Send height instead to set that side."),
        "height": "The image height in pixels, instead of width.",
    },
    "copy_shape_between": {
        "file_path": "The deck that receives the copy.",
        "source_path": (
            "The deck the shape comes from. It is only read, never changed, and it may be "
            "file_path itself to copy within one deck."),
        "shape": (
            "The shape or group on the source slide: its id, or its name (\"Group/Shape\" "
            "for a member of a group)."),
        "x": "Where the copy's left edge goes, in inches; leave out x and y to keep its place.",
        "y": "Where the copy's top edge goes, in inches; leave out x and y to keep its place.",
        "w": "The copy's width in inches; send w or h alone to keep the proportions.",
        "h": "The copy's height in inches; send w or h alone to keep the proportions.",
        "leave_out": (
            "Members of a copied group to drop, by id or name (\"Group/Shape\" when a name "
            "repeats). The group is refitted to what is left, so nothing else moves."),
        "name": "A name for the copy; leave it out to keep the source shape's name.",
    },
    "insert_shape": {"group": (
        "An existing group's id or name (\"Outer/Inner\" for a group inside a group) to put "
        "the new shape in. x and y stay slide positions, and the group's frame grows to "
        "hold the shape. Leave it out to place the shape on the slide.")},
    "create_snapshot": {"dest_dir": (
        "Leave it out to keep the snapshot in the deck's hidden backup folder, where "
        "manage_backups lists it and restores from it. Send a folder for a visible copy there "
        "instead, named after the deck: \"<deck name> (snapshot <label> YYYY-MM-DD "
        "HHMM).pptx\".")},
    **{op: {"shape": SHAPE_PARAM} for op in (
        "fit_text", "set_bullets", "get_autofit_state", "add_equation_to_shape")},
    "manage_backups": {
        "source": (
            "For action='restore': the backup to put back, as action='list' gives it in "
            "each backup's source: 'prev' (the deck before its last change), 'anchor' (the "
            "deck when the current editing session began) or a snapshot's .pptx path. slot "
            "is read as another name for source."),
        "scope": (
            "For action='purge': 'slots' (this deck's prev and anchor) or 'orphans' "
            "(backup folders whose deck is gone; a folder that holds a snapshot is kept)."),
    },
    "set_shape_adjustments": {
        "shape": SHAPE_PARAM,
        "adjustments": (
            'Handle name to value as stored, such as {"adj": 25000}; a value with a decimal '
            "point is a fraction (0.25 is 25000). Handles you leave out keep their values. "
            "Leave out adjustments and degrees to read the handles first."),
        "degrees": (
            "Angle handle name to degrees, clockwise from 3 o'clock (0 right, 90 down, 180 "
            "left, 270 up). An arc runs clockwise from adj1 to adj2, so "
            '{"adj1": 180, "adj2": 0} is its top half and {"adj1": 0, "adj2": 180} its '
            "bottom half."),
    },
}
ARG_HELP["copy_shape_between"].update(
    {key: ARG_HELP["set_shape"][key] for key in ("scale_text", "scale_lines", "no_wrap")})
HELP_OP_PARAM = ("An operation name, for its arguments and an example; leave it out to list "
                 "every tool and operation.")
OPS_HEADING = "Operations, one per line, as name: what it does:"
OPS_CONTINUED = ("More operations follow, in the same format, in the op parameter's "
                 "description.")
OP_PARAM_CONTINUED = (
    "The operation to run: one listed in this tool's description, or one of these, in the "
    "same format:")
HELP_USAGE = (
    "Call a tool as {op, args}. ppt_help with op gives that operation's arguments and an "
    "example.")
#: ppt_help apply_edits: the edits argument. {ops} is filled from the batch's
#: own table of ops, so the list cannot drift from what apply_edits accepts.
BATCH_EDITS_PARAM = (
    "The edits, applied in order in one save. Each edit names its op (\"op\": "
    "\"set_text\") and carries that op's fields beside it. The ops: {ops}. batchable_ops "
    "gives each one's fields, how it names its target, and an example.")
#: ppt_help for an edit that only apply_edits runs, such as set_text.
BATCH_ONLY_SUMMARY = (
    "{op} is an edit that apply_edits batches, not an operation of its own: send it in "
    "apply_edits' edits list, as the example shows.")
SETTINGS_OWNER = (
    "This server's launch settings, usually set by IT; they are fixed until the server "
    "restarts, and no tool call changes them.")
WORKFLOW_INDEX_NOTE = (
    "Call get_workflows with one of these task names for its full ordered steps.")
WORKFLOW_UNKNOWN_TASK = "There is no workflow with that name."
#: Legal text: the factual sentence followed by the copyright line and the
#: permission notice copied verbatim from python-pptx 1.0.2's LICENSE file
#: (dist-info/LICENSE, SHA-256 7652f271e46d0d533e9dc463f3b5fcbdcacf4d6a9c8d6b554d15efd0f37f6132).
TEMPLATE_NOTICE = (
    'The default deck template is a modified copy of the file shipped with python-pptx 1.0.2 (original SHA-256 e10cc9e120961f6bd4074a373c9c80d2a06c497157e8f4972977b7bea83a8f34; shipped file SHA-256 eb33715284d0a4838d3c1ffaa1245df153fef590ea5e77fa8f9d2fe4bc6867e6). Its thumbnail image and its description, last-modified-by and date properties were removed; every slide, layout, master and theme part is unchanged. It is used under the MIT License.\n'
    '\n'
    'Copyright (c) 2013 Steve Canny, https://github.com/scanny\n'
    '\n'
    'Permission is hereby granted, free of charge, to any person obtaining a copy\n'
    'of this software and associated documentation files (the "Software"), to deal\n'
    'in the Software without restriction, including without limitation the rights\n'
    'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell\n'
    'copies of the Software, and to permit persons to whom the Software is\n'
    'furnished to do so, subject to the following conditions:\n'
    '\n'
    'The above copyright notice and this permission notice shall be included in\n'
    'all copies or substantial portions of the Software.\n'
    '\n'
    'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR\n'
    'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,\n'
    'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE\n'
    'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER\n'
    'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,\n'
    'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN\n'
    'THE SOFTWARE.'
)

INTERNAL_ERROR_MESSAGE = (
    "Something inside the server failed while running this call; the arguments were not "
    "the cause.")

HINTS: dict[str, str] = {
    "AMBIGUOUS_LOCATION": (
        "Several targets matched; resend with one of the addresses the message lists."),
    "APP_BLOCKED": (
        "PowerPoint is not answering; wait, or close the dialog it is showing, then retry."),
    "APP_BUSY": (
        "PowerPoint is showing a dialog or running a command; clear it, then retry."),
    "APP_NOT_RUNNING": (
        "This operation needs PowerPoint installed, and for live operations the deck open "
        "in it; ppt_inspect powerpoint_status shows what is running."),
    "BAD_ARGS": (
        "The arguments did not fit this operation; ppt_help with this op gives its arguments "
        "and an example, and when the op belongs to another tool, details.tool names it."),
    "BAD_PARAMS": (
        "The request did not fit this operation; the message says what. ppt_help with this "
        "op gives its arguments and an example."),
    "CONFLICT": (
        "The target already exists (pass overwrite where the operation offers it), or "
        "PowerPoint disconnected during the call; the message says which."),
    "DOCUMENT_LOCKED": (
        "Close the file in PowerPoint, or wait for the other process, and retry; operations "
        "with a live argument can edit the open copy instead (live='auto')."),
    "INTERNAL_ERROR": (
        "Try once more; if it repeats, report it with the version from ppt_read "
        "get_server_info."),
    "LIMIT_EXCEEDED": (
        "The answer would pass the output limit; narrow the request (fewer slides, a scope, "
        "or limit and offset where offered)."),
    "NOT_AVAILABLE": (
        "This server's launch settings turned this off, usually by IT; no tool call can turn "
        "it back on."),
    "NOT_FOUND": (
        "Re-read the deck with ppt_read get_presentation_view or get_slide_info for the "
        "current slides, shape ids and anchors."),
    "PROTECTED": (
        "Remove the protection in PowerPoint and retry, or open the deck there "
        "and edit it live (live='auto')."),
    "PROTECTED_VIEW": "Click Enable Editing in PowerPoint, then retry.",
    "STALE_ANCHOR": (
        "The deck changed since the view was taken; re-run ppt_read get_presentation_view "
        "and resend with fresh anchors."),
    "TIMEOUT": (
        "When details.saved is false, nothing was saved and the call can be retried or "
        "narrowed; when it is null, the change had started, so check the deck before "
        "retrying."),
    "UNSUPPORTED_CONTENT": (
        "The deck holds something this operation does not handle, or is protected or "
        "damaged; the message names it, and the file was not changed."),
    "VALIDATION_FAILED": (
        "The original file was not changed; the message says what the produced file "
        "failed."),
}

# Final writer overrides for operation lines. Other operations use the first
# sentence of the already-shipped docstring unchanged. Target about 110
# characters, verb first, with the one surprise a caller needs.
OP_LINES: dict[str, str] = {
    # ppt_read
    "get_presentation_view": (
        "Read the deck as an anchored markdown view with shape and table-cell anchors; the "
        "cheap first read, and apply_edits takes its anchors."),
    "get_presentation_info": (
        "Read deck-level facts: slide count and size, slide order with ids and titles, "
        "masters, layouts and sections."),
    "get_slide_info": (
        "Read one slide in depth: layout, hidden flag, and every shape with its id, kind, "
        "geometry and text preview."),
    "get_text": (
        "Read the deck's text in reading order, optionally with speaker notes; pages in whole "
        "slides."),
    "find_text": (
        "Find text deck-wide or in scope (slides, a range, or slide_numbers from 1), exact "
        "case unless match_case=false; each match gives its slide, shape, paragraph and "
        "character offsets for edits."),
    "list_elements": (
        "List one kind of element (slides, shapes, placeholders, tables, charts, images, "
        "layouts, ...) with ids and locations."),
    "list_hyperlinks": (
        "List every link and click action on shapes, text, tables, notes, SmartArt, layouts "
        "and masters, broken ones flagged."),
    "get_table": (
        "Read one table in full: cell text, merge regions, column widths and row heights, "
        "with 0-based addresses."),
    "get_chart_data": (
        "Read what a chart holds: series names, categories, values and axis titles, from the "
        "chart's cached values."),
    "get_notes": (
        "Read one slide's speaker notes with their paragraph and run structure, to edit "
        "without losing formatting."),
    "extract_text": (
        "Read all slide text and every slide's speaker notes in one call; pages by slide."),
    "list_comments": (
        "List every comment, modern threads with replies and legacy comments, with the ids "
        "the comment operations take."),
    "list_equations": (
        "List every equation in slides, tables and notes with its location and characters."),
    "get_media_playback": (
        "Read how an embedded video or audio plays: start, loop, rewind, mute, volume and "
        "full screen."),
    "get_transitions": (
        "Read each slide's transition: kind, direction, speed, duration and advance "
        "settings."),
    "list_animations": (
        "Read one slide's animations in play order, with effects this server does not write "
        "counted separately."),
    "get_server_info": (
        "Report product, edition, version, tool and operation totals, platform and Python; "
        "needs no deck."),
    "get_workflows": (
        "List task recipes, or return one recipe's ordered steps naming the exact tool and "
        "operation."),
    # ppt_inspect
    "audit_accessibility": (
        "Audit alt text, table headers, reading order, missing titles and contrast; each "
        "finding names the fixing call."),
    "check_layout": (
        "Check slides for overlap, off-slide shapes, tiny text, contrast and more; findings "
        "name the fixing call."),
    "font_inventory": (
        "List every font in use and where, plus fonts declared where no visible text uses "
        "them; run before replace_fonts."),
    "comment_report": (
        "Report every comment thread by slide, with authors, dates, replies and resolved "
        "state, plus a markdown summary."),
    "compare_decks": (
        "Compare two decks slide by slide (matched by id, title, then position): text, "
        "layout, shape and notes changes."),
    "deck_statistics": (
        "Count slides, words in body and notes, shapes, images, tables and charts, and "
        "estimate speaking time."),
    "diagnose": (
        "Check this environment and, optionally, one file; output leaves out local paths "
        "unless verbose=true."),
    "extract_brand": (
        "Read a deck's theme colors, fonts and most-used literal fills, for apply_brand on "
        "another deck."),
    "zombie_check": (
        "Count running PowerPoint processes to spot one left behind by automation; never "
        "ends any."),
    "get_theme": (
        "Read a master's theme: its 12 color slots and its heading and body fonts."),
    "get_autofit_state": (
        "Report each text shape's autofit mode and shrink level, to find crowded slides."),
    "get_document_properties": (
        "Read the deck's document properties (title, author, dates, company and "
        "counters)."),
    "get_footer_support": (
        "Check whether a slide's design can show a footer, slide number and date; run before "
        "set_footer."),
    "list_master_elements": (
        "List a master's layouts, placeholders, decoration shapes and backgrounds, with "
        "which slides use each layout."),
    "get_export_engines": (
        "Report which render engines exist here (PowerPoint or LibreOffice) and what each "
        "can export."),
    "live_status": (
        "Check whether the user's PowerPoint responds, and which open decks have unsaved "
        "changes."),
    "powerpoint_status": (
        "Report whether PowerPoint is installed and running, and which decks the user has "
        "open."),
    # ppt_edit
    "apply_edits": (
        "Apply many edits (text, paragraphs, shapes, tables, speaker notes) in one call: one "
        "backup and one save; one stale anchor refuses all."),
    "search_and_replace": (
        "Find and replace across the deck, SmartArt text included, keeping each run's "
        "formatting."),
    "set_placeholder_text": (
        "Fill a layout placeholder (title, subtitle, body) with text that takes the layout's "
        "styling."),
    "fit_text": (
        "Shrink overflowing text to fit its box; this is an estimate, so check with "
        "export_slide_image."),
    "set_bullets": (
        "Set character, numbered or no bullets on a text shape's paragraphs."),
    "insert_textbox": (
        "Add a free text box at a position in inches, for labels and callouts."),
    "set_notes": (
        "Replace one slide's speaker notes, or many in one save with notes; dropping bold, "
        "bullets or links needs flatten=true."),
    "delete_notes": "Remove one slide's speaker notes entirely.",
    "set_hyperlink": (
        "Set a hyperlink (a URL or a jump to a slide) on a shape or text range, replacing any "
        "link there."),
    "insert_slide": (
        "Add a slide from a layout (by name or index), with the layout's placeholders ready "
        "to fill."),
    "delete_slide": (
        "Delete a slide and everything only it used; links to it are disabled. "
        "set_slide_hidden keeps it instead."),
    "duplicate_slide": (
        "Copy a slide with its notes, charts and embedded objects; the copy lands after the "
        "original by default."),
    "reorder_slides": (
        "Put every slide in a new order in one call; the order must list every slide "
        "exactly once."),
    "move_slide": "Move one slide to a new 0-based position.",
    "set_slide_hidden": (
        "Hide or unhide a slide; hidden slides keep their content but are skipped in the "
        "show."),
    "manage_section": (
        "Create, rename or delete a section, or move a slide into one (action); slides are "
        "never deleted."),
    "set_language": (
        "Set the proofing language on every text run across slides, layouts, masters, notes "
        "and charts."),
    # ppt_graphics
    "insert_shape": (
        "Insert a native shape (a preset like rect or ellipse, or a freeform path) at a "
        "position in inches."),
    "insert_connector": (
        "Draw a straight, elbow or curved connector; ends glued to shapes follow them when "
        "they move."),
    "set_shape": (
        "Edit one shape by id or name, inside groups too: position, size, rotation, fill, "
        "line, effect, text or name; only the values given change."),
    "set_shape_adjustments": (
        "Read or set a preset shape's adjust handles: an arc's start and end angles, a "
        "rounded corner, an arrow's head, an elbow's bend."),
    "delete_shape": (
        "Delete one shape, group or connector by id; connectors glued to it are kept and "
        "unglued."),
    "group_shapes": "Group two or more shapes by id into one group, like Ctrl+G.",
    "ungroup_shapes": (
        "Ungroup one group by id, keeping each member's position and size, like "
        "Ctrl+Shift+G."),
    "align_shapes": (
        "Align shapes left, center, right, top, middle or bottom, to each other or to the "
        "slide."),
    "distribute_shapes": (
        "Space three or more shapes evenly across or down; the outer two stay put."),
    "set_z_order": (
        "Bring a shape to the front or back, or one step forward or backward."),
    "svg_to_shapes": (
        "Turn SVG markup into a group of native, editable shapes (no images); unsupported "
        "features are reported."),
    "set_diagram_text": (
        "Replace the text of nodes in an existing SmartArt diagram, leaving its layout and "
        "styling alone."),
    "generate_diagram": (
        "Build a timeline, org chart, matrix, cycle or comparison diagram of editable shapes "
        "from one spec."),
    "copy_format": (
        "Copy one shape's fill, line, effects and text style onto other shapes."),
    "copy_position": (
        "Copy one shape's position and size onto the same-named shape on other slides."),
    "copy_shape_between": (
        "Copy one shape or group from a slide of another deck (or this one) onto a slide, "
        "with its pictures and charts; drop members, resize, scale text."),
    "insert_image": (
        "Place a picture (PNG, JPEG, GIF, BMP or TIFF) on a slide at a position in inches."),
    "replace_image": (
        "Swap the picture in an image shape, keeping its position, size, crop and effects."),
    "set_image": (
        "Move, resize, crop, rename or set alt text on a picture by shape id; only the values "
        "given change."),
    "replace_image_everywhere": (
        "Swap every copy of one picture across slides, layouts and masters, keeping each "
        "one's geometry."),
    "insert_video": "Embed an MP4 video on a slide; other formats are refused.",
    "insert_audio": "Embed MP3, M4A or WAV audio on a slide; other formats are refused.",
    "set_media_playback": (
        "Set how a video or audio plays: start, loop, rewind, mute, volume, full screen; "
        "only the values given change."),
    "insert_equation": (
        "Insert a LaTeX equation as native, editable PowerPoint math in a new text box."),
    "add_equation_to_shape": (
        "Add a LaTeX equation as a new math paragraph inside an existing shape's text."),
    "set_alt_text": (
        "Set or clear the alt text screen readers read, on any shape by id."),
    "set_reading_order": (
        "Set a slide's screen-reader order by listing every top-level shape id; this also "
        "changes stacking order."),
    # ppt_tables
    "create_table": (
        "Insert a native table at a position in inches, optionally filled from rows of "
        "data."),
    "set_table_cells": (
        "Write many table cells in one save: text, and per-cell bold, size, color and fill."),
    "merge_cells": (
        "Merge a rectangle of cells (0-based, inclusive); text from covered cells moves to the "
        "first cell."),
    "unmerge_cells": (
        "Split one merged region back into cells; its text stays in the top-left cell."),
    "insert_table_rows": "Insert empty rows before a 0-based row (at = row count appends).",
    "delete_table_rows": "Delete rows starting at a 0-based row, with their content.",
    "insert_table_cols": "Insert empty columns before a 0-based column.",
    "delete_table_cols": "Delete columns starting at a 0-based column, with their content.",
    "format_table_cells": (
        "Set borders, fill, margins and vertical alignment on a range of cells; text is left "
        "alone."),
    "set_column_widths": "Set table column widths in inches, for all columns or some.",
    "set_row_heights": (
        "Set table row heights in inches; rows still grow to fit their text."),
    "apply_table_style": (
        "Apply one of PowerPoint's built-in table styles by name, with its stripe and header "
        "options."),
    "import_table": (
        "Build or refill a table from a CSV or JSON file."),
    "create_chart": (
        "Insert a native, editable chart (bar, column, line, pie, scatter or combo) from "
        "categories and series."),
    "update_chart_data": (
        "Replace a chart's data in place, keeping its type and styling; the series count "
        "must stay the same."),
    "format_chart": (
        "Format a chart's titles, legend, number format, gridlines and data labels."),
    # ppt_design
    "set_slide_size": (
        "Change the slide size (16:9, 4:3, A4, Letter or inches); content is not rescaled."),
    "apply_layout": (
        "Switch a slide to another layout; matching placeholders keep their content, others "
        "are reported."),
    "set_theme_colors": "Set any of a master's 12 theme color slots to hex colors.",
    "set_theme_fonts": "Set a master's theme heading and body fonts.",
    "apply_brand": (
        "Apply an extract_brand result (theme colors and fonts) to every master of this "
        "deck."),
    "set_slide_background": (
        "Set or clear one slide's background: a color, gradient or picture."),
    "set_master_placeholder": (
        "Set a master placeholder's position, size and text defaults, which every slide "
        "inherits."),
    "set_layout_placeholder": (
        "Set a layout placeholder's position, size and text defaults, for slides using that "
        "layout."),
    "add_layout_placeholder": (
        "Add a placeholder to a layout; existing slides do not gain it, new slides do."),
    "remove_layout_placeholder": (
        "Remove a placeholder from a layout; refuses while slides use it unless "
        "force=true."),
    "insert_master_shape": (
        "Put a decoration shape (logo box, footer bar, rule line) on a master or one "
        "layout."),
    "delete_master_shape": (
        "Delete a decoration shape from a master or layout by id; placeholders are refused."),
    "create_layout": (
        "Create a new layout under a master, copied from an existing layout or blank."),
    "set_master_background": (
        "Set the background of a master or one layout: a color or a gradient."),
    "replace_fonts": (
        "Replace typefaces across the whole deck, including charts, SmartArt and unused "
        "declarations."),
    "replace_colors": (
        "Replace literal colors across the deck, optionally with theme colors that follow "
        "later theme edits."),
    "set_transition": (
        "Set or remove a slide transition (fade, push, wipe, split, cut or random)."),
    "add_entrance_animation": (
        "Add an appear, fade or wipe entrance animation to a shape, optionally paragraph by "
        "paragraph."),
    "clear_animations": (
        "Remove a slide's animations, or one shape's; transitions stay."),
    # ppt_review
    "add_comment": (
        "Add a modern threaded comment to a slide; a deck with legacy comments is refused. "
        "Leave author out to sign with the user's own name, or pass author=\"assistant\" to "
        "sign as \"AI Assistant via KS4PPT\"."),
    "reply_to_comment": (
        "Reply to a modern comment thread (reply to the thread root, not to a reply). Leave "
        "author out to sign with the user's own name, or pass author=\"assistant\" to sign "
        "as \"AI Assistant via KS4PPT\"."),
    "resolve_comment": "Mark a modern comment thread resolved, or reopen it.",
    "delete_comment": (
        "Delete a comment thread, or one reply; a thread with replies needs "
        "cascade_replies=true."),
    "anonymize_deck": (
        "Replace author names in properties and comments throughout the deck; a snapshot is "
        "saved first because this cannot be undone in the file."),
    # ppt_deck
    "create_presentation": (
        "Create a new .pptx, blank or from a template deck; an existing file is refused."),
    "copy_presentation": (
        "Copy a deck byte for byte to dest_path; an existing destination is refused unless "
        "overwrite=true."),
    "create_snapshot": (
        "Save a permanent copy of the deck in its hidden backup folder, or a visible one in "
        "dest_dir; never pruned or overwritten."),
    "manage_backups": (
        "List, restore or purge the automatic backups (prev and anchor slots); purge is a dry "
        "run unless dry_run=false."),
    "copy_slide_between": (
        "Copy one slide from another deck into this one; the source deck is not changed."),
    "merge_decks": "Append whole decks onto this deck, in order.",
    "split_deck": (
        "Split a deck into files by section or slide range; the source is not changed."),
    "export_pdf": (
        "Export the deck to PDF through PowerPoint, or LibreOffice when PowerPoint is absent."),
    "export_slide_image": (
        "Render slides to PNG files to check them: slides counts from 0 (or send slide_numbers, "
        "from 1) and width sets the pixels (default 1280); the deck is not changed."),
    "export_shape_image": (
        "Render one shape or group to a PNG at a set width; needs PowerPoint."),
    "export_handout": (
        "Export a handout or notes-page PDF; needs PowerPoint."),
    "export_table": "Export a table's content as CSV or JSON, to a file or inline.",
    "set_document_properties": (
        "Set the deck's document properties; an empty string clears a field."),
    "generate_agenda_slide": (
        "Insert an agenda slide listing the sections, each linked to its first slide."),
    "refresh_agenda_slide": (
        "Rebuild the agenda slide after reordering or section changes."),
    "compress_deck": (
        "Shrink a deck by downsampling images; purge_unused=true also drops unused parts, "
        "and dry_run=true reports without changing anything."),
    "set_show_properties": (
        "Set how the slide show runs: presenter, window or kiosk, looping, timings and "
        "range."),
    "manage_custom_show": (
        "Create, rename, delete or list custom shows (named slide sequences)."),
    "set_footer": (
        "Set footer text, slide numbers and dates per slide or deck-wide, where the design "
        "allows them."),
    "validate": (
        "Check the deck is sound: the file structure, then a real open in a hidden PowerPoint "
        "when available."),
    "live_save": (
        "Save the open deck, only when the user asks; live edits stay unsaved for them to "
        "review. A first key saved_user_work means the save wrote their own edits too."),
    "live_scroll_to": (
        "Scroll the open deck's window to a slide so the user can watch edits land."),
}


def all_strings() -> list[str]:
    """Return every writer-owned string, for the tests that scan the copy."""
    out = [
        INSTRUCTIONS,
        OP_PARAM,
        ARGS_PARAM,
        SLIDE_RULE,
        SLIDE_PARAM,
        SLIDE_NUMBER_PARAM["description"],
        SLIDE_NUMBERS_PARAM["description"],
        SCOPE_TWIN_PARAM,
        *TWIN_HELP.values(),
        SHAPE_PARAM,
        *(text for notes in ARG_NOTES.values() for text in notes.values()),
        *(text for notes in ARG_HELP.values() for text in notes.values()),
        HELP_OP_PARAM,
        OPS_HEADING,
        OPS_CONTINUED,
        OP_PARAM_CONTINUED,
        HELP_USAGE,
        BATCH_EDITS_PARAM,
        BATCH_ONLY_SUMMARY,
        SETTINGS_OWNER,
        WORKFLOW_INDEX_NOTE,
        WORKFLOW_UNKNOWN_TASK,
        OPEN_DECK_SENTENCE,
        TEMPLATE_NOTICE,
        INTERNAL_ERROR_MESSAGE,
        *ENGINE_COPY,
        *HINTS.values(),
        *OP_LINES.values(),
    ]
    for slot in TOOLS.values():
        out.extend((slot["title"], slot["description"]))
    return out


__all__ = [
    "ARG_NOTES", "ARGS_PARAM", "BAD_ARGS_HINT", "BAD_ARGS_MESSAGE", "BATCH_ATOMIC_ONLY",
    "HELP_OP_PARAM",
    "HELP_USAGE", "HINTS", "INSTRUCTIONS", "INTERNAL_ERROR_MESSAGE",
    "LIMIT_EXCEEDED_HINT", "LIMIT_EXCEEDED_MESSAGE", "MISSING_TITLE",
    "OPEN_DECK_SENTENCE", "OP_NOT_FOUND_HINT", "SET_TEXT_FORMATTING",
    "OP_NOT_FOUND_MESSAGE",
    "OP_LINES", "OP_PARAM", "OP_PARAM_CONTINUED",
    "OPS_CONTINUED", "OPS_HEADING", "SCOPE_PARAM", "SETTINGS_OWNER", "SLIDE_NUMBER_PARAM",
    "SLIDE_NUMBERS_PARAM",
    "SLIDE_PARAM", "SLIDE_RULE", "SLOT",
    "TEMPLATE_NOTICE", "TIMEOUT_MESSAGE", "TITLE_READ_AFTER_BODY", "TOOLS",
    "VIEW_OUTLINE_OMISSIONS", "VIEW_TEXT_OMISSIONS",
    "WRONG_TOOL_MESSAGE", "all_strings",
    "WORKFLOW_INDEX_NOTE", "WORKFLOW_UNKNOWN_TASK",
]
