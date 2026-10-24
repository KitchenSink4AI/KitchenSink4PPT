"""Workflow guidance: recommended tool sequences for common multi-step tasks.

Every recipe names the fixed surface tool and operation for each step. Tests
assert both against the live registry, so a workflow cannot point at an
operation that does not exist or at the wrong group tool.

The sequences are recommendations, not scripts: steps marked optional can be
skipped, and the notes carry the judgment calls a caller should know about.
"""

from __future__ import annotations

from ..core.errors import PptMcpError
from ..surface import copy as _copy

# task -> {summary, steps: [{tool, op, why, optional?}], notes: [...]}
WORKFLOWS: dict[str, dict] = {
    "migrate-from-v1": {
        "summary": (
            "Find the 2.0 call for any PowerPoint 1.3.1 tool name: the map "
            "names the grouped tool each one now runs under."
        ),
        "steps": [],
        "notes": [
            "Each 1.3.1 tool is now an operation of the grouped tool the map "
            "names; call it as <tool>(op=<old name>, args={...}). Some "
            "arguments changed, so check ppt_help for the operation before "
            "the first call.",
            "The two 1.3.1 tools that switched tool sets on and off are gone, "
            "because every operation is present from the start; their map "
            "rows have no target. To limit what the server offers, use the "
            "KS4P_READ_ONLY and KS4P_FORBID launch settings.",
        ],
        "generated_map": True,
    },
    "build-a-diagram": {
        "summary": (
            "Turn an SVG (or a from-scratch shape layout) into native, "
            "individually editable PowerPoint shapes with glued connectors, "
            "then verify by rendering."
        ),
        "steps": [
            {"tool": "ppt_read", "op": "get_presentation_view",
             "why": "see the target slide's current shapes and anchors first"},
            {"tool": "ppt_graphics", "op": "svg_to_shapes",
             "why": "compile the SVG into one grouped tree of native shapes "
                    "(no rasterization); returns every created shape id"},
            {"tool": "ppt_graphics", "op": "insert_connector", "optional": True,
             "why": "add glued arrows between shapes by id; glued ends follow "
                    "the shape when it moves"},
            {"tool": "ppt_deck", "op": "export_slide_image",
             "why": "render the slide to PNG and LOOK at it; layout bugs are "
                    "visual, not structural"},
            {"tool": "ppt_graphics", "op": "set_shape",
             "why": "nudge position, size, fill, or text by shape id and "
                    "re-render until it is right"},
        ],
        "notes": [
            "insert_shape builds presets and freeform geometry directly when "
            "there is no SVG to compile.",
            "align_shapes and distribute_shapes do mechanical cleanup that "
            "eyeballing coordinates cannot.",
        ],
    },
    "build-a-table-report": {
        "summary": (
            "Create or import a data table on a slide, style it, and export "
            "it back out as CSV/JSON."
        ),
        "steps": [
            {"tool": "ppt_tables", "op": "create_table",
             "why": "new table with optional initial data; import_table "
                    "builds one from an existing CSV/JSON file instead"},
            {"tool": "ppt_tables", "op": "set_table_cells",
             "why": "bulk cell edits: text plus per-cell bold/size/fill in "
                    "one call"},
            {"tool": "ppt_tables", "op": "merge_cells", "optional": True,
             "why": "span header cells; unmerge_cells reverses it"},
            {"tool": "ppt_tables", "op": "apply_table_style", "optional": True,
             "why": "one of the 74 built-in PowerPoint table styles by name"},
            {"tool": "ppt_deck", "op": "export_table", "optional": True,
             "why": "round-trip the finished table to CSV/JSON for records"},
        ],
        "notes": [
            "create_chart / update_chart_data live in the same tool for "
            "bar, column, line, and pie charts.",
            "Cell addresses are 0-based (row, col) into the full grid; merge "
            "continuations refuse edits and name the origin cell.",
        ],
    },
    "template-deck-setup": {
        "summary": (
            "Start a new deck FROM a template (theme, layouts, and masters "
            "intact) and shape its skeleton."
        ),
        "steps": [
            {"tool": "ppt_deck", "op": "create_presentation",
             "why": "byte-copies the template so theme colors, fonts, "
                    "layouts, and masters all carry over"},
            {"tool": "ppt_read", "op": "list_elements",
             "why": "kind='layouts' shows which layouts the template offers, "
                    "by name and part"},
            {"tool": "ppt_edit", "op": "insert_slide",
             "why": "add slides by layout name or index; repeat per section"},
            {"tool": "ppt_edit", "op": "set_placeholder_text",
             "why": "titles and body placeholders inherit the template's "
                    "styling automatically"},
            {"tool": "ppt_edit", "op": "move_slide", "optional": True,
             "why": "adjust running order; set_slide_hidden parks optional "
                    "slides without deleting them"},
        ],
        "notes": [
            "keep_slides=False (default) strips the template's slides and "
            "keeps only its design machinery.",
        ],
    },
    "render-and-review": {
        "summary": (
            "The verification loop: render slides to images, inspect, fix, "
            "re-render; finish with a full-deck validation and PDF."
        ),
        "steps": [
            {"tool": "ppt_inspect", "op": "get_export_engines",
             "why": "confirm which render engine exists here (PowerPoint COM "
                    "is ground truth; LibreOffice drifts on themes/fonts)"},
            {"tool": "ppt_deck", "op": "export_slide_image",
             "why": "PNG per slide; read the image to catch overflow, "
                    "overlap, and color mistakes"},
            {"tool": "ppt_inspect", "op": "get_autofit_state", "optional": True,
             "why": "flags text boxes that are shrinking or overflowing "
                    "without opening PowerPoint"},
            {"tool": "ppt_deck", "op": "validate",
             "why": "package payload check plus a real invisible-PowerPoint "
                    "open; repair prompts count as failure"},
            {"tool": "ppt_deck", "op": "export_pdf",
             "why": "the shareable artifact once the deck passes review"},
        ],
        "notes": [
            "ppt_edit/set_notes writes the speaker notes reviewers read "
            "alongside the rendered slides.",
            "ppt_graphics/set_shape and ppt_edit/format_text fix what the "
            "render shows; table repairs use ppt_tables.",
            "Exports never replace a file on their own: a re-render into the "
            "same output_dir sends overwrite=true.",
        ],
    },
    "refresh-an-existing-deck": {
        "summary": (
            "Inherit a deck somebody else built and fix it: copy, triage, "
            "render the flagged slides, repair, re-render, validate."
        ),
        "steps": [
            {"tool": "ppt_deck", "op": "copy_presentation",
             "why": "work on a copy under a new name; the original stays "
                    "recoverable however the pass goes"},
            {"tool": "ppt_read", "op": "get_presentation_view",
             "why": "the cheap read of what is actually there, with the "
                    "anchors apply_edits takes"},
            {"tool": "ppt_inspect", "op": "check_layout",
             "why": "overlap, off-slide, tiny or overflowing text, contrast; "
                    "findings carry shape ids and a fix call"},
            {"tool": "ppt_inspect", "op": "audit_accessibility", "optional": True,
             "why": "alt text, reading order, and title coverage, which the "
                    "layout checks do not cover"},
            {"tool": "ppt_inspect", "op": "get_autofit_state", "optional": True,
             "why": "names the shapes already shrinking their text, the "
                    "classic crowded-slide signal"},
            {"tool": "ppt_deck", "op": "export_slide_image",
             "why": "render the FLAGGED slides only and look; a heuristic "
                    "finding is a place to look, not a verdict"},
            {"tool": "ppt_edit", "op": "apply_edits",
             "why": "batch the repairs in one atomic save; set_shape, "
                    "format_text and set_table_cells all ride here"},
            {"tool": "ppt_deck", "op": "validate",
             "why": "payload check plus a real invisible-PowerPoint open "
                    "before the deck goes anywhere"},
            {"tool": "ppt_deck", "op": "export_pdf", "optional": True,
             "why": "the shareable artifact once the deck passes"},
        ],
        "notes": [
            "Render only the slides check_layout flagged; rendering the "
            "whole deck takes far longer.",
            "check_layout resolves inherited font size and inherited "
            "placeholder geometry, so findings name the layout or master "
            "when that is where the fix belongs.",
            "Re-run get_presentation_view after a batch that deletes or "
            "reorders shapes; anchors from the earlier read go stale.",
        ],
    },
    "one-call-diagram": {
        "summary": (
            "Build a timeline, org chart, matrix, cycle, or comparison "
            "diagram from one spec call, then tweak it by role ids."
        ),
        "steps": [
            {"tool": "ppt_graphics", "op": "generate_diagram",
             "why": "kind + spec + inch box in, a grouped tree of native "
                    "shapes with glued connectors out; the result maps "
                    "roles to shape ids"},
            {"tool": "ppt_deck", "op": "export_slide_image",
             "why": "render and LOOK; crowding warnings in the result tell "
                    "you where to check first"},
            {"tool": "ppt_graphics", "op": "set_shape",
             "why": "move or restyle any piece by its role's shape id; "
                    "glued connectors follow"},
        ],
        "notes": [
            "generate_diagram builds five kinds of diagram; svg_to_shapes "
            "covers any other diagram.",
            "Specs take theme colors by default, so a template change "
            "recolors the whole diagram.",
        ],
    },
    "batch-edit-from-view": {
        "summary": (
            "The token-cheap editing loop on the fixed surface: read the anchored "
            "view once, then apply many edits in one atomic call."
        ),
        "steps": [
            {"tool": "ppt_read", "op": "get_presentation_view",
             "why": "one anchored markdown projection of the deck: slide "
                    "anchors [s:id], shape anchors [a:hex], table cells "
                    "t:hex:rNcN"},
            {"tool": "ppt_edit", "op": "apply_edits",
             "why": "batch of {op, anchor, ...} edits; every anchor is "
                    "resolved BEFORE anything mutates, one backup and one "
                    "save for the whole batch"},
            {"tool": "ppt_read", "op": "get_presentation_view", "optional": True,
             "why": "re-read only when anchors went stale (shape deleted or "
                    "re-created); surviving anchors stay valid"},
        ],
        "notes": [
            "A stale or ambiguous anchor refuses the WHOLE batch and lists "
            "every failed op index; nothing is half-applied.",
            "search_and_replace inside apply_edits covers deck-wide text "
            "swaps without any anchors at all.",
        ],
    },
    "animate-a-build": {
        "summary": (
            "Give a deck slide transitions and click-by-click entrance "
            "builds, honestly bounded to the verified effect subset."
        ),
        "steps": [
            {"tool": "ppt_read", "op": "get_transitions",
             "why": "see what the deck already carries before overwriting "
                    "(morph and friends are reported, not authored)"},
            {"tool": "ppt_design", "op": "set_transition",
             "why": "one call for one slide or the whole deck; kind, "
                    "direction, millisecond duration, auto-advance"},
            {"tool": "ppt_design", "op": "add_entrance_animation",
             "why": "appear/fade/wipe per shape; trigger='click' builds "
                    "click groups, by_paragraph=True builds bullet lists "
                    "line by line"},
            {"tool": "ppt_read", "op": "list_animations",
             "why": "verify play order, triggers, and targets after "
                    "authoring"},
            {"tool": "ppt_design", "op": "clear_animations", "optional": True,
             "why": "reset a slide (or one shape) when a build sequence "
                    "should start over"},
        ],
        "notes": [
            "delete_shape prunes its own animation nodes; foreign effect "
            "trees survive edits untouched and are listed honestly.",
        ],
    },
    "review-cycle": {
        "summary": (
            "Run a comment-driven review pass: read every thread, reply, "
            "resolve what is settled, and report the deck's review state."
        ),
        "steps": [
            {"tool": "ppt_inspect", "op": "comment_report",
             "why": "the whole deck's threads grouped by slide, with a "
                    "markdown rendering for review notes"},
            {"tool": "ppt_review", "op": "add_comment",
             "why": "new feedback anchored to a slide, a position, or a "
                    "shape"},
            {"tool": "ppt_review", "op": "reply_to_comment",
             "why": "answer inside the thread; replies nest under the "
                    "root"},
            {"tool": "ppt_review", "op": "resolve_comment",
             "why": "mark settled threads; resolved=False reopens"},
            {"tool": "ppt_review", "op": "delete_comment", "optional": True,
             "why": "remove a thread (cascade_replies=True) or a single "
                    "reply once it is truly obsolete"},
        ],
        "notes": [
            "Decks with classic (legacy) comments are read-only here: the "
            "two comment systems never mix in one file.",
        ],
    },
    "live-session": {
        "summary": (
            "Edit a deck WHILE the user has it open in PowerPoint: route "
            "edits live, show the user, save only on request."
        ),
        "steps": [
            {"tool": "ppt_inspect", "op": "live_status",
             "why": "confirm PowerPoint is responsive and the target file "
                    "is open (and not read-only) before editing"},
            {"tool": "ppt_read", "op": "get_text",
             "why": "dual-mode reads route to the open copy automatically "
                    "when the file is locked (live='auto', the default)"},
            {"tool": "ppt_edit", "op": "set_placeholder_text",
             "why": "dual-mode edits land in the OPEN copy and stay "
                    "unsaved; the user sees them immediately"},
            {"tool": "ppt_deck", "op": "live_scroll_to",
             "why": "bring the edited slide into the user's view without "
                    "touching their selection"},
            {"tool": "ppt_deck", "op": "live_save",
             "why": "only when the user asks: nothing writes the user's "
                    "file until this call (or their own Ctrl+S); the "
                    "result's first key, saved_user_work, says when the "
                    "save also wrote their own unsaved work; "
                    "unsaved_before_our_edits says whether the "
                    "deck had unsaved changes before this server's first "
                    "edit in this run (true: the save also wrote their "
                    "own unsaved work; false: it was clean, and typing "
                    "they did after that is not told apart from the "
                    "edits; null: not known)"},
        ],
        "notes": [
            "Eleven operations are dual-mode (get_text, get_slide_info, "
            "search_and_replace, set_placeholder_text, insert_textbox, "
            "format_text, insert_shape, set_shape, insert_slide, "
            "delete_slide, set_notes); live='off' pins them file-only.",
            "Regex replace, character-range formatting, and other "
            "file-only parameters refuse loudly on the live path.",
        ],
    },
    "cross-deck-assembly": {
        "summary": (
            "Assemble a deck from slides living in other presentations, "
            "keeping either the destination design or the source's."
        ),
        "steps": [
            {"tool": "ppt_read", "op": "get_presentation_info",
             "why": "inventory both decks first: slide ids, sizes, and "
                    "section layout"},
            {"tool": "ppt_deck", "op": "copy_slide_between",
             "why": "pull a source slide in; design='link' restyles onto "
                    "the destination layout with appearance carried, "
                    "'import' keeps the source design family"},
            {"tool": "ppt_edit", "op": "move_slide",
             "why": "place the arrival in the running order"},
            {"tool": "ppt_deck", "op": "validate",
             "why": "cross-package surgery earns a real opens-clean check "
                    "before anyone presents the result"},
        ],
        "notes": [
            "Slide-size mismatches copy with a warning and no rescaling; "
            "check the seams with export_slide_image.",
            "Same-file copies refuse by design; duplicate_slide covers "
            "that case.",
        ],
    },
    "rebrand-pipeline": {
        "summary": (
            "Move a deck onto a new brand COMPLETELY: theme first, then "
            "the literal colors and fonts the theme cannot reach."
        ),
        "steps": [
            {"tool": "ppt_inspect", "op": "extract_brand",
             "why": "read the SOURCE deck's palette: theme slots, fonts, "
                    "and the most-used literal hex fills"},
            {"tool": "ppt_design", "op": "apply_brand",
             "why": "write the palette onto the target's theme(s); "
                    "theme-linked content re-resolves, literals do not"},
            {"tool": "ppt_inspect", "op": "font_inventory",
             "why": "census what actually remains: literal typefaces, "
                    "chart fonts, phantom declarations"},
            {"tool": "ppt_design", "op": "replace_colors",
             "why": "to_theme=True converts leftover literal hex to theme "
                    "references so they follow future edits; plain "
                    "hex-to-hex remaps work too"},
            {"tool": "ppt_design", "op": "replace_fonts",
             "why": "swap stray typefaces everywhere, charts and phantom "
                    "declarations included"},
            {"tool": "ppt_inspect", "op": "check_layout", "optional": True,
             "why": "contrast check after recoloring; new brand colors can "
                    "sink text into backgrounds"},
        ],
        "notes": [
            "apply_brand's result says honestly which slots and fonts it "
            "set; the sweeps close the literal-value gap it reports.",
        ],
    },
    "master-restyle": {
        "summary": (
            "Restyle a deck at the ROOT: edit the master and layout "
            "placeholders, decorations, and backgrounds every slide "
            "inherits."
        ),
        "steps": [
            {"tool": "ppt_inspect", "op": "list_master_elements",
             "why": "the orientation read: placeholders, layouts, slide "
                    "usage, txStyles defaults, background ownership"},
            {"tool": "ppt_design", "op": "set_master_placeholder",
             "why": "deck-wide text defaults ('all titles 28pt') and "
                    "geometry on the master placeholder"},
            {"tool": "ppt_design", "op": "set_layout_placeholder", "optional": True,
             "why": "per-layout overrides sitting between master and "
                    "slides"},
            {"tool": "ppt_design", "op": "insert_master_shape", "optional": True,
             "why": "logo box, footer bar, or rule line rendered by every "
                    "derived slide"},
            {"tool": "ppt_design", "op": "set_master_background", "optional": True,
             "why": "master- or layout-level background; slides with their "
                    "OWN background are flagged, not overwritten"},
            {"tool": "ppt_inspect", "op": "check_layout",
             "why": "verify the inherited result on real slides before "
                    "calling the restyle done"},
        ],
        "notes": [
            "Slides carrying explicit run formatting keep it; every edit "
            "reports affected_slides so surprises are visible.",
            "create_layout clones or builds new layouts when the deck "
            "needs a variant instead of an edit.",
        ],
    },
    "deck-merge": {
        "summary": (
            "Combine whole presentations into one deck with sections, an "
            "agenda, and a validation gate."
        ),
        "steps": [
            {"tool": "ppt_read", "op": "get_presentation_info",
             "why": "inventory the destination first: slide count, size, "
                    "existing sections"},
            {"tool": "ppt_deck", "op": "merge_decks",
             "why": "append each source deck in order; design='link' "
                    "restyles onto the destination, 'import' keeps source "
                    "designs; sections wrap each source"},
            {"tool": "ppt_deck", "op": "generate_agenda_slide", "optional": True,
             "why": "a jump-linked agenda built from the sections the "
                    "merge just created"},
            {"tool": "ppt_inspect", "op": "deck_statistics", "optional": True,
             "why": "size and speaking-time estimate of the combined deck"},
            {"tool": "ppt_deck", "op": "validate",
             "why": "cross-package surgery earns a real opens-clean check"},
        ],
        "notes": [
            "split_deck is the inverse: one file per section or range, "
            "dependencies kept per piece.",
            "For single slides use copy_slide_between (cross-deck-assembly "
            "recipe) instead of merging a whole deck.",
        ],
    },
    "accessibility-pass": {
        "summary": (
            "Audit a deck for screen-reader readiness and repair what the "
            "audit finds, with named fixes."
        ),
        "steps": [
            {"tool": "ppt_inspect", "op": "audit_accessibility",
             "why": "one read: missing alt text, headerless tables, "
                    "reading-order mismatches, missing titles, contrast; "
                    "each finding names its repairing call"},
            {"tool": "ppt_deck", "op": "export_slide_image", "optional": True,
             "why": "LOOK at flagged slides before describing images; alt "
                    "text written blind is worse than none"},
            {"tool": "ppt_graphics", "op": "set_alt_text",
             "why": "describe each flagged shape (empty string clears a "
                    "stale description)"},
            {"tool": "ppt_graphics", "op": "set_reading_order", "optional": True,
             "why": "rewrite a slide's shape order for screen readers; "
                    "z-order changes are reported for review"},
            {"tool": "ppt_inspect", "op": "audit_accessibility",
             "why": "re-run to confirm the findings list is empty or "
                    "consciously accepted"},
        ],
        "notes": [
            "Table header semantics come from create_table/apply_table_"
            "style first_row flags (ppt_tables).",
            "Reading order IS z-order in PowerPoint; reordering can change "
            "overlaps, which the result reports.",
        ],
    },
    "equation-authoring": {
        "summary": (
            "Put real, editable math on slides from LaTeX, then verify it "
            "reads back and renders."
        ),
        "steps": [
            {"tool": "ppt_graphics", "op": "insert_equation",
             "why": "LaTeX in, native PowerPoint math out (editable in the "
                    "equation editor, not an image), in a new box"},
            {"tool": "ppt_graphics", "op": "add_equation_to_shape", "optional": True,
             "why": "append a math paragraph inside an existing text "
                    "shape at a chosen position"},
            {"tool": "ppt_read", "op": "list_equations",
             "why": "each equation in a shape's text, a table cell or the speaker "
                    "notes, with its location and its characters, which get_text and "
                    "find_text show too, also in a text box PowerPoint saved with its "
                    "equation"},
            {"tool": "ppt_deck", "op": "export_slide_image",
             "why": "render and LOOK; spacing and symbol fidelity are "
                    "visual judgments"},
        ],
        "notes": [
            "Conversion failures refuse with the converter's message and "
            "leave the file untouched.",
            "Equations survive duplicate_slide and cross-deck copies as "
            "opaque math XML.",
        ],
    },
}


def get_workflows(task: str | None = None) -> dict:
    """Serve one workflow by name, or the index of all of them."""
    if task is None:
        return {
            "workflows": {
                name: {
                    "summary": wf["summary"],
                    "tools": sorted({step["tool"] for step in wf["steps"]}),
                }
                for name, wf in WORKFLOWS.items()
            },
            "note": _copy.WORKFLOW_INDEX_NOTE,
        }
    if task not in WORKFLOWS:
        raise PptMcpError(
            f"{_copy.WORKFLOW_UNKNOWN_TASK} Requested: {task!r}. "
            f"Available: {sorted(WORKFLOWS)}"
        )
    result = {"task": task, **WORKFLOWS[task]}
    if result.get("generated_map"):
        result["map"] = migration_map_v1()
    return result


REMOVED_IN_V2 = ("enable_tools", "disable_tools")


#: Operations added after 1.3.1; they have no 1.3.1 name to route.
NEW_IN_V2 = frozenset({
    "set_shape_adjustments",
    "copy_shape_between",
    "export_shape_image",
    "update_now",
    "update_notices",
})


def migration_map_v1() -> list[dict]:
    """Generate the released 1.3.1 -> grouped-surface routing map."""
    from .. import registry

    rows = [
        {"old": name, "tool": spec.group, "op": name}
        for name, spec in sorted(registry.public_ops().items())
        if name not in NEW_IN_V2
    ]
    rows.extend({"old": name, "tool": None, "op": None} for name in REMOVED_IN_V2)
    return rows
