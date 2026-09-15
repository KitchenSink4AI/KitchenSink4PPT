"""Tool titles and the write-shape annotations, declared by name.

Every registered tool ships four MCP annotations. `readOnlyHint` lives in
core/readonly.py and is not repeated here. The other three live here:

* **title** -- the human-readable name a client shows instead of the raw
  tool name. Derived mechanically: the tool name, split on underscores,
  Title Cased, with known acronyms upper-cased, a leading `com_` replaced
  by the host application's name, and short function words kept lowercase
  inside the phrase. A handful of names too short or too generic to read
  as a title take a phrase from the first clause of their own description
  instead; those are the only hand-written strings in the table and
  docs/TOOL_TITLES.md marks each one. Titles are unique within this
  server and none exceeds 40 characters.

* **destructiveHint** -- whether the tool may perform a destructive
  update. ONE rule decides it, and docs/TOOL_ANNOTATIONS.md states the
  rule alongside a per-tool reason so a reviewer can dispute any single
  row:

      false only when every code path either (a) adds new content or a new
      file without replacing anything that was already there, creation
      refusing an existing target, or (b) changes no user data at all
      (this session's tool surface, the viewport, a read performed through
      a hidden or already-running Office instance, a read taken through a
      temporary copy).

      true otherwise. That covers everything that deletes, replaces,
      overwrites, clears, reorders, moves, applies a batch of edits, saves
      over the document the user has open, or writes an output file it may
      silently overwrite. It also covers every tool that writes into an
      existing document, because the pre-write backup these servers take
      is defeatable by the tool's own `backup=False` argument and so
      cannot be claimed as guaranteed reversibility, and everything whose
      reversibility could not be PROVEN by reading the code.

  Read-only tools carry no destructiveHint: the field is meaningful only
  when readOnlyHint is false, and a value there would be noise.

* **idempotentHint** -- true only where repeating the identical call
  obviously lands the same state: the pack switches, whose own docstrings
  say idempotent, and the whole-value setters that take an address and a
  value and carry no action selector. Everything else is left unset rather
  than guessed.

* **openWorldHint** -- false on every tool. Nothing here reaches a remote
  service.

An unclassified name RAISES at registration, the same contract
core/readonly.py already holds this surface to.
"""

from __future__ import annotations

#: tool name -> the title a client shows. Unique, at most 40
#: characters, and mirrored into docs/TOOL_TITLES.md, which a test
#: checks against this table so the published English cannot drift
#: from what goes on the wire.
TITLES: dict[str, str] = {
    "add_comment":               "Add Comment",
    "add_entrance_animation":    "Add Entrance Animation",
    "add_equation_to_shape":     "Add Equation to Shape",
    "add_layout_placeholder":    "Add Layout Placeholder",
    "align_shapes":              "Align Shapes",
    "anonymize_deck":            "Anonymize Deck",
    "apply_brand":               "Apply Brand",
    "apply_edits":               "Apply Edits",
    "apply_layout":              "Apply Layout",
    "apply_table_style":         "Apply Table Style",
    "audit_accessibility":       "Audit Accessibility",
    "check_layout":              "Check Layout",
    "clear_animations":          "Clear Animations",
    "comment_report":            "Comment Report",
    "compare_decks":             "Compare Decks",
    "compress_deck":             "Compress Deck Media",
    "copy_format":               "Copy Format",
    "copy_position":             "Copy Position",
    "copy_presentation":         "Copy Presentation",
    "copy_slide_between":        "Copy Slide Between Decks",
    "create_chart":              "Create Chart",
    "create_layout":             "Create Layout",
    "create_presentation":       "Create Presentation",
    "create_snapshot":           "Create Snapshot",
    "create_table":              "Create Table",
    "deck_statistics":           "Deck Statistics",
    "delete_comment":            "Delete Comment",
    "delete_master_shape":       "Delete Master Shape",
    "delete_notes":              "Delete Notes",
    "delete_shape":              "Delete Shape",
    "delete_slide":              "Delete Slide",
    "delete_table_cols":         "Delete Table Columns",
    "delete_table_rows":         "Delete Table Rows",
    "diagnose":                  "Diagnose Presentation",
    "disable_tools":             "Disable Tools",
    "distribute_shapes":         "Distribute Shapes",
    "duplicate_slide":           "Duplicate Slide",
    "enable_tools":              "Enable Tools",
    "export_handout":            "Export Handout",
    "export_pdf":                "Export PDF",
    "export_slide_image":        "Export Slide Image",
    "export_table":              "Export Table",
    "extract_brand":             "Extract Brand",
    "extract_text":              "Extract Text",
    "find_text":                 "Find Text",
    "fit_text":                  "Fit Text",
    "font_inventory":            "Font Inventory",
    "format_chart":              "Format Chart",
    "format_table_cells":        "Format Table Cells",
    "format_text":               "Format Text",
    "generate_agenda_slide":     "Generate Agenda Slide",
    "generate_diagram":          "Generate Diagram",
    "get_autofit_state":         "Get Autofit State",
    "get_chart_data":            "Get Chart Data",
    "get_document_properties":   "Get Document Properties",
    "get_export_engines":        "Get Export Engines",
    "get_footer_support":        "Get Footer Support",
    "get_media_playback":        "Get Media Playback",
    "get_notes":                 "Get Notes",
    "get_presentation_info":     "Get Presentation Info",
    "get_presentation_view":     "Get Presentation View",
    "get_server_info":           "Get Server Info",
    "get_slide_info":            "Get Slide Info",
    "get_table":                 "Get Table",
    "get_text":                  "Get Text",
    "get_theme":                 "Get Theme",
    "get_transitions":           "Get Transitions",
    "get_workflows":             "Get Workflows",
    "group_shapes":              "Group Shapes",
    "import_table":              "Import Table",
    "insert_audio":              "Insert Audio",
    "insert_connector":          "Insert Connector",
    "insert_equation":           "Insert Equation",
    "insert_image":              "Insert Image",
    "insert_master_shape":       "Insert Master Shape",
    "insert_shape":              "Insert Shape",
    "insert_slide":              "Insert Slide",
    "insert_table_cols":         "Insert Table Columns",
    "insert_table_rows":         "Insert Table Rows",
    "insert_textbox":            "Insert Textbox",
    "insert_video":              "Insert Video",
    "list_animations":           "List Animations",
    "list_comments":             "List Comments",
    "list_elements":             "List Elements",
    "list_equations":            "List Equations",
    "list_hyperlinks":           "List Hyperlinks",
    "list_master_elements":      "List Master Elements",
    "live_save":                 "Live Save",
    "live_scroll_to":            "Live Scroll to Location",
    "live_status":               "Live Status",
    "manage_backups":            "Manage Backups",
    "manage_custom_show":        "Manage Custom Show",
    "manage_section":            "Manage Section",
    "merge_cells":               "Merge Cells",
    "merge_decks":               "Merge Decks",
    "move_slide":                "Move Slide",
    "powerpoint_status":         "PowerPoint Status",
    "refresh_agenda_slide":      "Refresh Agenda Slide",
    "remove_hyperlink":          "Remove Hyperlink",
    "remove_layout_placeholder": "Remove Layout Placeholder",
    "reorder_slides":            "Reorder Slides",
    "replace_colors":            "Replace Colors",
    "replace_fonts":             "Replace Fonts",
    "replace_image":             "Replace Image",
    "replace_image_everywhere":  "Replace Image Everywhere",
    "reply_to_comment":          "Reply to Comment",
    "resolve_comment":           "Resolve Comment",
    "search_and_replace":        "Search and Replace",
    "set_alt_text":              "Set Alt Text",
    "set_bullets":               "Set Bullets",
    "set_column_widths":         "Set Column Widths",
    "set_diagram_text":          "Set Diagram Text",
    "set_document_properties":   "Set Document Properties",
    "set_footer":                "Set Footer",
    "set_hyperlink":             "Set Hyperlink",
    "set_image":                 "Set Image",
    "set_language":              "Set Language",
    "set_layout_placeholder":    "Set Layout Placeholder",
    "set_master_background":     "Set Master Background",
    "set_master_placeholder":    "Set Master Placeholder",
    "set_media_playback":        "Set Media Playback",
    "set_notes":                 "Set Notes",
    "set_placeholder_text":      "Set Placeholder Text",
    "set_reading_order":         "Set Reading Order",
    "set_row_heights":           "Set Row Heights",
    "set_shape":                 "Set Shape",
    "set_show_properties":       "Set Show Properties",
    "set_slide_background":      "Set Slide Background",
    "set_slide_hidden":          "Set Slide Hidden",
    "set_slide_size":            "Set Slide Size",
    "set_table_cells":           "Set Table Cells",
    "set_theme_colors":          "Set Theme Colors",
    "set_theme_fonts":           "Set Theme Fonts",
    "set_transition":            "Set Transition",
    "set_z_order":               "Set Z-Order",
    "split_deck":                "Split Deck",
    "svg_to_shapes":             "SVG to Shapes",
    "ungroup_shapes":            "Ungroup Shapes",
    "unmerge_cells":             "Unmerge Cells",
    "update_chart_data":         "Update Chart Data",
    "validate":                  "Validate Presentation",
    "zombie_check":              "Zombie Process Check",
}

#: destructiveHint: true. Deletes, replaces, overwrites, clears,
#: reorders, moves, batch-edits, saves over the open document, writes
#: an output file it may overwrite, acts on a live page, or could not
#: be proven reversible. docs/TOOL_ANNOTATIONS.md carries the reason
#: for every row.
DESTRUCTIVE: frozenset[str] = frozenset({
    "align_shapes", "anonymize_deck", "apply_brand", "apply_edits",
    "apply_layout", "apply_table_style", "clear_animations",
    "compress_deck", "copy_format", "copy_position", "delete_comment",
    "delete_master_shape", "delete_notes", "delete_shape",
    "delete_slide", "delete_table_cols", "delete_table_rows",
    "distribute_shapes", "export_handout", "export_pdf",
    "export_slide_image", "export_table", "fit_text", "format_chart",
    "format_table_cells", "format_text", "group_shapes",
    "import_table", "live_save", "manage_backups",
    "manage_custom_show", "manage_section", "merge_cells",
    "move_slide", "refresh_agenda_slide", "remove_hyperlink",
    "remove_layout_placeholder", "reorder_slides", "replace_colors",
    "replace_fonts", "replace_image", "replace_image_everywhere",
    "reply_to_comment", "resolve_comment", "search_and_replace",
    "set_alt_text", "set_bullets", "set_column_widths",
    "set_diagram_text", "set_document_properties", "set_footer",
    "set_hyperlink", "set_image", "set_language",
    "set_layout_placeholder", "set_master_background",
    "set_master_placeholder", "set_media_playback", "set_notes",
    "set_placeholder_text", "set_reading_order", "set_row_heights",
    "set_shape", "set_show_properties", "set_slide_background",
    "set_slide_hidden", "set_slide_size", "set_table_cells",
    "set_theme_colors", "set_theme_fonts", "set_transition",
    "set_z_order", "ungroup_shapes", "unmerge_cells",
    "update_chart_data", "validate"
})

#: destructiveHint: false. Every code path either only ADDS, or
#: changes no user data at all. Each row's reason is in
#: docs/TOOL_ANNOTATIONS.md.
NON_DESTRUCTIVE: frozenset[str] = frozenset({
    "add_comment", "add_entrance_animation", "add_equation_to_shape",
    "add_layout_placeholder", "copy_presentation",
    "copy_slide_between", "create_chart", "create_layout",
    "create_presentation", "create_snapshot", "create_table",
    "disable_tools", "duplicate_slide", "enable_tools",
    "generate_agenda_slide", "generate_diagram", "insert_audio",
    "insert_connector", "insert_equation", "insert_image",
    "insert_master_shape", "insert_shape", "insert_slide",
    "insert_table_cols", "insert_table_rows", "insert_textbox",
    "insert_video", "live_scroll_to", "merge_decks", "split_deck",
    "svg_to_shapes"
})

#: idempotentHint: true. Repeating the identical call lands the same
#: state. Anything absent here is left UNSET rather than guessed.
IDEMPOTENT: frozenset[str] = frozenset({
    "disable_tools", "enable_tools", "set_alt_text", "set_bullets",
    "set_column_widths", "set_diagram_text", "set_document_properties",
    "set_footer", "set_hyperlink", "set_image", "set_language",
    "set_layout_placeholder", "set_master_background",
    "set_master_placeholder", "set_media_playback", "set_notes",
    "set_placeholder_text", "set_reading_order", "set_row_heights",
    "set_shape", "set_show_properties", "set_slide_background",
    "set_slide_hidden", "set_slide_size", "set_table_cells",
    "set_theme_colors", "set_theme_fonts", "set_transition"
})


def title(name: str) -> str:
    """The MCP title for one tool. Unknown names RAISE, because a tool that
    reaches tools/list without a title fails the directory's annotation
    requirement and a name-shaped fallback would hide that from the test."""
    try:
        return TITLES[name]
    except KeyError:
        raise RuntimeError(
            f"tool {name!r} has no title in this module. Every registered "
            f"tool needs one: the Anthropic directory requires a title on "
            f"every tool, and a generated fallback would pass the check "
            f"while shipping 'Com Export Pdf' to a user."
        ) from None


def destructive_hint(name: str) -> bool | None:
    """The MCP destructiveHint, or None for a read-only tool, where the
    field carries no meaning. An unclassified mutating name returns None
    and `annotations` raises on it."""
    if name in NON_DESTRUCTIVE:
        return False
    if name in DESTRUCTIVE:
        return True
    return None


def idempotent_hint(name: str) -> bool | None:
    """True where repeating the call lands the same state, else None. Never
    false: an unlisted tool is unclassified, not proven non-idempotent."""
    return True if name in IDEMPOTENT else None


def open_world_hint(name: str) -> bool:
    """False for every tool: this server talks to local files and to a local
    Office installation, never to a remote service."""
    return False


def annotations(name: str, read_only: bool) -> dict:
    """The full annotation dict for one tool, ready for registration.
    destructiveHint is omitted on read-only tools and idempotentHint is
    omitted where it was not classified, so an absent field means
    "undeclared" rather than "false"."""
    ann: dict = {
        "title": title(name),
        "readOnlyHint": read_only,
        "openWorldHint": open_world_hint(name),
    }
    if not read_only:
        destructive = destructive_hint(name)
        if destructive is None:
            raise RuntimeError(
                f"tool {name!r} is not classified DESTRUCTIVE or "
                f"NON_DESTRUCTIVE in this module. Every tool that can "
                f"change something must declare whether the change may be "
                f"destructive before it can be registered."
            )
        ann["destructiveHint"] = destructive
    idempotent = idempotent_hint(name)
    if idempotent is not None:
        ann["idempotentHint"] = idempotent
    return ann
