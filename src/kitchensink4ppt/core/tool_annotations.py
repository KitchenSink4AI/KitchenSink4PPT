"""Tool titles, declared by name.

Every registered tool ships a `title` annotation: the Anthropic Connectors
Directory requires one, and a client shows it to a human instead of the
raw tool name. `readOnlyHint` lives in core/readonly.py and is not repeated here.

The title is derived mechanically, so that a new tool gets one without
anyone inventing it: the tool name, split on underscores, Title Cased,
with known acronyms upper-cased, a leading `com_` replaced by the host
application's name, and short function words kept lowercase inside the
phrase. A handful of names too short or too generic to read as a title
take a phrase from the first clause of their own description instead;
those are the only hand-written strings in the table, and
docs/TOOL_TITLES.md marks each one.

Titles are unique within this server and none exceeds 40 characters. An
unknown name RAISES at registration, the same contract core/readonly.py
already holds this surface to.
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


def annotations(name: str, read_only: bool) -> dict:
    """The annotation dict for one tool, ready for registration."""
    return {
        "title": title(name),
        "readOnlyHint": read_only,
    }
