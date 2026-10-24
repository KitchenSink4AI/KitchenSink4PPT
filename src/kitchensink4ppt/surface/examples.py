"""Worked argument examples for the fixed PowerPoint operation surface.

The op sweep replaces the portable fixture names and ids below with paths
and ids from its own canary deck.  Keeping the examples as data also lets
``ppt_help`` show the same call shape the release gate executes.
"""

from __future__ import annotations

from typing import Any

from ..registry import GROUP_OPS

DECK = "presentation.pptx"
OTHER_DECK = "source-presentation.pptx"
PNG = "image.png"
CSV = "table.csv"

_NO_FILE = {
    "get_server_info", "get_workflows", "get_export_engines",
    "powerpoint_status", "zombie_check", "live_status",
    "update_now", "update_notices",
}

# Every operation has an entry.  Most operations begin with the canary deck;
# the operation-specific values below replace required selectors and choices.
EXAMPLES: dict[str, dict[str, Any]] = {
    name: ({} if name in _NO_FILE else {"file_path": DECK})
    for names in GROUP_OPS.values()
    for name in names
}

EXAMPLES.update({
    "get_presentation_view": {"file_path": DECK, "detail": "text"},
    "get_text": {"file_path": DECK, "live": "off"},
    "find_text": {"file_path": DECK, "query": "Canary"},
    "search_and_replace": {"file_path": DECK, "find": "Canary", "replace": "Verified", "live": "off"},
    "apply_edits": {"file_path": DECK, "edits": [{"op": "set_text", "anchor": "s:canary", "text": "Verified"}]},
    "insert_slide": {"file_path": DECK, "layout": 1, "position": 1, "live": "off"},
    "delete_slide": {"file_path": DECK, "slide": 2, "live": "off"},
    "duplicate_slide": {"file_path": DECK, "slide": 0},
    "reorder_slides": {"file_path": DECK, "order": [2, 1, 0]},
    "move_slide": {"file_path": DECK, "slide": 2, "to": 0},
    "set_slide_hidden": {"file_path": DECK, "slide": 1, "hidden": True},
    "manage_section": {"file_path": DECK, "action": "create", "name": "Canary section", "slide": 1},
    "set_placeholder_text": {"file_path": DECK, "slide": 0, "placeholder": 0, "text": "Canary title", "live": "off"},
    "fit_text": {"file_path": DECK, "slide": 0},
    "format_text": {"file_path": DECK, "slide": 0, "shape": 1, "bold": True, "live": "off"},
    "set_bullets": {"file_path": DECK, "slide": 0, "shape": 2, "style": "char", "char": "•"},
    "insert_textbox": {"file_path": DECK, "slide": 0, "text": "Canary text", "x": 1.0, "y": 1.0, "w": 2.0, "h": 1.0, "live": "off"},
    "set_notes": {"file_path": DECK, "slide": 0, "text": "Canary notes", "live": "off"},
    "delete_notes": {"file_path": DECK, "slide": 0},
    "set_hyperlink": {"file_path": DECK, "slide": 0, "target": 1, "url": "https://example.com"},
    "remove_hyperlink": {"file_path": DECK, "slide": 0, "target": 1},
    "get_slide_info": {"file_path": DECK, "slide": 0, "live": "off"},
    "list_elements": {"file_path": DECK, "kind": "shapes"},
    "get_table": {"file_path": DECK, "slide": 1, "table": {"index": 1}},
    "get_chart_data": {"file_path": DECK, "slide": 2, "chart": {"index": 0}},
    "get_notes": {"file_path": DECK, "slide": 0},
    "get_media_playback": {"file_path": DECK, "slide": 0, "shape": 1},
    "get_footer_support": {"file_path": DECK, "slide": 0},
    "get_autofit_state": {"file_path": DECK, "slide": 0},
    "compare_decks": {"file_path_a": DECK, "file_path_b": OTHER_DECK},
    "diagnose": {"file_path": DECK},
    "insert_shape": {"file_path": DECK, "slide": 0, "shape_type": "rect", "x": 1.0, "y": 1.0, "w": 2.0, "h": 1.0, "text": "Canary", "live": "off"},
    "insert_connector": {"file_path": DECK, "slide": 0, "kind": "straight", "start": [1.0, 1.0], "end": [3.0, 2.0]},
    "set_shape": {"file_path": DECK, "slide": 0, "shape": 1, "dx": 0.1, "live": "off"},
    "set_shape_adjustments": {"file_path": DECK, "slide": 0, "shape": 1},
    "delete_shape": {"file_path": DECK, "slide": 0, "shape": 1},
    "group_shapes": {"file_path": DECK, "slide": 0, "ids": [1, 2], "name": "Canary group"},
    "ungroup_shapes": {"file_path": DECK, "slide": 0, "group": 1},
    "align_shapes": {"file_path": DECK, "slide": 0, "ids": [1, 2], "mode": "left"},
    "distribute_shapes": {"file_path": DECK, "slide": 0, "ids": [1, 2, 3], "axis": "h"},
    "set_z_order": {"file_path": DECK, "slide": 0, "shape": 1, "action": "front"},
    "svg_to_shapes": {"file_path": DECK, "slide": 0, "svg": "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><rect width='10' height='10'/></svg>", "x": 1.0, "y": 1.0, "w": 1.0, "h": 1.0},
    "set_diagram_text": {"file_path": DECK, "slide": 0, "shape": 1, "nodes": [{"index": 0, "text": "Canary"}]},
    "generate_diagram": {"file_path": DECK, "slide": 0, "kind": "timeline", "spec": {"milestones": [{"label": "Start"}, {"label": "Finish"}]}, "x": 1.0, "y": 1.0, "w": 5.0, "h": 2.0},
    "copy_format": {"file_path": DECK, "slide": 0, "from_shape": 1, "to_shapes": [2]},
    "copy_position": {"file_path": DECK, "from_slide": 0, "shape": 1, "to_slides": [1]},
    "copy_shape_between": {"file_path": DECK, "source_path": DECK, "source_slide": 0, "shape": 1, "slide": 0, "x": 1.0, "y": 1.0},
    "insert_image": {"file_path": DECK, "slide": 0, "image": PNG, "x": 1.0, "y": 1.0, "w": 1.0},
    "replace_image": {"file_path": DECK, "slide": 3, "shape": 1, "image": PNG},
    "set_image": {"file_path": DECK, "slide": 3, "shape": 1, "alt_text": "Canary image"},
    "replace_image_everywhere": {"file_path": DECK, "old_image": PNG, "new_image": PNG},
    "insert_video": {"file_path": DECK, "slide": 0, "video": "video.mp4", "x": 1.0, "y": 1.0, "w": 2.0, "h": 1.0},
    "insert_audio": {"file_path": DECK, "slide": 0, "audio": "audio.wav", "x": 1.0, "y": 1.0},
    "set_media_playback": {"file_path": DECK, "slide": 0, "shape": 1, "start": "click"},
    "insert_equation": {"file_path": DECK, "slide": 0, "latex": "x^2", "x": 1.0, "y": 1.0},
    "add_equation_to_shape": {"file_path": DECK, "slide": 0, "shape": 2, "latex": "x^2"},
    "create_table": {"file_path": DECK, "slide": 0, "rows": 2, "cols": 2, "x": 1.0, "y": 2.0, "w": 4.0, "h": 1.5, "data": [["Name", "Qty"], ["Canary", 1]]},
    "set_table_cells": {"file_path": DECK, "slide": 1, "table": {"index": 1}, "cells": [{"row": 0, "col": 0, "text": "Canary"}]},
    "merge_cells": {"file_path": DECK, "slide": 1, "table": {"index": 1}, "r1": 0, "c1": 0, "r2": 0, "c2": 1},
    "unmerge_cells": {"file_path": DECK, "slide": 1, "table": {"index": 1}, "row": 0, "col": 0},
    "insert_table_rows": {"file_path": DECK, "slide": 1, "table": {"index": 1}, "at": 1},
    "delete_table_rows": {"file_path": DECK, "slide": 1, "table": {"index": 1}, "at": 1},
    "insert_table_cols": {"file_path": DECK, "slide": 1, "table": {"index": 1}, "at": 1},
    "delete_table_cols": {"file_path": DECK, "slide": 1, "table": {"index": 1}, "at": 1},
    "format_table_cells": {"file_path": DECK, "slide": 1, "table": {"index": 1}, "range": {"r1": 0, "c1": 0, "r2": 1, "c2": 1}, "fill": "#D9EAF7"},
    "set_column_widths": {"file_path": DECK, "slide": 1, "table": {"index": 1}, "widths": [2.0, 2.0]},
    "set_row_heights": {"file_path": DECK, "slide": 1, "table": {"index": 1}, "heights": [0.5, 0.5]},
    "apply_table_style": {"file_path": DECK, "slide": 1, "table": {"index": 1}, "first_row": True},
    "export_table": {"file_path": DECK, "slide": 1, "table": {"index": 1}, "path": "table-export.csv", "format": "csv"},
    "import_table": {"file_path": DECK, "slide": 0, "source": CSV, "x": 1.0, "y": 2.0},
    "create_chart": {"file_path": DECK, "slide": 0, "chart_type": "bar", "series": [{"name": "Qty", "values": [1, 2]}], "categories": ["A", "B"], "x": 1.0, "y": 2.0, "w": 5.0, "h": 3.0},
    "update_chart_data": {"file_path": DECK, "slide": 2, "chart": {"index": 0}, "series": [{"name": "Qty", "values": [2, 3]}], "categories": ["A", "B"]},
    "format_chart": {"file_path": DECK, "slide": 2, "chart": {"index": 0}, "title": "Canary chart"},
    "create_presentation": {"path": "new-presentation.pptx"},
    "copy_presentation": {"file_path": DECK, "dest_path": "presentation-copy.pptx"},
    "create_snapshot": {"file_path": DECK, "label": "canary", "dest_dir": "snapshots"},
    "manage_backups": {"action": "list", "file_path": DECK},
    "copy_slide_between": {"file_path": DECK, "source_path": OTHER_DECK, "slide": 0},
    "merge_decks": {"file_path": DECK, "sources": [OTHER_DECK]},
    "split_deck": {"file_path": DECK, "output_dir": "split", "by": "ranges", "ranges": [{"name": "first", "start": 0, "end": 1}]},
    "export_pdf": {"file_path": DECK, "output": "presentation.pdf", "engine": "com"},
    "export_slide_image": {"file_path": DECK, "output_dir": "slide-images", "slides": [0], "engine": "com"},
    "export_shape_image": {"file_path": DECK, "slide": 0, "shape": 1, "output": "shape.png", "width": 640},
    "export_handout": {"file_path": DECK, "output": "handout.pdf", "slides_per_page": 3},
    "set_document_properties": {"file_path": DECK, "properties": {"title": "Canary deck"}},
    "generate_agenda_slide": {"file_path": DECK, "position": 1, "title": "Agenda"},
    "set_show_properties": {"file_path": DECK, "show_type": "present", "loop": False},
    "manage_custom_show": {"file_path": DECK, "action": "create", "name": "Canary show", "slides": [0, 1]},
    "set_slide_size": {"file_path": DECK, "preset": "16:9"},
    "apply_layout": {"file_path": DECK, "slide": 0, "layout": 1},
    "set_theme_colors": {"file_path": DECK, "colors": {"accent1": "#4472C4"}},
    "set_theme_fonts": {"file_path": DECK, "major": "Arial", "minor": "Arial"},
    "apply_brand": {"file_path": DECK, "brand": {"colors": {"accent1": "#4472C4"}, "fonts": {"major_latin": "Arial", "minor_latin": "Arial"}}},
    "set_slide_background": {"file_path": DECK, "slide": 0, "fill": "#FFFFFF"},
    "set_master_placeholder": {"file_path": DECK, "ph": "title", "size": 24.0},
    "set_layout_placeholder": {"file_path": DECK, "layout": 1, "ph": "title", "size": 24.0},
    "add_layout_placeholder": {"file_path": DECK, "layout": 1, "ph_type": "body", "x": 1.0, "y": 2.0, "w": 5.0, "h": 2.0},
    "remove_layout_placeholder": {"file_path": DECK, "layout": 1, "ph": "body", "force": False},
    "insert_master_shape": {"file_path": DECK, "shape_type": "rect", "x": 0.0, "y": 0.0, "w": 1.0, "h": 0.2, "text": "Canary"},
    "delete_master_shape": {"file_path": DECK, "shape_id": 1},
    "create_layout": {"file_path": DECK, "name": "Canary layout", "based_on": 1},
    "set_master_background": {"file_path": DECK, "fill": "#FFFFFF"},
    "set_alt_text": {"file_path": DECK, "slide": 0, "shape": 1, "text": "Canary title"},
    "set_reading_order": {"file_path": DECK, "slide": 0, "order": [1, 2]},
    "set_footer": {"file_path": DECK, "footer": "Canary footer", "slide_number": True},
    "set_transition": {"file_path": DECK, "kind": "fade", "slide": 0},
    "add_entrance_animation": {"file_path": DECK, "slide": 0, "shape": 1, "effect": "fade"},
    "list_animations": {"file_path": DECK, "slide": 0},
    "clear_animations": {"file_path": DECK, "slide": 0},
    "add_comment": {"file_path": DECK, "slide": 0, "text": "Canary review"},
    "reply_to_comment": {"file_path": DECK, "slide": 0, "comment_id": "00000000-0000-0000-0000-000000000001", "text": "Canary reply"},
    "resolve_comment": {"file_path": DECK, "slide": 0, "comment_id": "00000000-0000-0000-0000-000000000001"},
    "delete_comment": {"file_path": DECK, "slide": 0, "comment_id": "00000000-0000-0000-0000-000000000001"},
    "update_now": {},
    "update_notices": {"action": "status"},
    "replace_fonts": {"file_path": DECK, "mapping": {"Arial": "Calibri"}},
    "replace_colors": {"file_path": DECK, "mapping": {"#000000": "#111111"}},
    "set_language": {"file_path": DECK, "lang": "en-US"},
})

__all__ = ["EXAMPLES"]
