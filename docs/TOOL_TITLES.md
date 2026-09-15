# Tool titles

Every tool this server registers ships a `title` annotation, which is what
the Anthropic Connectors Directory requires and what a client shows in
place of the raw tool name. This table is GENERATED from
`src/kitchensink4ppt/core/tool_annotations.py` and a test in `tests/unit/test_tool_annotations.py`
fails if the two ever disagree, so the English here is the English on the
wire.

## How a title is derived

The rule is mechanical, so that a new tool gets a title without anyone
inventing one:

1. split the tool name on underscores;
2. replace a leading `com_` with the host application's name and collapse
   an immediately repeated word (`com_word_status` -> `Word Status`);
3. upper-case known acronyms (`pdf` -> `PDF`, `svg` -> `SVG`), Title Case
   the rest, and keep short function words lowercase inside the phrase
   (`add_equation_to_shape` -> `Add Equation to Shape`);
4. a name too short or too generic to read as a title takes a phrase from
   the first clause of its own description instead. Those are the only
   hand-written strings here and the Source column marks them.

Titles are unique within this server and none exceeds 40 characters. No
title carries product or marketing language.

142 tools.

| Tool | Title | Source |
|---|---|---|
| `add_comment` | Add Comment | mechanical |
| `add_entrance_animation` | Add Entrance Animation | mechanical |
| `add_equation_to_shape` | Add Equation to Shape | mechanical |
| `add_layout_placeholder` | Add Layout Placeholder | mechanical |
| `align_shapes` | Align Shapes | mechanical |
| `anonymize_deck` | Anonymize Deck | mechanical |
| `apply_brand` | Apply Brand | mechanical |
| `apply_edits` | Apply Edits | mechanical |
| `apply_layout` | Apply Layout | mechanical |
| `apply_table_style` | Apply Table Style | mechanical |
| `audit_accessibility` | Audit Accessibility | mechanical |
| `check_layout` | Check Layout | mechanical |
| `clear_animations` | Clear Animations | mechanical |
| `comment_report` | Comment Report | mechanical |
| `compare_decks` | Compare Decks | mechanical |
| `compress_deck` | Compress Deck Media | first clause of its description |
| `copy_format` | Copy Format | mechanical |
| `copy_position` | Copy Position | mechanical |
| `copy_presentation` | Copy Presentation | mechanical |
| `copy_slide_between` | Copy Slide Between Decks | first clause of its description |
| `create_chart` | Create Chart | mechanical |
| `create_layout` | Create Layout | mechanical |
| `create_presentation` | Create Presentation | mechanical |
| `create_snapshot` | Create Snapshot | mechanical |
| `create_table` | Create Table | mechanical |
| `deck_statistics` | Deck Statistics | mechanical |
| `delete_comment` | Delete Comment | mechanical |
| `delete_master_shape` | Delete Master Shape | mechanical |
| `delete_notes` | Delete Notes | mechanical |
| `delete_shape` | Delete Shape | mechanical |
| `delete_slide` | Delete Slide | mechanical |
| `delete_table_cols` | Delete Table Columns | first clause of its description |
| `delete_table_rows` | Delete Table Rows | mechanical |
| `diagnose` | Diagnose Presentation | first clause of its description |
| `disable_tools` | Disable Tools | mechanical |
| `distribute_shapes` | Distribute Shapes | mechanical |
| `duplicate_slide` | Duplicate Slide | mechanical |
| `enable_tools` | Enable Tools | mechanical |
| `export_handout` | Export Handout | mechanical |
| `export_pdf` | Export PDF | mechanical |
| `export_slide_image` | Export Slide Image | mechanical |
| `export_table` | Export Table | mechanical |
| `extract_brand` | Extract Brand | mechanical |
| `extract_text` | Extract Text | mechanical |
| `find_text` | Find Text | mechanical |
| `fit_text` | Fit Text | mechanical |
| `font_inventory` | Font Inventory | mechanical |
| `format_chart` | Format Chart | mechanical |
| `format_table_cells` | Format Table Cells | mechanical |
| `format_text` | Format Text | mechanical |
| `generate_agenda_slide` | Generate Agenda Slide | mechanical |
| `generate_diagram` | Generate Diagram | mechanical |
| `get_autofit_state` | Get Autofit State | mechanical |
| `get_chart_data` | Get Chart Data | mechanical |
| `get_document_properties` | Get Document Properties | mechanical |
| `get_export_engines` | Get Export Engines | mechanical |
| `get_footer_support` | Get Footer Support | mechanical |
| `get_media_playback` | Get Media Playback | mechanical |
| `get_notes` | Get Notes | mechanical |
| `get_presentation_info` | Get Presentation Info | mechanical |
| `get_presentation_view` | Get Presentation View | mechanical |
| `get_server_info` | Get Server Info | mechanical |
| `get_slide_info` | Get Slide Info | mechanical |
| `get_table` | Get Table | mechanical |
| `get_text` | Get Text | mechanical |
| `get_theme` | Get Theme | mechanical |
| `get_transitions` | Get Transitions | mechanical |
| `get_workflows` | Get Workflows | mechanical |
| `group_shapes` | Group Shapes | mechanical |
| `import_table` | Import Table | mechanical |
| `insert_audio` | Insert Audio | mechanical |
| `insert_connector` | Insert Connector | mechanical |
| `insert_equation` | Insert Equation | mechanical |
| `insert_image` | Insert Image | mechanical |
| `insert_master_shape` | Insert Master Shape | mechanical |
| `insert_shape` | Insert Shape | mechanical |
| `insert_slide` | Insert Slide | mechanical |
| `insert_table_cols` | Insert Table Columns | first clause of its description |
| `insert_table_rows` | Insert Table Rows | mechanical |
| `insert_textbox` | Insert Textbox | mechanical |
| `insert_video` | Insert Video | mechanical |
| `list_animations` | List Animations | mechanical |
| `list_comments` | List Comments | mechanical |
| `list_elements` | List Elements | mechanical |
| `list_equations` | List Equations | mechanical |
| `list_hyperlinks` | List Hyperlinks | mechanical |
| `list_master_elements` | List Master Elements | mechanical |
| `live_save` | Live Save | mechanical |
| `live_scroll_to` | Live Scroll to Location | first clause of its description |
| `live_status` | Live Status | mechanical |
| `manage_backups` | Manage Backups | mechanical |
| `manage_custom_show` | Manage Custom Show | mechanical |
| `manage_section` | Manage Section | mechanical |
| `merge_cells` | Merge Cells | mechanical |
| `merge_decks` | Merge Decks | mechanical |
| `move_slide` | Move Slide | mechanical |
| `powerpoint_status` | PowerPoint Status | mechanical |
| `refresh_agenda_slide` | Refresh Agenda Slide | mechanical |
| `remove_hyperlink` | Remove Hyperlink | mechanical |
| `remove_layout_placeholder` | Remove Layout Placeholder | mechanical |
| `reorder_slides` | Reorder Slides | mechanical |
| `replace_colors` | Replace Colors | mechanical |
| `replace_fonts` | Replace Fonts | mechanical |
| `replace_image` | Replace Image | mechanical |
| `replace_image_everywhere` | Replace Image Everywhere | mechanical |
| `reply_to_comment` | Reply to Comment | mechanical |
| `resolve_comment` | Resolve Comment | mechanical |
| `search_and_replace` | Search and Replace | mechanical |
| `set_alt_text` | Set Alt Text | mechanical |
| `set_bullets` | Set Bullets | mechanical |
| `set_column_widths` | Set Column Widths | mechanical |
| `set_diagram_text` | Set Diagram Text | mechanical |
| `set_document_properties` | Set Document Properties | mechanical |
| `set_footer` | Set Footer | mechanical |
| `set_hyperlink` | Set Hyperlink | mechanical |
| `set_image` | Set Image | mechanical |
| `set_language` | Set Language | mechanical |
| `set_layout_placeholder` | Set Layout Placeholder | mechanical |
| `set_master_background` | Set Master Background | mechanical |
| `set_master_placeholder` | Set Master Placeholder | mechanical |
| `set_media_playback` | Set Media Playback | mechanical |
| `set_notes` | Set Notes | mechanical |
| `set_placeholder_text` | Set Placeholder Text | mechanical |
| `set_reading_order` | Set Reading Order | mechanical |
| `set_row_heights` | Set Row Heights | mechanical |
| `set_shape` | Set Shape | mechanical |
| `set_show_properties` | Set Show Properties | mechanical |
| `set_slide_background` | Set Slide Background | mechanical |
| `set_slide_hidden` | Set Slide Hidden | mechanical |
| `set_slide_size` | Set Slide Size | mechanical |
| `set_table_cells` | Set Table Cells | mechanical |
| `set_theme_colors` | Set Theme Colors | mechanical |
| `set_theme_fonts` | Set Theme Fonts | mechanical |
| `set_transition` | Set Transition | mechanical |
| `set_z_order` | Set Z-Order | first clause of its description |
| `split_deck` | Split Deck | mechanical |
| `svg_to_shapes` | SVG to Shapes | mechanical |
| `ungroup_shapes` | Ungroup Shapes | mechanical |
| `unmerge_cells` | Unmerge Cells | mechanical |
| `update_chart_data` | Update Chart Data | mechanical |
| `validate` | Validate Presentation | first clause of its description |
| `zombie_check` | Zombie Process Check | first clause of its description |
