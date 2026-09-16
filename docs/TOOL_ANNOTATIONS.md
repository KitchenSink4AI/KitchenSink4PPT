# Tool annotations

Every tool this server registers ships four MCP annotations: `title`
(see [TOOL_TITLES.md](TOOL_TITLES.md)), `readOnlyHint`, `destructiveHint`
and `openWorldHint`, plus `idempotentHint` where it is obviously true.
The tables are GENERATED from `src/kitchensink4ppt/core/tool_annotations.py` and a test in
`tests/unit/test_tool_annotations.py` fails if they disagree, so a row
here is what goes on the wire.

## The rule for `destructiveHint`

ONE rule decides every row, and it is stated here so a reviewer can
dispute any single one of them.

**`false`** only when every code path either

* **(a)** adds new content or a new file without replacing anything that
  was already there, with creation refusing an existing target, or
* **(b)** changes no user data at all: this session's tool surface, the
  viewport, a read performed through a hidden or already-running Office
  instance, or a read taken through a temporary copy.

**`true`** otherwise. That covers everything that deletes, replaces,
overwrites, clears, reorders, moves, applies a batch of edits, saves over
the document the user has open, writes an output file it may silently
overwrite, or acts on a live page. It also covers **every tool that writes
into an existing document**, because the pre-write backup these servers
take is defeatable by the tool's own `backup=False` argument and therefore
cannot be claimed as guaranteed reversibility. And it covers everything
whose reversibility could not be PROVEN by reading the code: an unproven
claim of safety is the one thing this annotation must not make.

Read-only tools carry no `destructiveHint`. The field is meaningful only
when `readOnlyHint` is false, and a value there would be noise.

`idempotentHint` is set true only where repeating the identical call
obviously lands the same state: the pack switches, whose own docstrings
say idempotent, and the whole-value setters that take an address and a
value and carry no action selector. Everything else is left unset rather
than guessed; an absent hint means undeclared, never false.

`openWorldHint` is `false` on every tool. Nothing here reaches a remote service.

## Tools that may perform destructive updates (76)

`destructiveHint: true`.

| Tool | Why |
|---|---|
| `align_shapes` | reorders or moves existing content |
| `anonymize_deck` | its own description opens with IRREVERSIBLE |
| `apply_brand` | writes over the styling or content already in place |
| `apply_edits` | applies a batch of addressed edits whose ops include replace and delete |
| `apply_layout` | writes over the styling or content already in place |
| `apply_table_style` | writes over the styling or content already in place |
| `clear_animations` | deletes the animations already on the slide |
| `compress_deck` | downsamples raster images and purges unused parts |
| `copy_format` | the destination's existing formatting is overwritten |
| `copy_position` | the matched shapes' position, size and rotation are overwritten |
| `delete_comment` | deletes existing content |
| `delete_master_shape` | deletes existing content |
| `delete_notes` | deletes existing content |
| `delete_shape` | deletes existing content |
| `delete_slide` | deletes existing content |
| `delete_table_cols` | deletes existing content |
| `delete_table_rows` | deletes existing content |
| `distribute_shapes` | reorders or moves existing content |
| `export_handout` | writes the output path with no proven existing-file check |
| `export_pdf` | writes the output path with no proven existing-file check |
| `export_slide_image` | writes the output paths with no proven existing-file check |
| `export_table` | writes the output path with no proven existing-file check |
| `fit_text` | rewrites explicit run sizes throughout the shape |
| `format_chart` | writes a new value over the one already stored |
| `format_table_cells` | writes a new value over the one already stored |
| `format_text` | writes a new value over the one already stored |
| `group_shapes` | reorders or moves existing content |
| `import_table` | addressing an existing table refills it in place, overwriting its cells |
| `live_save` | saves over the presentation the user has open |
| `manage_backups` | action='restore' overwrites the deck and action='purge' deletes backups |
| `manage_custom_show` | carries delete or remove actions alongside its read actions |
| `manage_section` | carries delete or remove actions alongside its read actions |
| `merge_cells` | text from covered cells moves into the origin and the cells go |
| `move_slide` | reorders or moves existing content |
| `refresh_agenda_slide` | rebuilds the agenda slide in place and drops stale links |
| `remove_hyperlink` | deletes existing content |
| `remove_layout_placeholder` | deletes existing content |
| `reorder_slides` | reorders or moves existing content |
| `replace_colors` | replaces existing content |
| `replace_fonts` | replaces existing content |
| `replace_image` | replaces existing content |
| `replace_image_everywhere` | replaces every occurrence of the image deck-wide |
| `reply_to_comment` | writes into an existing document, and the pre-write backup is defeatable by the tool's own backup argument |
| `resolve_comment` | rewrites existing content in place |
| `search_and_replace` | replaces matched text throughout the document |
| `set_alt_text` | writes a new value over the one already stored |
| `set_bullets` | writes a new value over the one already stored |
| `set_column_widths` | writes a new value over the one already stored |
| `set_diagram_text` | writes a new value over the one already stored |
| `set_document_properties` | writes a new value over the one already stored |
| `set_footer` | writes a new value over the one already stored |
| `set_hyperlink` | writes a new value over the one already stored |
| `set_image` | writes a new value over the one already stored |
| `set_language` | writes a new value over the one already stored |
| `set_layout_placeholder` | writes a new value over the one already stored |
| `set_master_background` | writes a new value over the one already stored |
| `set_master_placeholder` | writes a new value over the one already stored |
| `set_media_playback` | writes a new value over the one already stored |
| `set_notes` | writes a new value over the one already stored |
| `set_placeholder_text` | writes a new value over the one already stored |
| `set_reading_order` | rewrites the slide's whole top-level shape order |
| `set_row_heights` | writes a new value over the one already stored |
| `set_shape` | writes a new value over the one already stored |
| `set_show_properties` | writes a new value over the one already stored |
| `set_slide_background` | writes a new value over the one already stored |
| `set_slide_hidden` | writes a new value over the one already stored |
| `set_slide_size` | writes a new value over the one already stored |
| `set_table_cells` | writes a new value over the one already stored |
| `set_theme_colors` | writes a new value over the one already stored |
| `set_theme_fonts` | writes a new value over the one already stored |
| `set_transition` | writes a new value over the one already stored |
| `set_z_order` | writes a new value over the one already stored |
| `ungroup_shapes` | reorders or moves existing content |
| `unmerge_cells` | deletes existing content |
| `update_chart_data` | rewrites existing content in place |
| `validate` | writes into an existing document, and the pre-write backup is defeatable by the tool's own backup argument |
## Tools that may not (31)

`destructiveHint: false`.

| Tool | Why |
|---|---|
| `add_comment` | adds new content at a position; nothing existing is replaced or removed |
| `add_entrance_animation` | adds new content at a position; nothing existing is replaced or removed |
| `add_equation_to_shape` | adds new content at a position; nothing existing is replaced or removed |
| `add_layout_placeholder` | adds new content at a position; nothing existing is replaced or removed |
| `copy_presentation` | writes to a separate destination; an overwrite rotates the destination into its backup slot first |
| `copy_slide_between` | writes to a separate destination; an overwrite rotates the destination into its backup slot first |
| `create_chart` | adds new content at a position; nothing existing is replaced or removed |
| `create_layout` | adds new content at a position; nothing existing is replaced or removed |
| `create_presentation` | creates a new object or file and refuses an existing target |
| `create_snapshot` | creates a new object or file and refuses an existing target |
| `create_table` | adds new content at a position; nothing existing is replaced or removed |
| `disable_tools` | changes only this session's tool surface, which the counterpart tool reverses |
| `duplicate_slide` | adds new content at a position; nothing existing is replaced or removed |
| `enable_tools` | changes only this session's tool surface, which the counterpart tool reverses |
| `generate_agenda_slide` | adds new content at a position; nothing existing is replaced or removed |
| `generate_diagram` | adds new content at a position; nothing existing is replaced or removed |
| `insert_audio` | adds new content at a position; nothing existing is replaced or removed |
| `insert_connector` | adds new content at a position; nothing existing is replaced or removed |
| `insert_equation` | adds new content at a position; nothing existing is replaced or removed |
| `insert_image` | adds new content at a position; nothing existing is replaced or removed |
| `insert_master_shape` | adds new content at a position; nothing existing is replaced or removed |
| `insert_shape` | adds new content at a position; nothing existing is replaced or removed |
| `insert_slide` | adds new content at a position; nothing existing is replaced or removed |
| `insert_table_cols` | adds new content at a position; nothing existing is replaced or removed |
| `insert_table_rows` | adds new content at a position; nothing existing is replaced or removed |
| `insert_textbox` | adds new content at a position; nothing existing is replaced or removed |
| `insert_video` | adds new content at a position; nothing existing is replaced or removed |
| `live_scroll_to` | moves what the user is looking at; no document byte changes |
| `merge_decks` | appends to the destination; nothing already in it is replaced |
| `split_deck` | writes new files and refuses a collision; the source is never modified |
| `svg_to_shapes` | adds new content at a position; nothing existing is replaced or removed |
## Read-only tools (35)

`readOnlyHint: true`, and no `destructiveHint`: the field carries no
meaning for a tool that changes nothing. The classification itself lives
in `core/readonly.py` and is guarded by
`tests/unit/test_readonly_annotations.py`.

`audit_accessibility`, `check_layout`, `comment_report`, `compare_decks`, `deck_statistics`, `diagnose`, `extract_brand`, `extract_text`, `find_text`, `font_inventory`, `get_autofit_state`, `get_chart_data`, `get_document_properties`, `get_export_engines`, `get_footer_support`, `get_media_playback`, `get_notes`, `get_presentation_info`, `get_presentation_view`, `get_server_info`, `get_slide_info`, `get_table`, `get_text`, `get_theme`, `get_transitions`, `get_workflows`, `list_animations`, `list_comments`, `list_elements`, `list_equations`, `list_hyperlinks`, `list_master_elements`, `live_status`, `powerpoint_status`, `zombie_check`

## Idempotent tools (28)

`idempotentHint: true`.

`disable_tools`, `enable_tools`, `set_alt_text`, `set_bullets`, `set_column_widths`, `set_diagram_text`, `set_document_properties`, `set_footer`, `set_hyperlink`, `set_image`, `set_language`, `set_layout_placeholder`, `set_master_background`, `set_master_placeholder`, `set_media_playback`, `set_notes`, `set_placeholder_text`, `set_reading_order`, `set_row_heights`, `set_shape`, `set_show_properties`, `set_slide_background`, `set_slide_hidden`, `set_slide_size`, `set_table_cells`, `set_theme_colors`, `set_theme_fonts`, `set_transition`
