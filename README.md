<!-- mcp-name: io.github.KitchenSink4AI/kitchensink4ppt -->

# <img src="https://raw.githubusercontent.com/KitchenSink4AI/KitchenSink4PPT/main/bundle/icon.png" alt="" width="40" height="40"> KitchenSink4PPT

[![Tests](https://github.com/KitchenSink4AI/KitchenSink4PPT/actions/workflows/tests.yml/badge.svg)](https://github.com/KitchenSink4AI/KitchenSink4PPT/actions/workflows/tests.yml)
[![PyPI](https://img.shields.io/pypi/v/kitchensink4ppt)](https://pypi.org/project/kitchensink4ppt/)
[![License: AGPL-3.0](https://img.shields.io/badge/license-AGPL--3.0-blue)](https://github.com/KitchenSink4AI/KitchenSink4PPT/blob/main/LICENSE)

**Build and revise PowerPoint decks with your AI assistant, with diagrams made of shapes you can still edit.**

Your AI gets 193 actions in PowerPoint, a backup before every change, and a question instead of a guess when your request could mean more than one shape.

KitchenSink4PPT is an MCP server that lets Claude Code, the Codex CLI, Copilot CLI, Claude Desktop and other local MCP clients create and edit .pptx files. Its diagrams are built from native PowerPoint shapes and glued connectors, so the slide is still yours to adjust after the AI is done. This is the free Community edition, licensed under the AGPL-3.0.

[Product page](https://community.kitchensink4.ai/ppt/) · [llms.txt](https://community.kitchensink4.ai/ppt/llms.txt) · [Changelog](https://github.com/KitchenSink4AI/KitchenSink4PPT/blob/main/CHANGELOG.md)

**Platform:** the file tools run on Windows and Linux (macOS is untested in this release). Editing a deck that is open in PowerPoint needs Windows with PowerPoint. PDF and image export need PowerPoint on Windows or a supported LibreOffice.

## Install

The package needs Python 3.12 or newer. Run the commands below in PowerShell or a terminal, not in an AI chat.

**Claude Code or Codex CLI.** Install [uv](https://docs.astral.sh/uv/getting-started/installation/), run the line for your app, and start a new session:

```sh
claude mcp add ppt -s user -- uvx kitchensink4ppt
```

```sh
codex mcp add ppt -- uvx kitchensink4ppt
```

**Claude Desktop.** Install uv and restart Claude Desktop. Download the `.mcpb` file from the [latest release](https://github.com/KitchenSink4AI/KitchenSink4PPT/releases/latest), then open Settings > Extensions > Advanced settings > Install extension and choose the file. The first launch downloads the Python package, so it needs a network connection.

**Any other local MCP client.** Use `uvx` as the command and `kitchensink4ppt` as its argument, or install the package and run `kitchensink4ppt` as the server command:

```sh
pip install kitchensink4ppt
```

Optional: `pip install "kitchensink4ppt[metrics]"` adds measurement with your installed fonts for text fitting and overflow checks. The results are still estimates, so render the slide to confirm.

## What it can do

- Build a deck from a template and revise its text and speaker notes.
- Draw diagrams from editable shapes and glued connectors, or generate timelines, org charts, matrices, cycles and comparisons in one call.
- Turn SVG graphics into native PowerPoint shapes.
- Align objects and apply consistent fonts, colors, themes and layouts.
- Build tables and charts that colleagues can keep editing.
- Merge, split and compare decks, and work through review comments.
- Check slides for accessibility issues and apply the supported repairs.
- Export PDF or slide images when a rendering application is installed.

In total the server offers 145 operations in 9 tools, 193 counting each action an operation can take (for example each kind of transition or alignment). The whole tool list costs about 7.3k tokens of your AI's working memory at startup. Figures come from `scripts/measure_surface.py` and `scripts/count_operations.py`.

[KitchenSink4Word](https://github.com/KitchenSink4AI/KitchenSink4Word) (Word), [KitchenSink4XL](https://github.com/KitchenSink4AI/KitchenSink4XL) (Excel) and [KitchenSink4Web](https://github.com/KitchenSink4AI/KitchenSink4Web) (web pages) are built the same way.

## Real graphics, not pictures of graphics

`svg_to_shapes` compiles an SVG into a grouped tree of native PowerPoint shapes: custom geometry, gradients, strokes, text and nested groups. Nothing is rasterized, no external service is called, and building does not need PowerPoint installed. Every shape keeps its own id, so you can recolor one node, resize a box or rewrite a label afterward.

Connectors are glued. Move a node and its arrows follow, exactly as if someone had drawn the diagram by hand, so the person you hand the deck to can adjust any piece without regenerating it.

`insert_shape` (presets and freeform geometry), `insert_connector` (straight, elbow, curved, with arrowheads), `group_shapes`, `align_shapes`, `distribute_shapes` and `set_z_order` cover diagrams built from scratch, and `export_slide_image` renders a PNG so your AI can look at what it made and fix it.

## Tool packs and context cost

The server lists nine tools, all from the first message. Every tool except `ppt_help` takes an `op` (the operation to run) and `args` (its arguments, including the path to the .pptx). Everything is available from the start; nothing needs switching on. Token figures count the full `tools/list` entry the client receives and come from `scripts/measure_surface.py`. `ppt_help` returns the arguments and a working example for any operation, and unknown arguments are refused rather than ignored. Name a slide with `slide_number`, which counts from 1 the way people do, or with `slide`, which counts from 0; every answer about one slide gives its number and title.

| Tool | Operations | Approx. tokens | What it covers |
|---|---:|---:|---|
| `ppt_read` | 18 | ~1.0k | Read a deck as text with a stable address on every slide, find text, list elements, read tables, charts, notes, comments, equations and hyperlinks, `get_server_info` and `get_workflows` |
| `ppt_inspect` | 17 | ~0.9k | Audits that change nothing: accessibility, layout and overflow checks, font inventory, deck comparison, statistics, brand extraction, theme, document properties, `diagnose`, and PowerPoint status |
| `ppt_edit` | 19 | ~0.9k | Batch edits with `apply_edits`, search and replace, placeholder text, text fitting and formatting, bullets, hyperlinks, notes, and slide insert, delete, duplicate, reorder, move and hide |
| `ppt_graphics` | 27 | ~1.2k | Shapes and their adjust handles, glued connectors, the SVG compiler, one-call diagram generators, grouping, alignment, z-order, images, video and audio, equations, alt text and reading order, and copying a shape or group from another deck |
| `ppt_tables` | 16 | ~0.7k | Tables with cell edits, merge and unmerge, row and column insert and delete, borders and fills, widths and heights, built-in styles, CSV and JSON import; charts with editable data |
| `ppt_design` | 19 | ~0.9k | Slide size, layouts, theme colors and fonts, brand application, backgrounds, master and layout editing, font and color replacement across the deck, transitions and entrance animations |
| `ppt_review` | 7 | ~0.6k | Threaded comments (add, reply, resolve, delete) and anonymizing a deck for review. Also updates the server when you ask (`update_now`) and turns update notices on or off (`update_notices`) |
| `ppt_deck` | 22 | ~1.0k | New presentations from templates, copies, snapshots and backups, copying slides between decks, merge and split, PDF, PNG, shape image and handout export, agenda slides, footers, compression, validation, and saving or scrolling a deck open in PowerPoint |
| `ppt_help` | n/a | ~0.2k | The arguments and an example for any operation |
| **All nine** | **145** | **~7.3k** | Loaded once at the start; nothing to switch on later |

For a recipe, ask your AI to call `get_workflows`. It includes the build-a-diagram, template and table sequences with the right tool for each step. The server's own instructions add about 500 tokens to the figures above. Administrators who want a narrower server can set `KS4P_READ_ONLY=1` or list what to remove in `KS4P_FORBID` (see Configuration); what is removed no longer counts toward the cost.

## Safety model

- **Atomic, validated saves.** Every change rebuilds the package in memory, validates it, then swaps the new file in for the old one. A failed operation leaves the original exactly as it was. On Windows the deck keeps its Mark of the Web, so a downloaded deck still opens in Protected View, and its creation time; it keeps its own permissions wherever the file system and the folder allow it; where they do not, the deck takes its folder's permissions and the answer says so in a warning. Other programs may see the deck missing for a moment during a save, usually a few milliseconds; this server's reads wait it out.
- **Backups before every change.** The current version rotates into `prev.pptx` and `anchor.pptx` in a hidden `.ks4p-backups/` folder next to the deck. `create_snapshot` keeps a permanent copy in the same hidden folder, where nothing rotates or purges it; with `dest_dir` it writes a visible copy there instead, named after the deck so it never carries a newer leading date stamp than the deck. `manage_backups` lists and restores them all (a restore can itself be undone) and purges the rotating ones.
- **Untouched parts stay untouched.** Parts a call does not change are written back byte for byte, so themes, media and animations survive edits to other slides.
- **Refusals instead of guesses.** Ambiguous targets are refused with a list of candidates, merged-cell edits that would split a span are refused, and a batch with a stale anchor is refused before anything changes. Each refusal is structured (`{ok: false, error: {code, message, hint}}`) and names the next call to make.
- **Locks.** Changes to one file are serialized within a server and across processes. A deck open in PowerPoint is read from its saved file, and the answer says so; the file tools refuse to change it rather than risk it.
- **PowerPoint stays yours.** The server ends a PowerPoint process only if it started that process itself, and it cancels a hung call safely instead of killing the application.
- **Encrypted and read-only files.** An encrypted presentation is refused before any Office call, and a read-only file is reported as read-only.

## Privacy

The server runs on your computer. KitchenSink4AI never receives your files or usage data. Your AI app may send prompts, file contents and tool results to its own provider under that app's settings and terms. The update check described below contacts PyPI at most once a week and sends only the package name and the installed version. Choosing Update now downloads the new version from PyPI, or the new bundle from GitHub for Claude Desktop. The server's startup banner is switched off, so the framework it runs on makes no version check of its own. The [Privacy Policy](https://kitchensink4.ai/privacy/) covers the product, purchases and support.

## Configuration

| Variable | Effect |
|---|---|
| `KS4P_READ_ONLY` | `1` removes every tool that changes a file, leaving `ppt_read`, `ppt_inspect` and `ppt_help` |
| `KS4P_FORBID` | Comma-separated tool or operation names to remove, for example `export_pdf,ppt_review`; `ppt_app` is also accepted as a name, and removes everything that starts or attaches to PowerPoint |
| `KS4P_ALLOWED_ROOTS` | Optional sandbox: tools refuse to touch files outside these folders |
| `KS4P_CALL_TIMEOUT_S` | Seconds a single call may run before it is cancelled safely (default 120) |
| `KS4P_MAX_OUTPUT_CHARS` | Ceiling on the size of one reply, so a large deck is read in pages (default 52,000 characters) |
| `KS4P_UPDATE_CHECK` | `off` disables the update check completely (`KS4P_NO_UPDATE_CHECK=1` also works) |
| `KS4P_UPDATE_NOTICES` | `off` turns update notices off; `on` turns them back on after you turned them off |
| `KS4P_AUTHOR` | Who signs a comment or reply when the call names no author: `user` (the default) signs with your own name, your Office user name or, without one, your Windows sign-in name; `assistant` signs as "AI Assistant via KS4PPT", as does a call that passes `author="assistant"`, whatever the setting |
| `KS4P_COMMENT_AUTHOR` | A name that signs every comment and reply whose call names no author, in place of the `KS4P_AUTHOR` choice. An email address never signs, whether it is set here or sent as a call's `author`: the next choice signs and the answer says so |

A value the server cannot use stops startup with a plain message instead of being guessed at: an unreadable `KS4P_READ_ONLY`, `KS4P_CALL_TIMEOUT_S`, `KS4P_MAX_OUTPUT_CHARS` or `KS4P_AUTHOR`, a name in `KS4P_FORBID` that is not a tool or operation, or a retired setting that locked the tools, such as `KS4P_PACK_POLICY=locked`. A variable name the server does not know is ignored, so a misspelled one such as `KS4P_READ_ONY` has no effect. Check the spelling; `get_server_info` reports whether the server is read-only and what it forbids.

**Update check and update notices.** At most once a week, after you use a tool, the server asks PyPI in the background which version of kitchensink4ppt is newest. The request names only the package and the installed version and contains no decks, file names or identifiers. When a newer version exists and the server runs through uvx or the Claude Desktop extension, your AI app shows a short notice with Update now and Not now, where the app supports one. Otherwise your AI mentions the new version once and offers to update. It never downloads or installs a new version unless you choose Update now in its notice or ask your AI to update it. To turn notices off, ask your AI to turn off KitchenSink4PPT update notices; `KS4P_UPDATE_NOTICES=off` turns them off everywhere, and `KS4P_UPDATE_CHECK=off` turns the check off completely (the older `KS4P_NO_UPDATE_CHECK=1` also works). The setting is kept in the server's own state folder, so it survives updates.

## Upgrading from 1.x

Version 2.0 regroups the 142 flat tools of 1.3.1 into the nine tools above. Every 1.3.1 operation still exists, but you now call it as an operation of a group tool: `ppt_graphics` with `op: "svg_to_shapes"` instead of a tool named `svg_to_shapes`. Some arguments changed, so check `ppt_help` for an operation before the first call.

- [`docs/surface_map.md`](https://github.com/KitchenSink4AI/KitchenSink4PPT/blob/main/docs/surface_map.md) lists each 1.3.1 tool beside the tool and operation that replaces it.
- The `migrate-from-v1` recipe in `get_workflows` gives your AI the same map, so it can translate a saved prompt or script.
- `enable_tools` and `disable_tools` are gone, because every operation is present from the start. The launch settings `KS4P_MODE`, `KS4P_PACK_POLICY`, `KS4P_ALL_TOOLS` and `KS4P_LOCK_TOOLS` are retired. Settings from earlier versions are handled for you: an old default (`KS4P_MODE=lite`, `full` or `everything`, `KS4P_PACK_POLICY=auto`, `KS4P_ALL_TOOLS`, `KS4P_LOCK_TOOLS=false`) is ignored with a one-line notice, and a setting that locked the tools (`KS4P_PACK_POLICY=locked`, `KS4P_LOCK_TOOLS=true` or a `KS4P_MODE` list of tool groups) stops the start with a message naming what to use instead. To limit what the server offers, use `KS4P_READ_ONLY` and `KS4P_FORBID` (see Configuration).
- The install commands, the package name and the `kitchensink4ppt` command are unchanged.

The [changelog](https://github.com/KitchenSink4AI/KitchenSink4PPT/blob/main/CHANGELOG.md) lists everything else that changed.

## Testing

The file layer (packages, slides, text, graphics, tables, charts, comments, animations, themes, links, media, notes, masters, equations, accessibility, assembly, sweeps and export) is covered by a suite of 3,689 tests, including checks that generated decks open cleanly in PowerPoint, and the server passes a protocol round-trip suite over stdio. Tests that need sample decks generate synthetic ones, so the suite runs on any machine. Editing a deck that is open in PowerPoint runs every call through one process-wide lock, with dialog detection and bounded timeouts.

## Reporting problems

Please [open an issue](https://github.com/KitchenSink4AI/KitchenSink4PPT/issues/new?template=bug_report.yml) with what you asked for, what happened and what you expected. Ask your AI to run `diagnose` and paste the output; it is designed to be safe to share. Please do not attach decks that contain anything private. For security issues, follow [SECURITY.md](https://github.com/KitchenSink4AI/KitchenSink4PPT/blob/main/SECURITY.md).

## Business edition

The Business edition pairs these tools with a Windows installer, a signed update channel and support under the Business terms. Update checks tell you when a covered release is available, and nothing installs on its own. See [pricing](https://kitchensink4.ai/pricing/). The Community edition stays free under the AGPL, including business use that meets its terms.

## Contributing

Bug reports, fixes, tests and documentation are welcome. See [CONTRIBUTING.md](https://github.com/KitchenSink4AI/KitchenSink4PPT/blob/main/CONTRIBUTING.md); first-time contributors agree to the [Contributor License Agreement](https://github.com/KitchenSink4AI/KitchenSink4PPT/blob/main/CLA.md).

## License

KitchenSink4PPT is dual-licensed.

- **AGPL-3.0.** Free for anyone, including businesses, for any use that complies with the AGPL. Its terms include sharing your source code, with your modifications, when you distribute the software or make it available over a network.
- **Commercial license.** For organizations that want to build KitchenSink4PPT into their own products or services without the AGPL's source-sharing obligations. Contact licensing@kitchensink4.ai.

Copyright (c) 2026 Alvut Consulting, LLC. KitchenSink4AI is a product line of Alvut Consulting, LLC. KitchenSink4AI tools are built with the help of AI coding assistants, and every release goes through the full test suite before it ships.

Not affiliated with, endorsed by, or sponsored by Microsoft Corporation. Microsoft and PowerPoint are trademarks of the Microsoft group of companies. GitHub and Copilot are trademarks of the Microsoft group of companies. Claude is a trademark of Anthropic. Codex and OpenAI are trademarks of OpenAI. All other trademarks are the property of their respective owners.
