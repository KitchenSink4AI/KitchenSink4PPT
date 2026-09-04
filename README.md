<!-- mcp-name: io.github.KitchenSink4AI/kitchensink4ppt -->

# 🖌️ KitchenSink4PPT

[![Tests](https://github.com/KitchenSink4AI/KitchenSink4PPT/actions/workflows/tests.yml/badge.svg)](https://github.com/KitchenSink4AI/KitchenSink4PPT/actions/workflows/tests.yml)
[![PyPI](https://img.shields.io/pypi/v/kitchensink4ppt)](https://pypi.org/project/kitchensink4ppt/)
[![License: AGPL-3.0](https://img.shields.io/badge/license-AGPL--3.0-blue)](https://github.com/KitchenSink4AI/KitchenSink4PPT/blob/main/LICENSE)

**Build and revise PowerPoint decks with your AI assistant, with diagrams made of shapes you can still edit.**

KitchenSink4PPT is an MCP server that lets Claude Code, the Codex CLI, Copilot CLI, Claude Desktop and other local MCP clients create and edit .pptx files. Its diagrams are built from native PowerPoint shapes and glued connectors, so the slide is still yours to adjust after the AI is done. This is the free Community edition, licensed under the AGPL-3.0.

[Product page](https://kitchensink4.ai/KitchenSink4PPT/) · [llms.txt](https://kitchensink4.ai/KitchenSink4PPT/llms.txt) · [Changelog](https://github.com/KitchenSink4AI/KitchenSink4PPT/blob/main/CHANGELOG.md)

**Platform:** the file tools run on Windows, macOS and Linux. Editing a deck that is open in PowerPoint needs Windows with PowerPoint. PDF and image export need PowerPoint on Windows or a supported LibreOffice.

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

The server offers 142 tools with every pack enabled and 25 in the startup set.

## Real graphics, not pictures of graphics

`svg_to_shapes` compiles an SVG into a grouped tree of native PowerPoint shapes: custom geometry, gradients, strokes, text and nested groups. Nothing is rasterized, no external service is called, and building does not need PowerPoint installed. Every shape keeps its own id, so you can recolor one node, resize a box or rewrite a label afterward.

Connectors are glued. Move a node and its arrows follow, exactly as if someone had drawn the diagram by hand, so the person you hand the deck to can adjust any piece without regenerating it.

`insert_shape` (presets and freeform geometry), `insert_connector` (straight, elbow, curved, with arrowheads), `group_shapes`, `align_shapes`, `distribute_shapes` and `set_z_order` cover diagrams built from scratch, and `export_slide_image` renders a PNG so your AI can look at what it made and fix it.

## Tool packs and context cost

The server starts with a 25-tool core of about 7.2k tokens. The other packs are available on request, and the full set is about 39.7k tokens. Figures come from `scripts/measure_surface.py`.

| Pack | Tools | Approx. tokens | What it covers |
|---|---:|---:|---|
| **lite** (loaded at start) | 25 | 7.2k | Deck view, atomic batch edits, text find and replace (including SmartArt and decks open in PowerPoint), slide insert, delete, duplicate and reorder, placeholder text, hyperlinks with broken-link detection, copy, snapshots, backups, diagnostics, workflows, pack controls |
| graphics | 27 | 8.8k | Shapes, glued connectors, the SVG compiler, one-call diagram generators, images, video and audio, groups, alignment, z-order, text boxes, run formatting, bullets, format painter, native LaTeX equations |
| tables-charts | 19 | 5.4k | Tables with bulk cell edits, merge and unmerge, row and column insert and delete, borders and fills, 74 built-in styles, CSV and JSON import and export; bar, line, pie, scatter and combo charts with editable data, chart formatting and data readback |
| design | 25 | 6.8k | New presentations from templates, layouts, theme colors and fonts, brand extraction and application, layout checks, slide size, hidden and moved slides, backgrounds, master and layout editing, accessibility audit and repair |
| assembly-export | 28 | 7.3k | Speaker notes, sections, footers and slide numbers, PDF, PNG and handout export, opens-clean validation, text extraction, copying slides between decks, merge and split, agenda slides, deck statistics, document properties, slide shows, transitions and entrance animations |
| review-sweeps | 13 | 3.3k | Threaded comments, deck review reports, deck-to-deck comparison, and deck-wide sweeps: font inventory and replacement, color remapping to theme colors, proofing language, logo replacement, compression |
| com (Windows only) | 5 | 0.9k | PowerPoint status and editing the deck while it is open in PowerPoint: save, scroll to a slide, session status. Eleven file tools route to the open deck automatically with `live='auto'` |

Your AI turns a pack on by name:

```
enable_tools(packs=["graphics"])
```

The tool list grows in place, and the result reports the approximate token cost added. `disable_tools` shrinks it again, and `get_workflows` recipes name the right pack for each job. Pack names from 1.0 (`transitions-animations`, `review`, `sweeps`, `com-live`) still resolve to their current homes.

**Fewer permission prompts:** in Claude Desktop's tool permissions for this server, set the read-only tools to Always Allow. They cannot change anything.

## Safety model

- **Atomic, validated saves.** Every change rebuilds the package in memory, validates it, then replaces the file atomically. A failed operation leaves the original exactly as it was, and the file is never missing from its path.
- **Backups before every change.** The current version rotates into `prev.pptx` and `anchor.pptx` in a hidden `.ks4p-backups/` folder next to the deck. `manage_backups` lists, restores (a restore can itself be undone) and purges them, and `create_snapshot` keeps permanent time-stamped copies.
- **Untouched parts stay untouched.** Parts a call does not change are written back byte for byte, so themes, media and animations survive edits to other slides.
- **Refusals instead of guesses.** Ambiguous targets are refused with a list of candidates, merged-cell edits that would split a span are refused, and a batch with a stale anchor is refused before anything changes. Each refusal is structured (`{ok: false, error: {code, message, hint}}`) and names the next call to make.
- **Locks.** Changes to one file are serialized within a server and across processes, and decks open in PowerPoint are refused by the file tools rather than risked.

## Privacy

The server runs on your computer. KitchenSink4AI never receives your files or usage data. Your AI app may send prompts, file contents and tool results to its own provider under that app's settings and terms. The optional update check contacts PyPI and sends nothing beyond the request itself. The [Privacy Policy](https://kitchensink4.ai/privacy/) covers the product, purchases and support.

## Configuration

| Variable | Effect |
|---|---|
| `KS4P_MODE` | Tools at startup: `lite` (default), `full`, or a pack list such as `graphics,com` |
| `KS4P_PACK_POLICY` | `auto` (default) or `locked`, which fixes the tool set at startup |
| `KS4P_ALL_TOOLS` | `true` loads every pack at startup; `KS4P_MODE` takes precedence when set |
| `KS4P_LOCK_TOOLS` | `true` fixes the tool set at startup; `KS4P_PACK_POLICY` takes precedence when set |
| `KS4P_ALLOWED_ROOTS` | Optional sandbox: tools refuse to touch files outside these folders |
| `KS4P_UPDATE_CHECK` | `off` disables the update check completely (`KS4P_NO_UPDATE_CHECK=1` also works) |

`KS4P_ALL_TOOLS` and `KS4P_LOCK_TOOLS` are the **Load every tool at startup** and **Lock the tool set at startup** settings in the Claude Desktop extension. Both accept `true` or `false`, treat an empty value as off, and stop startup with a clear message on anything else. The server prints one line at startup saying what decided the tool set.

**Update check.** The server checks PyPI for a newer release only when you call `diagnose`: never at startup or on a timer, at most once every seven days, with a two-second limit. It never downloads or installs anything.

## Maturity and testing

KitchenSink4PPT is in beta. The file layer (packages, slides, text, graphics, tables, charts, comments, animations, themes, links, media, notes, masters, equations, accessibility, assembly, sweeps and export) is covered by a suite of about 1,600 tests, including checks that generated decks open cleanly in PowerPoint, and the server passes a protocol round-trip suite over stdio. Tests that need sample decks generate synthetic ones, so the suite runs on any machine. Editing a deck that is open in PowerPoint runs every call through one process-wide lock, with dialog detection and bounded timeouts.

It has not yet had a long life in the field, so keep the backup tools in mind for important decks and expect frequent point releases.

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

Not affiliated with, endorsed by, or sponsored by Microsoft Corporation. Microsoft and PowerPoint are trademarks of the Microsoft group of companies.

## Related

[KitchenSink4Word](https://github.com/KitchenSink4AI/KitchenSink4Word) (Word), [KitchenSink4XL](https://github.com/KitchenSink4AI/KitchenSink4XL) (Excel) and [KitchenSink4Web](https://github.com/KitchenSink4AI/KitchenSink4Web) (web pages) are built the same way.
