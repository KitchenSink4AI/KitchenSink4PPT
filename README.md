<!-- mcp-name: io.github.nometalalchemist/kitchensink4ppt -->
<!-- The line above verifies the name server.json declares today. The line below
     is the org namespace the next version bump moves to; both may sit here, because
     the registry looks for the one string that matches server.json. -->
<!-- mcp-name: io.github.KitchenSink4AI/kitchensink4ppt -->
# 🖌️ KitchenSink4PPT

[![Tests](https://github.com/KitchenSink4AI/KitchenSink4PPT/actions/workflows/tests.yml/badge.svg)](https://github.com/KitchenSink4AI/KitchenSink4PPT/actions/workflows/tests.yml)
[![PyPI](https://img.shields.io/pypi/v/kitchensink4ppt)](https://pypi.org/project/kitchensink4ppt/)
[![License: AGPL-3.0](https://img.shields.io/badge/license-AGPL--3.0-blue)](LICENSE)

[Landing page](https://kitchensink4ai.github.io/KitchenSink4PPT/) · [llms.txt](https://kitchensink4ai.github.io/KitchenSink4PPT/llms.txt) (machine-readable capability manifest for agents and LLM crawlers)

**Build and revise PowerPoint decks with your AI assistant, with diagrams made of shapes you can still edit.**

Create and revise real PowerPoint presentations from Claude Code, Codex CLI, Copilot CLI or any other MCP client that runs local tools. KitchenSink4PPT connects your assistant to .pptx files, and its diagrams are built from native shapes and connectors, so the slide is still yours to change after the AI is done. Files are processed on your computer; the only thing that leaves it is what your AI app sends to its own provider. The Community edition is free under the AGPL. The Business edition adds a Windows installer, a signed update channel, a license your company can approve and support.

**Works on:** Windows, macOS and Linux for the file tools. Live PowerPoint features need Windows with PowerPoint. PDF and image export need PowerPoint on Windows or a supported LibreOffice.

## Install

Pick the route for your AI app. The commands go in PowerShell on Windows or a terminal on macOS and Linux, not into an AI chat. The package routes need Python 3.12 or newer.

**Claude Desktop**

Install [uv](https://docs.astral.sh/uv/getting-started/installation/), then quit and reopen Claude Desktop. Download the `.mcpb` file from [KitchenSink4PPT releases](https://github.com/KitchenSink4AI/KitchenSink4PPT/releases/latest). In Claude Desktop open Settings, then Extensions, then Advanced settings, then Install extension, and choose the file. The bundle fetches the Python package the first time it starts, so the first launch needs a network connection. Restart your session and check that the tools show as connected.

**Claude Code or Codex CLI**

Install uv, then run the line for your app and restart your session:

```sh
claude mcp add ppt -s user -- uvx kitchensink4ppt
```

```sh
codex mcp add ppt -- uvx kitchensink4ppt
```

**Any other local MCP client**

Use `uvx` as the command and `kitchensink4ppt` as its argument, or install the package and use `kitchensink4ppt` as the server command:

```sh
pip install kitchensink4ppt
```

Then follow your client's guide for adding a local MCP server. Installing the package on its own does not connect it to an AI app.

**Business edition**

Compare the editions on the [pricing page](https://kitchensink4.ai/pricing/). Already purchased? Your Windows installer and download link are in your [license portal](https://get.kitchensink4.ai/my-license/).

## What it can do

142 tools with every pack enabled; 25 in the lite surface.

- Build a deck from a template and revise its text and speaker notes.
- Draw diagrams from editable shapes and connectors.
- Turn supported SVG graphics into native PowerPoint shapes.
- Align objects and apply consistent fonts, colors and layouts.
- Build tables and charts that colleagues can keep editing.
- Merge, split or compare decks and review comments.
- Check slides for accessibility issues and apply the supported repairs.
- Export PDF or slide images when a rendering application is installed.

What is available depends on the packs you enable and the applications installed. The full tool reference is below.

## Business edition

Need a license your company can approve and a setup someone supports? The Business edition pairs these tools with a Windows installer, a signed update channel and support under the Business terms. Update checks tell you when a covered release is available; nothing installs on its own. Compare the options on the [pricing page](https://kitchensink4.ai/pricing/). The Community edition stays free under the AGPL, including business use that meets its terms.

## Privacy Policy

The tools run on your computer, and KitchenSink4AI receives no documents and no usage data from them. Your AI app may send prompts, file contents and tool results to its own provider under that app's settings and terms. Installing downloads packages, and the Community version check contacts PyPI unless you disable it; those requests carry connection details such as your network address and never your documents. Cloud folders and backups follow their own settings. The [Privacy Policy](https://kitchensink4.ai/privacy/) covers the product, purchases and support records.

Not affiliated with, endorsed by, or sponsored by Microsoft Corporation.
Microsoft and PowerPoint are trademarks of the Microsoft group of companies.
Dual-licensed: AGPL-3.0, or a commercial license for organizations that need
to ship it in closed products ([details](#license)).

## The headline: real graphics, not pictures of graphics

Feed `svg_to_shapes` any SVG and it compiles into a grouped tree of NATIVE
PowerPoint shapes: custom geometry, gradients, strokes, text, nested groups.
No rasterizing, no external APIs, no PowerPoint install needed to build.
Every shape keeps its own id, so afterward you edit the diagram
semantically: recolor node 7, resize the box, rewrite its label.

Connectors are GLUED. Move the node and the arrows follow, exactly as if a
person had drawn the diagram by hand in PowerPoint. The human you hand the
deck to can nudge any piece in seconds, with no regeneration round trip.

`insert_shape` (presets and freeform geometry), `insert_connector`
(straight, elbow, curved, arrowheads), `group_shapes`, `align_shapes`,
`distribute_shapes`, and `set_z_order` cover from-scratch diagram building;
`export_slide_image` renders a PNG so the agent can look at what it made and
fix it.

## Pack inventory (142 tools total)

| Pack | Tools | ~Tokens | What is in it |
|---|---|---|---|
| lite core (always on) | 25 | 7.1k | anchored deck view, atomic batch edits, get/find/replace text (live-aware, SmartArt text included), slide insert/delete/duplicate/reorder, placeholder text, hyperlinks (set/remove/list with broken-link detection), info and enumeration, copy, snapshots, backups, diagnose, workflows, enable/disable_tools |
| graphics | 27 | 8.8k | shapes, glued connectors, SVG compiler, one-call diagram generators (timeline, org chart, matrix, cycle, comparison), images, video/audio embed, groups, align/distribute, z-order, text boxes, run formatting, bullets, format painter (copy_format/copy_position), native LaTeX equations |
| tables-charts | 19 | 5.3k | create table, bulk cells, merge/unmerge, row and column insert/delete, borders and fills, widths/heights, 74 built-in styles, CSV/JSON export/import, bar/line/pie/scatter/combo charts with editable data workbooks, chart formatting and data readback |
| design | 25 | 6.8k | create presentation FROM template, apply layouts, theme read AND write (colors, fonts), brand extract/apply, layout guardrail checks, slide size, hide/move slide, autofit report, slide and master/layout backgrounds, full master and layout editing (placeholders, decoration shapes, create_layout), accessibility audit and repair |
| assembly-export | 28 | 7.2k | speaker notes, sections, footers and slide numbers, PDF/PNG/handout export, engine detection, opens-clean validation, text extraction, cross-deck slide copy, deck merge and split, agenda slides, deck statistics, document properties, anonymize, slide-show setup and custom shows, slide transitions (fade/push/wipe/split/cut/random, millisecond durations, auto-advance) and bounded entrance animations (appear/fade/wipe, click builds, by-paragraph) |
| review-sweeps | 13 | 3.3k | modern threaded comments (add, replies, resolve, cascade delete, dual-system listing), whole-deck review report, structural deck-to-deck diff (compare_decks), and the deck-wide sweeps: font inventory/replace (incl. charts and phantom declarations), color remap and literal-to-theme unification, proofing language, whole-deck logo replace, compress/purge |
| com (Windows only) | 5 | 0.9k | PowerPoint status and zombie process check, plus editing the deck while it is OPEN in the user's PowerPoint: explicit save, scroll-to-slide, session status; eleven file tools route here automatically via `live='auto'` |

Full surface: about 39.4k tokens if you pin `KS4P_MODE=full` (numbers from
`scripts/measure_surface.py`, not hand-math).

v1.1 consolidated nine packs into six. The v1.0 names
`transitions-animations`, `review`, `sweeps`, and `com-live` still resolve
to their new homes in `enable_tools`, `disable_tools`, and `KS4P_MODE`, so
nothing that worked before stops working.

Structural table operations on the file itself (merging, inserting and
deleting rows AND columns, per-edge borders) exist in no other PowerPoint
MCP server; they were previously COM-or-nothing.

## Tiered loading: start light, grow mid-session

The server starts in lite mode: 25 tools, roughly 7.1k tokens of tool
context, covering reading, slide CRUD, text, hyperlinks, batch editing,
backups, and diagnostics. The other 117 tools are registered but disabled
until asked for:

```
enable_tools(packs=["graphics"])
```

The tool list grows in place (the server emits tools/list_changed and the
client re-fetches), and the result reports the approximate token cost added
so the tradeoff is visible. `disable_tools` shrinks the surface again, and
`get_workflows` ships recipes that name the right pack for each job.

Environment pins for hosts and power users:

| Variable | Effect |
|---|---|
| `KS4P_MODE` | startup surface: `lite` (default), `full`, or a pack list like `graphics,com` |
| `KS4P_PACK_POLICY` | `auto` (default) or `locked` (enable_tools refuses; surface fixed at startup) |
| `KS4P_ALL_TOOLS` | `true` loads every pack at startup; `false` or empty keeps lite. `KS4P_MODE` wins when set |
| `KS4P_LOCK_TOOLS` | `true` fixes the surface at startup; `false` or empty leaves it adjustable. `KS4P_PACK_POLICY` wins when set |
| `KS4P_ALLOWED_ROOTS` | opt-in path sandbox; tools refuse to touch files outside these roots |
| `KS4P_UPDATE_CHECK` | `off` turns the update check off completely: no network call, no cache file (the older `KS4P_NO_UPDATE_CHECK=1` still works) |

**Update check.** The server looks for a newer release on PyPI only when you
call `diagnose`, never at startup and never on a timer, at most one request
every seven days, capped at two seconds. The check is a single plain HTTPS
GET to pypi.org that sends nothing but the request itself. A failed check is
reported with its reason rather than hidden. Set `KS4P_UPDATE_CHECK=off` to turn it off
completely (the older `KS4P_NO_UPDATE_CHECK=1` still works). The server never
downloads or installs anything.

The last two are the checkboxes the `.mcpb` bundle shows in Claude Desktop:
"Load every tool at startup" and "Lock the tool set at startup". Both take
`true` or `false`, treat an empty value as off, and refuse to start on
anything else rather than guessing. The server writes one line to stderr at
startup naming what decided the surface.

Tip: in Claude Desktop's Tool permissions, set the Read-only tools group to
Always Allow: those tools cannot change anything, and it stops most
permission prompts.

## Safety story

The same discipline as KitchenSink4Word, applied from day one:

- **Atomic validated saves.** Every mutation rebuilds the package in
  memory, validates the payload, then atomically replaces the file. A
  failed operation leaves the original byte-identical; the file is never
  absent from its own path, even for an instant.
- **Two-slot backups.** Before each mutation the current content rotates
  into `prev.pptx` and `anchor.pptx` under a hidden `.ks4p-backups/`
  folder. `manage_backups` lists, restores (undoably), and purges;
  `create_snapshot` makes DTG-stamped permanent keepers.
- **Byte-identical passthrough.** Parts the tool did not touch are written
  back byte-for-byte, so themes, media, and animations survive edits to
  other slides untouched.
- **Conservative refusals.** Ambiguous targets refuse with a candidate
  list (no first-match guessing), merged-cell surgery that would split a
  span refuses, stale batch anchors refuse the whole batch before anything
  mutates. Refusals are structured (`{ok: false, error: {code, message,
  hint}}`) and tell you the exact next call to make.
- **Sandboxing, opt-in.** Set `KS4P_ALLOWED_ROOTS` and every path in and
  out is checked.
- **Locks.** Mutations of one file are serialized in-process and across
  processes; files open in PowerPoint are refused rather than corrupted.

## Maturity

Beta. The file layer (packages, slides, text, graphics, tables, charts,
comments, animations, themes, links, media, notes, masters, equations,
accessibility, deck assembly, sweeps, export) is covered by a 1,251-test
suite, including validation that generated decks open clean in real
PowerPoint, and the server passes a raw stdio protocol round-trip suite.
Live editing of decks open in PowerPoint runs every call through one
process-wide lock with dialog detection at the window layer and bounded
timeouts, and ships with its own COM gate scripts (`tests/com_gates/`).
It has not yet had a long field life; treat important decks with the
respect the backup tools make easy, and expect fast point releases.
[Before filing](https://github.com/KitchenSink4AI/KitchenSink4PPT/issues/new?template=bug_report.yml):
ask your AI to run `diagnose` and paste the output here; it is designed
to be safe to share.

## License

KitchenSink4PPT is dual-licensed:

**AGPL-3.0 (open source).** Free for anyone (individuals, academics, and
businesses) for any use that complies with the AGPL's terms. Those terms
include sharing source, including your modifications, when you distribute
the software or make it available over a network.

**Commercial license.** For organizations that want to build KitchenSink4PPT
into their own products or services without the AGPL's source-sharing
obligations. Contact licensing@kitchensink4.ai.

Copyright (c) 2026 Alvut Consulting, LLC. KitchenSink4AI is a product line
of Alvut Consulting, LLC.

Not affiliated with, endorsed by, or sponsored by Microsoft Corporation.
Microsoft and PowerPoint are trademarks of the Microsoft group of companies.

## Family

Sibling of [KitchenSink4Word](https://github.com/KitchenSink4AI/KitchenSink4Word)
(the same engineering for .docx; [site](https://kitchensink4ai.github.io/KitchenSink4Word/)).
