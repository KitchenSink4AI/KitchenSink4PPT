# KitchenSink4PPT .mcpb Bundle

This directory builds the MCP Bundle (`.mcpb`) for one-click Claude Desktop
installs and Smithery's Local (MCPB Bundle) listing. It follows the sibling
KitchenSink4Word bundle exactly; see that repo's `bundle/README-bundle.md`
for the field notes gathered during its Desktop installs.

**Design: thin uvx launcher.** The archive carries only `manifest.json` and
`icon.png`. The manifest's `mcp_config` tells the client to run
`uvx kitchensink4ppt==<version>`, so the actual server is fetched from PyPI
at first launch. Nothing from `src/` is bundled: pywin32's `.pth` bootstrap
and compiled wheels are not portable under `server/lib` bundling. End users
need uv installed (`uvx` ships with uv), which is the standard arrangement
for Python MCP servers.

## Files

| File | Role | In the archive? |
|---|---|---|
| `manifest.json` | The bundle manifest (binary/uvx shape, manifest_version 0.3) | yes |
| `icon.png` | 512x512 icon: the KitchenSink4AI app icon (ringed symbol on the dark tile), the same file in every product bundle | yes |
| `.mcpbignore` | Keeps everything but manifest + icon out of the archive | no |
| `kitchensink4ppt.mcpb` | The packed bundle | output |

## Rebuild per release

The pinned version appears in `manifest.json` **twice** and must match the
version just published to PyPI:

- `"version": "X.Y.Z"`
- `"args": ["kitchensink4ppt==X.Y.Z"]`

The pin is deliberate: uvx caches resolved environments, and an unpinned spec
can serve a stale cached version. **Every PyPI release therefore needs a
manifest bump and a repack.** After editing the two version strings:

```
mcpb pack bundle bundle/kitchensink4ppt.mcpb
.venv\Scripts\python.exe -X utf8 bundle\smoke_test.py
.venv\Scripts\python.exe -X utf8 scripts\bundle_release.py bundle\kitchensink4ppt.mcpb
```

`scripts/bundle_release.py` runs the same check the server runs before it opens a
downloaded bundle, then writes `kitchensink4ppt-X.Y.Z.mcpb` (the same bytes) and
`kitchensink4ppt-X.Y.Z.mcpb.sha256` beside the packed file.

(`mcpb` comes from `npm install -g @anthropic-ai/mcpb`.) The icon is the
KitchenSink4AI brand app icon, replaced as a file when the brand art changes.

## Settings checkboxes (`user_config`)

The 2.0 manifest has none, and its `mcp_config.env` passes one fixed setting,
`KS4P_INSTALL_ROUTE=mcpb`, which tells the server it runs from this bundle, so
`update_now` offers the new bundle rather than a `uvx` refresh. The two 1.x
checkboxes ("Load every tool at startup" and "Lock the tool set at startup") mapped to `KS4P_ALL_TOOLS` and
`KS4P_LOCK_TOOLS`, which are retired: the server ignores either one set to
`false` (with a one-line notice) and stops the start when `KS4P_LOCK_TOOLS` is
set to `true`, so the manifest does not pass them. Every operation is available from the start. To narrow the
surface, use the `KS4P_READ_ONLY` and `KS4P_FORBID` launch settings.

Tip worth repeating in install docs: in Claude Desktop's Tool permissions,
set the Read-only tools group to Always Allow. Those tools cannot change
anything, and it stops most permission prompts.

## Claude Desktop install (manual test / direct distribution)

Double-click `kitchensink4ppt.mcpb`, or drag it into the Claude Desktop
window, or Settings, Extensions, Advanced settings, Install Extension.
Desktop writes the manifest's `mcp_config` into its config; the first launch
resolves the package from PyPI (needs network and uv on the machine).

## Attach to a GitHub release

Attach the `.mcpb` as a release asset (do not commit the binary):

```
gh release upload vX.Y.Z bundle/kitchensink4ppt.mcpb bundle/kitchensink4ppt-X.Y.Z.mcpb bundle/kitchensink4ppt-X.Y.Z.mcpb.sha256 --repo KitchenSink4AI/KitchenSink4PPT
```

All three files go on every release from 2.0.0 on: `update_now` downloads the
versioned bundle and its `.sha256` from the release of the new version, and the
unversioned file keeps the site's `releases/latest/download` link working.

Because the manifest pins the PyPI version, each release's asset is specific
to that release. Never re-upload an old `.mcpb` to a new tag.

## Platform note

`compatibility.platforms` is declared `["win32", "darwin", "linux"]`. The
file-based majority of the toolset runs on Windows and Linux (macOS is untested in this release), and PDF and image
export fall back to LibreOffice headless where PowerPoint is absent; only the
`com` pack needs Windows with PowerPoint, and pywin32 is gated behind
`sys_platform == 'win32'` in the package metadata.

## Smithery note

The Smithery registry rejects (HTTP 400) any bundle whose manifest declares a
`tools` array (a limitation of the Smithery CLI). This manifest deliberately carries
**no** `tools` array, so it publishes as-is.
