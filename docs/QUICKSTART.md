## KitchenSink4PPT Quickstart

**Install.** Claude Desktop: the KitchenSink4PPT extension. Anywhere else:
`uvx kitchensink4ppt`.

**First deck.** Ask Claude: "Add a slide to pitch.pptx with a three-box
architecture diagram." Diagrams are built from real shapes with glued
connectors, so you can drag a box in PowerPoint afterward and the arrows
follow. Every save is validated, and a backup lands first.

**The habits that matter.** Big decks read in pages, never all at once, so
one look never floods the conversation. `diagnose` is the health check: run
it when anything seems off, and it also reports the version and update
state (weekly check, off with KS4P_UPDATE_CHECK=off). Close your own
PowerPoint before long automated runs; the server refuses to fight your open
copy for the file.
