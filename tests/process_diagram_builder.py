"""Process diagram builder: the acceptance diagram for the graphics engine.

Three labeled vertex nodes (Plan, Build, Ship), a central Hub node, three
straight connectors GLUED Hub -> vertex, three curved perimeter connectors
with arrowheads (P -> B -> S cycle), In and Out badges beside the Hub node, all
grouped. build() writes the before-tweak deck; tweak() moves the Hub node by
a delta and restyles one edge dashed red through set_shape, proving the
glue survives semantic post-editing.

Shared between tests/unit/test_shapes.py (acceptance test) and the COM
render gate script. This module WRITES files (it is a test/gate helper, not
an ops module).
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parents[1] / "src"))

import make_corpus  # noqa: E402  (lives beside this file)
from kitchensink4ppt.core.package import PptxPackage  # noqa: E402
from kitchensink4ppt.ops import shapes as shp  # noqa: E402
from kitchensink4ppt.ops import slides as sl  # noqa: E402
from kitchensink4ppt.ops.read import get_presentation_info  # noqa: E402

NODE_W, NODE_H = 1.9, 0.85
HUB_W, HUB_H = 1.7, 0.95
BADGE = 0.5

NODE_FILL = "2E5E8C"
NODE_LINE = {"width": 1.5, "color": "1B3A57"}
HUB_FILL = "C05A2E"
HUB_LINE = {"width": 1.75, "color": "7A3517"}
BADGE_FILL = "F2C14E"
BADGE_LINE = {"width": 1.0, "color": "8C6A1D"}
SPOKE = {"width": 1.75, "color": "404040"}
CYCLE = {
    "width": 2.25,
    "color": "2E8C5E",
    "tail": {"type": "triangle", "w": "med", "len": "med"},
}
TEXT = {"size": 16, "color": "FFFFFF", "bold": True}
BADGE_TEXT = {"size": 14, "color": "3B2F0B", "bold": True}


def build(path: Path) -> dict:
    """Build the deck at `path` (fresh synthetic base + one diagram slide)
    and return {"path", "slide", "ids": {...}, "group_id"}."""
    make_corpus.build_deck(path, seed=4, extra_slides=0)
    pkg = PptxPackage(path)
    slide = sl.insert_slide(pkg, 0)["index"]

    size = get_presentation_info(pkg)["slide_size"]
    sw, sh = size["cx_in"], size["cy_in"]
    cx = sw / 2

    def node(name, ccx, ccy, w, h, text, fill, line, style, shape="rounded_rect", adj=None):
        return shp.insert_shape(
            pkg, slide, shape,
            ccx - w / 2, ccy - h / 2, w, h,
            adjustments=adj, fill=fill, line=line,
            text=text, text_style=style, name=name,
        )["shape_id"]

    n_plan = node("Plan node", cx, 1.35, NODE_W, NODE_H, "Plan", NODE_FILL, NODE_LINE, TEXT, adj={"adj": 0.28})
    n_build = node("Build node", cx + 2.9, sh - 1.95, NODE_W, NODE_H, "Build", NODE_FILL, NODE_LINE, TEXT, adj={"adj": 0.28})
    n_ship = node("Ship node", cx - 2.9, sh - 1.95, NODE_W, NODE_H, "Ship", NODE_FILL, NODE_LINE, TEXT, adj={"adj": 0.28})
    n_hub = node("Hub node", cx, 3.85, HUB_W, HUB_H, "Hub", HUB_FILL, HUB_LINE, TEXT, adj={"adj": 0.22})
    badge_in = node("In badge", cx - HUB_W / 2 - 0.45, 3.85, BADGE, BADGE, "In", BADGE_FILL, BADGE_LINE, BADGE_TEXT, shape="ellipse")
    badge_out = node("Out badge", cx + HUB_W / 2 + 0.45, 3.85, BADGE, BADGE, "Out", BADGE_FILL, BADGE_LINE, BADGE_TEXT, shape="ellipse")

    def spoke(name, end_shape):
        return shp.insert_connector(
            pkg, slide, "straight",
            start_shape=n_hub, end_shape=end_shape,
            line=SPOKE, name=name,
        )["shape_id"]

    spoke_plan = spoke("Hub to Plan", n_plan)
    spoke_build = spoke("Hub to Build", n_build)
    spoke_ship = spoke("Hub to Ship", n_ship)

    def cycle(name, start_shape, start_site, end_shape, end_site):
        return shp.insert_connector(
            pkg, slide, "curved",
            start_shape=start_shape, start_site=start_site,
            end_shape=end_shape, end_site=end_site,
            line=CYCLE, name=name,
        )["shape_id"]

    # Perimeter cycle P -> B -> S -> P, bowing outward.
    edge_pb = cycle("Plan to Build", n_plan, 3, n_build, 0)
    edge_bs = cycle("Build to Ship", n_build, 2, n_ship, 2)
    edge_sp = cycle("Ship to Plan", n_ship, 1, n_plan, 1)

    ids = {
        "plan": n_plan, "build": n_build, "ship": n_ship, "hub": n_hub,
        "badge_in": badge_in, "badge_out": badge_out,
        "spoke_plan": spoke_plan, "spoke_build": spoke_build, "spoke_ship": spoke_ship,
        "edge_pb": edge_pb, "edge_bs": edge_bs, "edge_sp": edge_sp,
    }
    group_id = shp.group_shapes(
        pkg, slide, list(ids.values()), name="Process Overview"
    )["group_id"]
    pkg.save(do_backup=False)
    return {"path": path, "slide": slide, "ids": ids, "group_id": group_id}


def tweak(path: Path, info: dict) -> dict:
    """A manual-style nudge, agent-side: move the Hub node by a delta
    and turn the Ship -> Plan edge dashed red, both by shape id, both
    inside the group. Saves in place; returns the set_shape results."""
    pkg = PptxPackage(path)
    slide = info["slide"]
    ids = info["ids"]
    moved = shp.set_shape(pkg, slide, ids["hub"], dx=0.4, dy=-0.3)
    restyled = shp.set_shape(
        pkg, slide, ids["edge_sp"],
        line={"width": 2.25, "color": "C00000", "dash": "dash",
              "tail": {"type": "triangle", "w": "med", "len": "med"}},
    )
    pkg.save(do_backup=False)
    return {"moved": moved, "restyled": restyled}
