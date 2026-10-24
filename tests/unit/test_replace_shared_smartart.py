"""search_and_replace on SmartArt whose parts two frames share.

Older versions of duplicate_slide left the copy pointing at the source's
SmartArt parts, so decks exist where two frames (on two slides, or on one)
show one data model and one cached drawing. search_and_replace walked the
shared parts once per frame: a replacement that contains the search text
was applied twice ("Plan" -> "Plan B B"), and a call scoped to one slide
changed the SmartArt on the other slide too. Now each part is replaced once
per call, and a scoped call first gives the slide in scope its own copy of
a SmartArt it shares with a slide outside the scope, and says so.

The decks are made up and built by code.
"""

from __future__ import annotations

import copy
import posixpath

import pytest

from kitchensink4ppt.core.package import PptxPackage, qn, rels_name, resolve_target
from kitchensink4ppt.ops import read, slides
from kitchensink4ppt.ops import text as tx
from test_diagrams_write import add_smartart

_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _shared_deck(tmp_path, same_slide: bool):
    path = tmp_path / "shared.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)
    slides.insert_slide(pkg, 6)
    p1, p2 = pkg.slide_parts()[:2]
    sid = add_smartart(pkg, p1, ["Plan", "Draft", "Review"])
    tree1 = pkg.root(p1).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    frame = next(
        f for f in tree1.iter(qn("p:graphicFrame"))
        if f.find(f"{qn('p:nvGraphicFramePr')}/{qn('p:cNvPr')}").get("id") == str(sid)
    )
    dst = p1 if same_slide else p2
    twin = copy.deepcopy(frame)
    twin.find(f"{qn('p:nvGraphicFramePr')}/{qn('p:cNvPr')}").set("id", str(pkg.next_shape_id(dst)))
    rel_ids = twin.find(f".//{{{_DGM}}}relIds")
    if not same_slide:
        rels = {r.get("Id"): r for r in pkg.root(rels_name(p1))}
        for key in ("dm", "lo", "qs", "cs"):
            rel = rels[rel_ids.get(f"{{{_R}}}{key}")]
            target = resolve_target(p1, rel.get("Target"))
            new = pkg.add_relationship(
                p2, rel.get("Type"), posixpath.relpath(target, posixpath.dirname(p2))
            )
            rel_ids.set(f"{{{_R}}}{key}", new)
    pkg.root(dst).find(f"{qn('p:cSld')}/{qn('p:spTree')}").append(twin)
    pkg.mark_dirty(dst)
    pkg.save(do_backup=False)
    return path


def _texts(pkg):
    """[(slide index, [node texts])] for every SmartArt frame in the deck."""
    return [
        (item["slide_index"], [n["text"] for n in item["nodes"]])
        for item in read.diagram_text(pkg)["items"]
    ]


def _drawing_texts(pkg):
    out = []
    for name in sorted(n for n in pkg.part_names() if n.startswith("ppt/diagrams/drawing")):
        if name.endswith(".xml"):
            out.append("".join(t.text or "" for t in pkg.root(name).iter(qn("a:t"))))
    return out


@pytest.mark.parametrize("same_slide", [False, True], ids=["two slides", "one slide"])
def test_a_shared_smartart_is_replaced_once(tmp_path, same_slide):
    pkg = PptxPackage(_shared_deck(tmp_path, same_slide))
    res = tx.search_and_replace(pkg, "Plan", "Plan B")
    pkg = PptxPackage(pkg.save(do_backup=False))
    assert [t for _i, t in _texts(pkg)] == [["Plan B", "Draft", "Review"]] * 2
    assert all("Plan B B" not in t for t in _drawing_texts(pkg))
    assert res["total"] == 2  # the model and its cached drawing, once each


def test_a_scoped_replace_leaves_the_other_slide_alone(tmp_path):
    pkg = PptxPackage(_shared_deck(tmp_path, same_slide=False))
    res = tx.search_and_replace(pkg, "Plan", "Plan B", scope=0)
    pkg = PptxPackage(pkg.save(do_backup=False))
    assert _texts(pkg) == [(0, ["Plan B", "Draft", "Review"]), (1, ["Plan", "Draft", "Review"])]
    assert sorted(_drawing_texts(pkg)) == sorted(["Plan BDraftReview", "PlanDraftReview"])
    assert res["total"] == 2  # the model and its cached drawing, once each
    assert any("own copy" in w for w in res.get("warnings", []))
    parts0 = slides._diagram_graph(pkg, pkg.slide_parts()[0])[0]
    parts1 = slides._diagram_graph(pkg, pkg.slide_parts()[1])[0]
    assert parts0 and parts1 and not parts0 & parts1


def test_a_scoped_replace_with_no_match_copies_nothing(tmp_path):
    pkg = PptxPackage(_shared_deck(tmp_path, same_slide=False))
    before = sorted(pkg.part_names())
    res = tx.search_and_replace(pkg, "Nowhere", "X", scope=0)
    assert res["total"] == 0 and "warnings" not in res
    assert sorted(pkg.part_names()) == before
