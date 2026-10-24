"""Deleting a slide removes companion parts that only it reached, cycles included.

delete_slide garbage-collects what the deleted slide referenced by counting
references: a part survives while any relationship anywhere names it. Two
companion parts that name each other (A to B and B to A) and that nothing
else reaches kept each other alive, so the deleted slide's data stayed in
the file. Parts are collected by reachability now: what the rest of the
package still reaches stays, everything else the slide reached goes.

The companion parts and their relationship type are made up for the test
(an unknown relationship family, the way add-ins write their own data); the
deck is made up and built by code.
"""

from __future__ import annotations

import pytest

from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import slides as sl

RT = "http://schemas.example.com/relationships/companion"
CT = "application/vnd.example.companion+xml"
A, B = "ppt/companions/partA.xml", "ppt/companions/partB.xml"


def _part(pkg: PptxPackage, name: str, label: str) -> None:
    body = f'<?xml version="1.0" encoding="UTF-8"?><companion label="{label}"/>'.encode()
    pkg.add_part_with_content_type(name, body, CT)


def _deck(make_deck, edges: list[tuple[str, str]], from_slides: list[int]):
    """A made-up deck whose slides `from_slides` name part A, plus `edges`
    between the companion parts."""
    path = make_deck("companions.pptx")
    pkg = PptxPackage(path)
    _part(pkg, A, "made-up companion A")
    _part(pkg, B, "made-up companion B")
    for i in from_slides:
        slide_part = pkg.slide_parts()[i]
        pkg.add_relationship(slide_part, RT, "../companions/partA.xml")
    for src, dst in edges:
        pkg.add_relationship(src, RT, dst.rsplit("/", 1)[1])
    pkg.save(do_backup=False)
    return path


@pytest.mark.parametrize("edges", [
    [(A, B), (B, A)],
    [(A, B)],
    [(A, A)],
], ids=["two-part cycle", "chain", "self loop"])
def test_companions_only_the_deleted_slide_reached_go(make_deck, edges):
    path = _deck(make_deck, edges, from_slides=[2])
    pkg = PptxPackage(path)
    res = sl.delete_slide(pkg, 2)
    pkg.save(do_backup=False)

    pkg = PptxPackage(path)
    assert not pkg.has_part(A)  # before the fix: the cycle kept A and B alive
    if any(B in e for e in edges):
        assert not pkg.has_part(B)
    assert A in res["gc_parts"]


def test_a_cycle_another_slide_still_reaches_stays(make_deck):
    path = _deck(make_deck, [(A, B), (B, A)], from_slides=[2, 3])
    pkg = PptxPackage(path)
    res = sl.delete_slide(pkg, 2)
    pkg.save(do_backup=False)
    pkg = PptxPackage(path)
    assert pkg.has_part(A) and pkg.has_part(B)
    assert A not in res["gc_parts"] and B not in res["gc_parts"]


def test_the_design_and_notes_master_stay(make_deck):
    """The deleted slide reached its layout, master, theme and notes master
    too; the rest of the deck still reaches them, so none of them goes."""
    path = _deck(make_deck, [(A, B), (B, A)], from_slides=[0])
    pkg = PptxPackage(path)
    before = set(pkg.part_names())
    res = sl.delete_slide(pkg, 0)
    pkg.save(do_backup=False)
    gone = before - set(PptxPackage(path).part_names())
    kinds = {name.split("/")[1] for name in gone if name.count("/") >= 2}
    assert kinds <= {"slides", "notesSlides", "companions"}, sorted(gone)
    assert set(res["gc_parts"]) <= gone
