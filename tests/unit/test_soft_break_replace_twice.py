"""Several matches in one run, each replaced with text holding a soft break.

search_and_replace snapshots a paragraph's run map once and applies the
matches right to left. A replacement holding a soft break rebuilds the run
it lands in, so the next match to its left still pointed at the removed
run and the call failed with a ValueError ("not in list") whenever two
matches shared one run. The same text split into one run per match
worked, so harmless run fragmentation decided whether the call succeeded.
The answer must not depend on how the text is split into runs.

The deck is made up and built by code.
"""

from __future__ import annotations

import pytest
from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import read, shapes, text
from kitchensink4ppt.ops import slides as sl

SPLITS = {
    "one run": ["A A A"],
    "a run per match": ["A ", "A ", "A"],
    "uneven": ["A A", " A"],
}


def _deck_with_runs(make_deck, runs: list[str]):
    """A blank slide with one text box whose paragraph holds `runs`, every
    run bold so formatting survival is visible."""
    pkg = PptxPackage(make_deck("runs.pptx", extra_slides=0))
    slide = sl.insert_slide(pkg, 0)["index"]
    sid = shapes.insert_shape(pkg, slide, "rect", 1, 1, 4, 1, text="seed")["shape_id"]
    part = read.get_slide_info(pkg, slide)["part"]
    elem, _chain = shapes._find_shape(pkg, part, sid)
    p = elem.find(f"{qn('p:txBody')}/{qn('a:p')}")
    for r in p.findall(qn("a:r")):
        p.remove(r)
    end = p.find(qn("a:endParaRPr"))
    for chunk in runs:
        r = etree.Element(qn("a:r"))
        etree.SubElement(r, qn("a:rPr"), lang="en-US", b="1")
        etree.SubElement(r, qn("a:t")).text = chunk
        if end is not None:
            end.addprevious(r)
        else:
            p.append(r)
    pkg.mark_dirty(part)
    return pkg, slide, part, sid


@pytest.mark.parametrize("split", sorted(SPLITS))
@pytest.mark.parametrize("regex", [False, True], ids=["literal", "regex"])
def test_every_match_in_one_run_takes_a_soft_break(make_deck, split, regex):
    pkg, slide, part, sid = _deck_with_runs(make_deck, SPLITS[split])
    res = text.search_and_replace(pkg, "A", "X\vY", scope=slide, regex=regex)
    assert res["total"] == 3  # before the fix: ValueError when two matches shared a run
    elem, _chain = shapes._find_shape(pkg, part, sid)
    assert read.shape_text(elem) == "X\vY X\vY X\vY"
    p = elem.find(f"{qn('p:txBody')}/{qn('a:p')}")
    assert len(p.findall(qn("a:br"))) == 3
    assert all(r.find(qn("a:rPr")).get("b") == "1" for r in p.findall(qn("a:r")))
    pkg.save(do_backup=False)


def test_a_match_beside_a_spanning_match(make_deck):
    """One match spans two runs while another shares the second run."""
    pkg, slide, part, sid = _deck_with_runs(make_deck, ["see A", "B and AB"])
    res = text.search_and_replace(pkg, "AB", "1\v2", scope=slide)
    assert res["total"] == 2
    elem, _chain = shapes._find_shape(pkg, part, sid)
    assert read.shape_text(elem) == "see 1\v2 and 1\v2"
    pkg.save(do_backup=False)
