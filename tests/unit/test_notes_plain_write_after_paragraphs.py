"""set_notes(text=...) over notes this server wrote without formatting.

The refusal "would flatten N runs of mixed character formatting" counted every
run of a paragraph that has more than one run, even when the runs carry the
same properties. Notes written by set_notes(paragraphs=[...]) with no
formatting are exactly that. The refusal now counts only runs whose
properties actually differ.
"""

from __future__ import annotations

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import UnsupportedStructure
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import notes as nt

PLAIN_PARAGRAPHS = [
    {"runs": [{"text": "Welcome "}, {"text": "everyone "}, {"text": "today."}]},
    {"runs": [{"text": "Second "}, {"text": "point."}]},
    {"runs": [{"text": "Closing."}]},
]


def _body(pkg: PptxPackage, index: int = 0):
    part = nt.notes_part_for(pkg, pkg.slide_parts()[index])
    return nt._notes_body_sp(pkg.root(part)).find(qn("p:txBody"))


def test_plain_text_over_unformatted_multi_run_notes_made_by_paragraphs_new_notes(make_deck):
    pkg = PptxPackage(make_deck("p24a.pptx"))
    nt.set_notes(pkg, 0, paragraphs=PLAIN_PARAGRAPHS)
    assert sum(len(p.findall(qn("a:r"))) for p in _body(pkg).findall(qn("a:p"))) == 6
    out = nt.set_notes(pkg, 0, "A fresh line.\nAnd another.")
    assert out["changed"] is True
    assert "flattened" not in out
    assert nt.get_notes(pkg, 0)["text"] == "A fresh line.\nAnd another."


def test_plain_text_over_unformatted_multi_run_notes_made_by_paragraphs_existing_notes(make_deck):
    pkg = PptxPackage(make_deck("p24b.pptx"))
    nt.set_notes(pkg, 0, "seed")
    nt.set_notes(pkg, 0, paragraphs=PLAIN_PARAGRAPHS)
    out = nt.set_notes(pkg, 0, "Replaced.")
    assert out["changed"] is True and "flattened" not in out
    assert nt.get_notes(pkg, 0)["text"] == "Replaced."


def test_runs_that_only_differ_in_noise_attributes_are_not_a_loss(make_deck):
    """lang, dirty and spelling marks are not formatting."""
    pkg = PptxPackage(make_deck("p24c.pptx"))
    nt.set_notes(pkg, 0, paragraphs=PLAIN_PARAGRAPHS)
    runs = _body(pkg).findall(f"{qn('a:p')}/{qn('a:r')}")
    for i, r in enumerate(runs):
        rpr = r.find(qn("a:rPr"))
        if rpr is None:
            rpr = etree.Element(qn("a:rPr"))
            r.insert(0, rpr)
        rpr.set("lang", "en-US" if i % 2 else "en-GB")
        if i % 3 == 0:
            rpr.set("dirty", "0")
        if i % 3 == 1:
            rpr.set("err", "1")
    out = nt.set_notes(pkg, 0, "Replaced again.")
    assert out["changed"] is True and "flattened" not in out


def test_one_differently_sized_run_is_still_refused_and_only_it_is_counted(make_deck):
    pkg = PptxPackage(make_deck("p24d.pptx"))
    paras = [dict(p) for p in PLAIN_PARAGRAPHS]
    paras[0] = {"runs": [{"text": "Welcome "}, {"text": "everyone "},
                         {"text": "today.", "size_pt": 14}]}
    nt.set_notes(pkg, 0, paragraphs=paras)
    with pytest.raises(UnsupportedStructure) as exc:
        nt.set_notes(pkg, 0, "Replaced.")
    msg = str(exc.value)
    assert "1 run of mixed character formatting" in msg, msg
    assert "6 runs" not in msg
    out = nt.set_notes(pkg, 0, "Replaced.", flatten=True)
    assert out["changed"] is True


def test_a_run_with_a_property_this_server_does_not_read_is_still_a_difference(make_deck):
    pkg = PptxPackage(make_deck("p24e.pptx"))
    nt.set_notes(pkg, 0, paragraphs=PLAIN_PARAGRAPHS)
    first = _body(pkg).findall(f"{qn('a:p')}/{qn('a:r')}")[1]
    rpr = etree.Element(qn("a:rPr"))
    rpr.set("strike", "sngStrike")
    first.insert(0, rpr)
    with pytest.raises(UnsupportedStructure) as exc:
        nt.set_notes(pkg, 0, "Replaced.")
    assert "mixed character formatting" in str(exc.value)


def test_bold_italics_bullets_and_links_are_still_named(make_deck):
    pkg = PptxPackage(make_deck("p24f.pptx"))
    paras = [dict(p) for p in PLAIN_PARAGRAPHS]
    paras[1] = {"runs": [{"text": "Second ", "bold": True}, {"text": "point."}]}
    nt.set_notes(pkg, 0, paragraphs=paras)
    with pytest.raises(UnsupportedStructure) as exc:
        nt.set_notes(pkg, 0, "Replaced.")
    assert "1 bold run" in str(exc.value)
