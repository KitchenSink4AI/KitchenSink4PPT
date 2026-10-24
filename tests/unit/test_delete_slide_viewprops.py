"""Deleting a slide that PowerPoint's outline view lists.

PowerPoint keeps the outline view's collapsed slides in ppt/viewProps.xml
(p:outlineViewPr/p:sldLst/p:sld r:id="..." collapse="1"), each pointing at
the slide through a relationship of viewProps.xml. delete_slide dropped that
relationship with the other links to the deleted slide but left the p:sld
entry naming it, so the save check found a dangling relationship id in
ppt/viewProps.xml and refused every save: the slide could not be deleted at
all. The entry goes with its relationship now.

The deck is made up and built by code.
"""

from __future__ import annotations

import posixpath

from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, qn, rels_name
from kitchensink4ppt.ops.slides import RT_SLIDE, delete_slide

VIEW = "ppt/viewProps.xml"


def _list_in_outline(pkg: PptxPackage, parts: list[str]) -> list[str]:
    """Collapse the given slides in the outline view, as PowerPoint writes it."""
    root = pkg.root(VIEW)
    outline = etree.Element(qn("p:outlineViewPr"))
    cview = etree.SubElement(outline, qn("p:cViewPr"))
    scale = etree.SubElement(cview, qn("p:scale"))
    etree.SubElement(scale, qn("a:sx"), n="33", d="100")
    etree.SubElement(scale, qn("a:sy"), n="33", d="100")
    etree.SubElement(cview, qn("p:origin"), x="0", y="0")
    lst = etree.SubElement(outline, qn("p:sldLst"))
    rids = []
    for part in parts:
        rid = pkg.add_relationship(VIEW, RT_SLIDE, "slides/" + posixpath.basename(part))
        etree.SubElement(lst, qn("p:sld"), {qn("r:id"): rid, "collapse": "1"})
        rids.append(rid)
    root.find(qn("p:notesTextViewPr")).addprevious(outline)
    pkg.mark_dirty(VIEW)
    return rids


def _outline_rids(pkg: PptxPackage) -> list[str]:
    return [s.get(qn("r:id")) for s in pkg.root(VIEW).iter(qn("p:sld"))]


def test_delete_a_slide_the_outline_view_lists(make_deck):
    doc = make_deck("outline.pptx")
    pkg = PptxPackage(doc)
    parts = pkg.slide_parts()
    keep_rid, gone_rid = _list_in_outline(pkg, [parts[0], parts[2]])
    pkg.save(do_backup=False)

    pkg = PptxPackage(doc)
    delete_slide(pkg, 2)
    pkg.save(do_backup=False)  # refused before the fix: dangling id in viewProps.xml

    pkg = PptxPackage(doc)
    assert _outline_rids(pkg) == [keep_rid]
    targets = {rel.get("Id"): rel.get("Target") for rel in pkg.root(rels_name(VIEW))}
    assert gone_rid not in targets
    assert targets[keep_rid] == "slides/" + posixpath.basename(parts[0])


def test_delete_the_only_slide_the_outline_view_lists(make_deck):
    doc = make_deck("outline_one.pptx")
    pkg = PptxPackage(doc)
    _list_in_outline(pkg, [pkg.slide_parts()[1]])
    pkg.save(do_backup=False)

    pkg = PptxPackage(doc)
    delete_slide(pkg, 1)
    pkg.save(do_backup=False)

    pkg = PptxPackage(doc)
    assert _outline_rids(pkg) == []
    outline = pkg.root(VIEW).find(qn("p:outlineViewPr"))
    assert outline is not None and outline.find(qn("p:cViewPr")) is not None
