"""Text written into an empty SmartArt node keeps the paragraph in schema order.

A SmartArt node nobody has typed into holds a paragraph with nothing but
its end-of-paragraph mark: <a:p><a:endParaRPr/></a:p>, sometimes after an
a:pPr. set_diagram_text added the new run at the end of that paragraph,
after the mark, but a paragraph runs pPr, then its runs and breaks, then
endParaRPr. That happened in the data model and in the cached drawing
alike, and the save check does not look at order, so the out-of-order
XML was saved.

The deck is made up and built by code.
"""

from __future__ import annotations

import pytest
from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import diagrams as dg
from kitchensink4ppt.ops import read, slides
from test_diagrams_write import _DSP_NS, _model_id, add_smartart

_RANK = {qn("a:pPr"): 0, qn("a:endParaRPr"): 2}


def _in_order(p: etree._Element) -> bool:
    ranks = [_RANK.get(child.tag, 1) for child in p]
    return ranks == sorted(ranks)


def _empty(body: etree._Element, ppr: bool) -> None:
    """Leave one paragraph holding only its end mark (and a pPr)."""
    for p in body.findall(qn("a:p")):
        body.remove(p)
    p = etree.SubElement(body, qn("a:p"))
    if ppr:
        etree.SubElement(p, qn("a:pPr"), marL="0", indent="0")
    etree.SubElement(p, qn("a:endParaRPr"), lang="en-US", dirty="0")


def _deck_with_an_empty_node(tmp_path, ppr: bool):
    path = tmp_path / "diagram.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)
    part = pkg.slide_parts()[0]
    sid = add_smartart(pkg, part, ["Plan", "Draft", "Review"])
    data = pkg.root("ppt/diagrams/data1.xml")
    for pt in data.iter("{http://schemas.openxmlformats.org/drawingml/2006/diagram}pt"):
        if pt.get("modelId") == _model_id(1):
            _empty(pt.find("{http://schemas.openxmlformats.org/drawingml/2006/diagram}t"), ppr)
    drawing = pkg.root("ppt/diagrams/drawing1.xml")
    for sp in drawing.iter(f"{{{_DSP_NS}}}sp"):
        if sp.get("modelId") == _model_id(1):
            _empty(sp.find(f"{{{_DSP_NS}}}txBody"), ppr)
    pkg.mark_dirty("ppt/diagrams/data1.xml")
    pkg.mark_dirty("ppt/diagrams/drawing1.xml")
    pkg.save(do_backup=False)
    return path, sid


def _paragraphs(pkg: PptxPackage) -> list[etree._Element]:
    out = []
    for name in ("ppt/diagrams/data1.xml", "ppt/diagrams/drawing1.xml"):
        out.extend(pkg.root(name).iter(qn("a:p")))
    return out


@pytest.mark.parametrize("ppr", [False, True], ids=["bare", "with pPr"])
@pytest.mark.parametrize("value", ["Deliver", "Deliver\vthe draft", "Deliver\nthe draft"],
                         ids=["one line", "soft break", "two lines"])
def test_text_into_an_empty_node_keeps_schema_order(tmp_path, ppr, value):
    path, sid = _deck_with_an_empty_node(tmp_path, ppr)
    pkg = PptxPackage(path)
    assert all(_in_order(p) for p in _paragraphs(pkg))
    dg.set_diagram_text(pkg, 0, sid, nodes=[{"index": 1, "text": value}])
    pkg.save(do_backup=False)

    pkg = PptxPackage(path)
    bad = [etree.tostring(p).decode() for p in _paragraphs(pkg) if not _in_order(p)]
    assert bad == []  # before the fix: the run after a:endParaRPr in both parts
    nodes = [n["text"] for n in read.diagram_text(pkg)["items"][0]["nodes"]]
    assert nodes[0] == "Plan" and nodes[2] == "Review"
    assert nodes[1].replace("\v", " ").replace("\n", " ") == "Deliver" + (
        "" if value == "Deliver" else " the draft")


# The node list diagram_text reports and the one set_diagram_text indexes
# were built by two different filters: the read skipped empty nodes, the
# write skipped only the document and connector points, so presentation
# points that carry a text body counted too. On a SmartArt with empty nodes
# (the ordinary state of one nobody has filled in) or with such points, the
# index diagram_text gave for a node named another node when written to.


def _powerpoint_like(tmp_path, pres_points: bool):
    """Nodes Alpha, (empty), Gamma, (empty), (empty), the shape PowerPoint
    writes for a five-node list with two nodes typed in."""
    path = tmp_path / "list.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)
    part = pkg.slide_parts()[0]
    sid = add_smartart(pkg, part, ["Alpha", "x", "Gamma", "y", "z"],
                       drawing_from_slide=pres_points)
    dgm = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}"
    data = pkg.root("ppt/diagrams/data1.xml")
    for pt in data.iter(dgm + "pt"):
        if pt.get("modelId") in {_model_id(i) for i in (1, 3, 4)}:
            _empty(pt.find(dgm + "t"), ppr=False)
    pkg.mark_dirty("ppt/diagrams/data1.xml")
    pkg.save(do_backup=False)
    return path, sid


@pytest.mark.parametrize("pres_points", [False, True], ids=["nodes only", "with presentation points"])
def test_the_index_diagram_text_reports_names_the_same_node(tmp_path, pres_points):
    path, sid = _powerpoint_like(tmp_path, pres_points)
    pkg = PptxPackage(path)
    item = read.diagram_text(pkg)["items"][0]
    texts = [n["text"] for n in item["nodes"]]
    assert texts == ["Alpha", "", "Gamma", "", ""]  # before the fix: the empty nodes were left out
    assert item["text"] == "Alpha\nGamma"
    dg.set_diagram_text(pkg, 0, sid, nodes=[{"index": texts.index("Gamma"), "text": "Gamma two"}])
    after = [n["text"] for n in read.diagram_text(pkg)["items"][0]["nodes"]]
    assert after == ["Alpha", "", "Gamma two", "", ""]  # before the fix: an empty node took the text


def test_an_empty_node_can_be_filled_by_index(tmp_path):
    path, sid = _powerpoint_like(tmp_path, pres_points=False)
    pkg = PptxPackage(path)
    dg.set_diagram_text(pkg, 0, sid, nodes=[{"index": 1, "text": "Beta"}])
    pkg.save(do_backup=False)
    pkg = PptxPackage(path)
    assert [n["text"] for n in read.diagram_text(pkg)["items"][0]["nodes"]] == [
        "Alpha", "Beta", "Gamma", "", ""]
    assert all(_in_order(p) for p in _paragraphs(pkg))
