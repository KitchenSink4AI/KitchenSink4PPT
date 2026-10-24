"""Package-mode soft line breaks stay distinct from paragraphs."""

from __future__ import annotations

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import UnsupportedStructure
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import (
    batch,
    comments,
    diagrams,
    notes,
    read,
    shapes,
    tables,
    text,
)


def _first_text_shape(pkg: PptxPackage, slide_index: int = 0):
    part = pkg.slide_parts()[slide_index]
    tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    for element, kind, _z, _parent in read.iter_shapes(tree):
        if kind not in (
            "group", "picture", "table", "chart", "diagram", "ole"
        ):
            body = element.find(qn("p:txBody"))
            cnvpr = read._cnvpr(element)
            if body is not None and cnvpr is not None:
                return part, element, body, int(cnvpr.get("id"))
    raise AssertionError("synthetic deck has no text shape")


def _write_mixed_break_body(body: etree._Element) -> None:
    for paragraph in body.findall(qn("a:p")):
        body.remove(paragraph)
    first = etree.SubElement(body, qn("a:p"))
    etree.SubElement(
        etree.SubElement(first, qn("a:r")), qn("a:t")
    ).text = "Alpha"
    etree.SubElement(first, qn("a:br"))
    etree.SubElement(
        etree.SubElement(first, qn("a:r")), qn("a:t")
    ).text = "Beta"
    second = etree.SubElement(body, qn("a:p"))
    etree.SubElement(
        etree.SubElement(second, qn("a:r")), qn("a:t")
    ).text = "Gamma"


def _table_element(pkg: PptxPackage, slide_index: int, shape_id: int):
    part = pkg.slide_parts()[slide_index]
    tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    for element, kind, _z, _parent in read.iter_shapes(tree):
        cnvpr = read._cnvpr(element)
        if (
            kind == "table"
            and cnvpr is not None
            and int(cnvpr.get("id")) == shape_id
        ):
            return element
    raise AssertionError(f"table shape {shape_id} not found")


def test_read_ranges_replace_and_shape_rewrite_round_trip(make_deck):
    path = make_deck()
    pkg = PptxPackage(path)
    part, _shape, body, shape_id = _first_text_shape(pkg)
    _write_mixed_break_body(body)
    pkg.mark_dirty(part)
    pkg.save()

    reopened = PptxPackage(path)
    _part, shape, body, _sid = _first_text_shape(reopened)
    paragraphs = body.findall(qn("a:p"))
    assert read.paragraph_text(paragraphs[0]) == "Alpha\vBeta"
    assert read.shape_text(shape) == "Alpha\vBeta\nGamma"
    assert "Alpha\vBeta\nGamma" in read.get_text(reopened, scope=0)["text"]

    hits = [
        match for match in read.find_text(reopened, "\v", scope=0)["matches"]
        if match.get("shape_id") == shape_id
    ]
    assert [(hit["paragraph"], hit["start"], hit["end"]) for hit in hits] == [
        (0, 5, 6)
    ]
    text.format_text(
        reopened, 0, shape_id, paragraph=0, start=5, end=6, bold=True
    )
    break_props = body.findall(qn("a:p"))[0].find(
        f"{qn('a:br')}/{qn('a:rPr')}"
    )
    assert break_props.get("b") == "1"

    result = text.search_and_replace(reopened, "Alpha\vBeta", "One\vTwo", scope=0)
    assert result["total"] == 1
    assert read.shape_text(shape) == "One\vTwo\nGamma"
    reopened.save()

    saved = PptxPackage(path)
    _part, saved_shape, saved_body, _sid = _first_text_shape(saved)
    assert len(saved_body.findall(qn("a:p"))) == 2
    assert len(saved_body.findall(qn("a:p"))[0].findall(qn("a:br"))) == 1
    value = read.shape_text(saved_shape)
    shapes.set_shape(saved, 0, shape_id, text=value)
    saved.save()

    final = PptxPackage(path)
    _part, final_shape, final_body, _sid = _first_text_shape(final)
    assert read.shape_text(final_shape) == "One\vTwo\nGamma"
    assert len(final_body.findall(qn("a:p"))) == 2
    assert len(final_body.findall(qn("a:p"))[0].findall(qn("a:br"))) == 1


def test_placeholder_table_and_batch_writers_emit_a_br(make_deck):
    path = make_deck()
    pkg = PptxPackage(path)

    placeholder = text.set_placeholder_text(
        pkg, 0, "title", text="Title\vWrap\nSecond paragraph"
    )
    shape_id = placeholder["shape_id"]
    slide_tree = pkg.root(pkg.slide_parts()[0]).find(
        f"{qn('p:cSld')}/{qn('p:spTree')}"
    )
    placeholder_shape = next(
        element for element, _kind, _z, _parent in read.iter_shapes(slide_tree)
        if read._cnvpr(element) is not None
        and int(read._cnvpr(element).get("id")) == shape_id
    )
    assert read.shape_text(placeholder_shape) == "Title\vWrap\nSecond paragraph"
    batch.apply_edits(
        pkg,
        [{
            "op": "set_text",
            "slide": 0,
            "shape": shape_id,
            "text": "Batch\vWrap\nSecond paragraph",
        }],
    )

    table_info = read.list_elements(pkg, "tables")["items"][0]
    slide_index = table_info["slide_index"]
    table_id = table_info["id"]
    tables.set_table_cells(
        pkg,
        slide_index,
        {"shape_id": table_id},
        [{"row": 0, "col": 0, "text": "Cell\vWrap\nNext"}],
    )
    # A format-only rewrite reads the current cell text and must not turn its
    # soft break into a paragraph boundary.
    tables.set_table_cells(
        pkg,
        slide_index,
        {"shape_id": table_id},
        [{"row": 0, "col": 0, "bold": True}],
    )
    pkg.save()

    reopened = PptxPackage(path)
    _part, shape, body, _sid = _first_text_shape(reopened)
    if int(read._cnvpr(shape).get("id")) != shape_id:
        part = reopened.slide_parts()[0]
        tree = reopened.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
        shape = next(
            element for element, _kind, _z, _parent in read.iter_shapes(tree)
            if read._cnvpr(element) is not None
            and int(read._cnvpr(element).get("id")) == shape_id
        )
        body = shape.find(qn("p:txBody"))
    assert read.shape_text(shape) == "Batch\vWrap\nSecond paragraph"
    assert len(body.findall(qn("a:p"))) == 2
    assert len(body.findall(qn("a:p"))[0].findall(qn("a:br"))) == 1

    table = _table_element(reopened, slide_index, table_id)
    assert read.table_cells(table)[0][0] == "Cell\vWrap\nNext"
    cell_body = read.table_element(table).find(
        f"{qn('a:tr')}/{qn('a:tc')}/{qn('a:txBody')}"
    )
    assert len(cell_body.findall(qn("a:p"))) == 2
    assert len(cell_body.findall(qn("a:p"))[0].findall(qn("a:br"))) == 1


def test_notes_comments_and_diagram_text_use_same_break_contract(make_deck):
    path = make_deck()
    pkg = PptxPackage(path)
    notes.set_notes(pkg, 0, "Note\vWrap\nNext", flatten=True)
    pkg.save()

    reopened = PptxPackage(path)
    assert read.notes_text(reopened, reopened.slide_parts()[0]) == "Note\vWrap\nNext"
    rich = notes.get_notes(reopened, 0)
    assert rich["text"] == "Note\vWrap\nNext"
    assert any(run["text"] == "\v" for run in rich["paragraphs"][0]["runs"])
    assert notes.set_notes(reopened, 0, paragraphs=rich["paragraphs"])["changed"] is False

    comment_body = comments._txbody("Comment\vWrap\nNext")
    assert comments._txbody_text(comment_body) == "Comment\vWrap\nNext"
    assert len(comment_body.findall(qn("a:p"))) == 2
    assert len(comment_body.findall(qn("a:p"))[0].findall(qn("a:br"))) == 1

    diagram_body = etree.fromstring(
        '<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        "<a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Old</a:t></a:r></a:p></a:txBody>"
    )
    changed, _collapsed = diagrams._set_body_text(diagram_body, "Node\vWrap\nNext")
    assert changed is True
    assert [read.paragraph_text(p) for p in diagram_body.findall(qn("a:p"))] == [
        "Node\vWrap",
        "Next",
    ]


def test_search_replace_refuses_paragraph_break_but_accepts_soft_break(make_deck):
    pkg = PptxPackage(make_deck())
    _part, _shape, _body, shape_id = _first_text_shape(pkg)
    shapes.set_shape(pkg, 0, shape_id, text="Alpha")
    with pytest.raises(UnsupportedStructure, match="cannot insert a paragraph break"):
        text.search_and_replace(pkg, "Alpha", "One\nTwo", scope=0)
    assert read.find_text(pkg, "Alpha", scope=0)["count"] == 1

    result = text.search_and_replace(pkg, "Alpha", "One\vTwo", scope=0)
    assert result["total"] == 1
    result = text.search_and_replace(pkg, "ne\vTw", "X\vY", scope=0)
    assert result["total"] == 1
    _part, shape, body, _sid = _first_text_shape(pkg)
    assert read.shape_text(shape) == "OX\vYo"
    assert len(body.findall(qn("a:p"))) == 1
    assert len(body.findall(qn("a:p"))[0].findall(qn("a:br"))) == 1
