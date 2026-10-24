"""Regressions for title semantics and reading order."""

from __future__ import annotations

import base64
import struct
import zlib

import pytest
from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import access, design_check, media, shapes, slides
from kitchensink4ppt.ops.read import _cnvpr, iter_shapes
from kitchensink4ppt.core.messages import MISSING_TITLE, TITLE_READ_AFTER_BODY


@pytest.fixture()
def blank_deck(tmp_path):
    path = tmp_path / "title-semantics.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)
    return pkg


def _as_title_placeholder(
    pkg: PptxPackage, shape_id: int, placeholder_type: str = "title"
) -> None:
    """Turn a generated text shape into a valid, explicit title placeholder."""
    part = pkg.slide_parts()[0]
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    for elem, _kind, _z, _parent in iter_shapes(sp_tree):
        cnvpr = _cnvpr(elem)
        if cnvpr is None or int(cnvpr.get("id")) != shape_id:
            continue
        nvpr = elem.find(f"{qn('p:nvSpPr')}/{qn('p:nvPr')}")
        assert nvpr is not None
        etree.SubElement(nvpr, qn("p:ph"), type=placeholder_type)
        pkg.mark_dirty(part)
        return
    raise AssertionError(f"shape {shape_id} not found")


def _shape(pkg: PptxPackage, text: str, y: float) -> int:
    return shapes.insert_shape(
        pkg, 0, "rectangle", 1, y, 7, 0.7, text=text
    )["shape_id"]


def _picture(pkg: PptxPackage, y: float = 2.0) -> int:
    def chunk(tag: bytes, payload: bytes) -> bytes:
        body = tag + payload
        return (
            struct.pack(">I", len(payload))
            + body
            + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)
        )

    raw = b"\x00" + b"\xc8\x1e\x1e" * 4
    png = (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", 4, 1, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )
    encoded = base64.b64encode(png).decode("ascii")
    return media.insert_image(pkg, 0, encoded, 1, y, 2, 0.5)["shape_id"]


def _mark_decorative(pkg: PptxPackage, shape_id: int) -> None:
    part = pkg.slide_parts()[0]
    elem, _chain = shapes._find_shape(pkg, part, shape_id)
    cnvpr = _cnvpr(elem)
    assert cnvpr is not None
    ext_lst = etree.SubElement(cnvpr, qn("a:extLst"))
    ext = etree.SubElement(ext_lst, qn("a:ext"), uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}")
    decorative = etree.SubElement(
        ext,
        "{http://schemas.microsoft.com/office/drawing/2017/decorative}decorative",
    )
    decorative.set("val", "1")
    pkg.mark_dirty(part)


def _title(
    pkg: PptxPackage,
    text: str = "Quarterly Review",
    placeholder_type: str = "title",
) -> int:
    shape_id = _shape(pkg, text, 0.4)
    _as_title_placeholder(pkg, shape_id, placeholder_type)
    return shape_id


def _missing_title_findings(pkg: PptxPackage) -> list[dict]:
    return design_check.check_layout(
        pkg, 0, ["missing_title"]
    )["findings"]


def _title_order_findings(pkg: PptxPackage) -> list[dict]:
    return [
        finding
        for finding in access.audit_accessibility(pkg, 0)["findings"]
        if finding["check"] == "reading_order"
        and finding.get("reason") == "title_read_after_body"
    ]


def test_only_nonempty_title_placeholders_satisfy_missing_title(blank_deck):
    empty = _title(blank_deck, "")
    findings = _missing_title_findings(blank_deck)
    assert len(findings) == 1
    assert findings[0]["message"] == MISSING_TITLE
    assert findings[0]["shape_ids"] == [empty]
    assert findings[0]["reason"] == "empty_title_placeholder"

    shapes.delete_shape(blank_deck, 0, empty)
    findings = _missing_title_findings(blank_deck)
    assert len(findings) == 1
    assert findings[0]["message"] == MISSING_TITLE
    assert findings[0]["shape_ids"] == []
    assert findings[0]["reason"] == "no_title_placeholder"


def test_styled_top_textbox_does_not_count_as_title(blank_deck):
    styled = _shape(blank_deck, "Looks Like a Title", 0.1)
    findings = _missing_title_findings(blank_deck)
    assert len(findings) == 1
    assert findings[0]["message"] == MISSING_TITLE
    assert findings[0]["shape_ids"] == []
    assert styled not in findings[0]["shape_ids"]


@pytest.mark.parametrize("placeholder_type", ["title", "ctrTitle"])
def test_nonempty_title_placeholder_is_accepted(blank_deck, placeholder_type):
    _title(blank_deck, placeholder_type=placeholder_type)
    assert _missing_title_findings(blank_deck) == []


@pytest.mark.parametrize("extra_shapes", [0, 1, 4], ids=["two", "three", "six"])
def test_title_after_body_is_flagged_at_every_slide_size(blank_deck, extra_shapes):
    body = _shape(blank_deck, "Body", 2.0)
    for index in range(extra_shapes):
        _shape(blank_deck, f"Content {index + 1}", 3.0 + index * 0.7)
    title = _title(blank_deck)

    findings = _title_order_findings(blank_deck)
    assert len(findings) == 1
    finding = findings[0]
    assert finding["message"] == TITLE_READ_AFTER_BODY
    assert finding["shape_ids"][0] == body
    assert finding["shape_ids"][-1] == title
    assert finding["title_shape_id"] == title
    assert finding["body_before_title_ids"][0] == body
    assert finding["document_order"][-1] == title
    assert finding["grouped_title"] is False
    assert finding["heuristic"] is False
    assert finding["suggested_order"][0] == title


@pytest.mark.parametrize("extra_shapes", [0, 1, 4], ids=["two", "three", "six"])
def test_title_first_control_is_never_flagged(blank_deck, extra_shapes):
    _title(blank_deck)
    _shape(blank_deck, "Body", 2.0)
    for index in range(extra_shapes):
        _shape(blank_deck, f"Content {index + 1}", 3.0 + index * 0.7)
    assert _title_order_findings(blank_deck) == []


def test_office_decorative_picture_before_title_is_ignored(blank_deck):
    picture = _picture(blank_deck)
    _mark_decorative(blank_deck, picture)
    _title(blank_deck)

    assert _title_order_findings(blank_deck) == []


def test_ordinary_picture_before_title_is_still_flagged(blank_deck):
    picture = _picture(blank_deck)
    title = _title(blank_deck)

    findings = _title_order_findings(blank_deck)
    assert len(findings) == 1
    assert findings[0]["body_before_title_ids"] == [picture]
    assert findings[0]["title_shape_id"] == title


def test_grouped_title_after_body_is_flagged_with_honest_limits(blank_deck):
    body = _shape(blank_deck, "Body", 2.0)
    title = _title(blank_deck)
    group = shapes.group_shapes(blank_deck, 0, [body, title])["group_id"]

    findings = _title_order_findings(blank_deck)
    assert len(findings) == 1
    finding = findings[0]
    assert finding["message"] == TITLE_READ_AFTER_BODY
    assert finding["shape_ids"] == [body, title]
    assert finding["title_shape_id"] == title
    assert finding["title_group_id"] == group
    assert finding["grouped_title"] is True
    assert finding["heuristic"] is True
    assert finding["suggested_order"] is None
    assert "op='ungroup_shapes'" in finding["fix"]
    assert f"'group': {group}" in finding["fix"]


def test_grouped_title_first_control_is_not_flagged(blank_deck):
    title = _title(blank_deck)
    body = _shape(blank_deck, "Body", 2.0)
    shapes.group_shapes(blank_deck, 0, [title, body])
    assert _title_order_findings(blank_deck) == []
