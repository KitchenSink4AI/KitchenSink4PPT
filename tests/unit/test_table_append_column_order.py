"""Appending a column to a table keeps each row's cells before its extLst.

PowerPoint 2016 and later end every table row with an a:extLst holding the
row's a16:rowId (and every grid column with one holding its a16:colId).
insert_table_cols(at=ncols) appended the new cell at the very end of each
row, after that extLst, but a row lists its cells first and its extLst
last. Inserting anywhere else was already in order.

The deck is made up and built by code, with row and column ids written the
way PowerPoint writes them.
"""

from __future__ import annotations

import pytest
from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import tables as tb
from kitchensink4ppt.ops.read import get_slide_info

A16 = "http://schemas.microsoft.com/office/drawing/2014/main"
ROW_EXT = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}"
COL_EXT = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}"
TABLE_SLIDE = 2  # make_deck's table slide


def _ids(parent: etree._Element, uri: str, tag: str, val: int) -> None:
    lst = etree.SubElement(parent, qn("a:extLst"))
    ext = etree.SubElement(lst, qn("a:ext"), uri=uri)
    etree.SubElement(ext, f"{{{A16}}}{tag}", nsmap={"a16": A16}, val=str(val))


@pytest.fixture()
def table_deck(make_deck):
    path = make_deck("table.pptx")
    pkg = PptxPackage(path)
    part = get_slide_info(pkg, TABLE_SLIDE)["part"]
    tbl = next(pkg.root(part).iter(qn("a:tbl")))
    for n, gc in enumerate(tbl.find(qn("a:tblGrid")).findall(qn("a:gridCol"))):
        _ids(gc, COL_EXT, "colId", 20000 + n)
    for n, tr in enumerate(tbl.findall(qn("a:tr"))):
        _ids(tr, ROW_EXT, "rowId", 10000 + n)
    pkg.mark_dirty(part)
    pkg.save(do_backup=False)
    return path


def _rows(path):
    pkg = PptxPackage(path)
    part = get_slide_info(pkg, TABLE_SLIDE)["part"]
    return next(pkg.root(part).iter(qn("a:tbl"))).findall(qn("a:tr"))


def _row_in_order(tr: etree._Element) -> bool:
    tags = [child.tag for child in tr]
    return qn("a:extLst") not in tags or tags.index(qn("a:extLst")) == len(tags) - 1


@pytest.mark.parametrize("at, count", [(3, 1), (3, 2), (0, 1), (1, 1)],
                         ids=["append one", "append two", "first", "inside"])
def test_inserted_columns_land_before_the_row_extlst(table_deck, at, count):
    pkg = PptxPackage(table_deck)
    tb.insert_table_cols(pkg, TABLE_SLIDE, None, at, count)
    pkg.save(do_backup=False)

    rows = _rows(table_deck)
    assert all(_row_in_order(tr) for tr in rows)  # before the fix: a:tc after a:extLst on append
    assert all(len(tr.findall(qn("a:tc"))) == 3 + count for tr in rows)
    assert [tr.find(f"{qn('a:extLst')}/{qn('a:ext')}").get("uri") for tr in rows] == [ROW_EXT] * 3
    pkg = PptxPackage(table_deck)
    cells = tb.get_table(pkg, TABLE_SLIDE, None)
    assert cells["cols"] == 3 + count
