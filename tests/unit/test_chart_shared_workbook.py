"""Two charts that share one embedded workbook: changing one leaves the other alone.

Two classic charts can name the same embedded workbook part: decks
assembled by other tools do this, and a cross-deck copy keeps a sharing
it finds. update_chart_data rewrote the cells of that one workbook, so the
other chart's backing data changed while its cached numbers did not, and
the next Edit Data in PowerPoint would show chart A's numbers in chart B.
The edited chart now gets a workbook of its own first; the other chart
keeps the bytes it had.

The deck is made up and built by code.
"""

from __future__ import annotations

import io
import zipfile

from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, rels_name, resolve_target
from kitchensink4ppt.ops import charts as ch
from kitchensink4ppt.ops import slides as sl

RT_PACKAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
CATS = ["North", "South", "East"]
SERIES = [{"name": "Units", "values": [1, 2, 3]}]


def _workbook_of(pkg: PptxPackage, chart_part: str) -> str:
    for rel in pkg.root(rels_name(chart_part)):
        if rel.get("Type") == RT_PACKAGE:
            return resolve_target(chart_part, rel.get("Target"))
    raise AssertionError("no embedded workbook")


def _cell(blob: bytes, ref: str) -> str | None:
    with zipfile.ZipFile(io.BytesIO(blob)) as zf:
        sheet = etree.fromstring(zf.read("xl/worksheets/sheet1.xml"))
    for c in sheet.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"):
        if c.get("r") == ref:
            v = c.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
            return v.text if v is not None else None
    return None


def _two_charts(make_deck, share: bool):
    path = make_deck("charts.pptx", extra_slides=0)
    pkg = PptxPackage(path)
    slide = sl.insert_slide(pkg, 0)["index"]
    a = ch.create_chart(pkg, slide, "column", CATS, SERIES, 0.5, 1, 4, 3)
    b = ch.create_chart(pkg, slide, "column", CATS, SERIES, 5, 1, 4, 3)
    if share:  # chart B names chart A's workbook, as a shared bundle does
        rels = pkg.root(rels_name(b["chart_part"]))
        for rel in rels:
            if rel.get("Type") == RT_PACKAGE:
                rel.set("Target", "../embeddings/" + a["embedded_workbook"].rsplit("/", 1)[1])
        pkg.mark_dirty(rels_name(b["chart_part"]))
        pkg.remove_part(b["embedded_workbook"])
    pkg.save(do_backup=False)
    return path, slide, a, b


def test_updating_one_of_two_charts_that_share_a_workbook(make_deck):
    path, slide, a, b = _two_charts(make_deck, share=True)
    pkg = PptxPackage(path)
    shared = _workbook_of(pkg, a["chart_part"])
    assert _workbook_of(pkg, b["chart_part"]) == shared
    before = pkg.raw_part(shared)

    res = ch.update_chart_data(pkg, slide, a["shape_id"], CATS,
                               [{"name": "Units", "values": [9, 2, 3]}])
    pkg.save(do_backup=False)

    pkg = PptxPackage(path)
    assert _workbook_of(pkg, b["chart_part"]) == shared
    assert pkg.raw_part(shared) == before  # before the fix: chart B's backing cells became 9
    own = _workbook_of(pkg, a["chart_part"])
    assert own != shared
    assert _cell(pkg.raw_part(own), "B2") == "9"
    assert _cell(pkg.raw_part(shared), "B2") == "1"
    assert res["embedded_workbook"] == "updated_in_own_copy"


def test_a_chart_with_its_own_workbook_is_still_updated_in_place(make_deck):
    path, slide, a, _b = _two_charts(make_deck, share=False)
    pkg = PptxPackage(path)
    res = ch.update_chart_data(pkg, slide, a["shape_id"], CATS,
                               [{"name": "Units", "values": [9, 2, 3]}])
    pkg.save(do_backup=False)
    pkg = PptxPackage(path)
    assert _workbook_of(pkg, a["chart_part"]) == a["embedded_workbook"]
    assert _cell(pkg.raw_part(a["embedded_workbook"]), "B2") == "9"
    assert res["embedded_workbook"] == "updated_in_place"
