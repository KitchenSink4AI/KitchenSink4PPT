"""set_footer(date=True) after a fixed date brings the automatic date back.

Measured in PowerPoint 16 on a made-up deck: after set_footer(date="Fixed
1 Jan") and then set_footer(date=True), the earlier code left the fixed run and
PowerPoint read the slide's date as Fixed (HeadersFooters.DateAndTime
UseFormat 0, Text "Fixed 1 Jan"). PowerPoint's own switch to "Update
automatically" writes a date field (a:fld type="datetime1", a fresh id, the
current date as cached text). The product restores the layout's (or
master's) automatic date field, the one date=True clones onto a slide that
has no date yet; with no such field anywhere it writes PowerPoint's form.
"""

from __future__ import annotations

import datetime

from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import furniture as fu
from kitchensink4ppt.ops.read import slide_table


def _dt(pkg, index):
    return fu._find_ph(pkg.root(slide_table(pkg)[index]["part"]), "dt")


def _fields(sp):
    return [f.get("type") for f in sp.iter(qn("a:fld"))]


def _text(sp):
    return "".join(t.text or "" for t in sp.iter(qn("a:t")))


def _master_dt(pkg):
    return fu._find_ph(pkg.root("ppt/slideMasters/slideMaster1.xml"), "dt")


def _reload(pkg):
    return PptxPackage(pkg.save(do_backup=False))


def test_automatic_after_fixed_restores_the_date_field(make_deck):
    pkg = PptxPackage(make_deck("dt1.pptx", seed=1, extra_slides=1))
    fu.set_footer(pkg, 0, date="Fixed 1 Jan")
    assert _fields(_dt(pkg, 0)) == [] and _text(_dt(pkg, 0)) == "Fixed 1 Jan"
    res = fu.set_footer(pkg, 0, date=True)
    assert res["results"][0]["date"] == "set"
    pkg = _reload(pkg)
    sp = _dt(pkg, 0)
    assert _fields(sp) == _fields(_master_dt(pkg)) and _fields(sp)[0].startswith("datetime")
    assert "Fixed 1 Jan" not in _text(sp)


def test_deck_wide_automatic_reaches_the_fixed_slide_too(make_deck):
    pkg = PptxPackage(make_deck("dt2.pptx", seed=2, extra_slides=1))
    fu.set_footer(pkg, 0, date="Fixed 1 Jan")
    fu.set_footer(pkg, None, date=True)
    pkg = _reload(pkg)
    for index in range(len(slide_table(pkg))):
        sp = _dt(pkg, index)
        assert sp is not None and _fields(sp) and "Fixed" not in _text(sp)


def test_an_automatic_date_is_left_as_it_is(make_deck):
    pkg = PptxPackage(make_deck("dt3.pptx", seed=3, extra_slides=0))
    fu.set_footer(pkg, 0, date=True)
    before = etree.tostring(_dt(pkg, 0))
    fu.set_footer(pkg, 0, date=True)
    assert etree.tostring(_dt(pkg, 0)) == before


def test_with_no_date_field_anywhere_it_writes_powerpoints_field(make_deck):
    pkg = PptxPackage(make_deck("dt4.pptx", seed=4, extra_slides=0))
    for name in [n for n in pkg.part_names() if "slideLayout" in n or "slideMaster" in n]:
        if not name.endswith(".xml") or "/_rels/" in name:
            continue
        sp = fu._find_ph(pkg.root(name), "dt")
        if sp is not None:
            fu._set_sp_text(sp, "No field")
            pkg.mark_dirty(name)
    fu.set_footer(pkg, 0, date="Fixed 1 Jan")
    fu.set_footer(pkg, 0, date=True)
    pkg = _reload(pkg)
    sp = _dt(pkg, 0)
    flds = list(sp.iter(qn("a:fld")))
    assert [f.get("type") for f in flds] == ["datetime1"]
    assert flds[0].get("id").startswith("{") and len(flds[0].get("id")) == 38
    today = datetime.date.today()
    assert _text(sp) == f"{today.month}/{today.day}/{today.year}"
