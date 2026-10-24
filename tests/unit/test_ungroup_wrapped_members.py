"""ungroup_shapes keeps every member of the group.

PowerPoint stores an equation box, ink and a newer chart (chartex) inside an
mc:AlternateContent wrapper, and inside a group that wrapper is a member like
any shape. ungroup_shapes lifted only the plain shapes and then removed the
group, so the wrapped members were deleted while the answer said ok. Now each
wrapper moves out whole, every branch rescaled by the group's transform, and a
member the tool cannot place is refused before anything changes.

Every deck here is made up by code.
"""

from __future__ import annotations

import zipfile

import pytest
from lxml import etree
from pptx import Presentation
from pptx.util import Inches

from kitchensink4ppt.core.errors import UnsupportedStructure
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import shapes as sh

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
CX2 = "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
OMML = "http://schemas.openxmlformats.org/officeDocument/2006/math"
AC = f"{{{MC}}}AlternateContent"


def _sp(sid, name, x, y, cx, cy, text):
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/>'
        f'</p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/>'
        '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody>'
        f"<a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"
    )


def _equation(sid, x, y, cx, cy, *, declare_a14=True):
    decl = f' xmlns:a14="{A14}"' if declare_a14 else ""
    math = (
        f'<a14:m><m:oMathPara xmlns:m="{OMML}"><m:oMath><m:r><m:t>y=3</m:t></m:r>'
        "</m:oMath></m:oMathPara></a14:m>"
    )
    choice = (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="Made-up equation"/><p:cNvSpPr txBox="1"/>'
        f'<p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" '
        f'cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
        f"<p:txBody><a:bodyPr/><a:lstStyle/><a:p>{math}</a:p></p:txBody></p:sp>"
    )
    fallback = _sp(sid, "Made-up equation", x, y, cx, cy, "y = 3 as a picture")
    return (
        f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice{decl} Requires="a14">{choice}'
        f"</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"
    )


def _ink(sid, x, y, cx, cy):
    return (
        f'<mc:AlternateContent xmlns:mc="{MC}" xmlns:p14="{P14}"><mc:Choice Requires="p14">'
        f'<p:contentPart p14:bwMode="auto" r:id="rIdINK"><p14:nvContentPartPr><p14:cNvPr id="{sid}" '
        'name="Made-up ink"/><p14:cNvContentPartPr/><p14:nvPr/></p14:nvContentPartPr><p14:xfrm>'
        f'<a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></p14:xfrm></p:contentPart>'
        f'</mc:Choice><mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="{sid}" name="Made-up ink"/>'
        '<p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdPIC"/><a:stretch>'
        f'<a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        "</p:spPr></p:pic></mc:Fallback></mc:AlternateContent>"
    )


def _chartex(sid, x, y, cx, cy, *, declare_cx2=True):
    decl = f' xmlns:cx2="{CX2}"' if declare_cx2 else ""
    return (
        f'<mc:AlternateContent xmlns:mc="{MC}"{decl}><mc:Choice Requires="cx2">'
        f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{sid}" name="Made-up funnel"/>'
        '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm>'
        f'<a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></p:xfrm><a:graphic>'
        '<a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">'
        '<cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" '
        'r:id="rIdCHX"/></a:graphicData></a:graphic></p:graphicFrame></mc:Choice><mc:Fallback>'
        f"{_sp(sid, 'Made-up funnel', x, y, cx, cy, 'This chart needs a newer version')}"
        "</mc:Fallback></mc:AlternateContent>"
    )


def _group(gid, members, *, off=(914400, 914400), ext=(4572000, 1828800),
           ch_off=(914400, 914400), ch_ext=(4572000, 1828800), extra_ns=""):
    return (
        f'<p:grpSp xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"{extra_ns}><p:nvGrpSpPr>'
        f'<p:cNvPr id="{gid}" name="Made-up group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        f'<p:grpSpPr><a:xfrm><a:off x="{off[0]}" y="{off[1]}"/><a:ext cx="{ext[0]}" '
        f'cy="{ext[1]}"/><a:chOff x="{ch_off[0]}" y="{ch_off[1]}"/><a:chExt cx="{ch_ext[0]}" '
        f'cy="{ch_ext[1]}"/></a:xfrm></p:grpSpPr>{"".join(members)}</p:grpSp>'
    )


RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
_PARTS = (
    # (placeholder in the group XML, relationship type, part, content type, body)
    ("rIdINK", RT + "customXml", "ppt/ink/ink1.xml", "application/inkml+xml",
     b"<inkml:ink xmlns:inkml='http://www.w3.org/2003/InkML'/>"),
    ("rIdPIC", RT + "image", "ppt/media/image1.png", "image/png", b"\x89PNG made up"),
    ("rIdCHX", "http://schemas.microsoft.com/office/2014/relationships/chartEx",
     "ppt/charts/chartEx1.xml", "application/vnd.ms-office.chartex+xml",
     b"<cx:chartSpace xmlns:cx='http://schemas.microsoft.com/office/drawing/2014/chartex'/>"),
)


def _deck(tmp_path, group_xml, name="ungroup.pptx"):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(Inches(0.2), Inches(0.2), Inches(2), Inches(0.5)).text_frame.text = "Before"
    path = tmp_path / name
    prs.save(str(path))
    pkg = PptxPackage(path)
    part = pkg.slide_parts()[0]
    for marker, rel_type, target, ctype, body in _PARTS:
        if marker in group_xml:
            pkg.add_part_with_content_type(target, body, ctype)
            rid = pkg.add_relationship(part, rel_type, "../" + target.split("/", 1)[1])
            group_xml = group_xml.replace(marker, rid)
    tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    tree.append(etree.fromstring(group_xml))
    after = etree.fromstring(
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="90" name="After"/>'
        '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="6000000"/>'
        '<a:ext cx="914400" cy="457200"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>'
        "<a:p><a:r><a:t>After</a:t></a:r></a:p></p:txBody></p:sp>"
    )
    tree.append(after)
    pkg.mark_dirty(part)
    pkg.save(do_backup=False)
    return path


def _tree(path):
    pkg = PptxPackage(path)
    return pkg.root(pkg.slide_parts()[0]).find(f"{qn('p:cSld')}/{qn('p:spTree')}")


def _top(path):
    """The top-level members as (kind, id) in stacking order."""
    out = []
    for c in _tree(path):
        if c.tag == AC:
            out.append(("wrapped", int(next(c.iter("{*}cNvPr")).get("id"))))
        elif c.tag in (qn("p:sp"), qn("p:grpSp"), qn("p:pic")):
            out.append((etree.QName(c).localname, int(next(c.iter(qn("p:cNvPr"))).get("id"))))
    return out


def _boxes(ac):
    """(x, y, cx, cy) of every branch's shape in one wrapper."""
    out = []
    for xfrm in ac.iter(qn("a:xfrm"), qn("p:xfrm"), f"{{{P14}}}xfrm"):
        off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
        out.append((int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))))
    return out


def _ungroup(path, gid):
    pkg = PptxPackage(path)
    out = sh.ungroup_shapes(pkg, 0, gid)
    pkg.save(do_backup=False)
    return out


def test_ungroup_lifts_a_wrapped_equation_with_the_plain_member(tmp_path):
    deck = _deck(tmp_path, _group(20, [
        _sp(21, "Plain member", 914400, 914400, 2286000, 914400, "plain words"),
        _equation(22, 3200400, 1828800, 2286000, 914400),
    ]))
    before = etree.tostring(_tree(deck)).count(b"AlternateContent")
    out = _ungroup(deck, 20)
    assert _top(deck)[1:] == [("sp", 21), ("wrapped", 22), ("sp", 90)]
    assert sorted(out["freed_ids"]) == [21, 22]
    # Both wrappers survive: the shape-level one and the math inside it.
    assert etree.tostring(_tree(deck)).count(b"AlternateContent") == before
    ac = [c for c in _tree(deck) if c.tag == AC][0]
    assert ac.find(f"{{{MC}}}Choice").get("Requires") == "a14"
    assert len(list(ac.iter(f"{{{OMML}}}t"))) == 1


def test_ungroup_rescales_every_branch_of_a_wrapped_member(tmp_path):
    # The group draws its children at half size, moved right by one inch.
    deck = _deck(tmp_path, _group(
        20,
        [_sp(21, "Plain member", 0, 0, 2000000, 1000000, "plain words"),
         _equation(22, 2000000, 1000000, 2000000, 1000000)],
        off=(914400, 0), ext=(2000000, 1000000), ch_off=(0, 0), ch_ext=(4000000, 2000000),
    ))
    _ungroup(deck, 20)
    tree = _tree(deck)
    plain = [c for c in tree if c.tag == qn("p:sp") and c.find(f".//{qn('p:cNvPr')}").get("id") == "21"][0]
    ac = [c for c in tree if c.tag == AC][0]
    assert _boxes(plain) == [(914400, 0, 1000000, 500000)]
    assert _boxes(ac) == [(1914400, 500000, 1000000, 500000)] * 2


def test_ungroup_keeps_ink_and_a_newer_chart(tmp_path):
    deck = _deck(tmp_path, _group(30, [
        _ink(31, 914400, 914400, 1000000, 500000),
        _chartex(32, 2000000, 914400, 1500000, 1000000),
        _equation(33, 914400, 2000000, 1000000, 500000),
    ]))
    out = _ungroup(deck, 30)
    assert _top(deck)[1:] == [("wrapped", 31), ("wrapped", 32), ("wrapped", 33), ("sp", 90)]
    assert sorted(out["freed_ids"]) == [31, 32, 33]
    ink = [c for c in _tree(deck) if c.tag == AC][0]
    rid = ink.find(f".//{qn('p:contentPart')}").get(qn("r:id"))
    pkg = PptxPackage(deck)
    assert pkg.relationship_target(pkg.slide_parts()[0], rid) == "ppt/ink/ink1.xml"
    assert _boxes(ink) == [(914400, 914400, 1000000, 500000)] * 2


def test_ungroup_keeps_a_requires_prefix_the_group_declared(tmp_path):
    # Another producer declared the chart's version namespace on the group
    # itself. The wrapper names it only in Requires, which no XML library
    # tracks, so the move must declare it again.
    deck = _deck(tmp_path, _group(
        20,
        [_sp(21, "Plain member", 914400, 914400, 2286000, 914400, "plain words"),
         _chartex(22, 3200400, 1828800, 2286000, 914400, declare_cx2=False)],
        extra_ns=f' xmlns:cx2="{CX2}"',
    ))
    _ungroup(deck, 20)
    with zipfile.ZipFile(deck) as z:
        saved = etree.fromstring(z.read("ppt/slides/slide1.xml"))
    tree = saved.find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    choice = [c for c in tree if c.tag == AC][0].find(f"{{{MC}}}Choice")
    assert choice.get("Requires") == "cx2"
    assert choice.nsmap.get("cx2") == CX2


def test_ungroup_refuses_a_member_it_cannot_place(tmp_path):
    odd = '<x:thing xmlns:x="http://example.com/made-up/2040"/>'
    deck = _deck(tmp_path, _group(
        20,
        [_sp(21, "Plain member", 0, 0, 2000000, 1000000, "plain words"), odd],
        off=(914400, 0), ext=(2000000, 1000000), ch_off=(0, 0), ch_ext=(4000000, 2000000),
    ))
    before = deck.read_bytes()
    with pytest.raises(UnsupportedStructure) as refused:
        _ungroup(deck, 20)
    message = str(refused.value)
    assert "group 20" in message and "PowerPoint" in message and "nothing" in message
    assert deck.read_bytes() == before


def test_ungroup_moves_an_unknown_member_whole_when_the_group_does_not_scale(tmp_path):
    odd = '<x:thing xmlns:x="http://example.com/made-up/2040"/>'
    deck = _deck(tmp_path, _group(20, [
        _sp(21, "Plain member", 914400, 914400, 2286000, 914400, "plain words"), odd,
    ]))
    _ungroup(deck, 20)
    tree = _tree(deck)
    assert tree.find("{http://example.com/made-up/2040}thing") is not None
    assert tree.find(qn("p:grpSp")) is None
