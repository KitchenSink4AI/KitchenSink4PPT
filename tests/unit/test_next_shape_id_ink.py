"""A new shape never takes the id of an ink content part.

PowerPoint stores ink as p14:contentPart inside mc:AlternateContent, with its
id on p14:cNvPr. next_shape_id counted only p:cNvPr, so on a slide whose ink
has no picture fallback carrying the same id, the next shape could be given
the ink's id. The deck is made up and built by code.
"""

from __future__ import annotations

from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import shapes as shp
from kitchensink4ppt.ops.read import get_slide_info

P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def test_an_ink_part_id_is_not_given_to_a_new_shape(make_deck):
    pkg = PptxPackage(make_deck("ink.pptx", seed=1, extra_slides=0))
    part = get_slide_info(pkg, 0)["part"]
    tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    ink_id = pkg.next_shape_id(part) + 3
    ac = etree.SubElement(tree, f"{{{MC}}}AlternateContent", nsmap={"mc": MC})
    choice = etree.SubElement(ac, f"{{{MC}}}Choice", Requires="p14")
    cp = etree.SubElement(choice, f"{{{P14}}}contentPart", nsmap={"p14": P14})
    nv = etree.SubElement(cp, f"{{{P14}}}nvContentPartPr")
    etree.SubElement(nv, f"{{{P14}}}cNvPr", id=str(ink_id), name="Ink 1")
    assert pkg.next_shape_id(part) == ink_id + 1
    new = shp.insert_shape(pkg, 0, "rect", 1, 1, 1, 1)["shape_id"]
    assert new == ink_id + 1
