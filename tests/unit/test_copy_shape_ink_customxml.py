"""copy_shape_between refuses ink however the slide relates it.

PowerPoint relates an ink part to its slide with the customXml relationship
type (the published drawing extensions say so); the operation's table of
refused types names an older ink type instead. The refusal never depended on
that table: a shape holding p:contentPart is refused as ink by its element,
and an internal relationship of a type the operation does not know is
refused too. This pins the customXml form. Every deck here is made up.
"""

from __future__ import annotations

from lxml import etree

from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import shapes as shp

from test_copy_shape_between import _NS, _deck, _insert, _ok, _refused, dest  # noqa: F401


def test_refuses_ink_related_as_custom_xml(tmp_path, dest):  # noqa: F811
    deck = _deck(tmp_path / "ink.pptx")
    a = _insert(deck, "A", 1, 1)
    b = _insert(deck, "B", 3, 1)
    _ok("ppt_graphics", "group_shapes", {
        "file_path": deck, "slide": 0, "ids": [a, b], "name": "Inked"})
    pkg = PptxPackage(deck)
    part = pkg.slide_parts()[0]
    pkg.add_part_with_content_type(
        "ppt/ink/ink1.xml", b'<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"/>',
        "application/inkml+xml")
    rid = pkg.add_relationship(
        part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
        "../ink/ink1.xml")
    grp, _c = shp._find_shape(pkg, part, "Inked")
    grp.append(etree.fromstring(f'<p:contentPart {_NS} r:id="{rid}"/>'))
    pkg.mark_dirty(part)
    pkg.save(do_backup=False)
    message = _refused({"file_path": dest, "source_path": deck, "source_slide": 0,
                        "shape": "Inked", "slide": 1})
    assert "ink" in message
