"""duplicate_slide gives the copy its own chartex chart, ActiveX control, ink,
web add-in, slide-library record and VML drawing.

Only charts, OLE objects, embedded packages, tags, notes and SmartArt were
copied; every other part the slide related to stayed shared between the
slide and its copy, so a change to one (an ActiveX control's state, a chartex
chart's data) showed on both, and deleting either slide could take a part the
other still names. Now each of these is copied with what it owns (a chartex
chart's workbook, style and colors; an ActiveX control's binary), while the
pictures they use stay shared, as media are. Every part here is made up.
"""

from __future__ import annotations

from kitchensink4ppt.core.package import PptxPackage, rels_name, resolve_target
from kitchensink4ppt.ops import slides

RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
MS = "http://schemas.microsoft.com/office/"
OWNED = {
    "chartex": (MS + "2014/relationships/chartEx", "ppt/charts/chartEx1.xml",
                "application/vnd.ms-office.chartex+xml"),
    "activex": (RT + "control", "ppt/activeX/activeX1.xml",
                "application/vnd.ms-office.activeX+xml"),
    "ink": (RT + "customXml", "ppt/ink/ink1.xml", "application/inkml+xml"),
    "webext": (MS + "2011/relationships/webextension", "ppt/webextensions/webextension1.xml",
               "application/vnd.ms-office.webextension+xml"),
    "update": (RT + "slideUpdateInfo", "ppt/slideUpdateInfo/slideUpdateInfo1.xml",
               "application/vnd.openxmlformats-officedocument.presentationml.slideUpdateInfo+xml"),
    "vml": (RT + "vmlDrawing", "ppt/drawings/vmlDrawing1.vml",
            "application/vnd.openxmlformats-officedocument.vmlDrawing"),
}
CHILDREN = {
    "ppt/charts/chartEx1.xml": [
        (RT + "package", "ppt/embeddings/Made_Up_Book1.xlsx",
         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", b"PK made up"),
        (MS + "2011/relationships/chartStyle", "ppt/charts/style1.xml",
         "application/vnd.ms-office.chartstyle+xml", b'<cs:chartStyle xmlns:cs="urn:made-up"/>'),
        (MS + "2011/relationships/chartColorStyle", "ppt/charts/colors1.xml",
         "application/vnd.ms-office.chartcolorstyle+xml", b'<cs:colorStyle xmlns:cs="urn:made-up"/>'),
    ],
    "ppt/activeX/activeX1.xml": [
        (MS + "2006/relationships/activeXControlBinary", "ppt/activeX/activeX1.bin",
         "application/vnd.ms-office.activeX", b"made-up control state"),
    ],
    "ppt/drawings/vmlDrawing1.vml": [
        (RT + "image", "ppt/media/image1.png", "image/png", b"\x89PNG made up"),
    ],
}


def _deck(tmp_path):
    path = tmp_path / "owned.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)
    slide = pkg.slide_parts()[0]
    for kind, (rel_type, part, ctype) in OWNED.items():
        body = b"<x:made-up xmlns:x='urn:made-up'/>" if part.endswith(".xml") else (
            b"<xml xmlns:v='urn:schemas-microsoft-com:vml'/>")
        pkg.add_part_with_content_type(part, body, ctype)
        pkg.add_relationship(slide, rel_type, "../" + part.split("/", 1)[1])
        for child_type, child, child_ct, child_body in CHILDREN.get(part, []):
            pkg.add_part_with_content_type(child, child_body, child_ct)
            pkg.add_relationship(part, child_type, _relpath(part, child))
    pkg.save(do_backup=False)
    return path


def _relpath(src, dst):
    import posixpath
    return posixpath.relpath(dst, posixpath.dirname(src))


def _ctype(pkg, part):
    """The content type [Content_Types].xml gives a part: its Override, else
    the Default for its extension."""
    root = pkg.root("[Content_Types].xml")
    for node in root:
        if node.get("PartName") == "/" + part:
            return node.get("ContentType")
    ext = part.rsplit(".", 1)[-1].lower()
    for node in root:
        if (node.get("Extension") or "").lower() == ext:
            return node.get("ContentType")
    return None


def _targets(pkg, part):
    out = {}
    for rel in pkg.root(rels_name(part)):
        if rel.get("TargetMode") != "External":
            out[rel.get("Type")] = resolve_target(part, rel.get("Target"))
    return out


def test_the_copy_owns_its_parts(tmp_path):
    path = _deck(tmp_path)
    pkg = PptxPackage(path)
    slides.duplicate_slide(pkg, 0)
    pkg.save(do_backup=False)
    pkg = PptxPackage(path)
    first, copy_ = pkg.slide_parts()[:2]
    src, dst = _targets(pkg, first), _targets(pkg, copy_)
    for kind, (rel_type, part, ctype) in OWNED.items():
        assert src[rel_type] == part, kind
        assert dst[rel_type] != part, f"{kind} is still shared"
        assert pkg.part_bytes(dst[rel_type]) == pkg.part_bytes(part), kind
        assert _ctype(pkg, dst[rel_type]) == ctype, kind


def test_what_an_owned_part_owns_is_copied_and_pictures_stay_shared(tmp_path):
    path = _deck(tmp_path)
    pkg = PptxPackage(path)
    slides.duplicate_slide(pkg, 0)
    pkg.save(do_backup=False)
    pkg = PptxPackage(path)
    copy_ = pkg.slide_parts()[1]
    dst = _targets(pkg, copy_)
    chartex = _targets(pkg, dst[OWNED["chartex"][0]])
    for child_type, child, _ct, body in CHILDREN["ppt/charts/chartEx1.xml"]:
        assert chartex[child_type] != child and pkg.part_bytes(chartex[child_type]) == body
    activex = _targets(pkg, dst[OWNED["activex"][0]])
    (bin_type, bin_part, _ct, bin_body), = CHILDREN["ppt/activeX/activeX1.xml"]
    assert activex[bin_type] != bin_part and pkg.part_bytes(activex[bin_type]) == bin_body
    vml = _targets(pkg, dst[OWNED["vml"][0]])
    assert vml[RT + "image"] == "ppt/media/image1.png"
