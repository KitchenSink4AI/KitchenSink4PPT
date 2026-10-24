"""Text sizes and line widths that follow a resize.

A group draws its members through its transform, so their geometry scales
when the group is resized. Font sizes (a:rPr sz), text insets, paragraph
spacing in points and line widths (a:ln w) are absolute and do not: a figure
placed at a third of its size keeps full-size labels, which wrap, and
full-weight lines. These helpers scale them by one factor across a shape or
a whole group, when the caller asks.

A run that inherits its size is given the size it was drawn at, scaled, so
it shrinks with the rest; a line drawn from the theme's line styles is given
the theme's width, scaled. Tables and charts inside a group are left as they
are and named in the answer.
"""

from __future__ import annotations

from lxml import etree

from ..core.package import PptxPackage, qn
from ..core.xmlbool import attr_bool

#: ST_TextFontSize bounds, in hundredths of a point.
_SZ_MIN, _SZ_MAX = 100, 400000
#: a:bodyPr insets and their defaults when absent, in EMU.
_INSETS = (("lIns", 91440), ("tIns", 45720), ("rIns", 91440), ("bIns", 45720))
_SHAPES = ("p:sp", "p:cxnSp", "p:pic", "p:grpSp", "p:graphicFrame")


def _members(elem: etree._Element):
    """elem and every shape inside it, depth first."""
    yield elem
    if elem.tag == qn("p:grpSp"):
        for child in elem:
            if child.tag in tuple(qn(t) for t in _SHAPES):
                yield from _members(child)


def _shape_id(elem: etree._Element) -> int | None:
    for child in elem:
        if etree.QName(child).localname.startswith("nv"):
            cnvpr = child.find(qn("p:cNvPr"))
            if cnvpr is not None and (cnvpr.get("id") or "").isdigit():
                return int(cnvpr.get("id"))
    return None


def _scaled_int(raw: str | None, factor: float, low: int | None = None,
                high: int | None = None) -> str | None:
    if raw is None:
        return None
    try:
        value = round(int(raw) * factor)
    except ValueError:
        return None
    if low is not None:
        value = max(low, value)
    if high is not None:
        value = min(high, value)
    return str(value)


def _scale_attr(el, name: str, factor: float, low=None, high=None) -> bool:
    if el is None or el.get(name) is None:
        return False
    new = _scaled_int(el.get(name), factor, low, high)
    if new is None:
        return False
    el.set(name, new)
    return True


def _scale_ppr(ppr, factor: float) -> None:
    """Point-based spacing, bullet size, margins and the default run size
    of one a:pPr or a:lvlNpPr."""
    if ppr is None:
        return
    for name in ("marL", "indent"):
        _scale_attr(ppr, name, factor)
    for tag in ("a:lnSpc", "a:spcBef", "a:spcAft"):
        holder = ppr.find(qn(tag))
        if holder is not None:
            _scale_attr(holder.find(qn("a:spcPts")), "val", factor, 0, 158400)
    _scale_attr(ppr.find(qn("a:buSzPts")), "val", factor, 100, 400000)
    _scale_attr(ppr.find(qn("a:defRPr")), "sz", factor, _SZ_MIN, _SZ_MAX)
    for tab in ppr.iter(qn("a:tab")):
        _scale_attr(tab, "pos", factor)


def _scale_body(pkg, part, sp: etree._Element, body: etree._Element,
                factor: float) -> tuple[int, int]:
    """(sizes written, sizes that could not be established) for one text
    body. Insets are written out so they shrink with the box."""
    from . import text as _tx

    written = unknown = 0
    bodypr = body.find(qn("a:bodyPr"))
    if bodypr is not None:
        for name, default in _INSETS:
            value = int(bodypr.get(name, default))
            bodypr.set(name, str(round(value * factor)))
    lst = body.find(qn("a:lstStyle"))
    if lst is not None:
        for lvl in lst:
            _scale_ppr(lvl, factor)
    for p in body.iter(qn("a:p")):
        level = _tx._para_level(p)
        _scale_ppr(p.find(qn("a:pPr")), factor)
        holders = [child.find(qn("a:rPr")) if child.tag != qn("a:endParaRPr") else child
                   for child in p if child.tag in (qn("a:r"), qn("a:fld"), qn("a:br"),
                                                   qn("a:endParaRPr"))]
        runs = [child for child in p if child.tag in (qn("a:r"), qn("a:fld"), qn("a:br"),
                                                     qn("a:endParaRPr"))]
        for child, rpr in zip(runs, holders):
            if rpr is not None and rpr.get("sz") is not None:
                _scale_attr(rpr, "sz", factor, _SZ_MIN, _SZ_MAX)
                written += 1
                continue
            size = _tx._certain_size_pt(rpr, p, body, sp, level, pkg, part)
            if size is None:
                unknown += 1
                continue
            if rpr is None:
                rpr = etree.Element(qn("a:rPr"))
                child.insert(0, rpr)
            value = max(_SZ_MIN, min(_SZ_MAX, round(size * 100 * factor)))
            rpr.set("sz", str(value))
            written += 1
    return written, unknown


def scale_text(pkg: PptxPackage, part: str, elem: etree._Element,
               factor: float) -> dict:
    """Scale every text size, inset and point spacing in elem (a shape or
    a group) by factor."""
    sizes = unknown = 0
    skipped: list[int] = []
    for member in _members(elem):
        if member.tag == qn("p:graphicFrame"):
            skipped.append(_shape_id(member))
            continue
        body = member.find(qn("p:txBody"))
        if body is None:
            continue
        written, missing = _scale_body(pkg, part, member, body, factor)
        sizes += written
        unknown += missing
    out: dict = {"factor": round(factor, 4), "sizes": sizes}
    if unknown:
        out["sizes_not_established"] = unknown
    if skipped:
        out["tables_and_charts_not_scaled"] = [i for i in skipped if i is not None]
    return out


def _theme_line_widths(pkg: PptxPackage, part: str) -> list[int | None]:
    """The theme's line style widths (a:lnStyleLst), in order."""
    try:
        from .design import _theme_part_of
        from .inherit import layout_part_of, master_part_of

        master = master_part_of(pkg, layout_part_of(pkg, part))
        theme = _theme_part_of(pkg, master) if master else None
        if not theme or not pkg.has_part(theme):
            return []
        lst = pkg.root(theme).find(
            f"{qn('a:themeElements')}/{qn('a:fmtScheme')}/{qn('a:lnStyleLst')}")
    except Exception:
        return []
    if lst is None:
        return []
    widths = []
    for ln in lst.findall(qn("a:ln")):
        raw = ln.get("w")
        widths.append(int(raw) if raw and raw.isdigit() else None)
    return widths


def scale_lines(pkg: PptxPackage, part: str, elem: etree._Element,
                factor: float) -> dict:
    """Scale every outline width in elem by factor. An outline drawn from
    the theme's line styles is written out at the theme width, scaled."""
    from . import geometry as g

    theme = None
    scaled: list[int] = []
    left: list[int] = []
    for member in _members(elem):
        if member.tag in (qn("p:grpSp"), qn("p:graphicFrame")):
            continue
        sppr = member.find(qn("p:spPr"))
        if sppr is None:
            continue
        ln = sppr.find(qn("a:ln"))
        if ln is not None and ln.get("w") is not None:
            _scale_attr(ln, "w", factor, 0, 20116800)
            scaled.append(_shape_id(member))
            continue
        if ln is not None and ln.find(qn("a:noFill")) is not None:
            continue
        ref = member.find(f"{qn('p:style')}/{qn('a:lnRef')}")
        idx = int(ref.get("idx", "0")) if ref is not None and (ref.get("idx") or "0").isdigit() else 0
        if idx <= 0:
            if ln is not None:
                left.append(_shape_id(member))
            continue
        if theme is None:
            theme = _theme_line_widths(pkg, part)
        width = theme[idx - 1] if 0 < idx <= len(theme) else None
        if width is None:
            left.append(_shape_id(member))
            continue
        if ln is None:
            ln = etree.Element(qn("a:ln"))
            g.insert_spPr_child(sppr, ln)
            ln = sppr.find(qn("a:ln"))
        ln.set("w", str(max(0, round(width * factor))))
        scaled.append(_shape_id(member))
    out: dict = {"factor": round(factor, 4), "shapes": [i for i in scaled if i is not None]}
    if left:
        out["widths_not_established"] = [i for i in left if i is not None]
    return out


def unwrap_text_boxes(elem: etree._Element) -> list[int]:
    """Set every text box in elem (a shape or a group) to keep each
    paragraph on one line (bodyPr wrap="none"). Shapes that are not text
    boxes keep their wrapping."""
    changed: list[int] = []
    for member in _members(elem):
        if member.tag != qn("p:sp"):
            continue
        nv = member.find(f"{qn('p:nvSpPr')}/{qn('p:cNvSpPr')}")
        if not attr_bool(nv, "txBox"):
            continue
        bodypr = member.find(f"{qn('p:txBody')}/{qn('a:bodyPr')}")
        if bodypr is None or bodypr.get("wrap") == "none":
            continue
        bodypr.set("wrap", "none")
        sid = _shape_id(member)
        if sid is not None:
            changed.append(sid)
    return changed


__all__ = ["scale_lines", "scale_text", "unwrap_text_boxes"]
