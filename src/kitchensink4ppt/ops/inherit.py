"""Placeholder inheritance: what a slide shape gets from its layout, its
master, and the theme when it says nothing itself.

Three tools were blind in the same place and for the same reason. A
placeholder that carries an empty `<p:spPr/>` has no box of its own, so
`set_shape` refused to move it, `check_layout`'s overlap check could not
see it, and its runs have no `sz`, so the tiny_text check skipped exactly
the runs most likely to overflow: the ones rendering at the master's body
size on a templated deck. In each case the information was one or two hops
away in the layout or the master, and nobody made the hop.

Two resolutions live here, both read-only and both returning None rather
than a guess when the chain runs out:

- GEOMETRY. The slide placeholder's a:xfrm, else the matching layout
  placeholder's, else the master's.
- FONT SIZE. The run's own sz, else the paragraph's defRPr, else the
  shape's own lstStyle at that level, else the layout placeholder's
  lstStyle, else the master's txStyles (titleStyle for title-family
  placeholders, bodyStyle for body-family, otherStyle for everything
  else), else the presentation's defaultTextStyle.

MATCHING. PowerPoint matches a slide placeholder to its layout twin by
`idx` when one is present, and by type otherwise; the master has at most
one placeholder per family, so layout-to-master matching is by family
alone. `ctrTitle` inherits from the master's title, and the whole body
family (`body`, `obj`, `subTitle`, and a bare idx with no type) inherits
from the master's body.
"""

from __future__ import annotations

from lxml import etree

from ..core.package import PptxPackage, qn, resolve_target

_RT_SLIDE_LAYOUT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/"
    "relationships/slideLayout"
)
_RT_SLIDE_MASTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/"
    "relationships/slideMaster"
)

#: Placeholder type -> the master text style that backs it.
TITLE_TYPES = frozenset({"title", "ctrTitle"})
BODY_TYPES = frozenset({"body", "obj", "subTitle"})


def related(pkg: PptxPackage, part: str | None, rel_type: str) -> str | None:
    """The first related part of `rel_type`, or None."""
    if not part:
        return None
    try:
        rels = pkg.rels_for(part)
    except KeyError:
        return None
    for rel in rels.getroot():
        if rel.get("Type") == rel_type and rel.get("TargetMode") != "External":
            return resolve_target(part, rel.get("Target", ""))
    return None


def layout_part_of(pkg: PptxPackage, slide_part: str) -> str | None:
    return related(pkg, slide_part, _RT_SLIDE_LAYOUT)


def master_part_of(pkg: PptxPackage, layout_part: str | None) -> str | None:
    return related(pkg, layout_part, _RT_SLIDE_MASTER)


def placeholder_of(elem: etree._Element) -> etree._Element | None:
    """The p:ph element of a shape, or None when it is not a placeholder."""
    for child in elem:
        if etree.QName(child).localname.startswith("nv"):
            nvpr = child.find(qn("p:nvPr"))
            return nvpr.find(qn("p:ph")) if nvpr is not None else None
    return None


def placeholder_key(ph: etree._Element | None) -> tuple[str, str | None] | None:
    """(family, idx) for a p:ph, or None. A p:ph with an idx and no type is
    a body placeholder, which is what the schema default says and what the
    layouts this server writes rely on."""
    if ph is None:
        return None
    ph_type = ph.get("type") or "body"
    return ph_type, ph.get("idx")


def family_of(ph_type: str) -> str:
    """Which master text style and which master placeholder back a type."""
    if ph_type in TITLE_TYPES:
        return "title"
    if ph_type in BODY_TYPES:
        return "body"
    return "other"


def _placeholders(pkg: PptxPackage, part: str | None):
    if not part or not pkg.has_part(part):
        return []
    tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if tree is None:
        return []
    out = []
    for sp in tree.iter(qn("p:sp")):
        ph = placeholder_of(sp)
        if ph is not None:
            out.append((sp, ph))
    return out


def layout_twin(
    pkg: PptxPackage, layout_part: str | None, key: tuple[str, str | None]
) -> etree._Element | None:
    """The layout placeholder a slide placeholder inherits from: same idx
    when the slide names one, else same type, else same family."""
    ph_type, idx = key
    candidates = _placeholders(pkg, layout_part)
    if idx is not None:
        for sp, ph in candidates:
            if ph.get("idx") == idx:
                return sp
    for sp, ph in candidates:
        if (ph.get("type") or "body") == ph_type:
            return sp
    family = family_of(ph_type)
    for sp, ph in candidates:
        if family_of(ph.get("type") or "body") == family and family != "other":
            return sp
    return None


def master_twin(
    pkg: PptxPackage, master_part: str | None, key: tuple[str, str | None]
) -> etree._Element | None:
    """The master placeholder backing a family. The master carries at most
    one of each, so this matches by family, not by idx."""
    family = family_of(key[0])
    for sp, ph in _placeholders(pkg, master_part):
        if family_of(ph.get("type") or "body") == family:
            return sp
    return None


def _xfrm_of(sp: etree._Element) -> etree._Element | None:
    sppr = sp.find(qn("p:spPr"))
    return sppr.find(qn("a:xfrm")) if sppr is not None else None


def inherited_xfrm(
    pkg: PptxPackage, slide_part: str, elem: etree._Element
) -> etree._Element | None:
    """The a:xfrm a slide placeholder RENDERS with when it carries none of
    its own: the layout twin's, else the master twin's. None when the shape
    is not a placeholder or nothing up the chain has a box either."""
    key = placeholder_key(placeholder_of(elem))
    if key is None:
        return None
    layout_part = layout_part_of(pkg, slide_part)
    twin = layout_twin(pkg, layout_part, key)
    if twin is not None:
        xfrm = _xfrm_of(twin)
        if xfrm is not None:
            return xfrm
    master = master_twin(pkg, master_part_of(pkg, layout_part), key)
    if master is not None:
        return _xfrm_of(master)
    return None


def inherited_box(
    pkg: PptxPackage, slide_part: str, elem: etree._Element
) -> tuple[int, int, int, int] | None:
    """(x, y, cx, cy) in EMU from the inheritance chain, or None."""
    xfrm = inherited_xfrm(pkg, slide_part, elem)
    if xfrm is None:
        return None
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        return None
    try:
        return (
            int(off.get("x")), int(off.get("y")),
            int(ext.get("cx")), int(ext.get("cy")),
        )
    except (TypeError, ValueError):
        return None


# ------------------------------------------------------------- font size


def _lvl_size(container: etree._Element | None, level: int) -> float | None:
    """sz (in points) from a lstStyle-like container at one outline level."""
    if container is None:
        return None
    lvl = container.find(qn(f"a:lvl{min(max(level, 0), 8) + 1}pPr"))
    if lvl is None:
        return None
    defrpr = lvl.find(qn("a:defRPr"))
    if defrpr is None:
        return None
    return _sz_pt(defrpr.get("sz"))


def _sz_pt(raw) -> float | None:
    if not raw:
        return None
    try:
        return int(raw) / 100.0
    except (TypeError, ValueError):
        return None


#: Master text style part names, by placeholder family.
_TX_STYLE = {
    "title": "p:titleStyle",
    "body": "p:bodyStyle",
    "other": "p:otherStyle",
}


def resolve_run_size_pt(
    pkg: PptxPackage,
    slide_part: str,
    elem: etree._Element,
    paragraph: etree._Element,
    rpr: etree._Element | None,
    level: int = 0,
) -> tuple[float | None, str]:
    """(size in points, where it came from) for one run.

    The source string is part of the answer, not decoration: a check that
    acts on an inherited size should be able to say it was inherited, and a
    chain that runs dry must report "unresolved" rather than a default
    nobody can verify.
    """
    explicit = _sz_pt(rpr.get("sz")) if rpr is not None else None
    if explicit is not None:
        return explicit, "run"

    ppr = paragraph.find(qn("a:pPr"))
    if ppr is not None:
        # lxml elements are falsy when childless, so every lookup here is
        # an explicit `is not None` test.
        defrpr = ppr.find(qn("a:defRPr"))
        size = _sz_pt(defrpr.get("sz")) if defrpr is not None else None
        if size is not None:
            return size, "paragraph"
        try:
            level = int(ppr.get("lvl") or level)
        except (TypeError, ValueError):
            pass

    body = elem.find(qn("p:txBody"))
    own = body.find(qn("a:lstStyle")) if body is not None else None
    size = _lvl_size(own, level)
    if size is not None:
        return size, "shape"

    key = placeholder_key(placeholder_of(elem))
    if key is None:
        return None, "unresolved"

    layout_part = layout_part_of(pkg, slide_part)
    twin = layout_twin(pkg, layout_part, key)
    if twin is not None:
        tbody = twin.find(qn("p:txBody"))
        size = _lvl_size(
            tbody.find(qn("a:lstStyle")) if tbody is not None else None, level
        )
        if size is not None:
            return size, "layout"

    master_part = master_part_of(pkg, layout_part)
    mtwin = master_twin(pkg, master_part, key)
    if mtwin is not None:
        mbody = mtwin.find(qn("p:txBody"))
        size = _lvl_size(
            mbody.find(qn("a:lstStyle")) if mbody is not None else None, level
        )
        if size is not None:
            return size, "master placeholder"

    if master_part and pkg.has_part(master_part):
        styles = pkg.root(master_part).find(qn("p:txStyles"))
        if styles is not None:
            size = _lvl_size(
                styles.find(qn(_TX_STYLE[family_of(key[0])])), level
            )
            if size is not None:
                return size, "master"

    try:
        default = pkg.presentation().find(qn("p:defaultTextStyle"))
    except Exception:
        default = None
    size = _lvl_size(default, level)
    if size is not None:
        return size, "presentation default"
    return None, "unresolved"
