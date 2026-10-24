"""Copy one shape or group from a deck onto a slide (copy_shape_between).

A reused figure is usually one group: bands, labels, connectors, a picture.
Copying it by hand means a deep copy of its XML with new shape ids, every
relationship it uses recreated on the receiving slide, the group frame
rewritten for the new place, and the members that do not belong removed.
This module does that for one shape or group:

- The source deck is read from disk and never changed (asserted), so the
  source may be the deck being edited: a copy within one deck works.
- Every shape in the copy gets a fresh id on the receiving slide; connector
  glue follows the copied shapes, and glue to a shape left behind is dropped
  with a warning (the end keeps its place).
- Every relationship the copy uses is recreated on the receiving slide:
  pictures and media through the content-deduplicated media pool, charts
  with their workbook and style parts, tags, external links. A jump link to
  a slide keeps its slide within one deck and is refused across decks.
- Between designs (another deck, or another master in the same deck) the
  copy keeps its look: theme colours and +mj/+mn theme fonts are written as
  literals resolved against the SOURCE slide's theme, inherited text sizes
  are written out, and outlines drawn from the theme's line styles get the
  source theme's width. Within one design the copy stays linked to the
  theme.
- leave_out drops members before placing, and the groups that held them are
  refitted to what is left, so nothing moves.
- x, y default to where the shape is drawn on the source slide (through its
  groups); w or h alone keeps the aspect; scale_text, scale_lines and
  no_wrap work as they do on set_shape.

Refused, each with the reason: placeholders (their place and look come from
the layout), SmartArt, embedded objects, ink, a jump link or custom show of
another deck, relationship types this module does not know, a member of a
rotated or flipped group (its drawn place cannot be computed here), and a
new size for a table (column widths and row heights decide a table's size).
"""

from __future__ import annotations

import copy as _copy
from pathlib import Path

from lxml import etree

from ..core import safesave
from ..core.errors import (
    AmbiguousTarget,
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from ..core.package import NSMAP, RT_SLIDE, PptxPackage, qn, rels_name, resolve_target
from ..core.sandbox import check_path
from ..core.swapwait import require_deck
from . import geometry as g
from . import scaling as _scaling
from . import shapes as _sh
from .interdeck import (
    RT_SLIDE_MASTER,
    _bake_theme_refs,
    _import_chart,
    _import_leaf,
    _rel_target_of_type,
    _source_layout_of,
    _theme_bake_context,
)
from .read import resolve_slide
from .slides import _regenerate_creation_ids, _rel_target, _resolve_slide

_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
_MS = "http://schemas.microsoft.com/office/"

#: Internal relationship types a copied shape may use, by how they travel.
_MEDIA = frozenset({
    _RT + "image", _RT + "audio", _RT + "video",
    _MS + "2007/relationships/media", _MS + "2007/relationships/hdphoto",
})
_CHARTS = frozenset({_RT + "chart", _MS + "2014/relationships/chartEx"})
_LEAVES = frozenset({_RT + "tags"})
#: External relationship types a copied shape may use: links and linked media.
_EXTERNAL = _MEDIA | {_RT + "hyperlink"}

#: Internal relationship types refused, with the words the refusal uses.
_REFUSED_TYPES = {
    _RT + "diagramData": "SmartArt",
    _RT + "diagramLayout": "SmartArt",
    _RT + "diagramQuickStyle": "SmartArt",
    _RT + "diagramColors": "SmartArt",
    _MS + "2007/relationships/diagramDrawing": "SmartArt",
    _RT + "oleObject": "an embedded object",
    _RT + "package": "an embedded object",
    _RT + "control": "an ActiveX control",
    _MS + "2011/relationships/ink": "ink",
}

_URI_DIAGRAM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
_URI_OLE = "http://schemas.openxmlformats.org/presentationml/2006/ole"
_URI_TABLE = "http://schemas.openxmlformats.org/drawingml/2006/table"

_R_NS = "{" + NSMAP["r"] + "}"


def _what(elem: etree._Element) -> str:
    return f"shape {_sh._shape_id(elem)} ({_sh._shape_name(elem)!r})"


# ------------------------------------------------------------- refusals


def _refuse_structure(elem: etree._Element, cross_deck: bool) -> None:
    """Refuse what cannot travel as a shape, naming the shape and the
    reason, before anything is changed."""
    for node in elem.iter():
        if not isinstance(node.tag, str):
            continue
        local = etree.QName(node).localname
        if node.tag == qn("p:ph"):
            raise UnsupportedStructure(
                f"{_what(elem)} is or holds a placeholder, which takes its place "
                "and look from the slide layout and does not travel as a shape; "
                "copy the slide with copy_slide_between, or set the text with "
                "set_placeholder_text"
            )
        if local == "contentPart":
            raise UnsupportedStructure(
                f"{_what(elem)} holds ink, which this operation does not copy; "
                "copy the slide with copy_slide_between instead"
            )
        if node.tag == qn("a:graphicData"):
            uri = node.get("uri", "")
            if uri == _URI_DIAGRAM:
                raise UnsupportedStructure(
                    f"{_what(elem)} is SmartArt, whose parts this operation does "
                    "not copy; copy the slide with copy_slide_between instead"
                )
            if uri == _URI_OLE:
                raise UnsupportedStructure(
                    f"{_what(elem)} is an embedded object, which this operation "
                    "does not copy; copy the slide with copy_slide_between instead"
                )
        if local == "oleObj":
            raise UnsupportedStructure(
                f"{_what(elem)} holds an embedded object, which this operation "
                "does not copy; copy the slide with copy_slide_between instead"
            )
        if cross_deck and node.tag in (qn("a:hlinkClick"), qn("a:hlinkHover"),
                                       qn("a:hlinkMouseOver")):
            if (node.get("action") or "").startswith("ppaction://customshow"):
                raise UnsupportedStructure(
                    f"{_what(elem)} has a link that starts a custom show of the "
                    "source deck, which does not exist in this deck; remove the "
                    "link (remove_hyperlink) or copy within the source deck"
                )


def _is_table(elem: etree._Element) -> bool:
    data = elem.find(f"{qn('a:graphic')}/{qn('a:graphicData')}")
    return elem.tag == qn("p:graphicFrame") and data is not None and (
        data.get("uri") == _URI_TABLE)


# ------------------------------------------------------------ leave_out


def _leave_out(root: etree._Element, refs: list) -> tuple[list[int], list[str]]:
    """Remove the named members from the copy and refit the groups that held
    them. Returns (source ids removed, warnings)."""
    if root.tag != qn("p:grpSp"):
        raise PptMcpError(
            f"leave_out takes members of a group, and {_what(root)} is not a "
            "group; leave it out, or copy the group that holds the shape"
        )
    everything = list(_sh._all_shapes(root))
    chosen: list = []
    for ref in refs:
        if isinstance(ref, bool) or not isinstance(ref, (int, str)):
            raise PptMcpError(f"leave_out takes member ids or names, got {ref!r}")
        hits = [] if isinstance(ref, int) else _sh._named(everything, ref)
        if not hits and isinstance(ref, str) and "/" in ref:
            hits = _sh._path_hits(everything, ref)
        if not hits and (isinstance(ref, int) or ref.strip().isdigit()):
            hits = [(e, c) for e, c in everything if _sh._shape_id(e) == int(ref)]
        if not hits:
            listing = ", ".join(
                f"{_sh._shape_id(e)}:{_sh._shape_name(e)!r}" for e, _c in everything)
            raise TargetNotFound(
                f"leave_out names {ref!r}, which is not a member of "
                f"{_what(root)}; its members (id:name) are {listing}"
            )
        if len(hits) > 1:
            ids = ", ".join(str(_sh._shape_id(e)) for e, _c in hits)
            raise AmbiguousTarget(
                f"leave_out names {ref!r}, which {len(hits)} members carry "
                f"(ids {ids}); send the id, or \"Group/Shape\""
            )
        chosen.append((hits[0][0], _sh._chain_of(hits[0][0])))
    removed: list[int] = []
    touched: list[etree._Element] = []
    for member, chain in chosen:
        parent = member.getparent()
        if parent is None:  # already removed with a group that held it
            continue
        removed.append(_sh._shape_id(member))
        parent.remove(member)
        touched.extend(chain)
    # A group left with no members goes too; the copy itself may not.
    for grp in reversed(list(root.iter(qn("p:grpSp")))):
        if any(child.tag in _sh._SHAPE_TAGS for child in grp):
            continue
        if grp is root:
            raise PptMcpError(
                "leave_out names every member of the group, so nothing would be "
                "left to copy"
            )
        grp.getparent().remove(grp)
    warnings: list[str] = []
    seen: set[int] = set()
    # Innermost first: a group's frame is refitted after the groups inside it.
    alive = set(root.iter(qn("p:grpSp")))
    for grp in reversed(list(root.iter(qn("p:grpSp")))):
        if grp not in alive or id(grp) in seen or not any(grp is t for t in touched):
            continue
        seen.add(id(grp))
        _refitted, refit_warnings = _sh._refit_groups([grp])
        warnings.extend(refit_warnings)
    return removed, warnings


# --------------------------------------------------------- relationships


def _plan_relationships(
    src: PptxPackage, src_part: str, root: etree._Element, same_file: bool
) -> dict[str, tuple[str, str, str, bool]]:
    """rId -> (kind, type, source target, external) for every relationship
    the copy uses, refusing what cannot travel. Nothing is changed here."""
    rels: dict[str, etree._Element] = {}
    if src.has_part(rels_name(src_part)):
        rels = {rel.get("Id"): rel for rel in src.root(rels_name(src_part))}
    plan: dict[str, tuple[str, str, str, bool]] = {}
    for node in root.iter():
        if not isinstance(node.tag, str):
            continue
        for attr, rid in node.attrib.items():
            if not attr.startswith(_R_NS) or not rid or rid in plan:
                continue
            rel = rels.get(rid)
            if rel is None:
                raise UnsupportedStructure(
                    f"{_what(root)} refers to relationship {rid}, which the "
                    "source slide does not define; the source deck needs repair "
                    "before this shape can be copied"
                )
            rel_type = rel.get("Type", "")
            raw = rel.get("Target", "")
            if rel.get("TargetMode") == "External" and rel_type in _EXTERNAL:
                plan[rid] = ("external", rel_type, raw, True)
                continue
            if rel.get("TargetMode") == "External":
                raise UnsupportedStructure(
                    f"{_what(root)} uses a relationship of type {rel_type}, which "
                    "this operation does not know how to copy; nothing was changed"
                )
            target = resolve_target(src_part, raw)
            if rel_type in _MEDIA or rel_type in _LEAVES:
                plan[rid] = ("leaf", rel_type, target, False)
            elif rel_type in _CHARTS:
                plan[rid] = ("chart", rel_type, target, False)
            elif rel_type == RT_SLIDE:
                if not same_file:
                    raise UnsupportedStructure(
                        f"{_what(root)} has a link that jumps to a slide of the "
                        "source deck, which is not in this deck; remove the link "
                        "(remove_hyperlink) or copy within the source deck"
                    )
                plan[rid] = ("slide", rel_type, target, False)
            elif rel_type in _REFUSED_TYPES:
                raise UnsupportedStructure(
                    f"{_what(root)} uses {_REFUSED_TYPES[rel_type]}, which this "
                    "operation does not copy; copy the slide with "
                    "copy_slide_between instead"
                )
            else:
                raise UnsupportedStructure(
                    f"{_what(root)} uses a relationship of type {rel_type}, which "
                    "this operation does not know how to copy; nothing was changed"
                )
    return plan


def _carry_relationships(
    src: PptxPackage, dst: PptxPackage, dst_part: str, root: etree._Element,
    plan: dict, ctx: dict,
) -> None:
    """Recreate each planned relationship on the receiving slide and point
    the copy's references at the new ids."""
    new_ids: dict[str, str] = {}
    for rid, (kind, rel_type, target, _external) in plan.items():
        if kind == "external":
            new_ids[rid] = dst.add_relationship(dst_part, rel_type, target, external=True)
            continue
        if kind == "chart":
            new_target = _import_chart(src, dst, target, ctx)
        elif kind == "leaf":
            new_target = _import_leaf(src, dst, target, ctx)
        else:  # a slide of the same deck keeps its slide
            new_target = target
        new_ids[rid] = dst.add_relationship(
            dst_part, rel_type, _rel_target(dst_part, new_target))
    for node in root.iter():
        if not isinstance(node.tag, str):
            continue
        for attr, rid in list(node.attrib.items()):
            if attr.startswith(_R_NS) and rid in new_ids:
                node.set(attr, new_ids[rid])


# -------------------------------------------------------------- the design


def _master_of(pkg: PptxPackage, slide_part: str) -> str | None:
    layout = _source_layout_of(pkg, slide_part)
    return _rel_target_of_type(pkg, layout, RT_SLIDE_MASTER) if layout else None


def _override(pkg: PptxPackage, slide_part: str) -> dict:
    el = pkg.root(slide_part).find(
        f"{qn('p:clrMapOvr')}/{qn('a:overrideClrMapping')}")
    return dict(el.attrib) if el is not None else {}


def _faces(ctx: dict | None) -> tuple[str | None, str | None]:
    if not ctx:
        return None, None
    return ctx["fonts"]["mj"].get("lt"), ctx["fonts"]["mn"].get("lt")


# ------------------------------------------------------------------- main


def copy_shape_between(
    pkg: PptxPackage,
    source_path: str | Path,
    source_slide,
    shape,
    slide,
    *,
    x: float | None = None,
    y: float | None = None,
    w: float | None = None,
    h: float | None = None,
    scale_text: bool = False,
    scale_lines: bool = False,
    no_wrap: bool = False,
    leave_out: list | None = None,
    name: str | None = None,
) -> dict:
    """Copy one shape or group from slide `source_slide` of the deck at
    `source_path` onto slide `slide` of `pkg`, at the top of the slide.
    Positions and sizes are inches on the receiving slide; see the module
    docstring for what travels and what is refused."""
    for value, label in ((w, "w"), (h, "h")):
        if value is not None and float(value) <= 0:
            raise PptMcpError(f"{label} must be positive inches, got {value}")
    if (scale_text or scale_lines) and w is None and h is None:
        raise PptMcpError(
            "scale_text and scale_lines follow a change of size; send w or h with them"
        )
    # The receiving slide first: the answer leads with it, not the source's.
    rec = resolve_slide(pkg, slide)
    dst_part = rec["part"]

    src_file = Path(check_path(source_path, "read source presentation"))
    require_deck(src_file, f"no source presentation at {src_file}")
    same_file = safesave.canonical_key(src_file) == safesave.canonical_key(pkg.path)
    src = PptxPackage(src_file)  # read from disk; nothing below changes it
    src_part_count = len(src.part_names())
    try:
        src_index, src_part, _entry, src_slide_id, _rid = _resolve_slide(src, source_slide)
    except TargetNotFound as exc:
        count = len(src.slide_parts())
        if isinstance(source_slide, int) and not isinstance(source_slide, bool):
            raise TargetNotFound(
                f"source_slide {source_slide} is out of range: the source deck has "
                f"{count} slide{'' if count == 1 else 's'}, so source_slide counts "
                f"from 0 and runs 0 to {max(count - 1, 0)} (source_slide_number 1 to "
                f"{count})"
            ) from None
        raise TargetNotFound(f"in the source deck, {exc}") from None
    try:
        elem, chain = _sh._find_shape(src, src_part, shape)
    except (TargetNotFound, AmbiguousTarget) as exc:
        raise type(exc)(f"in the source deck, {exc}") from None

    turned = [grp for grp in chain if _sh._turned(grp)]
    if turned:
        raise UnsupportedStructure(
            f"{_what(elem)} sits in group {_sh._shape_id(turned[0])}, which is rotated "
            "or flipped, so where it is drawn cannot be worked out here; copy the "
            "whole group, or straighten it first"
        )
    if _is_table(elem) and (w is not None or h is not None):
        raise PptMcpError(
            f"{_what(elem)} is a table, whose size comes from its column widths and "
            "row heights; copy it without w and h, then use set_column_widths and "
            "set_row_heights"
        )
    warnings: list[str] = []
    copy_el = _copy.deepcopy(elem)
    old_root_id = _sh._shape_id(elem)
    left_out: list[int] = []
    if leave_out:
        left_out, leave_warnings = _leave_out(copy_el, list(leave_out))
        warnings.extend(leave_warnings)
    # Refused on what is left, so leaving out a member that cannot travel works.
    _refuse_structure(copy_el, cross_deck=not same_file)
    if _sh._xfrm_of(copy_el) is None:
        raise UnsupportedStructure(
            f"{_what(elem)} has no position of its own; nothing to place")
    plan = _plan_relationships(src, src_part, copy_el, same_file)

    # Where the copy is drawn on the source slide, through its groups.
    ax, bx, ay, by, _rot = _sh._chain_transform(chain)
    cx0, cy0, cw0, ch0 = _sh._xfrm_box(_sh._require_xfrm(copy_el))
    sx, sy, sw, sh = ax * cx0 + bx, ay * cy0 + by, ax * cw0, ay * ch0
    nx = g.in_to_emu(x) if x is not None else round(sx)
    ny = g.in_to_emu(y) if y is not None else round(sy)
    if w is None and h is None:
        nw, nh = round(sw), round(sh)
    elif h is None:
        nw = g.in_to_emu(w)
        nh = round(sh * nw / sw) if sw else round(sh)
    elif w is None:
        nh = g.in_to_emu(h)
        nw = round(sw * nh / sh) if sh else round(sw)
    else:
        nw, nh = g.in_to_emu(w), g.in_to_emu(h)
    ratios = [n / s for n, s in ((nw, sw), (nh, sh)) if s]
    factor = min(ratios) if ratios else 1.0
    g.check_emu_box(nx, ny, nw, nh, what=_what(elem))

    # Written against the SOURCE deck before any id or reference changes.
    same_design = same_file and (
        _master_of(src, src_part) == _master_of(pkg, dst_part)
        and _override(src, src_part) == _override(pkg, dst_part))
    result: dict = {}
    if scale_text or not same_design:
        sizes = _scaling.scale_text(src, src_part, copy_el, factor if scale_text else 1.0)
        if scale_text:
            result["text_scaled"] = sizes
        elif sizes.get("sizes"):
            result["text_sizes_written"] = sizes["sizes"]
    if scale_lines or not same_design:
        lines = _scaling.scale_lines(src, src_part, copy_el, factor if scale_lines else 1.0)
        if scale_lines:
            result["lines_scaled"] = lines
    if no_wrap:
        result["unwrapped"] = _scaling.unwrap_text_boxes(copy_el)
    baked: dict = {}
    if not same_design:
        bake_ctx = _theme_bake_context(src, _source_layout_of(src, src_part))
        if bake_ctx is None:
            warnings.append(
                "the source slide's design (layout, master and theme) could not be "
                "read, so theme colours and fonts in the copy follow this deck's theme"
            )
        else:
            bake_ctx = {**bake_ctx, "clr_map": {**bake_ctx["clr_map"],
                                                **_override(src, src_part)}}
            _bake_theme_refs(copy_el, bake_ctx, baked)
            dst_ctx = _theme_bake_context(pkg, _source_layout_of(pkg, dst_part))
            src_faces, dst_faces = _faces(bake_ctx), _faces(dst_ctx)
            if dst_ctx is not None and src_faces != dst_faces:
                warnings.append(
                    "text that takes its typeface from the theme rather than naming "
                    f"one is set in this deck's theme fonts ({dst_faces[1]} for body "
                    f"text) instead of the source's ({src_faces[1]}); set the face "
                    "with format_text where it must match"
                )

    # Place it: slide space now, the copy's own child space unchanged.
    if copy_el.tag == qn("p:grpSp"):
        _xfrm, off, ext, _choff, _chext = _sh._group_parts(copy_el)
    else:
        xfrm = _sh._require_xfrm(copy_el)
        off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
    off.set("x", str(nx))
    off.set("y", str(ny))
    ext.set("cx", str(nw))
    ext.set("cy", str(nh))
    if name is not None:
        _sh._cnvpr(copy_el).set("name", name)

    # Fresh ids on the receiving slide; glue follows the copied shapes.
    next_id = pkg.next_shape_id(dst_part)
    ids: dict[str, int] = {}
    for cnvpr in copy_el.iter(qn("p:cNvPr")):
        ids[cnvpr.get("id", "")] = next_id
        cnvpr.set("id", str(next_id))
        next_id += 1
    for cxn in copy_el.iter(qn("p:cxnSp")):
        holder = cxn.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}")
        if holder is None:
            continue
        for tag in ("a:stCxn", "a:endCxn"):
            end = holder.find(qn(tag))
            if end is None:
                continue
            if end.get("id") in ids:
                end.set("id", str(ids[end.get("id")]))
                continue
            holder.remove(end)
            warnings.append(
                f"connector {_sh._shape_id(cxn)} was glued to shape {end.get('id')} on "
                "the source slide, which was not copied; that end keeps its place, "
                "unglued"
            )
    _regenerate_creation_ids(copy_el)

    ctx = {"memo": {}, "copied_parts": [], "media_added": [], "media_reused": [],
           "imported_design_parts": [], "warnings": warnings, "baked": {}}
    _carry_relationships(src, pkg, dst_part, copy_el, plan, ctx)

    tree = _sh._sp_tree(pkg, dst_part)
    last = tree[-1] if len(tree) else None
    if last is not None and last.tag == qn("p:extLst"):
        last.addprevious(copy_el)
    else:
        tree.append(copy_el)
    pkg.mark_dirty(dst_part)

    timing = src.root(src_part).find(qn("p:timing"))
    if timing is not None:
        named = sorted({t.get("spid") for t in timing.iter(qn("p:spTgt"))} & set(ids))
        if named:
            warnings.append(
                "the source slide's animation or media timing names shape(s) "
                f"{', '.join(named)}; timing stays with the source slide, so the copy "
                "has none (set_media_playback sets a video's or sound's playback)"
            )

    assert not src._dirty and len(src.part_names()) == src_part_count, (
        "internal error: the SOURCE package was changed during a shape copy"
    )
    new_id = _sh._shape_id(copy_el)
    result = {
        "shape_id": new_id,
        "created": [new_id],
        "name": _sh._shape_name(copy_el),
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "box": {k: round(v / 914400, 4) for k, v in
                (("x", nx), ("y", ny), ("w", nw), ("h", nh))},
        "ids": ids,
        "source": {
            "path": str(src_file),
            "slide_index": src_index,
            "slide_number": src_index + 1,
            "slide_id": src_slide_id,
            "shape_id": old_root_id,
            "name": _sh._shape_name(elem),
        },
        "theme": "kept" if same_design else "written",
        **result,
    }
    if baked:
        result["theme_values_written"] = baked
    if left_out:
        result["left_out"] = left_out
    for key in ("copied_parts", "media_added"):
        if ctx[key]:
            result[key] = ctx[key]
    if ctx["media_reused"]:
        result["media_reused"] = sorted(set(ctx["media_reused"]))
    if warnings:
        result["warnings"] = warnings
    return result


__all__ = ["copy_shape_between"]
