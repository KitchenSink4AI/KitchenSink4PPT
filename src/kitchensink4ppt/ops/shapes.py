"""Shape tool layer: insert, connect, edit-in-place, group, align,
distribute, and z-order native shapes.

Contract (all ops modules): every function takes the open PptxPackage first,
mutates only the in-memory package, calls pkg.mark_dirty() on every part it
touches, and returns a summary dict reporting every created or changed shape
id. Nothing here writes to disk; the caller decides when to save.

Coordinates: public parameters are INCHES (floats); everything is stored as
EMU ints (1 in = 914400 EMU, 1 pt = 12700 EMU). Slide addressing follows
ops/read.py (0-based index or {"slide_id": N}); shapes are addressed by
their per-slide p:cNvPr id.

Connectors: glued by default when shape ids are given. Glue is stCxn/endCxn
(shape id + connection-site index into the target's cxnLst). Whenever a
glued shape is moved or resized by this module, every connector glued to it
has its xfrm box RE-DERIVED from the new site points, because PowerPoint
renders the stored box until the user nudges something; stale boxes are the
detached-connector bug. Freeform (custGeom) shapes emitted by this engine
carry no connection sites, so gluing to them is refused rather than guessed.

Group math: a grpSp maps child space to its frame by
slide = off + (child - chOff) * ext / chExt (per axis). Groups are created
with identity mapping (chOff/chExt = off/ext), and all edits on shapes
inside groups resolve through the full ancestor chain, so absolute inch
positions always mean SLIDE space. Ancestor group rotation is not folded
into this math; edits under a rotated group carry a warning.
"""

from __future__ import annotations

import copy as _copy

from lxml import etree

from ..core.errors import PptMcpError, TargetNotFound, UnsupportedStructure
from ..core.package import PptxPackage, qn
from . import geometry as g
from .read import iter_shapes, resolve_slide

# --------------------------------------------------------------- text_style

#: Keys geometry.txbody understands for the single-style text replace.
_TXBODY_STYLE_KEYS = {
    "size", "color", "bold", "italic", "align", "anchor", "font", "wrap",
}
#: Extra run-level keys wired through ops/text.py's rPr writer after the
#: body is built (geometry.txbody stays minimal on purpose).
_TXBODY_EXTRA_KEYS = {"underline"}
#: Aliases agents actually type, mapped to the canonical key.
_TXBODY_STYLE_ALIASES = {
    "font_size": "size",
    "size_pt": "size",
    "font_name": "font",
    "typeface": "font",
}


def _split_text_style(style: dict | None) -> tuple[dict | None, dict]:
    """Validate a text_style dict for the single-style replace: canonical
    aliases are folded in, unknown keys REFUSE loudly (they used to be
    silently dropped, so styled text landed unstyled with no warning).
    Returns (base style for geometry.txbody, extra run props for
    text._apply_run_props)."""
    if style is None:
        return None, {}
    if not isinstance(style, dict):
        raise PptMcpError(
            f"text_style must be a dict, got {type(style).__name__}"
        )
    folded: dict = {}
    for key, value in style.items():
        canon = _TXBODY_STYLE_ALIASES.get(key, key)
        if canon in folded and canon != key:
            raise PptMcpError(
                f"text_style gives {key!r} and its alias {canon!r}; "
                "pass one of them"
            )
        folded[canon] = value
    allowed = _TXBODY_STYLE_KEYS | _TXBODY_EXTRA_KEYS
    unknown = sorted(set(folded) - allowed)
    if unknown:
        raise PptMcpError(
            f"unknown text_style key(s): {', '.join(unknown)}; allowed: "
            f"{', '.join(sorted(allowed))} (rich per-run restyling is "
            "format_text's job)"
        )
    base = {k: v for k, v in folded.items() if k in _TXBODY_STYLE_KEYS}
    extra = {k: v for k, v in folded.items() if k in _TXBODY_EXTRA_KEYS}
    return base, extra


#: Bullet children of a:pPr, plus the marker that switches bullets OFF.
#: Carrying these is what keeps set_bullets(style="none") from being undone
#: by a later retext.
_BULLET_TAGS = ("a:buNone", "a:buAutoNum", "a:buChar", "a:buBlip", "a:buFont")


def _merge_rpr(old_rpr: etree._Element, new_rpr: etree._Element):
    """Old run properties, with everything the caller just named laid on
    top. Returns (merged, carried) where carried names the categories that
    came from the old run.

    The freshly built a:rPr states exactly what text_style asked for, plus
    lang and dirty, so overlaying it onto a copy of the old one preserves
    colour, typeface, size, bold and italic that nobody asked to change.
    """
    from ._runmap import FILL_CHOICE_TAGS, RPR_ORDER, rank_insert, remove_children

    merged = _copy.deepcopy(old_rpr)
    merged.tag = new_rpr.tag
    carried: set[str] = set()
    if merged.get("sz"):
        carried.add("size")
    for attr, value in new_rpr.attrib.items():
        if attr == "lang" and merged.get("lang"):
            continue
        if attr in ("sz", "b", "i", "u"):
            carried.discard({"sz": "size"}.get(attr, attr))
        merged.set(attr, value)
    if merged.find(qn("a:solidFill")) is not None or merged.find(
        qn("a:gradFill")
    ) is not None:
        carried.add("run_color")
    if merged.find(qn("a:latin")) is not None:
        carried.add("font")
    for child in new_rpr:
        tag = etree.QName(child).localname
        if f"a:{tag}" in FILL_CHOICE_TAGS:
            remove_children(merged, FILL_CHOICE_TAGS)
            carried.discard("run_color")
        elif tag in ("latin", "ea", "cs"):
            remove_children(merged, ("a:latin", "a:ea", "a:cs"))
            carried.discard("font")
        else:
            remove_children(merged, (f"a:{tag}",))
        rank_insert(merged, _copy.deepcopy(child), RPR_ORDER)
    return merged, carried


#: a:bodyPr autofit children. The ELEMENT is the mode and gets carried;
#: the attributes on it are PowerPoint's cache of how far it had to shrink
#: the OLD text, and carrying those onto new text is a rendering defect.
_AUTOFIT_TAGS = ("a:normAutofit", "a:spAutoFit", "a:noAutofit")
_AUTOFIT_CACHE_ATTRS = ("fontScale", "lnSpcReduction")


#: What clearing the cache actually costs the caller, stated in the
#: result. PowerPoint does NOT recompute on open (measured: a cleared
#: normAutofit stays cleared through an open-and-save round trip), which
#: is what ops/text.py's _AUTOFIT_CAVEAT has said all along. So the frame
#: renders at 100% and its fit is genuinely unknown until something
#: re-fits it.
_AUTOFIT_RESET_NOTE = (
    "the cached shrink belonged to the old text and was cleared; "
    "PowerPoint does not recompute it on open, so this frame now renders "
    "at full size and its fit is unknown. Re-fit with fit_text, or check "
    'with check_layout(checks=["overflow"]).'
)


def _reset_autofit_cache(bodypr: etree._Element) -> bool:
    """Drop the cached autofit numbers, keep the mode. Returns whether
    anything was actually dropped.

    PowerPoint renders from `normAutofit@fontScale` until the frame is
    next edited in the app, so a box that once overflowed kept shrinking
    text that now fits: a placeholder retexted from a paragraph to three
    words still rendered at 40%. Clearing the element instead of the
    numbers would change the shape's autofit BEHAVIOUR, which nobody
    asked for.

    Clearing does NOT re-fit anything. PowerPoint recomputes the cache
    only when the frame is edited in the application, so a replacement
    LONGER than the original will now overflow visibly rather than shrink
    to a scale computed for text that is no longer there. Both are wrong
    renderings of the new text; an overflow is at least the honest one,
    and check_layout's overflow check sees it. The caller is told, and
    fit_text is the remedy: doing it here would silently rewrite every
    run's size on an edit that only asked to change the words.
    """
    dropped = False
    for tag in _AUTOFIT_TAGS:
        el = bodypr.find(qn(tag))
        if el is None:
            continue
        for attr in _AUTOFIT_CACHE_ATTRS:
            if el.get(attr) is not None:
                el.attrib.pop(attr)
                dropped = True
    return dropped


def _carry_text_properties(
    old: etree._Element, new: etree._Element, named: set
) -> tuple[list[str], dict]:
    """Carry the previous text body's properties onto its replacement, for
    everything the caller did not name.

    A single-style replace used to build a brand-new body from defaults, so
    replacing a body's text re-centred paragraphs that were left-aligned,
    put bullet glyphs back on paragraphs where set_bullets(style="none")
    had removed them, reset the vertical anchor, and dropped explicit run
    colour. None of that was requested and none of it was reported;
    `changed: ["text"]` was all the caller got back. Text replacement now
    touches text, and touches formatting only where the caller named it.

    Returns (carried categories, facts). `facts` reports what the carry
    could NOT preserve, because a `preserved` list that only ever says
    what survived is a half-truth on a body whose runs disagreed.
    """
    from ._runmap import remove_children

    carried: set[str] = set()
    facts: dict = {}

    old_bodypr = old.find(qn("a:bodyPr"))
    new_bodypr = new.find(qn("a:bodyPr"))
    if old_bodypr is not None and new_bodypr is not None:
        kept_anchor = new_bodypr.get("anchor")
        kept_wrap = new_bodypr.get("wrap")
        replacement = _copy.deepcopy(old_bodypr)
        if _reset_autofit_cache(replacement):
            facts["autofit_scale_reset"] = _AUTOFIT_RESET_NOTE
        if replacement.find(qn("a:normAutofit")) is not None or (
            replacement.find(qn("a:spAutoFit")) is not None
        ):
            carried.add("autofit_mode")
        if "anchor" in named:
            if kept_anchor is None:
                replacement.attrib.pop("anchor", None)
            else:
                replacement.set("anchor", kept_anchor)
        elif replacement.get("anchor") != kept_anchor:
            carried.add("anchor")
        if "wrap" in named:
            if kept_wrap is None:
                replacement.attrib.pop("wrap", None)
            else:
                replacement.set("wrap", kept_wrap)
        new.replace(new_bodypr, replacement)

    old_lst = old.find(qn("a:lstStyle"))
    new_lst = new.find(qn("a:lstStyle"))
    if old_lst is not None and len(old_lst) and new_lst is not None:
        new.replace(new_lst, _copy.deepcopy(old_lst))
        carried.add("list_style")

    old_paras = old.findall(qn("a:p"))
    if not old_paras:
        return sorted(carried), facts

    # A paragraph whose runs disagreed about their formatting cannot be
    # carried: the replacement is one run, so it can only take ONE run's
    # properties, and it takes the first. That loss predates this carry
    # (a full rebuild lost it too), but reporting `preserved` over it
    # without saying so would be the carry claiming credit for a wreck.
    if any(_runs_disagree(p) for p in old_paras):
        facts["runs_collapsed_to_first"] = True
    # Spreading the FIRST run's hyperlink across the whole replacement
    # would silently relink text that pointed somewhere else, or nowhere.
    # When the source runs did not agree on their link, no link is carried
    # rather than the wrong one, and the caller is told.
    #
    # The decision is PER PARAGRAPH. It used to be one flag for the whole
    # body, so a paragraph mixing linked and unlinked runs stripped the
    # link off every OTHER paragraph, including ones where every run
    # pointed at the same place (second review, G3, 2026-09-22).
    disagrees = {id(p): _runs_disagree_on_links(p) for p in old_paras}
    dropped_at: list[int] = []

    # `carried` is body-global, so a paragraph whose level the caller
    # overrode must not erase the accurate report from a paragraph that
    # DID carry one (review finding N2). Counted, then folded in at the end.
    levels_carried = 0

    for i, para in enumerate(new.findall(qn("a:p"))):
        # A replacement with MORE paragraphs than the original takes the
        # last one's shape, which is what pressing Enter in PowerPoint does.
        source = old_paras[min(i, len(old_paras) - 1)]
        drop_links = disagrees[id(source)]
        if drop_links:
            dropped_at.append(i)

        old_ppr = source.find(qn("a:pPr"))
        new_ppr = para.find(qn("a:pPr"))
        if new_ppr is not None:
            kept_align = new_ppr.get("algn")
            # An outline level stated by the caller outranks the one on the
            # paragraph being replaced, the same way a named alignment does.
            # Nothing in the single-style replace path sets lvl, so this is
            # inert there; the placeholder path uses it.
            kept_lvl = new_ppr.get("lvl")
            if old_ppr is None:
                # The original inherited everything. Preserve that rather
                # than pinning the default alignment onto it.
                if "align" not in named and kept_lvl is None:
                    para.remove(new_ppr)
            else:
                replacement = _copy.deepcopy(old_ppr)
                if "align" in named:
                    if kept_align is None:
                        replacement.attrib.pop("algn", None)
                    else:
                        replacement.set("algn", kept_align)
                elif replacement.get("algn") != kept_align:
                    carried.add("alignment")
                if any(
                    replacement.find(qn(t)) is not None for t in _BULLET_TAGS
                ):
                    carried.add("bullets")
                old_lvl = old_ppr.get("lvl")
                if kept_lvl is not None:
                    replacement.set("lvl", kept_lvl)
                if old_lvl and (kept_lvl is None or kept_lvl == old_lvl):
                    levels_carried += 1
                para.replace(new_ppr, replacement)

        src_run_rpr = _first_run_rpr(source)
        src_end_rpr = source.find(qn("a:endParaRPr"))
        # A destination endParaRPr is the paragraph's INSERTION style, so
        # its source is the old paragraph's endParaRPr when there is one:
        # taking the first visible run's style instead loses a 24 pt
        # insertion point behind 10 pt body text (review finding M2).
        sources = {
            "a:rPr": src_run_rpr if src_run_rpr is not None else src_end_rpr,
            "a:endParaRPr": (
                src_end_rpr if src_end_rpr is not None else src_run_rpr
            ),
        }
        for tag, src_rpr in sources.items():
            if src_rpr is None:
                continue
            if tag == "a:rPr":
                targets = [
                    rpr
                    for holder in _RUN_FAMILY
                    for rpr in para.findall(f"{qn(holder)}/{qn('a:rPr')}")
                ]
            else:
                targets = para.findall(qn("a:endParaRPr"))
            for rpr in targets:
                merged, got = _merge_rpr(src_rpr, rpr)
                if drop_links:
                    remove_children(merged, _HLINK_TAGS)
                rpr.getparent().replace(rpr, merged)
                carried |= got
    if dropped_at:
        facts["hyperlinks_dropped"] = (
            "paragraph index "
            + ", ".join(str(n) for n in dropped_at)
            + ": the replaced runs did not share one hyperlink, so none was "
            "carried onto that paragraph's new text; re-apply it with "
            "set_hyperlink"
        )
    if levels_carried > 0:
        carried.add("level")
    return sorted(carried), facts


#: The a:p children that can carry character properties. a:fld (a slide
#: number, a date) and a:br (a line break) hold a real a:rPr just as a:r
#: does, and a paragraph can consist of nothing else (review finding B1).
_RUN_FAMILY = ("a:r", "a:fld", "a:br")

#: a:r and a:fld carry text; a:br does not.
_TEXT_BEARING = ("a:r", "a:fld")

#: The hyperlink elements a RUN carries. The hover one is a:hlinkMouseOver
#: (a:hlinkHover is the SHAPE-level spelling, inside p:cNvPr, and is not a
#: legal child of a:rPr). Comparing the wrong name made every hover link
#: compare equal to every other, so a disagreeing one was spread over the
#: replacement in silence (second review, G3, 2026-09-22). The names and
#: their order match _runmap.RPR_ORDER.
_HLINK_TAGS = ("a:hlinkClick", "a:hlinkMouseOver")

#: rPr attributes that say nothing about how a run LOOKS. Office writes
#: and rewrites these constantly, so two runs that differ only here are
#: the same run as far as a reader is concerned.
_RPR_NOISE_ATTRS = frozenset({"dirty", "smtClean", "smtId", "err"})


def _first_run_rpr(para: etree._Element) -> etree._Element | None:
    """The first character-property element in the paragraph, in DOCUMENT
    order across a:r, a:fld and a:br.

    Looking only at a:r meant a placeholder holding nothing but a slide
    number field lost its size, colour and East Asian font the moment its
    text was replaced, and reported nothing lost (review finding B1).
    """
    for child in para:
        if f"a:{etree.QName(child).localname}" in _RUN_FAMILY:
            rpr = child.find(qn("a:rPr"))
            if rpr is not None:
                return rpr
    return None


def _element_signature(el: etree._Element) -> tuple:
    """A whole element subtree reduced to something comparable, with
    attribute order normalized and Office's bookkeeping attributes
    dropped. Recursive, so nothing inside a:rPr is invisible to it."""
    attrs = tuple(
        sorted(
            (k, v)
            for k, v in el.attrib.items()
            if etree.QName(k).localname not in _RPR_NOISE_ATTRS
        )
    )
    return (el.tag, attrs, tuple(_element_signature(c) for c in el))


def _run_signature(run: etree._Element) -> tuple:
    """What a run looks like, reduced to something comparable.

    The old version compared a hand-picked handful of attributes plus one
    fill and a:latin, so runs differing in East Asian or complex-script
    typeface, rtl, language, effects, highlight or hyperlink compared
    EQUAL and were collapsed without a word (review finding M1). The whole
    a:rPr subtree is compared now; over-reporting a collapse is a report,
    under-reporting one is a silent loss.
    """
    rpr = run.find(qn("a:rPr"))
    if rpr is None:
        return ()
    return _element_signature(rpr)


def _runs_disagree(para: etree._Element) -> bool:
    """True when a paragraph's runs are not all formatted alike.

    a:r and a:fld count with or without properties: a text-bearing run
    with no a:rPr renders from inherited properties, which is genuinely
    different from an explicit one. A bare a:br does not, because it shows
    no glyphs, so counting it would report a collapse on every ordinary
    two-line paragraph.
    """
    sigs = {
        _run_signature(r)
        for holder in _TEXT_BEARING
        for r in para.findall(qn(holder))
    }
    sigs |= {
        _run_signature(b)
        for b in para.findall(qn("a:br"))
        if b.find(qn("a:rPr")) is not None
    }
    return len(sigs) > 1


def _runs_disagree_on_links(para: etree._Element) -> bool:
    """True when the paragraph's text-bearing runs do not all point at the
    same place (counting 'nowhere' as a place)."""
    links = set()
    for holder in _TEXT_BEARING:
        for run in para.findall(qn(holder)):
            rpr = run.find(qn("a:rPr"))
            if rpr is None:
                links.add(())
                continue
            links.add(
                tuple(
                    _element_signature(el)
                    for tag in _HLINK_TAGS
                    for el in rpr.findall(qn(tag))
                )
            )
    return len(links) > 1


def _apply_extra_run_props(body: etree._Element, extra: dict) -> None:
    """Apply the run-level keys geometry.txbody does not know (currently
    underline) onto every a:rPr / a:endParaRPr of a freshly built body,
    via ops/text.py's schema-order-aware writer."""
    if not extra:
        return
    from .text import _apply_run_props

    for tag in ("a:rPr", "a:endParaRPr"):
        for rpr in body.iter(qn(tag)):
            _apply_run_props(rpr, underline=extra.get("underline"))


# ------------------------------------------------------------------ presets

#: Friendly name -> DrawingML preset. Any other name can be passed through
#: verbatim with the "prst:" prefix (e.g. "prst=heptagon" is "prst:heptagon").
PRESETS = {
    "rect": "rect",
    "rectangle": "rect",
    "rounded_rect": "roundRect",
    "ellipse": "ellipse",
    "circle": "ellipse",
    "triangle": "triangle",
    "right_triangle": "rtTriangle",
    "diamond": "diamond",
    "chevron": "chevron",
    "pentagon_arrow": "homePlate",
    "hexagon": "hexagon",
    "pentagon": "pentagon",
    "parallelogram": "parallelogram",
    "trapezoid": "trapezoid",
    "octagon": "octagon",
    "star4": "star4",
    "star5": "star5",
    "arrow_right": "rightArrow",
    "arrow_left": "leftArrow",
    "arrow_up": "upArrow",
    "arrow_down": "downArrow",
    "arrow_left_right": "leftRightArrow",
    "arrow_up_down": "upDownArrow",
    "arrow_quad": "quadArrow",
    "arrow_bent": "bentArrow",
    "arrow_u_turn": "uturnArrow",
    "callout_rect": "wedgeRectCallout",
    "callout_rounded_rect": "wedgeRoundRectCallout",
    "callout_ellipse": "wedgeEllipseCallout",
    "frame": "frame",
    "plus": "mathPlus",
    "minus": "mathMinus",
    "pie": "pie",
    "arc": "arc",
    "donut": "donut",
    "cloud": "cloud",
    "heart": "heart",
    "lightning": "lightningBolt",
    "flowchart_process": "flowChartProcess",
    "flowchart_decision": "flowChartDecision",
    "flowchart_terminator": "flowChartTerminator",
    "flowchart_data": "flowChartInputOutput",
    "flowchart_document": "flowChartDocument",
}

_CONNECTOR_PRST = {
    "straight": "line",
    "elbow": "bentConnector3",
    "curved": "curvedConnector3",
}

#: Fractional (x, y) connection-site positions per preset, in cxnLst order.
#: rect family: 0=top, 1=left, 2=bottom, 3=right (verified ground truth).
_SITES_4 = ((0.5, 0.0), (0.0, 0.5), (0.5, 1.0), (1.0, 0.5))
#: ellipse: 8 sites, 0=top then counterclockwise (verified: idx 2 = left).
_IL = 0.14645
_IR = 1.0 - _IL
_SITES_ELLIPSE = (
    (0.5, 0.0), (_IL, _IL), (0.0, 0.5), (_IL, _IR),
    (0.5, 1.0), (_IR, _IR), (1.0, 0.5), (_IR, _IL),
)
_PRST_SITES = {
    "rect": _SITES_4,
    "roundRect": _SITES_4,
    "snip1Rect": _SITES_4,
    "snip2SameRect": _SITES_4,
    "diamond": _SITES_4,
    "flowChartProcess": _SITES_4,
    "flowChartDecision": _SITES_4,
    "flowChartTerminator": _SITES_4,
    "flowChartInputOutput": _SITES_4,
    "hexagon": _SITES_4,
    "ellipse": _SITES_ELLIPSE,
}

_SHAPE_TAGS = tuple(qn(t) for t in ("p:sp", "p:pic", "p:graphicFrame", "p:grpSp", "p:cxnSp"))


# --------------------------------------------------------------- resolution


def _sp_tree(pkg: PptxPackage, part: str) -> etree._Element:
    tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if tree is None:
        raise UnsupportedStructure(f"{part} has no p:spTree")
    return tree


def _cnvpr(elem: etree._Element) -> etree._Element | None:
    for child in elem:
        if etree.QName(child).localname.startswith("nv"):
            return child.find(qn("p:cNvPr"))
    return None


def _shape_id(elem: etree._Element) -> int | None:
    cnvpr = _cnvpr(elem)
    if cnvpr is None:
        return None
    try:
        return int(cnvpr.get("id", ""))
    except ValueError:
        return None


def _find_shape(
    pkg: PptxPackage, part: str, shape_id: int
) -> tuple[etree._Element, list[etree._Element]]:
    """(shape element, ancestor grpSp chain outermost-first) for one shape id.
    Searches the whole tree including group interiors."""
    if isinstance(shape_id, bool) or not isinstance(shape_id, int):
        raise PptMcpError(f"shape id must be an int, got {shape_id!r}")

    def _walk(container, chain):
        for child in container:
            if child.tag not in _SHAPE_TAGS:
                continue
            if _shape_id(child) == shape_id:
                return child, chain
            if child.tag == qn("p:grpSp"):
                found = _walk(child, chain + [child])
                if found is not None:
                    return found
        return None

    hit = _walk(_sp_tree(pkg, part), [])
    if hit is None:
        known = sorted(
            i for i in (
                _shape_id(e) for e, _k, _z, _p in iter_shapes(_sp_tree(pkg, part))
            ) if i is not None
        )
        raise TargetNotFound(
            f"no shape with id {shape_id} on that slide; ids present: {known}"
        )
    return hit


def _spPr_of(elem: etree._Element) -> etree._Element:
    tag = "p:grpSpPr" if elem.tag == qn("p:grpSp") else "p:spPr"
    pr = elem.find(qn(tag))
    if pr is None:
        raise UnsupportedStructure(
            f"shape {_shape_id(elem)} has no {tag}; refusing to guess"
        )
    return pr


def _xfrm_of(elem: etree._Element) -> etree._Element | None:
    if elem.tag == qn("p:graphicFrame"):
        return elem.find(qn("p:xfrm"))
    return _spPr_of(elem).find(qn("a:xfrm"))


def _xfrm_box(xfrm: etree._Element) -> tuple[int, int, int, int]:
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        raise UnsupportedStructure("a:xfrm without off/ext")
    return (
        int(off.get("x")), int(off.get("y")),
        int(ext.get("cx")), int(ext.get("cy")),
    )


def _require_xfrm(elem: etree._Element) -> etree._Element:
    xfrm = _xfrm_of(elem)
    if xfrm is None:
        raise UnsupportedStructure(
            f"shape {_shape_id(elem)} has no explicit geometry (a placeholder "
            "inheriting from its layout); set an absolute position and size "
            "first, or edit the placeholder through the text tools"
        )
    return xfrm


def _seed_xfrm(
    pkg: PptxPackage, part: str, elem: etree._Element
) -> etree._Element | None:
    """Give a layout-inheriting placeholder its FIRST explicit a:xfrm,
    copied from whatever it currently renders with.

    The old refusal was circular: it told the caller to set an absolute
    position and size first, and set_shape is the tool that does that, so
    a full x/y/w/h call could not get past it. The box a placeholder
    renders with is one or two hops up the chain (ops/inherit.py), so the
    honest move is to seed from there and let the caller's values land on
    top. A shape whose chain has no box either still refuses.
    """
    from . import inherit as _inh

    box = _inh.inherited_box(pkg, part, elem)
    if box is None:
        return None
    x, y, cx, cy = box
    sppr = _spPr_of(elem)
    g.insert_spPr_child(sppr, g.xfrm_element(x, y, cx, cy))
    return _xfrm_of(elem)


# ------------------------------------------------------- coordinate algebra


def _chain_transform(chain: list[etree._Element]) -> tuple[float, float, float, float, bool]:
    """(ax, bx, ay, by, rotated) mapping CHILD space to SLIDE space:
    slide_x = ax * child_x + bx (same for y). chain is outermost-first.
    rotated is True when any ancestor group carries a rotation (the affine
    math here ignores it; callers surface a warning)."""
    ax = ay = 1.0
    bx = by = 0.0
    rotated = False
    for grp in chain:
        xfrm = grp.find(f"{qn('p:grpSpPr')}/{qn('a:xfrm')}")
        if xfrm is None:
            continue
        if xfrm.get("rot") or xfrm.get("flipH") or xfrm.get("flipV"):
            rotated = True
        x, y, cx, cy = _xfrm_box(xfrm)
        choff = xfrm.find(qn("a:chOff"))
        chext = xfrm.find(qn("a:chExt"))
        chx = int(choff.get("x")) if choff is not None else x
        chy = int(choff.get("y")) if choff is not None else y
        chcx = int(chext.get("cx")) if chext is not None else cx
        chcy = int(chext.get("cy")) if chext is not None else cy
        if chcx == 0 or chcy == 0:
            raise UnsupportedStructure(
                "ancestor group has zero child extent; cannot resolve coordinates"
            )
        sx = cx / chcx
        sy = cy / chcy
        # compose: slide = prior(off + (u - chOff) * s)
        bx = bx + ax * (x - chx * sx)
        by = by + ay * (y - chy * sy)
        ax *= sx
        ay *= sy
    if ax == 0 or ay == 0:
        raise UnsupportedStructure(
            "ancestor group has zero extent; cannot resolve coordinates"
        )
    return ax, bx, ay, by, rotated


def _slide_box(elem: etree._Element, chain: list[etree._Element]) -> tuple[float, float, float, float]:
    """(x, y, cx, cy) of a shape in SLIDE EMU space, resolved through groups."""
    x, y, cx, cy = _xfrm_box(_require_xfrm(elem))
    ax, bx, ay, by, _rot = _chain_transform(chain)
    return ax * x + bx, ay * y + by, ax * cx, ay * cy


# ------------------------------------------------------------ shape builder


def _nv_pr(sp: etree._Element, tag: str, shape_id: int, name: str) -> etree._Element:
    nv = etree.SubElement(sp, qn(tag))
    cnvpr = etree.SubElement(nv, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", name)
    return nv


def _resolve_preset(shape_type: str) -> str:
    if shape_type.startswith("prst:"):
        raw = shape_type[len("prst:"):]
        if not raw or not raw[0].isalpha():
            raise PptMcpError(f"invalid raw preset name {raw!r}")
        return raw
    prst = PRESETS.get(shape_type)
    if prst is None:
        raise PptMcpError(
            f"unknown shape_type {shape_type!r}; one of: "
            f"{', '.join(sorted(PRESETS))}, freeform, or 'prst:<DrawingML "
            f"preset name>' verbatim"
        )
    return prst


def _freeform_geom(path, cx_emu: int, cy_emu: int) -> tuple[etree._Element, list[str]]:
    """custGeom from a freeform path spec in LOCAL INCHES (origin = shape
    top-left, extent = w x h). Returns (geom, warnings)."""
    if not path:
        raise PptMcpError(
            'shape_type="freeform" needs a path: a list of commands '
            '[["move", x, y], ["line", x, y], ["cubic", c1x, c1y, c2x, c2y, '
            'x, y], ["close"]] in local inches, or a list of such lists for '
            "multiple paths"
        )
    warnings: list[str] = []
    if path and isinstance(path[0], dict):
        specs = path
    elif path and isinstance(path[0], (list, tuple)) and path[0] and isinstance(
        path[0][0], (list, tuple)
    ):
        specs = [{"commands": p} for p in path]
    else:
        specs = [{"commands": path}]
    converted = []
    for spec in specs:
        cmds = []
        contours = 0
        for cmd in spec.get("commands", []):
            op = str(cmd[0])
            coords = [g.in_to_emu(v) for v in cmd[1:]]
            if op == "move":
                contours += 1
            cmds.append((op, *coords))
        if contours > 1:
            warnings.append(
                "multi-contour path: PowerPoint fills overlapping contours "
                "even-odd regardless of winding; same-winding overlaps will "
                "show holes (split into separate paths for solid fill)"
            )
        converted.append(
            {
                "commands": cmds,
                "fill": spec.get("fill", "norm"),
                "stroke": spec.get("stroke", True),
            }
        )
    return g.cust_geom(converted, cx_emu, cy_emu), warnings


def insert_shape(
    pkg: PptxPackage,
    slide,
    shape_type: str,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    adjustments=None,
    path=None,
    fill=None,
    line=None,
    effect=None,
    text: str | None = None,
    text_style: dict | None = None,
    name: str | None = None,
    rotation: float = 0.0,
    flip_h: bool = False,
    flip_v: bool = False,
) -> dict:
    """Insert one native shape at x, y with size w x h (inches, slide space).

    shape_type: a preset name from PRESETS, "prst:<name>" verbatim, or
    "freeform" with `path` (commands in local inches; see _freeform_geom).
    adjustments: preset adjust values (dict or list; float = fraction).
    fill / line / effect / text / text_style: see ops/geometry.py specs.
    Without an explicit fill the shape carries PowerPoint's default theme
    style block (accent1) so it renders theme-native.
    """
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    sp_tree = _sp_tree(pkg, part)
    for value, label in ((w, "w"), (h, "h")):
        if float(value) <= 0:
            raise PptMcpError(f"{label} must be positive inches, got {value}")
    shape_id = pkg.next_shape_id(part)
    warnings: list[str] = []

    x_emu, y_emu = g.in_to_emu(x), g.in_to_emu(y)
    cx_emu, cy_emu = g.in_to_emu(w), g.in_to_emu(h)

    sp = etree.SubElement(sp_tree, qn("p:sp"))
    display = name or f"{shape_type.replace('prst:', '')} {shape_id}"
    nv = _nv_pr(sp, "p:nvSpPr", shape_id, display)
    etree.SubElement(nv, qn("p:cNvSpPr"))
    etree.SubElement(nv, qn("p:nvPr"))
    sppr = etree.SubElement(sp, qn("p:spPr"))
    sppr.append(
        g.xfrm_element(
            x_emu, y_emu, cx_emu, cy_emu,
            rot=rotation, flip_h=flip_h, flip_v=flip_v,
        )
    )
    if shape_type == "freeform":
        if adjustments:
            raise PptMcpError("freeform shapes take a path, not adjustments")
        geom, warnings = _freeform_geom(path, cx_emu, cy_emu)
        sppr.append(geom)
    else:
        if path is not None:
            raise PptMcpError(
                'path is only valid with shape_type="freeform"'
            )
        sppr.append(g.prst_geom(_resolve_preset(shape_type), adjustments))
    fill_el = g.fill_element(fill)
    if fill_el is not None:
        sppr.append(fill_el)
    line_el = g.line_element(line)
    if line_el is not None:
        sppr.append(line_el)
    effect_el = g.effect_element(effect)
    if effect_el is not None:
        sppr.append(effect_el)
    # PowerPoint writes p:style on every inserted shape; explicit spPr
    # fills/lines override it, missing pieces render theme-native.
    sp.append(g.default_style())
    # p:sp requires a txBody (repair risk without one).
    base_style, extra_style = _split_text_style(text_style)
    body = g.txbody(text if text is not None else "", base_style)
    _apply_extra_run_props(body, extra_style)
    sp.append(body)
    pkg.mark_dirty(part)
    result = {
        "shape_id": shape_id,
        "created": [shape_id],
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "type": shape_type,
        "name": display,
    }
    if warnings:
        result["warnings"] = warnings
    return result


# --------------------------------------------------------------- connectors


def _prst_of(elem: etree._Element) -> str | None:
    geom = _spPr_of(elem).find(qn("a:prstGeom"))
    return geom.get("prst") if geom is not None else None


def _site_table(elem: etree._Element) -> tuple[tuple[tuple[float, float], ...], bool]:
    """(fractional site table, exact) for a shape. Falls back to the 4-site
    rect table (approximate box only; PowerPoint still glues by the real
    site). Refuses custGeom shapes: this engine emits them with an empty
    cxnLst, so there is nothing to glue to."""
    prst = _prst_of(elem)
    if prst is None:
        if _spPr_of(elem).find(qn("a:custGeom")) is not None:
            raise UnsupportedStructure(
                f"shape {_shape_id(elem)} is a freeform (custGeom) without "
                "connection sites; glue to a preset shape or use coordinate "
                "mode (start=/end=)"
            )
        raise UnsupportedStructure(
            f"shape {_shape_id(elem)} has no preset geometry to glue to"
        )
    table = _PRST_SITES.get(prst)
    if table is not None:
        return table, True
    return _SITES_4, False


def _site_point(
    pkg: PptxPackage, part: str, shape_id: int, idx: int | None, toward: tuple[float, float] | None
) -> tuple[float, float, int, bool]:
    """(slide x, slide y, resolved idx, exact) of one connection site.
    idx=None picks the site nearest `toward` (slide EMU point)."""
    elem, chain = _find_shape(pkg, part, shape_id)
    table, exact = _site_table(elem)
    x, y, cx, cy = _slide_box(elem, chain)
    points = [(x + fx * cx, y + fy * cy) for fx, fy in table]
    if idx is None:
        if toward is None:
            idx = 0
        else:
            idx = min(
                range(len(points)),
                key=lambda i: (points[i][0] - toward[0]) ** 2
                + (points[i][1] - toward[1]) ** 2,
            )
    if not 0 <= idx < len(points):
        raise PptMcpError(
            f"connection site {idx} out of range for shape {shape_id} "
            f"({len(points)} sites, 0..{len(points) - 1})"
        )
    return points[idx][0], points[idx][1], idx, exact


def _shape_center(pkg: PptxPackage, part: str, shape_id: int) -> tuple[float, float]:
    elem, chain = _find_shape(pkg, part, shape_id)
    x, y, cx, cy = _slide_box(elem, chain)
    return x + cx / 2, y + cy / 2


def _connector_style() -> etree._Element:
    style = etree.Element(qn("p:style"))
    lnref = etree.SubElement(style, qn("a:lnRef"))
    lnref.set("idx", "1")
    etree.SubElement(lnref, qn("a:schemeClr")).set("val", "accent1")
    fillref = etree.SubElement(style, qn("a:fillRef"))
    fillref.set("idx", "0")
    etree.SubElement(fillref, qn("a:schemeClr")).set("val", "accent1")
    effref = etree.SubElement(style, qn("a:effectRef"))
    effref.set("idx", "0")
    etree.SubElement(effref, qn("a:schemeClr")).set("val", "accent1")
    fontref = etree.SubElement(style, qn("a:fontRef"))
    fontref.set("idx", "minor")
    etree.SubElement(fontref, qn("a:schemeClr")).set("val", "tx1")
    return style


def _set_connector_box(
    cxnsp: etree._Element,
    p1: tuple[float, float],
    p2: tuple[float, float],
    chain: list[etree._Element],
) -> None:
    """Write the connector's xfrm from two SLIDE-space endpoints, converting
    into the connector's parent space and encoding direction as flips."""
    ax, bx, ay, by, _rot = _chain_transform(chain)
    q1 = ((p1[0] - bx) / ax, (p1[1] - by) / ay)
    q2 = ((p2[0] - bx) / ax, (p2[1] - by) / ay)
    x = min(q1[0], q2[0])
    y = min(q1[1], q2[1])
    cx = abs(q2[0] - q1[0])
    cy = abs(q2[1] - q1[1])
    sppr = _spPr_of(cxnsp)
    g.insert_spPr_child(
        sppr,
        g.xfrm_element(
            round(x), round(y), round(cx), round(cy),
            flip_h=q2[0] < q1[0], flip_v=q2[1] < q1[1],
        ),
    )


def insert_connector(
    pkg: PptxPackage,
    slide,
    kind: str = "straight",
    *,
    start: tuple[float, float] | list | None = None,
    end: tuple[float, float] | list | None = None,
    start_shape: int | None = None,
    end_shape: int | None = None,
    start_site: int | None = None,
    end_site: int | None = None,
    line=None,
    name: str | None = None,
) -> dict:
    """Insert a connector. kind: straight | elbow | curved.

    Glued mode (the default whenever shape ids are given): start_shape /
    end_shape are shape ids; start_site / end_site are connection-site
    indexes (rect family: 0=top 1=left 2=bottom 3=right; ellipse: 8 sites
    0=top counterclockwise). Omitted sites auto-pick the nearest facing
    pair. Coordinate mode: start / end as (x, y) inches. The two modes mix
    per endpoint. line: see ops/geometry.py (width, color, dash, head/tail
    arrowheads).
    """
    prst = _CONNECTOR_PRST.get(kind)
    if prst is None:
        raise PptMcpError(
            f"unknown connector kind {kind!r}; one of: "
            f"{', '.join(sorted(_CONNECTOR_PRST))}"
        )
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    if start_shape is None and start is None:
        raise PptMcpError("connector needs start_shape (glued) or start=(x, y)")
    if end_shape is None and end is None:
        raise PptMcpError("connector needs end_shape (glued) or end=(x, y)")

    # Aim points for auto site selection.
    if start_shape is not None and end_shape is not None:
        start_aim = _shape_center(pkg, part, end_shape)
        end_aim = _shape_center(pkg, part, start_shape)
    else:
        start_aim = (
            (g.in_to_emu(end[0]), g.in_to_emu(end[1])) if end is not None else None
        )
        end_aim = (
            (g.in_to_emu(start[0]), g.in_to_emu(start[1])) if start is not None else None
        )

    if start_shape is not None:
        p1x, p1y, start_site, _ = _site_point(pkg, part, start_shape, start_site, start_aim)
        p1 = (p1x, p1y)
    else:
        p1 = (g.in_to_emu(start[0]), g.in_to_emu(start[1]))
    if end_shape is not None:
        p2x, p2y, end_site, _ = _site_point(pkg, part, end_shape, end_site, end_aim)
        p2 = (p2x, p2y)
    else:
        p2 = (g.in_to_emu(end[0]), g.in_to_emu(end[1]))

    sp_tree = _sp_tree(pkg, part)
    shape_id = pkg.next_shape_id(part)
    display = name or f"{kind} connector {shape_id}"
    cxnsp = etree.SubElement(sp_tree, qn("p:cxnSp"))
    nv = etree.SubElement(cxnsp, qn("p:nvCxnSpPr"))
    cnvpr = etree.SubElement(nv, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", display)
    cnvcxn = etree.SubElement(nv, qn("p:cNvCxnSpPr"))
    etree.SubElement(cnvcxn, qn("a:cxnSpLocks"))
    if start_shape is not None:
        st = etree.SubElement(cnvcxn, qn("a:stCxn"))
        st.set("id", str(start_shape))
        st.set("idx", str(start_site))
    if end_shape is not None:
        en = etree.SubElement(cnvcxn, qn("a:endCxn"))
        en.set("id", str(end_shape))
        en.set("idx", str(end_site))
    etree.SubElement(nv, qn("p:nvPr"))
    sppr = etree.SubElement(cxnsp, qn("p:spPr"))
    _set_connector_box(cxnsp, p1, p2, [])
    if prst == "line":
        sppr.append(g.prst_geom(prst))
    else:
        sppr.append(g.prst_geom(prst, {"adj1": 50000}))
    line_el = g.line_element(line)
    if line_el is not None:
        sppr.append(line_el)
    cxnsp.append(_connector_style())
    pkg.mark_dirty(part)
    return {
        "shape_id": shape_id,
        "created": [shape_id],
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "kind": kind,
        "glued": {
            "start": {"shape_id": start_shape, "site": start_site}
            if start_shape is not None
            else None,
            "end": {"shape_id": end_shape, "site": end_site}
            if end_shape is not None
            else None,
        },
        "name": display,
    }


# ------------------------------------------------------------- reroute glue


def _connector_endpoints_slide(
    cxnsp: etree._Element, chain: list[etree._Element]
) -> tuple[tuple[float, float], tuple[float, float]]:
    """Current (start, end) slide-space endpoints derived from the stored
    xfrm box and flips."""
    xfrm = _require_xfrm(cxnsp)
    x, y, cx, cy = _xfrm_box(xfrm)
    flip_h = xfrm.get("flipH") == "1"
    flip_v = xfrm.get("flipV") == "1"
    sx = x + (cx if flip_h else 0)
    sy = y + (cy if flip_v else 0)
    ex = x + (0 if flip_h else cx)
    ey = y + (0 if flip_v else cy)
    ax, bx, ay, by, _rot = _chain_transform(chain)
    return (
        (ax * sx + bx, ay * sy + by),
        (ax * ex + bx, ay * ey + by),
    )


def _iter_connectors(container, chain=()):
    for child in container:
        if child.tag == qn("p:cxnSp"):
            yield child, list(chain)
        elif child.tag == qn("p:grpSp"):
            yield from _iter_connectors(child, (*chain, child))


def reroute_connectors(pkg: PptxPackage, part: str, moved_ids: set[int]) -> list[int]:
    """Re-derive the xfrm of every connector glued to any shape in moved_ids
    (stale boxes render as detached connectors until the user nudges them).
    Returns the rerouted connector ids."""
    rerouted: list[int] = []
    sp_tree = _sp_tree(pkg, part)
    for cxnsp, chain in _iter_connectors(sp_tree):
        cnvcxn = cxnsp.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}")
        if cnvcxn is None:
            continue
        st = cnvcxn.find(qn("a:stCxn"))
        en = cnvcxn.find(qn("a:endCxn"))
        st_id = int(st.get("id")) if st is not None else None
        en_id = int(en.get("id")) if en is not None else None
        if not ({st_id, en_id} & moved_ids):
            continue
        cur_start, cur_end = _connector_endpoints_slide(cxnsp, chain)
        p1, p2 = cur_start, cur_end
        try:
            if st_id is not None:
                p1x, p1y, _idx, _exact = _site_point(
                    pkg, part, st_id, int(st.get("idx", "0")), None
                )
                p1 = (p1x, p1y)
            if en_id is not None:
                p2x, p2y, _idx, _exact = _site_point(
                    pkg, part, en_id, int(en.get("idx", "0")), None
                )
                p2 = (p2x, p2y)
        except (UnsupportedStructure, TargetNotFound, PptMcpError):
            # Endpoint shape has no derivable sites (e.g. a PowerPoint-made
            # freeform); leave the connector's current box alone.
            continue
        _set_connector_box(cxnsp, p1, p2, chain)
        cid = _shape_id(cxnsp)
        if cid is not None:
            rerouted.append(cid)
    if rerouted:
        pkg.mark_dirty(part)
    return rerouted


# ------------------------------------------------------------ edit-in-place


def set_shape(
    pkg: PptxPackage,
    slide,
    shape: int,
    *,
    x: float | None = None,
    y: float | None = None,
    dx: float | None = None,
    dy: float | None = None,
    w: float | None = None,
    h: float | None = None,
    rotation: float | None = None,
    flip_h: bool | None = None,
    flip_v: bool | None = None,
    fill=None,
    line=None,
    effect=None,
    text: str | None = None,
    text_style: dict | None = None,
    name: str | None = None,
) -> dict:
    """Edit one shape in place by id: move (absolute x/y or delta dx/dy,
    inches, SLIDE space even inside groups), resize (w/h inches), rotate,
    flip, restyle (fill/line/effect specs per ops/geometry.py; pass "none"
    to remove), replace text (simple single-style), rename. Connectors glued
    to the shape are rerouted automatically. Only the parameters given are
    touched."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    elem, chain = _find_shape(pkg, part, shape)
    changed: list[str] = []
    warnings: list[str] = []
    preserved: list[str] = []
    text_facts: dict = {}

    if (x is not None or y is not None) and (dx is not None or dy is not None):
        raise PptMcpError("use absolute x/y or delta dx/dy, not both")

    geo_change = any(v is not None for v in (x, y, dx, dy, w, h))
    if geo_change and _xfrm_of(elem) is None:
        seeded = _seed_xfrm(pkg, part, elem)
        if seeded is not None:
            warnings.append(
                f"shape {shape} had no explicit geometry; seeded its xfrm "
                "from the box it inherited before applying the change"
            )
    if geo_change:
        xfrm = _require_xfrm(elem)
        ax, bx, ay, by, rotated = _chain_transform(chain)
        if rotated:
            warnings.append(
                "an ancestor group carries rotation or flip; slide-space "
                "coordinate math ignores it, verify the result visually"
            )
        cur_x, cur_y, cur_cx, cur_cy = _xfrm_box(xfrm)
        slide_x = ax * cur_x + bx
        slide_y = ay * cur_y + by
        if x is not None:
            slide_x = g.in_to_emu(x)
        if y is not None:
            slide_y = g.in_to_emu(y)
        if dx is not None:
            slide_x += g.in_to_emu(dx)
        if dy is not None:
            slide_y += g.in_to_emu(dy)
        new_x = round((slide_x - bx) / ax)
        new_y = round((slide_y - by) / ay)
        new_cx = round(g.in_to_emu(w) / ax) if w is not None else cur_cx
        new_cy = round(g.in_to_emu(h) / ay) if h is not None else cur_cy
        if new_cx < 0 or new_cy < 0:
            raise PptMcpError("w and h must be positive inches")
        g.check_emu_box(new_x, new_y, new_cx, new_cy, what=f"shape {shape}")
        off = xfrm.find(qn("a:off"))
        ext = xfrm.find(qn("a:ext"))
        off.set("x", str(new_x))
        off.set("y", str(new_y))
        ext.set("cx", str(new_cx))
        ext.set("cy", str(new_cy))
        changed.append("geometry")
        if elem.tag == qn("p:grpSp"):
            warnings.append(
                "group resized: children scale through ext/chExt, positions "
                "were not rewritten"
            )
    if rotation is not None or flip_h is not None or flip_v is not None:
        xfrm = _require_xfrm(elem)
        if rotation is not None:
            if rotation % 360 == 0:
                xfrm.attrib.pop("rot", None)
            else:
                xfrm.set("rot", str(g.deg_to_60000(rotation % 360)))
        for flag, attr in ((flip_h, "flipH"), (flip_v, "flipV")):
            if flag is not None:
                if flag:
                    xfrm.set(attr, "1")
                else:
                    xfrm.attrib.pop(attr, None)
        changed.append("transform")

    if fill is not None:
        if elem.tag == qn("p:cxnSp"):
            raise PptMcpError(
                "connectors have no fill; restyle the line instead"
            )
        g.insert_spPr_child(_spPr_of(elem), g.fill_element(fill))
        changed.append("fill")
    if line is not None:
        if elem.tag == qn("p:grpSp"):
            raise PptMcpError(
                "groups have no outline of their own; restyle the member "
                "shapes instead"
            )
        g.insert_spPr_child(_spPr_of(elem), g.line_element(line))
        changed.append("line")
    if effect is not None:
        g.insert_spPr_child(_spPr_of(elem), g.effect_element(effect))
        changed.append("effect")

    if text is not None:
        if elem.tag != qn("p:sp"):
            raise PptMcpError(
                f"shape {shape} is not a text-capable autoshape; text applies "
                "to p:sp shapes only"
            )
        old = elem.find(qn("p:txBody"))
        base_style, extra_style = _split_text_style(text_style)
        new_body = g.txbody(text, base_style)
        if old is not None:
            named = set(base_style or {}) | set(extra_style)
            carried, carry_facts = _carry_text_properties(
                old, new_body, named
            )
            if carried:
                preserved = carried
            text_facts.update(carry_facts)
            elem.replace(old, new_body)
        else:
            elem.append(new_body)
        _apply_extra_run_props(new_body, extra_style)
        changed.append("text")
    elif text_style is not None:
        raise PptMcpError(
            "text_style only applies together with text (single-style "
            "replace); rich in-place restyling is the text pack's job"
        )

    if name is not None:
        cnvpr = _cnvpr(elem)
        if cnvpr is None:
            raise UnsupportedStructure(f"shape {shape} has no cNvPr to rename")
        cnvpr.set("name", name)
        changed.append("name")

    if not changed:
        raise PptMcpError("set_shape called with nothing to change")
    pkg.mark_dirty(part)
    rerouted: list[int] = []
    if "geometry" in changed or "transform" in changed:
        moved = {shape}
        if elem.tag == qn("p:grpSp"):
            moved |= {
                i for i in (
                    _shape_id(e) for e, _k, _z, _p in iter_shapes(elem)
                ) if i is not None
            }
        rerouted = reroute_connectors(pkg, part, moved)
    result = {
        "shape_id": shape,
        "changed": changed,
        "changed_ids": [shape],
        "rerouted_connectors": rerouted,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
    }
    if preserved:
        result["preserved"] = preserved
    # What the carry could NOT preserve, stated rather than left to be
    # discovered in a render.
    result.update(text_facts)
    if warnings:
        result["warnings"] = warnings
    return result


def delete_shape(pkg: PptxPackage, slide, shape: int) -> dict:
    """Delete one shape by id. Connectors glued to it lose that glue (the
    stCxn/endCxn reference is dropped, the connector keeps its geometry),
    matching PowerPoint's own delete behavior. Timing nodes and build
    entries referencing the deleted ids are pruned in the same pass (a
    dangling spid makes PowerPoint silently repair the slide); the counts
    surface as timing_report."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    elem, _chain = _find_shape(pkg, part, shape)
    deleted_ids = {shape}
    if elem.tag == qn("p:grpSp"):
        deleted_ids |= {
            i for i in (
                _shape_id(e) for e, _k, _z, _p in iter_shapes(elem)
            ) if i is not None
        }
    elem.getparent().remove(elem)
    unglued: list[int] = []
    for cxnsp, _chain2 in _iter_connectors(_sp_tree(pkg, part)):
        cnvcxn = cxnsp.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}")
        if cnvcxn is None:
            continue
        for tag in ("a:stCxn", "a:endCxn"):
            ref = cnvcxn.find(qn(tag))
            if ref is not None and int(ref.get("id", "-1")) in deleted_ids:
                cnvcxn.remove(ref)
                cid = _shape_id(cxnsp)
                if cid is not None and cid not in unglued:
                    unglued.append(cid)
    from .animations import _prune_spids

    timing_report = _prune_spids(pkg.root(part), deleted_ids)
    pkg.mark_dirty(part)
    return {
        "deleted": sorted(deleted_ids),
        "unglued_connectors": unglued,
        "timing_report": timing_report,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
    }


# ------------------------------------------------------------------ groups


def _top_level_members(
    pkg: PptxPackage, part: str, ids: list[int]
) -> list[etree._Element]:
    if not ids or len(ids) < 2:
        raise PptMcpError("group_shapes needs at least 2 shape ids")
    if len(set(ids)) != len(ids):
        raise PptMcpError(f"duplicate shape ids in {ids}")
    sp_tree = _sp_tree(pkg, part)
    elems = []
    for sid in ids:
        elem, chain = _find_shape(pkg, part, sid)
        if chain:
            raise UnsupportedStructure(
                f"shape {sid} is already inside group {_shape_id(chain[-1])}; "
                "ungroup first or group at the top level"
            )
        if elem.getparent() is not sp_tree:
            raise UnsupportedStructure(f"shape {sid} is not a top-level shape")
        elems.append(elem)
    return elems


def group_shapes(pkg: PptxPackage, slide, ids: list[int], *, name: str | None = None) -> dict:
    """Group top-level shapes into one p:grpSp. The group frame is the union
    bounding box; chOff/chExt are written IDENTITY-equal to off/ext so the
    children keep their coordinates and visual positions exactly."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    elems = _top_level_members(pkg, part, ids)
    boxes = [_xfrm_box(_require_xfrm(e)) for e in elems]
    min_x = min(b[0] for b in boxes)
    min_y = min(b[1] for b in boxes)
    max_x = max(b[0] + b[2] for b in boxes)
    max_y = max(b[1] + b[3] for b in boxes)

    sp_tree = _sp_tree(pkg, part)
    group_id = pkg.next_shape_id(part)
    display = name or f"Group {group_id}"
    grp = etree.Element(qn("p:grpSp"))
    nv = _nv_pr(grp, "p:nvGrpSpPr", group_id, display)
    etree.SubElement(nv, qn("p:cNvGrpSpPr"))
    etree.SubElement(nv, qn("p:nvPr"))
    grppr = etree.SubElement(grp, qn("p:grpSpPr"))
    grppr.append(
        g.xfrm_element(
            min_x, min_y, max_x - min_x, max_y - min_y,
            tag="a:xfrm",
            ch_off=(min_x, min_y),
            ch_ext=(max_x - min_x, max_y - min_y),
        )
    )
    # The group takes the z-position of the front-most member.
    anchor = max(elems, key=lambda e: list(sp_tree).index(e))
    anchor.addnext(grp)
    for elem in sorted(elems, key=lambda e: list(sp_tree).index(e)):
        sp_tree.remove(elem)
        grp.append(elem)
    pkg.mark_dirty(part)
    return {
        "group_id": group_id,
        "created": [group_id],
        "member_ids": list(ids),
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "name": display,
    }


def ungroup_shapes(pkg: PptxPackage, slide, group: int) -> dict:
    """Dissolve one group: children get the group transform baked into their
    xfrm (position and scale) so visual positions are preserved, then move up
    to the group's parent at the group's z-position. Rotated or flipped
    groups are refused (baking rotation into children is lossy)."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    elem, _chain = _find_shape(pkg, part, group)
    if elem.tag != qn("p:grpSp"):
        raise PptMcpError(f"shape {group} is not a group")
    xfrm = elem.find(f"{qn('p:grpSpPr')}/{qn('a:xfrm')}")
    if xfrm is None:
        raise UnsupportedStructure(f"group {group} has no transform")
    if xfrm.get("rot") or xfrm.get("flipH") or xfrm.get("flipV"):
        raise UnsupportedStructure(
            f"group {group} is rotated or flipped; ungroup it in PowerPoint "
            "(baking rotation into children is not supported)"
        )
    ax, bx, ay, by, _rot = _chain_transform([elem])
    parent = elem.getparent()
    freed: list[int] = []
    children = [c for c in elem if c.tag in _SHAPE_TAGS]
    marker = elem
    for child in children:
        cxfrm = _xfrm_of(child)
        if cxfrm is not None:
            cx, cy, ccx, ccy = _xfrm_box(cxfrm)
            off = cxfrm.find(qn("a:off"))
            ext = cxfrm.find(qn("a:ext"))
            off.set("x", str(round(ax * cx + bx)))
            off.set("y", str(round(ay * cy + by)))
            ext.set("cx", str(round(ax * ccx)))
            ext.set("cy", str(round(ay * ccy)))
            if child.tag == qn("p:grpSp"):
                # Child group frame scaled; chOff/chExt untouched keeps its
                # interior mapping consistent.
                pass
        elem.remove(child)
        marker.addnext(child)
        marker = child
        cid = _shape_id(child)
        if cid is not None:
            freed.append(cid)
    parent.remove(elem)
    pkg.mark_dirty(part)
    return {
        "ungrouped": group,
        "freed_ids": freed,
        "changed_ids": freed,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
    }


# -------------------------------------------------------- align / distribute


_ALIGN_MODES = ("left", "center", "right", "top", "middle", "bottom")


def _same_parent_shapes(
    pkg: PptxPackage, part: str, ids: list[int], minimum: int
) -> tuple[list[etree._Element], list[etree._Element]]:
    if len(set(ids)) != len(ids):
        raise PptMcpError(f"duplicate shape ids in {ids}")
    if len(ids) < minimum:
        raise PptMcpError(f"need at least {minimum} shape ids, got {len(ids)}")
    elems = []
    chains = []
    for sid in ids:
        elem, chain = _find_shape(pkg, part, sid)
        elems.append(elem)
        chains.append(chain)
    parents = {id(e.getparent()) for e in elems}
    if len(parents) > 1:
        raise UnsupportedStructure(
            "shapes live in different containers (mixed group membership); "
            "align/distribute within one container at a time"
        )
    return elems, chains[0]


def align_shapes(
    pkg: PptxPackage, slide, ids: list[int], mode: str, *, to: str = "selection"
) -> dict:
    """Align shapes: left | center | right | top | middle | bottom, against
    the selection bounds (default) or the slide (to="slide", top-level
    shapes only). Glued connectors reroute automatically."""
    if mode not in _ALIGN_MODES:
        raise PptMcpError(
            f"unknown align mode {mode!r}; one of: {', '.join(_ALIGN_MODES)}"
        )
    if to not in ("selection", "slide"):
        raise PptMcpError(f'align target must be "selection" or "slide", got {to!r}')
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    elems, chain = _same_parent_shapes(pkg, part, ids, 1 if to == "slide" else 2)
    boxes = [_xfrm_box(_require_xfrm(e)) for e in elems]
    if to == "slide":
        if chain:
            raise UnsupportedStructure(
                'to="slide" needs top-level shapes (these are inside a group)'
            )
        pres = pkg.presentation()
        sldsz = pres.find(qn("p:sldSz"))
        if sldsz is None:
            raise UnsupportedStructure("presentation has no p:sldSz")
        lo_x, lo_y = 0, 0
        hi_x, hi_y = int(sldsz.get("cx")), int(sldsz.get("cy"))
    else:
        lo_x = min(b[0] for b in boxes)
        lo_y = min(b[1] for b in boxes)
        hi_x = max(b[0] + b[2] for b in boxes)
        hi_y = max(b[1] + b[3] for b in boxes)
    moved: list[int] = []
    for elem, box in zip(elems, boxes):
        bx, by, bcx, bcy = box
        nx, ny = bx, by
        if mode == "left":
            nx = lo_x
        elif mode == "center":
            nx = round((lo_x + hi_x) / 2 - bcx / 2)
        elif mode == "right":
            nx = hi_x - bcx
        elif mode == "top":
            ny = lo_y
        elif mode == "middle":
            ny = round((lo_y + hi_y) / 2 - bcy / 2)
        elif mode == "bottom":
            ny = hi_y - bcy
        if (nx, ny) != (bx, by):
            xfrm = _require_xfrm(elem)
            off = xfrm.find(qn("a:off"))
            off.set("x", str(int(nx)))
            off.set("y", str(int(ny)))
            sid = _shape_id(elem)
            if sid is not None:
                moved.append(sid)
    if moved:
        pkg.mark_dirty(part)
    rerouted = reroute_connectors(pkg, part, set(moved)) if moved else []
    return {
        "mode": mode,
        "to": to,
        "changed_ids": moved,
        "rerouted_connectors": rerouted,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
    }


def distribute_shapes(pkg: PptxPackage, slide, ids: list[int], axis: str = "h") -> dict:
    """Distribute shapes with even gaps along axis "h" or "v". The first and
    last shape (by position) stay put; the shapes between are respaced so
    every gap is equal. Needs at least 3 shapes."""
    if axis not in ("h", "v"):
        raise PptMcpError(f'axis must be "h" or "v", got {axis!r}')
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    elems, _chain = _same_parent_shapes(pkg, part, ids, 3)
    ai = 0 if axis == "h" else 1
    si = 2 if axis == "h" else 3
    items = sorted(
        ((e, _xfrm_box(_require_xfrm(e))) for e in elems),
        key=lambda pair: pair[1][ai],
    )
    span_lo = items[0][1][ai]
    span_hi = items[-1][1][ai] + items[-1][1][si]
    total_size = sum(box[si] for _e, box in items)
    gap = (span_hi - span_lo - total_size) / (len(items) - 1)
    moved: list[int] = []
    cursor = float(span_lo)
    for i, (elem, box) in enumerate(items):
        target = round(cursor)
        interior = 0 < i < len(items) - 1
        if interior and target != box[ai]:
            xfrm = _require_xfrm(elem)
            off = xfrm.find(qn("a:off"))
            off.set("x" if axis == "h" else "y", str(target))
            sid = _shape_id(elem)
            if sid is not None:
                moved.append(sid)
        cursor += box[si] + gap
    if moved:
        pkg.mark_dirty(part)
    rerouted = reroute_connectors(pkg, part, set(moved)) if moved else []
    return {
        "axis": axis,
        "changed_ids": moved,
        "rerouted_connectors": rerouted,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
    }


# ----------------------------------------------------------------- z-order


_Z_ACTIONS = ("front", "back", "forward", "backward")


def set_z_order(pkg: PptxPackage, slide, shape: int, action: str) -> dict:
    """Move a shape in the stacking order of its container: front | back |
    forward | backward. Document order IS z-order (later = on top)."""
    if action not in _Z_ACTIONS:
        raise PptMcpError(
            f"unknown z-order action {action!r}; one of: {', '.join(_Z_ACTIONS)}"
        )
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    elem, _chain = _find_shape(pkg, part, shape)
    parent = elem.getparent()
    siblings = [c for c in parent if c.tag in _SHAPE_TAGS]
    idx = siblings.index(elem)
    moved = False
    if action == "front" and idx < len(siblings) - 1:
        siblings[-1].addnext(elem)
        moved = True
    elif action == "back" and idx > 0:
        siblings[0].addprevious(elem)
        moved = True
    elif action == "forward" and idx < len(siblings) - 1:
        siblings[idx + 1].addnext(elem)
        moved = True
    elif action == "backward" and idx > 0:
        siblings[idx - 1].addprevious(elem)
        moved = True
    if moved:
        pkg.mark_dirty(part)
    new_siblings = [c for c in parent if c.tag in _SHAPE_TAGS]
    return {
        "shape_id": shape,
        "action": action,
        "changed_ids": [shape] if moved else [],
        "z": new_siblings.index(elem),
        "of": len(new_siblings),
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
    }
