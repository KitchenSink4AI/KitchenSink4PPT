"""Text engine: placeholder population, text boxes, run-level formatting,
bullets, deck-wide search and replace, and autofit state reporting.

Contract (all ops modules): every function takes the open PptxPackage first,
mutates only the in-memory package, calls pkg.mark_dirty() on every part it
touches, and returns a summary dict. Nothing here writes to disk; the caller
decides when to save. Errors come from the core.errors taxonomy with
actionable messages. First-match is forbidden: any selector matching more
than one target refuses, listing the candidates.

Addressing:
- slide: 0-based presentation-order index (int) or {"slide_id": N}.
- shape: the p:cNvPr id (int, unique per slide; digits as text too), the
  shape name (str, searched through groups; refuses when several shapes
  share the name), "Group/Shape", or the view anchor "a:3f9c2a" from
  get_presentation_view. One resolver for all of them: shapes._find_shape.
- placeholder: a type string ("title", "subtitle", "body", "content", or a
  raw p:ph type value) or an idx int (the p:ph idx, 0 when absent).
- character ranges are PER PARAGRAPH (paragraph index + start/end offsets),
  matching what find_text reports; edits resolve through ops/_runmap.py,
  never per-run (a match can span several fragmented runs).

UNIT CONVENTION (one rule, everywhere a position or size is accepted):
floats are INCHES; ints below 10000 are INCHES; ints of 10000 or more are
EMU (no real slide coordinate expressed in inches reaches 10000). Pass
unit="in" or unit="emu" to make it explicit.

Fields (a:fld, slide numbers/dates) render cached text that PowerPoint
recomputes; edits overlapping a field refuse rather than rewrite the cache.
"""

from __future__ import annotations

import copy as _copy
import math
import re as _stdlib_re

from lxml import etree

from ..core import budget as _budget
from ..core.errors import (
    AmbiguousTarget,
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from ..core.package import PptxPackage, qn, resolve_target, seat_ahead_of_extlst
from ..core.xmlbool import xml_bool
from . import _regex
from . import inherit as _inh
from ._casenote import CaseTally
from ._runmap import (
    BULLET_CHOICE_TAGS,
    FILL_CHOICE_TAGS,
    PPR_ORDER,
    RPR_ORDER,
    build_map,
    ensure_child,
    ensure_pPr,
    ensure_rPr,
    rank_insert,
    remove_children,
    replace_range,
    split_for_range,
)
from ._textcontent import append_text as _append_paragraph_text
from .read import (
    _dgm,
    chart_text_records,
    in_wrapper,
    iter_shapes,
    notes_bodies,
    notes_part_for,
    paragraph_text,
    resolve_slide,
    shape_text,
    slides_in_scope,
    table_element,
    txbody_paragraphs,
)
from .read import cnvpr_id

EMU_PER_INCH = 914400
EMU_PER_POINT = 12700

#: ST_SchemeColorVal tokens accepted wherever a color is accepted.
SCHEME_COLOR_TOKENS = frozenset(
    {
        "bg1",
        "tx1",
        "bg2",
        "tx2",
        "dk1",
        "lt1",
        "dk2",
        "lt2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hlink",
        "folHlink",
        "phClr",
    }
)

#: ST_TextAutonumberScheme values for set_bullets(style="autonum").
AUTONUM_TYPES = frozenset(
    {
        "alphaLcParenBoth",
        "alphaUcParenBoth",
        "alphaLcParenR",
        "alphaUcParenR",
        "alphaLcPeriod",
        "alphaUcPeriod",
        "arabicParenBoth",
        "arabicParenR",
        "arabicPeriod",
        "arabicPlain",
        "romanLcParenBoth",
        "romanUcParenBoth",
        "romanLcParenR",
        "romanUcParenR",
        "romanLcPeriod",
        "romanUcPeriod",
        "circleNumDbPlain",
        "circleNumWdBlackPlain",
        "circleNumWdWhitePlain",
        "arabicDbPeriod",
        "arabicDbPlain",
        "ea1ChsPeriod",
        "ea1ChsPlain",
        "ea1ChtPeriod",
        "ea1ChtPlain",
        "ea1JpnChsDbPeriod",
        "ea1JpnKorPlain",
        "ea1JpnKorPeriod",
        "arabic1Minus",
        "arabic2Minus",
        "hebrew2Minus",
        "thaiAlphaPeriod",
        "thaiAlphaParenR",
        "thaiAlphaParenBoth",
        "thaiNumPeriod",
        "thaiNumParenR",
        "thaiNumParenBoth",
        "hindiAlphaPeriod",
        "hindiNumPeriod",
        "hindiNumParenR",
        "hindiAlpha1Period",
    }
)

_ALIGN_VALUES = {
    "left": "l",
    "center": "ctr",
    "right": "r",
    "justify": "just",
    "distribute": "dist",
    "l": "l",
    "ctr": "ctr",
    "r": "r",
    "just": "just",
    "dist": "dist",
    "justLow": "justLow",
    "thaiDist": "thaiDist",
}

#: a:bodyPr @anchor, in the words the rest of the surface uses (the same
#: vocabulary geometry.txbody's text_style takes).
_ANCHOR_VALUES = {"top": "t", "middle": "ctr", "center": "ctr", "bottom": "b"}

_UNDERLINE_VALUES = {
    "none",
    "sng",
    "dbl",
    "heavy",
    "dotted",
    "dottedHeavy",
    "dash",
    "dashHeavy",
    "dashLong",
    "dashLongHeavy",
    "dotDash",
    "dotDashHeavy",
    "dotDotDash",
    "dotDotDashHeavy",
    "wavy",
    "wavyHeavy",
    "wavyDbl",
    "words",
}

#: Placeholder type aliases (values are the p:ph type strings they match,
#: FIRST entry = the preferred exact type). "body" and "content" are one
#: alias set: a Title-and-Content layout's content placeholder is type
#: "obj", and agents asking for "body" used to dead-end on NOT_FOUND.
#: When a slide carries both types, the exact
#: type wins; several same-type matches still refuse as ambiguous.
_PH_ALIASES = {
    "title": ("title", "ctrTitle"),
    "subtitle": ("subTitle",),
    "body": ("body", "obj"),
    "content": ("obj", "body"),
}

#: Shape kinds that carry no directly editable text body (find_text parity).
_NO_TEXT_KINDS = ("group", "picture", "media", "chart", "diagram", "ole", "graphicFrame")

_HEX_COLOR = _stdlib_re.compile(r"\A#?([0-9A-Fa-f]{6})\Z")

_AUTOFIT_CAVEAT = (
    "normAutofit fontScale/lnSpcReduction are a rendering cache written by "
    "PowerPoint at the last edit, not an instruction it recomputes on open; "
    "PowerPoint re-fits only when the frame is edited again, so cached "
    "values may not reflect the current text."
)


# ------------------------------------------------------------------- units


def _to_emu(value, unit: str | None, what: str) -> int:
    """The module unit convention: floats are inches; ints below 10000 are
    inches; ints of 10000 or more are EMU. unit='in'/'emu' overrides."""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise PptMcpError(
            f"{what} must be a number (float/int inches, or int EMU); got "
            f"{value!r}"
        )
    # Overflow guard: a non-finite value, or a float
    # multiply that overflows to inf (values near 1e308), must refuse as
    # BAD_PARAMS, never escape as a raw OverflowError from int().
    if isinstance(value, float) and not math.isfinite(value):
        raise PptMcpError(f"{what} = {value!r} is not a finite number")

    def _scaled(v) -> int:
        emu = v * EMU_PER_INCH
        if isinstance(emu, float) and not math.isfinite(emu):
            raise PptMcpError(
                f"{what} = {v!r} inches is not a representable coordinate"
            )
        return int(round(emu))

    if unit is not None:
        if unit == "emu":
            return int(round(value))
        if unit == "in":
            return _scaled(value)
        raise PptMcpError(f"unit must be 'in' or 'emu', got {unit!r}")
    if isinstance(value, float):
        return _scaled(value)
    if value >= 10000:
        return value
    return value * EMU_PER_INCH


def _emu_to_in(emu: int) -> float:
    return round(emu / EMU_PER_INCH, 3)


# ------------------------------------------------------- shape resolution


def _sp_tree(pkg: PptxPackage, part: str) -> etree._Element | None:
    return pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")


def _cnvpr_of(elem: etree._Element) -> etree._Element | None:
    for child in elem:
        if etree.QName(child).localname.startswith("nv"):
            return child.find(qn("p:cNvPr"))
    return None


def _ph_of(sp: etree._Element) -> etree._Element | None:
    nv = sp.find(qn("p:nvSpPr"))
    if nv is None:
        return None
    nvpr = nv.find(qn("p:nvPr"))
    return nvpr.find(qn("p:ph")) if nvpr is not None else None


def _shape_id(elem: etree._Element) -> int | None:
    cnvpr = _cnvpr_of(elem)
    return cnvpr_id(cnvpr)


def _shape_name(elem: etree._Element) -> str:
    cnvpr = _cnvpr_of(elem)
    return cnvpr.get("name", "") if cnvpr is not None else ""


def _resolve_shape(
    pkg: PptxPackage, rec: dict, shape
) -> tuple[etree._Element, str]:
    """(shape element, kind) for a shape selector on one slide record,
    through the shape operations' own resolver (shapes._find_shape), so
    every operation that takes a shape takes the same addresses: the
    p:cNvPr id (digits as text too, when no shape has them as its name),
    the shape name (searched through groups; a repeated name refuses with
    every candidate), "Group/Shape", or the view anchor
    get_presentation_view prints ("a:3f9c2a", brackets optional)."""
    from .read import _shape_kind
    from .shapes import _find_shape

    if _sp_tree(pkg, rec["part"]) is None:
        raise TargetNotFound(f"slide {rec['index']} has no shape tree")
    if isinstance(shape, bool) or not isinstance(shape, (int, str)):
        raise PptMcpError(
            "shape must be a shape id, a shape name, \"Group/Shape\" or a "
            f"view anchor such as a:3f9c2a; got {shape!r}"
        )
    elem, _chain = _find_shape(pkg, rec["part"], shape)
    return elem, _shape_kind(elem)


def _require_txbody(
    elem: etree._Element, kind: str, rec: dict, *, create: bool = False
) -> etree._Element:
    if etree.QName(elem).localname != "sp":
        extra = (
            " Table cell text is reachable through search_and_replace; "
            "per-cell text and formatting go through set_table_cells."
            if kind == "table"
            else ""
        )
        exc = UnsupportedStructure(
            f"shape {_shape_id(elem)} on slide {rec['index']} is a {kind} "
            f"and has no directly editable text body.{extra}"
        )
        if kind == "table":
            exc.hint_tools = ["set_table_cells"]
        raise exc
    body = elem.find(qn("p:txBody"))
    if body is None:
        if not create:
            raise UnsupportedStructure(
                f"shape {_shape_id(elem)} on slide {rec['index']} has no "
                "text body"
            )
        body = etree.SubElement(elem, qn("p:txBody"))  # txBody is last in p:sp
        etree.SubElement(body, qn("a:bodyPr"))
        etree.SubElement(body, qn("a:lstStyle"))
        # A p:txBody with zero a:p violates the content model and makes
        # PowerPoint refuse the whole deck, naming nothing. Same shape
        # ops/masters.py builds.
        etree.SubElement(body, qn("a:p"))
    return body


# --------------------------------------------------------------- formatting


def _fill_element(color: str) -> etree._Element:
    """a:solidFill carrying srgbClr (hex) or schemeClr (theme token)."""
    fill = etree.Element(qn("a:solidFill"))
    m = _HEX_COLOR.match(color) if isinstance(color, str) else None
    if m:
        clr = etree.SubElement(fill, qn("a:srgbClr"))
        clr.set("val", m.group(1).upper())
        return fill
    if color in SCHEME_COLOR_TOKENS:
        clr = etree.SubElement(fill, qn("a:schemeClr"))
        clr.set("val", color)
        return fill
    raise PptMcpError(
        f"color must be a 6-digit hex string ('1F4E79' or '#1F4E79') or a "
        f"theme token ({', '.join(sorted(SCHEME_COLOR_TOKENS))}); got "
        f"{color!r}"
    )


def _apply_run_props(
    rpr: etree._Element,
    *,
    font: str | None = None,
    size_pt: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | str | None = None,
    color: str | None = None,
) -> None:
    """Write character properties onto one a:rPr, schema-order aware."""
    if size_pt is not None:
        if not 1 <= size_pt <= 4000:
            raise PptMcpError(
                f"size_pt must be between 1 and 4000 points, got {size_pt!r}"
            )
        rpr.set("sz", str(int(round(size_pt * 100))))
    if bold is not None:
        rpr.set("b", "1" if bold else "0")
    if italic is not None:
        rpr.set("i", "1" if italic else "0")
    if underline is not None:
        if isinstance(underline, bool):
            rpr.set("u", "sng" if underline else "none")
        elif underline in _UNDERLINE_VALUES:
            rpr.set("u", underline)
        else:
            raise PptMcpError(
                f"underline must be a bool or one of "
                f"{', '.join(sorted(_UNDERLINE_VALUES))}; got {underline!r}"
            )
    if color is not None:
        remove_children(rpr, FILL_CHOICE_TAGS)  # fill choice is exclusive
        rank_insert(rpr, _fill_element(color), RPR_ORDER)
    if font is not None:
        # CJK glyphs render through a:ea (and complex scripts through a:cs),
        # so a latin-only write silently leaves Hangul/Kanji in the theme
        # font. Mirror the typeface into all three slots.
        for tag in ("a:latin", "a:ea", "a:cs"):
            el = ensure_child(rpr, tag, RPR_ORDER)
            el.set("typeface", font)


def _apply_paragraph_props(
    p: etree._Element,
    *,
    align: str | None = None,
    line_spacing: float | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
) -> None:
    if (
        align is None
        and line_spacing is None
        and space_before is None
        and space_after is None
    ):
        return
    ppr = ensure_pPr(p)
    if align is not None:
        if align not in _ALIGN_VALUES:
            raise PptMcpError(
                f"align must be one of left, center, right, justify, "
                f"distribute; got {align!r}"
            )
        ppr.set("algn", _ALIGN_VALUES[align])
    if line_spacing is not None:
        if not 0 < line_spacing <= 1000:
            raise PptMcpError(
                f"line_spacing must be positive (a multiple of single "
                f"spacing when 10 or less, points when above 10); got "
                f"{line_spacing!r}"
            )
        lnspc = ensure_child(ppr, "a:lnSpc", PPR_ORDER)
        for child in list(lnspc):
            lnspc.remove(child)
        if line_spacing <= 10:
            spc = etree.SubElement(lnspc, qn("a:spcPct"))
            spc.set("val", str(int(round(line_spacing * 100000))))
        else:
            spc = etree.SubElement(lnspc, qn("a:spcPts"))
            spc.set("val", str(int(round(line_spacing * 100))))
    for name, tag, value in (
        ("space_before", "a:spcBef", space_before),
        ("space_after", "a:spcAft", space_after),
    ):
        if value is None:
            continue
        if isinstance(value, bool) or not 0 <= value <= 1584:
            raise PptMcpError(
                f"{name} must be 0..1584 points; got {value!r}"
            )
        holder = ensure_child(ppr, tag, PPR_ORDER)
        for child in list(holder):
            holder.remove(child)
        pts = etree.SubElement(holder, qn("a:spcPts"))
        pts.set("val", str(int(round(value * 100))))


# ---------------------------------------------------------- paragraph build


def _parse_text_paragraphs(text: str) -> list[dict]:
    """'\\n' separates paragraphs; leading tabs set the outline level
    (one tab per level, 0..8). No tab is a statement too: it is level 0."""
    out = []
    for line in text.split("\n"):
        level = 0
        while line.startswith("\t") and level < 8:
            level += 1
            line = line[1:]
        # EVERY line of this shorthand states its level, and a line with no
        # tab states level 0. Treating "no tab" as "said nothing" let a
        # replacement keep the nesting of the paragraph it replaced, so
        # "Top level\n\tSecond level\nTop again" came back from PowerPoint
        # at levels 1, 2, 2.
        out.append({"text": line, "level": level, "level_explicit": True})
    return out


def _normalize_paragraphs(paragraphs) -> list[dict]:
    out = []
    for i, item in enumerate(paragraphs):
        if not isinstance(item, dict) or "text" not in item:
            raise PptMcpError(
                f"paragraphs[{i}] must be a dict with 'text' (and optional "
                f"'level' 0..8); got {item!r}"
            )
        level = item.get("level", 0)
        if not isinstance(level, int) or not 0 <= level <= 8:
            raise PptMcpError(
                f"paragraphs[{i}] level must be an int 0..8, got {level!r}"
            )
        out.append({
            "text": str(item["text"]),
            "level": level,
            # A stated level, including a stated 0, is an instruction and
            # outranks the level carried from the paragraph it replaces.
            "level_explicit": "level" in item,
        })
    return out


def _build_paragraph(spec: dict, run_props: dict | None = None) -> etree._Element:
    p = etree.Element(qn("a:p"))
    if spec["level"] or spec.get("level_explicit"):
        ppr = etree.SubElement(p, qn("a:pPr"))
        ppr.set("lvl", str(spec["level"]))
    if spec["text"]:
        configure = None
        if run_props:
            configure = lambda rpr: _apply_run_props(rpr, **run_props)
        _append_paragraph_text(p, spec["text"], configure)
    return p


def _landing_spots(body: etree._Element) -> None:
    """Give every freshly built paragraph the empty a:pPr / a:rPr /
    a:endParaRPr that the carry needs to write the old properties INTO.

    _carry_text_properties only fills elements that already exist (it was
    written against geometry.txbody, which always emits them). A bare
    paragraph built here has none, so without this the carry silently did
    nothing and paragraph formatting kept dying. Anything the carry leaves
    empty is removed again by _strip_empty_props.
    """
    from .shapes import _RUN_FAMILY

    for p in body.findall(qn("a:p")):
        if p.find(qn("a:pPr")) is None:
            p.insert(0, etree.Element(qn("a:pPr")))
        runs = [el for holder in _RUN_FAMILY for el in p.findall(qn(holder))]
        for r in runs:
            ensure_rPr(r)
        if not runs and p.find(qn("a:endParaRPr")) is None:
            # An empty line still has a height, and that lives on
            # endParaRPr. endParaRPr is last in a:p.
            p.append(etree.Element(qn("a:endParaRPr")))


def _strip_empty_props(body: etree._Element) -> None:
    """Drop the landing spots nothing was carried into, so a body that had
    no formatting to keep comes out exactly as it used to."""
    from .shapes import _RUN_FAMILY

    for p in body.findall(qn("a:p")):
        for tag in ("a:pPr", "a:endParaRPr"):
            el = p.find(qn(tag))
            if el is not None and not len(el) and not el.attrib:
                p.remove(el)
        for holder in _RUN_FAMILY:
            for r in p.findall(qn(holder)):
                rpr = r.find(qn("a:rPr"))
                if rpr is not None and not len(rpr) and not rpr.attrib:
                    r.remove(rpr)


def _replace_body_paragraphs(
    body: etree._Element,
    specs: list[dict],
    run_props: dict | None = None,
    *,
    preserve: bool = True,
) -> tuple[int, list[str], dict]:
    """Replace the body's paragraphs with `specs`, carrying the formatting
    of the paragraphs being replaced onto the new ones.

    Returns (paragraph count, preserved categories, facts). The rebuild
    used to emit bare a:p / a:r, which threw away marL, indent, spcAft,
    the bullet definition, defRPr and every run's size, colour and
    typeface: a body whose formatting lived at paragraph level came back
    unstyled. The carry is the one in
    ops/shapes.py, so both text-writing routes preserve the same things
    and report them the same way.
    """
    from .shapes import _carry_text_properties

    old = _copy.deepcopy(body) if preserve else None
    for p in body.findall(qn("a:p")):
        body.remove(p)
    for spec in specs:
        body.append(_build_paragraph(spec, run_props))
    if not specs:
        # A text body with zero a:p is exactly what PowerPoint refuses to
        # open; an empty replacement leaves one empty paragraph instead.
        body.append(etree.Element(qn("a:p")))
    preserved: list[str] = []
    facts: dict = {}
    if old is not None and old.findall(qn("a:p")):
        _landing_spots(body)
        preserved, facts = _carry_text_properties(
            old, body, set(run_props or {})
        )
        _strip_empty_props(body)
    return len(specs), preserved, facts


def _restore_layout_placeholder(
    pkg: PptxPackage,
    slide_part: str,
    slide_tree: etree._Element,
    selector,
) -> tuple[etree._Element, str, int] | None:
    """Recreate one deleted slide placeholder from its layout skeleton.

    The new shape carries only identity and placeholder binding. Geometry
    and formatting remain inherited from the layout, matching PowerPoint's
    Reset behavior and the skeleton used for a newly inserted slide.
    """
    if not (
        isinstance(selector, str)
        or (isinstance(selector, int) and not isinstance(selector, bool))
    ):
        raise PptMcpError(
            "placeholder selector must be a type string or an idx int; "
            f"got {selector!r}"
        )
    layout_part = None
    try:
        rels = pkg.rels_for(slide_part)
    except KeyError:
        return None
    for rel in rels.getroot():
        if rel.get("Type", "").endswith("/slideLayout") and (
            rel.get("TargetMode") != "External"
        ):
            layout_part = resolve_target(slide_part, rel.get("Target", ""))
            break
    if layout_part is None or not pkg.has_part(layout_part):
        return None
    layout_tree = _sp_tree(pkg, layout_part)
    if layout_tree is None:
        return None
    found: list[tuple[etree._Element, str, int]] = []
    for shape in layout_tree.findall(qn("p:sp")):
        ph = _ph_of(shape)
        if ph is None:
            continue
        ph_type = ph.get("type", "obj")
        ph_idx = int(ph.get("idx", "0"))
        if isinstance(selector, int):
            match = ph_idx == selector
        else:
            match = ph_type in _PH_ALIASES.get(selector, (selector,))
        if match:
            found.append((shape, ph_type, ph_idx))
    if len(found) > 1 and selector in ("body", "content"):
        primary = _PH_ALIASES[selector][0]
        exact = [item for item in found if item[1] == primary]
        if exact:
            found = exact
    if len(found) > 1:
        listing = ", ".join(f"type={t} idx={i}" for _s, t, i in found)
        raise AmbiguousTarget(
            f"{len(found)} layout placeholders match {selector!r}: "
            f"{listing}. Address by idx (int) instead."
        )
    if not found:
        return None

    layout_shape, ph_type, ph_idx = found[0]
    source_ph = _ph_of(layout_shape)
    shape_id = pkg.next_shape_id(slide_part)
    new_shape = seat_ahead_of_extlst(slide_tree, etree.SubElement(slide_tree, qn("p:sp")))
    nvsp = etree.SubElement(new_shape, qn("p:nvSpPr"))
    cnvpr = etree.SubElement(nvsp, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    source_name = _cnvpr_of(layout_shape)
    cnvpr.set(
        "name",
        (source_name.get("name") if source_name is not None else None)
        or f"{ph_type} Placeholder {shape_id}",
    )
    cnvsp = etree.SubElement(nvsp, qn("p:cNvSpPr"))
    etree.SubElement(cnvsp, qn("a:spLocks")).set("noGrp", "1")
    nvpr = etree.SubElement(nvsp, qn("p:nvPr"))
    new_ph = etree.SubElement(nvpr, qn("p:ph"))
    for attr in ("type", "orient", "sz", "idx"):
        value = source_ph.get(attr) if source_ph is not None else None
        if value is not None:
            new_ph.set(attr, value)
    etree.SubElement(new_shape, qn("p:spPr"))
    body = etree.SubElement(new_shape, qn("p:txBody"))
    etree.SubElement(body, qn("a:bodyPr"))
    etree.SubElement(body, qn("a:lstStyle"))
    etree.SubElement(body, qn("a:p"))
    pkg.mark_dirty(slide_part)
    return new_shape, ph_type, ph_idx


# ========================================================== public API


def set_placeholder_text(
    pkg: PptxPackage,
    slide,
    placeholder,
    text: str | None = None,
    *,
    paragraphs: list[dict] | None = None,
) -> dict:
    """Replace a placeholder's text content. `placeholder` targets by type
    ("title" also matches ctrTitle; "subtitle"; "body" and "content" are
    one alias set matching both body and obj placeholders, with the exact
    type preferred when a slide has both; or a raw p:ph type value) or by
    idx (int; a p:ph without idx counts as 0). Content: `text` with '\\n'
    paragraph breaks, '\\v' soft line breaks, and leading tabs for bullet
    levels, or
    `paragraphs=[{"text": ..., "level": 0..8}, ...]`. Existing paragraphs
    are fully replaced, but their formatting is carried onto the new text
    (indents, spacing, bullet definition, run size, colour and typeface)
    and reported as `preserved`; bodyPr and lstStyle are untouched. A
    stated level wins over the carried one. Several matching placeholders
    refuse, listing candidates for idx addressing."""
    if (text is None) == (paragraphs is None):
        raise PptMcpError(
            "pass exactly one of text (str) or paragraphs (list of dicts)"
        )
    rec = resolve_slide(pkg, slide)
    sp_tree = _sp_tree(pkg, rec["part"])
    if sp_tree is None:
        raise TargetNotFound(f"slide {rec['index']} has no shape tree")

    candidates: list[tuple[etree._Element, str, int]] = []
    inventory: list[str] = []
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        if kind != "placeholder":
            continue
        ph = _ph_of(elem)
        ph_type = ph.get("type", "obj")
        ph_idx = int(ph.get("idx", "0"))
        inventory.append(f"type={ph_type} idx={ph_idx} (shape {_shape_id(elem)})")
        if isinstance(placeholder, int) and not isinstance(placeholder, bool):
            if ph_idx == placeholder:
                candidates.append((elem, ph_type, ph_idx))
        elif isinstance(placeholder, str):
            wanted = _PH_ALIASES.get(placeholder, (placeholder,))
            if ph_type in wanted:
                candidates.append((elem, ph_type, ph_idx))
        else:
            raise PptMcpError(
                "placeholder selector must be a type string or an idx int; "
                f"got {placeholder!r}"
            )
    restored_from_layout = False
    if not candidates:
        restored = _restore_layout_placeholder(
            pkg, rec["part"], sp_tree, placeholder
        )
        if restored is not None:
            candidates = [restored]
            restored_from_layout = True
        else:
            raise TargetNotFound(
                f"no placeholder {placeholder!r} on slide {rec['index']}; "
                f"placeholders present: {', '.join(inventory) or 'none'}, "
                "and its layout has no unique matching placeholder to restore"
            )
    if len(candidates) > 1 and placeholder in ("body", "content"):
        # The body/content alias set matches both types; when the slide
        # carries both, the asked-for exact type wins instead of refusing.
        primary = _PH_ALIASES[placeholder][0]
        exact = [c for c in candidates if c[1] == primary]
        if exact:
            candidates = exact
    if len(candidates) > 1:
        listing = ", ".join(f"type={t} idx={i}" for _e, t, i in candidates)
        raise AmbiguousTarget(
            f"{len(candidates)} placeholders on slide {rec['index']} match "
            f"{placeholder!r}: {listing}. Address by idx (int) instead."
        )

    elem, ph_type, ph_idx = candidates[0]
    specs = (
        _parse_text_paragraphs(text)
        if text is not None
        else _normalize_paragraphs(paragraphs)
    )
    body = _require_txbody(elem, "placeholder", rec, create=True)
    count, preserved, facts = _replace_body_paragraphs(body, specs)
    pkg.mark_dirty(rec["part"])
    out = {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": _shape_id(elem),
        "placeholder_type": ph_type,
        "placeholder_idx": ph_idx,
        "paragraphs": count,
        "characters": sum(len(s["text"]) for s in specs),
        "restored_from_layout": restored_from_layout,
    }
    if preserved:
        out["preserved"] = preserved
    out.update(facts)
    return out


def insert_textbox(
    pkg: PptxPackage,
    slide,
    text: str,
    x,
    y,
    w,
    h,
    *,
    unit: str | None = None,
    name: str | None = None,
    font: str | None = None,
    size_pt: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | str | None = None,
    color: str | None = None,
    align: str | None = None,
    wrap: bool = True,
) -> dict:
    """New text box at (x, y) sized (w, h). Units per the module rule:
    floats and small ints are inches, ints of 10000+ are EMU; unit='in' or
    'emu' overrides. Text splits on '\\n' with leading tabs as levels; '\\v'
    inserts a soft line break inside the current paragraph.
    Optional run formatting applies to every created run; `align` to every
    paragraph. The box gets prstGeom rect, noFill, and square wrapping
    (wrap=False disables wrapping)."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    sp_tree = _sp_tree(pkg, part)
    if sp_tree is None:
        raise TargetNotFound(f"slide {rec['index']} has no shape tree")
    x_emu = _to_emu(x, unit, "x")
    y_emu = _to_emu(y, unit, "y")
    w_emu = _to_emu(w, unit, "w")
    h_emu = _to_emu(h, unit, "h")
    if w_emu <= 0 or h_emu <= 0:
        raise PptMcpError(
            f"textbox size must be positive; got w={w_emu} EMU, h={h_emu} EMU"
        )
    from . import geometry as _g

    _g.check_emu_box(x_emu, y_emu, w_emu, h_emu, what="textbox")

    shape_id = pkg.next_shape_id(part)
    box_name = name or f"TextBox {shape_id - 1}"

    sp = seat_ahead_of_extlst(sp_tree, etree.SubElement(sp_tree, qn("p:sp")))
    nvsp = etree.SubElement(sp, qn("p:nvSpPr"))
    cnv = etree.SubElement(nvsp, qn("p:cNvPr"))
    cnv.set("id", str(shape_id))
    cnv.set("name", box_name)
    cnvsp = etree.SubElement(nvsp, qn("p:cNvSpPr"))
    cnvsp.set("txBox", "1")
    etree.SubElement(nvsp, qn("p:nvPr"))

    sppr = etree.SubElement(sp, qn("p:spPr"))
    xfrm = etree.SubElement(sppr, qn("a:xfrm"))
    off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(x_emu))
    off.set("y", str(y_emu))
    ext = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(w_emu))
    ext.set("cy", str(h_emu))
    geom = etree.SubElement(sppr, qn("a:prstGeom"))
    geom.set("prst", "rect")
    etree.SubElement(geom, qn("a:avLst"))
    etree.SubElement(sppr, qn("a:noFill"))

    body = etree.SubElement(sp, qn("p:txBody"))
    bodypr = etree.SubElement(body, qn("a:bodyPr"))
    bodypr.set("wrap", "square" if wrap else "none")
    bodypr.set("rtlCol", "0")
    etree.SubElement(body, qn("a:lstStyle"))

    run_props = {
        k: v
        for k, v in (
            ("font", font),
            ("size_pt", size_pt),
            ("bold", bold),
            ("italic", italic),
            ("underline", underline),
            ("color", color),
        )
        if v is not None
    }
    specs = _parse_text_paragraphs(text)
    for spec in specs:
        p = _build_paragraph(spec, run_props or None)
        _apply_paragraph_props(p, align=align)
        body.append(p)

    pkg.mark_dirty(part)
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": shape_id,
        "name": box_name,
        "paragraphs": len(specs),
        "geometry": {
            "x": x_emu,
            "y": y_emu,
            "cx": w_emu,
            "cy": h_emu,
            "x_in": _emu_to_in(x_emu),
            "y_in": _emu_to_in(y_emu),
            "cx_in": _emu_to_in(w_emu),
            "cy_in": _emu_to_in(h_emu),
        },
    }


def format_text(
    pkg: PptxPackage,
    slide,
    shape,
    *,
    paragraph: int | None = None,
    start: int | None = None,
    end: int | None = None,
    font: str | None = None,
    size_pt: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | str | None = None,
    color: str | None = None,
    align: str | None = None,
    line_spacing: float | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    anchor: str | None = None,
    wrap: bool | None = None,
) -> dict:
    """Run-level formatting over a whole shape, one paragraph, or a
    character range within one paragraph (start/end require `paragraph`;
    find_text reports paragraph indexes and offsets). Fragmented runs are
    split at the range boundaries, cloning their rPr, so formatting never
    bleeds outside the range. `color` accepts hex ('1F4E79') or a theme
    token ('accent1', which stays theme-linked as schemeClr). align and
    line_spacing, space_before and space_after are paragraph properties
    and apply to the touched paragraphs (space_before/space_after are
    points, 0..1584, written as a:spcBef/a:spcAft). rPr children are written in schema order.

    anchor (top|middle|bottom) and wrap are a:bodyPr properties of the
    SHAPE, so they apply to the whole text frame whatever paragraph or
    range is selected. They live here because the only other route to them
    was set_shape's text_style, which replaces the whole body with one flat
    string and so flattens bullet levels, per-paragraph alignment and
    per-run formatting to change one attribute."""
    run_props = {
        k: v
        for k, v in (
            ("font", font),
            ("size_pt", size_pt),
            ("bold", bold),
            ("italic", italic),
            ("underline", underline),
            ("color", color),
        )
        if v is not None
    }
    if (
        not run_props
        and align is None
        and line_spacing is None
        and space_before is None
        and space_after is None
        and anchor is None
        and wrap is None
    ):
        raise PptMcpError(
            "nothing to do: pass at least one of font, size_pt, bold, "
            "italic, underline, color, align, line_spacing, space_before, "
            "space_after, anchor, wrap"
        )
    if anchor is not None and anchor not in _ANCHOR_VALUES:
        raise PptMcpError(
            f"anchor must be one of {', '.join(sorted(_ANCHOR_VALUES))}; "
            f"got {anchor!r}"
        )
    if (start is None) != (end is None):
        raise PptMcpError("start and end must be given together")
    if start is not None and paragraph is None:
        raise PptMcpError(
            "character ranges are per paragraph: pass paragraph= with "
            "start/end (find_text reports paragraph indexes and offsets)"
        )

    rec = resolve_slide(pkg, slide)
    elem, kind = _resolve_shape(pkg, rec, shape)
    body = _require_txbody(elem, kind, rec)
    paras = body.findall(qn("a:p"))
    if not paras:
        exc = TargetNotFound(
            f"shape {_shape_id(elem)} on slide {rec['index']} has an empty "
            "text body; add text first (set_placeholder_text/insert_textbox)"
        )
        raise exc
    if paragraph is not None:
        if not 0 <= paragraph < len(paras):
            raise TargetNotFound(
                f"paragraph {paragraph} out of range; shape "
                f"{_shape_id(elem)} has {len(paras)} paragraphs"
            )
        targets = [paras[paragraph]]
    else:
        targets = paras

    runs_formatted = 0
    for p in targets:
        if start is not None:
            plain = paragraph_text(p)
            if not 0 <= start < end <= len(plain):
                raise TargetNotFound(
                    f"range [{start}, {end}) out of bounds; paragraph "
                    f"{paragraph} has {len(plain)} characters"
                )
            try:
                covered = split_for_range(p, start, end)
            except ValueError as exc:
                raise TargetNotFound(str(exc)) from exc
            for el in covered:
                if run_props:
                    _apply_run_props(ensure_rPr(el), **run_props)
                runs_formatted += 1
        else:
            for child in p:
                if etree.QName(child).localname in ("r", "br", "fld"):
                    if run_props:
                        _apply_run_props(ensure_rPr(child), **run_props)
                    runs_formatted += 1
        _apply_paragraph_props(
            p, align=align, line_spacing=line_spacing,
            space_before=space_before, space_after=space_after,
        )

    frame_changed: list[str] = []
    if anchor is not None or wrap is not None:
        bodypr = body.find(qn("a:bodyPr"))
        if bodypr is None:
            bodypr = etree.Element(qn("a:bodyPr"))
            body.insert(0, bodypr)
        if anchor is not None:
            bodypr.set("anchor", _ANCHOR_VALUES[anchor])
            frame_changed.append("anchor")
        if wrap is not None:
            bodypr.set("wrap", "square" if wrap else "none")
            frame_changed.append("wrap")

    pkg.mark_dirty(rec["part"])
    out = {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": _shape_id(elem),
        "paragraphs": len(targets),
        "runs_formatted": runs_formatted,
    }
    if frame_changed:
        out["frame_changed"] = frame_changed
    return out


def set_bullets(
    pkg: PptxPackage,
    slide,
    shape,
    style: str,
    *,
    paragraphs: int | list[int] | None = None,
    char: str = "•",
    char_font: str = "Arial",
    num_type: str = "arabicPeriod",
    start_at: int | None = None,
    level: int | None = None,
    size_pct: float | None = None,
    color: str | None = None,
) -> dict:
    """Per-paragraph bullet overrides on one shape. style: "char" (literal
    bullet `char` in `char_font`), "autonum" (numbered, `num_type` from the
    ST_TextAutonumberScheme values, optional `start_at`), or "none".
    `paragraphs` selects indexes (int or list; None = all). Existing bullet
    properties on the touched paragraphs are replaced. Optional size_pct
    (25..400, percent of text size), color (hex or theme token), and level
    (0..8, which also writes hanging-indent geometry: marL = 0.25 in + 0.5 in
    per level, indent = -0.25 in, so the bullet does not sit at the text
    edge). With level omitted, a char or autonum bullet on a plain text box
    still gets that hanging geometry at the paragraph's own level (0 when
    none), unless the paragraph already carries its own marL; placeholders
    keep the indent they inherit from the layout. style "none" drops any
    marL/indent from the paragraph so it inherits again. Writes per-paragraph a:pPr overrides only; shape
    and master lstStyle defaults are never touched."""
    aliases = {"char": "char", "autonum": "autonum", "number": "autonum", "none": "none"}
    if style not in aliases:
        raise PptMcpError(
            f"style must be 'char', 'autonum' (alias 'number'), or 'none'; "
            f"got {style!r}"
        )
    style = aliases[style]
    if style == "autonum" and num_type not in AUTONUM_TYPES:
        raise PptMcpError(
            f"num_type {num_type!r} is not a known autonumber scheme; "
            f"valid: {', '.join(sorted(AUTONUM_TYPES))}"
        )
    if start_at is not None and not 1 <= start_at <= 32767:
        raise PptMcpError(f"start_at must be 1..32767, got {start_at!r}")
    if size_pct is not None and not 25 <= size_pct <= 400:
        raise PptMcpError(
            f"size_pct must be 25..400 (percent of text size), got {size_pct!r}"
        )
    if level is not None and not (isinstance(level, int) and 0 <= level <= 8):
        raise PptMcpError(f"level must be an int 0..8, got {level!r}")

    rec = resolve_slide(pkg, slide)
    elem, kind = _resolve_shape(pkg, rec, shape)
    body = _require_txbody(elem, kind, rec)
    is_placeholder = _ph_of(elem) is not None
    paras = body.findall(qn("a:p"))
    if not paras:
        raise TargetNotFound(
            f"shape {_shape_id(elem)} on slide {rec['index']} has an empty "
            "text body; nothing to bullet"
        )
    if paragraphs is None:
        indexes = list(range(len(paras)))
    else:
        indexes = [paragraphs] if isinstance(paragraphs, int) else list(paragraphs)
        for i in indexes:
            if not (isinstance(i, int) and 0 <= i < len(paras)):
                raise TargetNotFound(
                    f"paragraph index {i!r} out of range; shape "
                    f"{_shape_id(elem)} has {len(paras)} paragraphs"
                )

    bullet_family = (
        "a:buClrTx",
        "a:buClr",
        "a:buSzTx",
        "a:buSzPct",
        "a:buSzPts",
        "a:buFontTx",
        "a:buFont",
    ) + BULLET_CHOICE_TAGS

    for i in indexes:
        ppr = ensure_pPr(paras[i])
        remove_children(ppr, bullet_family)
        if style == "none":
            # The glyph goes; so does the hanging indent an earlier bullet
            # call wrote, or the text would stay pushed in with no bullet.
            for attr in ("marL", "indent"):
                if attr in ppr.attrib:
                    del ppr.attrib[attr]
        elif level is None and not is_placeholder and ppr.get("marL") is None:
            # A bullet with no hanging indent sits glued to the text. A
            # plain text box inherits marL=0, so give it the level-0 (or
            # its own lvl) geometry; placeholders inherit a real indent
            # from the layout and are left alone.
            try:
                own = int(ppr.get("lvl", "0"))
            except ValueError:
                own = 0
            own = min(max(own, 0), 8)
            ppr.set("marL", str(228600 + 457200 * own))
            ppr.set("indent", "-228600")
        if level is not None:
            ppr.set("lvl", str(level))
            if style != "none":
                # Default hanging-indent geometry (quarter inch hang, half
                # inch per level) so an explicit level does not leave the
                # bullet glued to the text edge.
                ppr.set("marL", str(228600 + 457200 * level))
                ppr.set("indent", "-228600")
        if style == "none":
            rank_insert(ppr, etree.Element(qn("a:buNone")), PPR_ORDER)
            continue
        if color is not None:
            buclr = etree.Element(qn("a:buClr"))
            buclr.append(_fill_element(color)[0])  # the srgbClr/schemeClr child
            rank_insert(ppr, buclr, PPR_ORDER)
        if size_pct is not None:
            szpct = etree.Element(qn("a:buSzPct"))
            szpct.set("val", str(int(round(size_pct * 1000))))
            rank_insert(ppr, szpct, PPR_ORDER)
        if style == "char":
            bufont = etree.Element(qn("a:buFont"))
            bufont.set("typeface", char_font)
            rank_insert(ppr, bufont, PPR_ORDER)
            buchar = etree.Element(qn("a:buChar"))
            buchar.set("char", char)
            rank_insert(ppr, buchar, PPR_ORDER)
        else:  # autonum
            aunum = etree.Element(qn("a:buAutoNum"))
            aunum.set("type", num_type)
            if start_at is not None:
                aunum.set("startAt", str(start_at))
            rank_insert(ppr, aunum, PPR_ORDER)

    pkg.mark_dirty(rec["part"])
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": _shape_id(elem),
        "style": style,
        "paragraphs_updated": len(indexes),
    }


# ------------------------------------------------------- search and replace


def _compile_matcher(find: str, replace: str, regex: bool, match_case: bool):
    """Returns a callable text -> list of (start, end, replacement),
    left-to-right and non-overlapping. Regex goes through the ReDoS guard;
    literal patterns use stdlib re on an escaped pattern (safe)."""
    if regex:
        pattern = find if match_case else "(?i)" + find

        def _matches(text: str) -> list[tuple[int, int, str]]:
            out = []
            for m in _regex.finditer(pattern, text):
                s, e = m.span()
                if s == e:
                    continue  # zero-width matches are meaningless replacements
                try:
                    out.append((s, e, m.expand(replace)))
                except Exception as exc:
                    raise PptMcpError(
                        f"replacement template {replace!r} failed against "
                        f"match {m.group(0)!r}: {exc}"
                    ) from exc
            return out

        return _matches

    flags = 0 if match_case else _stdlib_re.IGNORECASE
    compiled = _stdlib_re.compile(_stdlib_re.escape(find), flags)

    def _matches(text: str) -> list[tuple[int, int, str]]:
        return [(m.start(), m.end(), replace) for m in compiled.finditer(text)]

    return _matches


def _replace_in_paragraph(
    p: etree._Element, matcher, where: str, skipped: list | None = None
) -> int:
    text, segments = build_map(p)
    if not text:
        return 0
    spans = matcher(text)
    if not spans:
        return 0
    wrapped = [seg for seg in segments if seg.kind == "wrapped"]
    if wrapped:
        # An equation (or a run for a newer version) in a compatibility
        # wrapper is never rewritten: a match touching one is left and
        # counted in `skipped`, and the others go ahead.
        kept = [m for m in spans if not any(w.start < m[1] and w.end > m[0] for w in wrapped)]
        if skipped is not None and len(kept) < len(spans):
            skipped.append(len(spans) - len(kept))
        spans = kept
        if not spans:
            return 0
    # Snapshot once, apply right-to-left: offsets left of each edit stay
    # valid and replacement output can never be re-matched.
    for s, e, rep in reversed(spans):
        if "\n" in rep:
            raise UnsupportedStructure(
                f"{where}: a search-and-replace match is confined to one "
                "paragraph, so its replacement cannot insert a paragraph "
                "break. Use '\\v' for a soft line break, or rewrite the "
                "shape/placeholder/cell text to add paragraphs."
            )
        try:
            replace_range(p, segments, s, e, rep)
        except UnsupportedStructure as exc:
            raise UnsupportedStructure(f"{where}: {exc}") from exc
        # A replacement holding a soft break rebuilds the runs it touched,
        # so the next match to the left (it may share one of them) resolves
        # through a fresh map. Text left of this edit did not move, so the
        # snapshot's offsets for the remaining matches still hold.
        _text, segments = build_map(p)
    return len(spans)


def _slide_paragraph_sites(pkg: PptxPackage, rec: dict):
    """Yield (a:p, where-description) for every editable paragraph on one
    slide, mirroring find_text's traversal (shapes in reading order, groups
    recursed, table cells row by row)."""
    sp_tree = _sp_tree(pkg, rec["part"])
    if sp_tree is None:
        return
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        sid = _shape_id(elem)
        if kind == "table":
            tbl = table_element(elem)
            if tbl is None:
                continue
            for r, tr in enumerate(tbl.findall(qn("a:tr"))):
                for c, tc in enumerate(tr.findall(qn("a:tc"))):
                    for p in tc.findall(f"{qn('a:txBody')}/{qn('a:p')}"):
                        yield p, (
                            f"slide {rec['index']} table shape {sid} "
                            f"cell r{r + 1}c{c + 1}"
                        ), rec["part"]
        elif kind in _NO_TEXT_KINDS:
            continue
        else:
            for p in txbody_paragraphs(elem):
                yield p, f"slide {rec['index']} shape {sid}", rec["part"]


def _wrapped_slide_paragraphs(pkg: PptxPackage, rec: dict):
    """Yield every paragraph of the shapes PowerPoint keeps in a slide-level
    compatibility wrapper on one slide (a text box or table saved with an
    equation), from the branch PowerPoint shows: the paragraphs the reads
    show and _slide_paragraph_sites, which edits walk, does not reach."""
    sp_tree = _sp_tree(pkg, rec["part"])
    if sp_tree is None:
        return
    for elem, kind, _z, _parent in iter_shapes(sp_tree, wrapped=True):
        if not in_wrapper(elem):
            continue
        if kind == "table":
            tbl = table_element(elem)
            if tbl is not None:
                yield from tbl.iter(qn("a:p"))
        elif kind not in _NO_TEXT_KINDS:
            yield from txbody_paragraphs(elem)


def _diagram_paragraph_sites(pkg: PptxPackage, rec: dict, seen: set[str] | None = None):
    """Yield (a:p, where, part) for every SmartArt paragraph reachable from
    one slide: the data model AND the cached drawing, because a replacement
    that lands in one and not the other leaves the old words on screen for
    every viewer that does not regenerate the diagram. A part named in
    `seen` is skipped and every part walked is added to it, so parts that
    two frames share are walked once per call."""
    from .diagrams import _dsp, diagram_frames, drawing_part_of

    seen = set() if seen is None else seen
    for _frame, data_part in diagram_frames(pkg, rec["part"]):
        if data_part is None or not pkg.has_part(data_part) or data_part in seen:
            continue
        seen.add(data_part)
        for t in pkg.root(data_part).iter(_dgm("t")):
            for p in t.findall(qn("a:p")):
                yield p, f"slide {rec['index']} SmartArt", data_part
        drawing = drawing_part_of(pkg, data_part, rec["part"])
        if drawing is None or drawing in seen:
            continue
        seen.add(drawing)
        for body in pkg.root(drawing).iter(_dsp("txBody")):
            for p in body.findall(qn("a:p")):
                yield p, f"slide {rec['index']} SmartArt (cached)", drawing


def _smartart_sharers_outside(pkg: PptxPackage, rec: dict, in_scope: set[str]) -> list[int]:
    """Indexes of slides outside the scope that reach any SmartArt part this
    slide reaches (decks copied by older tools share them)."""
    from .slides import _diagram_graph

    mine = _diagram_graph(pkg, rec["part"])[0]
    if not mine:
        return []
    return [
        other["index"]
        for other in slides_in_scope(pkg, None)
        if other["part"] not in in_scope and mine & _diagram_graph(pkg, other["part"])[0]
    ]


def _give_slide_own_smartart(pkg: PptxPackage, slide_part: str) -> None:
    """Copy every SmartArt part the slide reaches and point the slide at the
    copies, as duplicate_slide does for a duplicated slide."""
    from ..core.package import rels_name
    from .slides import _clone_diagram_bundle

    rels_part = rels_name(slide_part)
    _clone_diagram_bundle(pkg, slide_part, slide_part, pkg.root(rels_part))
    pkg.mark_dirty(rels_part)


def search_and_replace(
    pkg: PptxPackage,
    find: str,
    replace: str,
    *,
    scope=None,
    regex: bool = False,
    match_case: bool = True,
    include_notes: bool = False,
    remove_empty_paragraphs: bool = False,
) -> dict:
    """Deck-wide (or scoped) text replacement, fragmented-run safe: matches
    resolve through the run map, so a hit spanning three runs is one edit
    that lands in the first run's formatting. Spans are snapshotted per
    paragraph and applied right-to-left, so a replacement containing the
    search text is never re-matched ('partner' -> 'partner-network'
    cannot loop). Matches cannot cross paragraph boundaries. regex=True runs
    the pattern through the ReDoS guard and supports group references in
    `replace`; match_case=False matches case-insensitively. With
    match_case=True (the default) text in another case is never changed,
    and when some was left, case_note says how much, on which slides, and
    that match_case=False includes it (ops/_casenote.py). A match
    overlapping a field (slide number/date) refuses the whole call and
    nothing is written to disk. ``\v`` in a replacement emits a soft line
    break; ``\n`` refuses because one match cannot span or create paragraph
    boundaries. Returns per-slide counts and the total.
    Chart labels are intentionally read-only; matching labels are counted
    separately in chart_matches_skipped and matches_found. So are matches
    inside an equation or other content PowerPoint stores in a compatibility
    wrapper (wrapped_matches_skipped, with a warning naming the slides).

    remove_empty_paragraphs=True removes each paragraph a replacement left
    with no text but spaces and line breaks, in slide shapes, table cells and
    notes (not SmartArt, whose data model and drawing would disagree). A
    paragraph that was empty before keeps its place, and a text body keeps
    at least one paragraph."""
    if not find:
        raise PptMcpError("search_and_replace needs a non-empty find string")
    matcher = _compile_matcher(find, replace, regex, match_case)
    tally = CaseTally(find, regex) if match_case else None

    def _probe(p, slide_index: int) -> None:
        """Count other-case matches in one paragraph before it is changed,
        leaving out those in an equation or other wrapped content, which a
        replace in any case leaves too (wrapped_matches_skipped)."""
        if tally is not None:
            text, segments = build_map(p)
            if text:
                wrapped = [(seg.start, seg.end) for seg in segments if seg.kind == "wrapped"]
                tally.add(slide_index, text, matcher(text), skip=wrapped)

    slides_out: list[dict] = []
    total = 0
    skipped_total = 0
    removed_total = 0
    warnings: list[str] = []
    wrapped_skips: list[int] = []
    wrapped_slides: list[int] = []
    recs = slides_in_scope(pkg, scope)
    in_scope = {r["part"] for r in recs}
    seen_diagram_parts: set[str] = set()
    for rec in recs:
        count = 0
        emptied: list = []
        skips_before = sum(wrapped_skips)
        for p, where, _part in _slide_paragraph_sites(pkg, rec):
            _probe(p, rec["index"])
            hits = _replace_in_paragraph(p, matcher, where, wrapped_skips)
            count += hits
            if hits and remove_empty_paragraphs and _blank(p):
                emptied.append(p)
        # A shape PowerPoint keeps in a slide-level wrapper (a text box saved
        # with an equation): find_text finds its words, and no edit rewrites
        # a wrapper, so its matches are left and counted as skipped.
        for p in _wrapped_slide_paragraphs(pkg, rec):
            found = len(matcher(build_map(p)[0]))
            if found:
                wrapped_skips.append(found)
        sharers = _smartart_sharers_outside(pkg, rec, in_scope)
        if sharers and any(
            matcher(build_map(p)[0]) for p, _w, _d in _diagram_paragraph_sites(pkg, rec)
        ):
            _give_slide_own_smartart(pkg, rec["part"])
            listed = ", ".join(str(i) for i in sharers)
            warnings.append(
                f"slide {rec['index']}: its SmartArt shared parts with slide(s) "
                f"{listed}, outside the scope; slide {rec['index']} now has its "
                "own copy, so the replacement does not reach those slides"
            )
        diagram_count = 0
        diagram_parts: set[str] = set()
        for p, where, dpart in _diagram_paragraph_sites(pkg, rec, seen_diagram_parts):
            _probe(p, rec["index"])
            hits = _replace_in_paragraph(p, matcher, where, wrapped_skips)
            if hits:
                diagram_count += hits
                diagram_parts.add(dpart)
        for dpart in diagram_parts:
            pkg.mark_dirty(dpart)
        chart_matches_skipped = 0
        sp_tree = _sp_tree(pkg, rec["part"])
        if sp_tree is not None:
            for elem, kind, _z, _parent in iter_shapes(sp_tree):
                if kind != "chart":
                    continue
                for record in chart_text_records(pkg, rec["part"], elem):
                    chart_matches_skipped += len(matcher(record["text"]))
        notes_count = 0
        if include_notes:
            npart = notes_part_for(pkg, rec["part"])
            if npart is not None and pkg.has_part(npart):
                # PowerPoint keeps a notes body that holds an equation in a
                # wrapper, with a copy in its Fallback. Both copies are
                # changed, so they stay in step; only the body PowerPoint
                # shows is counted.
                shown = notes_bodies(pkg, npart)
                notes_changed = False
                for sp in pkg.root(npart).iter(qn("p:sp")):
                    ph = _ph_of(sp)
                    if ph is None or ph.get("type") != "body":
                        continue
                    counted = any(sp is body for body in shown)
                    for p in txbody_paragraphs(sp):
                        if counted:
                            _probe(p, rec["index"])
                        hits = _replace_in_paragraph(
                            p, matcher, f"slide {rec['index']} notes",
                            wrapped_skips if counted else [],
                        )
                        if counted:
                            notes_count += hits
                        notes_changed = notes_changed or bool(hits)
                        if hits and remove_empty_paragraphs and _blank(p):
                            emptied.append(p)
                if notes_changed:
                    pkg.mark_dirty(npart)
        removed = _drop_paragraphs(emptied)
        removed_total += removed
        wrapped_here = sum(wrapped_skips) - skips_before
        if wrapped_here:
            wrapped_slides.append(rec["index"] + 1)
        if count:
            pkg.mark_dirty(rec["part"])
        if count or notes_count or diagram_count or chart_matches_skipped or wrapped_here:
            entry = {
                "slide_index": rec["index"],
                "slide_id": rec["slide_id"],
                "count": count,
            }
            if include_notes:
                entry["notes_count"] = notes_count
            if diagram_count:
                entry["diagram_count"] = diagram_count
            if chart_matches_skipped:
                entry["chart_matches_skipped"] = chart_matches_skipped
            if wrapped_here:
                entry["wrapped_matches_skipped"] = wrapped_here
            if remove_empty_paragraphs:
                entry["paragraphs_removed"] = removed
            slides_out.append(entry)
        total += count + notes_count + diagram_count
        skipped_total += chart_matches_skipped

    out = {
        "find": find,
        "replace": replace,
        "regex": regex,
        "match_case": match_case,
        "notes_included": include_notes,
        "total": total,
        "matches_found": total + skipped_total + sum(wrapped_skips),
        "chart_matches_skipped": skipped_total,
        "slides": slides_out,
    }
    if remove_empty_paragraphs:
        out["paragraphs_removed"] = removed_total
    if wrapped_skips:
        skipped_here = sum(wrapped_skips)
        out["wrapped_matches_skipped"] = skipped_here
        listed = ", ".join(str(n) for n in wrapped_slides)
        warnings.append(
            f"{skipped_here} match{'es' if skipped_here != 1 else ''} on slide_number "
            f"{listed} sit in an equation or other content PowerPoint stores in "
            "a compatibility wrapper, and were left unchanged; edit them in PowerPoint"
        )
    case_note = tally.note("replace", total) if tally is not None else None
    if case_note:
        out["case_note"] = case_note
    if warnings:
        out["warnings"] = warnings
    return out


def _blank(p: etree._Element) -> bool:
    """True when a paragraph shows no text: nothing but spaces and breaks."""
    text, _segments = build_map(p)
    return not text.replace("\v", "").strip()


def _drop_paragraphs(paragraphs: list) -> int:
    """Remove the paragraphs a replacement emptied, keeping at least one in
    each text body (the schema needs one, and PowerPoint refuses a body
    without). Returns how many were removed."""
    removed = 0
    for p in paragraphs:
        body = p.getparent()
        if body is not None and len(body.findall(qn("a:p"))) > 1:
            body.remove(p)
            removed += 1
    return removed


# --------------------------------------------------------------- autofit


def _pct_value(raw: str | None, default: float) -> float:
    """Autofit percent attribute: integer thousandths ('62500') or the
    later-edition percent string ('62.5%'). Read both."""
    if raw is None:
        return default
    raw = raw.strip()
    if raw.endswith("%"):
        return float(raw[:-1])
    return int(raw) / 1000.0


def _ins(bodypr: etree._Element | None, attr: str, default: int) -> int:
    if bodypr is None:
        return default
    raw = bodypr.get(attr)
    return int(raw) if raw is not None else default


def _first_run_size(body: etree._Element) -> int | None:
    """First explicit sz (centipoints) among the body's rPr/defRPr, or None
    (size inherited from the placeholder chain)."""
    for tag in ("a:rPr", "a:defRPr"):
        for el in body.iter(qn(tag)):
            sz = el.get("sz")
            if sz is not None:
                try:
                    return int(sz)
                except ValueError:
                    continue
    return None


def _first_run_style(body: etree._Element) -> tuple[bool, bool]:
    """(bold, italic) from the first rPr/defRPr that states either."""
    bold = italic = False
    for tag in ("a:rPr", "a:defRPr"):
        for el in body.iter(qn(tag)):
            if el.get("b") is not None or el.get("i") is not None:
                return xml_bool(el.get("b")), xml_bool(el.get("i"))
    return bold, italic


def _theme_font(pkg, part: str | None, token: str) -> str | None:
    """The theme's real typeface behind a +mn-lt / +mj-lt reference."""
    if pkg is None or not part:
        return None
    try:
        from .design import _theme_part_of
        from .inherit import layout_part_of, master_part_of

        master = master_part_of(pkg, layout_part_of(pkg, part))
        if not master:
            return None
        theme_part = _theme_part_of(pkg, master)
        if not pkg.has_part(theme_part):
            return None
        tag = "a:majorFont" if token.startswith("+mj") else "a:minorFont"
        latin = pkg.root(theme_part).find(
            f"{qn('a:themeElements')}/{qn('a:fontScheme')}/{qn(tag)}/"
            f"{qn('a:latin')}"
        )
    except Exception:
        return None
    if latin is None:
        return None
    face = latin.get("typeface")
    return face or None


def _latin_typeface(
    body: etree._Element, pkg=None, part: str | None = None
) -> str | None:
    """The latin typeface this body's text is set in.

    Run first, then paragraph/shape defRPr, then the theme's minor font.
    A `+mn-lt` / `+mj-lt` reference resolves to the theme's actual face,
    which is what a font file can be looked up by."""
    face = None
    for tag in ("a:rPr", "a:defRPr"):
        for el in body.iter(qn(tag)):
            latin = el.find(qn("a:latin"))
            if latin is not None and latin.get("typeface"):
                face = latin.get("typeface")
                break
        if face:
            break
    if face is None:
        face = "+mn-lt"
    if face.startswith("+"):
        return _theme_font(pkg, part, face)
    return face


#: DrawingML caps marL/indent at 51206400 EMU (56 inches). A value outside
#: that is malformed, and a malformed indent silently read as zero is a
#: measurement of a frame nobody has.
_MAR_LIMIT_EMU = 51206400


class _Unmeasurable(Exception):
    """A run the measured path cannot resolve WITH CERTAINTY.

    The measured path is allowed to be unavailable; it is not allowed to
    guess. Anything it cannot establish raises this, the whole body falls
    back to the estimate, and `reason` becomes the method_reason the caller
    reads, naming the run that forced it.
    """

    def __init__(self, reason: str) -> None:
        super().__init__(reason)
        self.reason = reason


def _para_level(p: etree._Element) -> int:
    """The paragraph's outline level, 0..8."""
    ppr = p.find(qn("a:pPr"))
    if ppr is None:
        return 0
    try:
        return min(max(int(ppr.get("lvl") or 0), 0), 8)
    except (TypeError, ValueError):
        return 0


def _para_margins(p: etree._Element, chain=()) -> tuple[int, int]:
    """(marL, indent) in EMU for one paragraph; indent is negative for a
    hanging indent, which is how a bulleted list is written.

    Resolved PER ATTRIBUTE: the paragraph's own a:pPr first, then the
    a:lvlNpPr elements behind it (`chain`, from inherit.level_ppr_chain).
    Reading only the paragraph meant a bulleted body whose marL lives in
    the master measured as if it had the whole frame. Nothing in the chain
    defining an attribute leaves the DrawingML default of zero, which is
    what PowerPoint uses too."""
    containers = []
    own = p.find(qn("a:pPr"))
    if own is not None:
        containers.append(own)
    containers.extend(chain)

    def _int(name: str, low: int) -> int:
        for container in containers:
            raw = container.get(name)
            if raw is None:
                continue
            try:
                value = int(raw)
            except (TypeError, ValueError):
                raise _Unmeasurable(
                    f"a paragraph states a {name} of {raw!r}, which is not a "
                    "number of EMU"
                ) from None
            if not low <= value <= _MAR_LIMIT_EMU:
                raise _Unmeasurable(
                    f"a paragraph states a {name} of {value} EMU, outside "
                    "the range DrawingML allows"
                )
            return value
        return 0
    # The two attributes do NOT share a range, and treating them as if they
    # did accepted a negative marL as a confident measurement: marL is
    # ST_TextMargin, which is NONNEGATIVE, while indent is ST_TextIndent,
    # which is signed because a hanging indent is how every bulleted list
    # is written. A negative marL is
    # schema-invalid whether the paragraph states it or inherits it, so it
    # is unmeasurable; a negative indent stays measurable.
    return _int("marL", 0), _int("indent", -_MAR_LIMIT_EMU)


def _body_typeface(
    body: etree._Element, pkg=None, part: str | None = None
) -> str | None:
    """The typeface a run inherits when NOTHING in its own chain names one:
    the theme's minor font, which is what DrawingML falls back to.

    It used to take the first a:defRPr anywhere in the shape's a:lstStyle,
    regardless of outline level, so a level-two run could be measured in
    level one's face while the result still said font-metrics. The list
    styles are read at the run's ACTUAL level
    now, by _rpr_chain over inherit.level_ppr_chain, the same effective
    chain the size resolver walks; this is only what is left when that
    chain names nothing. Deliberately does NOT look at other runs."""
    return _theme_font(pkg, part, "+mn-lt")


def _rpr_chain(r: etree._Element, p: etree._Element, chain=()):
    """The rPr sources for one run, nearest first: its own, the paragraph's
    defRPr, then the defRPr of each a:lvlNpPr standing behind the paragraph
    (`chain`, from inherit.level_ppr_chain)."""
    out = [r.find(qn("a:rPr"))]
    ppr = p.find(qn("a:pPr"))
    out.append(ppr.find(qn("a:defRPr")) if ppr is not None else None)
    for lvl in chain:
        out.append(lvl.find(qn("a:defRPr")))
    return [el for el in out if el is not None]


def _certain_size_pt(
    rpr, p: etree._Element, body: etree._Element, elem, level: int,
    pkg=None, part: str | None = None,
) -> float | None:
    """The run's size in points, or None when it cannot be ESTABLISHED.

    One resolver, not two: this is inherit.resolve_run_size_pt, the same
    chain check_layout's "text too small" check acts on, which walks the
    run's sz, the paragraph's defRPr, the shape's own lstStyle at this
    level, the layout twin, the master twin, the master text style and the
    presentation default. It reaches that last step only for a
    placeholder, so a plain text box asks for the presentation default
    separately rather than falling back to the body's first run."""
    if elem is None and body is not None:
        elem = body.getparent()
    if elem is None:
        return _inh._sz_pt(rpr.get("sz")) if rpr is not None else None
    size, _src = _inh.resolve_run_size_pt(
        pkg, part if pkg is not None else None, elem, p, rpr, level
    )
    if size is None:
        size = _inh.presentation_default_size_pt(pkg, level)
    return size


def _resolve_run(
    r: etree._Element, p: etree._Element, body: etree._Element,
    inherited_face: str | None, default_pt: float, pkg=None,
    part: str | None = None, elem=None, level: int = 0, chain=(),
):
    """(family, size_pt, bold, italic, spc_pt) for ONE run, each property
    resolved from that run's own chain rather than from the first run in
    the body. Raises _Unmeasurable when the size cannot be established: a
    28pt run inherited from a level-one list style used to be measured at
    the body default, which is large enough to reverse a verdict."""
    rpr = r.find(qn("a:rPr"))
    rpr_chain = _rpr_chain(r, p, chain)

    size = _certain_size_pt(rpr, p, body, elem, level, pkg, part)
    if size is None:
        raise _Unmeasurable(
            "the run's font size is inherited and could not be resolved"
        )

    spc_pt = 0.0
    for el in rpr_chain:
        raw = el.get("spc")
        if raw:
            try:
                spc_pt = int(raw) / 100.0
            except (TypeError, ValueError):
                raise _Unmeasurable(
                    "the run's character spacing (a:rPr/@spc) could not be "
                    "read"
                ) from None
            break

    def flag(name: str) -> bool:
        for el in rpr_chain:
            if el.get(name) is not None:
                return xml_bool(el.get(name))
        return False

    face = None
    for el in rpr_chain:
        latin = el.find(qn("a:latin"))
        if latin is not None and latin.get("typeface"):
            face = latin.get("typeface")
            break
    if face is None:
        face = inherited_face
    elif face.startswith("+"):
        face = _theme_font(pkg, part, face)
    return face, size, flag("b"), flag("i"), spc_pt


#: Run-level elements that carry text. a:fld (a slide number or date) holds
#: cached text that renders like any other run and takes up the same space.
_TEXT_RUN_TAGS = ("a:r", "a:fld")


def _paragraph_runs(
    p: etree._Element, body: etree._Element, inherited_face: str | None,
    default_pt: float, pkg=None, part: str | None = None, elem=None,
    chain=None,
):
    """The paragraph as a list of fontmetrics.StyledRun. a:br becomes a
    hard newline carried on the preceding run's style.

    Raises _Unmeasurable for a run whose typeface or size cannot be
    established, naming that run rather than the first one in the body."""
    from . import fontmetrics as _fm

    level = _para_level(p)
    if chain is None:
        chain = _inh.level_ppr_chain(
            pkg, part, elem if elem is not None else body.getparent(), level
        )
    runs: list = []
    for child in p:
        tag = etree.QName(child).localname
        if tag == "br":
            if runs:
                last = runs[-1]
                runs[-1] = last._replace(text=last.text + "\n")
            else:
                runs.append(_fm.StyledRun("\n", inherited_face or "",
                                          default_pt))
            continue
        if f"a:{tag}" not in _TEXT_RUN_TAGS:
            continue
        t = child.find(qn("a:t"))
        text = t.text or "" if t is not None else ""
        if not text:
            continue
        family, size, bold, italic, spc_pt = _resolve_run(
            child, p, body, inherited_face, default_pt, pkg, part,
            elem, level, chain,
        )
        if not family:
            raise _Unmeasurable(
                "a run's typeface does not resolve through the theme"
            )
        runs.append(
            _fm.StyledRun(text, family, size, bold, italic, spc_pt)
        )
    return runs


def _measured_overflow(
    body: etree._Element,
    bodypr: etree._Element | None,
    inner_w: int,
    inner_h: int,
    pt: float,
    line_h: float,
    pkg=None,
    part: str | None = None,
    lnspc_reduction_pct: float = 0.0,
    font_scale_pct: float = 100.0,
    elem=None,
) -> dict | None:
    """Lines and widest line measured PER RUN against each run's own font.

    Every run is resolved separately (its own face, size, bold, italic and
    character spacing), the wrap runs across run boundaries, and a line's
    height comes from the largest point size ON THAT LINE. Returns None
    when the metrics extra is not installed, and raises _Unmeasurable when
    a run cannot be resolved with certainty; either way the caller falls
    back to the estimate and says why. A body measured in one run's format,
    or at a size nobody could establish, must never be labelled
    font-metrics."""
    from . import fontmetrics as _fm

    if not _fm.available():
        return None
    if elem is None:
        elem = body.getparent()
    inherited = _body_typeface(body, pkg, part)
    wrap = (bodypr.get("wrap") if bodypr is not None else None) or "square"
    scale = font_scale_pct / 100.0
    height_mult = 1.2 * (1.0 - lnspc_reduction_pct / 100.0) * EMU_PER_POINT

    total_h = 0.0
    lines = 0
    widest_pt = 0.0
    faces: dict[str, str] = {}
    runs_measured = 0
    for p in body.findall(qn("a:p")):
        level = _para_level(p)
        chain = _inh.level_ppr_chain(pkg, part, elem, level)
        runs = _paragraph_runs(p, body, inherited, pt / scale if scale else pt,
                               pkg, part, elem, chain)
        runs = [r._replace(size_pt=r.size_pt * scale) for r in runs]
        runs_measured += len(runs)
        for run in runs:
            # Readability first: metrics_readable is False only when a file
            # WAS found and would not parse, which is a different errand for
            # the reader than a font that is not installed.
            if _fm.metrics_readable(
                run.family, bold=run.bold, italic=run.italic
            ) is False:
                raise _Unmeasurable(
                    f"a font file for {run.family} was found but its metrics "
                    "could not be read"
                )
            name = _fm.font_file_name(
                run.family, bold=run.bold, italic=run.italic
            )
            if name is None:
                raise _Unmeasurable(
                    f"no font file for {run.family} was found on this machine"
                )
            faces[run.family] = name
        if not runs:
            lines += 1
            total_h += pt * height_mult
            continue
        mar_l, indent = _para_margins(p, chain)
        first_w = (inner_w - mar_l - indent) / EMU_PER_POINT
        rest_w = (inner_w - mar_l) / EMU_PER_POINT
        if wrap == "none":
            # No wrapping: every hard break is a line and the width is what
            # decides whether it fits.
            laid = _fm.wrap_styled(runs, 10**9, 10**9)
        else:
            laid = _fm.wrap_styled(runs, first_w, rest_w)
        if laid is None:
            raise _Unmeasurable(
                "a paragraph could not be laid out against its fonts"
            )
        for line in laid:
            lines += 1
            total_h += line.max_size_pt * height_mult
            # The line's width as PowerPoint renders it: the advance sum
            # plus the measured trailing allowance, then the indent it
            # starts at. Judging a non-wrapping frame on the raw sum is
            # what let a label 3 pt too wide report as fitting.
            widest_pt = max(
                widest_pt,
                _fm.calibrated_line_pt(line.width_pt)
                + (mar_l + indent) / EMU_PER_POINT,
            )

    height_ratio = total_h / inner_h if inner_h > 0 else 0.0
    width_ratio = widest_pt / (inner_w / EMU_PER_POINT) if inner_w > 0 else 0.0
    # A wrapping frame cannot overflow sideways: the wrap already accounted
    # for the width. A frame with wrap="none" can, and that is the case the
    # old model could not see at all.
    overflow = height_ratio > 1.0 or (wrap == "none" and width_ratio > 1.0)
    ratio = max(height_ratio, width_ratio) if wrap == "none" else height_ratio
    used = sorted(set(faces.values()))
    return {
        "heuristic": False,
        "method": "font-metrics",
        "font_file": used[0] if len(used) == 1 else ", ".join(used),
        "typeface": (
            sorted(faces)[0] if len(faces) == 1 else ", ".join(sorted(faces))
        ),
        "runs_measured": runs_measured,
        "likely_overflow": overflow,
        "fill_ratio": round(ratio, 2),
        "height_ratio": round(height_ratio, 2),
        "width_ratio": round(width_ratio, 2),
        "estimated_lines": lines,
        "wrap": wrap,
        "note": (
            "every run measured against its own installed font; a "
            f"{_fm.WIDTH_SAFETY_PAD_PT:g} pt allowance, calibrated on "
            "tested PowerPoint layouts, is added to each line, so a line "
            "that only just fits may be reported as wrapping; PowerPoint's "
            "rendering is still the final authority"
        ),
    }


def _overflow_heuristic(
    elem: etree._Element,
    body: etree._Element,
    bodypr: etree._Element | None,
    font_scale_pct: float,
    lnspc_reduction_pct: float,
    box: tuple[float, float, float, float] | None = None,
    size_pt: float | None = None,
    pkg=None,
    part: str | None = None,
    frame: tuple[int, int] | None = None,
) -> dict | None:
    """Fit estimate against the frame's inner box.

    Two models, and every result says which one produced it. When the
    optional metrics extra is installed AND the run's font file can be
    found on this machine, the text is measured against that font's own
    advance widths and wrapped by word ("method": "font-metrics"). Failing
    either condition it falls back to the original average-glyph-width
    estimate ("method": "estimate", with the reason). The honest fit
    authority is still PowerPoint's own renderer.

    `box` supplies (x, y, cx, cy) for a shape with no a:xfrm of its own,
    and `size_pt` a font size resolved through the layout/master chain.
    Callers that can resolve those pass them; without them a placeholder
    inheriting its geometry cannot be estimated at all and an unsized run
    falls back to a flat guess. `pkg`/`part` let the typeface resolve
    through the theme; without them the measured path is unavailable.
    `frame` (cx, cy) in EMU replaces the box outright: fit_text measures a
    shape that grows with its text against the room it can grow into."""
    xfrm = elem.find(f"{qn('p:spPr')}/{qn('a:xfrm')}")
    ext = xfrm.find(qn("a:ext")) if xfrm is not None else None
    if frame is not None:
        cx, cy = int(frame[0]), int(frame[1])
    elif ext is None and box is not None:
        cx, cy = int(box[2]), int(box[3])
    elif ext is None:
        return {
            "heuristic": True,
            "method": "estimate",
            "method_reason": "no geometry to measure against",
            "likely_overflow": None,
            "note": (
                "shape has no explicit geometry (inherited from the "
                "layout placeholder); fit cannot be estimated here"
            ),
        }
    else:
        cx, cy = int(ext.get("cx")), int(ext.get("cy"))
    inner_w = cx - _ins(bodypr, "lIns", 91440) - _ins(bodypr, "rIns", 91440)
    inner_h = cy - _ins(bodypr, "tIns", 45720) - _ins(bodypr, "bIns", 45720)
    if inner_w <= 0 or inner_h <= 0:
        return {
            "heuristic": True,
            "method": "estimate",
            "method_reason": "frame insets consume the whole shape",
            "likely_overflow": True,
            "note": "frame insets consume the whole shape",
        }
    raw = _first_run_size(body)
    if raw is not None:
        base_pt = raw / 100.0
    elif size_pt is not None:
        base_pt = size_pt
    else:
        base_pt = 18.0  # last-resort guess when nothing resolves
    pt = base_pt * (font_scale_pct / 100.0)
    char_w = 0.5 * pt * EMU_PER_POINT  # average glyph width model
    line_h = 1.2 * pt * EMU_PER_POINT * (1.0 - lnspc_reduction_pct / 100.0)

    measured = None
    unmeasurable = None
    try:
        measured = _measured_overflow(
            body, bodypr, inner_w, inner_h, pt, line_h, pkg, part,
            lnspc_reduction_pct=lnspc_reduction_pct,
            font_scale_pct=font_scale_pct,
            elem=elem,
        )
    except _Unmeasurable as exc:
        # The measurement refused to guess, and it knows which run and why.
        unmeasurable = exc.reason
    except Exception:
        measured = None  # measurement never becomes the thing that raises
    if measured is not None:
        return measured

    lines = 0
    for p in body.findall(qn("a:p")):
        plain = paragraph_text(p)
        for chunk in plain.split("\v"):  # a:br forces a visual line
            lines += max(1, math.ceil(len(chunk) * char_w / inner_w))
    est_h = lines * line_h
    ratio = est_h / inner_h
    return {
        "heuristic": True,
        "method": "estimate",
        "method_reason": unmeasurable or _estimate_reason(body, pkg, part),
        "likely_overflow": ratio > 1.0,
        "fill_ratio": round(ratio, 2),
        "estimated_lines": lines,
        "note": (
            "text-length-vs-frame-area estimate at the current font size, "
            "no real font metrics; trust PowerPoint's rendering, not this"
        ),
    }


def _estimate_reason(body: etree._Element, pkg=None, part: str | None = None) -> str:
    """Why this result is an estimate rather than a measurement.

    Four closed reasons, and they are not interchangeable: an extra that is
    not installed, a typeface that does not resolve, a typeface with no file
    on this machine, and a file that exists but whose metrics will not parse
    (the last one used to report the third, which sends someone to install a
    font they already have)."""
    from . import fontmetrics as _fm

    if not _fm.available():
        return _fm.unavailable_reason()
    try:
        face = _latin_typeface(body, pkg, part)
    except Exception:
        face = None
    if not face:
        return "the run's typeface does not resolve through the theme"
    try:
        readable = _fm.metrics_readable(face)
    except Exception:
        readable = None
    if readable is False:
        return (
            f"a font file for {face} was found but its metrics could not "
            "be read"
        )
    return f"no font file for {face} was found on this machine"


def _autofit_record(elem: etree._Element, pkg=None, part: str | None = None) -> dict:
    body = elem.find(qn("p:txBody"))
    bodypr = body.find(qn("a:bodyPr")) if body is not None else None
    mode = "inherited"
    font_scale = 100.0
    lnspc_reduction = 0.0
    if bodypr is not None:
        norm = bodypr.find(qn("a:normAutofit"))
        if norm is not None:
            mode = "normAutofit"
            font_scale = _pct_value(norm.get("fontScale"), 100.0)
            lnspc_reduction = _pct_value(norm.get("lnSpcReduction"), 0.0)
        elif bodypr.find(qn("a:spAutoFit")) is not None:
            mode = "spAutoFit"
        elif bodypr.find(qn("a:noAutofit")) is not None:
            mode = "none"
    rec = {
        "shape_id": _shape_id(elem),
        "name": _shape_name(elem),
        "autofit": mode,
    }
    if mode == "normAutofit":
        rec["font_scale_pct"] = font_scale
        rec["line_spacing_reduction_pct"] = lnspc_reduction
    if body is not None:
        rec["overflow"] = _overflow_heuristic(
            elem, body, bodypr, font_scale, lnspc_reduction,
            pkg=pkg, part=part,
        )
    return rec


def _autofit_slide_records(pkg: PptxPackage, rec: dict) -> list[dict]:
    """Autofit records for every text-bearing shape on one slide record."""
    sp_tree = _sp_tree(pkg, rec["part"])
    shapes: list[dict] = []
    if sp_tree is not None:
        for elem, _kind, _z, _parent in iter_shapes(sp_tree):
            if (
                etree.QName(elem).localname == "sp"
                and elem.find(qn("p:txBody")) is not None
            ):
                shapes.append(_autofit_record(elem, pkg, rec["part"]))
    return shapes


def get_autofit_state(pkg: PptxPackage, slide=None, shape=None, *,
                      limit=None, offset: int = 0) -> dict:
    """Autofit state of one shape, every text-bearing shape on a slide, or
    (slide=None) every text-bearing shape on EVERY slide: the bodyPr autofit
    mode (normAutofit with fontScale/lnSpcReduction, spAutoFit, none, or
    inherited when the bodyPr carries no autofit child), plus a rough
    overflow heuristic. The caveat is structural, not a bug: cached
    normAutofit values reflect the last PowerPoint edit and are only
    recomputed when the frame is edited again inside PowerPoint."""
    if slide is None:
        if shape is not None:
            raise PptMcpError(
                "shape scoping needs an explicit slide; pass slide together "
                "with shape (slide=None means all slides)"
            )
        slides_out = [
            {
                "slide_index": rec["index"],
                "slide_number": rec["index"] + 1,
                "slide_id": rec["slide_id"],
                "shapes": _autofit_slide_records(pkg, rec),
            }
            for rec in slides_in_scope(pkg, None)
        ]
        # Deck-wide, this is one record per text-bearing shape on every
        # slide: 880,000 characters on the heaviest corpus deck.
        kept, page = _budget.page_items(
            slides_out, limit=limit, offset=offset,
            overhead=_budget.overhead_of({"caveat": _AUTOFIT_CAVEAT}),
            unit="slides",
            narrow_hint="narrow with slide=<index>",
        )
        out = {"caveat": _AUTOFIT_CAVEAT, "slides": kept}
        if page is not None:
            out["page"] = page
        return out
    rec = resolve_slide(pkg, slide)
    if shape is not None:
        elem, kind = _resolve_shape(pkg, rec, shape)
        _require_txbody(elem, kind, rec)
        shapes = [_autofit_record(elem, pkg, rec["part"])]
    else:
        shapes = _autofit_slide_records(pkg, rec)
    header = {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "caveat": _AUTOFIT_CAVEAT,
    }
    kept, page = _budget.page_items(
        shapes, limit=limit, offset=offset,
        overhead=_budget.overhead_of(header),
        unit="shapes",
        narrow_hint="narrow with shape=<id>",
        shrink=_budget.shrink_record,
    )
    return {
        **header,
        "shapes": kept,
        **({"page": page} if page is not None else {}),
    }


# ------------------------------------------------------------------ fit_text

#: The a:bodyPr autofit choice by element name, in get_autofit_state's words.
_AUTOFIT_MODE_OF = {
    "normAutofit": "normAutofit",
    "spAutoFit": "spAutoFit",
    "noAutofit": "none",
}

_FIT_NOT_OVERFLOWING = (
    "the heuristic does not estimate this shape as overflowing (or its "
    "geometry is inherited); nothing was changed"
)
_FIT_GROWS_ON_SLIDE = (
    "the shape is set to resize to fit its text, and at this size it grows "
    "and stays on the slide, so there is nothing to fit; nothing was changed"
)
_FIT_OFF_SLIDE = (
    "the shape is set to resize to fit its text, but it starts at or past "
    "the slide's edge, so there is no room on the slide to fit its text "
    "into; nothing was changed. Move it onto the slide with set_shape, then "
    "fit again"
)
_FIT_SIZE_UNKNOWN = (
    "a run's size is inherited and could not be established, so it cannot "
    "be scaled without guessing; nothing was changed. Set the sizes with "
    "format_text, then fit again"
)
_FIT_NOTE_AUTOFIT = (
    "Each shape keeps its own autofit setting: text set to shrink on "
    "overflow is fitted through that setting's font scale, other text by "
    "rewriting its sizes, and a shape set to resize to fit its text is "
    "fitted to the room it can grow into on the slide."
)


def _explicit_size_elements(body: etree._Element) -> list[etree._Element]:
    """Every element in the body carrying an explicit sz (rPr, defRPr,
    endParaRPr): the set a uniform scale rewrites."""
    out = []
    for tag in ("a:rPr", "a:defRPr", "a:endParaRPr"):
        for el in body.iter(qn(tag)):
            if el.get("sz") is not None:
                out.append(el)
    return out


def _autofit_choice(bodypr: etree._Element | None) -> str | None:
    """The autofit mode a bodyPr states itself, or None when it states
    none."""
    if bodypr is None:
        return None
    for child in bodypr:
        if not isinstance(child.tag, str):
            continue
        mode = _AUTOFIT_MODE_OF.get(etree.QName(child).localname)
        if mode is not None:
            return mode
    return None


def _effective_autofit(elem: etree._Element, pkg=None,
                       part: str | None = None) -> tuple[str, str]:
    """(mode, where it is set) for one shape: its own bodyPr, else, for a
    placeholder, its layout twin's, else its master twin's. Nothing set
    anywhere is no autofit ("none", "default"), which is how PowerPoint
    draws it."""
    body = elem.find(qn("p:txBody"))
    mode = _autofit_choice(
        body.find(qn("a:bodyPr")) if body is not None else None)
    if mode is not None:
        return mode, "shape"
    key = _inh.placeholder_key(_inh.placeholder_of(elem))
    if key is not None and pkg is not None and part:
        layout = _inh.layout_part_of(pkg, part)
        master = _inh.master_part_of(pkg, layout)
        for where, twin in (
            ("layout", _inh.layout_twin(pkg, layout, key)),
            ("master", _inh.master_twin(pkg, master, key)),
        ):
            if twin is None:
                continue
            tbody = twin.find(qn("p:txBody"))
            mode = _autofit_choice(
                tbody.find(qn("a:bodyPr")) if tbody is not None else None)
            if mode is not None:
                return mode, where
    return "none", "default"


def _set_normautofit(
    bodypr: etree._Element, font_scale_pct: float | None
) -> None:
    """Make normAutofit the body's autofit mode, replacing any existing
    autofit choice. font_scale_pct None or 100 writes a bare normAutofit."""
    for tag in ("a:noAutofit", "a:normAutofit", "a:spAutoFit"):
        for el in bodypr.findall(qn(tag)):
            bodypr.remove(el)
    norm = etree.Element(qn("a:normAutofit"))
    if font_scale_pct is not None and font_scale_pct < 100:
        norm.set("fontScale", str(int(round(font_scale_pct * 1000))))
    # The autofit choice sits after prstTxWarp, before scene3d/sp3d/extLst.
    warp = bodypr.find(qn("a:prstTxWarp"))
    if warp is not None:
        warp.addnext(norm)
    else:
        bodypr.insert(0, norm)


def _slide_size(pkg) -> tuple[int, int]:
    """(cx, cy) of the slide in EMU; the DrawingML default when unstated."""
    sld = pkg.presentation().find(qn("p:sldSz")) if pkg is not None else None
    if sld is None:
        return 9144000, 6858000
    return int(sld.get("cx")), int(sld.get("cy"))


def _frame_box(elem: etree._Element, pkg=None, part: str | None = None):
    """(x, y, cx, cy, own, grouped) in slide EMU: the box the text is laid
    out in. A shape's own box is mapped through any groups it sits in; a
    placeholder with no box of its own is drawn in its layout's (or
    master's), and own is False. None when there is no box anywhere."""
    from . import shapes as _shapes

    xfrm = elem.find(f"{qn('p:spPr')}/{qn('a:xfrm')}")
    if (xfrm is not None and xfrm.find(qn("a:off")) is not None
            and xfrm.find(qn("a:ext")) is not None):
        chain = _shapes._chain_of(elem)
        try:
            x, y, cx, cy = _shapes._slide_box(elem, chain)
        except Exception:
            return None
        return int(x), int(y), int(cx), int(cy), True, bool(chain)
    if pkg is not None and part:
        box = _inh.inherited_box(pkg, part, elem)
        if box is not None:
            return (*box, False, False)
    return None


def _inherited_run_sizes(elem, body, pkg=None, part=None):
    """(run, its rPr or None, size in points) for every text run whose rPr
    states no size, resolved through the placeholder chain BEFORE anything
    changes. None when one of them cannot be established."""
    out = []
    for p in body.findall(qn("a:p")):
        level = _para_level(p)
        for child in p:
            if child.tag not in (qn("a:r"), qn("a:fld")):
                continue
            rpr = child.find(qn("a:rPr"))
            if rpr is not None and rpr.get("sz") is not None:
                continue
            size = _certain_size_pt(rpr, p, body, elem, level, pkg, part)
            if size is None:
                return None
            out.append((child, rpr, size))
    return out


def _run_sizes_pt(body: etree._Element, inherited) -> list[float]:
    """The size every text run is drawn at: its own sz, or the inherited
    size resolved for it."""
    sizes = []
    for p in body.findall(qn("a:p")):
        for child in p:
            if child.tag not in (qn("a:r"), qn("a:fld")):
                continue
            rpr = child.find(qn("a:rPr"))
            if rpr is not None and rpr.get("sz") is not None:
                try:
                    sizes.append(int(rpr.get("sz")) / 100.0)
                except ValueError:
                    continue
    sizes.extend(size for _run, _rpr, size in inherited or ())
    return sizes


def _floor_pct(sizes: list[float], min_size: float,
               fallback_pt: float) -> float:
    """The smallest uniform scale that keeps every run at min_size or above,
    measured against the smallest run that is not already at or below it
    (small print the caller chose does not hold the fit back)."""
    eligible = [size for size in sizes if size > min_size]
    if eligible:
        base = min(eligible)
    elif sizes:
        return 100.0
    else:
        base = fallback_pt
    return max(1.0, min(100.0, min_size / base * 100.0))


def _floored_pt(size_pt: float, scale_pct: float, min_size: float) -> float:
    """A run's size after a uniform scale: never below min_size, and never
    larger than it was (a run already at or below min_size keeps its
    size; min_size is a floor for shrinking, not a size to grow to)."""
    if size_pt <= min_size:
        return size_pt
    return max(min_size, size_pt * scale_pct / 100.0)


def _largest_fitting_scale(fits, floor_pct: float) -> tuple[float, bool]:
    """(the largest uniform scale in [floor, 100] that fits, found to about
    1% and rounded down to 0.1, and whether even the floor fits)."""
    if not fits(floor_pct):
        return floor_pct, False
    lo, hi = floor_pct, 100.0
    for _ in range(12):
        mid = (lo + hi) / 2.0
        if fits(mid):
            lo = mid
        else:
            hi = mid
        if hi - lo < 1.0:
            break
    return max(floor_pct, math.floor(lo * 10) / 10), True


def _fit_one_shape(elem: etree._Element, min_size: float, pkg=None,
                   part: str | None = None) -> tuple[dict | None, str | None]:
    """Fit one shape's text with ONE mechanism, keeping its autofit setting.

    Returns (report, None) when it changed the shape, or (None, reason)
    when it left it alone; a None reason means the estimate sees no
    overflow it can measure.

    - Shrink text on overflow (normAutofit, the shape's own or inherited):
      the fit is written as that setting's fontScale, on the shape, and no
      run size changes. Rewriting the sizes as well left the shrink on, so
      PowerPoint shrank the text a second time when it drew the slide.
    - Resize shape to fit text (spAutoFit): the box grows with its text, so
      the text is measured against the room from the box's top to the
      slide's bottom edge (and to the right edge when it does not wrap).
      Text that fits there needs nothing; text that does not has its run
      sizes scaled, and the box's stored height is set to the estimate.
    - No autofit: the run sizes are rewritten.
    Inherited sizes are written out at the size they are drawn, then scaled.
    """
    body = elem.find(qn("p:txBody"))
    if body is None:
        return None, None
    bodypr = body.find(qn("a:bodyPr"))
    mode, where = _effective_autofit(elem, pkg, part)
    box = _frame_box(elem, pkg, part)
    if box is None:
        return None, None
    x, y, cx, cy, own_box, grouped = box
    frame = (cx, cy)
    fitted_to = "frame"
    vertical = _ins(bodypr, "tIns", 45720) + _ins(bodypr, "bIns", 45720)
    if mode == "spAutoFit":
        slide_w, slide_h = _slide_size(pkg)
        wrap = (bodypr.get("wrap") if bodypr is not None else None) or "square"
        room_h = slide_h - y
        room_w = slide_w - x if wrap == "none" else cx
        if room_h <= vertical or room_w <= 0:
            return None, _FIT_OFF_SLIDE
        frame = (room_w, room_h)
        fitted_to = "slide"
    inherited = _inherited_run_sizes(elem, body, pkg, part)
    first = _first_run_size(body)
    if first is not None:
        size_hint = first / 100.0
    elif inherited:
        size_hint = inherited[0][2]
    else:
        size_hint = None

    def estimate(scale_pct: float) -> dict | None:
        return _overflow_heuristic(
            elem, body, bodypr, scale_pct, 0.0, size_pt=size_hint,
            pkg=pkg, part=part, frame=frame,
        )

    est0 = estimate(100.0)
    if not est0 or est0.get("likely_overflow") is not True:
        if mode == "spAutoFit" and est0 and est0.get("likely_overflow") is False:
            return None, _FIT_GROWS_ON_SLIDE
        return None, None

    def fits(scale_pct: float) -> bool:
        est = estimate(scale_pct)
        if not est or est.get("likely_overflow") is None:
            return True
        return not est["likely_overflow"]

    if mode != "normAutofit" and inherited is None:
        return None, _FIT_SIZE_UNKNOWN
    sizes = _run_sizes_pt(body, inherited)
    floor_pct = _floor_pct(sizes, min_size, size_hint or 18.0)
    scale_pct, fits_floor = _largest_fitting_scale(fits, floor_pct)
    before: dict = {
        "sizes_pt": sorted(set(sizes)) or "inherited",
        "fill_ratio": est0.get("fill_ratio"),
    }
    after: dict = {}
    if mode == "normAutofit":
        # PowerPoint's own shrink stays the one mechanism: the fit becomes
        # its fontScale, replacing a cache drawn for older text, and the
        # run sizes stay as they are.
        own = bodypr.find(qn("a:normAutofit")) if bodypr is not None else None
        before["font_scale_pct"] = (
            _pct_value(own.get("fontScale"), 100.0) if own is not None
            else 100.0)
        if bodypr is None:
            bodypr = etree.Element(qn("a:bodyPr"))
            body.insert(0, bodypr)
        _set_normautofit(bodypr, scale_pct)
        applied = "normAutofit_scale"
        est_after = estimate(scale_pct)
        after["sizes_pt"] = before["sizes_pt"]
        after["font_scale_pct"] = scale_pct if scale_pct < 100 else None
    else:
        for el in _explicit_size_elements(body):
            old = int(el.get("sz")) / 100.0
            new = _floored_pt(old, scale_pct, min_size)
            el.set("sz", str(int(round(new * 100))))
        for run, rpr, size in inherited:
            if rpr is None:
                rpr = etree.Element(qn("a:rPr"))
                run.insert(0, rpr)
            new = _floored_pt(size, scale_pct, min_size)
            rpr.set("sz", str(int(round(new * 100))))
        applied = "run_sizes"
        est_after = estimate(100.0)
        after["sizes_pt"] = sorted(set(_run_sizes_pt(body, ()))) or "inherited"
        after["font_scale_pct"] = None
        if mode == "spAutoFit" and own_box and not grouped and est_after:
            ratio = est_after.get("height_ratio", est_after.get("fill_ratio"))
            ext = elem.find(f"{qn('p:spPr')}/{qn('a:xfrm')}/{qn('a:ext')}")
            if ratio is not None and ext is not None:
                needed = int(math.ceil(ratio * (frame[1] - vertical))) + vertical
                new_cy = max(1, min(frame[1], needed))
                old_cy = int(ext.get("cy"))
                if new_cy != old_cy:
                    ext.set("cy", str(new_cy))
                    before["box_h_in"] = _emu_to_in(old_cy)
                    after["box_h_in"] = _emu_to_in(new_cy)
    after["fill_ratio"] = (est_after or {}).get("fill_ratio")
    report = {
        "shape_id": _shape_id(elem),
        "name": _shape_name(elem),
        "applied": applied,
        "autofit": mode,
        "autofit_from": where,
        "autofit_changed": False,
        "fitted_to": fitted_to,
        "method": est0.get("method"),
        "method_reason": est0.get("method_reason"),
        "font_file": est0.get("font_file"),
        "scale_pct": round(scale_pct, 1),
        "before": before,
        "after": after,
        "still_overflowing": (not fits_floor) or bool(
            est_after and est_after.get("likely_overflow") is True),
    }
    return report, None


def fit_text(
    pkg: PptxPackage, slide, shape=None, *, min_size: float = 10.0
) -> dict:
    """Approximate text-fit: estimate the largest uniform font scale
    (floored at min_size pt) at which a shape's text fits, using the same
    overflow estimate as get_autofit_state (measured against the installed
    font when the metrics extra can, an average-glyph-width HEURISTIC
    otherwise), then apply it through ONE mechanism, keeping the shape's
    autofit setting (its own, or inherited from the layout or master):
    shrink text on overflow (normAutofit) gets the scale as its fontScale
    and its run sizes stay; resize shape to fit text (spAutoFit) is fitted
    to the room it can grow into on the slide, by its run sizes; no autofit
    gets its run sizes rewritten. Inherited sizes are written out first. A
    run already at or below min_size keeps its size. shape=None treats
    every text shape on the slide that the estimate flags as overflowing; a
    shape that still cannot fit at min_size is applied at the floor and
    reported with still_overflowing=true. This is an ESTIMATE: verify
    visually with export_slide_image before presenting."""
    if not isinstance(min_size, (int, float)) or isinstance(min_size, bool):
        raise PptMcpError(f"min_size must be a number in points, got {min_size!r}")
    if not 1 <= float(min_size) <= 400:
        raise PptMcpError(f"min_size must be 1..400 pt, got {min_size!r}")
    min_size = float(min_size)
    rec = resolve_slide(pkg, slide)
    fitted: list[dict] = []
    skipped: list[dict] = []
    if shape is not None:
        elem, kind = _resolve_shape(pkg, rec, shape)
        _require_txbody(elem, kind, rec)
        if not shape_text(elem).strip():
            raise PptMcpError(
                f"shape {_shape_id(elem)} on slide {rec['index']} has no "
                "text to fit"
            )
        report, reason = _fit_one_shape(elem, min_size, pkg, rec["part"])
        if report is None:
            skipped.append({
                "shape_id": _shape_id(elem),
                "reason": reason or _FIT_NOT_OVERFLOWING,
            })
        else:
            fitted.append(report)
    else:
        sp_tree = _sp_tree(pkg, rec["part"])
        if sp_tree is not None:
            for elem, _kind, _z, _parent in iter_shapes(sp_tree):
                if etree.QName(elem).localname != "sp":
                    continue
                if elem.find(qn("p:txBody")) is None:
                    continue
                if not shape_text(elem).strip():
                    continue
                report, _reason = _fit_one_shape(
                    elem, min_size, pkg, rec["part"])
                if report is not None:
                    fitted.append(report)
    if fitted:
        pkg.mark_dirty(rec["part"])
    methods = {r.get("method") for r in fitted if r.get("method")}
    measured = methods == {"font-metrics"}
    if measured:
        how = (
            "measured against the installed font's own advance widths "
            "(no kerning); the fit is still an estimate, so "
            "render-to-verify with export_slide_image before presenting. "
        )
    else:
        how = (
            "average-glyph-width heuristic, no real font metrics; the "
            "fit is an estimate, so render-to-verify with "
            "export_slide_image before presenting. "
        )
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "fitted": fitted,
        "skipped": skipped,
        "min_size_pt": min_size,
        "estimate": True,
        "method": "font-metrics" if measured else "estimate",
        "note": how + _FIT_NOTE_AUTOFIT,
    }
