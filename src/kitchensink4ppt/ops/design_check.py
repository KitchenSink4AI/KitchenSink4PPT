"""Layout guardrails: check_layout, the read-only design reviewer.

The gap this closes: every ecosystem server can MAKE a slide; none can tell
the agent the slide it made is bad. check_layout runs a battery of named
checks over one slide or the whole deck and returns findings an agent can
act on directly: shape ids, severities, and a fix hint naming the exact
tool call that repairs the problem.

Contract (all ops modules): functions take the open PptxPackage first and
return plain dicts. Everything here is READ-ONLY: no mark_dirty, no disk.

Honesty rules (binding):
- Every check is a static-XML heuristic, not a renderer. Each check that
  approximates says so in its caveat, which ships in the result. The fit
  authority remains PowerPoint's renderer (export_slide_images + look).
- A check that cannot evaluate a shape (image fill, a backdrop no static
  read can resolve) SKIPS it and counts the skip in the check's caveat
  data rather than guessing. Inherited geometry and inherited font size
  are NOT in that class any more: they are resolved through the layout
  and master chain (ops/inherit.py), because a placeholder that takes its
  box and its size from the layout is the normal case on a templated
  deck, and skipping it made the checks blind exactly where the risk is
  highest.

The checks and their heuristics:

- overlap: pairwise slide-space bbox intersection, at the top level AND
  between siblings inside each group. A top-level pair is flagged when the
  intersection exceeds min_overlap_pct (default 40%) of the SMALLER
  shape's area; an in-group pair when it exceeds min_group_overlap_in
  (default 0.05") in BOTH dimensions, because a diagram's real collisions
  are corner clips a percentage-of-area rule never reaches. Deliberate
  containment is not overlap: when one box fully contains the other AND
  sits behind it in z-order it is a background/panel, and the pair is
  skipped. In-group members whose name matches exclude_names (connectors,
  ticks, spines, bands and the rest of the families meant to touch) are
  left out. Connectors and hidden shapes never participate.
- off_slide: bbox vs p:sldSz. Fully outside = error (invisible content);
  partially outside = warning once the overhang exceeds
  partial_tolerance_in (default 0.1", so deliberate full-bleed edges do
  not fire).
- tiny_text: run sizes below a floor. Shapes whose whole text is one short
  line (<= label_max_chars, default 30) count as labels (floor
  label_min_pt, default 10); everything else is body (floor body_min_pt,
  default 14). Table cells use the label floor and are judged on explicit
  sizes only, since a cell inherits from the table style part. Everywhere
  else a run with no explicit sz is RESOLVED through the chain (paragraph
  defRPr, shape lstStyle, layout placeholder, master placeholder, master
  txStyles, presentation default) and the finding names the source;
  strict=True also reports runs that resolve nowhere.
- overflow: ops/text.py's Phase 3 estimate (average glyph width vs the
  frame's inner box) honoring cached normAutofit scales. Labeled
  heuristic: no real font metrics, so treat a hit as "render and look",
  not proof.
- empty_placeholder: text-family placeholders (title, ctrTitle, subTitle,
  body, obj) with an empty text body on a slide. Furniture (sldNum, dt,
  ftr) and object placeholders that hold non-text content never fire.
- missing_title: no title/ctrTitle placeholder carrying text. Severity
  info, not warning: section breaks and full-bleed visuals legitimately
  have no title, but screen readers and Outline view want one. Only a
  title placeholder with text counts, or a shape the CALLER declares a
  title through title_shape_names. Text near the top of the slide does
  not: it used to pass as a de-facto title, which hid exactly the defect
  this check exists for (a deleted title placeholder whose first bullet
  sits near the top; punch-list #928). Such text is reported as
  candidate_shape_ids instead.
- contrast: effective text color vs the color the text is READ AGAINST,
  WCAG-ish ratio. schemeClr resolves through the slide's clrMap override
  chain and the master's theme; lumMod/lumOff/tint/shade are approximated
  in HLS space. Run color resolves through the whole chain (rPr,
  paragraph defRPr, shape lstStyle, p:style fontRef, layout placeholder,
  master placeholder, master txStyles, mapped tx1) and the finding names
  the source. Backdrop resolution: the shape's own fill first, where
  a:noFill in spPr beats the p:style fillRef; a shape that paints nothing
  of its own takes the composite of every shape BENEATH the text's centre,
  by z, down to the slide/layout/master p:bg and theme bg1, alpha
  included. That composite is what finds text placed on top of something
  else, which is the commonest real contrast failure in a designed deck.
  Table cells are judged against their own tcPr fill. Gradients are
  averaged; a picture or pattern anywhere in the stack, or a cell taking
  its fill from the table style part, is reported as its own info finding
  rather than passing silently. Flags below min_ratio (default 4.5; large
  text >= 18pt or bold >= 14pt uses large_min_ratio, default 3.0); ratios
  under 2.0 escalate to error.
- diagram_glue: the un-tweakable-diagram smell. A connector end with no
  stCxn/endCxn glue whose endpoint touches a shape's bbox (within
  touch_tolerance_in, default 0.05") LOOKS attached but will not follow
  the shape when it moves. Glued ends and ends floating in space are fine.
"""

from __future__ import annotations

import colorsys
import re as _re
import unicodedata as _ud

from lxml import etree

from ..core import budget as _budget
from ..core.errors import PptMcpError, UnsupportedStructure
from ..core.package import PptxPackage, qn, resolve_target
from . import inherit as _inh
from .design import COLOR_SLOTS, _theme_part_of
from .read import (
    _ph,
    iter_shapes,
    shape_text,
    slides_in_scope,
    table_element,
    txbody_paragraphs,
)
from .shapes import (
    _connector_endpoints_slide,
    _iter_connectors,
    _shape_id,
    _slide_box,
)
from .text import _overflow_heuristic, _pct_value

EMU_PER_INCH = 914400

_RT_SLIDE_LAYOUT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)
_RT_SLIDE_MASTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
)

#: Shape families that are MEANT to touch their neighbours. A diagram's
#: connectors, tick marks, spines and bands sit on or against the things
#: they annotate, so pairing them inside a group produces nothing but
#: noise. Matched case-insensitively against the shape name.
DEFAULT_TOUCHING_NAMES = (
    "connector", "tick", "spine", "leader", "marker", "band", "curve",
    "arrow", "axis", "bracket", "callout", "rule", "divider", "underline",
)


#: Check registry: name -> (default options, caveat text). Order is
#: presentation order in results.
CHECKS: dict[str, tuple[dict, str]] = {
    "overlap": (
        {
            "min_overlap_pct": 40.0,
            "min_group_overlap_in": 0.05,
            "exclude_names": list(DEFAULT_TOUCHING_NAMES),
        },
        "bbox intersection, between top-level shapes and between siblings "
        "INSIDE each group (a percentage of the smaller shape at the top "
        "level, an absolute min_group_overlap_in in both dimensions inside "
        "a group, where a corner clip no percentage rule reaches is still "
        "a defect); skipped as deliberate composition: a shape fully "
        "inside another that sits behind it (background/panel), two "
        "untexted autoshapes overlapping (template band layering), content "
        "sitting >=85% on an untexted shape behind it (label-on-panel), "
        "and in-group members whose name matches exclude_names, the "
        "families meant to touch; placeholders with inherited geometry are "
        "resolved through the layout; rotation is ignored (boxes are "
        "unrotated). An in-group finding is geometry: a badge deliberately "
        "pinned half-on a box reads the same as a collision, so each one "
        "carries overlap_in for the caller to judge with",
    ),
    "off_slide": (
        {"partial_tolerance_in": 0.1},
        "bbox vs slide bounds; overhang under the tolerance is treated as "
        "deliberate full-bleed",
    ),
    "tiny_text": (
        {
            "body_min_pt": 14.0,
            "label_min_pt": 10.0,
            "label_max_chars": 30,
            "strict": False,
        },
        "run sizes resolved through the chain (run, paragraph, shape "
        "lstStyle, layout placeholder, master placeholder, master "
        "txStyles, presentation default); findings carry size_sources and "
        "inherited_size so an inherited size can be told from one written "
        "on the slide. Table cells are judged on explicit sizes only, "
        "since a cell inherits from the table style part rather than this "
        "chain. strict=True adds an info finding for runs whose size does "
        "not resolve anywhere instead of passing them",
    ),
    "overflow": (
        {"min_fill_ratio": 1.4, "min_fill_ratio_metrics": 1.0},
        "two models, and every finding names the one that produced it. "
        "With the optional metrics extra installed and each run's font file "
        "present, every run is measured against its own font and the wrap "
        "carries a per-line allowance calibrated against PowerPoint, so "
        "that model needs no extra margin here and min_fill_ratio_metrics "
        "is 1.0. Otherwise it is the original text-length-vs-frame-area "
        "estimate with no real font metrics, suppressed under the wider "
        "min_fill_ratio because that model's error is wider and "
        "uncalibrated. Shapes with spAutoFit are skipped because their "
        "frame grows with the text; confirm with export_slide_images "
        "before acting",
    ),
    "empty_placeholder": (
        {},
        "text-family placeholders only; picture/table/chart placeholders "
        "holding their content never fire",
    ),
    "missing_title": (
        {"title_shape_names": []},
        "screen readers and Outline view identify a slide by its title "
        "placeholder, so only a title or ctrTitle placeholder with text "
        "counts (a hidden or off-slide one included, which is how "
        "PowerPoint hides a title); a text box or bullet near the top of "
        "the slide does NOT, and is named in candidate_shape_ids. A deck "
        "that titles its slides with text boxes by design can declare "
        "their shape names in title_shape_names (case-insensitive exact "
        "match); screen readers still do not treat those as titles",
    ),
    "contrast": (
        {"min_ratio": 4.5, "large_min_ratio": 3.0},
        "WCAG-style ratio. Run colour resolves through the whole chain "
        "(rPr, paragraph defRPr, shape lstStyle, p:style fontRef, layout "
        "placeholder, master placeholder, master txStyles, mapped tx1) and "
        "the finding names the source. The backdrop is the shape's own "
        "fill when it has one (a:noFill in spPr beats the p:style "
        "fillRef), otherwise every shape beneath the text's centre "
        "composited by z down to the slide background, alpha included. "
        "Table cells are judged against their own tcPr fill, falling back "
        "to a table style DEFINED in tableStyles.xml, whose parts layer in "
        "the ECMA-376 CT_TableStyle order (wholeTbl, row bands, column "
        "bands, last/first column, last row, first row, corner cells, each "
        "beating the one before). PowerPoint's built-in styles are "
        "referenced by GUID and their definitions are not in the file, so "
        "those stay unresolved, as does a corner cell whose emphasis flags "
        "renderers disagree about. A backdrop that cannot "
        "be resolved to a colour, and a run whose colour is not stated "
        "anywhere this read can reach, are reported as info findings "
        "(reason unresolved_colour_source) and NEVER as a failure: a gate "
        "severity built on a guessed colour is worse than no check. "
        "Gradients are averaged and "
        "lumMod/lumOff/tint/shade are approximated, so treat borderline "
        "ratios as render-and-look",
    ),
    "diagram_glue": (
        {"touch_tolerance_in": 0.05},
        "an unglued connector end touching a shape's bbox looks attached "
        "but will not follow the shape when it moves",
    ),
}


# ---------------------------------------------------------------- options


def _is_word_list(default) -> bool:
    """Whether an option's default marks it as a list of name words, the
    shape exclude_names has."""
    return (
        isinstance(default, list)
        and all(isinstance(d, str) for d in default)
    )


def _normalize_checks(checks) -> list[tuple[str, dict]]:
    """Validate the checks array: None = all checks with defaults; entries
    are check names or {"check": name, <option>: value} dicts."""
    if checks is None:
        return [(name, dict(defaults)) for name, (defaults, _c) in CHECKS.items()]
    if isinstance(checks, str):
        checks = [checks]
    # A bare option dict is a single-entry battery; the docstring's own
    # example uses this shape (insane round 2 L3).
    if isinstance(checks, dict):
        checks = [checks]
    if not isinstance(checks, list) or not checks:
        raise PptMcpError(
            f"checks must be a non-empty list from {list(CHECKS)} (or None "
            "for all of them)"
        )
    out: list[tuple[str, dict]] = []
    seen: set[str] = set()
    for entry in checks:
        if isinstance(entry, str):
            name, opts = entry.strip().lower(), {}
        elif isinstance(entry, dict):
            if "check" not in entry:
                raise PptMcpError(
                    f'check dict needs a "check" key naming one of '
                    f"{list(CHECKS)}, got {entry!r}"
                )
            name = str(entry["check"]).strip().lower()
            opts = {k: v for k, v in entry.items() if k != "check"}
        else:
            raise PptMcpError(
                f"invalid checks entry {entry!r}: use a check name or "
                '{"check": name, option: value}'
            )
        if name not in CHECKS:
            raise PptMcpError(
                f"unknown check {name!r}; one of: {', '.join(CHECKS)}"
            )
        defaults, _caveat = CHECKS[name]
        unknown = sorted(set(opts) - set(defaults))
        if unknown:
            raise PptMcpError(
                f"unknown option(s) for check {name!r}: {', '.join(unknown)}"
                f"; valid: {sorted(defaults) or 'none'}"
            )
        merged = dict(defaults)
        for k, v in opts.items():
            # A word list is the one option kind a bare string is a
            # reasonable request for, and list("band") would silently
            # spell it out into four one-letter words.
            if _is_word_list(defaults[k]) and isinstance(v, str):
                merged[k] = [v]
                continue
            try:
                merged[k] = type(defaults[k])(v)
            except (TypeError, ValueError):
                raise PptMcpError(
                    f"option {k}={v!r} for check {name!r} must be "
                    f"{type(defaults[k]).__name__}"
                ) from None
        # The container was validated above and the ELEMENTS were not, so
        # a number inside exclude_names reached .casefold() and raised a
        # raw AttributeError instead of the refusal every other bad value
        # gets.
        for k, v in merged.items():
            if not _is_word_list(defaults[k]):
                continue
            for entry in v:
                if not isinstance(entry, str):
                    raise PptMcpError(
                        f"option {k} entry {entry!r} for check {name!r} "
                        f"must be str, not {type(entry).__name__}"
                    )
        if name not in seen:
            seen.add(name)
            out.append((name, merged))
    return out


# ------------------------------------------------------------ slide context


def _gentle_box(
    elem: etree._Element,
    chain: list,
    pkg: PptxPackage | None = None,
    part: str | None = None,
) -> tuple[tuple[float, float, float, float] | None, bool]:
    """(slide-space bbox, inherited) for one shape. Never raises.

    A placeholder with an empty p:spPr has no box of its own, and returning
    None for it used to make it invisible to every geometry check: the
    overlap and off_slide checks passed slides where a body placeholder
    printed straight through the caption below it. The box it RENDERS with
    is in the layout or the master, so the chain is resolved and the answer
    says it was inherited.
    """
    try:
        return _slide_box(elem, chain), False
    except (UnsupportedStructure, PptMcpError, ValueError, TypeError):
        pass
    if pkg is not None and part is not None and not chain:
        box = _inh.inherited_box(pkg, part, elem)
        if box is not None:
            return tuple(float(v) for v in box), True
    return None, False


def _slide_size(pkg: PptxPackage) -> tuple[int, int]:
    sld_sz = pkg.presentation().find(qn("p:sldSz"))
    if sld_sz is None:  # ECMA default: 10 x 7.5 in
        return 9144000, 6858000
    return int(sld_sz.get("cx")), int(sld_sz.get("cy"))


def _sp_tree_of(pkg: PptxPackage, part: str) -> etree._Element | None:
    return pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")


def _is_hidden(elem: etree._Element) -> bool:
    for child in elem:
        if etree.QName(child).localname.startswith("nv"):
            cnvpr = child.find(qn("p:cNvPr"))
            return cnvpr is not None and cnvpr.get("hidden") == "1"
    return False


class _SlideCtx:
    """Everything the per-slide checks share, computed once."""

    def __init__(self, pkg: PptxPackage, rec: dict):
        self.pkg = pkg
        self.rec = rec
        self.part = rec["part"]
        self.sp_tree = _sp_tree_of(pkg, self.part)
        self.slide_cx, self.slide_cy = _slide_size(pkg)
        # Top-level shapes (direct spTree children), for the geometry checks.
        self.top: list[dict] = []
        # All shapes (groups recursed) with slide-space boxes, for hit tests.
        self.all: list[dict] = []
        if self.sp_tree is not None:
            for elem, kind, z, parent in iter_shapes(self.sp_tree):
                srec = {
                    "elem": elem,
                    "kind": kind,
                    "id": _shape_id(elem),
                    "name": _shape_name(elem),
                    "z": z,
                    "hidden": _is_hidden(elem),
                }
                if parent is None:
                    srec["box"], srec["box_inherited"] = _gentle_box(
                        elem, [], pkg, self.part
                    )
                    self.top.append(srec)
                    if kind != "group":
                        self.all.append(srec)
                else:
                    # box resolved through the ancestor chain lazily below
                    srec["box"] = None
                    srec["box_inherited"] = False
                    srec["group_id"] = parent
                    if kind != "group":
                        self.all.append(srec)
            # Resolve grouped shapes' slide-space boxes via _iter with chains.
            self._resolve_group_boxes()
        self._resolver: _ColorResolver | None = None
        # Per-check counters a check wants VISIBLE in the result.
        # Filtering nobody can see is filtering nobody can check.
        self.stats: dict[str, dict[str, int]] = {}

    def _resolve_group_boxes(self) -> None:
        chains: dict[int, list] = {}
        # id -> the id of the OUTERMOST group it lives in, so the overlap
        # check can compare the members of one diagram against each other.
        roots: dict[int, int] = {}

        def _walk(container, chain):
            for child in container:
                if child.tag == qn("p:grpSp"):
                    _walk(child, chain + [child])
                else:
                    sid = _shape_id(child)
                    if sid is not None and chain:
                        chains[sid] = chain
                        roots[sid] = _shape_id(chain[0])

        _walk(self.sp_tree, [])
        for srec in self.all:
            srec["group_root"] = roots.get(srec["id"])
            if srec["box"] is None and srec["id"] in chains:
                srec["box"], srec["box_inherited"] = _gentle_box(
                    srec["elem"], chains[srec["id"]]
                )

    def resolver(self) -> "_ColorResolver":
        if self._resolver is None:
            self._resolver = _ColorResolver(self.pkg, self.part)
        return self._resolver

    def finding(self, check: str, severity: str, message: str, fix: str, **extra) -> dict:
        # Contract (production-test finding): every finding carries the
        # structured location fields, never prose alone. shape_ids is
        # always present (possibly empty) so clients can address shapes
        # without parsing the message.
        out = {
            "check": check,
            "severity": severity,
            "slide_index": self.rec["index"],
            "slide_id": self.rec["slide_id"],
            "shape_ids": [],
            "message": message,
            "fix": fix,
        }
        out.update(extra)
        return out


def _shape_name(elem: etree._Element) -> str:
    for child in elem:
        if etree.QName(child).localname.startswith("nv"):
            cnvpr = child.find(qn("p:cNvPr"))
            return cnvpr.get("name", "") if cnvpr is not None else ""
    return ""


def _label(srec: dict) -> str:
    name = srec.get("name") or ""
    return f"shape {srec['id']}" + (f" ({name!r})" if name else "")


# ------------------------------------------------------------ color algebra


_PRST_CLR = {"black": "000000", "white": "FFFFFF"}


def _related(pkg: PptxPackage, part: str, rel_type: str) -> str | None:
    try:
        rels = pkg.rels_for(part)
    except KeyError:
        return None
    for rel in rels.getroot():
        if rel.get("Type") == rel_type and rel.get("TargetMode") != "External":
            return resolve_target(part, rel.get("Target", ""))
    return None


class _ColorResolver:
    """Resolve DrawingML color elements to RRGGBB hex for one slide,
    honoring the clrMap override chain (slide/layout clrMapOvr with
    a:overrideClrMapping, else the master's p:clrMap) and the master's
    theme clrScheme. Unresolvable colors come back None, never a guess."""

    def __init__(self, pkg: PptxPackage, slide_part: str):
        self.pkg = pkg
        self.slide_part = slide_part
        self.layout_part = _related(pkg, slide_part, _RT_SLIDE_LAYOUT)
        self.master_part = (
            _related(pkg, self.layout_part, _RT_SLIDE_MASTER)
            if self.layout_part
            else None
        )
        self.theme_colors = self._theme_colors()
        self.clr_map = self._clr_map()

    def _theme_colors(self) -> dict[str, str]:
        if not self.master_part or not self.pkg.has_part(self.master_part):
            return {}
        try:
            theme_part = _theme_part_of(self.pkg, self.master_part)
        except (UnsupportedStructure, PptMcpError):
            return {}
        if not self.pkg.has_part(theme_part):
            return {}
        scheme = self.pkg.root(theme_part).find(
            f"{qn('a:themeElements')}/{qn('a:clrScheme')}"
        )
        if scheme is None:
            return {}
        out: dict[str, str] = {}
        for slot in COLOR_SLOTS:
            el = scheme.find(qn(f"a:{slot}"))
            if el is None:
                continue
            srgb = el.find(qn("a:srgbClr"))
            if srgb is not None and srgb.get("val"):
                out[slot] = srgb.get("val").upper()
                continue
            sysclr = el.find(qn("a:sysClr"))
            if sysclr is not None and sysclr.get("lastClr"):
                out[slot] = sysclr.get("lastClr").upper()
        return out

    def _clr_map(self) -> dict[str, str]:
        # Slide, then layout, may carry p:clrMapOvr/a:overrideClrMapping.
        for part in (self.slide_part, self.layout_part):
            if not part or not self.pkg.has_part(part):
                continue
            ovr = self.pkg.root(part).find(
                f"{qn('p:clrMapOvr')}/{qn('a:overrideClrMapping')}"
            )
            if ovr is not None:
                return dict(ovr.attrib)
        if self.master_part and self.pkg.has_part(self.master_part):
            cmap = self.pkg.root(self.master_part).find(qn("p:clrMap"))
            if cmap is not None:
                return dict(cmap.attrib)
        return {
            "bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2",
        }

    def scheme_hex(self, val: str) -> str | None:
        slot = self.clr_map.get(val, val)
        return self.theme_colors.get(slot)

    def resolve(self, color_el: etree._Element | None) -> str | None:
        """One a:srgbClr / a:schemeClr / a:sysClr / a:prstClr / a:scrgbClr
        element to hex, with lumMod/lumOff/tint/shade applied (approx)."""
        if color_el is None:
            return None
        local = etree.QName(color_el).localname
        base: str | None = None
        if local == "srgbClr":
            base = (color_el.get("val") or "").upper()
        elif local == "schemeClr":
            val = color_el.get("val", "")
            if val == "phClr":  # placeholder color: context-dependent
                return None
            base = self.scheme_hex(val)
        elif local == "sysClr":
            base = (color_el.get("lastClr") or "").upper() or None
        elif local == "prstClr":
            base = _PRST_CLR.get(color_el.get("val", ""))
        elif local == "scrgbClr":
            try:
                base = "".join(
                    f"{round(int(color_el.get(k, '0')) / 100000 * 255):02X}"
                    for k in ("r", "g", "b")
                )
            except ValueError:
                base = None
        if not base or len(base) != 6:
            return None
        return _apply_transforms(base, color_el)


def _apply_transforms(hexstr: str, el: etree._Element) -> str:
    """Approximate lumMod/lumOff/tint/shade (HLS for luminance ops)."""
    try:
        r, gg, b = (int(hexstr[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    except ValueError:
        return hexstr
    for child in el:
        local = etree.QName(child).localname
        raw = child.get("val")
        if raw is None:
            continue
        # _pct_value handles both '50000' (thousandths) and '50%' -> 50.0
        try:
            val = _pct_value(raw, 100.0)
        except ValueError:
            continue
        frac = val / 100.0
        if local == "shade":
            r, gg, b = r * frac, gg * frac, b * frac
        elif local == "tint":
            r, gg, b = (c * frac + (1 - frac) for c in (r, gg, b))
        elif local in ("lumMod", "lumOff"):
            h, lum, s = colorsys.rgb_to_hls(r, gg, b)
            lum = lum * frac if local == "lumMod" else min(1.0, max(0.0, lum + frac))
            r, gg, b = colorsys.hls_to_rgb(h, lum, s)
    clamp = lambda c: min(255, max(0, round(c * 255)))  # noqa: E731
    return f"{clamp(r):02X}{clamp(gg):02X}{clamp(b):02X}"


def _rel_luminance(hexstr: str) -> float:
    def chan(c: float) -> float:
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4

    r, g, b = (int(hexstr[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    return 0.2126 * chan(r) + 0.7152 * chan(g) + 0.0722 * chan(b)


def contrast_ratio(hex1: str, hex2: str) -> float:
    l1, l2 = _rel_luminance(hex1), _rel_luminance(hex2)
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


_COLOR_TAGS = tuple(
    qn(t) for t in ("a:srgbClr", "a:schemeClr", "a:sysClr", "a:prstClr", "a:scrgbClr")
)


def _first_color_child(parent: etree._Element | None) -> etree._Element | None:
    if parent is None:
        return None
    for child in parent:
        if child.tag in _COLOR_TAGS:
            return child
    return None


#: Sentinel: the shape paints no pixels of its own; look behind it.
_TRANSPARENT = "transparent"


def _own_fill_hex(
    sppr_owner: etree._Element, resolver: _ColorResolver
) -> tuple[str | None, str | None]:
    """(hex | _TRANSPARENT | None, skip_reason) for a shape's OWN fill:
    explicit spPr fill first, then the p:style fillRef. _TRANSPARENT means
    noFill (or no fill information at all); skip_reason is set (hex None)
    for picture/pattern fills."""
    sppr = sppr_owner.find(qn("p:spPr"))
    if sppr is not None:
        for child in sppr:
            local = etree.QName(child).localname
            if local == "solidFill":
                return resolver.resolve(_first_color_child(child)), None
            if local == "gradFill":
                hexes = [
                    h
                    for gs in child.iter(qn("a:gs"))
                    if (h := resolver.resolve(_first_color_child(gs)))
                ]
                if hexes:
                    avg = tuple(
                        round(sum(int(h[i:i + 2], 16) for h in hexes) / len(hexes))
                        for i in (0, 2, 4)
                    )
                    return "".join(f"{c:02X}" for c in avg), None
                return None, "gradient with unresolvable stops"
            if local in ("blipFill", "pattFill"):
                return None, f"{local} (image/pattern) fill"
            if local == "noFill":
                return _TRANSPARENT, None
    style = sppr_owner.find(qn("p:style"))
    if style is not None:
        fillref = style.find(qn("a:fillRef"))
        if fillref is not None:
            try:
                idx = int(fillref.get("idx", "0"))
            except ValueError:
                idx = 0
            if idx > 0:
                resolved = resolver.resolve(_first_color_child(fillref))
                if resolved:
                    return resolved, None
                return None, "fillRef color unresolvable (phClr)"
            return _TRANSPARENT, None
    return _TRANSPARENT, None


def _own_fill_alpha(sppr_owner: etree._Element) -> float | None:
    """Opacity (0..1) of the shape's OWN explicit solid fill, or None when
    fully opaque / no explicit solid fill. Reads the a:alpha child of the
    solidFill's color element (both '10000' thousandths and '10%' forms).
    Production-test finding: tinted fills (accent1 at 10% alpha, the
    matrix-cell default) used to be judged at full strength, flagging dark
    text on a pale wash as a contrast error."""
    sppr = sppr_owner.find(qn("p:spPr"))
    if sppr is None:
        return None
    solid = sppr.find(qn("a:solidFill"))
    if solid is None:
        return None
    color_el = _first_color_child(solid)
    if color_el is None:
        return None
    alpha_el = color_el.find(qn("a:alpha"))
    if alpha_el is None:
        return None
    try:
        pct = _pct_value(alpha_el.get("val"), 100.0)
    except ValueError:
        return None
    return min(1.0, max(0.0, pct / 100.0))


def _blend_hex(top_hex: str, under_hex: str, alpha: float) -> str:
    """Composite a translucent top color over an opaque backdrop."""
    out = []
    for i in (0, 2, 4):
        t = int(top_hex[i:i + 2], 16)
        u = int(under_hex[i:i + 2], 16)
        out.append(f"{round(t * alpha + u * (1.0 - alpha)):02X}")
    return "".join(out)


def _backdrop_hex(
    srec: dict, ctx: "_SlideCtx", resolver: _ColorResolver
) -> tuple[str | None, str | None]:
    """The colour a shape's text is actually read against: every visible
    shape BENEATH it whose box contains the text's centre, composited in
    z order down to the slide background.

    This used to demand that the shape beneath fully CONTAIN the one in
    front, and it took only the smallest such shape rather than the stack.
    That missed the commonest real contrast failure in a designed deck:
    an unfilled text box floating on a band, where the label overhangs the
    band it sits on. The centre is what sits on the colour, and a
    half-transparent band over another band is what the eye sees, so the
    whole stack gets composited.

    Returns (hex, None) or (None, reason). A picture or pattern anywhere
    in the stack, or a background no static read can resolve, comes back
    as a REASON, which the caller reports rather than passing silently.
    """
    box = srec.get("box")
    base = _background_hex(resolver)
    if base is None:
        return None, "slide background is a picture, pattern or gradient"
    if box is None:
        return base, None

    cx = box[0] + box[2] / 2.0
    cy = box[1] + box[3] / 2.0
    order = {id(s["elem"]): i for i, s in enumerate(ctx.all)}
    mine = order.get(id(srec["elem"]), -1)

    stack: list[tuple[str, float]] = []
    for other in ctx.all:
        if other is srec or other["hidden"] or other["box"] is None:
            continue
        if other["kind"] == "connector":
            continue
        if order.get(id(other["elem"]), -1) >= mine:
            continue  # in front of the text, or the text itself
        ox, oy, ocx, ocy = other["box"]
        if not (ox <= cx <= ox + ocx and oy <= cy <= oy + ocy):
            continue
        if other["kind"] in ("picture", "chart", "diagram"):
            # A photograph under the text has no single colour to measure
            # against, and picking one would be a confident guess.
            return None, f"{other['kind']} beneath the text ({_label(other)})"
        hexval, skip = _own_fill_hex(other["elem"], resolver)
        if skip:
            return None, f"{skip} on {_label(other)}, beneath the text"
        if hexval is None or hexval == _TRANSPARENT:
            continue
        alpha = _own_fill_alpha(other["elem"])
        stack.append((hexval, 1.0 if alpha is None else alpha))

    current = base
    for hexval, alpha in stack:  # document order is bottom-up
        current = _blend_hex(hexval, current, alpha)
    return current, None


def _background_hex(resolver: _ColorResolver) -> str | None:
    """Effective slide background: p:bg on the slide, else layout, else
    master, else mapped bg1."""
    pkg = resolver.pkg
    for part in (resolver.slide_part, resolver.layout_part, resolver.master_part):
        if not part or not pkg.has_part(part):
            continue
        bg = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:bg')}")
        if bg is None:
            continue
        bgpr = bg.find(qn("p:bgPr"))
        if bgpr is not None:
            solid = bgpr.find(qn("a:solidFill"))
            if solid is not None:
                return resolver.resolve(_first_color_child(solid))
            return None  # gradient/picture background: not resolvable here
        bgref = bg.find(qn("p:bgRef"))
        if bgref is not None:
            resolved = resolver.resolve(_first_color_child(bgref))
            if resolved:
                return resolved
    return resolver.scheme_hex("bg1")


# ----------------------------------------------------------------- checks


def _boxes_intersect_area(a, b) -> float:
    ax, ay, acx, acy = a
    bx, by, bcx, bcy = b
    w = min(ax + acx, bx + bcx) - max(ax, bx)
    h = min(ay + acy, by + bcy) - max(ay, by)
    return w * h if w > 0 and h > 0 else 0.0


def _contains(outer, inner, eps: float = 9525.0) -> bool:
    ox, oy, ocx, ocy = outer
    ix, iy, icx, icy = inner
    return (
        ix >= ox - eps
        and iy >= oy - eps
        and ix + icx <= ox + ocx + eps
        and iy + icy <= oy + ocy + eps
    )


#: Anything that is not a word character separates one name token from the
#: next, so "lane band", "band-2" and "Band" all match "band" while
#: "Bandwidth" and "Brand" do not. \w is UNICODE-aware by default in
#: Python 3: an ASCII-only class dropped every non-Latin character, so a
#: shape named "밴드 one" tokenized to {"one"} and exclude_names=["밴드"]
#: tokenized to nothing at all, which put the whole option out of reach
#: for anyone naming shapes in Korean, Japanese, Chinese, Cyrillic or
#: accented Latin.
_NAME_TOKENS = _re.compile(r"\w+")


def _nfc_fold(text: str) -> str:
    """One spelling, one case, for both sides of the comparison.

    NFC first, because a combining mark is not a word character: the
    decomposed spelling of "bänd" (a + U+0308 + nd), which is what a name
    authored on macOS routinely carries, tokenized to {"ba", "nd"} and
    could never match the precomposed exclusion word the user typed.
    casefold, not lower, because casefold is the case-insensitive
    comparison Unicode actually defines.
    """
    return _ud.normalize("NFC", text or "").casefold()


def _name_tokens(text: str) -> set[str]:
    """Casefolded word tokens of a name, normalized before tokenizing so
    a decomposed name splits where a reader would split it."""
    return {
        t.casefold()
        for t in _NAME_TOKENS.findall(_ud.normalize("NFC", text or ""))
    }


def _touching_family(s: dict, words) -> bool:
    """Whether a shape's name puts it in a family meant to touch.

    Substring matching silently exempted "Bandwidth chart" and "Brand box"
    from the in-group pass, because both contain "band"; the same hole
    swallowed anything containing rule, tick, arrow or axis. Names are
    matched a TOKEN at a time now.
    """
    return bool(_name_tokens(s.get("name") or "") & set(words))


def _decoration(s: dict) -> bool:
    """Untexted plain shape: template band, strip or panel material."""
    return (
        etree.QName(s["elem"]).localname == "sp"
        and not shape_text(s["elem"]).strip()
    )


def _deliberate_composition(a: dict, b: dict, inter: float) -> bool:
    """True when two intersecting boxes are a designed stack, not a
    collision. The same three rules hold inside a group as at the top
    level, and applying them there is what keeps a diagram's labels from
    being reported against the shapes they are labelling."""
    # Full inclusion with the outer BEHIND: a background or a panel.
    if _contains(a["box"], b["box"]) and a["z"] < b["z"]:
        return True
    if _contains(b["box"], a["box"]) and b["z"] < a["z"]:
        return True
    # Decoration on decoration: template layering.
    if _decoration(a) and _decoration(b):
        return True
    # Label on panel: the front shape sits almost entirely on an untexted
    # shape behind it.
    back, front_s = (a, b) if a["z"] < b["z"] else (b, a)
    front_area = front_s["box"][2] * front_s["box"][3]
    return bool(
        _decoration(back) and front_area > 0 and inter / front_area >= 0.85
    )


def _check_overlap(ctx: _SlideCtx, opts: dict) -> list[dict]:
    findings = []
    cand = [
        s
        for s in ctx.top
        if not s["hidden"]
        and s["kind"] != "connector"
        and s["box"] is not None
        and s["box"][2] > 0
        and s["box"][3] > 0
    ]
    min_pct = float(opts["min_overlap_pct"])

    for i in range(len(cand)):
        for j in range(i + 1, len(cand)):
            a, b = cand[i], cand[j]
            inter = _boxes_intersect_area(a["box"], b["box"])
            if inter <= 0:
                continue
            if _deliberate_composition(a, b, inter):
                continue
            smaller = min(a["box"][2] * a["box"][3], b["box"][2] * b["box"][3])
            pct = inter / smaller * 100.0
            if pct < min_pct:
                continue
            front = b if b["z"] > a["z"] else a
            findings.append(
                ctx.finding(
                    "overlap",
                    "warning",
                    f"{_label(a)} and {_label(b)} overlap by "
                    f"{round(pct)}% of the smaller shape; {_label(front)} "
                    "is in front and hides the other",
                    f"set_shape(slide={ctx.rec['index']}, shape="
                    f"{front['id']}, dx=..., dy=...) to nudge it clear, or "
                    f"align_shapes/distribute_shapes(slide="
                    f"{ctx.rec['index']}, ids=[{a['id']}, {b['id']}]) to "
                    "lay them out; if the stacking is deliberate, "
                    "set_z_order controls which shows",
                    shape_ids=[a["id"], b["id"]],
                    overlap_pct=round(pct, 1),
                )
            )
    findings.extend(_check_group_overlap(ctx, opts))
    return findings


def _check_group_overlap(ctx: _SlideCtx, opts: dict) -> list[dict]:
    """Overlap between SIBLINGS inside a group.

    The top-level pass treats a group as one box, on the reasoning that a
    group's interior is its own business. That reasoning fails on exactly
    the slides this server's own diagram tooling produces, since
    generate_diagram and svg_to_shapes both return grouped output: every
    real collision in a timeline is between two boxes inside the group,
    and the check went blind on all of them.

    Two things keep this usable as a gate rather than a hint. Members of
    families meant to touch (connectors, ticks, spines, bands) are
    excluded by name, and the bar is an ABSOLUTE overlap in BOTH
    dimensions rather than a percentage: a corner clip of a couple of
    hundredths of an inch between two timeline boxes is a real defect that
    no percentage-of-area rule would ever reach.
    """
    # `or DEFAULT` made an EMPTY list fall back to the default, so there
    # was no way to ask for an unfiltered in-group pass. Absent and empty
    # are different requests.
    words = opts["exclude_names"] if "exclude_names" in opts else (
        DEFAULT_TOUCHING_NAMES
    )
    if isinstance(words, str):
        words = (words,)
    # Normalized and casefolded the same way the names are, so a non-ASCII
    # exclusion word compares equal to the token it is meant to match
    # whichever spelling either side carries. (_normalize_checks already
    # refuses a None here, so there is no None branch to write.)
    words = tuple(_nfc_fold(w) for w in words)
    bar = float(opts.get("min_group_overlap_in", 0.05)) * EMU_PER_INCH

    groups: dict[int, list[dict]] = {}
    suppressed = 0
    for s in ctx.all:
        root = s.get("group_root")
        if root is None or s["hidden"] or s["kind"] == "connector":
            continue
        if s["box"] is None or s["box"][2] <= 0 or s["box"][3] <= 0:
            continue
        if _touching_family(s, words):
            suppressed += 1
            continue
        groups.setdefault(root, []).append(s)
    ctx.stats.setdefault("overlap", {})["suppressed_by_exclude_names"] = (
        ctx.stats.get("overlap", {}).get("suppressed_by_exclude_names", 0)
        + suppressed
    )

    findings = []
    for root, members in groups.items():
        for i in range(len(members)):
            for j in range(i + 1, len(members)):
                a, b = members[i], members[j]
                ax, ay, acx, acy = a["box"]
                bx, by, bcx, bcy = b["box"]
                ow = min(ax + acx, bx + bcx) - max(ax, bx)
                oh = min(ay + acy, by + bcy) - max(ay, by)
                if ow <= bar or oh <= bar:
                    continue
                if _deliberate_composition(
                    a, b, _boxes_intersect_area(a["box"], b["box"])
                ):
                    continue
                front = b if b["z"] > a["z"] else a
                findings.append(
                    ctx.finding(
                        "overlap",
                        "warning",
                        f"inside group {root}, {_label(a)} and {_label(b)} "
                        f"overlap by {ow / EMU_PER_INCH:.3f} x "
                        f"{oh / EMU_PER_INCH:.3f} in; {_label(front)} is in "
                        "front",
                        f"set_shape(slide={ctx.rec['index']}, "
                        f"shape={front['id']}, dx=..., dy=...) moves the "
                        "member itself (moving the group moves both), or "
                        f"ungroup_shapes(slide={ctx.rec['index']}, "
                        f"shape={root}) first if the layout needs rebuilding",
                        shape_ids=[a["id"], b["id"]],
                        group_id=root,
                        overlap_in=[
                            round(ow / EMU_PER_INCH, 3),
                            round(oh / EMU_PER_INCH, 3),
                        ],
                    )
                )
    return findings


def _check_off_slide(ctx: _SlideCtx, opts: dict) -> list[dict]:
    findings = []
    tol = float(opts["partial_tolerance_in"]) * EMU_PER_INCH
    for s in ctx.top:
        if s["hidden"] or s["box"] is None:
            continue
        x, y, cx, cy = s["box"]
        fully_out = (
            x + cx <= 0 or y + cy <= 0 or x >= ctx.slide_cx or y >= ctx.slide_cy
        )
        overhang = max(
            0.0, -x, -y, (x + cx) - ctx.slide_cx, (y + cy) - ctx.slide_cy
        )
        if fully_out:
            findings.append(
                ctx.finding(
                    "off_slide",
                    "error",
                    f"{_label(s)} lies entirely OUTSIDE the slide "
                    f"(at {round(x / EMU_PER_INCH, 2)}, "
                    f"{round(y / EMU_PER_INCH, 2)} in on a "
                    f"{round(ctx.slide_cx / EMU_PER_INCH, 2)} x "
                    f"{round(ctx.slide_cy / EMU_PER_INCH, 2)} in slide); it "
                    "renders nowhere",
                    f"set_shape(slide={ctx.rec['index']}, shape={s['id']}, "
                    f"x={_clamp_in(x, cx, ctx.slide_cx)}, "
                    f"y={_clamp_in(y, cy, ctx.slide_cy)}) brings it fully "
                    "on-slide, or delete_shape if it is debris",
                    shape_ids=[s["id"]],
                    extent="full",
                )
            )
        elif overhang > tol:
            findings.append(
                ctx.finding(
                    "off_slide",
                    "warning",
                    f"{_label(s)} extends "
                    f"{round(overhang / EMU_PER_INCH, 2)} in past the slide "
                    "edge; the overhang is cut off when presenting",
                    f"set_shape(slide={ctx.rec['index']}, shape={s['id']}, "
                    f"x={_clamp_in(x, cx, ctx.slide_cx)}, "
                    f"y={_clamp_in(y, cy, ctx.slide_cy)}) pulls it inside "
                    "(or shrink with w=/h=)",
                    shape_ids=[s["id"]],
                    extent="partial",
                    overhang_in=round(overhang / EMU_PER_INCH, 2),
                )
            )
    return findings


def _clamp_in(pos: float, size: float, bound: float) -> float:
    return round(max(0.0, min(pos, bound - size)) / EMU_PER_INCH, 2)


def _explicit_sizes(paragraphs: list[etree._Element]) -> tuple[list[float], int]:
    """(explicit run sizes in pt for runs that carry text, skipped count).

    Table cells only. A cell's inherited size comes from the table style
    part, not from the placeholder chain, so resolving it the way a shape's
    runs resolve would produce a confident wrong number.
    """
    sizes: list[float] = []
    skipped = 0
    for p in paragraphs:
        for r in p.findall(qn("a:r")):
            t = r.find(qn("a:t"))
            if t is None or not (t.text or "").strip():
                continue
            rpr = r.find(qn("a:rPr"))
            sz = rpr.get("sz") if rpr is not None else None
            if sz is None:
                skipped += 1
                continue
            try:
                sizes.append(int(sz) / 100.0)
            except ValueError:
                skipped += 1
    return sizes, skipped


def _run_sizes(
    ctx: "_SlideCtx",
    elem: etree._Element,
    paragraphs: list[etree._Element],
) -> tuple[list[tuple[float, str]], int]:
    """([(size in pt, where it came from)], unresolved count) for every run
    that carries text.

    This used to read a:rPr @sz and skip anything without one, which is
    every ordinary body run on a templated deck: those runs render at the
    master's size, around 24pt on a 16:9 Title-and-Content layout, so the
    check went blind exactly where the overflow risk is highest. The chain
    is resolved through ops/inherit.py now, and a run whose size still
    cannot be found is counted rather than quietly passed.
    """
    # A cached normAutofit scale is what the text ACTUALLY renders at, and
    # _check_overflow in this same battery already applies it. Ignoring it
    # here meant a 28pt run shrunk to 7pt produced no finding while the
    # check beside it called the same shape crowded: two checks in one
    # battery disagreeing about the size of the same text.
    # Reported as its own field. Appending it to the SOURCE string made
    # the source unequal to "run", which is the test the finding uses to
    # decide whether a size was INHERITED, so a size written on the slide
    # as sz="2800" came back inherited_size=true and sent the caller
    # looking in the layout for a fix that belongs on the slide.
    scale = _autofit_scale(elem)
    sizes: list[tuple[float, str]] = []
    unresolved = 0
    for p in paragraphs:
        for r in p.findall(qn("a:r")):
            t = r.find(qn("a:t"))
            if t is None or not (t.text or "").strip():
                continue
            size, source = _inh.resolve_run_size_pt(
                ctx.pkg, ctx.part, elem, p, r.find(qn("a:rPr"))
            )
            if size is None:
                unresolved += 1
                continue
            if scale != 1.0:
                size = round(size * scale, 2)
            sizes.append((size, source))
    return sizes, unresolved, scale


def _autofit_scale(elem: etree._Element) -> float:
    """The cached normAutofit fontScale as a multiplier, or 1.0."""
    body = elem.find(qn("p:txBody"))
    bodypr = body.find(qn("a:bodyPr")) if body is not None else None
    norm = bodypr.find(qn("a:normAutofit")) if bodypr is not None else None
    if norm is None or norm.get("fontScale") is None:
        return 1.0
    try:
        return _pct_value(norm.get("fontScale"), 100.0) / 100.0
    except (ValueError, TypeError):
        return 1.0


def _check_tiny_text(ctx: _SlideCtx, opts: dict) -> list[dict]:
    findings = []
    body_min = float(opts["body_min_pt"])
    label_min = float(opts["label_min_pt"])
    label_chars = int(opts["label_max_chars"])
    strict = bool(opts.get("strict", False))
    for s in ctx.all:
        if s["hidden"]:
            continue
        unresolved = 0
        if s["kind"] == "table":
            tbl = table_element(s["elem"])
            if tbl is None:
                continue
            sized: list[tuple[float, str]] = []
            for tc in tbl.iter(qn("a:tc")):
                cell_sizes, skipped = _explicit_sizes(
                    tc.findall(f"{qn('a:txBody')}/{qn('a:p')}")
                )
                sized.extend((sz, "run") for sz in cell_sizes)
                unresolved += skipped
            scale = 1.0
            floor, role = label_min, "table"
        else:
            paras = txbody_paragraphs(s["elem"])
            if not paras:
                continue
            sized, unresolved, scale = _run_sizes(ctx, s["elem"], paras)
            text = shape_text(s["elem"]).strip()
            is_label = len(text) <= label_chars and "\n" not in text
            floor = label_min if is_label else body_min
            role = "label" if is_label else "body"

        if strict and unresolved:
            findings.append(
                ctx.finding(
                    "tiny_text",
                    "info",
                    f"{_label(s)} has {unresolved} run(s) whose font size "
                    "does not resolve through the layout, the master or "
                    "the presentation default; size was not checked there",
                    "render the slide (export_slide_images) and read the "
                    "size off the picture, or set one explicitly with "
                    f"format_text(slide={ctx.rec['index']}, "
                    f"shape={s['id']}, size_pt=...)",
                    shape_ids=[s["id"]],
                    unresolved_runs=unresolved,
                )
            )
        too_small = sorted({sz for sz, _src in sized if sz < floor})
        if not too_small:
            continue
        # Where the offending sizes came from decides where the fix goes:
        # an inherited size is the layout's or the master's problem until
        # somebody overrides it on the slide.
        sources = sorted({
            src for sz, src in sized if sz < floor and src != "run"
        })
        note = (
            f"; size inherited from the {', '.join(sources)} text style"
            if sources else ""
        )
        if scale != 1.0:
            note += (
                f"; shrunk to {round(scale * 100)}% of the written size by "
                "the frame's cached autofit"
            )
        findings.append(
            ctx.finding(
                "tiny_text",
                "warning",
                f"{_label(s)} has {role} text at "
                f"{', '.join(f'{sz:g}pt' for sz in too_small)}, below the "
                f"{floor:g}pt {role} floor; unreadable from the back of "
                f"the room{note}",
                f"format_text(slide={ctx.rec['index']}, shape={s['id']}, "
                f"size_pt={floor:g}) raises every run in the shape (add "
                "paragraph=/start=/end= to target one run)",
                shape_ids=[s["id"]],
                sizes_pt=too_small,
                floor_pt=floor,
                size_sources=sorted({src for _sz, src in sized}),
                inherited_size=bool(sources),
                **({"autofit_scale": round(scale, 4)} if scale != 1.0 else {}),
            )
        )
    return findings


def _check_overflow(ctx: _SlideCtx, opts: dict) -> list[dict]:
    findings = []
    for s in ctx.all:
        if s["hidden"] or etree.QName(s["elem"]).localname != "sp":
            continue
        body = s["elem"].find(qn("p:txBody"))
        if body is None or not shape_text(s["elem"]).strip():
            continue
        bodypr = body.find(qn("a:bodyPr"))
        font_scale, lnspc = 100.0, 0.0
        if bodypr is not None:
            if bodypr.find(qn("a:spAutoFit")) is not None:
                continue  # the frame grows with the text; it cannot overflow
            norm = bodypr.find(qn("a:normAutofit"))
            if norm is not None:
                font_scale = _pct_value(norm.get("fontScale"), 100.0)
                lnspc = _pct_value(norm.get("lnSpcReduction"), 0.0)
        # Both inputs this estimate needs are inheritable, and both used to
        # dead-end: no xfrm returned "cannot be estimated" and an unsized
        # run fell back to a flat 18pt guess. The layout and master know
        # the answers, so the resolved values go in.
        paras = body.findall(qn("a:p"))
        resolved_pt = None
        for p in paras:
            r = p.find(qn("a:r"))
            if r is None:
                continue
            resolved_pt, _src = _inh.resolve_run_size_pt(
                ctx.pkg, ctx.part, s["elem"], p, r.find(qn("a:rPr"))
            )
            if resolved_pt is not None:
                break
        est = _overflow_heuristic(
            s["elem"], body, bodypr, font_scale, lnspc,
            box=s["box"] if s.get("box_inherited") else None,
            size_pt=resolved_pt,
            pkg=ctx.pkg,
            part=ctx.part,
        )
        if not est or est.get("likely_overflow") is not True:
            continue
        ratio = est.get("fill_ratio")
        measured = est.get("method") == "font-metrics"
        # The suppression threshold is an allowance for MODEL error, so it
        # tracks the model. The character-count estimate keeps the wide 1.4x
        # margin it always had; a measurement of the font's own advance
        # widths does not need it, and the 2026-09-21 field run is what
        # proved that a real ~1.1x overflow was being swallowed by it.
        floor = float(
            opts["min_fill_ratio_metrics"] if measured
            else opts["min_fill_ratio"]
        )
        if ratio is not None and ratio < floor:
            continue  # within the model's error margin; do not cry wolf
        if measured:
            how = f"measured against {est.get('font_file')}, no kerning"
        else:
            how = "HEURISTIC estimate with no real font metrics"
        findings.append(
            ctx.finding(
                "overflow",
                "warning",
                f"{_label(s)} likely overflows its frame"
                + (f" (estimated {ratio}x the available height)" if ratio else "")
                + "; " + how,
                f"format_text(slide={ctx.rec['index']}, shape={s['id']}, "
                f"size_pt=...) to shrink the text, or set_shape(slide="
                f"{ctx.rec['index']}, shape={s['id']}, h=...) to grow the "
                "frame; confirm with export_slide_images first",
                shape_ids=[s["id"]],
                fill_ratio=ratio,
                heuristic=not measured,
                method=est.get("method"),
                font_file=est.get("font_file"),
            )
        )
    return findings


#: Placeholder types whose emptiness is a content bug, not furniture.
_TEXT_PH_TYPES = {"title", "ctrTitle", "subTitle", "body", "obj", None}


def _check_empty_placeholder(ctx: _SlideCtx, opts: dict) -> list[dict]:
    findings = []
    for s in ctx.all:
        if s["hidden"] or s["kind"] != "placeholder":
            continue
        ph = _ph(s["elem"])
        ph_type = ph.get("type") if ph is not None else None
        if ph_type not in _TEXT_PH_TYPES:
            continue
        if not txbody_paragraphs(s["elem"]):
            continue  # no text body at all: a picture/content placeholder
        if shape_text(s["elem"]).strip():
            continue
        findings.append(
            ctx.finding(
                "empty_placeholder",
                "warning",
                f"{_label(s)} is an empty {ph_type or 'content'} "
                "placeholder; it renders as invisible dead space (or "
                "prompt text in edit view)",
                f"set_placeholder_text(slide={ctx.rec['index']}, ...) to "
                f"fill it, or delete_shape(slide={ctx.rec['index']}, "
                f"shape={s['id']}) to clear it",
                shape_ids=[s["id"]],
                placeholder_type=ph_type,
            )
        )
    return findings


def _check_missing_title(ctx: _SlideCtx, opts: dict) -> list[dict]:
    empty_title_id = None
    for s in ctx.all:
        if s["kind"] != "placeholder":
            continue
        ph = _ph(s["elem"])
        if ph is None or ph.get("type") not in ("title", "ctrTitle"):
            continue
        if shape_text(s["elem"]).strip():
            return []
        empty_title_id = s["id"]
    # A title the CALLER declared: textbox-styled decks name their title
    # boxes consistently, and saying so is an explicit statement, unlike
    # the position guess this replaced.
    declared = {n.casefold() for n in opts.get("title_shape_names") or []}
    if declared:
        for s in ctx.all:
            if s["hidden"] or (s["name"] or "").casefold() not in declared:
                continue
            if shape_text(s["elem"]).strip():
                return []
    # Text near the top used to count as a de-facto title and silence the
    # finding. That hid the commonest real defect: a slide whose title
    # placeholder was deleted, leaving its first bullet near the top
    # (punch-list #928). Such text is now reported as the likely title
    # instead, and the slide is flagged.
    candidates = [
        s
        for s in ctx.all
        if not s["hidden"]
        and s["box"] is not None
        and etree.QName(s["elem"]).localname == "sp"
        and s["box"][1] < ctx.slide_cy * 0.25
        and shape_text(s["elem"]).strip()
    ]
    if empty_title_id is not None:
        fix = (
            f"set_placeholder_text(slide={ctx.rec['index']}, "
            'placeholder="title", text=...) fills the existing empty title'
        )
    else:
        fix = (
            "this slide has no title placeholder, and apply_layout never "
            "adds one: insert_slide on a layout with a title gives a slide "
            'that has it, set_placeholder_text(placeholder="title", '
            "text=...) fills it, and this slide's content moves there"
        )
    message = (
        "slide has no title text; screen readers and Outline view "
        "identify slides by their title"
    )
    if candidates:
        message += (
            f"; {', '.join(_label(s) for s in candidates)} near the top "
            "may look like a title but is not one to a screen reader"
        )
    return [
        ctx.finding(
            "missing_title",
            "info",
            message,
            fix,
            shape_ids=[empty_title_id] if empty_title_id is not None else [],
            candidate_shape_ids=[s["id"] for s in candidates],
        )
    ]


def _lvl_defrpr(body: etree._Element | None, lvl: int):
    """The a:defRPr of one outline level in a txBody's own lstStyle."""
    if body is None:
        return None
    lst = body.find(qn("a:lstStyle"))
    if lst is None:
        return None
    lvlppr = lst.find(qn(f"a:lvl{min(max(lvl, 0), 8) + 1}pPr"))
    return lvlppr.find(qn("a:defRPr")) if lvlppr is not None else None


def _inherited_defrprs(ctx: "_SlideCtx", elem: etree._Element, lvl: int):
    """[(defRPr, source)] from the layout placeholder, the master
    placeholder, and the master text style, in that order."""
    out = []
    key = _inh.placeholder_key(_inh.placeholder_of(elem))
    if key is None:
        return out
    layout_part = _inh.layout_part_of(ctx.pkg, ctx.part)
    master_part = _inh.master_part_of(ctx.pkg, layout_part)
    for part_twin, label in (
        (_inh.layout_twin(ctx.pkg, layout_part, key), "layout"),
        (_inh.master_twin(ctx.pkg, master_part, key), "master placeholder"),
    ):
        if part_twin is None:
            continue
        defrpr = _lvl_defrpr(part_twin.find(qn("p:txBody")), lvl)
        if defrpr is not None:
            out.append((defrpr, label))
    if master_part and ctx.pkg.has_part(master_part):
        styles = ctx.pkg.root(master_part).find(qn("p:txStyles"))
        if styles is not None:
            node = styles.find(
                qn(_MASTER_TX_STYLE[_inh.family_of(key[0])])
            )
            if node is not None:
                lvlppr = node.find(qn(f"a:lvl{min(max(lvl, 0), 8) + 1}pPr"))
                defrpr = (
                    lvlppr.find(qn("a:defRPr"))
                    if lvlppr is not None else None
                )
                if defrpr is not None:
                    out.append((defrpr, "master"))
    return out


#: Placeholder family -> the master text style backing it.
_MASTER_TX_STYLE = {
    "title": "p:titleStyle",
    "body": "p:bodyStyle",
    "other": "p:otherStyle",
}


def _run_text_color(
    r: etree._Element,
    p: etree._Element,
    body: etree._Element,
    resolver: _ColorResolver,
    ctx: "_SlideCtx | None" = None,
    elem: etree._Element | None = None,
    fallback: tuple[str, str] | None = None,
) -> tuple[str | None, float | None, bool, str]:
    """(hex or None, explicit size pt or None, bold, source) for one run.

    The chain is the whole chain now: run rPr, paragraph defRPr, the
    shape's own lstStyle, the shape's p:style fontRef, the layout
    placeholder, the master placeholder, the master text styles, and
    finally the theme's mapped tx1. It used to stop after the shape's own
    XML and return None for everything else, which silently exempted every
    ordinary bulleted body slide from the contrast check: those runs carry
    no colour of their own, which is exactly why they are worth checking.
    """
    lvl = 0
    ppr = p.find(qn("a:pPr"))
    if ppr is not None and ppr.get("lvl"):
        try:
            lvl = int(ppr.get("lvl"))
        except ValueError:
            lvl = 0
    chain: list[tuple[etree._Element | None, str]] = [
        (r.find(qn("a:rPr")), "run"),
        (ppr.find(qn("a:defRPr")) if ppr is not None else None, "paragraph"),
        (_lvl_defrpr(body, lvl), "shape"),
    ]
    if ctx is not None and elem is not None:
        chain.extend(_inherited_defrprs(ctx, elem, lvl))

    hexval = None
    source = "unresolved"
    size = None
    bold = None
    for props, label in chain:
        if props is None:
            continue
        if hexval is None:
            solid = props.find(qn("a:solidFill"))
            if solid is not None:
                hexval = resolver.resolve(_first_color_child(solid))
                if hexval:
                    source = label
        if size is None and props.get("sz") is not None:
            try:
                size = int(props.get("sz")) / 100.0
            except ValueError:
                pass
        if bold is None and props.get("b") is not None:
            bold = props.get("b") == "1"

    if hexval is None and elem is not None:
        # p:style/fontRef is how an inserted autoshape gets its text
        # colour, and it is why white-on-white happens after a retext.
        fontref = elem.find(f"{qn('p:style')}/{qn('a:fontRef')}")
        if fontref is not None:
            hexval = resolver.resolve(_first_color_child(fontref))
            if hexval:
                source = "style/fontRef"
    if hexval is None and fallback is not None:
        # A caller that CAN reach the governing part (the table style)
        # hands its answer in, so the theme default is not reached and the
        # finding is a reading rather than a guess.
        hexval, source = fallback
    if hexval is None:
        hexval = resolver.scheme_hex("tx1")
        if hexval:
            source = "theme default"
    return hexval, size, bool(bold), source


#: tblStyle parts in ASCENDING precedence, which is exactly the child
#: sequence ECMA-376 Part 1 gives CT_TableStyle (20.1.4.2.26): later in
#: the sequence wins. Confirmed by rendering a deck-local style whose
#: every part declares a different fill and sampling the pixels: a header
#: row beats a column band, a column band beats a row band, first/last
#: column beats both bands, and a corner cell beats everything.
#:
#: The round-2 version of this list was walked FIRST-match, so wholeTbl
#: beat every emphasis part and the resolver reported white-on-white where
#: the renderer paints a navy header. Read it highest-first (see _parts).
_TBL_STYLE_ORDER = (
    "wholeTbl",
    "band1H", "band2H", "band1V", "band2V",
    "lastCol", "firstCol", "lastRow",
    "seCell", "swCell",
    "firstRow",
    "neCell", "nwCell",
)

#: A corner part and the two emphases that switch it on. ECMA and
#: PowerPoint treat a corner as the INTERSECTION of a row emphasis and a
#: column emphasis; LibreOffice paints corners whatever the flags say
#: (measured). _TableStyle refuses to pick a side where they disagree.
_TBL_CORNERS = {
    "nwCell": ("firstRow", "firstCol"),
    "neCell": ("firstRow", "lastCol"),
    "swCell": ("lastRow", "firstCol"),
    "seCell": ("lastRow", "lastCol"),
}

#: Sentinel: the cell's colour depends on which renderer opens the deck,
#: so this resolver declines to state one. It flows to the same
#: unresolved path a missing style does, which downgrades the finding to
#: info rather than asserting a colour nobody can verify statically.
_RENDERER_DEPENDENT = object()


class _TableStyle:
    """One deck-local a:tblStyle, resolved per cell.

    Only styles DEFINED in tableStyles.xml can be read. PowerPoint's
    built-in styles are referenced by GUID and their definitions live
    inside the application, not in the file, so a deck using one has a
    tableStyles.xml carrying nothing but a `def` attribute. That is why
    the unresolved path still has to exist: there is genuinely nothing to
    read, and anything this check said about those colours would be a
    guess.
    """

    def __init__(self, style: etree._Element, flags: dict,
                 resolver: _ColorResolver):
        self._style = style
        self._flags = flags
        self._resolver = resolver

    def _row_emphasis(self, row: int, rows: int) -> str | None:
        if self._flags.get("firstRow") and row == 0:
            return "firstRow"
        if self._flags.get("lastRow") and row == rows - 1:
            return "lastRow"
        return None

    def _col_emphasis(self, col: int, cols: int) -> str | None:
        if self._flags.get("firstCol") and col == 0:
            return "firstCol"
        if self._flags.get("lastCol") and col == cols - 1:
            return "lastCol"
        return None

    def _parts(self, row: int, col: int, rows: int, cols: int) -> list[str]:
        """Which style parts apply to one cell, HIGHEST precedence first.

        Banding skips an emphasised first or last row (and column), which
        is what PowerPoint does and what the render confirms: with a
        header on, the first BODY row is band1, not band2.
        """
        active = ["wholeTbl"]
        row_emph = self._row_emphasis(row, rows)
        col_emph = self._col_emphasis(col, cols)

        if self._flags.get("bandRow") and row_emph is None:
            body = row - (1 if self._flags.get("firstRow") else 0)
            active.append("band1H" if body % 2 == 0 else "band2H")
        if self._flags.get("bandCol") and col_emph is None:
            body = col - (1 if self._flags.get("firstCol") else 0)
            active.append("band1V" if body % 2 == 0 else "band2V")
        if col_emph:
            active.append(col_emph)
        if row_emph:
            active.append(row_emph)
        if row_emph and col_emph:
            for corner, (need_row, need_col) in _TBL_CORNERS.items():
                if need_row == row_emph and need_col == col_emph:
                    active.append(corner)
        # Highest precedence first, so the first part that carries a value
        # is the one that wins. Walking this list the other way is what
        # made wholeTbl beat everything.
        return [p for p in reversed(_TBL_STYLE_ORDER) if p in active]

    def _ambiguous_corner(self, row, col, rows, cols) -> bool:
        """True when this cell sits in a geometric corner whose corner
        part the style DEFINES, but the emphasis flags do not switch on.
        ECMA and PowerPoint leave it alone; LibreOffice paints it. Rather
        than pick a winner, the resolver declines to state a colour."""
        if self._row_emphasis(row, rows) and self._col_emphasis(col, cols):
            return False
        for corner, (need_row, need_col) in _TBL_CORNERS.items():
            at_row = (row == 0) if need_row == "firstRow" else (row == rows - 1)
            at_col = (col == 0) if need_col == "firstCol" else (col == cols - 1)
            if at_row and at_col and self._style.find(qn(f"a:{corner}")) is not None:
                return True
        return False

    def text_color(self, row, col, rows, cols):
        """Hex, None when the style says nothing, or _RENDERER_DEPENDENT."""
        if self._ambiguous_corner(row, col, rows, cols):
            return _RENDERER_DEPENDENT
        for part in self._parts(row, col, rows, cols):
            node = self._style.find(qn(f"a:{part}"))
            if node is None:
                continue
            txs = node.find(qn("a:tcTxStyle"))
            if txs is None:
                continue
            hexval = self._resolver.resolve(_first_color_child(txs))
            if hexval:
                return hexval
        return None

    def fill(self, row, col, rows, cols):
        """Hex, _TRANSPARENT, None, or _RENDERER_DEPENDENT."""
        if self._ambiguous_corner(row, col, rows, cols):
            return _RENDERER_DEPENDENT
        for part in self._parts(row, col, rows, cols):
            node = self._style.find(qn(f"a:{part}"))
            if node is None:
                continue
            fill = node.find(f"{qn('a:tcStyle')}/{qn('a:fill')}")
            if fill is None:
                continue
            solid = fill.find(qn("a:solidFill"))
            if solid is not None:
                hexval = self._resolver.resolve(_first_color_child(solid))
                if hexval:
                    return hexval
            if fill.find(qn("a:noFill")) is not None:
                return _TRANSPARENT
        return None


def _table_style_for(ctx: "_SlideCtx", tbl: etree._Element):
    """The deck-local style backing one table, or None when the style is a
    built-in referenced only by GUID (the common case)."""
    part = "ppt/tableStyles.xml"
    if not ctx.pkg.has_part(part):
        return None
    try:
        lst = ctx.pkg.root(part)
    except (KeyError, PptMcpError):
        return None
    tblpr = tbl.find(qn("a:tblPr"))
    style_id = None
    if tblpr is not None:
        node = tblpr.find(qn("a:tableStyleId"))
        if node is not None and node.text:
            style_id = node.text.strip()
    if style_id is None:
        style_id = lst.get("def")
    if style_id is None:
        return None
    for style in lst.findall(qn("a:tblStyle")):
        if (style.get("styleId") or "").strip() == style_id:
            flags = {
                flag: (tblpr.get(flag) == "1") if tblpr is not None else False
                for flag in (
                    "firstRow", "lastRow", "firstCol", "lastCol",
                    "bandRow", "bandCol",
                )
            }
            return _TableStyle(style, flags, ctx.resolver())
    return None


def _cell_fill_hex(
    tc: etree._Element, resolver: _ColorResolver
) -> tuple[str | None, str | None]:
    """(hex | _TRANSPARENT | None, reason) for one table cell's own fill.
    A cell with no a:tcPr fill takes its colour from the table STYLE part,
    which this read does not resolve, so it comes back as a reason."""
    tcpr = tc.find(qn("a:tcPr"))
    if tcpr is None:
        return None, "cell fill comes from the table style part"
    for child in tcpr:
        local = etree.QName(child).localname
        if local == "solidFill":
            return resolver.resolve(_first_color_child(child)), None
        if local == "noFill":
            return _TRANSPARENT, None
        if local in ("blipFill", "pattFill", "gradFill"):
            return None, f"{local} cell fill"
    return None, "cell fill comes from the table style part"


def _worst_run(
    ctx: _SlideCtx,
    elem: etree._Element,
    body: etree._Element,
    paragraphs: list,
    bg_hex: str,
    resolver: _ColorResolver,
    min_ratio: float,
    large_min: float,
    fallback: tuple[str, str] | None = None,
) -> dict | None:
    worst: dict | None = None
    for p in paragraphs:
        for r in p.findall(qn("a:r")):
            t = r.find(qn("a:t"))
            if t is None or not (t.text or "").strip():
                continue
            fg_hex, size, bold, source = _run_text_color(
                r, p, body, resolver, ctx, elem, fallback
            )
            if fg_hex is None:
                continue
            ratio = contrast_ratio(fg_hex, bg_hex)
            large = (size is not None) and (
                size >= 18.0 or (bold and size >= 14.0)
            )
            threshold = large_min if large else min_ratio
            if ratio >= threshold:
                continue
            if worst is None or ratio < worst["ratio"]:
                worst = {
                    "ratio": ratio,
                    "fg": fg_hex,
                    "threshold": threshold,
                    "source": source,
                    "sample": (t.text or "").strip()[:40],
                }
    return worst


#: Colour sources that are a FALLBACK, not a reading. The governing part
#: is out of reach (a built-in table style, a placeholder chain that runs
#: dry), so the colour is the theme's default rather than the slide's.
_GUESSED_SOURCES = {"theme default", "unresolved"}


def _contrast_finding(ctx, s, worst, bg_hex) -> dict:
    suggested = "000000" if _rel_luminance(bg_hex) > 0.35 else "FFFFFF"
    severity = "error" if worst["ratio"] < 2.0 else "warning"
    where = (
        "" if worst["source"] == "run"
        else f", colour from the {worst['source']}"
    )
    if worst["source"] in _GUESSED_SOURCES:
        # A check must never fabricate a gate-severity failure out of a
        # colour it had to guess. An ordinary dark-header table built with
        # this server's own tools used to come back as an ERROR reading
        # "text #000000 on #1F3864" when the header text is white: the
        # fill was explicit, the text colour lived in the built-in table
        # style, and the fallback to tx1 invented the black. "Cannot know"
        # is a result, and it is an info.
        return ctx.finding(
            "contrast",
            "info",
            f"{_label(s)}: text on #{bg_hex} could NOT be checked, because "
            f"its colour is not stated anywhere this read can reach "
            f"({worst['sample']!r}); the theme default would give "
            f"{round(worst['ratio'], 2)}:1, which is a guess, not a finding",
            f"export_slide_image(slide={ctx.rec['index']}) and look, or "
            f"format_text(slide={ctx.rec['index']}, shape={s['id']}, "
            f'color="{suggested}") to state it',
            shape_ids=[s["id"]],
            reason="unresolved_colour_source",
            fill_color=bg_hex,
            color_source=worst["source"],
        )
    return ctx.finding(
        "contrast",
        severity,
        f"{_label(s)}: text #{worst['fg']} on #{bg_hex} has contrast "
        f"{round(worst['ratio'], 2)}:1, below the {worst['threshold']}:1 "
        f"target ({worst['sample']!r}){where}; approximated from resolved "
        "solid colors, not a render",
        f"format_text(slide={ctx.rec['index']}, shape={s['id']}, "
        f'color="{suggested}") fixes the text, or set_shape(slide='
        f"{ctx.rec['index']}, shape={s['id']}, fill=...) changes "
        "the background",
        shape_ids=[s["id"]],
        ratio=round(worst["ratio"], 2),
        text_color=worst["fg"],
        fill_color=bg_hex,
        color_source=worst["source"],
    )


def _check_contrast(ctx: _SlideCtx, opts: dict) -> list[dict]:
    findings = []
    min_ratio = float(opts["min_ratio"])
    large_min = float(opts["large_min_ratio"])
    resolver = ctx.resolver()
    for s in ctx.all:
        if s["hidden"]:
            continue
        if s["kind"] == "table":
            findings.extend(
                _check_table_contrast(ctx, s, resolver, min_ratio, large_min)
            )
            continue
        if etree.QName(s["elem"]).localname != "sp":
            continue
        if not shape_text(s["elem"]).strip():
            continue
        # a:noFill in spPr beats the p:style fillRef: the renderer honours
        # the explicit refusal to paint, and a check that read the fillRef
        # anyway invented a colour the slide does not show.
        bg_hex, skip_reason = _own_fill_hex(s["elem"], resolver)
        if bg_hex == _TRANSPARENT:
            bg_hex, skip_reason = _backdrop_hex(s, ctx, resolver)
        elif bg_hex is not None:
            # Alpha-aware: a translucent solid fill renders as a wash over
            # whatever is behind it, so judge the COMPOSITED color, not the
            # full-strength one (tinted-fill false positive).
            alpha = _own_fill_alpha(s["elem"])
            if alpha is not None and alpha < 0.995:
                behind, behind_skip = _backdrop_hex(s, ctx, resolver)
                if behind is None:
                    skip_reason = behind_skip or (
                        "translucent fill over an unresolvable backdrop"
                    )
                else:
                    bg_hex = _blend_hex(bg_hex, behind, alpha)
        if skip_reason or bg_hex is None:
            # An unresolvable backdrop is a RESULT, not a silent pass: the
            # text may be invisible and this check cannot tell.
            findings.append(
                ctx.finding(
                    "contrast",
                    "info",
                    f"{_label(s)} carries text over an unresolvable "
                    f"backdrop ({skip_reason or 'no colour resolved'}); "
                    "its contrast was NOT checked",
                    f"export_slide_image(slide={ctx.rec['index']}) and look, "
                    f"or set_shape(slide={ctx.rec['index']}, "
                    f"shape={s['id']}, fill=...) to give it a known fill",
                    shape_ids=[s["id"]],
                    unresolvable_backdrop=skip_reason or "no colour resolved",
                )
            )
            continue
        body = s["elem"].find(qn("p:txBody"))
        if body is None:
            continue
        worst = _worst_run(
            ctx, s["elem"], body, txbody_paragraphs(s["elem"]), bg_hex,
            resolver, min_ratio, large_min,
        )
        if worst is not None:
            findings.append(_contrast_finding(ctx, s, worst, bg_hex))
    return findings


def _check_table_contrast(
    ctx: _SlideCtx,
    s: dict,
    resolver: _ColorResolver,
    min_ratio: float,
    large_min: float,
) -> list[dict]:
    """Cell text against the cell's own fill. One finding for the worst
    cell, plus an info per reason a cell was declined: a built-in style
    whose definition is not in the file, and a corner cell whose emphasis
    flags renderers disagree about. The two were counted together and
    both reported as the first, which told a reader to go looking for a
    missing style part that was not the problem."""
    tbl = table_element(s["elem"])
    if tbl is None:
        return []
    # A style DEFINED in tableStyles.xml is readable, so its cell colours
    # are resolved rather than guessed. A built-in style is referenced by
    # GUID only and its definition is not in the file, which is what the
    # unresolved path below is for.
    style = _table_style_for(ctx, tbl)
    rows = tbl.findall(qn("a:tr"))
    row_count = len(rows)
    col_count = len(tbl.findall(f"{qn('a:tblGrid')}/{qn('a:gridCol')}"))

    worst: dict | None = None
    worst_bg = None
    unresolved = 0
    renderer_dependent = 0
    for row_i, tr in enumerate(rows):
        for col_i, tc in enumerate(tr.findall(qn("a:tc"))):
            body = tc.find(qn("a:txBody"))
            if body is None:
                continue
            paragraphs = body.findall(qn("a:p"))
            if not any(
                (t.text or "").strip()
                for p in paragraphs for t in p.iter(qn("a:t"))
            ):
                continue
            corner = False
            bg_hex, reason = _cell_fill_hex(tc, resolver)
            if reason and style is not None:
                styled = style.fill(row_i, col_i, row_count, col_count)
                if styled is _RENDERER_DEPENDENT:
                    reason = "corner cell renderers disagree about"
                    corner = True
                elif styled is not None:
                    bg_hex, reason = styled, None
            if bg_hex == _TRANSPARENT:
                # Whatever the backdrop pass decides, it decides for its
                # own reasons; the corner one no longer applies.
                bg_hex, reason = _backdrop_hex(s, ctx, resolver)
                corner = False
            if reason or bg_hex is None:
                if corner:
                    renderer_dependent += 1
                else:
                    unresolved += 1
                continue
            fallback = None
            if style is not None:
                styled_text = style.text_color(
                    row_i, col_i, row_count, col_count
                )
                if styled_text is _RENDERER_DEPENDENT:
                    renderer_dependent += 1
                    continue
                if styled_text:
                    fallback = (styled_text, "table style")
            cell_worst = _worst_run(
                ctx, s["elem"], body, paragraphs, bg_hex, resolver,
                min_ratio, large_min, fallback,
            )
            if cell_worst and (
                worst is None or cell_worst["ratio"] < worst["ratio"]
            ):
                worst, worst_bg = cell_worst, bg_hex
    out = []
    if worst is not None:
        out.append(_contrast_finding(ctx, s, worst, worst_bg))
    remedy = (
        f"export_slide_image(slide={ctx.rec['index']}) and look, or "
        f"format_table_cells(slide={ctx.rec['index']}, "
        f"table={s['id']}, ..., fill=...) to set explicit cell fills"
    )
    # One finding per REASON. A declined corner used to be counted with
    # the built-in-style cells and reported as one of them, so the
    # message named a missing style part that was not the problem.
    for count, what in (
        (unresolved, "cell(s) whose fill comes from the table style part"),
        (renderer_dependent,
         "corner cell(s) whose emphasis flags renderers disagree about"),
    ):
        if not count:
            continue
        out.append(
            ctx.finding(
                "contrast",
                "info",
                f"{_label(s)} has {count} {what}, which this check does "
                "not resolve; their contrast was NOT checked",
                remedy,
                shape_ids=[s["id"]],
                unresolved_cells=count,
            )
        )
    return out


def _check_diagram_glue(ctx: _SlideCtx, opts: dict) -> list[dict]:
    if ctx.sp_tree is None:
        return []
    findings = []
    tol = float(opts["touch_tolerance_in"]) * EMU_PER_INCH
    targets = [
        s
        for s in ctx.all
        if s["kind"] != "connector" and not s["hidden"] and s["box"] is not None
    ]
    for cxnsp, chain in _iter_connectors(ctx.sp_tree):
        cid = _shape_id(cxnsp)
        cnvcxn = cxnsp.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}")
        st = cnvcxn.find(qn("a:stCxn")) if cnvcxn is not None else None
        en = cnvcxn.find(qn("a:endCxn")) if cnvcxn is not None else None
        try:
            p_start, p_end = _connector_endpoints_slide(cxnsp, chain)
        except (UnsupportedStructure, PptMcpError, ValueError):
            continue
        loose: list[tuple[str, dict]] = []
        for glued, point, end_name in ((st, p_start, "start"), (en, p_end, "end")):
            if glued is not None:
                continue
            hit = _touched_shape(point, targets, tol)
            if hit is not None:
                loose.append((end_name, hit))
        if not loose:
            continue
        touched = ", ".join(
            f"{end_name} touches {_label(hit)}" for end_name, hit in loose
        )
        glued_count = (1 if st is not None else 0) + (1 if en is not None else 0)
        kwargs = []
        for end_name, hit in loose:
            kwargs.append(f"{end_name}_shape={hit['id']}")
        if st is not None:
            kwargs.insert(0, "start_shape=<current glued id>")
        if en is not None:
            kwargs.append("end_shape=<current glued id>")
        findings.append(
            ctx.finding(
                "diagram_glue",
                "warning",
                f"connector {cid} has {glued_count} of 2 ends glued but "
                f"{touched}; the loose end(s) LOOK attached and will be "
                "left behind the moment the shape moves",
                f"delete_shape(slide={ctx.rec['index']}, shape={cid}) and "
                f"re-create with insert_connector(slide="
                f"{ctx.rec['index']}, {', '.join(kwargs)}) so both ends "
                "are glued and follow their shapes",
                shape_ids=[cid] + [hit["id"] for _n, hit in loose],
                connector_id=cid,
                glued_ends=glued_count,
                loose_ends=[
                    {"end": end_name, "touches_shape": hit["id"]}
                    for end_name, hit in loose
                ],
            )
        )
    return findings


def _touched_shape(point, targets: list[dict], tol: float) -> dict | None:
    """The FRONTMOST shape whose (tolerance-expanded) bbox contains the
    point, preferring the smallest such shape (a point on a node inside a
    panel glues to the node, not the panel)."""
    px, py = point
    hits = []
    for s in targets:
        x, y, cx, cy = s["box"]
        if x - tol <= px <= x + cx + tol and y - tol <= py <= y + cy + tol:
            hits.append(s)
    if not hits:
        return None
    return min(hits, key=lambda s: s["box"][2] * s["box"][3])


_CHECK_FNS = {
    "overlap": _check_overlap,
    "off_slide": _check_off_slide,
    "tiny_text": _check_tiny_text,
    "overflow": _check_overflow,
    "empty_placeholder": _check_empty_placeholder,
    "missing_title": _check_missing_title,
    "contrast": _check_contrast,
    "diagram_glue": _check_diagram_glue,
}

_SEV_RANK = {"error": 0, "warning": 1, "info": 2}


# ================================================================ public API


def check_layout(pkg: PptxPackage, slide=None, checks=None, *,
                 limit=None, offset: int = 0) -> dict:
    """Run the design guardrail battery over `slide` (a selector, a list of
    selectors, or None for the whole deck). `checks` selects and tunes the
    battery: None = everything with defaults; entries are check names
    ("overlap") or option dicts ({"check": "tiny_text", "body_min_pt": 12}).

    Returns per-check findings with shape ids, severities (error > warning
    > info), and a fix hint naming the exact tool call that repairs the
    problem, plus per-check caveats stating what each heuristic can and
    cannot see. Read-only; nothing is modified."""
    plan = _normalize_checks(checks)
    recs = slides_in_scope(pkg, slide)
    findings: list[dict] = []
    stats: dict[str, dict[str, int]] = {}
    for rec in recs:
        ctx = _SlideCtx(pkg, rec)
        for name, opts in plan:
            findings.extend(_CHECK_FNS[name](ctx, opts))
        for check, counters in ctx.stats.items():
            bucket = stats.setdefault(check, {})
            for key, value in counters.items():
                bucket[key] = bucket.get(key, 0) + value
    findings.sort(
        key=lambda f: (_SEV_RANK.get(f["severity"], 3), f["slide_index"])
    )
    summary: dict[str, int] = {}
    for f in findings:
        summary[f["check"]] = summary.get(f["check"], 0) + 1
    # The counts and the per-check summary are the sweep's whole value and
    # stay complete; only the finding list pages. On the heaviest corpus
    # deck this call used to return 1.4 million characters.
    header = {
        "slides_checked": len(recs),
        "checks_run": [name for name, _o in plan],
        "finding_count": len(findings),
        "by_severity": {
            sev: sum(1 for f in findings if f["severity"] == sev)
            for sev in ("error", "warning", "info")
        },
        "by_check": summary,
    }
    tail = {
        "checks": [
            {"check": name, "caveat": CHECKS[name][1], **stats.get(name, {})}
            for name, _o in plan
        ],
        "caveats": {name: CHECKS[name][1] for name, _o in plan},
        "note": (
            "static-XML heuristics, not a renderer; for final visual "
            "verification use export_slide_images and look at the PNGs"
        ),
    }
    kept, page = _budget.page_items(
        findings, limit=limit, offset=offset,
        # the caveats and the note are measured, not guessed: they run to
        # nearly a thousand characters and a guess put the answer over.
        overhead=_budget.overhead_of({**header, **tail}),
        unit="findings",
        narrow_hint="narrow with slide=<index> or checks=[<one check>]",
        shrink=_budget.shrink_record,
    )
    result = {**header, "findings": kept, **tail}
    if page is not None:
        result["page"] = page
    return result
