"""Theme editing: color scheme, font scheme, and deck-to-deck brand transfer.

Why this matters: everything the graphics engine and the template ship as
schemeClr (accent1..accent6, tx/bg tokens) re-resolves against the theme at
render time. Editing the 12 clrScheme slots therefore recolors every
theme-linked fill, line, chart series, and diagram in the deck in one write,
which is the whole point of building diagrams theme-native.

Contract (all ops modules): functions take the open PptxPackage first
(extract_brand also accepts a path, it is read-only), mutate only the
in-memory package via mark_dirty on the theme part(s), and return plain
dicts. Saving (atomic, validated) is the server layer's job.

Structural rules baked in:
- Each slide master relates to ONE theme part; several masters may share
  one. set_theme_colors/set_theme_fonts edit the SELECTED master's theme;
  apply_brand walks every master and dedupes shared theme parts.
- a:clrScheme children are schema-ordered (dk1, lt1, dk2, lt2,
  accent1..accent6, hlink, folHlink); slots are edited in place, never
  reordered. A slot is rewritten as a plain a:srgbClr (replacing sysClr
  where present, exactly what PowerPoint does when a user customizes a
  theme color).
- a:majorFont/a:minorFont carry a:latin, a:ea, a:cs in that order; the ea
  (East Asian) slot is a first-class parameter because CJK decks resolve
  +mj-ea/+mn-ea through it and an empty ea typeface silently falls back.
- Explicit srgbClr fills on shapes do NOT follow theme edits; that is why
  extract_brand reports them separately instead of pretending the theme
  covers them.
"""

from __future__ import annotations

import os
from collections import Counter

from lxml import etree

from ..core.errors import OpError, PptMcpError, UnsupportedStructure
from ..core.package import PptxPackage, qn
from .design import COLOR_SLOTS, _resolve_master, _theme_part_of, get_theme

_FONT_KEYS = ("latin", "ea", "cs")
_FONT_TAGS = {"latin": "a:latin", "ea": "a:ea", "cs": "a:cs"}

#: Longest typeface a theme font slot takes. A Windows face name holds at
#: most 31 characters (LOGFONT's LF_FACESIZE of 32 counts the terminator)
#: and Office stores that face name, so no installed font has a longer one.
#: A longer value only bloats the theme and every answer that echoes it.
MAX_TYPEFACE_CHARS = 31


def _check_typeface(typeface: str, label: str) -> str:
    """Refuse a typeface no installed font could have, before any write."""
    if len(typeface) > MAX_TYPEFACE_CHARS:
        raise OpError(
            "BAD_ARGS",
            f"{label} typeface is {len(typeface):,} characters, but a font name has at "
            f"most {MAX_TYPEFACE_CHARS} (the Windows face-name limit), so no installed "
            "font has it. Nothing was changed.",
            details={
                "field": label,
                "chars": len(typeface),
                "max_chars": MAX_TYPEFACE_CHARS,
            },
        )
    return typeface


_RUN_FONT_TAGS = ("a:latin", "a:ea", "a:cs", "a:sym")
_BRAND_FONT_KEYS = {
    f"{family}_{script}": (family, script)
    for family in ("major", "minor")
    for script in _FONT_KEYS
}
_COLOR_MAP_KEYS = (
    "bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3",
    "accent4", "accent5", "accent6", "hlink", "folHlink",
)
_COLOR_MAP_VALUES = set(COLOR_SLOTS)


# ---------------------------------------------------------------- helpers


def _norm_hex(value, slot: str) -> str:
    """'#1F4E79' / '1f4e79' / '#abc' -> 'RRGGBB' uppercase, or refuse."""
    if not isinstance(value, str) or not value:
        raise PptMcpError(
            f"color for slot {slot!r} must be an RRGGBB hex string, got "
            f"{value!r}"
        )
    raw = value[1:] if value.startswith("#") else value
    if len(raw) == 3 and all(c in "0123456789abcdefABCDEF" for c in raw):
        raw = "".join(c * 2 for c in raw)
    if len(raw) != 6 or not all(c in "0123456789abcdefABCDEF" for c in raw):
        raise PptMcpError(
            f"invalid hex color {value!r} for slot {slot!r}: use RRGGBB "
            "(no alpha channel)"
        )
    return raw.upper()


def _clr_scheme(pkg: PptxPackage, theme_part: str) -> etree._Element:
    if not pkg.has_part(theme_part):
        raise UnsupportedStructure(
            f"theme part {theme_part} is missing from the package"
        )
    scheme = pkg.root(theme_part).find(
        f"{qn('a:themeElements')}/{qn('a:clrScheme')}"
    )
    if scheme is None:
        raise UnsupportedStructure(f"{theme_part} has no a:clrScheme")
    return scheme


def _font_scheme(pkg: PptxPackage, theme_part: str) -> etree._Element:
    if not pkg.has_part(theme_part):
        raise UnsupportedStructure(
            f"theme part {theme_part} is missing from the package"
        )
    scheme = pkg.root(theme_part).find(
        f"{qn('a:themeElements')}/{qn('a:fontScheme')}"
    )
    if scheme is None:
        raise UnsupportedStructure(f"{theme_part} has no a:fontScheme")
    return scheme


def _write_slot(scheme: etree._Element, slot: str, hexval: str) -> None:
    """Rewrite one clrScheme slot's content as a plain a:srgbClr."""
    el = scheme.find(qn(f"a:{slot}"))
    if el is None:
        raise UnsupportedStructure(
            f"clrScheme has no a:{slot} slot; refusing to guess where the "
            "missing slot belongs in schema order"
        )
    for child in list(el):
        el.remove(child)
    srgb = etree.SubElement(el, qn("a:srgbClr"))
    srgb.set("val", hexval)


def _set_typeface(font_el: etree._Element, key: str, typeface: str) -> None:
    """Set a:latin / a:ea / a:cs @typeface, creating the child in schema
    order (latin, ea, cs come first, before any a:font entries)."""
    node = font_el.find(qn(_FONT_TAGS[key]))
    if node is None:
        node = etree.Element(qn(_FONT_TAGS[key]))
        rank = _FONT_KEYS.index(key)
        inserted = False
        for i, child in enumerate(font_el):
            local = etree.QName(child).localname
            if local not in _FONT_KEYS or _FONT_KEYS.index(local) > rank:
                font_el.insert(i, node)
                inserted = True
                break
        if not inserted:
            font_el.append(node)
    node.set("typeface", typeface)


def _norm_font_spec(spec, name: str) -> dict[str, str]:
    """A font parameter: a bare typeface string means {"latin": s}; a dict
    may set any of latin/ea/cs."""
    if spec is None:
        return {}
    if isinstance(spec, str):
        if not spec.strip():
            raise PptMcpError(f"{name} typeface must be a non-empty string")
        return {"latin": _check_typeface(spec.strip(), name)}
    if isinstance(spec, dict):
        unknown = sorted(set(spec) - set(_FONT_KEYS))
        if unknown:
            raise PptMcpError(
                f"unknown {name} font key(s): {', '.join(unknown)}; valid: "
                f"latin, ea, cs"
            )
        out = {}
        for k, v in spec.items():
            if not isinstance(v, str) or not v.strip():
                raise PptMcpError(
                    f"{name}.{k} typeface must be a non-empty string, got "
                    f"{v!r}"
                )
            out[k] = _check_typeface(v.strip(), f"{name}.{k}")
        if not out:
            raise PptMcpError(f"{name} font dict is empty; nothing to set")
        return out
    raise PptMcpError(
        f"{name} must be a typeface string or a dict of latin/ea/cs, got "
        f"{type(spec).__name__}"
    )


def _explicit_font_runs_in_master_scope(
    pkg: PptxPackage, master_part: str
) -> dict:
    """Visible slide runs with literal fonts under ``master_part``.

    Theme references (for example ``+mn-lt``) continue to resolve through
    the font scheme and are therefore not overrides. A run is counted once
    even when it declares literal typefaces for more than one script.
    """
    from .masters import _layout_of_slide
    from .read import _layouts_of_master, slide_table

    layout_parts = set(_layouts_of_master(pkg, master_part))
    scoped_slides = []
    details = []
    total = 0

    for rec in slide_table(pkg):
        if _layout_of_slide(pkg, rec["part"]) not in layout_parts:
            continue
        scoped_slides.append(rec)
        run_count = 0
        typeface_runs: Counter[str] = Counter()
        root = pkg.root(rec["part"])
        for run in root.iter(qn("a:r"), qn("a:fld")):
            text = run.find(qn("a:t"))
            if text is None or not text.text:
                continue
            rpr = run.find(qn("a:rPr"))
            if rpr is None:
                continue
            literal_faces = set()
            for tag in _RUN_FONT_TAGS:
                face = rpr.find(qn(tag))
                typeface = (
                    face.get("typeface", "").strip()
                    if face is not None
                    else ""
                )
                if typeface and not typeface.startswith("+"):
                    literal_faces.add(typeface)
            if not literal_faces:
                continue
            run_count += 1
            for typeface in literal_faces:
                typeface_runs[typeface] += 1

        if run_count:
            total += run_count
            details.append(
                {
                    "slide_index": rec["index"],
                    "slide_id": rec["slide_id"],
                    "part": rec["part"],
                    "explicit_font_run_count": run_count,
                    "typefaces": [
                        {"typeface": face, "run_count": typeface_runs[face]}
                        for face in sorted(typeface_runs)
                    ],
                }
            )

    return {
        "affected_slides": {
            "count": len(scoped_slides),
            "slide_ids": [rec["slide_id"] for rec in scoped_slides],
        },
        "explicit_font_run_count": total,
        "explicit_font_affected_slides": {
            "count": len(details),
            "slide_ids": [rec["slide_id"] for rec in details],
            "details": details,
        },
    }


def _normalize_color_map(mapping: dict) -> dict[str, str]:
    if not isinstance(mapping, dict) or not mapping:
        raise PptMcpError(
            "color_map must be a non-empty dict of semantic token to theme "
            f"slot; keys: {', '.join(_COLOR_MAP_KEYS)}"
        )
    unknown = sorted(set(mapping) - set(_COLOR_MAP_KEYS))
    if unknown:
        raise PptMcpError(
            f"unknown color_map key(s): {', '.join(unknown)}; valid: "
            + ", ".join(_COLOR_MAP_KEYS)
        )
    out = {}
    for key, value in mapping.items():
        if value not in _COLOR_MAP_VALUES:
            raise PptMcpError(
                f"color_map.{key} must name a theme slot; got {value!r}; "
                f"valid: {', '.join(COLOR_SLOTS)}"
            )
        out[key] = value
    return out


def _master_color_map(pkg: PptxPackage, master_part: str) -> dict[str, str]:
    node = pkg.root(master_part).find(qn("p:clrMap"))
    if node is None:
        raise UnsupportedStructure(f"{master_part} has no p:clrMap")
    return {key: node.get(key, "") for key in _COLOR_MAP_KEYS}


def _layout_master(pkg: PptxPackage, layout_part: str) -> str:
    from ..core.package import resolve_target
    from .masters import RT_SLIDE_MASTER

    for rel in pkg.rels_for(layout_part).getroot():
        if rel.get("Type") == RT_SLIDE_MASTER and rel.get("TargetMode") != "External":
            return resolve_target(layout_part, rel.get("Target", ""))
    raise UnsupportedStructure(f"{layout_part} has no slideMaster relationship")


def _effective_layout_color_map(
    pkg: PptxPackage, layout_part: str
) -> dict[str, str]:
    root = pkg.root(layout_part)
    override = root.find(f"{qn('p:clrMapOvr')}/{qn('a:overrideClrMapping')}")
    if override is not None:
        return {key: override.get(key, "") for key in _COLOR_MAP_KEYS}
    return _master_color_map(pkg, _layout_master(pkg, layout_part))


def _set_color_map(pkg, *, master, layout, mapping):
    from . import masters as _masters
    from .slides import _resolve_layout

    if layout is not None and master is not None:
        raise PptMcpError("pass either master= or layout= for color_map, not both")
    if layout is None:
        master_part, _all = _resolve_master(pkg, master)
        normalized = _normalize_color_map(mapping)
        node = pkg.root(master_part).find(qn("p:clrMap"))
        if node is None:
            raise UnsupportedStructure(f"{master_part} has no p:clrMap")
        final = _master_color_map(pkg, master_part)
        final.update(normalized)
        for key, value in final.items():
            node.set(key, value)
        pkg.mark_dirty(master_part)
        return {
            "scope": "master", "part": master_part,
            "color_map": final,
            "affected_slides": _masters._affected(pkg, "master", master_part),
        }

    layout_part = _resolve_layout(pkg, layout)
    root = pkg.root(layout_part)
    previous_effective = _effective_layout_color_map(pkg, layout_part)
    ovr = root.find(qn("p:clrMapOvr"))
    if ovr is None:
        ovr = etree.Element(qn("p:clrMapOvr"))
        csld = root.find(qn("p:cSld"))
        if csld is None:
            raise UnsupportedStructure(f"{layout_part} has no p:cSld")
        csld.addnext(ovr)
    for child in list(ovr):
        ovr.remove(child)
    if mapping == "inherit":
        etree.SubElement(ovr, qn("a:masterClrMapping"))
        final = _master_color_map(pkg, _layout_master(pkg, layout_part))
        inherited = True
    else:
        normalized = _normalize_color_map(mapping)
        final = previous_effective
        final.update(normalized)
        node = etree.SubElement(ovr, qn("a:overrideClrMapping"))
        for key, value in final.items():
            node.set(key, value)
        inherited = False
    pkg.mark_dirty(layout_part)
    return {
        "scope": "layout", "part": layout_part, "inherited": inherited,
        "color_map": final,
        "affected_slides": _masters._affected(pkg, "layout", layout_part),
    }


# ================================================================ public API


def set_theme_colors(
    pkg: PptxPackage, master=None, colors: dict | None = None,
    *, color_map=None, layout=None,
) -> dict:
    """Set any subset of the 12 theme color slots (dk1, lt1, dk2, lt2,
    accent1..accent6, hlink, folHlink) to RRGGBB hex values on one master's
    theme (master: None = first master, 0-based index, or name). Every
    schemeClr-linked fill, line, chart, and diagram in the deck re-resolves
    against the new values; explicit srgbClr fills do not move (see
    extract_brand). sysClr slots are rewritten as srgbClr, matching
    PowerPoint's own customize behavior."""
    if colors is None and color_map is None:
        raise PptMcpError(
            "nothing to do: pass colors= and/or color_map="
        )
    if layout is not None and colors is not None:
        raise PptMcpError(
            "layout= applies only to color_map; theme palette colors belong "
            "to a master"
        )
    map_result = None
    if color_map is not None:
        map_result = _set_color_map(
            pkg, master=master, layout=layout, mapping=color_map
        )
    if colors is None:
        return map_result
    if not isinstance(colors, dict) or not colors:
        raise PptMcpError(
            f"colors must be a non-empty dict of slot -> RRGGBB hex; valid "
            f"slots: {', '.join(COLOR_SLOTS)}"
        )
    unknown = sorted(set(colors) - set(COLOR_SLOTS))
    if unknown:
        raise PptMcpError(
            f"unknown color slot(s): {', '.join(unknown)}; valid slots: "
            f"{', '.join(COLOR_SLOTS)}"
        )
    normalized = {slot: _norm_hex(val, slot) for slot, val in colors.items()}

    master_part, _all = _resolve_master(pkg, master)
    theme_part = _theme_part_of(pkg, master_part)
    scheme = _clr_scheme(pkg, theme_part)
    for slot, hexval in normalized.items():
        _write_slot(scheme, slot, hexval)
    pkg.mark_dirty(theme_part)
    result = {
        "master": master_part,
        "theme_part": theme_part,
        "set": normalized,
        "colors": get_theme(pkg, master)["colors"],
    }
    if map_result is not None:
        result["color_map_set"] = map_result
    return result


def set_theme_fonts(
    pkg: PptxPackage,
    master=None,
    major=None,
    minor=None,
    ea: str | None = None,
) -> dict:
    """Set the theme font scheme on one master's theme. `major` (headings)
    and `minor` (body) each take a typeface string (sets the latin slot) or
    a dict with any of {"latin", "ea", "cs"}. `ea` is a convenience: one
    East Asian typeface applied to BOTH major and minor (CJK decks resolve
    +mj-ea/+mn-ea through these slots, and an empty ea typeface silently
    falls back at render time). At least one parameter is required."""
    major_spec = _norm_font_spec(major, "major")
    minor_spec = _norm_font_spec(minor, "minor")
    if ea is not None:
        if not isinstance(ea, str) or not ea.strip():
            raise PptMcpError("ea typeface must be a non-empty string")
        face = _check_typeface(ea.strip(), "ea")
        major_spec.setdefault("ea", face)
        minor_spec.setdefault("ea", face)
    if not major_spec and not minor_spec:
        raise PptMcpError(
            "nothing to do: pass major=, minor= (typeface string or "
            '{"latin"/"ea"/"cs": ...} dict), and/or ea='
        )

    master_part, _all = _resolve_master(pkg, master)
    override_report = _explicit_font_runs_in_master_scope(pkg, master_part)
    theme_part = _theme_part_of(pkg, master_part)
    scheme = _font_scheme(pkg, theme_part)
    written: dict[str, dict] = {}
    for slot_tag, spec, label in (
        ("a:majorFont", major_spec, "major"),
        ("a:minorFont", minor_spec, "minor"),
    ):
        if not spec:
            continue
        font_el = scheme.find(qn(slot_tag))
        if font_el is None:
            raise UnsupportedStructure(f"{theme_part} has no {slot_tag}")
        for key, typeface in spec.items():
            _set_typeface(font_el, key, typeface)
        written[label] = spec
    pkg.mark_dirty(theme_part)
    result = {
        "master": master_part,
        "theme_part": theme_part,
        "set": written,
        "fonts": get_theme(pkg, master)["fonts"],
        **override_report,
    }
    if result["explicit_font_run_count"]:
        result["warnings"] = [
            "text with its own font set directly kept that font: "
            "explicit_font_run_count counts those runs and "
            "explicit_font_affected_slides lists the slides they are on"
        ]
    return result


def extract_brand(pkg_or_path, top_fills: int = 8) -> dict:
    """Read a deck's effective palette: the first master's theme colors
    (12 slots) and fonts, PLUS the most-used explicit srgbClr solid-fill
    colors across the slides with usage counts. The explicit list is the
    honest half: shapes filled with literal hex do NOT follow theme edits,
    so a brand transfer that only copies the theme misses them. Read-only;
    accepts an open PptxPackage or a file path."""
    if isinstance(pkg_or_path, (str, os.PathLike)):
        pkg = PptxPackage(pkg_or_path)
    elif isinstance(pkg_or_path, PptxPackage):
        pkg = pkg_or_path
    else:
        raise PptMcpError(
            "extract_brand takes a .pptx path or an open package, got "
            f"{type(pkg_or_path).__name__}"
        )
    if not isinstance(top_fills, int) or isinstance(top_fills, bool) or top_fills < 0:
        raise PptMcpError(f"top_fills must be a non-negative int, got {top_fills!r}")

    theme = get_theme(pkg)
    counts: Counter[str] = Counter()
    slide_parts = pkg.slide_parts()
    for part in slide_parts:
        root = pkg.root(part)
        for solid in root.iter(qn("a:solidFill")):
            srgb = solid.find(qn("a:srgbClr"))
            if srgb is not None and srgb.get("val"):
                counts[srgb.get("val").upper()] += 1
    return {
        "source": pkg.path.name,
        "theme_name": theme["name"],
        "colors": theme["colors"],
        "fonts": theme["fonts"],
        "explicit_fills": [
            {"hex": hexval, "count": n}
            for hexval, n in counts.most_common(top_fills)
        ],
        "explicit_fill_total": sum(counts.values()),
        "scope": f"{len(slide_parts)} slide part(s); explicit-fill counts "
        "cover a:solidFill/a:srgbClr on slides (shape, line, and text "
        "fills alike), not layouts or masters",
    }


def apply_brand(pkg: PptxPackage, brand: dict) -> dict:
    """Write an extract_brand result onto THIS deck's theme(s): all 12
    color slots (empty slots in the brand are skipped) and the major/minor
    latin/ea/cs typefaces, applied to EVERY master's theme so the whole
    deck re-resolves (shared theme parts are edited once). Accepts colors
    as extract_brand emits them ({slot: {"hex": ...}}) or as plain
    {slot: "RRGGBB"}; fonts.major/minor take a dict of latin/ea/cs or a
    plain typeface string (coerced to the latin slot, like
    set_theme_fonts). Explicit srgbClr fills in this deck are NOT touched;
    the brand's explicit_fills list is informational."""
    if not isinstance(brand, dict) or not (
        brand.get("colors") or brand.get("fonts")
    ):
        raise PptMcpError(
            'brand must be an extract_brand result (or at least {"colors": '
            '{...}} / {"fonts": {...}})'
        )

    colors_in = brand.get("colors") or {}
    if not isinstance(colors_in, dict):
        raise PptMcpError(
            "brand colors must be a dict of slot -> RRGGBB hex (or "
            "extract_brand's {slot: {'hex': ...}} entries), got "
            f"{type(colors_in).__name__}"
        )
    normalized: dict[str, str] = {}
    unknown = sorted(set(colors_in) - set(COLOR_SLOTS))
    if unknown:
        raise PptMcpError(
            f"unknown color slot(s) in brand: {', '.join(unknown)}; valid: "
            f"{', '.join(COLOR_SLOTS)}"
        )
    for slot, entry in colors_in.items():
        raw = entry.get("hex") if isinstance(entry, dict) else entry
        if not raw:
            continue  # slot the source theme could not express; skip honestly
        normalized[slot] = _norm_hex(raw, slot)

    fonts_in = brand.get("fonts") or {}
    if not isinstance(fonts_in, dict):
        raise PptMcpError(
            "brand fonts must be a dict with major/minor entries, got "
            f"{type(fonts_in).__name__}"
        )
    allowed_font_keys = {"major", "minor"} | set(_BRAND_FONT_KEYS)
    unknown_fonts = sorted(set(fonts_in) - allowed_font_keys)
    if unknown_fonts:
        raise PptMcpError(
            f"unknown brand font key(s): {', '.join(unknown_fonts)}; valid: {', '.join(sorted(allowed_font_keys))}; nothing was "
            "changed"
        )

    font_specs: dict[str, dict[str, str]] = {}
    for label in ("major", "minor"):
        if label in fonts_in:
            spec = fonts_in[label]
            if isinstance(spec, dict):
                unknown_nested = sorted(set(spec) - set(_FONT_KEYS))
                if unknown_nested:
                    raise PptMcpError(
                        f"unknown brand fonts.{label} font key(s): "
                        f"{', '.join(unknown_nested)}; valid: latin, ea, cs"
                    )
                # extract_brand represents unresolved script slots as empty
                # strings. They remain intentionally absent on transfer.
                cleaned = {
                    key: value
                    for key, value in spec.items()
                    if not (isinstance(value, str) and not value.strip())
                }
                if cleaned:
                    font_specs[label] = _norm_font_spec(
                        cleaned, f"brand fonts.{label}"
                    )
            else:
                normalized_family = _norm_font_spec(
                    spec, f"brand fonts.{label}"
                )
                if normalized_family:
                    font_specs[label] = normalized_family
    for flat_key, (label, script) in _BRAND_FONT_KEYS.items():
        if flat_key not in fonts_in:
            continue
        if script in font_specs.get(label, {}):
            raise PptMcpError(
                f"brand fonts sets both {label}.{script} and {flat_key}, which name "
                "the same font; pass only one. Nothing was changed."
            )
        normalized_flat = _norm_font_spec(
            {script: fonts_in[flat_key]}, f"brand fonts.{flat_key}"
        )
        font_specs.setdefault(label, {}).update(normalized_flat)

    if fonts_in and not font_specs:
        raise PptMcpError(
            "brand carries no usable colors or fonts (all slots empty)"
        )
    if not normalized and not font_specs:
        raise PptMcpError(
            "brand carries no usable colors or fonts (all slots empty)"
        )

    # Every master, each theme part once.
    from .read import _master_parts

    masters = _master_parts(pkg)
    if not masters:
        raise UnsupportedStructure("presentation has no slide masters")
    themes_done: set[str] = set()
    for master_part in masters:
        theme_part = _theme_part_of(pkg, master_part)
        if theme_part in themes_done:
            continue
        themes_done.add(theme_part)
        if normalized:
            scheme = _clr_scheme(pkg, theme_part)
            for slot, hexval in normalized.items():
                _write_slot(scheme, slot, hexval)
        if font_specs:
            fscheme = _font_scheme(pkg, theme_part)
            for label, tag in (("major", "a:majorFont"), ("minor", "a:minorFont")):
                spec = font_specs.get(label)
                if not spec:
                    continue
                font_el = fscheme.find(qn(tag))
                if font_el is None:
                    raise UnsupportedStructure(f"{theme_part} has no {tag}")
                for key, typeface in spec.items():
                    _set_typeface(font_el, key, typeface)
        pkg.mark_dirty(theme_part)

    return {
        "source": brand.get("source", ""),
        "themes_updated": sorted(themes_done),
        "masters": masters,
        "colors_set": normalized,
        "fonts_set": font_specs,
        "note": (
            "theme-linked (schemeClr) content re-resolves everywhere; "
            "explicit srgbClr fills in this deck were not changed"
        ),
    }
