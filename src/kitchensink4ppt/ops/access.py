"""Accessibility: audit_accessibility, set_alt_text, set_reading_order.

Parity with the Word editor's accessibility audit (audit + honest-skip
philosophy) expressed in THIS server's finding format: every finding follows
check_layout's shape (check / severity / slide_index / slide_id / message /
fix / shape_ids), so an agent that already consumes check_layout output can
consume this audit unchanged.

Contract (all ops modules): functions take the open PptxPackage first and
return plain dicts. audit_accessibility is READ-ONLY; set_alt_text and
set_reading_order mutate only the in-memory package and call mark_dirty().

Division of labor (deliberate, to avoid duplicated logic):
- missing_title and contrast are check_layout's checks. This audit CALLS
  check_layout for those two categories and merges the findings verbatim,
  so there is exactly one implementation of each heuristic in the codebase
  and the audit is still a one-stop report. Their caveats are carried over.
- alt_text, table_headers, and reading_order are implemented here: they are
  accessibility semantics, not layout guardrails.

Honesty rules (binding, same as design_check):
- reading_order is a HEURISTIC (spTree order vs a visual top-left band
  sort); it flags gross mismatches only, says so in the finding, and its
  suggested order is a starting point, not truth (multi-column layouts
  legitimately read column-first).
- Shapes whose geometry cannot be resolved are skipped, never guessed.
- Autoshapes and text boxes are NOT flagged for alt text: screen readers
  read their text content; untexted autoshapes are treated as decorative.
  Only content a screen reader cannot derive text from is flagged
  (pictures, charts, diagrams, OLE objects, unknown graphicFrames).
"""

from __future__ import annotations

from lxml import etree

from ..core import budget as _budget
from ..core.errors import PptMcpError, TargetNotFound, UnsupportedStructure
from ..core.messages import TITLE_READ_AFTER_BODY
from ..core.package import PptxPackage, qn
from ..core.xmlbool import attr_bool
from .design_check import CHECKS as _DC_CHECKS
from .design_check import _gentle_box, check_layout
from .read import (
    _cnvpr,
    _ph,
    iter_shapes,
    resolve_slide,
    shape_text,
    slides_in_scope,
    stack_members,
    wrapped_shape_id,
)
from .read import cnvpr_id
from .shapes import _SHAPE_TAGS, _find_shape, _sp_tree

EMU_PER_INCH = 914400

#: Shape kinds a screen reader cannot derive text from: these need descr.
_GRAPHICAL_KINDS = ("picture", "media", "chart", "diagram", "ole", "graphicFrame")

#: The audit's own checks (implemented here) and the delegated ones.
_OWN_CHECKS = ("alt_text", "table_headers", "reading_order")
_DELEGATED_CHECKS = ("missing_title", "contrast")
ALL_CHECKS = _OWN_CHECKS + _DELEGATED_CHECKS

_CAVEATS = {
    "alt_text": (
        "flags pictures, video and audio, charts, diagrams, OLE objects, and unknown "
        "graphicFrames whose cNvPr carries no (or an empty) descr; "
        "autoshapes/text boxes are not flagged because screen readers read "
        "their text, and untexted autoshapes are treated as decorative"
    ),
    "table_headers": (
        "a table whose tblPr lacks firstRow='1' gives assistive technology "
        "no header-row semantics; single-row tables are skipped (their only "
        "row being a header needs human judgment)"
    ),
    "reading_order": (
        "HEURISTIC: compares spTree document order (what screen readers "
        "follow) against a top-left visual band sort over CONTENT shapes "
        "only (text-bearing shapes, pictures, tables, charts, groups; "
        "untexted decorative autoshapes are ignored - counting them cried "
        "wolf on real decks); only gross mismatches are flagged, and "
        "multi-column layouts that legitimately read column-first can "
        "still trip it - review before reordering"
    ),
}


# ------------------------------------------------------------- own checks


def _top_level_records(pkg: PptxPackage, part: str) -> list[dict]:
    """Top-level shape records with slide-space boxes (None when the shape
    has no explicit geometry) in document order."""
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    out: list[dict] = []
    if sp_tree is None:
        return out
    for elem, kind, _z, parent in iter_shapes(sp_tree):
        if parent is not None:
            continue
        cnvpr = _cnvpr(elem)
        out.append(
            {
                "elem": elem,
                "kind": kind,
                "id": cnvpr_id(cnvpr),
                "name": cnvpr.get("name", "") if cnvpr is not None else "",
                "hidden": attr_bool(cnvpr, "hidden"),
                # Reading order wants the box a placeholder RENDERS with,
                # so the inherited one counts here too.
                "box": _gentle_box(elem, [], pkg, part)[0],
            }
        )
    return out


def _label(rec: dict) -> str:
    name = rec.get("name") or ""
    return f"shape {rec['id']}" + (f" ({name!r})" if name else "")


def _check_alt_text(pkg: PptxPackage, srec: dict) -> list[dict]:
    findings = []
    part = srec["part"]
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if sp_tree is None:
        return findings
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        if kind not in _GRAPHICAL_KINDS:
            continue
        cnvpr = _cnvpr(elem)
        if cnvpr is None:
            continue  # no cNvPr to carry descr; nothing to judge
        if attr_bool(cnvpr, "hidden"):
            continue
        if _is_office_decorative({"elem": elem}):
            continue
        if (cnvpr.get("descr") or "").strip():
            continue
        sid_raw = cnvpr.get("id")
        sid = int(sid_raw) if sid_raw and sid_raw.isdigit() else None
        name = cnvpr.get("name", "")
        findings.append(
            {
                "check": "alt_text",
                "severity": "warning",
                "slide_index": srec["index"],
                "slide_id": srec["slide_id"],
                "message": (
                    f"{kind} shape {sid}" + (f" ({name!r})" if name else "")
                    + " has no alt text; a screen reader announces it as an "
                    "unnamed object"
                ),
                "fix": (
                    f"set_alt_text(slide={srec['index']}, shape={sid}, "
                    'text="...") - describe the content, or state that it '
                    "is decorative"
                ),
                "shape_ids": [sid] if sid is not None else [],
                "kind": kind,
            }
        )
    return findings


def _check_table_headers(pkg: PptxPackage, srec: dict) -> list[dict]:
    findings = []
    part = srec["part"]
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if sp_tree is None:
        return findings
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        if kind != "table":
            continue
        data = elem.find(f"{qn('a:graphic')}/{qn('a:graphicData')}")
        tbl = data.find(qn("a:tbl")) if data is not None else None
        if tbl is None:
            continue
        rows = tbl.findall(qn("a:tr"))
        if len(rows) <= 1:
            continue  # single-row table: header judgment is human territory
        tblpr = tbl.find(qn("a:tblPr"))
        if attr_bool(tblpr, "firstRow"):
            continue
        cnvpr = _cnvpr(elem)
        sid = cnvpr_id(cnvpr)
        findings.append(
            {
                "check": "table_headers",
                "severity": "warning",
                "slide_index": srec["index"],
                "slide_id": srec["slide_id"],
                "message": (
                    f"table {sid} ({len(rows)} rows) has no header-row "
                    "semantics (tblPr firstRow flag); assistive technology "
                    "cannot announce column context"
                ),
                "fix": (
                    f"apply_table_style(slide={srec['index']}, table={sid}, "
                    "first_row=True) marks the first row as the header"
                ),
                "shape_ids": [sid] if sid is not None else [],
                "rows": len(rows),
            }
        )
    return findings


#: Shapes whose y-centers fall within this band merge into one visual row.
_ROW_BAND_EMU = round(0.5 * EMU_PER_INCH)

#: Gross-mismatch bar: at least this fraction of shape pairs inverted, and
#: at least this many inverted pairs, before the heuristic speaks up.
_INVERSION_FRACTION = 0.4
_MIN_INVERSIONS = 2


def _visual_order(records: list[dict]) -> list[dict]:
    """Top-left reading order: sort by y, band shapes whose vertical centers
    sit within _ROW_BAND_EMU into rows, order rows internally by x."""
    by_y = sorted(records, key=lambda r: r["box"][1] + r["box"][3] / 2)
    rows: list[list[dict]] = []
    for rec in by_y:
        yc = rec["box"][1] + rec["box"][3] / 2
        if rows:
            first = rows[-1][0]
            if abs(yc - (first["box"][1] + first["box"][3] / 2)) <= _ROW_BAND_EMU:
                rows[-1].append(rec)
                continue
        rows.append([rec])
    ordered: list[dict] = []
    for row in rows:
        ordered.extend(sorted(row, key=lambda r: r["box"][0]))
    return ordered


#: Furniture placeholder types whose corner positions are conventional;
#: their spTree position says nothing about content reading order.
_FURNITURE_PH = ("sldNum", "dt", "ftr")
_DECORATIVE_NS = "http://schemas.microsoft.com/office/drawing/2017/decorative"
_DECORATIVE_EXT_URI = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}"


def _is_office_decorative(rec: dict) -> bool:
    """Whether Office marked this shape to be skipped by screen readers."""
    cnvpr = _cnvpr(rec["elem"])
    if cnvpr is None:
        return False
    marker = cnvpr.find(f".//{{{_DECORATIVE_NS}}}decorative")
    return marker is not None and attr_bool(marker, "val", True)


def _is_content_shape(rec: dict) -> bool:
    """Only CONTENT shapes participate in the order walk: text-bearing
    shapes and graphical objects. Untexted autoshapes are decoration
    (bands, accents) that authors legitimately add late in document order,
    and furniture placeholders (slide number/date/footer) sit in
    conventional corners; counting either made the heuristic noisier on a
    real 131-slide deck."""
    if _is_office_decorative(rec):
        return False
    if rec["kind"] == "placeholder":
        ph = _ph(rec["elem"])
        if ph is not None and ph.get("type") in _FURNITURE_PH:
            return False
    if rec["kind"] in _GRAPHICAL_KINDS or rec["kind"] in ("table", "group"):
        return True
    return bool(shape_text(rec["elem"]).strip())


def _document_order_records(pkg: PptxPackage, part: str) -> list[dict]:
    """Visible shape records in XML preorder, including members of groups.

    The ordinary inversion heuristic deliberately works only at the top
    level because ``set_reading_order`` can only reorder top-level shapes.
    The title rule needs the complete XML sequence: a title inside a group
    can still follow body content inside that group. The parent id is kept so
    the resulting finding can state that the automatic top-level repair is
    unavailable rather than pretending a group member is directly movable.
    """
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if sp_tree is None:
        return []
    records: list[dict] = []
    for elem, kind, _z, parent in iter_shapes(sp_tree):
        cnvpr = _cnvpr(elem)
        sid = cnvpr_id(cnvpr)
        records.append(
            {
                "elem": elem,
                "kind": kind,
                "id": sid,
                "parent_id": parent,
                "hidden": attr_bool(cnvpr, "hidden"),
            }
        )
    return records


def _title_placeholder_type(rec: dict) -> str | None:
    if rec["kind"] != "placeholder":
        return None
    ph = _ph(rec["elem"])
    if ph is None:
        return None
    ph_type = ph.get("type")
    return ph_type if ph_type in ("title", "ctrTitle") else None


def _check_title_before_body(
    pkg: PptxPackage,
    srec: dict,
    top_level_records: list[dict],
) -> list[dict]:
    """Flag every non-empty title placeholder that follows body content.

    This is a direct XML-order fact rather than the visual-order heuristic.
    Grouped members are reported too, but labeled heuristic because assistive
    technologies differ in how they announce a group and the top-level
    reorder operation cannot move an individual member.
    """
    ordered = [
        rec
        for rec in _document_order_records(pkg, srec["part"])
        if rec["id"] is not None
        and not rec["hidden"]
        and rec["kind"] != "connector"
        and rec["kind"] != "group"
        and _is_content_shape(rec)
    ]
    document_order = [rec["id"] for rec in ordered]
    top_level_order = [
        rec["id"] for rec in top_level_records if rec["id"] is not None
    ]
    findings: list[dict] = []
    for position, title in enumerate(ordered):
        title_type = _title_placeholder_type(title)
        if title_type is None or not shape_text(title["elem"]).strip():
            continue
        preceding = [
            rec
            for rec in ordered[:position]
            if _title_placeholder_type(rec) is None
        ]
        if not preceding:
            continue
        title_id = title["id"]
        preceding_ids = [rec["id"] for rec in preceding]
        group_id = title["parent_id"]
        grouped = group_id is not None
        if grouped:
            suggested = None
            fix = (
                "call ppt_graphics with op='ungroup_shapes' and args="
                f"{{'slide': {srec['index']}, 'group': {group_id}}}"
            )
        else:
            suggested = [title_id] + [
                sid for sid in top_level_order if sid != title_id
            ]
            fix = (
                "call ppt_graphics with op='set_reading_order' and args="
                f"{{'slide': {srec['index']}, 'order': {suggested}}}"
            )
        findings.append(
            {
                "check": "reading_order",
                "severity": "warning",
                "slide_index": srec["index"],
                "slide_id": srec["slide_id"],
                "message": TITLE_READ_AFTER_BODY,
                "fix": fix,
                "shape_ids": preceding_ids + [title_id],
                "reason": "title_read_after_body",
                "title_shape_id": title_id,
                "title_placeholder_type": title_type,
                "body_before_title_ids": preceding_ids,
                "document_order": document_order,
                "suggested_order": suggested,
                "title_group_id": group_id,
                "grouped_title": grouped,
                "heuristic": grouped,
            }
        )
    return findings


def _check_reading_order(pkg: PptxPackage, srec: dict) -> list[dict]:
    all_records = [
        r
        for r in _top_level_records(pkg, srec["part"])
        if r["id"] is not None
    ]
    findings = _check_title_before_body(pkg, srec, all_records)
    records = [
        r
        for r in all_records
        if not r["hidden"]
        and r["box"] is not None
        and r["kind"] != "connector"
        and _is_content_shape(r)
    ]
    if len(records) < 3:
        return findings  # title order is meaningful even with two shapes
    doc_order = [r["id"] for r in records]
    visual = [r["id"] for r in _visual_order(records)]
    rank = {sid: i for i, sid in enumerate(visual)}
    inversions = 0
    pairs = 0
    for i in range(len(doc_order)):
        for j in range(i + 1, len(doc_order)):
            pairs += 1
            if rank[doc_order[i]] > rank[doc_order[j]]:
                inversions += 1
    if pairs == 0:
        return findings
    frac = inversions / pairs
    if inversions < _MIN_INVERSIONS or frac < _INVERSION_FRACTION:
        return findings
    # set_reading_order needs a COMPLETE permutation, so the suggestion
    # keeps non-content shapes (decorations, connectors, hidden) at their
    # current relative positions ahead of the reordered content: earliest
    # in spTree = painted behind, which is where decoration belongs.
    content = set(doc_order)
    suggested = [r["id"] for r in all_records if r["id"] not in content]
    suggested.extend(visual)
    findings.append(
        {
            "check": "reading_order",
            "severity": "info",
            "slide_index": srec["index"],
            "slide_id": srec["slide_id"],
            "message": (
                f"spTree reading order of the content shapes {doc_order} "
                f"disagrees grossly with the visual top-left order {visual} "
                f"({inversions}/{pairs} pairs inverted); screen readers "
                "follow spTree order. HEURISTIC: a multi-column layout may "
                "legitimately read this way - review before reordering"
            ),
            "fix": (
                "call ppt_graphics with op='set_reading_order' and args="
                f"{{'slide': {srec['index']}, 'order': {suggested}}} to "
                "rewrite spTree order (decorative shapes kept first/behind; "
                "note: spTree order is also z-order)"
            ),
            "shape_ids": doc_order,
            "document_order": doc_order,
            "visual_order": visual,
            "suggested_order": suggested,
            "inverted_pairs": inversions,
            "pair_count": pairs,
            "heuristic": True,
        }
    )
    return findings


# =============================================================== public API


def audit_accessibility(pkg: PptxPackage, scope=None, *,
                        limit=None, offset: int = 0) -> dict:
    """One-stop accessibility audit over `scope` (None = whole deck, a slide
    selector, or a list of selectors).

    Checks: alt_text (graphical shapes without descr), table_headers
    (tables without firstRow semantics), reading_order (spTree vs visual
    top-left order, gross mismatches only, labeled heuristic), plus
    missing_title and contrast DELEGATED to check_layout (single
    implementation; findings and caveats merged verbatim).

    Findings use check_layout's format: check, severity, slide_index,
    slide_id, message, fix (naming the exact repairing tool call),
    shape_ids. Read-only; nothing is modified."""
    recs = slides_in_scope(pkg, scope)
    findings: list[dict] = []
    for rec in recs:
        srec = {"index": rec["index"], "slide_id": rec["slide_id"], "part": rec["part"]}
        findings.extend(_check_alt_text(pkg, srec))
        findings.extend(_check_table_headers(pkg, srec))
        findings.extend(_check_reading_order(pkg, srec))

    # Delegated categories: one implementation, living in check_layout.
    delegated = check_layout(
        pkg,
        scope if scope is not None else None,
        list(_DELEGATED_CHECKS),
    )
    findings.extend(delegated["findings"])

    sev_rank = {"error": 0, "warning": 1, "info": 2}
    findings.sort(key=lambda f: (sev_rank.get(f["severity"], 3), f["slide_index"]))
    by_check: dict[str, int] = {}
    for f in findings:
        by_check[f["check"]] = by_check.get(f["check"], 0) + 1
    caveats = dict(_CAVEATS)
    for name in _DELEGATED_CHECKS:
        caveats[name] = _DC_CHECKS[name][1]
    # Counts and the per-check summary stay complete: they are the audit's
    # answer. Only the finding list pages.
    header = {
        "slides_checked": len(recs),
        "checks_run": list(ALL_CHECKS),
        "finding_count": len(findings),
        "by_severity": {
            sev: sum(1 for f in findings if f["severity"] == sev)
            for sev in ("error", "warning", "info")
        },
        "by_check": by_check,
    }
    kept, page = _budget.page_items(
        findings, limit=limit, offset=offset,
        overhead=_budget.overhead_of(header) + len(str(caveats)) + 300,
        unit="findings",
        narrow_hint="narrow with scope=<slide index or list of indexes>",
        shrink=_budget.shrink_record,
    )
    result = {
        **header,
        "findings": kept,
        "caveats": caveats,
        "note": (
            "static-XML checks; missing_title and contrast come from "
            "check_layout (same implementation, merged here so the audit "
            "is one-stop). reading_order is a labeled heuristic."
        ),
    }
    if page is not None:
        result["page"] = page
    return result


def _set_office_decorative(cnvpr: etree._Element, enabled: bool) -> None:
    ext_lst = cnvpr.find(qn("a:extLst"))
    markers = (
        [] if ext_lst is None
        else ext_lst.findall(f".//{{{_DECORATIVE_NS}}}decorative")
    )
    if enabled:
        if markers:
            markers[0].set("val", "1")
            for marker in markers[1:]:
                marker.getparent().remove(marker)
            return
        if ext_lst is None:
            ext_lst = etree.SubElement(cnvpr, qn("a:extLst"))
        ext = etree.SubElement(ext_lst, qn("a:ext"), uri=_DECORATIVE_EXT_URI)
        etree.SubElement(
            ext, f"{{{_DECORATIVE_NS}}}decorative",
            nsmap={"adec": _DECORATIVE_NS},
        ).set("val", "1")
        return
    for marker in markers:
        ext = marker.getparent()
        ext.remove(marker)
        if len(ext) == 0:
            ext.getparent().remove(ext)
    if ext_lst is not None and len(ext_lst) == 0:
        cnvpr.remove(ext_lst)


def set_alt_text(
    pkg: PptxPackage,
    slide,
    shape: int,
    text: str,
    decorative: bool | None = None,
) -> dict:
    """Set (or clear, with an empty string) the alt text of ANY shape by id:
    autoshapes, text boxes, pictures, tables, charts, diagrams, groups,
    connectors. Writes cNvPr/@descr on the shape's own nv*Pr family element
    (the attribute every screen reader and the PowerPoint accessibility
    checker read). media.set_image covers pictures; this covers everything."""
    if not isinstance(text, str):
        raise PptMcpError(f"alt text must be a string, got {text!r}")
    if decorative is not None and not isinstance(decorative, bool):
        raise PptMcpError(f"decorative must be true or false, got {decorative!r}")
    if decorative is True and text.strip():
        raise PptMcpError(
            "a decorative shape cannot also carry alt text; pass text=''"
        )
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    elem, _chain = _find_shape(pkg, part, shape)
    cnvpr = _cnvpr(elem)
    if cnvpr is None:
        raise UnsupportedStructure(
            f"shape {shape} has no cNvPr element to carry alt text"
        )
    previous = cnvpr.get("descr")
    previous_decorative = _is_office_decorative({"elem": elem})
    if decorative is True:
        cnvpr.attrib.pop("descr", None)
        _set_office_decorative(cnvpr, True)
    elif decorative is False:
        _set_office_decorative(cnvpr, False)
    elif text.strip() and previous_decorative:
        # Describing an object makes it content again; do not leave mutually
        # contradictory Office accessibility state behind.
        _set_office_decorative(cnvpr, False)
    if text.strip():
        cnvpr.set("descr", text)
        action = "set"
    else:
        cnvpr.attrib.pop("descr", None)
        action = "cleared"
    pkg.mark_dirty(part)
    return {
        "shape_id": shape,
        "changed_ids": [shape],
        "action": action,
        "alt_text": text if text.strip() else None,
        "previous": previous,
        "decorative": _is_office_decorative({"elem": elem}),
        "previous_decorative": previous_decorative,
        "kind": etree.QName(elem).localname,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
    }


def set_reading_order(pkg: PptxPackage, slide, order: list[int]) -> dict:
    """Rewrite the spTree order of one slide's TOP-LEVEL shapes to `order`
    (a complete permutation of the slide's top-level shape ids, first read =
    first in the list).

    spTree order is BOTH the screen-reader reading order and the z-order
    (later = painted on top), so reordering can change which shapes overlap
    visually; the result reports every shape whose relative depth changed so
    the caller can verify (check_layout's overlap check, or export and
    look). A partial list is refused rather than guessing where unlisted
    shapes should land."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    sp_tree = _sp_tree(pkg, part)
    # Shapes, and the mc:AlternateContent wrappers PowerPoint stores an
    # equation text box or ink in: each holds one place in the order.
    shape_elems = stack_members(sp_tree)
    id_map: dict[int, etree._Element] = {}
    old_order: list[int] = []
    wrapped_ids: set[int] = set()
    for el in shape_elems:
        if el.tag in _SHAPE_TAGS:
            cnvpr = _cnvpr(el)
            sid = cnvpr_id(cnvpr)
        else:
            sid = wrapped_shape_id(el)
            wrapped_ids.add(sid)
        if sid is None:
            raise UnsupportedStructure(
                "a top-level shape has no cNvPr id; cannot reorder this "
                "slide safely"
            )
        id_map[sid] = el
        old_order.append(sid)
    if not isinstance(order, list) or not all(
        isinstance(i, int) and not isinstance(i, bool) for i in order
    ):
        raise PptMcpError("order must be a list of shape ids (ints)")
    if len(order) != len(set(order)):
        raise PptMcpError(f"duplicate shape ids in order: {order}")
    if set(order) != set(old_order):
        missing = sorted(set(old_order) - set(order))
        unknown = sorted(set(order) - set(old_order))
        detail = []
        if missing:
            detail.append(f"missing ids {missing}")
            held = sorted(set(missing) & wrapped_ids)
            if held:
                detail.append(
                    f"{held} {'is a shape' if len(held) == 1 else 'are shapes'} PowerPoint "
                    "stores in a compatibility wrapper (an equation text box or ink); "
                    "include it where it should be read"
                )
        if unknown:
            detail.append(f"unknown ids {unknown}")
        raise TargetNotFound(
            "order must be a complete permutation of the slide's top-level "
            f"shape ids {sorted(old_order)}: {'; '.join(detail)}"
        )
    if order == old_order:
        return {
            "changed": False,
            "order": old_order,
            "z_order_changes": [],
            "slide_index": rec["index"],
            "slide_id": rec["slide_id"],
        }
    # Re-append in the new order: appending MOVES each element to the end,
    # and non-shape spTree children (nvGrpSpPr, grpSpPr) stay ahead of all
    # shapes exactly as the schema wants. A trailing p:extLst stays last.
    tail = sp_tree.find(qn("p:extLst"))
    for sid in order:
        if tail is not None:
            tail.addprevious(id_map[sid])
        else:
            sp_tree.append(id_map[sid])
    old_rank = {sid: i for i, sid in enumerate(old_order)}
    new_rank = {sid: i for i, sid in enumerate(order)}
    z_changes = [
        {"shape_id": sid, "z_from": old_rank[sid], "z_to": new_rank[sid]}
        for sid in order
        if old_rank[sid] != new_rank[sid]
    ]
    pkg.mark_dirty(part)
    return {
        "changed": True,
        "order": order,
        "previous_order": old_order,
        "changed_ids": [c["shape_id"] for c in z_changes],
        "z_order_changes": z_changes,
        "warning": (
            "spTree order is also z-order: shapes later in the list now "
            "paint ON TOP of earlier ones; verify overlapping regions "
            "visually with ppt_inspect op='check_layout' (overlap), or "
            "ppt_deck op='export_slide_image'"
        ),
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
    }
