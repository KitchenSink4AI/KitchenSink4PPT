"""Accessibility: audit_accessibility, set_alt_text, set_reading_order.

Sibling parity with word-mcp's ops/accessibility.py (audit + honest-skip
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
- One reading-order rule is NOT a gross-mismatch vote: the slide title is
  read first. A title placeholder that a screen reader reaches after a
  content shape sitting visually below or beside it is flagged on any
  slide, two shapes included (punch-list #928: the vote needed three
  shapes and two inverted pairs, so the commonest case, a title moved
  behind its body, never fired). Shapes the title is painted over and
  shapes marked decorative are exempt: the first cannot move after the
  title without covering it, and screen readers skip the second.
- No suggested reorder changes which of two overlapping shapes paints on
  top (spTree order is also z-order; a shape with unknown geometry counts
  as overlapping everything). When the title cannot be read first without
  such a change, the finding names the chain of overlapping shapes in
  z_order_blocked, and a title_first finding suggests no order at all
  (VERIFY #53 P-B1: the old repair moved a picture above the highlight
  drawn on it).
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
from ..core.package import PptxPackage, qn
from .design_check import CHECKS as _DC_CHECKS
from .design_check import _gentle_box, _run_battery
from .read import (
    _cnvpr,
    _ph,
    iter_shapes,
    resolve_slide,
    shape_text,
    slides_in_scope,
)
from .shapes import _SHAPE_TAGS, _find_shape, _sp_tree

EMU_PER_INCH = 914400

#: Shape kinds a screen reader cannot derive text from: these need descr.
_GRAPHICAL_KINDS = ("picture", "chart", "diagram", "ole", "graphicFrame")

#: The audit's own checks (implemented here) and the delegated ones.
_OWN_CHECKS = ("alt_text", "table_headers", "reading_order")
_DELEGATED_CHECKS = ("missing_title", "contrast")
ALL_CHECKS = _OWN_CHECKS + _DELEGATED_CHECKS

_CAVEATS = {
    "alt_text": (
        "flags pictures, charts, diagrams, OLE objects, and unknown "
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
        "still trip it - review before reordering. Separately, and on any "
        "slide, two shapes included: a title placeholder read AFTER a "
        "content shape that sits visually below or beside it is flagged "
        "as a warning (rule title_first), because screen readers announce "
        "a slide by reading its title first; shapes the title is painted "
        "over and shapes marked decorative are exempt"
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
                "id": int(cnvpr.get("id")) if cnvpr is not None else None,
                "name": cnvpr.get("name", "") if cnvpr is not None else "",
                "hidden": bool(cnvpr is not None and cnvpr.get("hidden") == "1"),
                # Reading order wants the box a placeholder RENDERS with,
                # so the inherited one counts here too.
                "box": _gentle_box(elem, [], pkg, part)[0],
            }
        )
    return out


#: Named apart from generators._label on purpose: the read-only guard
#: (test_readonly_annotations) walks the call graph by function NAME, and
#: that one draws shapes.
def _rec_label(rec: dict) -> str:
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
        if cnvpr.get("hidden") == "1":
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
        if tblpr is not None and tblpr.get("firstRow") == "1":
            continue
        cnvpr = _cnvpr(elem)
        sid = int(cnvpr.get("id")) if cnvpr is not None else None
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


def _is_content_shape(rec: dict) -> bool:
    """Only CONTENT shapes participate in the order walk: text-bearing
    shapes and graphical objects. Untexted autoshapes are decoration
    (bands, accents) that authors legitimately add late in document order,
    and furniture placeholders (slide number/date/footer) sit in
    conventional corners; counting either made the heuristic noisier on a
    real 131-slide deck (noise calibration, 2026-08-31)."""
    if rec["kind"] == "placeholder":
        ph = _ph(rec["elem"])
        if ph is not None and ph.get("type") in _FURNITURE_PH:
            return False
    if rec["kind"] in _GRAPHICAL_KINDS or rec["kind"] == "table":
        return True
    if rec["kind"] == "group":
        # A group is content when something inside it is: grouping two
        # untexted accents must not turn them into content (VERIFY #53
        # P-m1: the same shapes ungrouped were ignored).
        return any(
            kind in _GRAPHICAL_KINDS
            or kind == "table"
            or (kind != "group" and shape_text(el).strip())
            for el, kind, _z, _p in iter_shapes(rec["elem"])
        )
    return bool(shape_text(rec["elem"]).strip())


#: Title placeholder types: the shape a screen reader should announce first.
_TITLE_PH = ("title", "ctrTitle")

#: Office's "Mark as decorative" flag (cNvPr/a:extLst). Screen readers skip
#: a shape carrying it.
_DECORATIVE = "{http://schemas.microsoft.com/office/drawing/2017/decorative}decorative"


def _is_decorative(elem: etree._Element) -> bool:
    cnvpr = _cnvpr(elem)
    if cnvpr is None:
        return False
    return any(
        (d.get("val") or "").lower() in ("1", "true")
        for d in cnvpr.iter(_DECORATIVE)
    )


def _boxes_intersect(a, b) -> bool:
    ax, ay, acx, acy = a
    bx, by, bcx, bcy = b
    return ax < bx + bcx and bx < ax + acx and ay < by + bcy and by < ay + acy


def _may_overlap(a, b) -> bool:
    """Boxes that intersect, or either box unknown: a shape whose geometry
    cannot be resolved is assumed to overlap everything, so no suggested
    reorder ever moves another shape past it."""
    if a is None or b is None:
        return True
    return _boxes_intersect(a, b)


def _overlap_preds(all_records: list[dict]) -> dict[int, set[int]]:
    """For each shape, the shapes EARLIER in spTree that it overlaps: the
    ones it must stay painted above. spTree order is also z-order, so a
    reorder that changes the relative order of such a pair changes which
    one covers the other."""
    ids = [r["id"] for r in all_records]
    box = {r["id"]: r["box"] for r in all_records}
    preds: dict[int, set[int]] = {sid: set() for sid in ids}
    for i, a in enumerate(ids):
        for b in ids[i + 1:]:
            if _may_overlap(box[a], box[b]):
                preds[b].add(a)
    return preds


def _z_safe_order(
    all_records: list[dict], desired: list[int], first: int | None = None
) -> list[int]:
    """The permutation closest to `desired` that keeps every overlapping
    pair in its current relative order (VERIFY #53 P-B1: the title_first
    repair moved a picture above the highlight circle drawn on it, hiding
    the circle). A stable topological sort: at each step the shape placed
    is, of those whose overlapped predecessors are all placed, the one
    `desired` puts first. Current order satisfies every constraint, so a
    result always exists; it may fall short of `desired`, which callers
    check.

    `first` (the slide title) is placed as early as the constraints allow:
    it and every shape that must stay behind it (a chain of overlaps ends
    at it) go before anything else. Without that, shapes the title does
    not depend on could be placed ahead of it while it waits for a
    background to be placed, and the title would be read late for no
    reason (found on a real 131-slide deck)."""
    preds = _overlap_preds(all_records)
    rank = {sid: i for i, sid in enumerate(desired)}
    ahead: set[int] = set()
    if first is not None:
        stack = [first]
        while stack:
            node = stack.pop()
            if node in ahead:
                continue
            ahead.add(node)
            stack.extend(preds[node])
    remaining = [r["id"] for r in all_records]
    placed: set[int] = set()
    out: list[int] = []
    while remaining:
        ready = [s for s in remaining if preds[s] <= placed]
        nxt = min(
            ready, key=lambda s: (s not in ahead, rank.get(s, len(rank)))
        )
        out.append(nxt)
        placed.add(nxt)
        remaining.remove(nxt)
    return out


def _overlap_chain(all_records: list[dict], start: int, goal: int) -> list[int]:
    """Shortest chain start -> ... -> goal of shapes, each later in spTree
    than the one before and overlapping it. Its existence is exactly why
    `goal` cannot move ahead of `start` without changing which of two
    overlapping shapes paints on top. [] when there is none."""
    preds = _overlap_preds(all_records)
    succs: dict[int, list[int]] = {sid: [] for sid in preds}
    for sid, ps in preds.items():
        for p in ps:
            succs[p].append(sid)
    order = {r["id"]: i for i, r in enumerate(all_records)}
    parent: dict[int, int | None] = {start: None}
    frontier = [start]
    while frontier:
        nxt_frontier = []
        for node in frontier:
            for s in sorted(succs[node], key=order.get):
                if s in parent:
                    continue
                parent[s] = node
                if s == goal:
                    chain = [s]
                    while parent[chain[-1]] is not None:
                        chain.append(parent[chain[-1]])
                    return chain[::-1]
                nxt_frontier.append(s)
        frontier = nxt_frontier
    return []


def _title_read_late(
    records: list[dict], visual: list[int]
) -> tuple[dict | None, list[dict]]:
    """(the slide's title record, the content shapes read BEFORE it that
    sit visually after it). The title is the first title/ctrTitle
    placeholder with text among the content records.

    A shape read before the title is only a defect when it sits below or
    beside the title: text ABOVE the title (a kicker line) reads in the
    order the eye takes. Two kinds are exempt, because reordering cannot
    or need not help: a shape the title is painted over (a background
    picture or panel has to stay earlier in spTree, which is also
    z-order, or it would cover the title) and a shape marked decorative
    (screen readers skip it)."""
    title = None
    for r in records:
        if r["kind"] != "placeholder":
            continue
        ph = _ph(r["elem"])
        if ph is not None and ph.get("type") in _TITLE_PH:
            title = r
            break
    if title is None:
        return None, []
    rank = {sid: i for i, sid in enumerate(visual)}
    early = []
    for r in records:
        if r is title:
            break
        if _is_decorative(r["elem"]):
            continue
        if _boxes_intersect(r["box"], title["box"]):
            continue
        if rank[r["id"]] > rank[title["id"]]:
            early.append(r)
    return title, early


def _check_reading_order(pkg: PptxPackage, srec: dict) -> list[dict]:
    all_records = [
        r
        for r in _top_level_records(pkg, srec["part"])
        if r["id"] is not None
    ]
    records = [
        r
        for r in all_records
        if not r["hidden"]
        and r["box"] is not None
        and r["kind"] != "connector"
        # Screen readers skip a shape marked decorative, so it has no
        # reading position to get wrong (VERIFY #53 P-m3: the vote let a
        # decorative picture drive a reorder that put it in front).
        and not _is_decorative(r["elem"])
        and _is_content_shape(r)
    ]
    if len(records) < 2:
        return []  # one shape has no order to get wrong
    doc_order = [r["id"] for r in records]
    visual = [r["id"] for r in _visual_order(records)]
    title, early = _title_read_late(records, visual)

    # The gross-mismatch vote needs enough shapes for a fraction to mean
    # anything; the title rule does not.
    gross = False
    inversions = pairs = 0
    if len(records) >= 3:
        rank = {sid: i for i, sid in enumerate(visual)}
        for i in range(len(doc_order)):
            for j in range(i + 1, len(doc_order)):
                pairs += 1
                if rank[doc_order[i]] > rank[doc_order[j]]:
                    inversions += 1
        gross = (
            pairs > 0
            and inversions >= _MIN_INVERSIONS
            and inversions / pairs >= _INVERSION_FRACTION
        )
    if not gross and not early:
        return []

    early_ids = [r["id"] for r in early]
    title_note = ""
    if early:
        title_note = (
            f"the slide title, {_rec_label(title)}, is read AFTER "
            f"{', '.join(_rec_label(r) for r in early)}, which "
            f"{'sits' if len(early) == 1 else 'sit'} below or beside it, "
            "so a screen reader announces the body before the title"
        )

    if gross:
        # set_reading_order needs a COMPLETE permutation, so the target
        # keeps non-content shapes (decorations, connectors, hidden) at
        # their current relative positions ahead of the reordered content:
        # earliest in spTree = painted behind, which is where decoration
        # belongs. The suggestion is then the closest order to that target
        # that changes no overlapping pair's z-order (VERIFY #53 P-B1).
        content = set(doc_order)
        target = [r["id"] for r in all_records if r["id"] not in content]
        target.extend(visual)
        suggested = _z_safe_order(
            all_records, target, title["id"] if early else None
        )
    else:
        # Minimal repair target: every shape read too early moves to just
        # after the title, in its current relative order. The suggestion
        # is the closest order to that which changes no overlapping pair's
        # z-order: a shape drawn on a moved one travels with it.
        target = []
        for r in all_records:
            if r["id"] in early_ids:
                continue
            target.append(r["id"])
            if r["id"] == title["id"]:
                target.extend(early_ids)
        suggested = _z_safe_order(all_records, target, title["id"])

    # Early shapes the title still cannot precede: each has a chain of
    # overlapping shapes leading to the title, so any order that reads the
    # title first changes which of two overlapping shapes paints on top.
    blocked = []
    if early:
        pos = {sid: i for i, sid in enumerate(suggested)}
        for sid in early_ids:
            if pos[sid] < pos[title["id"]]:
                blocked.append(
                    {"shape_id": sid,
                     "chain": _overlap_chain(all_records, sid, title["id"])}
                )
    by_id = {r["id"]: r for r in all_records}
    blocked_note = "; ".join(
        " -> ".join(_rec_label(by_id[s]) for s in b["chain"]) for b in blocked
    )

    if gross:
        message = (
            f"spTree reading order of the content shapes {doc_order} "
            f"disagrees grossly with the visual top-left order {visual} "
            f"({inversions}/{pairs} pairs inverted); screen readers "
            "follow spTree order. HEURISTIC: a multi-column layout may "
            "legitimately read this way - review before reordering"
        )
        if title_note:
            message += f". Also, {title_note}"
        fix = (
            f"set_reading_order(slide={srec['index']}, "
            f"order={suggested}) rewrites spTree order (decorative "
            "shapes kept first/behind; note: spTree order is also "
            "z-order)"
        )
        if blocked:
            # [[COPY]] wording pending: see the PR #39 repair report.
            fix += (
                "; it cannot read the title first without changing which "
                f"of two overlapping shapes paints on top: {blocked_note}"
            )
        finding = {
            "check": "reading_order",
            "severity": "warning" if early else "info",
            "slide_index": srec["index"],
            "slide_id": srec["slide_id"],
            "message": message,
            "fix": fix,
            "shape_ids": doc_order,
            "rule": "visual_order",
            "document_order": doc_order,
            "visual_order": visual,
            "suggested_order": suggested,
            "inverted_pairs": inversions,
            "pair_count": pairs,
            "heuristic": True,
        }
    else:
        if blocked:
            # No reorder is suggested: every order that reads the title
            # first would hide or uncover part of the slide.
            # [[COPY]] wording pending: see the PR #39 repair report.
            fix = (
                "no set_reading_order is suggested: reading the title "
                "first would change which of two overlapping shapes paints "
                f"on top ({blocked_note}); separate the overlapping shapes "
                "(set_shape) or reorder by hand and check the slide image "
                "(export_slide_images)"
            )
        else:
            # [[COPY]] wording pending: see the PR #39 repair report.
            fix = (
                f"set_reading_order(slide={srec['index']}, "
                f"order={suggested}) moves the title ahead of "
                f"{'it' if len(early_ids) == 1 else 'them'} "
                "(spTree order is also z-order; no two overlapping shapes "
                "change which one paints on top)"
            )
        finding = {
            "check": "reading_order",
            "severity": "warning",
            "slide_index": srec["index"],
            "slide_id": srec["slide_id"],
            "message": (
                f"{title_note[0].upper()}{title_note[1:]}; screen readers "
                "follow spTree order and expect the title first"
            ),
            "fix": fix,
            "shape_ids": [title["id"], *early_ids],
            "rule": "title_first",
            "document_order": doc_order,
            "visual_order": visual,
            "heuristic": False,
        }
        if not blocked:
            finding["suggested_order"] = suggested
    if early:
        finding["title_id"] = title["id"]
        finding["read_before_title"] = early_ids
    if blocked:
        finding["z_order_blocked"] = blocked
    return [finding]


# =============================================================== public API


def audit_accessibility(pkg: PptxPackage, scope=None, *,
                        limit=None, offset: int = 0) -> dict:
    """One-stop accessibility audit over `scope` (None = whole deck, a slide
    selector, or a list of selectors).

    Checks: alt_text (graphical shapes without descr), table_headers
    (tables without firstRow semantics), reading_order (spTree vs visual
    top-left order: the gross-mismatch vote, labeled heuristic, and the
    title_first rule, which is not a heuristic), plus
    missing_title and contrast DELEGATED to check_layout (single
    implementation; EVERY finding merged, not one budget page, and the
    caveats verbatim).

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
    # Merge EVERY finding, not check_layout's first budget page: the audit
    # pages the merged list itself below, and its counts must be complete.
    _plan, _recs, delegated, _stats = _run_battery(
        pkg,
        scope if scope is not None else None,
        list(_DELEGATED_CHECKS),
    )
    findings.extend(delegated)

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


def set_alt_text(pkg: PptxPackage, slide, shape: int, text: str) -> dict:
    """Set (or clear, with an empty string) the alt text of ANY shape by id:
    autoshapes, text boxes, pictures, tables, charts, diagrams, groups,
    connectors. Writes cNvPr/@descr on the shape's own nv*Pr family element
    (the attribute every screen reader and the PowerPoint accessibility
    checker read). media.set_image covers pictures; this covers everything."""
    if not isinstance(text, str):
        raise PptMcpError(f"alt text must be a string, got {text!r}")
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    elem, _chain = _find_shape(pkg, part, shape)
    cnvpr = _cnvpr(elem)
    if cnvpr is None:
        raise UnsupportedStructure(
            f"shape {shape} has no cNvPr element to carry alt text"
        )
    previous = cnvpr.get("descr")
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
    shape_elems = [c for c in sp_tree if c.tag in _SHAPE_TAGS]
    id_map: dict[int, etree._Element] = {}
    old_order: list[int] = []
    for el in shape_elems:
        cnvpr = _cnvpr(el)
        sid = int(cnvpr.get("id")) if cnvpr is not None else None
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
    # shapes exactly as the schema wants.
    for sid in order:
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
            "visually (check_layout overlap, or export_slide_images)"
        ),
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
    }
