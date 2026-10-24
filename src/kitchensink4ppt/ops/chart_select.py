"""Shared chart discovery and selector resolution."""

from __future__ import annotations

from ..core.errors import AmbiguousTarget, PptMcpError, TargetNotFound
from ..core.package import NSMAP, PptxPackage, qn, resolve_target
from .read import iter_shapes, resolve_slide
from .read import cnvpr_id

_C = NSMAP["c"]
_CX = "http://schemas.microsoft.com/office/drawing/2014/chartex"
_R = NSMAP["r"]
_PACKAGE_REL_SUFFIXES = ("/package", "/oleObject")


def charts_on_slide(pkg: PptxPackage, part: str) -> list[dict]:
    """Chart frames in document order, including classic and chartex charts."""
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    out: list[dict] = []
    if sp_tree is None:
        return out
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        if kind not in ("chart", "graphicFrame"):
            continue
        gdata = elem.find(f"{qn('a:graphic')}/{qn('a:graphicData')}")
        if gdata is None:
            continue
        uri = gdata.get("uri")
        if uri == _C:
            chart_el = gdata.find(f"{{{_C}}}chart")
            chart_kind = "chart"
        elif uri == _CX:
            chart_el = gdata.find(f"{{{_CX}}}chart")
            chart_kind = "chartex"
        else:
            continue
        nv = elem.find(qn("p:nvGraphicFramePr"))
        cnvpr = nv.find(qn("p:cNvPr")) if nv is not None else None
        rid = chart_el.get(f"{{{_R}}}id") if chart_el is not None else None
        chart_part = None
        if rid:
            try:
                chart_part = pkg.relationship_target(part, rid)
            except (KeyError, PptMcpError):
                pass
        out.append(
            {
                "index": len(out),
                "shape_id": cnvpr_id(cnvpr),
                "name": cnvpr.get("name", "") if cnvpr is not None else "",
                "kind": chart_kind,
                "part": chart_part,
                "chart_part": chart_part,
                "chartex": chart_kind == "chartex",
            }
        )
    return out


def _candidates(charts: list[dict]) -> str:
    return ", ".join(
        f"index {c['index']} (shape id {c['shape_id']}, {c['kind']})"
        for c in charts
    )


def _selector_int(value, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise PptMcpError(f"chart {label} must be an integer, got {value!r}")
    return value


def resolve_chart(pkg: PptxPackage, slide, chart) -> dict:
    """Resolve explicit index/shape-id selectors and safe legacy integers.

    A bare integer remains compatible when it has exactly one meaning. If it
    is both a valid index and another chart's shape id, resolution refuses so
    callers cannot silently read, edit, and format different charts.
    """
    slide_rec = resolve_slide(pkg, slide)
    part = slide_rec["part"]
    charts = charts_on_slide(pkg, part)
    if not charts:
        raise TargetNotFound(f"slide index {slide_rec['index']} has no charts")

    candidates = _candidates(charts)
    chosen = None
    if chart is None:
        if len(charts) > 1:
            raise AmbiguousTarget(
                f"slide index {slide_rec['index']} has {len(charts)} charts; "
                "pass `chart` as {'index': N} or {'shape_id': N}. Candidates: "
                f"{candidates}"
            )
        chosen = charts[0]
    elif isinstance(chart, dict):
        if set(chart) == {"index"}:
            index = _selector_int(chart["index"], "index")
            if not 0 <= index < len(charts):
                raise TargetNotFound(
                    f"no chart at index {index} on slide index "
                    f"{slide_rec['index']}. Candidates: {candidates}"
                )
            chosen = charts[index]
        elif set(chart) == {"shape_id"}:
            shape_id = _selector_int(chart["shape_id"], "shape id")
            matches = [c for c in charts if c["shape_id"] == shape_id]
            if len(matches) > 1:
                raise AmbiguousTarget(
                    f"shape id {shape_id} occurs on {len(matches)} charts; "
                    f"use an explicit chart index. Candidates: {candidates}"
                )
            chosen = matches[0] if matches else None
            if chosen is None:
                raise TargetNotFound(
                    f"no chart with shape id {shape_id} on slide index "
                    f"{slide_rec['index']}; chart shape ids there: "
                    f"{[c['shape_id'] for c in charts]}. Candidates: {candidates}"
                )
        else:
            raise PptMcpError(
                f"invalid chart selector {chart!r}: use exactly "
                "{'index': N}, {'shape_id': N}, or None for the only chart"
            )
    elif isinstance(chart, int) and not isinstance(chart, bool):
        by_index = charts[chart] if 0 <= chart < len(charts) else None
        shape_matches = [c for c in charts if c["shape_id"] == chart]
        if len(shape_matches) > 1:
            raise AmbiguousTarget(
                f"shape id {chart} occurs on {len(shape_matches)} charts; "
                f"use an explicit chart index. Candidates: {candidates}"
            )
        by_shape = shape_matches[0] if shape_matches else None
        if by_index is not None and by_shape is not None and by_index is not by_shape:
            raise AmbiguousTarget(
                f"legacy integer chart selector {chart} matches chart index "
                f"{chart} (shape id {by_index['shape_id']}) and shape id "
                f"{chart} (index {by_shape['index']}); use {{'index': {chart}}} "
                f"or {{'shape_id': {chart}}}. Candidates: {candidates}"
            )
        chosen = by_index or by_shape
        if chosen is None:
            raise TargetNotFound(
                f"{chart} is neither a chart index (slide has {len(charts)}) "
                f"nor a chart shape id; chart shape ids there: "
                f"{[c['shape_id'] for c in charts]}. Candidates: {candidates}"
            )
    else:
        raise PptMcpError(
            f"invalid chart selector {chart!r}: use a chart shape id or "
            "0-based index via {'shape_id': N} / {'index': N}, a safe legacy "
            "integer, or None for the only chart"
        )

    return {
        "slide_part": part,
        "slide_index": slide_rec["index"],
        "slide_id": slide_rec["slide_id"],
        **chosen,
    }


def chart_data_source(pkg: PptxPackage, chart_part: str) -> dict:
    """Describe the workbook relationship named by c:externalData."""
    external_data = pkg.root(chart_part).find(f"{{{_C}}}externalData")
    if external_data is None:
        return {"kind": "none"}
    rid = external_data.get(f"{{{_R}}}id")
    if not rid:
        return {"kind": "missing", "relationship_id": None}
    try:
        rels = pkg.rels_for(chart_part).getroot()
    except KeyError:
        return {"kind": "missing", "relationship_id": rid}
    rel = next((item for item in rels if item.get("Id") == rid), None)
    if rel is None:
        return {"kind": "missing", "relationship_id": rid}
    rel_type = rel.get("Type", "")
    if not rel_type.endswith(_PACKAGE_REL_SUFFIXES):
        return {
            "kind": "unsupported",
            "relationship_id": rid,
            "relationship_type": rel_type,
        }
    target = rel.get("Target", "")
    if rel.get("TargetMode") == "External":
        return {
            "kind": "linked",
            "relationship_id": rid,
            "target": target,
        }
    return {
        "kind": "embedded",
        "relationship_id": rid,
        "part": resolve_target(chart_part, target),
    }
