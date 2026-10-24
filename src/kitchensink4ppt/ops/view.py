"""View layer: get_presentation_view, the anchored markdown projection of a
deck, plus resolve_anchor, the scheme's inverse.

Contract: functions take a PptxPackage first, never touch disk, never call
mark_dirty(), never import FastMCP or win32. The view is DETERMINISTIC: two
reads of identical bytes produce identical output.

THE ANCHOR SCHEME (apply_edits resolves these; it must stay
re-derivable from package content alone):
- Slide anchors are the durable p:sldId id verbatim: "s:258" in the slide
  header "## Slide 3 [s:258]". Slide ids survive reordering; 1-based
  display numbers and 0-based indices do not.
- Shape anchors are a short stable content-derived hash, NOT a list
  position: sha1("<slide_id>/<shape_id>") hexdigest, displayed as the
  first 6 hex chars ("[a:3f9c2a]"), extended to 10 (then the full digest)
  only for the rare shapes whose 6-char prefixes collide within one deck.
  slide_id is the p:sldId id, shape_id the shape's p:cNvPr id (unique per
  slide). Both inputs survive slide reorders and edits to OTHER shapes,
  so anchors are stable across exactly those operations; deleting or
  re-id-ing the shape itself invalidates its anchor, which is the point,
  with one exception: a new shape takes one more than the highest id left
  on its slide (as in PowerPoint), so after the highest-numbered shape is
  deleted the next new shape takes its id and answers to its old anchor.
  Slide ids work the same way. The delete answers say so when it happens;
  read the view again before reusing anchors across deletes and inserts.
- Table cells are addressed "t:<shape anchor>:rNcN" with 1-BASED row and
  column numbers (r1c1 = top-left cell).
- resolve_anchor(pkg, anchor) recomputes digests for every shape in the
  deck and matches the given hex as a digest PREFIX. Zero matches raise
  TargetNotFound (stale anchor: re-run get_presentation_view); several
  matches raise AmbiguousTarget listing candidates, with
  {"shape": {"slide": i, "id": n}} as the unambiguous fallback address.

Detail levels: "outline" = slide headers + titles only; "text" (default) =
one block per text-bearing shape, bullets as indented "-" lines, tables as
pipe tables, speaker notes as "> " quoted blocks; "full" = every shape
(text or not) with geometry in inches. Hidden slides are marked (hidden).
Multi-paragraph shape text renders one "-" line per paragraph, indented
two spaces per a:pPr outline level; single paragraphs render inline as
prose. This is a projection for agents, not a fidelity claim.
"""

from __future__ import annotations

import hashlib
import re

from ..core import budget as _budget
from ..core import messages as _messages
from ..core.errors import AmbiguousTarget, PptMcpError, TargetNotFound
from ..core.package import PptxPackage, qn
from . import read as _read

_DETAILS = ("outline", "text", "full")

#: Kinds the view shows by name and box only (detail="full"); they hold no text.
_NO_TEXT_KINDS = ("group", "picture", "media", "chart", "diagram", "ole", "graphicFrame", "ink")

_ANCHOR_LEN = 6
_ANCHOR_LEN_EXTENDED = 10


def _digest(slide_id: int, shape_id: int) -> str:
    return hashlib.sha1(f"{slide_id}/{shape_id}".encode()).hexdigest()


def _shape_map(pkg: PptxPackage) -> list[dict]:
    """Every shape in the deck (groups recursed) with its full digest:
    {"slide_index","slide_id","part","shape_id","kind","elem","digest"}."""
    out = []
    for rec in _read.slide_table(pkg):
        sp_tree = pkg.root(rec["part"]).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
        if sp_tree is None:
            continue
        for elem, kind, _z, _parent in _read.iter_shapes(sp_tree, wrapped=True):
            cnvpr = _read._cnvpr(elem)
            sid = _read.cnvpr_id(cnvpr)
            if sid is None:
                continue  # no id (a damaged deck): no anchor to give
            out.append(
                {
                    "slide_index": rec["index"],
                    "slide_id": rec["slide_id"],
                    "part": rec["part"],
                    "shape_id": sid,
                    "kind": kind,
                    "elem": elem,
                    "digest": _digest(rec["slide_id"], sid),
                }
            )
    return out


def _display_anchors(shape_map: list[dict]) -> dict[tuple[int, int], str]:
    """(slide_id, shape_id) -> display anchor. 6 hex chars, extended only
    for colliding prefixes so the projection stays short AND unambiguous."""
    by_prefix: dict[str, list[dict]] = {}
    for s in shape_map:
        by_prefix.setdefault(s["digest"][:_ANCHOR_LEN], []).append(s)
    out = {}
    for prefix, group in by_prefix.items():
        if len(group) == 1:
            s = group[0]
            out[(s["slide_id"], s["shape_id"])] = prefix
            continue
        longer: dict[str, list[dict]] = {}
        for s in group:
            longer.setdefault(s["digest"][:_ANCHOR_LEN_EXTENDED], []).append(s)
        for lp, lgroup in longer.items():
            for s in lgroup:
                out[(s["slide_id"], s["shape_id"])] = (
                    lp if len(lgroup) == 1 else s["digest"]
                )
    return out


# ------------------------------------------------------------ anchor inverse


_ANCHOR_RE = re.compile(
    r"^(?:s:(?P<sid>\d+)|(?P<t>t:)?(?:a:)?(?P<hex>[0-9a-f]{4,40})"
    r"(?(t):r(?P<row>\d+)c(?P<col>\d+)))$"
)

#: A shape anchor as the view prints it, "a:3f9c2a", brackets optional. A
#: `shape` argument takes this form; the bare hex apply_edits also reads is
#: left to it, because a bare word such as "face" could be a shape's name.
_SHAPE_ANCHOR_RE = re.compile(r"\[?a:([0-9a-fA-F]{4,40})\]?")
#: The view's slide and table-cell anchors, which name no shape.
_OTHER_ANCHOR_RE = re.compile(
    r"\[?(?:s:\d+|t:(?:a:)?[0-9a-fA-F]{4,40}:r\d+c\d+)\]?")


def shape_anchor_hex(ref) -> str | None:
    """The lowercase hex of a shape anchor ("a:3f9c2a" or "[a:3f9c2a]"), or
    None when `ref` is not one."""
    if not isinstance(ref, str):
        return None
    m = _SHAPE_ANCHOR_RE.fullmatch(ref.strip())
    return m.group(1).lower() if m else None


def is_slide_or_cell_anchor(ref) -> bool:
    """Whether `ref` is the view's slide anchor (s:N) or a table-cell anchor
    (t:hex:rNcN): an address, but not of a shape."""
    return isinstance(ref, str) and _OTHER_ANCHOR_RE.fullmatch(ref.strip()) is not None


def anchor_digest(slide_id: int, shape_id: int) -> str:
    """The full digest a shape's anchor is a prefix of."""
    return _digest(slide_id, shape_id)


def resolve_anchor(pkg: PptxPackage, anchor: str) -> dict:
    """Resolve a view anchor back to its target. Accepts "s:<slide_id>",
    "a:<hex>" (or bare hex), and "t:<hex>:rNcN" (1-based row/col). Returns
    {"kind", "slide_index", "slide_id", "part", ...} with "shape_id" and
    "shape_type" for shapes, plus 0-based "row"/"col" for table cells.
    Stale anchors raise TargetNotFound; prefix collisions raise
    AmbiguousTarget with unambiguous candidate addresses."""
    if not isinstance(anchor, str):
        raise PptMcpError(f"anchor must be a string, got {type(anchor).__name__}")
    m = _ANCHOR_RE.match(anchor.strip())
    if m is None:
        raise PptMcpError(
            f"malformed anchor {anchor!r}: expected 's:<slide_id>', "
            "'a:<hex>', or 't:<hex>:rNcN'"
        )

    if m.group("sid") is not None:
        sid = int(m.group("sid"))
        rec = _read.resolve_slide(pkg, {"slide_id": sid})
        return {"kind": "slide", "slide_index": rec["index"],
                "slide_id": rec["slide_id"], "part": rec["part"]}

    hexpart = m.group("hex")
    hits = [s for s in _shape_map(pkg) if s["digest"].startswith(hexpart)]
    if not hits:
        raise TargetNotFound(
            f"anchor {anchor!r} matches no shape in {pkg.path.name}; the "
            "view is stale, re-run get_presentation_view"
        )
    if len(hits) > 1 and len({(s["slide_id"], s["shape_id"]) for s in hits}) == 1:
        from .shapes import _shared_id_refusal

        raise _shared_id_refusal(hits[0]["shape_id"], [s["elem"] for s in hits])
    if len(hits) > 1:
        cands = [
            {"shape": {"slide": s["slide_index"], "id": s["shape_id"]},
             "anchor": s["digest"][:_ANCHOR_LEN_EXTENDED]}
            for s in hits
        ]
        raise AmbiguousTarget(
            f"anchor {anchor!r} matches {len(hits)} shapes; use a longer "
            f"prefix or a shape address. Candidates: {cands}"
        )
    hit = hits[0]
    out = {
        "kind": "shape",
        "slide_index": hit["slide_index"],
        "slide_id": hit["slide_id"],
        "part": hit["part"],
        "shape_id": hit["shape_id"],
        "shape_type": hit["kind"],
    }
    if m.group("t"):
        if hit["kind"] != "table":
            raise PptMcpError(
                f"anchor {anchor!r} uses cell addressing but the shape is a "
                f"{hit['kind']}, not a table"
            )
        cells = _read.table_cells(hit["elem"])
        row, col = int(m.group("row")) - 1, int(m.group("col")) - 1
        if not (0 <= row < len(cells)) or not cells or not (0 <= col < len(cells[0])):
            raise TargetNotFound(
                f"cell r{row + 1}c{col + 1} out of range; table is "
                f"{len(cells)}x{len(cells[0]) if cells else 0}"
            )
        out["kind"] = "cell"
        out["row"], out["col"] = row, col
    return out


# --------------------------------------------------------------- rendering


def _para_level(p) -> int:
    ppr = p.find(qn("a:pPr"))
    if ppr is None:
        return 0
    try:
        return int(ppr.get("lvl", "0"))
    except ValueError:
        return 0


def _shape_label(elem, kind: str) -> str:
    if kind == "placeholder":
        ph = _read._ph(elem)
        t = ph.get("type", "obj") if ph is not None else "obj"
        return {"ctrTitle": "title", "subTitle": "subtitle"}.get(t, t)
    return kind


def _text_lines(elem) -> list[str]:
    """Shape text as view lines: one paragraph inline, several as indented
    '-' bullet lines."""
    paras = _read.txbody_paragraphs(elem)
    texts = [_read.paragraph_text(p) for p in paras]
    if not any(texts):
        return []
    if len(texts) == 1:
        return [texts[0]]
    return [
        "  " * _para_level(p) + "- " + t
        for p, t in zip(paras, texts)
    ]


def _pipe_cell(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", " ")


def _render_table(elem, anchor: str) -> list[str]:
    cells = _read.table_cells(elem)
    rows = len(cells)
    cols = len(cells[0]) if cells else 0
    lines = [
        f"[a:{anchor}] table {rows}x{cols} "
        f"(cell anchors t:{anchor}:rNcN, 1-based; tool row/col args 0-based):"
    ]
    for row in cells:
        lines.append("| " + " | ".join(_pipe_cell(c) for c in row) + " |")
    return lines


def _geometry_note(elem) -> str:
    geo = _read._geometry(elem)
    if geo is None:
        return " @ (layout-inherited)"
    return (
        f" @ ({geo['x_in']:.2f}, {geo['y_in']:.2f}) "
        f"{geo['cx_in']:.2f}x{geo['cy_in']:.2f} in"
    )


def _slide_title(pkg: PptxPackage, part: str) -> str | None:
    return _read.slide_title(pkg, part)


# ------------------------------------------------------------------ public


def get_presentation_view(
    pkg: PptxPackage,
    scope=None,
    detail: str = "text",
    *,
    limit=None,
    offset: int = 0,
) -> dict:
    """The anchored markdown projection (module docstring has the scheme).
    scope: None = whole deck, or a slide selector / list of selectors.
    detail: "outline" | "text" (default) | "full".

    The projection is assembled slide by slide under the output budget
    (core/budget.py). A deck that fits renders whole, exactly as before. A
    deck that does not renders as many whole slides as fit and carries a
    `page` block naming the total, the count rendered, the count omitted,
    and the offset that continues; slides are never cut mid-way, so every
    anchor in the view is a complete, usable one. `anchor_count` counts the
    anchors actually rendered. `slide_count` is the deck's slide count, as
    in every other answer; a scoped view adds `slides_in_scope`."""
    if detail not in _DETAILS:
        raise PptMcpError(
            f"unknown detail {detail!r}; one of: {', '.join(_DETAILS)}"
        )
    slides = _read.slides_in_scope(pkg, scope)
    deck_count = len(slides) if scope is None else len(_read.slide_table(pkg))
    anchors = _display_anchors(_shape_map(pkg))

    header: list[str] = [
        f"# {pkg.path.name} ({len(slides)} slide"
        f"{'s' if len(slides) != 1 else ''}, detail={detail})"
    ]
    if detail != "outline":
        header.append(
            "Anchors: [s:N] slide id, [a:hex] shape, t:hex:rNcN table cell "
            "(rNcN is 1-based; row/col args to table tools are 0-based)."
        )
    if detail == "outline":
        header.append(_messages.VIEW_OUTLINE_OMISSIONS)
    elif detail == "text":
        header.append(_messages.VIEW_TEXT_OMISSIONS)

    blocks: list[str] = []
    block_anchors: list[int] = []

    for rec in slides:
        part = rec["part"]
        root = pkg.root(part)
        hidden = " (hidden)" if _read._slide_hidden(root) else ""
        lines: list[str] = [
            "",
            f"## Slide {rec['index'] + 1} [s:{rec['slide_id']}]{hidden}",
        ]
        anchor_count = 0

        if detail == "outline":
            title = _slide_title(pkg, part)
            if title:
                lines.append(title)
            blocks.append("\n".join(lines))
            block_anchors.append(anchor_count)
            continue

        _lp, layout_name = _read._layout_info(pkg, part)
        if layout_name:
            lines.append(f"Layout: {layout_name}")

        sp_tree = root.find(f"{qn('p:cSld')}/{qn('p:spTree')}")
        if sp_tree is not None:
            for elem, kind, _z, _parent in _read.iter_shapes(sp_tree, wrapped=True):
                cnvpr = _read._cnvpr(elem)
                if cnvpr is None:
                    continue
                sid = _read.cnvpr_id(cnvpr)
                anchor = anchors.get((rec["slide_id"], sid)) if sid is not None else None
                # A shape without an id (a damaged deck) still shows; it has no anchor.
                tag = f"[a:{anchor}]" if anchor else "[no id]"
                geo = _geometry_note(elem) if detail == "full" else ""
                if kind == "table":
                    lines.append("")
                    table_lines = _render_table(elem, anchor or "")
                    if not anchor:
                        table_lines[0] = "[no id] table:"
                    lines.extend(table_lines)
                    anchor_count += 1
                    continue
                if kind in _NO_TEXT_KINDS:
                    if detail == "full":
                        name = cnvpr.get("name", "")
                        label = f"{kind} \"{name}\"" if name else kind
                        lines.append(f"{tag} {label}{geo}")
                        anchor_count += 1
                    continue
                text_lines = _text_lines(elem)
                if not text_lines and detail != "full":
                    continue
                label = _shape_label(elem, kind)
                if len(text_lines) == 1:
                    lines.append(f"{tag} {label}{geo}: {text_lines[0]}")
                else:
                    lines.append(f"{tag} {label}{geo}:")
                    lines.extend(text_lines)
                anchor_count += 1

        notes = _read.notes_text(pkg, part)
        if notes:
            lines.append("Notes:")
            lines.extend("> " + ln for ln in notes.split("\n"))

        blocks.append("\n".join(lines))
        block_anchors.append(anchor_count)

    head = "\n".join(header)
    kept, page = _budget.page_blocks(
        blocks,
        limit=limit,
        offset=offset,
        overhead=_budget.overhead_of(
            {"view": head, "slide_count": len(slides), "detail": detail,
             "anchor_count": 0}
        ),
        unit="slides",
        narrow_hint="narrow the read with scope=<slide index or list of indexes>",
    )
    start = page["offset"] if page else 0
    rendered_anchors = sum(block_anchors[start : start + len(kept)])

    result = {
        "view": "\n".join([head, *kept]),
        "slide_count": deck_count,
        "detail": detail,
        "anchor_count": rendered_anchors,
    }
    if scope is not None:
        result = {
            "view": result.pop("view"), "slide_count": deck_count,
            "slides_in_scope": len(slides), **result,
        }
    if page is not None:
        result["page"] = page
    return result
