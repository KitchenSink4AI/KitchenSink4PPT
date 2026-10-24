"""apply_edits: the batch layer. Many edits, one lock, one backup, one save.

Contract: every anchor and location in the batch is
resolved FIRST, against the presentation's current state, before ANYTHING
mutates. Any resolution failure refuses the whole batch (STALE_ANCHOR when a
view anchor went stale, BAD_PARAMS otherwise) listing every failed op index,
so a refused batch leaves the file untouched. Only then are the edits applied
in order; an apply-time failure raises, the server's _edit wrapper never
saves, and the file is again untouched. `changed` maps op index to each op's
result so a follow-up batch chains without re-reading the deck.

Supported ops are a curated allowlist, not the whole tool surface: batch mode
multiplies mistakes, so structural surgery (slides, merges, groups) stays out
and goes through its own tools. Each edit dict: {"op": <name>, ...params},
addressed by a view "anchor" from get_presentation_view or by explicit
{"slide", "shape"/"table"/"placeholder"} keys, matching the standalone tools;
"slide_number" (counting from 1) or "slide_id" may stand in for "slide"
(counting from 0), and each edit's result carries the slide_number of the
slide it changed. A set_notes edit must name its slide by slide_id or
slide_number: a deck-wide notes pass addressed by 0-based index lands one
slide off in silence, so a bare index is refused there.

Sequencing caveat, by design: edits apply in list order against the live
package, so an edit that targets a shape a PRIOR edit in the same batch
deleted fails at apply time (whole batch aborted, nothing saved). Resolution
guarantees targets exist at batch START, not that the batch is internally
consistent; keep deletes last or in their own batch.

Paragraph edits (set_paragraph, insert_paragraph, delete_paragraph) are the
exception: each binds the paragraph it names when the batch is resolved, so
every paragraph address in one batch counts the shape's paragraphs as they
were before the batch, and deleting paragraphs 0 and 2 deletes those two.
Combinations whose addresses would contradict each other are refused at
resolution: set_text (which replaces every paragraph) beside a paragraph
edit on the same shape, format_text by paragraph beside an insert or delete
on the same shape (format_text counts paragraphs as earlier edits left
them), two edits changing one paragraph, and deletes that would leave a
shape no paragraph.
"""

from __future__ import annotations

from copy import deepcopy as _deepcopy

from lxml import etree

from ..core import messages as _messages
from ..core import slide_echo as _slide_echo
from ..core.errors import PptMcpError, TargetNotFound, UnsupportedStructure
from ..core.package import PptxPackage, qn
from ._runmap import build_map, remove_children, replace_range
from . import notes as _notes
from . import read as _read
from . import shapes as _shapes
from . import tables as _tables
from . import text as _text
from . import view as _view

# op -> (accepted param keys beyond op/location, required param keys)
_OPS: dict[str, tuple[set[str], set[str]]] = {
    "set_text": ({"text"}, {"text"}),
    "set_shape": (
        {"x", "y", "dx", "dy", "w", "h", "rotation", "flip_h", "flip_v",
         "fill", "line", "effect", "text", "text_style", "name"},
        set(),
    ),
    # format_text's `anchor` (vertical text anchor) rides here as
    # `text_anchor`: a bare "anchor" key in an edit is already the view
    # anchor that ADDRESSES the shape, and one word cannot mean both the
    # location and a property being written.
    "format_text": (
        {"paragraph", "paragraph_number", "start", "end", "font", "size_pt", "bold", "italic",
         "underline", "color", "align", "line_spacing", "space_before",
         "space_after", "text_anchor",
         "wrap"},
        set(),
    ),
    "delete_shape": (set(), set()),
    "set_table_cells": ({"cells"}, {"cells"}),
    "set_placeholder_text": (
        {"placeholder", "text", "paragraphs"}, {"placeholder"}
    ),
    "search_and_replace": (
        {"find", "replace", "scope", "regex", "match_case", "include_notes",
         "remove_empty_paragraphs"},
        {"find", "replace"},
    ),
    # One slide's speaker notes, replaced whole as set_notes replaces them:
    # exactly one of text or paragraphs (checked in _resolve_notes).
    "set_notes": ({"text", "paragraphs", "flatten"}, set()),
    # One paragraph of a shape, named by paragraph (counted from 0, as
    # find_text and format_text count) or paragraph_number (from 1), as the
    # shape was before the batch.
    "set_paragraph": ({"paragraph", "paragraph_number", "text"}, {"text"}),
    "insert_paragraph": ({"paragraph", "paragraph_number", "text"}, {"text"}),
    "delete_paragraph": ({"paragraph", "paragraph_number"}, set()),
}
#: The paragraph edits, which bind their paragraph at resolution.
_PARAGRAPH_OPS = ("set_paragraph", "insert_paragraph", "delete_paragraph")
_LOCATION_KEYS = {"anchor", "slide", "slide_id", "shape", "table"}
NOTES_SLIDE_RULE = (
    "a set_notes edit names its slide by slide_id (as get_presentation_info "
    "lists them) or slide_number (counted from 1), never by a 0-based slide "
    'index: send {"slide_id": 257, ...} or {"slide_number": 4, ...}'
)
#: format_text's vertical-anchor vocabulary. Seeing one of these under the
#: `anchor` key means the caller wanted text_anchor, not a view address.
_TEXT_ANCHOR_WORDS = {"top", "middle", "center", "bottom"}
# ops that address one shape (anchor or slide+shape)
_SHAPE_OPS = {"set_text", "set_shape", "format_text", "delete_shape", *_PARAGRAPH_OPS}

_PARAGRAPH_ADDRESS = (
    "anchor, or slide (or slide_number) with shape, plus paragraph counted from 0 as "
    "find_text reports it (or paragraph_number counted from 1, not both). Every "
    "paragraph address in one batch counts the shape's paragraphs as they were before "
    "the batch, so several edits to one shape need no renumbering")
#: How each batchable op names what it edits, as ppt_help apply_edits says
#: it; _resolve_one below is what reads these keys.
_ADDRESS: dict[str, str] = {
    "set_text": (
        "anchor (a shape anchor, or a table cell's t:hex:rNcN), or slide "
        "(or slide_number) with shape"),
    "set_shape": "anchor, or slide (or slide_number) with shape",
    "format_text": "anchor, or slide (or slide_number) with shape",
    "delete_shape": "anchor, or slide (or slide_number) with shape",
    "set_table_cells": "anchor (the table's), or slide (or slide_number) with table",
    "set_placeholder_text": "slide (or slide_number) with placeholder",
    "search_and_replace": "nothing: the whole deck, or scope to narrow it",
    "set_notes": (
        'slide_id or slide_number (or slide as {"slide_id": N}); a bare 0-based '
        "slide is refused. The slide's whole notes are replaced, as set_notes does"),
    "set_paragraph": (
        _PARAGRAPH_ADDRESS + ". The paragraph keeps its own formatting: the new text takes "
        "its first run's look, and a '\\v' in the text is a line break"),
    "insert_paragraph": (
        _PARAGRAPH_ADDRESS + ". paragraph is where the new one goes (leave it out to add it "
        "last); it takes the formatting of the paragraph now in that place, or of the last "
        "one when it goes last"),
    "delete_paragraph": _PARAGRAPH_ADDRESS + ". A shape keeps at least one paragraph",
}
#: What a batchable op does that its fields do not say, for ppt_help.
_NOTE: dict[str, str] = {
    "set_text": (
        _messages.SET_TEXT_FORMATTING + " On a table cell anchor, set_text writes that "
        "cell as set_table_cells does."),
    "set_shape": "With text, it replaces the shape's text the way set_text does.",
}
#: One edit of each kind, as it goes into the edits list.
_EXAMPLE: dict[str, dict] = {
    "set_text": {"anchor": "a:3f9c2a", "text": "New text"},
    "set_shape": {"anchor": "a:3f9c2a", "x": 1.0, "y": 1.5},
    "format_text": {"anchor": "a:3f9c2a", "bold": True, "size_pt": 24},
    "delete_shape": {"anchor": "a:3f9c2a"},
    "set_table_cells": {
        "anchor": "a:7d01e4", "cells": [{"row": 0, "col": 0, "text": "Total"}]},
    "set_placeholder_text": {
        "slide_number": 2, "placeholder": "title", "text": "Results"},
    "search_and_replace": {"find": "draft", "replace": "final"},
    "set_notes": {"slide_number": 1, "text": "What to say on this slide."},
    "set_paragraph": {"anchor": "a:3f9c2a", "paragraph": 1, "text": "New wording"},
    "insert_paragraph": {"anchor": "a:3f9c2a", "paragraph": 1, "text": "A new line"},
    "delete_paragraph": {"anchor": "a:3f9c2a", "paragraph_number": 1},
}
if set(_ADDRESS) != set(_OPS) or set(_EXAMPLE) != set(_OPS):  # pragma: no cover
    raise RuntimeError("every batchable op needs an address and an example")


def batchable_ops() -> list[dict]:
    """Every op apply_edits batches, with its fields and how it names its
    target, generated from the table _resolve_one enforces, so ppt_help can
    never list an op the batch refuses or leave one out."""
    rows = []
    for name, (allowed, required) in _OPS.items():
        row = {
            "op": name,
            "fields": sorted(allowed),
            "required": sorted(required),
            "address": _ADDRESS.get(name),
            "example": {"op": name, **_EXAMPLE.get(name, {})},
        }
        if name in _NOTE:
            row["note"] = _NOTE[name]
        rows.append(row)
    return rows


def _resolve_one(pkg: PptxPackage, edit: dict, named: set | None = None) -> dict:
    """Validate one edit and resolve its target against current state.
    `named` is the set of keys the edit's slide was sent under (slide,
    slide_number, slide_id). Returns a target record consumed by
    _apply_one. Raises on any problem."""
    if not isinstance(edit, dict) or "op" not in edit:
        raise PptMcpError('each edit must be a dict with an "op" key')
    op = edit["op"]
    if op not in _OPS:
        raise PptMcpError(
            f"op {op!r} is not batchable; supported: {sorted(_OPS)}"
        )
    allowed, required = _OPS[op]
    extra = set(edit) - allowed - _LOCATION_KEYS - {"op"}
    if extra:
        raise PptMcpError(
            f"op {op!r} does not accept {sorted(extra)}; "
            f"allowed params: {sorted(allowed)}"
        )
    missing = required - set(edit)
    if missing:
        raise PptMcpError(f"op {op!r} is missing required {sorted(missing)}")

    if op == "search_and_replace":
        scope = edit.get("scope")
        if scope is not None:
            _read.slides_in_scope(pkg, scope)  # validates, result discarded
        return {"kind": "none"}

    if op == "set_notes":
        return _resolve_notes(pkg, edit, named or set())

    if op == "set_placeholder_text":
        if "slide" not in edit:
            raise PptMcpError('set_placeholder_text needs a "slide" key')
        rec = _read.resolve_slide(pkg, edit["slide"])
        return {"kind": "slide", "slide_index": rec["index"]}

    # Shape- and table-addressed ops: anchor or explicit keys.
    if "anchor" in edit:
        # An agent reads format_text(anchor="middle") and writes it here,
        # where `anchor` already means the view anchor that ADDRESSES the
        # shape. It refused correctly, with a message about anchor grammar
        # that named no way forward.
        if op == "format_text" and isinstance(edit["anchor"], str) and (
            edit["anchor"].strip().lower() in _TEXT_ANCHOR_WORDS
        ):
            raise PptMcpError(
                f'op \'format_text\' got anchor={edit["anchor"]!r}, which '
                'reads as the VIEW anchor that addresses a shape. The '
                'vertical text anchor rides as "text_anchor" here: '
                f'{{"op": "format_text", ..., "text_anchor": '
                f'{edit["anchor"]!r}}}'
            )
        info = _view.resolve_anchor(pkg, edit["anchor"])
        if info["kind"] == "slide":
            raise PptMcpError(
                f"op {op!r} needs a shape or cell anchor, got the slide "
                f"anchor {edit['anchor']!r}"
            )
        if edit.get("slide") is not None:
            # An anchor already names its slide. A second name that
            # disagrees used to be ignored, and the edit landed on the
            # anchor's slide while the caller meant another one.
            named = _read.resolve_slide(pkg, edit["slide"])
            if named["index"] != info["slide_index"]:
                raise PptMcpError(
                    f"anchor {edit['anchor']!r} is on slide_number "
                    f"{info['slide_index'] + 1}, but this edit also names "
                    f"slide_number {named['index'] + 1}. Send the anchor "
                    "alone, or the slide with a shape."
                )
        target = {
            "kind": info["kind"],  # "shape" or... cells carry row/col below
            "slide_index": info["slide_index"],
            "shape_id": info["shape_id"],
            "shape_type": info["shape_type"],
        }
        if "row" in info:  # table-cell anchor t:<hex>:rNcN
            target["kind"] = "cell"
            target["row"], target["col"] = info["row"], info["col"]
    elif "slide" in edit:
        rec = _read.resolve_slide(pkg, edit["slide"])
        sel = edit.get("table") if op == "set_table_cells" else edit.get("shape")
        if op == "set_table_cells":
            trec = _tables.resolve_table(pkg, edit["slide"], sel)
            target = {
                "kind": "shape",
                "slide_index": rec["index"],
                "shape_id": trec["shape_id"],
                "shape_type": "table",
            }
        else:
            if not isinstance(sel, int) or isinstance(sel, bool):
                raise PptMcpError(
                    f'op {op!r} needs "shape" (an int shape id) alongside '
                    '"slide", or a view "anchor"'
                )
            _shapes._find_shape(pkg, rec["part"], sel)  # existence check
            target = {
                "kind": "shape",
                "slide_index": rec["index"],
                "shape_id": sel,
                "shape_type": None,
            }
    else:
        raise PptMcpError(
            f'op {op!r} needs an "anchor" or a "slide" (+ "shape"/"table") '
            "location"
        )

    if target["kind"] == "cell" and op != "set_text":
        exc = PptMcpError(
            f"op {op!r} cannot target a table CELL anchor; cell anchors "
            "work with set_text (or use set_table_cells on the table)"
        )
        exc.hint_tools = ["set_table_cells"]
        raise exc
    if op == "set_table_cells" and target.get("shape_type") not in (
        "table", None
    ):
        raise PptMcpError(
            f"set_table_cells target is a {target['shape_type']}, not a table"
        )
    if op in _PARAGRAPH_OPS:
        target.update(_resolve_paragraph(pkg, edit, target))
    elif op == "format_text":
        index = _paragraph_index(edit)
        if index is not None:
            target["paragraph"] = index
    return target


def _paragraph_index(edit: dict) -> int | None:
    """The 0-based paragraph an edit names, from paragraph (counted from 0)
    or paragraph_number (counted from 1); None when it names none. Both
    sent is refused, as every 1-based twin is; a null is not sent."""
    index, number = edit.get("paragraph"), edit.get("paragraph_number")
    if index is not None and number is not None:
        raise PptMcpError(
            "paragraph and paragraph_number were both sent: send paragraph "
            "counted from 0 (as find_text reports it) or paragraph_number "
            "counted from 1, not both"
        )
    if number is not None:
        if isinstance(number, bool) or not isinstance(number, int) or number < 1:
            raise PptMcpError(
                "paragraph_number counts from 1: the first paragraph is "
                "paragraph_number 1"
            )
        return number - 1
    if index is not None:
        if isinstance(index, bool) or not isinstance(index, int) or index < 0:
            raise PptMcpError(
                "paragraph counts from 0, as find_text reports it: the first "
                "paragraph is paragraph 0"
            )
        return index
    return None


def _one_paragraph_text(text) -> str:
    if not isinstance(text, str):
        raise PptMcpError(f"text must be a string, got {type(text).__name__}")
    if "\n" in text:
        raise PptMcpError(
            "one paragraph's text cannot hold a paragraph break ('\\n'); send "
            "one insert_paragraph per new paragraph, or '\\v' for a line "
            "break inside this one"
        )
    return text


def _resolve_paragraph(pkg: PptxPackage, edit: dict, target: dict) -> dict:
    """Bind a paragraph edit to the paragraph it names, as the shape is now
    (before anything in the batch is written)."""
    op = edit["op"]
    index = _paragraph_index(edit)
    rec = _read.resolve_slide(pkg, target["slide_index"])
    elem, _chain = _shapes._find_shape(pkg, rec["part"], target["shape_id"])
    sid = target["shape_id"]
    if etree.QName(elem).localname != "sp":
        raise UnsupportedStructure(
            f"shape {sid} is a {_read._shape_kind(elem)}, which has no paragraphs "
            "of its own; paragraph edits take a shape that holds text"
        )
    body = elem.find(qn("p:txBody"))
    paras = body.findall(qn("a:p")) if body is not None else []
    count = len(paras)
    plural = "" if count == 1 else "s"
    bound = {"kind": "paragraph", "part": rec["part"], "slide_id": rec["slide_id"],
             "elem": elem, "body": body, "count": count}
    if op == "insert_paragraph":
        _one_paragraph_text(edit["text"])
        if index is None or index == count:
            return {**bound, "para": None, "index": count}
        if not 0 <= index < count:
            raise TargetNotFound(
                f"paragraph {index} is out of range for a new paragraph; shape {sid} "
                f"has {count} paragraph{plural}, so a new one goes at paragraph 0 to "
                f"{count} ({count} adds it last)"
            )
        return {**bound, "para": paras[index], "index": index}
    if index is None:
        raise PptMcpError(
            f"{op} needs paragraph (counted from 0, as find_text reports it) or "
            "paragraph_number (counted from 1)"
        )
    if not 0 <= index < count:
        raise TargetNotFound(
            f"paragraph {index} is out of range; shape {sid} has {count} "
            f"paragraph{plural}, so paragraph runs 0 to {count - 1} "
            f"(paragraph_number 1 to {count})"
            if count else f"shape {sid} has no text, so it has no paragraph {index}"
        )
    if op == "set_paragraph":
        _one_paragraph_text(edit["text"])
        _visible, segments = build_map(paras[index])
        if any(seg.kind == "fld" for seg in segments):
            raise UnsupportedStructure(
                f"paragraph {index} of shape {sid} holds a field (a slide number or "
                "a date) that PowerPoint fills in when it draws the slide, so its "
                "text cannot be rewritten; use search_and_replace on the words "
                "around it"
            )
        if any(seg.kind == "wrapped" for seg in segments):
            raise UnsupportedStructure(
                f"paragraph {index} of shape {sid} holds an equation or other "
                "content PowerPoint stores in a compatibility wrapper, which "
                "rewriting the paragraph would lose; use search_and_replace on "
                "the words around it, or edit it in PowerPoint"
            )
    return {**bound, "para": paras[index], "index": index}


def _paragraph_conflicts(edits: list, targets: list[dict], failures: list[dict]) -> None:
    """Refuse batches whose paragraph addresses would contradict each other
    (module docstring): every one is found before anything is written."""
    rewrites: dict = {}
    formats: dict = {}
    shifts: dict = {}
    users: dict = {}
    changed: dict = {}
    deleted: dict = {}
    inserted: dict = {}
    counts: dict = {}
    for i, (edit, target) in enumerate(zip(edits, targets)):
        if not target or not isinstance(edit, dict):
            continue
        op = edit.get("op")
        key = (target.get("slide_index"), target.get("shape_id"))
        if (op == "set_text" and target.get("kind") == "shape") or (
            op == "set_shape" and edit.get("text") is not None
        ):
            rewrites.setdefault(key, i)
        elif op == "format_text" and target.get("paragraph") is not None:
            formats.setdefault(key, i)
        elif op in _PARAGRAPH_OPS:
            users.setdefault(key, i)
            counts[key] = target["count"]
            if op != "insert_paragraph":
                where = (*key, target["index"])
                if where in changed:
                    failures.append({"index": i, "op": op, "error": (
                        f"edits {changed[where]} and {i} both change paragraph "
                        f"{target['index']} of shape {key[1]}; send one of them")})
                else:
                    changed[where] = i
            if op == "delete_paragraph":
                deleted.setdefault(key, {})[target["index"]] = i
            if op == "insert_paragraph":
                inserted[key] = inserted.get(key, 0) + 1
            if op != "set_paragraph":
                shifts.setdefault(key, i)
    for key, i in users.items():
        if key in rewrites:
            j = rewrites[key]
            failures.append({"index": max(i, j), "op": edits[max(i, j)].get("op"), "error": (
                f"edit {j} replaces every paragraph of shape {key[1]} (set_text), so the "
                f"paragraph address of edit {i} would no longer exist; send the paragraph "
                "edits in a second batch")})
    for key, i in shifts.items():
        if key in formats:
            j = formats[key]
            failures.append({"index": max(i, j), "op": edits[max(i, j)].get("op"), "error": (
                f"format_text (edit {j}) counts the paragraphs of shape {key[1]} as earlier "
                "edits leave them, while insert_paragraph and delete_paragraph (edit "
                f"{i}) count them as they were before the batch; send format_text in a "
                "second batch")})
    for key, gone in deleted.items():
        if counts[key] - len(gone) + inserted.get(key, 0) < 1:
            failures.append({"index": max(gone.values()), "op": "delete_paragraph", "error": (
                f"edits {sorted(gone.values())} would delete every paragraph of shape "
                f"{key[1]}; a text body keeps at least one, so empty one with "
                "set_paragraph and text \"\" or delete the shape")})


def _look(source):
    """A run's character properties copied from `source` (an a:rPr or
    a:endParaRPr), without its hyperlinks: new text never inherits a link."""
    if source is None:
        return None
    rpr = _deepcopy(source)
    rpr.tag = qn("a:rPr")
    remove_children(rpr, ("a:hlinkClick", "a:hlinkMouseOver"))
    return rpr


def _text_nodes(text: str, look) -> list:
    """a:r and a:br elements for `text`, each carrying a copy of `look`."""
    nodes = []
    chunks = text.split("\v")
    for n, chunk in enumerate(chunks):
        if chunk:
            run = etree.Element(qn("a:r"))
            if look is not None:
                run.append(_deepcopy(look))
            etree.SubElement(run, qn("a:t")).text = chunk
            nodes.append(run)
        if n < len(chunks) - 1:
            br = etree.Element(qn("a:br"))
            if look is not None:
                br.append(_deepcopy(look))
            nodes.append(br)
    return nodes


def _write_paragraph(para, text: str) -> None:
    """Replace one paragraph's text, keeping its paragraph properties and
    its first run's look; an empty paragraph takes its insertion style."""
    visible, segments = build_map(para)
    if segments:
        replace_range(para, segments, 0, len(visible), text)
        return
    end = para.find(qn("a:endParaRPr"))
    for node in _text_nodes(text, _look(end)):
        if end is not None:
            end.addprevious(node)
        else:
            para.append(node)


def _new_paragraph(template, text: str):
    """A paragraph with `template`'s paragraph properties and first-run look."""
    para = etree.Element(qn("a:p"))
    if template is None:
        for node in _text_nodes(text, None):
            para.append(node)
        return para
    ppr = template.find(qn("a:pPr"))
    if ppr is not None:
        para.append(_deepcopy(ppr))
    end = template.find(qn("a:endParaRPr"))
    first = _shapes._first_run_rpr(template)
    for node in _text_nodes(text, _look(first if first is not None else end)):
        para.append(node)
    if end is not None:
        tail = _deepcopy(end)
        remove_children(tail, ("a:hlinkClick", "a:hlinkMouseOver"))
        para.append(tail)
    return para


def _apply_paragraph(pkg: PptxPackage, edit: dict, target: dict) -> dict:
    op = edit["op"]
    body, para, elem = target["body"], target["para"], target["elem"]
    sid, index = target["shape_id"], target["index"]
    gone = PptMcpError(
        f"paragraph {index} of shape {sid} was removed or replaced by an earlier "
        "edit in this batch; send this edit in a batch of its own"
    )
    if body is not None and body.getparent() is not elem:
        raise gone
    if para is not None and para.getparent() is not body:
        raise gone
    if op == "set_paragraph":
        _write_paragraph(para, edit["text"])
    elif op == "delete_paragraph":
        if len(body.findall(qn("a:p"))) <= 1:
            raise PptMcpError(
                f"paragraph {index} is the only paragraph left in shape {sid}; a "
                "text body keeps at least one, so empty it with set_paragraph and "
                'text "" or delete the shape'
            )
        body.remove(para)
    else:
        if body is None:
            body = _text._require_txbody(
                elem, "autoshape", {"index": target["slide_index"]}, create=True)
            _write_paragraph(body.find(qn("a:p")), edit["text"])
        elif para is not None:
            para.addprevious(_new_paragraph(para, edit["text"]))
        else:
            paras = body.findall(qn("a:p"))
            last = paras[-1] if paras else None
            new = _new_paragraph(last, edit["text"])
            if last is not None:
                last.addnext(new)
            else:
                body.append(new)
    pkg.mark_dirty(target["part"])
    return {
        "slide_index": target["slide_index"],
        "slide_id": target["slide_id"],
        "shape_id": sid,
        "paragraph": index,
        "paragraphs": len(body.findall(qn("a:p"))),
    }


def _resolve_notes(pkg: PptxPackage, edit: dict, named: set) -> dict:
    """A set_notes edit: what it writes is checked here, before anything
    in the batch is written, and its slide is named by slide_id or
    slide_number."""
    for key in ("anchor", "shape", "table"):
        if key in edit:
            raise PptMcpError(
                f"a set_notes edit names a slide, not a {key}: " + NOTES_SLIDE_RULE)
    text, paragraphs = edit.get("text"), edit.get("paragraphs")
    if (text is None) == (paragraphs is None):
        raise PptMcpError(
            "a set_notes edit takes exactly one of text (plain, '\\n' between "
            "paragraphs) or paragraphs (the structure get_notes returns)")
    if text is not None and not isinstance(text, str):
        raise PptMcpError(f"text must be a string, got {type(text).__name__}")
    if paragraphs is not None:
        _notes._normalize_paragraphs(paragraphs)  # refuses a bad shape now
    flatten = edit.get("flatten", False)
    if not isinstance(flatten, bool):
        raise PptMcpError("flatten must be true or false")
    selector = edit.get("slide")
    if selector is None:
        raise PptMcpError(NOTES_SLIDE_RULE)
    if not (named & {"slide_id", "slide_number"}) and not isinstance(selector, dict):
        raise PptMcpError(NOTES_SLIDE_RULE)
    rec = _read.resolve_slide(pkg, selector)
    return {"kind": "notes", "slide_index": rec["index"], "slide_id": rec["slide_id"]}


def _apply_one(pkg: PptxPackage, edit: dict, target: dict) -> dict:
    op = edit["op"]
    params = {
        k: v for k, v in edit.items() if k not in _LOCATION_KEYS and k != "op"
    }

    if op == "search_and_replace":
        find = params.pop("find")
        replace = params.pop("replace")
        return _text.search_and_replace(pkg, find, replace, **params)

    if op == "set_placeholder_text":
        placeholder = params.pop("placeholder")
        return _text.set_placeholder_text(
            pkg, target["slide_index"], placeholder, **params
        )

    if op == "set_notes":
        return _notes.set_notes(
            pkg, target["slide_index"], params.get("text"),
            paragraphs=params.get("paragraphs"),
            flatten=params.get("flatten", False),
        )

    slide = target["slide_index"]
    if op == "set_text":
        if target["kind"] == "cell":
            return _tables.set_table_cells(
                pkg, slide, {"shape_id": target["shape_id"]},
                [{"row": target["row"], "col": target["col"],
                  "text": edit["text"]}],
            )
        return _shapes.set_shape(
            pkg, slide, target["shape_id"], text=edit["text"]
        )
    if op == "set_shape":
        return _shapes.set_shape(pkg, slide, target["shape_id"], **params)
    if op == "format_text":
        if "text_anchor" in params:
            params["anchor"] = params.pop("text_anchor")
        params.pop("paragraph_number", None)
        if target.get("paragraph") is not None:
            params["paragraph"] = target["paragraph"]
        return _text.format_text(pkg, slide, target["shape_id"], **params)
    if op in _PARAGRAPH_OPS:
        return _apply_paragraph(pkg, edit, target)
    if op == "delete_shape":
        return _shapes.delete_shape(pkg, slide, target["shape_id"])
    if op == "set_table_cells":
        return _tables.set_table_cells(
            pkg, slide, {"shape_id": target["shape_id"]}, edit["cells"]
        )
    raise PptMcpError(f"unhandled op {op!r}")  # unreachable; _OPS gates entry


#: Other names an edit may use, as the standalone operations read them:
#: shape_id for shape on the edits that address a shape, font_size for
#: format_text's size_pt. set_table_cells addresses its table by table, so
#: a shape_id there stays unknown and is refused, as the standalone
#: operation refuses it.
_EDIT_SYNONYMS: dict[str, dict[str, str]] = {
    **{op: {"shape_id": "shape"} for op in _SHAPE_OPS},
    "format_text": {"shape_id": "shape", "font_size": "size_pt"},
}


def _edit_synonyms(edit):
    """Read an edit's other names as the names it takes. Two names for one
    value that disagree refuse rather than pick one; a null is not sent."""
    if not isinstance(edit, dict) or not isinstance(edit.get("op"), str):
        return edit
    names = _EDIT_SYNONYMS.get(edit["op"], {})
    if not any(name in edit for name in names):
        return edit
    out = dict(edit)
    for synonym, canonical in names.items():
        if synonym not in out:
            continue
        theirs = out.pop(synonym)
        ours = out.get(canonical)
        if theirs is None:
            continue
        if ours is not None and ours != theirs:
            raise PptMcpError(
                f"{synonym} and {canonical} were both sent, with different values; "
                f"an edit reads {synonym} as another name for {canonical}, so send "
                "one of them"
            )
        out[canonical] = theirs
    return out


def _slide_address_edit(pkg: PptxPackage, edit) -> tuple:
    """Accept slide_number (counting from 1) and slide_id wherever an edit
    takes slide, as slide. Returns the edit and the set of keys its slide
    was sent under. Names that disagree refuse rather than pick one; a null
    name is treated as not sent, like any optional key."""
    if not isinstance(edit, dict):
        return edit, set()
    out = dict(edit)
    named: list[tuple[str, object]] = []
    if out.get("slide") is not None:
        named.append(("slide", out["slide"]))
    else:
        out.pop("slide", None)
    if "slide_number" in out:
        number = out.pop("slide_number")
        if number is not None:
            if isinstance(number, bool) or not isinstance(number, int) or number < 1:
                raise PptMcpError(
                    "slide_number counts from 1: the first slide is slide_number 1"
                )
            named.append(("slide_number", number - 1))
    if "slide_id" in out:
        sid = out.pop("slide_id")
        if sid is not None:
            if isinstance(sid, bool) or not isinstance(sid, int) or sid < 0:
                raise PptMcpError(
                    "slide_id is a slide's whole-number id, as "
                    "get_presentation_info lists them"
                )
            named.append(("slide_id", {"slide_id": sid}))
    if not named:
        return out, set()
    first_name, first = named[0]
    for name, value in named[1:]:
        if value == first:
            continue
        if {first_name, name} == {"slide", "slide_number"} and all(
            isinstance(v, int) and not isinstance(v, bool) for v in (first, value)
        ):
            raise PptMcpError(
                "an edit's slide and slide_number name different slides; slide "
                "counts from 0 and slide_number from 1, so the second slide is "
                "slide 1 or slide_number 2. Send one of them."
            )
        a = _read.resolve_slide(pkg, first)["index"]
        b = _read.resolve_slide(pkg, value)["index"]
        if a != b:
            raise PptMcpError(
                f"an edit's {first_name} and {name} name different slides "
                f"(slide_number {a + 1} and slide_number {b + 1}). Send one "
                "of them."
            )
    # A slide_id names the slide durably; keep it when one was sent.
    out["slide"] = next((v for n, v in named if n == "slide_id"), first)
    return out, {name for name, _v in named}


def _numbered(result):
    """Put the 1-based slide_number beside a per-edit slide_index."""
    index = result.get("slide_index") if isinstance(result, dict) else None
    if (
        isinstance(index, int) and not isinstance(index, bool)
        and "slide_number" not in result
    ):
        return {"slide_number": index + 1, **result}
    return result


def set_notes_list(pkg: PptxPackage, notes, *, flatten: bool = False) -> dict:
    """set_notes(notes=[...]): every entry is a set_notes edit, so every
    slide is resolved before anything is written and the whole list is one
    save. An entry's own flatten wins over the call's."""
    if not isinstance(notes, list) or not notes:
        raise PptMcpError(
            "notes must be a non-empty list of {slide_id or slide_number, "
            "text or paragraphs}")
    edits = []
    for i, entry in enumerate(notes):
        if not isinstance(entry, dict):
            raise PptMcpError(
                f'notes[{i}] must be an object such as {{"slide_id": 257, '
                '"text": "..."}')
        unknown = sorted(set(entry) - {
            "slide_id", "slide_number", "slide", "text", "paragraphs", "flatten"})
        if unknown:
            raise PptMcpError(
                f"notes[{i}] does not take {unknown}; each entry takes slide_id "
                "or slide_number, and text or paragraphs")
        edits.append({"op": "set_notes", "flatten": flatten, **entry})
    out = apply_edits(pkg, edits)
    return {
        "applied": out["applied"],
        "notes": [out["changed"][str(i)] for i in range(len(edits))],
        "warnings": out["warnings"],
    }


def apply_edits(pkg: PptxPackage, edits: list[dict], atomic: bool = True) -> dict:
    if atomic is not True:
        raise PptMcpError(_messages.BATCH_ATOMIC_ONLY)
    if not isinstance(edits, list) or not edits:
        raise PptMcpError('edits must be a non-empty list of {"op": ...} dicts')
    # Step 1: resolve EVERYTHING before anything mutates.
    prepared: list = []
    targets: list[dict] = []
    failures: list[dict] = []
    hint_tools: list[str] = []
    stale = False
    not_found = False
    for i, raw in enumerate(edits):
        edit = raw
        try:
            edit, named = _slide_address_edit(pkg, _edit_synonyms(raw))
            with _slide_echo.numbered("slide_number" in named and "slide_id" not in named):
                targets.append(_resolve_one(pkg, edit, named))
        except TargetNotFound as exc:
            # STALE_ANCHOR only when a view anchor was actually in play; an
            # explicit slide+shape miss is a plain wrong address, and the
            # stale-anchor hint would send the caller on a futile re-view
            # loop.
            if isinstance(edit, dict) and "anchor" in edit:
                stale = True
            else:
                not_found = True
            failures.append(
                {"index": i,
                 "op": edit.get("op") if isinstance(edit, dict) else None,
                 "error": str(exc)}
            )
            for tool in getattr(exc, "hint_tools", None) or []:
                if tool not in hint_tools:
                    hint_tools.append(tool)
            targets.append({})
        except Exception as exc:
            failures.append(
                {"index": i,
                 "op": edit.get("op") if isinstance(edit, dict) else None,
                 "error": f"{type(exc).__name__}: {exc}"}
            )
            # Per-edit hint_tools must survive the aggregation, or every
            # instrumented raise site under apply_edits is dead code on the
            # wire (discoverability round, contract point 3).
            for tool in getattr(exc, "hint_tools", None) or []:
                if tool not in hint_tools:
                    hint_tools.append(tool)
            targets.append({})
        prepared.append(edit)
    # Two notes edits for one slide: the second would replace the first in
    # silence, which is an addressing mistake far more often than intent.
    notes_at: dict[int, int] = {}
    for i, target in enumerate(targets):
        if target.get("kind") != "notes":
            continue
        index = target["slide_index"]
        if index in notes_at:
            failures.append({
                "index": i, "op": "set_notes",
                "error": (f"edits {notes_at[index]} and {i} both set the notes "
                          f"of slide_number {index + 1}; send one of them"),
            })
        else:
            notes_at[index] = i
    _paragraph_conflicts(prepared, targets, failures)
    if failures:
        err = PptMcpError(
            f"batch refused, nothing was applied: {len(failures)} of "
            f"{len(edits)} edits failed resolution. Failures: {failures}. "
            "Fix or drop the failed edits (re-run get_presentation_view if "
            "anchors went stale) and resend the batch."
        )
        if stale:
            err.code = "STALE_ANCHOR"
        elif not_found:
            err.code = "NOT_FOUND"
        else:
            err.code = "BAD_PARAMS"
        err.detail = {"failures": failures}
        if hint_tools:
            err.hint_tools = hint_tools
        raise err

    # Step 2: apply in order. A failure here raises; the caller's _edit
    # wrapper never saves, so the on-disk file stays untouched.
    changed: dict[str, dict] = {}
    warnings: list[str] = []
    for i, (edit, target) in enumerate(zip(prepared, targets)):
        try:
            result = _apply_one(pkg, edit, target)
        except PptMcpError as exc:
            err = PptMcpError(
                f"edit {i} ({edit['op']}) failed during apply; the batch "
                f"was ABANDONED and the file is unchanged: {exc}"
            )
            # Keep the raise site's hint_tools alive across the re-wrap
            # (e.g. a batched format_text on a table cell points at
            # set_table_cells; discoverability round, contract point 3).
            tools = getattr(exc, "hint_tools", None)
            if tools:
                err.hint_tools = list(tools)
            raise err from exc
        if isinstance(result, dict):
            warnings.extend(result.pop("warnings", []) or [])
            # An entry's results sit under changed; its case note is lifted
            # where the caller reads cautions (search_and_replace's case_note).
            case_note = result.pop("case_note", None)
            if case_note:
                warnings.append(f"edit {i}: {case_note}")
        changed[str(i)] = _numbered(result)

    return {"applied": len(edits), "changed": changed, "warnings": warnings}
