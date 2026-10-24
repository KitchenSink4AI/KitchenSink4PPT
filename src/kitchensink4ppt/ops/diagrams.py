"""SmartArt text: the write side of what read.py already sees.

Contract (all ops modules): every function takes the open PptxPackage first,
mutates only the in-memory package, calls pkg.mark_dirty() on every part it
touches, and returns a summary dict. Nothing here writes to disk.

SUBSTITUTION ONLY, and deliberately so. Authoring or relayouting a diagram
from a file is a guess: PowerPoint owns the layout algorithm and regenerates
geometry in-app, which is why full SmartArt authoring stays out of
scope in this codebase. Changing the WORDS in nodes that
already exist is a different operation, it needs no layout knowledge, and it
closes a real asymmetry: get_text, find_text, and diagram_text all read
SmartArt, so an agent that finds a term inside a diagram and cannot replace
it hits a wall the tool surface never admits to.

TWO STORAGE SITES, BOTH WRITTEN. A diagram keeps its text twice:

- `ppt/diagrams/dataN.xml` (dgm:dataModel/dgm:ptLst/dgm:pt/dgm:t) is the
  MODEL. PowerPoint re-reads it and regenerates everything else from it.
- `ppt/diagrams/drawingN.xml` (dsp:sp/dsp:txBody, related from the data part
  in older files or directly from the slide in current PowerPoint files) is
  the CACHED RENDERING, matched to the model by modelId. PowerPoint
  rebuilds it, but every other consumer draws it as-is: LibreOffice, the
  thumbnail, any viewer that never opens the deck in PowerPoint.

Writing the model alone leaves the old words on screen for all of those,
which is the same class of defect as the SVG dual blip. So both are written,
the drawing sync is counted in the result, and a diagram whose drawing part
is missing or whose modelIds do not line up says so in warnings rather than
reporting a clean success.

The caveat that always ships in the result: PowerPoint may still regenerate
the drawing from the model when the deck is opened, so the cached rendering
is a best effort and the model is the authority.
"""

from __future__ import annotations

from lxml import etree

from ..core import xmlsafe
from ..core.errors import PptMcpError, TargetNotFound, UnsupportedStructure
from ..core.package import PptxPackage, qn, resolve_target
from .read import (
    _diagram_data_part,
    _dgm,
    iter_shapes,
    resolve_slide,
)

DSP_NS = "http://schemas.microsoft.com/office/drawing/2008/diagram"
RT_DIAGRAM_DRAWING = (
    "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"
)

REGEN_NOTE = (
    "SmartArt text is substituted in the data model and in the cached "
    "drawing; PowerPoint may still regenerate the drawing from the model on "
    "open, and layout, connections, styles, and colors are never touched."
)


def _dsp(name: str) -> str:
    return f"{{{DSP_NS}}}{name}"


# ------------------------------------------------------------------ finding


def _drawing_targets(pkg: PptxPackage, source_part: str) -> list[str]:
    try:
        rels = pkg.rels_for(source_part)
    except KeyError:
        return []
    targets = []
    for rel in rels.getroot():
        if rel.get("Type") != RT_DIAGRAM_DRAWING or (
            rel.get("TargetMode") == "External"
        ):
            continue
        target = resolve_target(source_part, rel.get("Target", ""))
        if pkg.has_part(target):
            targets.append(target)
    return targets


def _model_ids(root: etree._Element) -> set[str]:
    return {
        node.get("modelId")
        for node in root.iter()
        if node.get("modelId")
    }


def drawing_part_of(
    pkg: PptxPackage,
    data_part: str,
    slide_part: str | None = None,
) -> str | None:
    """Resolve a SmartArt cached drawing in legacy and current packages.

    Older producers relate drawingN.xml from dataN.xml. Current PowerPoint
    commonly relates it from the slide instead. When a slide has several
    diagrams, model ids bind the data model to exactly one cached drawing;
    ambiguity refuses resolution rather than updating the wrong cache.
    """
    legacy = _drawing_targets(pkg, data_part)
    if len(legacy) == 1:
        return legacy[0]
    if slide_part is None:
        return None
    candidates = _drawing_targets(pkg, slide_part)
    if len(candidates) == 1:
        return candidates[0]
    if not candidates:
        return None
    wanted = _model_ids(pkg.root(data_part))
    matches = []
    for part in candidates:
        drawing_ids = _model_ids(pkg.root(part))
        if drawing_ids and drawing_ids <= wanted:
            matches.append(part)
    return matches[0] if len(matches) == 1 else None


def diagram_frames(pkg: PptxPackage, slide_part: str):
    """(frame element, data part) for every SmartArt frame on one slide."""
    sp_tree = pkg.root(slide_part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if sp_tree is None:
        return
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        if kind != "diagram":
            continue
        yield elem, _diagram_data_part(pkg, slide_part, elem)


def _text_nodes(root: etree._Element) -> list[tuple[str | None, etree._Element]]:
    """(modelId, dgm:t) for every text node of a data model (ordinary and
    assistant nodes, empty ones included), in document order: the same
    list read.diagram_text reports, so an index names the same node in
    both. The doc, presentation and connector points carry no user text."""
    from .read import DIAGRAM_TEXT_POINT_TYPES

    out = []
    for pt in root.iter(_dgm("pt")):
        t = pt.find(_dgm("t"))
        if t is None or pt.get("type", "node") not in DIAGRAM_TEXT_POINT_TYPES:
            continue
        out.append((pt.get("modelId"), t))
    return out


def _paragraphs_with_text(body: etree._Element) -> list[etree._Element]:
    return [
        p for p in body.findall(qn("a:p"))
        if p.find(qn("a:r")) is not None
    ]


def _set_body_text(body: etree._Element, value: str) -> tuple[bool, bool]:
    """Write `value` into a dgm:t or dsp:txBody, keeping the first run's
    rPr. Returns (changed, collapsed_runs).

    Substitution discipline: the first run of the first text paragraph keeps
    its character properties and takes the new text; any further runs in that
    paragraph go, because a caller supplying one string cannot say how a
    split should be distributed. Extra paragraphs are removed only when the
    new text has fewer lines than the old."""
    lines = str(value).split("\n")
    paras = _paragraphs_with_text(body)
    if not paras:
        paras = body.findall(qn("a:p"))
    if not paras:
        p = etree.SubElement(body, qn("a:p"))
        paras = [p]
    changed = False
    collapsed = False
    template = None  # an a:rPr to clone onto runs this call creates
    ppr_template = None  # the a:pPr to clone onto paragraphs it creates
    for i, line in enumerate(lines):
        if i < len(paras):
            p = paras[i]
            if ppr_template is None:
                ppr_template = p.find(qn("a:pPr"))
        else:
            p = etree.SubElement(body, qn("a:p"))
            if ppr_template is not None:
                p.append(xmlsafe.fromstring(etree.tostring(ppr_template)))
            changed = True
        for old_break in p.findall(qn("a:br")):
            p.remove(old_break)
            changed = True
        runs = p.findall(qn("a:r"))
        if runs:
            keep = runs[0]
            if template is None:
                rpr = keep.find(qn("a:rPr"))
                template = rpr if rpr is not None else None
            for extra in runs[1:]:
                p.remove(extra)
                changed = True
                collapsed = True
        else:
            # An empty node's paragraph holds only its end mark (and maybe
            # a pPr); a paragraph runs pPr, runs, endParaRPr, so the new
            # run goes before the mark, never after it.
            keep = etree.Element(qn("a:r"))
            if template is not None:
                keep.append(xmlsafe.fromstring(etree.tostring(template)))
            etree.SubElement(keep, qn("a:t"))
            end = p.find(qn("a:endParaRPr"))
            if end is not None:
                end.addprevious(keep)
            else:
                p.append(keep)
            changed = True
        t = keep.find(qn("a:t"))
        if t is None:
            t = etree.SubElement(keep, qn("a:t"))
        chunks = line.split("\v")
        if (t.text or "") != chunks[0]:
            t.text = chunks[0]
            changed = True
        cursor = keep
        for chunk in chunks[1:]:
            br = etree.Element(qn("a:br"))
            if template is not None:
                br.append(xmlsafe.fromstring(etree.tostring(template)))
            cursor.addnext(br)
            cursor = br
            changed = True
            if chunk:
                run = etree.Element(qn("a:r"))
                if template is not None:
                    run.append(xmlsafe.fromstring(etree.tostring(template)))
                etree.SubElement(run, qn("a:t")).text = chunk
                cursor.addnext(run)
                cursor = run
    for stale in paras[len(lines):]:
        body.remove(stale)
        changed = True
    return changed, collapsed


def _drawing_bodies(
    pkg: PptxPackage,
    drawing_part: str,
    data_root: etree._Element | None = None,
) -> dict[str, etree._Element]:
    """Data-model id -> cached drawing text body.

    Simplified/legacy drawings put the data id directly on ``dsp:sp``.
    PowerPoint uses a presentation-point id there; that ``dgm:pt`` carries
    ``prSet/@presAssocID`` back to the editable data point. Expose both keys
    so substitution reaches either package shape.
    """
    out: dict[str, etree._Element] = {}
    associations: dict[str, str] = {}
    if data_root is not None:
        for point in data_root.iter(_dgm("pt")):
            model_id = point.get("modelId")
            props = point.find(_dgm("prSet"))
            associated = props.get("presAssocID") if props is not None else None
            if model_id and associated:
                associations[model_id] = associated
    root = pkg.root(drawing_part)
    for sp in root.iter(_dsp("sp")):
        model_id = sp.get("modelId")
        body = sp.find(_dsp("txBody"))
        if model_id and body is not None:
            out[model_id] = body
            associated = associations.get(model_id)
            if associated:
                previous = out.get(associated)
                has_text = any((t.text or "") for t in body.iter(qn("a:t")))
                if previous is None or has_text:
                    out[associated] = body
    return out


# =============================================================== public API


def set_diagram_text(pkg: PptxPackage, slide, shape: int, nodes: list) -> dict:
    """Replace the TEXT of existing SmartArt nodes. `nodes` is a list of
    {"model_id": ...} or {"index": N} (diagram_text reports both, in the
    same order) plus "text"; '\\n' splits paragraphs within that node and
    '\\v' inserts a soft line break inside a paragraph. Layout, geometry,
    connections, styles, and colors are never
    touched, and no node is created or removed: this is substitution into a
    diagram PowerPoint already laid out.

    Both storage sites are written, the data model and the cached drawing,
    so the new words show up in viewers that never regenerate the diagram.
    Every node is resolved before anything is written, so an unknown node
    refuses without a partial edit."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    if not isinstance(nodes, list) or not nodes:
        raise PptMcpError(
            'nodes must be a non-empty list of {"index" or "model_id", '
            '"text"} dicts (diagram_text reports both)'
        )

    target_frame = None
    data_part = None
    for frame, dp in diagram_frames(pkg, part):
        cnvpr = frame.find(f"{qn('p:nvGraphicFramePr')}/{qn('p:cNvPr')}")
        if cnvpr is not None and cnvpr.get("id") == str(shape):
            target_frame, data_part = frame, dp
            break
    if target_frame is None:
        raise UnsupportedStructure(
            f"shape {shape} on slide {rec['index']} is not a SmartArt "
            "(diagram) frame; list_elements kind='shapes' and diagram_text "
            "show which shapes are diagrams"
        )
    if data_part is None or not pkg.has_part(data_part):
        raise UnsupportedStructure(
            f"SmartArt shape {shape} has no resolvable diagram data part; "
            "its text is unreachable and this refuses to guess"
        )

    root = pkg.root(data_part)
    text_nodes = _text_nodes(root)
    by_id = {mid: body for mid, body in text_nodes if mid}

    resolved: list[tuple[str | None, etree._Element, str]] = []
    for i, spec in enumerate(nodes):
        if not isinstance(spec, dict) or "text" not in spec:
            raise PptMcpError(f"nodes[{i}] must be a dict carrying 'text'")
        if not isinstance(spec["text"], str):
            raise PptMcpError(f"nodes[{i}]['text'] must be a string")
        model_id = spec.get("model_id")
        index = spec.get("index")
        if (model_id is None) == (index is None):
            raise PptMcpError(
                f"nodes[{i}] needs exactly one of 'model_id' or 'index'"
            )
        if model_id is not None:
            if model_id not in by_id:
                raise TargetNotFound(
                    f"nodes[{i}]: no diagram node with model_id "
                    f"{model_id!r}; diagram_text lists the ids this "
                    "diagram has"
                )
            resolved.append((model_id, by_id[model_id], spec["text"]))
        else:
            if not isinstance(index, int) or not 0 <= index < len(text_nodes):
                raise TargetNotFound(
                    f"nodes[{i}]: index {index} is outside this diagram's "
                    f"{len(text_nodes)} text nodes (0-based, diagram_text "
                    "order)"
                )
            mid, body = text_nodes[index]
            resolved.append((mid, body, spec["text"]))

    warnings: list[str] = []
    changed_nodes = 0
    collapsed = 0
    for _mid, body, value in resolved:
        did, coll = _set_body_text(body, value)
        if did:
            changed_nodes += 1
        if coll:
            collapsed += 1
    if changed_nodes:
        pkg.mark_dirty(data_part)
    if collapsed:
        warnings.append(
            f"{collapsed} node(s) held their text in several runs; the "
            "replacement is one run carrying the first run's formatting"
        )

    drawing_part = drawing_part_of(pkg, data_part, part)
    synced = 0
    if drawing_part is None:
        warnings.append(
            "this diagram has no cached drawing part, so only the data model "
            "was written; PowerPoint rebuilds the rendering on open, other "
            "viewers may show nothing until it does"
        )
    else:
        bodies = _drawing_bodies(pkg, drawing_part, root)
        missed = []
        for mid, _body, value in resolved:
            target = bodies.get(mid) if mid else None
            if target is None:
                missed.append(mid)
                continue
            did, _coll = _set_body_text(target, value)
            if did:
                synced += 1
        if synced:
            pkg.mark_dirty(drawing_part)
        if missed:
            warnings.append(
                f"{len(missed)} node(s) have no matching shape in the cached "
                "drawing, so those words change only in the data model until "
                "PowerPoint regenerates the diagram"
            )

    result = {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": shape,
        "data_part": data_part,
        "drawing_part": drawing_part,
        "nodes_changed": changed_nodes,
        "nodes_requested": len(resolved),
        "drawing_synced": synced,
        "note": REGEN_NOTE,
    }
    if warnings:
        result["warnings"] = warnings
    return result
