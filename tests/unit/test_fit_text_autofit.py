"""fit_text fits text through ONE mechanism and keeps the shape's own autofit
setting.

Field report: fit_text on a text box set to "resize shape to fit text"
(a:spAutoFit) scaled every run to 61% AND replaced the setting with "shrink
text on overflow" (a:normAutofit). PowerPoint then shrank the text a second
time when it drew the slide (to about 8 pt), while the answer said the text
fitted at 10 to 14.6 pt.

The rule now, for every autofit state:
- "shrink text on overflow" (normAutofit, set on the shape or inherited from
  the layout or master): the fit is written as that setting's own font scale
  and the run sizes are left alone;
- "resize shape to fit text" (spAutoFit): the shape grows with its text, so
  the text is fitted to the room the shape can grow into on the slide, by its
  run sizes, and the setting stays;
- no autofit (noAutofit, or nothing set anywhere): the run sizes are
  rewritten and nothing is added.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import shapes as shp
from kitchensink4ppt.ops import slides as sl
from kitchensink4ppt.ops import text as tx
from kitchensink4ppt.ops.read import get_slide_info

_NS = (
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
)
_EMU = 914400
_AUTOFIT = ("normAutofit", "spAutoFit", "noAutofit")

_LONG = (
    "A paragraph with a good number of words in it, long enough that it "
    "wraps across the width of the box more than once at this size, and then "
    "keeps going for a while longer."
)


@pytest.fixture()
def deck(make_deck):
    """(pkg, slide index of a new Blank slide); the deck is 10 x 7.5 in."""
    pkg = PptxPackage(make_deck("fit_autofit.pptx", extra_slides=0))
    slide = sl.insert_slide(pkg, "Blank")["index"]
    return pkg, slide


def _tree(pkg, slide):
    part = get_slide_info(pkg, slide)["part"]
    return part, pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")


def _add_box(pkg, slide, *, autofit, y_in, h_in, sizes=(24, 20, 16, 16),
             text=_LONG, sid=77, w_in=9.0, wrap="square"):
    """A text box written the way other tools write them, with the given
    autofit child ("spAutoFit", "noAutofit", "normAutofit", a full element
    string, or None for none at all) and one paragraph per size."""
    part, tree = _tree(pkg, slide)
    if autofit is None:
        fit = ""
    elif autofit.startswith("<"):
        fit = autofit
    else:
        fit = f"<a:{autofit}/>"
    paras = "".join(
        f'<a:p><a:r><a:rPr lang="en-US" sz="{int(size * 100)}"/>'
        f"<a:t>{text}</a:t></a:r></a:p>"
        for size in sizes
    )
    xml = (
        f'<p:sp {_NS}><p:nvSpPr><p:cNvPr id="{sid}" name="Box {sid}"/>'
        '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{int(0.5 * _EMU)}" y="{int(y_in * _EMU)}"/>'
        f'<a:ext cx="{int(w_in * _EMU)}" cy="{int(h_in * _EMU)}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody><a:bodyPr wrap="{wrap}" rtlCol="0">{fit}</a:bodyPr>'
        f"<a:lstStyle/>{paras}</p:txBody></p:sp>"
    )
    tree.append(etree.fromstring(xml))
    pkg.mark_dirty(part)
    return sid


def _elem(pkg, slide, sid):
    part, _tree_ = _tree(pkg, slide)
    elem, _chain = shp._find_shape(pkg, part, sid)
    return elem


def _bodypr(pkg, slide, sid):
    return _elem(pkg, slide, sid).find(f"{qn('p:txBody')}/{qn('a:bodyPr')}")


def _autofit_children(pkg, slide, sid):
    return [
        etree.QName(child).localname
        for child in _bodypr(pkg, slide, sid)
        if etree.QName(child).localname in _AUTOFIT
    ]


def _run_sizes(pkg, slide, sid):
    body = _elem(pkg, slide, sid).find(qn("p:txBody"))
    return [r.get("sz") for r in body.iter(qn("a:rPr"))]


def _box_bottom_emu(pkg, slide, sid):
    xfrm = _elem(pkg, slide, sid).find(f"{qn('p:spPr')}/{qn('a:xfrm')}")
    return int(xfrm.find(qn("a:off")).get("y")) + int(
        xfrm.find(qn("a:ext")).get("cy"))


_SLIDE_H = int(7.5 * _EMU)


# ----------------------------------------------------- resize shape to fit


def test_spautofit_box_past_the_slide_keeps_its_setting_and_shrinks_once(deck):
    """The field case: a box that grows with its text and has grown past
    the slide's bottom edge. The text is fitted to the room between the
    box's top and the slide's bottom by its run sizes, the box still
    resizes to fit its text, and nothing asks PowerPoint to shrink it
    again."""
    pkg, slide = deck
    sid = _add_box(pkg, slide, autofit="spAutoFit", y_in=4.5, h_in=0.6)
    before = [int(s) for s in _run_sizes(pkg, slide, sid)]
    out = tx.fit_text(pkg, slide, sid, min_size=8)
    assert len(out["fitted"]) == 1, out
    rec = out["fitted"][0]
    assert _autofit_children(pkg, slide, sid) == ["spAutoFit"]
    assert rec["autofit"] == "spAutoFit"
    assert rec["autofit_changed"] is False
    assert rec["applied"] == "run_sizes"
    after = [int(s) for s in _run_sizes(pkg, slide, sid)]
    assert all(a < b for a, b in zip(after, before)), (before, after)
    # One uniform scale, applied once.
    ratios = [a / b for a, b in zip(after, before)]
    assert max(ratios) - min(ratios) < 0.01, ratios
    assert abs(ratios[0] * 100 - rec["scale_pct"]) < 1.0, (ratios, rec)
    # The box it renders as fits on the slide now.
    assert _box_bottom_emu(pkg, slide, sid) <= _SLIDE_H
    assert rec["still_overflowing"] is False


def test_spautofit_box_that_stays_on_the_slide_is_left_alone(deck):
    """A box that resizes to fit its text and stays on the slide when it
    grows has nothing to fit, whatever height it was saved with."""
    pkg, slide = deck
    sid = _add_box(pkg, slide, autofit="spAutoFit", y_in=1.0, h_in=0.4,
                   sizes=(18, 14))
    before_sizes = _run_sizes(pkg, slide, sid)
    before_xml = etree.tostring(_elem(pkg, slide, sid))
    out = tx.fit_text(pkg, slide, sid, min_size=8)
    assert out["fitted"] == []
    assert len(out["skipped"]) == 1
    assert "grow" in out["skipped"][0]["reason"]
    assert _run_sizes(pkg, slide, sid) == before_sizes
    assert etree.tostring(_elem(pkg, slide, sid)) == before_xml


# ------------------------------------------------------------- no autofit


def test_noautofit_keeps_noautofit_and_rewrites_the_sizes(deck):
    pkg, slide = deck
    sid = _add_box(pkg, slide, autofit="noAutofit", y_in=1.0, h_in=1.2)
    before = [int(s) for s in _run_sizes(pkg, slide, sid)]
    out = tx.fit_text(pkg, slide, sid, min_size=8)
    rec = out["fitted"][0]
    assert _autofit_children(pkg, slide, sid) == ["noAutofit"]
    assert rec["autofit"] == "none" and rec["applied"] == "run_sizes"
    assert rec["autofit_changed"] is False
    after = [int(s) for s in _run_sizes(pkg, slide, sid)]
    assert all(a < b for a, b in zip(after, before))


def test_a_box_with_no_autofit_setting_gets_none_added(deck):
    """Nothing set on the shape or above it means no autofit; fitting it
    rewrites the sizes and adds no setting of its own."""
    pkg, slide = deck
    sid = _add_box(pkg, slide, autofit=None, y_in=1.0, h_in=1.2)
    out = tx.fit_text(pkg, slide, sid, min_size=8)
    rec = out["fitted"][0]
    assert _autofit_children(pkg, slide, sid) == []
    assert rec["autofit"] == "none" and rec["autofit_from"] == "default"
    assert rec["applied"] == "run_sizes"


def test_a_run_already_below_the_floor_is_never_enlarged(deck):
    """min_size is a floor for shrinking, not a size to grow small print
    to: an 8 pt run in a fit with min_size 10 stays 8 pt."""
    pkg, slide = deck
    sid = _add_box(pkg, slide, autofit="noAutofit", y_in=1.0, h_in=1.0,
                   sizes=(24, 8, 24))
    tx.fit_text(pkg, slide, sid, min_size=10)
    sizes = [int(s) for s in _run_sizes(pkg, slide, sid)]
    assert sizes[1] == 800, sizes
    assert all(s >= 1000 for i, s in enumerate(sizes) if i != 1), sizes


# --------------------------------------------- shrink text on overflow


def test_normautofit_is_fitted_by_its_scale_and_the_sizes_stay(deck):
    pkg, slide = deck
    sid = _add_box(pkg, slide, autofit="normAutofit", y_in=1.0, h_in=1.2)
    before = _run_sizes(pkg, slide, sid)
    out = tx.fit_text(pkg, slide, sid, min_size=8)
    rec = out["fitted"][0]
    assert _run_sizes(pkg, slide, sid) == before  # one mechanism: the scale
    assert rec["autofit"] == "normAutofit"
    assert rec["applied"] == "normAutofit_scale"
    assert rec["autofit_changed"] is False
    norm = _bodypr(pkg, slide, sid).find(qn("a:normAutofit"))
    assert norm is not None
    scale = int(norm.get("fontScale"))
    assert 1000 <= scale < 100000
    assert abs(scale / 1000 - rec["after"]["font_scale_pct"]) < 0.01
    assert rec["after"]["sizes_pt"] == rec["before"]["sizes_pt"]


def test_a_stale_normautofit_cache_is_replaced_not_multiplied(deck):
    """PowerPoint's cached shrink belongs to the text it last drew. The fit
    writes its own scale in place of it (one scale, never a second one on
    top) and leaves the run sizes alone."""
    pkg, slide = deck
    sid = _add_box(
        pkg, slide, y_in=1.0, h_in=1.2,
        autofit='<a:normAutofit fontScale="40000" lnSpcReduction="20000"/>',
    )
    before = _run_sizes(pkg, slide, sid)
    out = tx.fit_text(pkg, slide, sid, min_size=8)
    rec = out["fitted"][0]
    assert _run_sizes(pkg, slide, sid) == before
    norm = _bodypr(pkg, slide, sid).find(qn("a:normAutofit"))
    assert norm.get("lnSpcReduction") is None
    assert int(norm.get("fontScale")) == round(rec["scale_pct"] * 1000)
    assert rec["before"]["font_scale_pct"] == 40.0


# ---------------------------------------------- inherited from the layout


def _content_slide(pkg, words: int = 260):
    """A Title and Content slide whose body placeholder holds a lot of
    text and states no autofit and no box of its own."""
    idx = sl.insert_slide(pkg, "Title and Content")["index"]
    tx.set_placeholder_text(
        pkg, idx, "body", " ".join(["overflowing"] * words))
    part = get_slide_info(pkg, idx)["part"]
    tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    for sp in tree.iter(qn("p:sp")):
        ph = sp.find(f"{qn('p:nvSpPr')}/{qn('p:nvPr')}/{qn('p:ph')}")
        if ph is not None and ph.get("type") in (None, "body", "obj") \
                and ph.get("idx") is not None:
            return idx, sp
    raise AssertionError("no body placeholder")


def test_inherited_normautofit_is_fitted_through_its_scale(deck):
    """The body placeholder states no autofit; the master's says shrink text
    on overflow. The fit uses that setting's scale, written on the shape,
    and writes no run sizes."""
    pkg, _slide = deck
    idx, sp = _content_slide(pkg)
    bodypr = sp.find(f"{qn('p:txBody')}/{qn('a:bodyPr')}")
    assert not [c for c in bodypr if etree.QName(c).localname in _AUTOFIT]
    sid = int(sp.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}").get("id"))
    out = tx.fit_text(pkg, idx, sid, min_size=8)
    assert len(out["fitted"]) == 1, out
    rec = out["fitted"][0]
    assert rec["autofit"] == "normAutofit"
    assert rec["autofit_from"] in ("layout", "master")
    assert rec["applied"] == "normAutofit_scale"
    norm = bodypr.find(qn("a:normAutofit"))
    assert norm is not None and int(norm.get("fontScale")) < 100000
    body = sp.find(qn("p:txBody"))
    assert all(r.get("sz") is None for r in body.iter(qn("a:rPr")))


def test_inherited_no_autofit_writes_the_sizes_out_and_adds_nothing(deck):
    """The layout says do not autofit. The runs inherit their size; the fit
    writes each one out at its inherited size times the scale and adds no
    autofit setting to the shape."""
    from kitchensink4ppt.ops import inherit as inh

    pkg, _slide = deck
    idx, sp = _content_slide(pkg)
    part = get_slide_info(pkg, idx)["part"]
    layout = inh.layout_part_of(pkg, part)
    key = inh.placeholder_key(inh.placeholder_of(sp))
    twin = inh.layout_twin(pkg, layout, key)
    twin_bodypr = twin.find(f"{qn('p:txBody')}/{qn('a:bodyPr')}")
    for child in list(twin_bodypr):
        if etree.QName(child).localname in _AUTOFIT:
            twin_bodypr.remove(child)
    twin_bodypr.insert(0, etree.Element(qn("a:noAutofit")))
    pkg.mark_dirty(layout)
    sid = int(sp.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}").get("id"))
    out = tx.fit_text(pkg, idx, sid, min_size=8)
    rec = out["fitted"][0]
    assert rec["autofit"] == "none" and rec["autofit_from"] == "layout"
    assert rec["applied"] == "run_sizes"
    bodypr = sp.find(f"{qn('p:txBody')}/{qn('a:bodyPr')}")
    assert not [c for c in bodypr if etree.QName(c).localname in _AUTOFIT]
    sizes = [r.get("sz") for r in sp.find(qn("p:txBody")).iter(qn("a:rPr"))]
    assert sizes and all(s is not None for s in sizes)
    inherited = rec["before"]["sizes_pt"]
    assert inherited != "inherited"
    assert all(int(s) / 100 < max(inherited) for s in sizes)


# ------------------------------------------------------------ slide-wide


def test_slide_wide_fit_gives_each_shape_its_own_mechanism(deck):
    pkg, slide = deck
    grow = _add_box(pkg, slide, autofit="spAutoFit", y_in=4.5, h_in=0.6,
                    sid=81)
    shrink = _add_box(pkg, slide, autofit="normAutofit", y_in=0.3, h_in=1.0,
                      sid=82, sizes=(20, 20))
    grow_sizes = _run_sizes(pkg, slide, grow)
    shrink_sizes = _run_sizes(pkg, slide, shrink)
    out = tx.fit_text(pkg, slide, None, min_size=8)
    by_id = {r["shape_id"]: r for r in out["fitted"]}
    assert by_id[grow]["applied"] == "run_sizes"
    assert by_id[shrink]["applied"] == "normAutofit_scale"
    assert _autofit_children(pkg, slide, grow) == ["spAutoFit"]
    assert _autofit_children(pkg, slide, shrink) == ["normAutofit"]
    assert _run_sizes(pkg, slide, grow) != grow_sizes
    assert _run_sizes(pkg, slide, shrink) == shrink_sizes


def test_the_answer_names_the_setting_it_kept():
    """The tool's own description no longer promises to switch shrink text
    on overflow on: it keeps each shape's setting."""
    import inspect

    from kitchensink4ppt import server

    doc = inspect.getdoc(server.fit_text) or ""
    assert "enables normAutofit" not in doc
    assert "autofit" in doc.lower()


# ------------------------------------------------- drawn by PowerPoint


def _com_gate():
    """Skip unless this machine has PowerPoint and no other PowerPoint is
    running: the check starts its own hidden instance and must not share
    one with a person or another job."""
    import sys

    if sys.platform != "win32":
        pytest.skip("PowerPoint rendering is Windows-only")
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        pytest.skip("pywin32 not installed")
    from kitchensink4ppt.com import bridge

    if not bridge.powerpoint_installed():
        pytest.skip("PowerPoint is not installed on this machine")
    import time

    deadline = time.monotonic() + 30  # an instance that is quitting
    running = bridge.powerpnt_pids()
    while (running is None or running) and time.monotonic() < deadline:
        time.sleep(1.0)
        running = bridge.powerpnt_pids()
    if running is None or running:
        pytest.skip(
            "SKIPPED-POWERPOINT-BUSY: another PowerPoint is running, so the "
            "render check did NOT run")


def _drawn(deck: Path, out_dir: Path, names: list[str]) -> dict:
    """Open the deck in a hidden PowerPoint, draw slide 2 to a PNG, keep a
    copy as PowerPoint holds it after drawing, and read each named shape:
    its top and height in points, the slide's height and its line count."""
    from kitchensink4ppt.com import bridge

    seen: dict = {}
    copy_path = out_dir / "held_by_powerpoint.pptx"

    def read(pres):
        # Every proxy lives in this frame and is released when it returns,
        # so the hidden PowerPoint can quit when the session ends.
        slide = pres.Slides.Item(2)
        slide.Export(str((out_dir / "slide2.png").resolve()), "PNG", 1280, 960)
        seen["slide_h"] = float(pres.PageSetup.SlideHeight)
        for name in names:
            shape = slide.Shapes.Item(name)
            try:
                lines = int(shape.TextFrame.TextRange.Lines(1, 10000).Count)
            except Exception:
                lines = None
            seen[name] = {
                "top": float(shape.Top),
                "height": float(shape.Height),
                "lines": lines,
                "autosize": int(shape.TextFrame2.AutoSize),
            }
        pres.SaveCopyAs(str(copy_path.resolve()))

    with bridge._powerpoint() as session:
        pres = bridge.open_presentation(session, deck, read_only=True)
        try:
            read(pres)
        finally:
            del pres
    seen["copy"] = copy_path
    return seen


def _shape_xml(path: Path, name: str):
    pkg = PptxPackage(path)
    part = get_slide_info(pkg, 1)["part"]
    tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    for sp in tree.iter(qn("p:sp")):
        cnv = sp.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}")
        if cnv is not None and cnv.get("name") == name:
            return sp
    raise AssertionError(f"no shape {name!r} in {path.name}")


@pytest.mark.live
@pytest.mark.timeout(600)
def test_powerpoint_draws_the_fitted_text_once(make_deck, tmp_path):
    """The field case drawn by PowerPoint itself. After fit_text, the box
    set to resize to fit its text still is, PowerPoint adds no shrink of
    its own (the copy it holds after drawing carries no fontScale, and its
    run sizes are the ones fit_text wrote), and the box PowerPoint grows
    ends on the slide. The box set to shrink on overflow keeps its run
    sizes; whatever scale PowerPoint lands on is its one scale."""
    _com_gate()
    deck = make_deck("drawn.pptx", extra_slides=0)
    pkg = PptxPackage(deck)
    slide = sl.insert_slide(pkg, "Blank", position=1)["index"]
    grow = _add_box(pkg, slide, autofit="spAutoFit", y_in=4.5, h_in=0.6, sid=81)
    shrink = _add_box(pkg, slide, autofit="normAutofit", y_in=0.3, h_in=1.2,
                      sid=82, sizes=(20, 20))
    out = tx.fit_text(pkg, slide, None, min_size=8)
    by_id = {r["shape_id"]: r for r in out["fitted"]}
    assert set(by_id) == {grow, shrink}, out
    written = {sid: _run_sizes(pkg, slide, sid) for sid in (grow, shrink)}
    pkg.save()

    seen = _drawn(Path(deck), tmp_path, [f"Box {grow}", f"Box {shrink}"])
    grown = seen[f"Box {grow}"]
    assert grown["autosize"] == 1  # msoAutoSizeShapeToFitText, unchanged
    assert grown["top"] + grown["height"] <= seen["slide_h"] + 1.0, seen
    held = _shape_xml(seen["copy"], f"Box {grow}")
    bodypr = held.find(f"{qn('p:txBody')}/{qn('a:bodyPr')}")
    assert bodypr.find(qn("a:normAutofit")) is None
    assert bodypr.find(qn("a:spAutoFit")) is not None
    sizes = [r.get("sz") for r in held.find(qn("p:txBody")).iter(qn("a:rPr"))]
    assert sizes == written[grow]
    kept = _shape_xml(seen["copy"], f"Box {shrink}")
    sizes = [r.get("sz") for r in kept.find(qn("p:txBody")).iter(qn("a:rPr"))]
    assert sizes == written[shrink]
    assert seen[f"Box {shrink}"]["autosize"] == 2  # msoAutoSizeTextToFitShape
