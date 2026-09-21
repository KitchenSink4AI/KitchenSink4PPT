"""Round 2: the second review's findings on the measured text-fit model.

B1  the conservative calibration must sit INSIDE the measurement, because
    the overflow boolean is decided against 1.0 before any suppression
    threshold is consulted;
M1  a mixed-format body must be measured run by run, or not called measured;
M2  a font collection holds several unrelated families and face zero is not
    an answer;
N3  a font file that exists but will not parse is a different fallback
    reason from a font file that is absent.

The near-boundary cases carry REAL numbers: every `true_ratio` below was
read from PowerPoint's own TextRange.BoundWidth in a hidden instance on
2026-09-22, not computed by this model.
"""

from __future__ import annotations

import pytest
from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import design_check as dc
from kitchensink4ppt.ops import fontmetrics as fm
from kitchensink4ppt.ops import shapes as shp
from kitchensink4ppt.ops import slides as sl
from kitchensink4ppt.ops import text as tx
from kitchensink4ppt.ops.read import iter_shapes

_HAVE_METRICS = fm.available()
_HAVE_CALIBRI = _HAVE_METRICS and fm.find_font_file("Calibri") is not None
needs_calibri = pytest.mark.skipif(
    not _HAVE_CALIBRI, reason="needs the metrics extra and an installed Calibri"
)
needs_metrics = pytest.mark.skipif(
    not _HAVE_METRICS, reason="needs the metrics extra (fonttools)"
)


@pytest.fixture
def deck(tmp_path):
    p = tmp_path / "r2.pptx"
    sl.create_presentation(str(p))
    pkg = PptxPackage(str(p))
    sl.insert_slide(pkg, layout="Blank")
    return pkg


def _elem(pkg, sid):
    for e, _k, _z, _p in iter_shapes(tx._sp_tree(pkg, pkg.slide_parts()[0])):
        if tx._shape_id(e) == sid:
            return e
    raise AssertionError(f"shape {sid} not found")


def _add(pkg, w, h, text, pt, wrap=None, font="Calibri"):
    sid = shp.insert_shape(
        pkg, 0, "rect", 0.6, 0.6, w, h, text=text,
        text_style={"size": pt, "font": font},
    )["shape_id"]
    if wrap is not None:
        _elem(pkg, sid).find(qn("p:txBody")).find(qn("a:bodyPr")).set(
            "wrap", wrap
        )
        pkg.mark_dirty(pkg.slide_parts()[0])
    return sid


def _overflow(pkg, sid):
    return tx.get_autofit_state(pkg, 0, sid)["shapes"][0]["overflow"]


# ===================================================== B1: the calibration


def test_the_calibration_is_a_per_line_allowance_not_a_per_token_one():
    """A pad added per token would multiply by the number of words."""
    assert fm.calibrated_line_pt(100.0) == 100.0 + fm.WIDTH_SAFETY_PAD_PT
    assert fm.WIDTH_SAFETY_PAD_PT > 0


@needs_calibri
def test_text_width_pt_itself_stays_uncalibrated():
    """The per-token measure must not carry the line allowance, or a wrap
    over N words would over-count the line by N pads."""
    raw = fm.raw_text_width_pt("Snyder (1997)", "Calibri", 12)
    assert fm.text_width_pt("Snyder (1997)", "Calibri", 12) == raw


@needs_calibri
def test_the_wrap_takes_the_allowance_off_the_line():
    """A string that fits the raw width but not the calibrated width must
    wrap, because PowerPoint wraps it."""
    text = "abandonment entrapment"
    raw = fm.raw_text_width_pt(text, "Calibri", 12)
    # A frame exactly as wide as the raw sum: uncalibrated this is one line.
    lines = fm.wrap_styled(
        [fm.StyledRun(text, "Calibri", 12)], raw
    )
    assert lines is not None
    assert len(lines) == 2, (
        "with the per-line allowance taken off, a string that exactly fills "
        "the raw width no longer fits on one line"
    )


#: (text, box width in inches, pt, PowerPoint's true width ratio). Measured
#: against TextRange.BoundWidth in a hidden PowerPoint on 2026-09-22. These
#: are non-wrapping labels, so width alone decides.
_BAND = [
    ("Christensen (2015)", 1.6198, 12, 0.9868),
    ("Sukin and Seo (2024)", 1.6345, 12, 1.0179),
    ("Henry (2022)", 1.0800, 12, 1.0504),
    ("Glaser (2012)", 1.0884, 12, 1.0671),
    ("Alvut and Mote (2025)", 1.6706, 12, 1.0660),
    ("Burton (1990)", 1.1088, 12, 1.1013),
    ("Azar (1990)", 0.9193, 12, 1.1329),
    ("OPCON transfer", 1.1812, 12, 1.1287),
    ("Combined Forces Command", 1.5886, 12, 1.1405),
    ("Mutual Defense Treaty", 1.5652, 12, 1.1525),
    ("Snyder (1997)", 0.744, 12, 1.797),
]


@needs_calibri
@pytest.mark.parametrize("text,box_in,pt,true_ratio", _BAND)
def test_no_false_negative_in_the_boundary_band(deck, text, box_in, pt,
                                                true_ratio):
    """The whole point of B1: a real overflow must never come back as
    "fits" because the model measures narrow."""
    sid = _add(deck, box_in, 0.5, text, pt, wrap="none")
    rec = _overflow(deck, sid)
    assert rec["method"] == "font-metrics"
    if true_ratio > 1.0:
        assert rec["likely_overflow"] is True, (
            f"PowerPoint renders {text!r} at {true_ratio}x its frame and the "
            "model said it fits"
        )
    else:
        assert rec["likely_overflow"] is False, (
            f"PowerPoint renders {text!r} at {true_ratio}x its frame, which "
            "fits, and the model cried wolf"
        )


@needs_calibri
def test_a_real_overflow_just_over_the_edge_survives_check_layout(deck):
    """B1's concrete failure scenario: at the old 1.05 gate a label a few
    percent over its frame was discarded twice, once by the uncalibrated
    boolean and once by the threshold."""
    sid = _add(deck, 1.6345, 0.5, "Sukin and Seo (2024)", 12, wrap="none")
    deck.save()
    out = dc.check_layout(deck, slide=0, checks=["overflow"])
    hits = [f for f in out["findings"] if sid in f.get("shape_ids", [])]
    assert hits, "a 1.018x width overflow must reach the caller"
    assert hits[0]["method"] == "font-metrics"


def test_the_metrics_threshold_is_one_because_the_model_is_calibrated():
    opts, _caveat = dc.CHECKS["overflow"]
    assert opts["min_fill_ratio_metrics"] == 1.0
    assert opts["min_fill_ratio"] == 1.4


# ======================================================= M1: per-run truth


def _set_runs(pkg, sid, runs):
    """Rewrite a shape's body as explicit runs: (text, size_pt, bold, face)."""
    body = _elem(pkg, sid).find(qn("p:txBody"))
    for p in body.findall(qn("a:p")):
        body.remove(p)
    p = etree.SubElement(body, qn("a:p"))
    for text, size, bold, face in runs:
        r = etree.SubElement(p, qn("a:r"))
        rpr = etree.SubElement(r, qn("a:rPr"))
        rpr.set("lang", "en-US")
        rpr.set("sz", str(int(size * 100)))
        if bold:
            rpr.set("b", "1")
        latin = etree.SubElement(rpr, qn("a:latin"))
        latin.set("typeface", face)
        t = etree.SubElement(r, qn("a:t"))
        t.text = text
    pkg.mark_dirty(pkg.slide_parts()[0])


@needs_calibri
def test_a_later_bigger_run_is_what_overflows_and_is_seen(deck):
    """M1's concrete failure: a 10 pt label followed by a 30 pt value. Read
    in the first run's format the body fits comfortably; read honestly it
    does not."""
    sid = _add(deck, 2.0, 0.55, "placeholder", 10)
    _set_runs(deck, sid, [
        ("Label ", 10, False, "Calibri"),
        ("SUBSTANTIAL VALUE", 30, True, "Calibri"),
    ])
    rec = _overflow(deck, sid)
    assert rec["method"] == "font-metrics"
    assert rec["runs_measured"] == 2
    assert rec["likely_overflow"] is True


@needs_calibri
def test_line_height_comes_from_the_largest_size_on_the_line(deck):
    """A 30 pt run on a line makes that line 30 pt tall, whatever the run
    before it was set in."""
    small = _add(deck, 4.0, 0.6, "placeholder", 10)
    _set_runs(deck, small, [("Label one two", 10, False, "Calibri")])
    low = _overflow(deck, small)["height_ratio"]
    _set_runs(deck, small, [
        ("Label ", 10, False, "Calibri"),
        ("BIG", 30, False, "Calibri"),
    ])
    high = _overflow(deck, small)["height_ratio"]
    assert high > low * 2, (
        "the tall run must set the line height, not the first run"
    )


@needs_calibri
def test_a_mixed_family_body_names_every_face_it_measured(deck):
    sid = _add(deck, 3.0, 0.8, "placeholder", 12)
    _set_runs(deck, sid, [
        ("Calibri here ", 12, False, "Calibri"),
        ("Georgia there", 12, False, "Georgia"),
    ])
    rec = _overflow(deck, sid)
    if rec["method"] != "font-metrics":
        pytest.skip("Georgia is not installed here")
    assert rec["runs_measured"] == 2
    assert "," in rec["typeface"]
    assert "," in rec["font_file"]


@needs_calibri
def test_an_unmeasurable_run_refuses_the_measured_label(deck):
    """Never label a body font-metrics when one of its runs could not be
    measured. The estimate is the honest answer and says why."""
    sid = _add(deck, 2.0, 0.6, "placeholder", 12)
    _set_runs(deck, sid, [
        ("fine ", 12, False, "Calibri"),
        ("unmeasurable", 12, False, "NoSuchFontFamilyAnywhere"),
    ])
    rec = _overflow(deck, sid)
    assert rec["method"] == "estimate"
    assert rec["heuristic"] is True


# ================================================ M2: collections and faces


@needs_metrics
def test_a_collection_face_is_selected_by_name_not_by_position():
    """cambria.ttc carries Cambria at index 0 and Cambria Math at index 1.
    Taking index 0 for both measured Cambria Math as Cambria."""
    plain = fm.find_font_face("Cambria")
    math = fm.find_font_face("Cambria Math")
    if plain is None or math is None:
        pytest.skip("Cambria/Cambria Math not installed here")
    assert plain.path == math.path, "both live in the same collection here"
    assert plain.index != math.index, (
        "they are different faces and must resolve to different indices"
    )


@needs_metrics
def test_two_faces_of_one_collection_measure_differently():
    if fm.find_font_face("Cambria Math") is None:
        pytest.skip("Cambria Math not installed here")
    a = fm.raw_text_width_pt("x+y=z", "Cambria", 20)
    b = fm.raw_text_width_pt("x+y=z", "Cambria Math", 20)
    assert a is not None and b is not None
    assert a != b, (
        "a path-keyed cache served the second face the first face's widths"
    )


@needs_metrics
def test_a_collection_face_is_named_with_its_index():
    """"cambria.ttc" alone names two fonts, so the report must not."""
    name = fm.font_file_name("Cambria Math")
    if name is None:
        pytest.skip("Cambria Math not installed here")
    assert name.endswith("#1") or "#" in name


@needs_metrics
def test_the_table_cache_is_keyed_by_face_not_by_path():
    faces = [
        f for f in (fm.find_font_face("Cambria"),
                    fm.find_font_face("Cambria Math"))
        if f is not None
    ]
    if len(faces) != 2:
        pytest.skip("Cambria/Cambria Math not installed here")
    for f in faces:
        fm._load_table(f)
    keys = {k for k in fm._TABLES if k[0] == str(faces[0].path)}
    assert len({k[1] for k in keys}) >= 2, (
        f"both faces of {faces[0].path.name} must cache separately, got {keys}"
    )


@needs_metrics
def test_a_multi_family_cjk_collection_resolves_its_members():
    got = {
        name: fm.font_file_name(name)
        for name in ("MS Gothic", "MS PGothic", "MS UI Gothic")
    }
    present = {k: v for k, v in got.items() if v}
    if len(present) < 2:
        pytest.skip("the MS Gothic collection is not installed here")
    # Same file, different faces inside it.
    assert len({v.split("#")[0] for v in present.values()}) == 1
    assert len(set(present.values())) == len(present), present


# ======================================================== N3: honest reasons


@needs_calibri
def test_a_font_file_that_will_not_parse_says_so(deck, monkeypatch):
    """Found but unreadable is not the same as missing, and telling someone
    to install a font they already have wastes their afternoon."""
    sid = _add(deck, 2.0, 0.6, "some words here", 12)
    monkeypatch.setattr(fm, "metrics_readable", lambda *a, **k: False)
    monkeypatch.setattr(fm, "font_file_name", lambda *a, **k: None)
    rec = _overflow(deck, sid)
    assert rec["method"] == "estimate"
    assert "could not be read" in rec["method_reason"]
    assert "no font file" not in rec["method_reason"]


@needs_metrics
def test_metrics_readable_separates_the_three_outcomes(tmp_path):
    assert fm.metrics_readable("NoSuchFontFamilyAnywhere") is None
    if fm.find_font_face("Calibri") is not None:
        assert fm.metrics_readable("Calibri") is True
    # A file that exists and is not a font: found, unreadable.
    junk = tmp_path / "junk.ttf"
    junk.write_bytes(b"not a font at all")
    assert fm._load_table(fm.FontFace(junk, 0)) is None
