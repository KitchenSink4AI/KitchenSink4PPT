"""Real advance-width measurement in the text-fit model.

Field report 2026-09-21 (P2): the fit model was one estimator, 0.5 x point
size per character with a character-count wrap, and it predicted a two-line
wrap for a Calibri 12 pt label that PowerPoint laid out as three. The
estimator also wraps by character count, which no renderer does, so a frame
whose words do not pack neatly reads as fitting when it does not.

These tests pin: the measured path flags the field case, a roomy frame stays
quiet either way, the estimate path is byte-for-byte what it always was when
the extra is absent, and every result names the model that produced it.
"""

from __future__ import annotations

import pytest

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
    not _HAVE_CALIBRI,
    reason="needs the metrics extra and an installed Calibri",
)
needs_metrics = pytest.mark.skipif(
    not _HAVE_METRICS, reason="needs the metrics extra (fonttools)"
)


@pytest.fixture
def deck(tmp_path):
    p = tmp_path / "fit.pptx"
    sl.create_presentation(str(p))
    pkg = PptxPackage(str(p))
    sl.insert_slide(pkg, layout="Blank")
    return pkg


def _add(pkg, x, y, w, h, text, pt, wrap=None):
    sid = shp.insert_shape(
        pkg, 0, "rect", x, y, w, h, text=text,
        text_style={"size": pt, "font": "Calibri"},
    )["shape_id"]
    if wrap is not None:
        part = pkg.slide_parts()[0]
        for elem, _k, _z, _p in iter_shapes(tx._sp_tree(pkg, part)):
            if tx._shape_id(elem) == sid:
                elem.find(qn("p:txBody")).find(qn("a:bodyPr")).set("wrap", wrap)
        pkg.mark_dirty(part)
    return sid


def _overflow(pkg, sid):
    return tx.get_autofit_state(pkg, 0, sid)["shapes"][0]["overflow"]


# --------------------------------------------------------------- the model


@needs_calibri
def test_field_case_word_wrap_beats_the_character_count(deck):
    """The 2026-09-21 shape of the defect: a descriptor the character-count
    model packs into two lines that really wraps to three, in a frame only
    tall enough for two."""
    sid = _add(deck, 1, 1, 2.5, 0.60,
               "Alliance politics: abandonment and entrapment", 14)
    measured = _overflow(deck, sid)
    assert measured["method"] == "font-metrics"
    assert measured["estimated_lines"] == 3
    assert measured["likely_overflow"] is True


@needs_calibri
def test_the_same_shape_passes_under_the_old_estimate(deck, monkeypatch):
    """Fail-first evidence, kept as a test: with measurement switched off
    the identical shape is reported as fitting."""
    sid = _add(deck, 1, 1, 2.5, 0.60,
               "Alliance politics: abandonment and entrapment", 14)
    monkeypatch.setattr(fm, "available", lambda: False)
    est = _overflow(deck, sid)
    assert est["method"] == "estimate"
    assert est["estimated_lines"] == 2
    assert est["likely_overflow"] is False


@needs_calibri
def test_non_wrapping_label_overflows_sideways(deck):
    """The Snyder label: wrap="none", so PowerPoint runs the text out of
    the frame instead of wrapping it. The old model always wrapped, so it
    could not see this class of defect at all."""
    sid = _add(deck, 1, 1, 0.744, 0.50, "Snyder (1997)", 12, wrap="none")
    measured = _overflow(deck, sid)
    assert measured["method"] == "font-metrics"
    assert measured["wrap"] == "none"
    assert measured["width_ratio"] > 1.0
    assert measured["likely_overflow"] is True


@needs_calibri
def test_non_wrapping_label_passes_under_the_old_estimate(deck, monkeypatch):
    sid = _add(deck, 1, 1, 0.744, 0.50, "Snyder (1997)", 12, wrap="none")
    monkeypatch.setattr(fm, "available", lambda: False)
    est = _overflow(deck, sid)
    assert est["likely_overflow"] is False


@needs_calibri
def test_a_roomy_frame_stays_quiet(deck):
    sid = _add(deck, 1, 1, 6.0, 2.0, "short label", 14)
    measured = _overflow(deck, sid)
    assert measured["method"] == "font-metrics"
    assert measured["likely_overflow"] is False
    assert measured["fill_ratio"] < 0.5


@needs_calibri
def test_paragraph_hanging_indent_narrows_the_line(deck):
    """marL/indent take width away from the text, which the old model
    ignored entirely."""
    sid = _add(deck, 1, 1, 2.2, 1.2,
               "Alliance politics: abandonment and entrapment", 12)
    plain = _overflow(deck, sid)["estimated_lines"]
    part = deck.slide_parts()[0]
    for elem, _k, _z, _p in iter_shapes(tx._sp_tree(deck, part)):
        if tx._shape_id(elem) == sid:
            for p in elem.find(qn("p:txBody")).findall(qn("a:p")):
                ppr = p.find(qn("a:pPr"))
                if ppr is None:
                    from lxml import etree

                    ppr = etree.SubElement(p, qn("a:pPr"))
                    p.insert(0, ppr)
                ppr.set("marL", "1143000")   # 1.25 in
                ppr.set("indent", "0")
    deck.mark_dirty(part)
    indented = _overflow(deck, sid)["estimated_lines"]
    assert indented > plain


# ------------------------------------------------- honesty of the result


def test_every_result_states_its_method(deck):
    sid = _add(deck, 1, 1, 2.0, 0.6, "some words here to lay out", 12)
    rec = _overflow(deck, sid)
    assert rec["method"] in ("font-metrics", "estimate")
    if rec["method"] == "font-metrics":
        assert rec["font_file"]
        assert rec["heuristic"] is False
    else:
        assert rec["method_reason"]
        assert rec["heuristic"] is True


def test_estimate_path_when_the_extra_is_absent(deck, monkeypatch):
    """Without fonttools the model is exactly what it always was: the
    0.5 x point size character width and the 1.2 line height."""
    import math

    text = "Alliance politics: abandonment and entrapment"
    sid = _add(deck, 1, 1, 2.5, 0.60, text, 14)
    monkeypatch.setattr(fm, "available", lambda: False)
    rec = _overflow(deck, sid)
    assert rec["method"] == "estimate"
    assert "fonttools" in rec["method_reason"]
    inner_w = int(2.5 * 914400) - 2 * 91440
    inner_h = int(0.60 * 914400) - 2 * 45720
    char_w = 0.5 * 14 * 12700
    line_h = 1.2 * 14 * 12700
    lines = max(1, math.ceil(len(text) * char_w / inner_w))
    assert rec["estimated_lines"] == lines
    assert rec["fill_ratio"] == round(lines * line_h / inner_h, 2)


@needs_metrics
def test_missing_font_falls_back_and_says_which_font(deck, monkeypatch):
    """The typeface resolved but no file for it exists here. Needs the extra
    installed, because without it the earlier reason fires first and is the
    correct one to report."""
    sid = _add(deck, 1, 1, 2.0, 0.6, "some words here", 12)
    monkeypatch.setattr(fm, "font_file_name", lambda *a, **k: None)
    rec = _overflow(deck, sid)
    assert rec["method"] == "estimate"
    assert "Calibri" in rec["method_reason"]


@needs_calibri
def test_fit_text_reports_the_method(deck):
    sid = _add(deck, 1, 1, 2.0, 0.5,
               "Alliance politics: abandonment and entrapment", 24)
    out = tx.fit_text(deck, 0, sid, min_size=8)
    assert out["method"] == "font-metrics"
    assert out["estimate"] is True  # still an estimate of a renderer
    assert out["fitted"][0]["method"] == "font-metrics"
    assert out["fitted"][0]["font_file"]


def test_fit_text_note_is_honest_without_the_extra(deck, monkeypatch):
    sid = _add(deck, 1, 1, 2.0, 0.5,
               "Alliance politics: abandonment and entrapment", 24)
    monkeypatch.setattr(fm, "available", lambda: False)
    out = tx.fit_text(deck, 0, sid, min_size=8)
    assert out["method"] == "estimate"
    assert "no real font metrics" in out["note"]


# ------------------------------------------------------- check_layout


@needs_calibri
def test_check_layout_flags_a_measured_overflow_the_estimate_suppressed(deck):
    """min_fill_ratio 1.4 was an allowance for MODEL error. A measurement
    does not need that much of one, and a real ~1.1x overflow was being
    swallowed by it."""
    sid = _add(deck, 1, 1, 2.5, 0.72,
               "Alliance politics: abandonment and entrapment", 14)
    deck.save()
    out = dc.check_layout(deck, slide=0, checks=["overflow"])
    hits = [f for f in out["findings"] if sid in f.get("shape_ids", [])]
    assert hits, "a measured 1.16x overflow must be reported"
    assert hits[0]["method"] == "font-metrics"
    assert hits[0]["heuristic"] is False
    assert 1.05 <= hits[0]["fill_ratio"] < 1.4


@needs_calibri
def test_check_layout_estimate_path_keeps_the_wide_margin(deck, monkeypatch):
    sid = _add(deck, 1, 1, 2.5, 0.72,
               "Alliance politics: abandonment and entrapment", 14)
    deck.save()
    monkeypatch.setattr(fm, "available", lambda: False)
    out = dc.check_layout(deck, slide=0, checks=["overflow"])
    hits = [f for f in out["findings"] if sid in f.get("shape_ids", [])]
    assert not hits, (
        "under the character-count estimate this shape is inside the 1.4x "
        "model-error margin and must stay suppressed"
    )


def test_overflow_check_declares_both_thresholds():
    opts, caveat = dc.CHECKS["overflow"]
    assert opts["min_fill_ratio"] == 1.4
    # 1.0, not 1.05: the conservative calibration moved INSIDE the
    # measurement in round 2, so this model needs no extra margin here.
    assert opts["min_fill_ratio_metrics"] == 1.0
    assert "measured against its own font" in caveat
    assert "per-line allowance" in caveat


# ------------------------------------------------------- the module itself


def test_metrics_is_never_a_runtime_requirement():
    import tomllib
    from pathlib import Path

    root = Path(__file__).resolve().parents[2]
    data = tomllib.loads((root / "pyproject.toml").read_text(encoding="utf-8"))
    required = " ".join(data["project"]["dependencies"]).lower()
    assert "fonttools" not in required
    assert "fonttools" in " ".join(
        data["project"]["optional-dependencies"]["metrics"]
    ).lower()


def test_fontmetrics_imports_without_the_extra(monkeypatch):
    """The module must import and answer on a machine with no fonttools."""
    import builtins

    real_import = builtins.__import__

    def _blocked(name, *args, **kwargs):
        if name.startswith("fontTools"):
            raise ImportError("blocked for the test")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", _blocked)
    monkeypatch.setattr(fm, "_AVAILABLE", None)  # re-ask, do not reuse
    assert fm.available() is False
    assert fm.text_width_pt("abc", "Calibri", 12) is None


@pytest.mark.parametrize(
    "face,expected",
    [
        ("Calibri (TrueType)", ("Calibri", False, False)),
        ("Calibri Bold (TrueType)", ("Calibri", True, False)),
        ("Calibri Bold Italic (TrueType)", ("Calibri", True, True)),
        ("Arial Narrow Italic (TrueType)", ("Arial Narrow", False, True)),
        ("Segoe UI", ("Segoe UI", False, False)),
    ],
)
def test_style_words_are_split_off_the_family(face, expected):
    assert fm._split_style(face) == expected


@needs_calibri
def test_bold_and_italic_pick_their_own_face():
    plain = fm.find_font_file("Calibri")
    bold = fm.find_font_file("Calibri", bold=True)
    assert plain is not None and bold is not None
    assert plain != bold
    wide = fm.text_width_pt("Hamburgefonstiv", "Calibri", 20, bold=True)
    narrow = fm.text_width_pt("Hamburgefonstiv", "Calibri", 20)
    assert wide > narrow


@needs_calibri
def test_word_wrap_uses_word_boundaries():
    # Two words that each fit but do not fit together.
    one = fm.text_width_pt("abandonment", "Calibri", 12)
    assert fm.wrapped_line_count(
        "abandonment entrapment", "Calibri", 12, one * 1.5
    ) == 2
    assert fm.wrapped_line_count(
        "abandonment entrapment", "Calibri", 12, one * 3
    ) == 1


@needs_calibri
def test_a_word_wider_than_the_line_breaks_inside_itself():
    n = fm.wrapped_line_count("internationalization", "Calibri", 12, 20.0)
    assert n is not None and n >= 2


@needs_calibri
def test_an_unknown_family_measures_nothing():
    assert fm.find_font_file("NoSuchFontFamilyAnywhere") is None
    assert fm.text_width_pt("x", "NoSuchFontFamilyAnywhere", 12) is None
