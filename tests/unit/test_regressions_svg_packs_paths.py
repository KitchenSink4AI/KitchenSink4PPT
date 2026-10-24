"""Regressions for hostile and out-of-range inputs and for error mapping.

Coordinate ceiling in the geometry path, plus payload validation.
SVG parser and recursion errors refuse in-envelope, at the source.
script/foreignObject content is warned about, never silently dropped.
Calls to a disabled tool signpost the pack that enables it (stdio).
Retired KS4P_MODE values fail closed at startup.
An explicit-address miss is NOT_FOUND, not STALE_ANCHOR.
A directory path is BAD_PARAMS, not DOCUMENT_LOCKED.
The pack hint never fires on echoed user input.
The a4 and letter slide-size presets are accepted.
apply_edits documents atomic as fixed True.
"""

from __future__ import annotations

import hashlib
import inspect
import json
import shutil
import zipfile
from pathlib import Path

import pytest

from kitchensink4ppt import server
from kitchensink4ppt.core.errors import ValidationFailed
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import geometry as g
from kitchensink4ppt.settings import Settings, SettingsError
from kitchensink4ppt.surface import envelope

_SVG_NS = 'xmlns="http://www.w3.org/2000/svg"'


def _fn(name: str):
    fn = getattr(server, name)

    def call(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except BaseException as exc:
            return envelope.refusal(exc)

    return call


def _md5(path) -> str:
    return hashlib.md5(Path(path).read_bytes()).hexdigest()


# ------------------------------------------------- coordinate ceiling


def test_c1_svg_giant_target_box_refused(make_deck):
    deck = make_deck("c1.pptx")
    before = _md5(deck)
    out = _fn("svg_to_shapes")(
        file_path=str(deck), slide=0,
        svg=f'<svg {_SVG_NS}><rect width="5" height="5"/></svg>',
        x=0, y=0, w=1e9, h=1e9,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "2147483647" in out["error"]["message"]
    assert _md5(deck) == before, "refused call must not touch the file"


def test_m3_svg_box_past_limit_but_below_schema_max_refused(make_deck):
    # 3000 in = 2.74e9 EMU: past PowerPoint's 2^31-1 but far below the
    # ECMA schema maximum; must refuse identically.
    deck = make_deck("m3.pptx")
    out = _fn("svg_to_shapes")(
        file_path=str(deck), slide=0,
        svg=f'<svg {_SVG_NS}><rect width="5" height="5"/></svg>',
        x=0, y=0, w=3000, h=3000,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"


def test_c2_insert_shape_oversize_refused_within_bounds_ok(make_deck):
    deck = make_deck("c2.pptx")
    before = _md5(deck)
    out = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect",
        x=0, y=0, w=10000, h=10000,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "2147483647" in out["error"]["message"]
    assert _md5(deck) == before

    ok = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect",
        x=0, y=0, w=1000, h=1000,
    )
    assert ok["ok"] is True, "in-bounds large geometry must still work"


def test_c2_set_shape_oversize_resize_refused(make_deck):
    deck = make_deck("c2b.pptx")
    made = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect", x=1, y=1, w=2, h=1,
    )
    assert made["ok"] is True
    sid = made["changed"]["shape_id"]
    before = _md5(deck)
    out = _fn("set_shape")(
        file_path=str(deck), slide=0, shape=sid, w=10000.0,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "2147483647" in out["error"]["message"]
    assert _md5(deck) == before
    # oversize move refused too
    out = _fn("set_shape")(
        file_path=str(deck), slide=0, shape=sid, x=100000.0,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"


def test_c2_textbox_table_chart_oversize_refused(make_deck):
    deck = make_deck("c2c.pptx")
    out = _fn("insert_textbox")(
        file_path=str(deck), slide=0, text="x", x=0, y=0, w=9999.0, h=9999.0,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    out = _fn("create_table")(
        file_path=str(deck), slide=0, rows=2, cols=2,
        x=0, y=0, w=9999, h=9999,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    out = _fn("create_chart")(
        file_path=str(deck), slide=0, chart_type="bar",
        categories=["a", "b"], series=[{"name": "s", "values": [1, 2]}],
        x=0, y=0, w=9999, h=9999,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"


def test_save_refuses_out_of_range_coordinates(make_deck):
    """The chokepoint behind the chokepoints: even a direct in-memory tamper
    cannot reach disk with an out-of-range coordinate."""
    deck = make_deck("val1.pptx")
    made = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect", x=1, y=1, w=2, h=1,
    )
    assert made["ok"] is True
    pkg = PptxPackage(deck)
    part = pkg.slide_parts()[0]
    ext = pkg.root(part).iter(qn("a:ext"))
    target = next(iter(ext))
    target.set("cx", "914400000000000")
    pkg.mark_dirty(part)
    with pytest.raises(ValidationFailed, match="2147483647"):
        pkg.save()


def test_validate_payload_flags_unopenable_deck(make_deck, tmp_path, monkeypatch):
    """payload_valid must never say True for a deck PowerPoint cannot open:
    a deck carrying a 9e14 EMU extent fails the payload check (the worst
    case was validate blessing the broken file)."""
    import kitchensink4ppt.com.bridge as bridge

    def _no_com(fp, **_kw):
        raise RuntimeError("COM stubbed out for the unit test")

    monkeypatch.setattr(bridge, "com_validate_opens_clean", _no_com)
    deck = make_deck("val2.pptx")
    made = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect", x=0, y=0, w=2, h=1,
    )
    assert made["ok"] is True
    # Rebuild the zip with the extent tampered to the adversarial value.
    bad = tmp_path / "val2_bad.pptx"
    with zipfile.ZipFile(deck) as zin, zipfile.ZipFile(bad, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.startswith("ppt/slides/slide"):
                data = data.replace(b'cx="1828800"', b'cx="914400000000000"')
            zout.writestr(info, data)
    with zipfile.ZipFile(bad) as zf:
        tampered = any(
            b"914400000000000" in zf.read(n)
            for n in zf.namelist() if n.startswith("ppt/slides/slide")
        )
    assert tampered, "test setup failed to plant the oversize extent"
    out = _fn("validate")(file_path=str(bad))
    assert out["ok"] is False
    assert out.get("payload_valid") is not True
    if "payload_valid" in out:
        assert out["payload_valid"] is False
        assert "2147483647" in out["payload_error"]
    else:
        assert out["error"]["code"] in {
            "BAD_PARAMS", "VALIDATION_FAILED", "INTERNAL_ERROR"
        }


def test_geometry_check_emu_box_names_dimension():
    from kitchensink4ppt.core.errors import PptMcpError

    with pytest.raises(PptMcpError, match=r"w = .*2147483647"):
        g.check_emu_box(0, 0, g.MAX_EMU + 1, 100)
    # far-edge overflow (each value legal, sum past the limit) refuses too
    with pytest.raises(PptMcpError, match=r"x\+w"):
        g.check_emu_box(g.MAX_EMU - 10, 0, 100, 100)
    g.check_emu_box(0, 0, g.MAX_EMU, 1)  # at the limit is legal


# ------------------------------------- SVG hostile-input refusals


def _svg_call(deck, svg):
    return _fn("svg_to_shapes")(
        file_path=str(deck), slide=0, svg=svg, x=1, y=1, w=3, h=3,
    )


def test_h1_malformed_xml_refuses_in_envelope(make_deck):
    deck = make_deck("h1a.pptx")
    out = _svg_call(deck, f'<svg {_SVG_NS}><rect width=8')
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "XML" in out["error"]["message"]


def test_h1_non_svg_root_refuses_in_envelope(make_deck):
    deck = make_deck("h1b.pptx")
    out = _svg_call(deck, "<html><body>hi</body></html>")
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "<svg>" in out["error"]["message"]


def test_h2_deep_nesting_refuses_in_envelope(make_deck):
    deck = make_deck("h2a.pptx")
    svg = (
        f'<svg {_SVG_NS}>' + "<g>" * 2000 + "<rect/>" + "</g>" * 2000
        + "</svg>"
    )
    out = _svg_call(deck, svg)
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"


def test_h2_use_ancestor_cycle_refuses_in_envelope(make_deck):
    deck = make_deck("h2b.pptx")
    svg = (
        f'<svg {_SVG_NS} xmlns:xlink="http://www.w3.org/1999/xlink">'
        '<g id="a"><use xlink:href="#a"/></g></svg>'
    )
    out = _svg_call(deck, svg)
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "cycle" in out["error"]["message"]


def test_h2_use_chain_cycle_refuses_in_envelope(make_deck):
    deck = make_deck("h2c.pptx")
    svg = (
        f'<svg {_SVG_NS}>'
        '<defs><g id="a"><use href="#b"/></g>'
        '<g id="b"><use href="#a"/></g></defs>'
        '<use href="#a"/></svg>'
    )
    out = _svg_call(deck, svg)
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "cycle" in out["error"]["message"]


def test_m1_undefined_entity_refuses_in_envelope(make_deck):
    deck = make_deck("m1a.pptx")
    out = _svg_call(deck, f'<svg {_SVG_NS}><rect width="5" height="5"/>&xxe;</svg>')
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"


def test_m1_entity_bomb_refuses_in_envelope(make_deck):
    deck = make_deck("m1b.pptx")
    ents = "".join(
        '<!ENTITY l%d "%s">' % (i, ("&l%d;" % (i - 1)) * 10)
        for i in range(1, 8)
    )
    svg = (
        f'<!DOCTYPE svg [<!ENTITY l0 "lol">{ents}]>'
        f'<svg {_SVG_NS}><text>&l7;</text></svg>'
    )
    out = _svg_call(deck, svg)
    assert out["ok"] is False
    assert out["error"]["code"] in ("BAD_PARAMS", "UNSUPPORTED_CONTENT")


def test_m1_external_entity_never_leaks(make_deck):
    deck = make_deck("m1c.pptx")
    svg = (
        '<!DOCTYPE svg [<!ENTITY xxe SYSTEM "file:///c:/windows/win.ini">]>'
        f'<svg {_SVG_NS}><text>&xxe;</text></svg>'
    )
    out = _svg_call(deck, svg)
    assert out["ok"] is False
    assert "[fonts]" not in str(out).lower()
    assert "extensions" not in out["error"]["message"].lower()


def test_widened_catch_set_maps_deliberately():
    assert envelope._classify(RecursionError()) == "UNSUPPORTED_CONTENT"
    assert envelope._classify(AttributeError("x")) == "INTERNAL_ERROR"
    refused = envelope.refusal(RuntimeError("secret"))
    assert refused["error"]["code"] == "INTERNAL_ERROR"
    assert "secret" not in refused["error"]["message"]


# --------------------------------------------- silent drops become warnings


def test_m2_script_and_foreignobject_warned(make_deck):
    deck = make_deck("m2.pptx")
    out = _svg_call(
        deck,
        f'<svg {_SVG_NS}><script>alert(1)</script>'
        '<foreignObject><body>hi</body></foreignObject>'
        '<rect width="5" height="5" fill="red"/></svg>',
    )
    assert out["ok"] is True
    joined = " ".join(out["warnings"])
    assert "script" in joined
    assert "foreignObject" in joined
    skipped = out["changed"].get("skipped") or {}
    assert skipped.get("script") == 1
    assert skipped.get("foreignObject") == 1


# --------------------------------------------- retired mode fails closed


def test_m5_mode_list_containing_lite_refuses_startup(monkeypatch):
    monkeypatch.setenv("KS4P_MODE", "lite,graphics")
    with pytest.raises(SettingsError, match="KS4P_MODE"):
        Settings.from_env()


def test_m5_mode_only_lite_refuses_startup(monkeypatch):
    monkeypatch.setenv("KS4P_MODE", "lite,")
    with pytest.raises(SettingsError, match="KS4P_MODE"):
        Settings.from_env()


# ------------------------------------------- explicit miss is NOT_FOUND


def test_m6_explicit_shape_miss_is_not_found(make_deck):
    deck = make_deck("m6.pptx")
    out = _fn("apply_edits")(
        file_path=str(deck),
        edits=[{"op": "set_text", "slide": 0, "shape": 99999, "text": "x"}],
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "NOT_FOUND"
    assert "stale" not in out["error"]["hint"].lower()


def test_m6_anchor_miss_still_stale_anchor(make_deck):
    deck = make_deck("m6b.pptx")
    out = _fn("apply_edits")(
        file_path=str(deck),
        edits=[{"op": "set_text", "anchor": "a:fffffff", "text": "x"}],
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "STALE_ANCHOR"


# --------------------------------------------- directory is not "locked"


def test_m7_directory_path_is_bad_params(tmp_path):
    out = _fn("get_presentation_info")(file_path=str(tmp_path))
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "directory" in out["error"]["message"]
    assert "PowerPoint" not in out["error"]["hint"]


# ------------------------------------- no pack hint on echoed user input


def test_m8_echoed_tool_name_no_pack_hint(make_deck):
    deck = make_deck("m8.pptx")
    for bad_kind in ("create_chart", "set_shape", "merge_cells",
                     "apply_table_style"):
        out = _fn("list_elements")(file_path=str(deck), kind=bad_kind)
        assert out["ok"] is False
        assert out["error"]["code"] == "BAD_PARAMS"
        assert bad_kind in out["error"]["message"]
        assert "enable_tools" not in out["error"]["hint"], (
            f"echoed input {bad_kind!r} must not trigger the pack hint"
        )


# ------------------------------------------------- a4 / letter presets


def test_m9_a4_and_letter_presets_accepted(make_deck):
    deck = make_deck("m9.pptx")
    out = _fn("set_slide_size")(file_path=str(deck), preset="a4")
    assert out["ok"] is True
    pkg = PptxPackage(deck)
    sldsz = pkg.presentation().find(qn("p:sldSz"))
    assert (sldsz.get("cx"), sldsz.get("cy")) == ("9906000", "6858000")
    assert sldsz.get("type") == "A4"
    out = _fn("set_slide_size")(file_path=str(deck), preset="Letter")
    assert out["ok"] is True
    pkg = PptxPackage(deck)
    sldsz = pkg.presentation().find(qn("p:sldSz"))
    assert sldsz.get("type") == "letter"


# ---------------------------------------------------- atomic stays guarded


def test_l1_atomic_still_guarded_without_stale_v1_copy(make_deck):
    desc = inspect.getdoc(server.apply_edits) or ""
    assert "atomic" in desc.lower()
    assert "v1" not in desc.lower()
    deck = make_deck("l1.pptx")
    out = _fn("apply_edits")(
        file_path=str(deck),
        edits=[{"op": "set_text", "slide": 0, "shape": 2, "text": "x"}],
        atomic=False,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "atomic" in out["error"]["message"]


# ------------------------------------------- over-the-wire behavior


@pytest.mark.timeout(90)
def test_m4_fixed_graphics_tool_dispatches_operation(make_deck, tmp_path):
    from test_stdio_roundtrip import _Server, _handshake

    deck_src = make_deck("m4_src.pptx")
    deck = tmp_path / "m4.pptx"
    shutil.copy2(deck_src, deck)
    srv = _Server()
    try:
        _handshake(srv)
        result = srv.request(
            "tools/call",
            {"name": "ppt_graphics",
             "arguments": {"op": "insert_shape", "args": {
                 "file_path": str(deck), "slide": 0,
                 "shape_type": "rect", "x": 1, "y": 1,
                 "w": 2, "h": 1,
             }}},
        )
        assert not result.get("isError"), result
        text = "".join(c.get("text", "") for c in result.get("content", []))
        payload = json.loads(text)
        assert payload["ok"] is True
        assert payload["changed"]["shape_id"]
    finally:
        srv.close()


@pytest.mark.timeout(60)
def test_m5_stdio_legacy_mode_refuses_startup(make_deck):
    from test_stdio_roundtrip import _Server

    srv = _Server(env_extra={"KS4P_MODE": "lite,graphics"})
    try:
        assert srv.proc.wait(timeout=30) == 2
        srv._stderr_reader.join(timeout=5)
        stderr = "\n".join(srv.stderr_lines)
        assert "KS4P_MODE" in stderr
        assert "KS4P_READ_ONLY" in stderr
        assert "KS4P_FORBID" in stderr
    finally:
        srv.close()
