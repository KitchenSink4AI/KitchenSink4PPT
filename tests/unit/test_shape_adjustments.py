"""A preset shape's adjust handles can be read and set by name.

An arc's start and end angles, a rounded rectangle's corner and an arrow's
head are the preset's adjust handles (a:avLst a:gd). set_shape_adjustments
reads them and sets them on an existing shape: names are checked against the
preset's own handle list, angle handles take degrees, and every answer lists
the handles with their current and default values.
"""

from __future__ import annotations

import asyncio
import hashlib
from pathlib import Path

import pytest
from lxml import etree
from pptx import Presentation

from kitchensink4ppt import server  # noqa: F401  (the import registers ops)
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import shapes as shp
from kitchensink4ppt.registry import public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.help import help_payload
from kitchensink4ppt.surface.plan import build

OP = "set_shape_adjustments"


def _plan():
    return build(Settings(), public_ops())


def _call(tool: str, op: str, args: dict):
    return asyncio.run(dispatch(_plan(), tool, {"op": op, "args": args}))


def _deck(tmp_path: Path) -> str:
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    path = tmp_path / "adjust.pptx"
    prs.save(path)
    return str(path)


def _result(payload: dict) -> dict:
    """An edit's own answer sits under "changed"; a read answers flat."""
    inner = payload.get("changed")
    return inner if isinstance(inner, dict) else payload


def _insert(deck: str, shape_type: str, **extra) -> int:
    payload, is_error = _call("ppt_graphics", "insert_shape", {
        "file_path": deck, "slide": 0, "shape_type": shape_type,
        "x": 1, "y": 1, "w": 3, "h": 2, **extra})
    assert is_error is False, payload
    return _result(payload)["shape_id"]


def _gds(deck: str, shape_id: int) -> dict[str, str]:
    pkg = PptxPackage(deck)
    part = pkg.slide_parts()[0]
    elem, _chain = shp._find_shape(pkg, part, shape_id)
    geom = elem.find(f"{qn('p:spPr')}/{qn('a:prstGeom')}")
    return {gd.get("name"): gd.get("fmla") for gd in geom.iter(qn("a:gd"))}


def _md5(path: str) -> str:
    return hashlib.md5(Path(path).read_bytes()).hexdigest()


def test_reading_lists_every_handle_with_its_default_and_changes_nothing(tmp_path):
    deck = _deck(tmp_path)
    arc = _insert(deck, "arc")
    before = _md5(deck)
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": arc})
    assert is_error is False, payload
    assert payload["preset"] == "arc"
    assert "updated" not in payload and "saved" not in payload
    rows = {row["name"]: row for row in payload["adjustments"]}
    assert list(rows) == ["adj1", "adj2"]
    assert rows["adj1"]["value"] == 16200000 and rows["adj1"]["default"] == 16200000
    assert rows["adj1"]["degrees"] == 270.0
    assert rows["adj2"]["value"] == 0 and rows["adj2"]["degrees"] == 0.0
    assert _md5(deck) == before


def test_arc_angles_are_set_in_degrees_and_stored_in_the_file(tmp_path):
    deck = _deck(tmp_path)
    arc = _insert(deck, "arc")
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": arc,
        "degrees": {"adj1": 200, "adj2": 340}})
    assert is_error is False, payload
    result = _result(payload)
    assert sorted(result["updated"]) == ["adj1", "adj2"]
    rows = {row["name"]: row for row in result["adjustments"]}
    assert rows["adj1"]["value"] == 12000000 and rows["adj1"]["degrees"] == 200.0
    assert rows["adj2"]["value"] == 20400000 and rows["adj2"]["degrees"] == 340.0
    assert _gds(deck, arc) == {"adj1": "val 12000000", "adj2": "val 20400000"}


def test_negative_and_full_turn_degrees_fold_into_one_turn(tmp_path):
    deck = _deck(tmp_path)
    arc = _insert(deck, "arc")
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": arc,
        "degrees": {"adj1": -90, "adj2": 360}})
    assert is_error is False, payload
    assert _gds(deck, arc) == {"adj1": "val 16200000", "adj2": "val 0"}


def test_raw_values_are_stored_as_given_and_other_handles_are_kept(tmp_path):
    deck = _deck(tmp_path)
    arrow = _insert(deck, "arrow_right", adjustments={"adj1": 30000})
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": arrow,
        "adjustments": {"adj2": 75000}})
    assert is_error is False, payload
    assert _result(payload)["updated"] == ["adj2"]
    assert _gds(deck, arrow) == {"adj1": "val 30000", "adj2": "val 75000"}
    rows = {row["name"]: row["value"] for row in _result(payload)["adjustments"]}
    assert rows == {"adj1": 30000, "adj2": 75000}


def test_a_fraction_is_read_as_insert_shape_reads_it(tmp_path):
    deck = _deck(tmp_path)
    box = _insert(deck, "rounded_rect")
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": box,
        "adjustments": {"adj": 0.25}})
    assert is_error is False, payload
    assert _gds(deck, box) == {"adj": "val 25000"}


def test_an_unknown_handle_is_refused_with_the_presets_own_list(tmp_path):
    deck = _deck(tmp_path)
    arc = _insert(deck, "arc")
    before = _md5(deck)
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": arc,
        "adjustments": {"adj3": 5}})
    assert is_error is True
    message = payload["error"]["message"]
    assert "adj3" in message and "adj1" in message and "adj2" in message
    assert _md5(deck) == before


def test_degrees_on_a_handle_that_is_not_an_angle_is_refused(tmp_path):
    deck = _deck(tmp_path)
    box = _insert(deck, "rounded_rect")
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": box, "degrees": {"adj": 45}})
    assert is_error is True
    assert "adjustments" in payload["error"]["message"]


def test_a_fraction_for_an_angle_handle_is_refused_rather_than_guessed(tmp_path):
    deck = _deck(tmp_path)
    arc = _insert(deck, "arc")
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": arc,
        "adjustments": {"adj1": 0.5}})
    assert is_error is True
    assert "degrees" in payload["error"]["message"]


def test_a_handle_named_twice_is_refused(tmp_path):
    deck = _deck(tmp_path)
    arc = _insert(deck, "arc")
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": arc,
        "adjustments": {"adj1": 0}, "degrees": {"adj1": 90}})
    assert is_error is True
    assert "adj1" in payload["error"]["message"]


@pytest.mark.parametrize("value", [True, "wide", None, [1]])
def test_a_value_that_is_not_a_number_is_refused(tmp_path, value):
    deck = _deck(tmp_path)
    box = _insert(deck, "rounded_rect")
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": box,
        "adjustments": {"adj": value}})
    assert is_error is True
    assert payload["error"]["code"] == "BAD_PARAMS", payload
    assert "adj" in payload["error"]["message"]


def test_a_preset_without_handles_says_so(tmp_path):
    deck = _deck(tmp_path)
    rect = _insert(deck, "rect")
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": rect})
    assert is_error is False, payload
    assert payload["adjustments"] == []
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": rect,
        "adjustments": {"adj": 5000}})
    assert is_error is True
    assert "no adjust handles" in payload["error"]["message"]


def test_a_freeform_or_a_group_is_refused(tmp_path):
    deck = _deck(tmp_path)
    free = _insert(deck, "freeform", path=[{"commands": [
        ["move", 0, 0], ["line", 1, 0], ["line", 1, 1], ["close"]]}])
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": free})
    assert is_error is True
    assert "preset" in payload["error"]["message"]
    a = _insert(deck, "rect")
    b = _insert(deck, "ellipse")
    group, is_error = _call("ppt_graphics", "group_shapes", {
        "file_path": deck, "slide": 0, "ids": [a, b]})
    assert is_error is False, group
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": _result(group)["group_id"]})
    assert is_error is True
    assert "group" in payload["error"]["message"]


def test_an_elbow_connector_bend_is_a_handle_too(tmp_path):
    deck = _deck(tmp_path)
    a = _insert(deck, "rect")
    b = _insert(deck, "ellipse", )
    con, is_error = _call("ppt_graphics", "insert_connector", {
        "file_path": deck, "slide": 0, "kind": "elbow",
        "start_shape": a, "end_shape": b})
    assert is_error is False, con
    con = _result(con)
    cid = con["shape_id"] if "shape_id" in con else con["connector_id"]
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": cid,
        "adjustments": {"adj1": 25000}})
    assert is_error is False, payload
    assert _result(payload)["preset"].startswith("bentConnector")
    assert {row["name"]: row["value"] for row in _result(payload)["adjustments"]}["adj1"] == 25000


def test_slide_number_and_a_shape_name_address_the_shape(tmp_path):
    deck = _deck(tmp_path)
    _insert(deck, "arc", name="Swoosh")
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide_number": 1, "shape": "Swoosh",
        "degrees": {"adj2": 90}})
    assert is_error is False, payload
    assert {row["name"]: row["value"] for row in _result(payload)["adjustments"]}["adj2"] == 5400000


def test_the_table_covers_every_drawingml_preset_and_names_the_angles():
    from kitchensink4ppt.ops import guides

    assert len(guides.GUIDES) == 187
    assert guides.GUIDES["arc"] == (("adj1", 16200000), ("adj2", 0))
    assert guides.ANGLES["arc"] == frozenset({"adj1", "adj2"})
    assert guides.ANGLES["circularArrow"] == frozenset({"adj2", "adj3", "adj4"})
    for preset, names in guides.ANGLES.items():
        assert names <= {name for name, _default in guides.GUIDES[preset]}


def test_insert_shape_refuses_a_handle_its_preset_does_not_have(tmp_path):
    """insert_shape used to write such a handle and warn that it changed
    nothing; PowerPoint refuses to open a deck with one, so it is
    refused, naming the preset's own handles, and nothing is written."""
    deck = _deck(tmp_path)
    before = _md5(deck)
    payload, is_error = _call("ppt_graphics", "insert_shape", {
        "file_path": deck, "slide": 0, "shape_type": "arc",
        "x": 1, "y": 1, "w": 2, "h": 2, "adjustments": {"angle": 5}})
    assert is_error is True, payload
    message = payload["error"]["message"]
    assert "'angle'" in message and "adj1" in message and "adj2" in message
    assert _md5(deck) == before


# ---------------------------------------------- what PowerPoint opens
#
# Measured on PowerPoint 16 with the product's own opens-clean check (an
# invisible PowerPoint, alerts off, a full load): PowerPoint refuses to open a
# deck in which a preset shape's avLst names only some of the preset's
# handles (even the first one alone), lists them out of the preset's order,
# names a guide the preset lacks, repeats one, or holds a value past a signed
# 64-bit integer. It opens every whole list in the preset's order, whatever
# the values: it pins them when it draws. So every write stores the whole list.


def _listed(deck: str, shape_id: int) -> list[tuple[str, str]]:
    pkg = PptxPackage(deck)
    part = pkg.slide_parts()[0]
    elem, _chain = shp._find_shape(pkg, part, shape_id)
    geom = elem.find(f"{qn('p:spPr')}/{qn('a:prstGeom')}")
    return [(gd.get("name"), gd.get("fmla")) for gd in geom.iter(qn("a:gd"))]


def _set(deck: str, shape_id: int, **values) -> dict:
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": shape_id, **values})
    assert is_error is False, payload
    return payload


def test_setting_one_handle_writes_the_whole_list_in_the_presets_order(tmp_path):
    """A deck did not open in PowerPoint because set_shape_adjustments wrote
    only the handles it was sent (an arc's adj1 alone)."""
    deck = _deck(tmp_path)
    arc = _insert(deck, "arc")
    _set(deck, arc, degrees={"adj1": 90})
    assert _listed(deck, arc) == [("adj1", "val 5400000"), ("adj2", "val 0")]
    block = _insert(deck, "prst:blockArc")
    _set(deck, block, adjustments={"adj3": 30000})
    assert _listed(deck, block) == [
        ("adj1", "val 10800000"), ("adj2", "val 0"), ("adj3", "val 30000")]
    star = _insert(deck, "prst:star5")
    _set(deck, star, adjustments={"vf": 120000})
    assert _listed(deck, star) == [
        ("adj", "val 19098"), ("hf", "val 105146"), ("vf", "val 120000")]


def test_values_already_stored_are_kept_when_another_handle_is_set(tmp_path):
    deck = _deck(tmp_path)
    arc = _insert(deck, "arc")
    _set(deck, arc, degrees={"adj1": 90})
    _set(deck, arc, degrees={"adj2": 180})
    assert _listed(deck, arc) == [("adj1", "val 5400000"), ("adj2", "val 10800000")]


def _garble(deck: str, shape_id: int, gds: list[tuple[str, str]]) -> None:
    pkg = PptxPackage(deck)
    part = pkg.slide_parts()[0]
    elem, _chain = shp._find_shape(pkg, part, shape_id)
    avlst = elem.find(f"{qn('p:spPr')}/{qn('a:prstGeom')}/{qn('a:avLst')}")
    for gd in list(avlst):
        avlst.remove(gd)
    for name, fmla in gds:
        gd = etree.SubElement(avlst, qn("a:gd"))
        gd.set("name", name)
        gd.set("fmla", fmla)
    pkg.mark_dirty(part)
    pkg.save()


def test_a_list_powerpoint_refuses_is_named_on_reading_and_rewritten_whole(tmp_path):
    """A list an earlier build could leave (a handle missing, out of order, a
    guide the preset lacks, one named twice): reading names the problem, and setting
    any handle writes the whole list, keeping the first stored value of each
    known handle and dropping the rest."""
    deck = _deck(tmp_path)
    arc = _insert(deck, "arc", name="Gauge")
    _garble(deck, arc, [("adj2", "val 0"), ("zzz", "val 5"),
                        ("adj1", "val 5400000"), ("adj1", "val 1")])
    payload = _set(deck, arc)
    warnings = " ".join(payload.get("warnings", []))
    assert "PowerPoint" in warnings and "set_shape_adjustments" in warnings, payload
    _set(deck, arc, degrees={"adj2": 180})
    assert _listed(deck, arc) == [("adj1", "val 5400000"), ("adj2", "val 10800000")]
    payload = _set(deck, arc)
    assert not payload.get("warnings")


def test_insert_shape_writes_the_whole_list(tmp_path):
    deck = _deck(tmp_path)
    arc = _insert(deck, "arc", adjustments={"adj2": 5400000})
    assert _listed(deck, arc) == [("adj1", "val 16200000"), ("adj2", "val 5400000")]
    star = _insert(deck, "prst:star5", adjustments=[0.3])
    assert _listed(deck, star) == [
        ("adj", "val 30000"), ("hf", "val 105146"), ("vf", "val 110557")]
    box = _insert(deck, "rounded_rect", adjustments={"adj": 0.12})
    assert _listed(deck, box) == [("adj", "val 12000")]


@pytest.mark.parametrize("value", [2 ** 31, -(2 ** 31) - 1, 10 ** 30])
def test_a_value_past_32_bits_is_refused(tmp_path, value):
    deck = _deck(tmp_path)
    box = _insert(deck, "rounded_rect")
    before = _md5(deck)
    payload, is_error = _call("ppt_graphics", OP, {
        "file_path": deck, "slide": 0, "shape": box, "adjustments": {"adj": value}})
    assert is_error is True, payload
    assert "2147483647" in payload["error"]["message"]
    assert _md5(deck) == before
    payload, is_error = _call("ppt_graphics", "insert_shape", {
        "file_path": deck, "slide": 0, "shape_type": "rounded_rect",
        "x": 1, "y": 1, "w": 2, "h": 1, "adjustments": {"adj": value}})
    assert is_error is True and _md5(deck) == before


def test_the_32_bit_extremes_are_stored(tmp_path):
    deck = _deck(tmp_path)
    box = _insert(deck, "rounded_rect")
    _set(deck, box, adjustments={"adj": 2 ** 31 - 1})
    assert _listed(deck, box) == [("adj", "val 2147483647")]
    _set(deck, box, adjustments={"adj": -(2 ** 31)})
    assert _listed(deck, box) == [("adj", "val -2147483648")]


def test_validate_names_a_shape_whose_list_powerpoint_refuses(tmp_path, monkeypatch):
    """validate's file check says what PowerPoint will refuse, where, and how
    to repair it, before any PowerPoint is asked (here none is)."""
    from kitchensink4ppt.com import bridge

    def no_powerpoint(*args, **kwargs):
        raise RuntimeError("no PowerPoint in this test")

    monkeypatch.setattr(bridge, "com_validate_opens_clean", no_powerpoint)
    deck = _deck(tmp_path)
    arc = _insert(deck, "arc", name="Gauge")
    _garble(deck, arc, [("adj1", "val 5400000")])
    payload, is_error = _call("ppt_deck", "validate", {"file_path": deck})
    assert is_error is False, payload
    body = _result(payload)
    assert body["payload_valid"] is False
    error = body["payload_error"]
    assert "Gauge" in error and "slide 1" in error and "set_shape_adjustments" in error
    _set(deck, arc, degrees={"adj2": 90})
    payload, is_error = _call("ppt_deck", "validate", {"file_path": deck})
    assert _result(payload)["payload_valid"] is True, payload


def test_help_documents_the_operation_and_the_graphics_tool_lists_it():
    plan = _plan()
    info = help_payload(plan, OP)
    assert info["tool"] == "ppt_graphics"
    props = info["args_schema"]["properties"]
    for name in ("shape", "adjustments", "degrees"):
        assert name in props
    assert "clockwise" in props["degrees"]["description"]
    tool = plan.tools["ppt_graphics"]
    listed = tool.description + str(tool.input_schema)
    assert f"- {OP}:" in listed
