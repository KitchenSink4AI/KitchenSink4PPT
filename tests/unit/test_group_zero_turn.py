"""A group saved with rot="0" or flipH="0" is neither rotated nor flipped.

rot="0", flipH="0", flipV="0" and the words "false" are explicit spellings
of "no turn", and producers other than PowerPoint write them. Two checks
asked only whether the attribute was there: ungroup_shapes refused such a
group as "rotated or flipped", and moving one of its members carried the
warning that an ancestor group rotates or flips. A zero turn is no turn
now; a real rotation or flip still refuses and still warns.

The deck is made up and built by code.
"""

from __future__ import annotations

import pytest

from kitchensink4ppt.core.errors import UnsupportedStructure
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import geometry as g
from kitchensink4ppt.ops import shapes as shp
from kitchensink4ppt.ops import slides as sl
from kitchensink4ppt.ops.read import get_slide_info

NO_TURN = [
    {"rot": "0"},
    {"flipH": "0", "flipV": "0"},
    {"flipH": "false", "flipV": "false"},
    {"rot": "0", "flipH": "0", "flipV": "false"},
]
TURNED = [
    {"rot": "5400000"},
    {"flipH": "1"},
    {"flipV": "true"},
    {"rot": "0", "flipH": "true"},
]
ROTATION_WARNING = "rotation or flip"


def _grouped(make_deck, tmp_path, attrs: dict, name: str):
    """A saved deck with one group of two rectangles whose group xfrm
    carries `attrs`; returns (path, slide, group id, member ids)."""
    path = make_deck(name, extra_slides=0)
    pkg = PptxPackage(path)
    slide = sl.insert_slide(pkg, 0)["index"]
    a = shp.insert_shape(pkg, slide, "rect", 1, 1, 1, 1)["shape_id"]
    b = shp.insert_shape(pkg, slide, "rect", 3, 2, 1.5, 1)["shape_id"]
    gid = shp.group_shapes(pkg, slide, [a, b])["group_id"]
    shp.set_shape(pkg, slide, gid, w=2.5, h=1.5)  # a scaled group
    part = get_slide_info(pkg, slide)["part"]
    grp, _chain = shp._find_shape(pkg, part, gid)
    xfrm = grp.find(f"{qn('p:grpSpPr')}/{qn('a:xfrm')}")
    for key, value in attrs.items():
        xfrm.set(key, value)
    pkg.mark_dirty(part)
    pkg.save(do_backup=False)
    return path, slide, gid, (a, b)


def _slide_boxes(pkg, slide, ids):
    part = get_slide_info(pkg, slide)["part"]
    out = []
    for sid in ids:
        elem, chain = shp._find_shape(pkg, part, sid)
        out.append(shp._slide_box(elem, chain))
    return out


@pytest.mark.parametrize("attrs", NO_TURN, ids=lambda d: " ".join(f"{k}={v}" for k, v in d.items()))
def test_ungroup_a_group_with_a_zero_turn(make_deck, tmp_path, attrs):
    path, slide, gid, members = _grouped(make_deck, tmp_path, attrs, "zero.pptx")
    pkg = PptxPackage(path)
    before = _slide_boxes(pkg, slide, members)
    res = shp.ungroup_shapes(pkg, slide, gid)  # refused before the fix as rotated or flipped
    assert sorted(res["freed_ids"]) == sorted(members)
    after = _slide_boxes(pkg, slide, members)
    for old, new in zip(before, after):
        assert new == pytest.approx(old, abs=2)
    pkg.save(do_backup=False)


@pytest.mark.parametrize("attrs", NO_TURN, ids=lambda d: " ".join(f"{k}={v}" for k, v in d.items()))
def test_moving_a_member_of_a_zero_turn_group_does_not_warn(make_deck, tmp_path, attrs):
    path, slide, _gid, (a, _b) = _grouped(make_deck, tmp_path, attrs, "member.pptx")
    pkg = PptxPackage(path)
    res = shp.set_shape(pkg, slide, a, x=1.25, y=1.5)
    assert not any(ROTATION_WARNING in w for w in res.get("warnings", []))
    box = _slide_boxes(pkg, slide, [a])[0]
    assert box[0] == pytest.approx(g.in_to_emu(1.25), abs=2)
    assert box[1] == pytest.approx(g.in_to_emu(1.5), abs=2)
    pkg.save(do_backup=False)


@pytest.mark.parametrize("attrs", TURNED, ids=lambda d: " ".join(f"{k}={v}" for k, v in d.items()))
def test_a_real_turn_still_refuses_and_warns(make_deck, tmp_path, attrs):
    path, slide, gid, (a, _b) = _grouped(make_deck, tmp_path, attrs, "turned.pptx")
    pkg = PptxPackage(path)
    with pytest.raises(UnsupportedStructure, match="rotated or flipped"):
        shp.ungroup_shapes(pkg, slide, gid)
    res = shp.set_shape(pkg, slide, a, dx=0.25)
    assert any(ROTATION_WARNING in w for w in res.get("warnings", []))
