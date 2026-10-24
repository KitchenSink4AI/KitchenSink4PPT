"""Copies get creation ids of their own, in the attribute each kind uses.

PowerPoint marks every slide, notes page, layout and master with
p14:creationId val="<unsigned 32-bit integer>" and every shape with
a16:creationId id="{GUID}". When a slide was copied, both kinds were
renewed by writing a GUID into an `id` attribute. p14:creationId has no
such attribute, so the copy kept its source's val (two slides shared one
creation id, and a comment added on the copy was anchored with the
source's) and carried an attribute the schema does not define. The old
unit test read `id` on both kinds, so it passed on None against a GUID,
and the repeat test skipped on every deck built by code, since none had
creation ids.

The decks are made up and built by code, with creation ids written where
and how PowerPoint writes them.
"""

from __future__ import annotations

import posixpath
import random
import shutil
import uuid
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, qn, rels_name
from kitchensink4ppt.ops import comments as cm
from kitchensink4ppt.ops import slides as sl
from kitchensink4ppt.ops.assembly import merge_decks
from kitchensink4ppt.ops.interdeck import copy_slide_between
from kitchensink4ppt.ops.masters import create_layout

P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
A16 = "http://schemas.microsoft.com/office/drawing/2014/main"
SLIDE_EXT = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
SHAPE_EXT = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"
P14_ID = f"{{{P14}}}creationId"
A16_ID = f"{{{A16}}}creationId"
_NV = ("p:nvSpPr", "p:nvPicPr", "p:nvGraphicFramePr", "p:nvGrpSpPr", "p:nvCxnSpPr")
_KINDS = ("ppt/slides/", "ppt/notesSlides/", "ppt/slideLayouts/", "ppt/slideMasters/")
RT_NOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"


def _xml_parts(pkg: PptxPackage, kinds=_KINDS) -> list[str]:
    return [p for p in pkg.part_names()
            if p.endswith(".xml") and p.startswith(kinds) and "/_rels/" not in p]


def _stamp(path: Path, seed: int) -> Path:
    """Write creation ids where and how PowerPoint does (measured on decks
    PowerPoint made): a p14 val in the extLst of p:cSld (the last child of
    cSld) for every slide, notes page, layout and master, and an a16 GUID
    in the cNvPr extLst of every shape."""
    rng = random.Random(seed)
    pkg = PptxPackage(path)
    for part in _xml_parts(pkg):
        root = pkg.root(part)
        if root.find(f".//{P14_ID}") is None:  # the template's layouts carry theirs
            csld = root.find(qn("p:cSld"))
            ext_lst = csld.find(qn("p:extLst"))
            if ext_lst is None:
                ext_lst = etree.SubElement(csld, qn("p:extLst"))
            ext = etree.SubElement(ext_lst, qn("p:ext"), uri=SLIDE_EXT)
            etree.SubElement(ext, P14_ID, nsmap={"p14": P14},
                             val=str(rng.randrange(1, 2**32)))
        for tag in _NV:
            for nv in root.iter(qn(tag)):
                cnv = nv.find(qn("p:cNvPr"))
                a_lst = cnv.find(qn("a:extLst"))
                if a_lst is None:
                    a_lst = etree.SubElement(cnv, qn("a:extLst"))
                a_ext = etree.SubElement(a_lst, qn("a:ext"), uri=SHAPE_EXT)
                guid = str(uuid.UUID(int=rng.getrandbits(128))).upper()
                etree.SubElement(a_ext, A16_ID, nsmap={"a16": A16}, id="{" + guid + "}")
        pkg.mark_dirty(part)
    pkg.save(do_backup=False)
    return path


@pytest.fixture()
def stamped(make_deck):
    return _stamp(make_deck("stamped.pptx"), seed=7)


def _val(pkg: PptxPackage, part: str) -> str:
    (el,) = list(pkg.root(part).iter(P14_ID))
    return el.get("val")


def _guids(pkg: PptxPackage, part: str) -> set[str]:
    return {el.get("id") for el in pkg.root(part).iter(A16_ID)}


def _notes_of(pkg: PptxPackage, slide_part: str) -> str:
    for rel in pkg.root(rels_name(slide_part)):
        if rel.get("Type") == RT_NOTES:
            return posixpath.normpath(posixpath.join(
                posixpath.dirname(slide_part), rel.get("Target")))
    raise AssertionError(f"{slide_part} has no notes page")


def _assert_distinct_and_well_formed(pkg: PptxPackage) -> None:
    """Every p14:creationId carries only val, an unsigned 32-bit integer
    above zero, and no two slides (no two layouts) share one."""
    for part in _xml_parts(pkg):
        for el in pkg.root(part).iter(P14_ID):
            assert set(el.attrib) == {"val"}, (part, dict(el.attrib))
            assert 0 < int(el.get("val")) < 2**32, (part, el.get("val"))
    for kind in ("ppt/slides/", "ppt/slideLayouts/", "ppt/notesSlides/"):
        vals = [_val(pkg, p) for p in _xml_parts(pkg, (kind,))]
        assert len(set(vals)) == len(vals), (kind, vals)


def test_the_stamped_deck_is_distinct_to_begin_with(stamped):
    _assert_distinct_and_well_formed(PptxPackage(stamped))


def test_duplicate_slide_gives_the_copy_its_own_val(stamped):
    pkg = PptxPackage(stamped)
    src = pkg.slide_parts()[0]
    res = sl.duplicate_slide(pkg, 0)
    pkg.save(do_backup=False)

    pkg = PptxPackage(stamped)
    assert _val(pkg, res["part"]) != _val(pkg, src)  # before the fix: the source's val
    assert _val(pkg, _notes_of(pkg, res["part"])) != _val(pkg, _notes_of(pkg, src))
    assert _guids(pkg, res["part"]) and not (_guids(pkg, res["part"]) & _guids(pkg, src))
    assert len(_guids(pkg, res["part"])) == len(_guids(pkg, src))
    _assert_distinct_and_well_formed(pkg)  # before the fix: an id attribute on p14:creationId


def test_repeated_duplication_keeps_every_creation_id_distinct(stamped):
    pkg = PptxPackage(stamped)
    parts = [pkg.slide_parts()[1]]
    for _ in range(3):
        parts.append(sl.duplicate_slide(pkg, 1)["part"])
    pkg.save(do_backup=False)

    pkg = PptxPackage(stamped)
    _assert_distinct_and_well_formed(pkg)
    sets = [_guids(pkg, p) for p in parts]
    assert all(sets)
    for i in range(len(sets)):
        for j in range(i + 1, len(sets)):
            assert not (sets[i] & sets[j]), (parts[i], parts[j])


def test_a_comment_on_the_copy_is_anchored_to_the_copy(stamped):
    pkg = PptxPackage(stamped)
    src = pkg.slide_parts()[0]
    res = sl.duplicate_slide(pkg, 0)
    cm.add_comment(pkg, res["index"], "Check the second figure", author="Pat Example")
    pkg.save(do_backup=False)

    pkg = PptxPackage(stamped)
    marks = [el.get("cId") for part in pkg.part_names() if part.startswith("ppt/comments/")
             for el in pkg.root(part).iter() if etree.QName(el).localname == "sldMk"]
    assert marks == [_val(pkg, res["part"])]
    assert marks[0] != _val(pkg, src)  # before the fix: anchored with the source slide's id


def test_copy_slide_between_avoids_the_destination_vals(stamped, tmp_path):
    dst = tmp_path / "dst.pptx"
    shutil.copy2(stamped, dst)  # the destination already uses the source's vals
    pkg = PptxPackage(dst)
    copy_slide_between(pkg, stamped, 0)
    pkg.save(do_backup=False)
    _assert_distinct_and_well_formed(PptxPackage(dst))


def test_merge_decks_gives_every_merged_slide_its_own_val(stamped, tmp_path):
    dst = tmp_path / "dst.pptx"
    shutil.copy2(stamped, dst)
    pkg = PptxPackage(dst)
    merge_decks(pkg, [str(stamped)])
    pkg.save(do_backup=False)
    pkg = PptxPackage(dst)
    assert len(pkg.slide_parts()) == 12
    _assert_distinct_and_well_formed(pkg)


def test_create_layout_from_another_gets_its_own_val(stamped):
    pkg = PptxPackage(stamped)
    create_layout(pkg, 0, "Made-up layout", based_on=1)
    pkg.save(do_backup=False)
    _assert_distinct_and_well_formed(PptxPackage(stamped))


def test_regenerate_writes_val_on_p14_and_id_on_a16():
    root = etree.fromstring(
        f'<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        f'xmlns:p14="{P14}" xmlns:a16="{A16}">'
        f'<a:ext><a16:creationId id="{{00000000-0000-0000-0000-00000000000A}}"/></a:ext>'
        f'<p:ext><p14:creationId val="12345" id="{{0000000B-0000-0000-0000-000000000000}}"/></p:ext>'
        f'</p:sld>')
    taken = {"12345"}
    assert sl._regenerate_creation_ids(root, taken) == 2
    (p14,) = list(root.iter(P14_ID))
    (a16,) = list(root.iter(A16_ID))
    assert set(p14.attrib) == {"val"} and p14.get("val") != "12345"
    assert p14.get("val") in taken and 0 < int(p14.get("val")) < 2**32
    assert set(a16.attrib) == {"id"}
    assert a16.get("id") != "{00000000-0000-0000-0000-00000000000A}"
    assert uuid.UUID(a16.get("id").strip("{}"))
