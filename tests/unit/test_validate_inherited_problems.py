"""A save refuses only the damage it would add, never a problem the deck
already had.

The save check used to refuse any payload with a dangling relationship, an r:id
with no relationship or a malformed part, wherever it was, so a deck that
arrived with one (a part nothing references, or an unused relationship to a
picture that is gone) could not be changed at all, though the change did not
touch it. The check now compares with the deck as it is on disk: a problem the
deck already had is kept as it was and named in the answer's warnings; a new
one is still refused. Restoring a backup follows the same rule against the
deck it replaces. The decks are made up and built by code.
"""

from __future__ import annotations

import asyncio
import zipfile

import pytest

from kitchensink4ppt import server  # noqa: F401  (the import registers ops)
from kitchensink4ppt.core.errors import ValidationFailed
from kitchensink4ppt.core.package import RT_SLIDE, PptxPackage
from kitchensink4ppt.ops import backups, text
from kitchensink4ppt.registry import public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.plan import build

_RT_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _rewrite(path, changes: dict[str, bytes]) -> None:
    """Rewrite the deck's ZIP with some entries replaced or added."""
    with zipfile.ZipFile(path) as zf:
        entries = [(i, zf.read(i.filename)) for i in zf.infolist()]
    names = {i.filename for i, _ in entries}
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for info, data in entries:
            zf.writestr(info.filename, changes.get(info.filename, data))
        for name, data in changes.items():
            if name not in names:
                zf.writestr(name, data)


def _rels(*rels: str) -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_REL_NS}">{"".join(rels)}</Relationships>'
    ).encode()


def _with_orphan(make_deck):
    """A part nothing references, whose own rels point at a missing picture."""
    path = make_deck("orphan.pptx")
    _rewrite(path, {
        "ppt/extra/orphan1.xml": b'<?xml version="1.0"?><note>left over</note>',
        "ppt/extra/_rels/orphan1.xml.rels": _rels(
            f'<Relationship Id="rId1" Type="{_RT_IMAGE}" Target="../media/gone.png"/>'),
    })
    return path


def _with_unused_dangling_rel(make_deck):
    """Slide 1's rels carry an unused relationship to a missing picture."""
    path = make_deck("inherited.pptx")
    with zipfile.ZipFile(path) as zf:
        rels = zf.read("ppt/slides/_rels/slide1.xml.rels")
    extra = f'<Relationship Id="rId99" Type="{_RT_IMAGE}" Target="../media/gone.png"/>'
    _rewrite(path, {"ppt/slides/_rels/slide1.xml.rels":
                    rels.replace(b"</Relationships>", extra.encode() + b"</Relationships>")})
    return path


def _edit_slide_two(path) -> PptxPackage:
    pkg = PptxPackage(path)
    text.insert_textbox(pkg, 1, "A new note", 1.0, 1.0, 3.0, 1.0)
    pkg.save()
    return pkg


def test_orphan_part_with_a_dangling_rel_does_not_block_an_edit(make_deck):
    path = _with_orphan(make_deck)
    pkg = _edit_slide_two(path)
    with zipfile.ZipFile(path) as zf:
        assert zf.read("ppt/extra/orphan1.xml") == b'<?xml version="1.0"?><note>left over</note>'
    assert any("already had" in w and "orphan1.xml.rels" in w for w in pkg.save_warnings), \
        pkg.save_warnings


def test_inherited_dangling_rel_does_not_block_an_edit_and_is_named(make_deck):
    path = _with_unused_dangling_rel(make_deck)
    plan = build(Settings(), public_ops())
    args = {"file_path": str(path), "slide": 1, "text": "A new note",
            "x": 1.0, "y": 1.0, "w": 3.0, "h": 1.0}
    payload, is_error = asyncio.run(dispatch(plan, "ppt_edit", {"op": "insert_textbox", "args": args}))
    assert is_error is False, payload
    warned = " ".join(str(w) for w in payload.get("warnings", []))
    assert "already had" in warned and "rId99" in warned, payload


def test_a_new_problem_is_still_refused_beside_an_old_one(make_deck):
    path = _with_unused_dangling_rel(make_deck)
    original = path.read_bytes()
    pkg = PptxPackage(path)
    pkg.add_relationship("ppt/presentation.xml", RT_SLIDE, "slides/slide999.xml")
    with pytest.raises(ValidationFailed, match="slide999"):
        pkg.save(do_backup=False)
    assert path.read_bytes() == original


def test_restore_keeps_the_rule_against_the_deck_it_replaces(make_deck):
    path = _with_unused_dangling_rel(make_deck)
    _edit_slide_two(path)  # the prev slot now holds the deck before the edit
    out = backups.restore_backup(str(path), "prev")
    assert out["restored"] == str(path.resolve()), out
    with zipfile.ZipFile(path) as zf:
        assert b"rId99" in zf.read("ppt/slides/_rels/slide1.xml.rels")
