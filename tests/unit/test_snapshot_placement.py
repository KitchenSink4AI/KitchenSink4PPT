"""Where create_snapshot puts its copy, what it calls it, and that nothing
removes it.

Field report: four snapshots of 20261018_2155_<deck>.pptx landed beside it
as 20261018_2200_<deck>_<label>.pptx and later. Where the newest leading
date stamp marks the current version, every snapshot outranked the working
deck, and three held superseded content. The copies also kept the deck's
modification time, so a file named 22:19 said 22:15. The automatic prev and
anchor backups were hidden; snapshots were not.

Now a snapshot goes into the deck's hidden backup folder by default, where
manage_backups lists and restores it and nothing rotates or purges it. With
dest_dir it is a visible copy named after the deck, its own name first, so
it never carries a newer leading date stamp than the deck. The time in the
name and the file's modification time are both the time the copy was
written.
"""

from __future__ import annotations

import datetime as real_dt
import os
import time
import types
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core import safesave
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import backups as bk

_DC_TITLE = "{http://purl.org/dc/elements/1.1/}title"
_CORE = "docProps/core.xml"
_STAMP = real_dt.datetime(2031, 9, 6, 12, 34, 5)


class _Frozen(real_dt.datetime):
    @classmethod
    def now(cls, tz=None):
        return cls(2031, 9, 6, 12, 34, 5)


@pytest.fixture()
def frozen(monkeypatch):
    monkeypatch.setattr(
        bk, "_dt", types.SimpleNamespace(datetime=_Frozen, date=real_dt.date))


def _mutate(doc: Path, title: str) -> None:
    pkg = PptxPackage(doc)
    root = pkg.root(_CORE)
    t = root.find(_DC_TITLE)
    if t is None:
        t = etree.SubElement(root, _DC_TITLE)
    t.text = title
    pkg.mark_dirty(_CORE)
    pkg.save()


def _visible_decks(folder: Path) -> list[str]:
    return sorted(p.name for p in folder.iterdir()
                  if p.is_file() and p.suffix == ".pptx")


# ------------------------------------------------------------ the default


def test_a_snapshot_goes_into_the_hidden_backup_folder_by_default(make_deck):
    deck = make_deck("20261018_2155_Sales_Summary.pptx")
    old_style = deck.parent / "20261018_2200_Sales_Summary_old.pptx"
    old_style.write_bytes(deck.read_bytes())  # made by an older build
    result = bk.create_snapshot(str(deck), label="with notes")
    snap = Path(result["snapshot"])
    assert snap.is_absolute() and snap.is_file()
    assert snap.parent == safesave.slot_dir(deck)
    assert snap.parent.parent.name == safesave.BACKUP_DIR_NAME
    assert snap.name.startswith("snapshot-")
    assert snap.read_bytes() == deck.read_bytes()
    assert result["hidden"] is True
    assert "hidden" in result["location"]
    # Nothing new beside the deck, and an older build's copy is left alone.
    assert _visible_decks(deck.parent) == sorted([deck.name, old_style.name])
    assert old_style.read_bytes() == deck.read_bytes()


def test_the_newest_stamp_beside_the_deck_is_still_the_deck(make_deck):
    """The convention the pilot's folder used: the highest leading stamp is
    the current version. A snapshot never takes that place."""
    deck = make_deck("20261018_2155_Sales_Summary.pptx")
    for label in ("one", "two", "three"):
        bk.create_snapshot(str(deck), label=label)
    assert max(_visible_decks(deck.parent)) == deck.name


def test_manage_backups_lists_and_restores_a_hidden_snapshot(make_deck):
    deck = make_deck("deck.pptx")
    before = deck.read_bytes()
    snap = Path(bk.create_snapshot(str(deck), label="keeper")["snapshot"])
    _mutate(deck, "changed after the snapshot")
    assert deck.read_bytes() != before
    listing = bk.manage_backups("list", file_path=str(deck))
    assert [Path(item["source"]) for item in listing["snapshots"]] == [snap]
    restored = bk.manage_backups(
        "restore", file_path=str(deck), source=listing["snapshots"][0]["source"])
    assert restored["from"] == f"file {snap.name}"
    assert deck.read_bytes() == before
    assert snap.is_file()  # restoring from it does not use it up


# ------------------------------------------------------ the time it was made


def test_the_name_and_the_modification_time_are_the_time_it_was_written(
        make_deck, frozen):
    deck = make_deck("deck.pptx")
    long_ago = time.mktime((2020, 1, 1, 0, 0, 0, 0, 0, -1))
    os.utime(deck, (long_ago, long_ago))
    snap = Path(bk.create_snapshot(str(deck))["snapshot"])
    assert "20310906_1234" in snap.name
    assert abs(snap.stat().st_mtime - _STAMP.timestamp()) < 1.0
    keep = deck.parent / "keep"
    keep.mkdir()
    visible = Path(bk.create_snapshot(str(deck), dest_dir=str(keep))["snapshot"])
    assert "2031-09-06 1234" in visible.name
    assert abs(visible.stat().st_mtime - _STAMP.timestamp()) < 1.0


# --------------------------------------------------------------- dest_dir


def test_a_dest_dir_snapshot_is_visible_and_named_after_the_deck(
        make_deck, frozen):
    deck = make_deck("20261018_2155_Sales_Summary.pptx")
    keep = deck.parent / "keep"
    keep.mkdir()
    first = bk.create_snapshot(str(deck), label="with notes", dest_dir=str(keep))
    assert Path(first["snapshot"]).name == (
        "20261018_2155_Sales_Summary (snapshot with notes 2031-09-06 1234).pptx")
    assert first["hidden"] is False
    assert Path(first["snapshot"]).parent == keep
    plain = bk.create_snapshot(str(deck), dest_dir=str(keep))
    assert Path(plain["snapshot"]).name == (
        "20261018_2155_Sales_Summary (snapshot 2031-09-06 1234).pptx")
    again = bk.create_snapshot(str(deck), dest_dir=str(keep))
    assert Path(again["snapshot"]).name == (
        "20261018_2155_Sales_Summary (snapshot 2031-09-06 1234) (2).pptx")
    # Beside the deck too, max() over the names still picks the deck.
    beside = bk.create_snapshot(str(deck), dest_dir=str(deck.parent))
    assert Path(beside["snapshot"]).name.startswith(deck.stem + " (snapshot")
    assert max(_visible_decks(deck.parent)) == deck.name


# ------------------------------------------------------- nothing removes it


def test_rotation_purges_and_sweeps_never_remove_a_hidden_snapshot(make_deck):
    deck = make_deck("deck.pptx")
    snap = Path(bk.create_snapshot(str(deck), label="keep me")["snapshot"])
    for n in range(3):  # prev and anchor rotate around it
        _mutate(deck, f"edit {n}")
    assert snap.is_file()
    bk.manage_backups("purge", file_path=str(deck), scope="slots", dry_run=False)
    assert snap.is_file()
    assert bk.manage_backups("list", file_path=str(deck))["snapshots"]
    # The stale-temp sweep looks at the same folder; an old snapshot stays.
    old = time.time() - 30 * 24 * 3600
    os.utime(snap, (old, old))
    safesave._sweep_slot_temps(str(deck), set())
    assert snap.is_file()
    # The deck goes away: its folder is an orphan, and the snapshot is all
    # that is left of it. A purge of orphans keeps it and says so.
    deck.unlink()
    out = bk.manage_backups("purge", directory=str(deck.parent),
                            scope="orphans", dry_run=False)
    assert snap.is_file()
    assert [s for kept in out["kept"] for s in kept["snapshots"]] == [str(snap)]
    assert out["deleted"] == []
    listing = bk.manage_backups("list", directory=str(deck.parent))
    orphan = listing["orphaned_folders"][0]
    assert [Path(s["source"]) for s in orphan["snapshots"]] == [snap]


def test_a_long_named_deck_stays_matched_to_its_snapshots_after_a_slot_purge(
        make_deck):
    """A deck name too long for the folder is kept in a breadcrumb file. A
    slot purge used to delete the breadcrumb even when the folder still held
    snapshots, which turned the folder into an orphan of a deck that exists."""
    name = "Q3 board review, appendix tables and the revenue bridge, final " \
           "version agreed.pptx"
    deck = make_deck(name)
    _mutate(deck, "first edit")
    snap = Path(bk.create_snapshot(str(deck))["snapshot"])
    assert snap.parent.name != deck.name  # the long-name folder
    bk.manage_backups("purge", file_path=str(deck), scope="slots", dry_run=False)
    listing = bk.manage_backups("list", directory=str(deck.parent))
    assert listing["orphaned_folders"] == []
    entry = next(p for p in listing["presentations"]
                 if Path(p["presentation"]) == deck)
    assert [Path(s["source"]) for s in entry["snapshots"]] == [snap]


# ------------------------------------------------------- what it tells you


def test_the_help_says_where_the_copy_goes():
    import inspect

    from kitchensink4ppt import server
    from kitchensink4ppt.surface import copy

    doc = inspect.getdoc(server.create_snapshot) or ""
    assert "hidden" in doc and "dest_dir" in doc
    assert "named after" in doc
    assert "hidden" in copy.OP_LINES["create_snapshot"]


#: (phrase, why it is not said) about where a visible copy lists.
RETIRED_SORT_CLAIMS = (
    ("sorts right after the deck",
     "a space sorts before a dot, so the copy lists BEFORE the deck in an "
     "ordinal listing and in Explorer with extensions shown"),
    ("never sorts ahead of it", "the same: it sorts ahead of the deck in an ordinal listing"),
    ("sorts after the deck", "the same"),
)


def test_what_it_says_about_a_visible_copy_is_true(make_deck, frozen):
    """The answer, the help and the public text said a dest_dir copy
    "sorts right after the deck and never ahead of it". In an ordinal
    listing (dir, a sorted() list, Explorer with extensions shown) it sorts
    BEFORE the deck. What holds everywhere is the reason for the name: the
    deck's whole name comes first, so the copy never carries a newer leading
    date stamp than the deck, and max() over the names still picks the deck.
    A name that sorted after the deck would make max() pick the copy, the
    field report's fault, so the name stays and the sentence changes."""
    import inspect

    from kitchensink4ppt import server
    from kitchensink4ppt.surface import copy

    deck = make_deck("20261018_2238_Sales_Summary.pptx")
    keep = deck.parent / "keep"
    keep.mkdir()
    out = bk.create_snapshot(str(deck), label="before review", dest_dir=str(keep))
    snap = Path(out["snapshot"])
    assert snap.name.startswith(deck.stem + " ")  # the deck's stamp leads
    assert max([deck.name, snap.name]) == deck.name
    root = Path(__file__).resolve().parents[2]
    said = {
        "answer": out["location"],
        "create_snapshot help": inspect.getdoc(server.create_snapshot) or "",
        "backups docstring": inspect.getdoc(bk.create_snapshot) or "",
        "dest_dir help": copy.ARG_HELP["create_snapshot"]["dest_dir"],
        "README.md": (root / "README.md").read_text(encoding="utf-8"),
        "CHANGELOG.md": (root / "CHANGELOG.md").read_text(encoding="utf-8"),
    }
    for where, text in said.items():
        flat = " ".join(text.split())
        for phrase, why in RETIRED_SORT_CLAIMS:
            assert phrase not in flat, f"{where} says {phrase!r} ({why})"
    assert "leading date stamp" in out["location"]
