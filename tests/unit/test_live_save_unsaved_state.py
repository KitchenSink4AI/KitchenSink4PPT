"""Live edits stay unsaved for the user; live_save says whose unsaved work it wrote.

An agent edited a deck open in PowerPoint, then called live_save unasked. The
result said the deck had unsaved changes, which did not say whether they were
the agent's own edit or the user's work. The server now records, at its first
live edit to a deck, whether the deck already had unsaved changes, and
live_save reports it as unsaved_before_our_edits (true, false or null when
unknown). The wording that tells a model to leave the save to the user is
pinned here too. PowerPoint is replaced by a fake: no COM, no process.
"""

from __future__ import annotations

import sys

import pytest

from kitchensink4ppt.com import live, live_ops, xproc
from kitchensink4ppt.surface import copy as ks_copy

DECK = "C:/decks/Quarterly.pptx"

windows_only = pytest.mark.skipif(
    sys.platform != "win32", reason="the live layer is Windows-only"
)


class _Pres:
    def __init__(self, path=DECK, saved=True):
        self.FullName = path
        self.Saved = saved
        self.Final = False
        self.ReadOnly = False
        self.saves = 0

    def Save(self):
        self.saves += 1
        self.Saved = True


class _PyWin:
    class com_error(Exception):
        pass


class _Com:
    """What the session scaffolding needs from COM, replaced."""

    def __init__(self, monkeypatch, tmp_path, pres):
        self.pres = pres
        # The deck file's write stamp as the live layer reads it; None = unreadable.
        self.stamp = ("saved-at-start",)
        monkeypatch.setattr(live, "_file_stamp", lambda p: self.stamp, raising=False)
        folder = tmp_path / "live-locks"
        monkeypatch.setattr(xproc, "_lock_dir", lambda: folder)
        monkeypatch.setattr(live, "_com_modules", lambda: (None, _PyWin(), None))
        monkeypatch.setattr(live, "_ensure_com", lambda pythoncom: "already")
        monkeypatch.setattr(live, "_attach_app", lambda w, p: object())
        monkeypatch.setattr(
            live, "_resolve_presentation",
            lambda pc, pw, w, app, path: (app, self.pres),
        )
        monkeypatch.setattr(live, "probe_ready", lambda *a, **k: None)
        monkeypatch.setattr(live, "_start_undo_boundary", lambda app: True)
        live._UNSAVED_BEFORE_OUR_EDITS.clear()

    def user_saves(self):
        """The user presses Ctrl+S: the deck is clean and the file was written."""
        self.pres.Saved = True
        self.stamp = (self.stamp[0] + "+",) if self.stamp else ("saved-later",)

    def edit(self, *, fail=False, dirties=True):
        def body(session):
            if fail:
                raise RuntimeError("the edit did not land")
            if dirties:
                self.pres.Saved = False
            return {"edited": True}

        return live.run_live(DECK, "live_set_text", body)


@pytest.fixture()
def com(monkeypatch, tmp_path):
    def make(saved=True):
        return _Com(monkeypatch, tmp_path, _Pres(saved=saved))

    yield make
    live._UNSAVED_BEFORE_OUR_EDITS.clear()


@windows_only
def test_a_save_after_our_edit_to_a_clean_deck_says_the_deck_was_clean(com):
    c = com(saved=True)
    c.edit()
    out = live_ops.live_save(DECK)
    assert out["unsaved_before_our_edits"] is False
    assert out["save_confirmed"] is True
    assert "save_note" not in out


@windows_only
def test_a_save_after_our_edit_to_a_deck_with_the_users_unsaved_work_says_so(com):
    c = com(saved=False)  # the user typed something first
    c.edit()
    out = live_ops.live_save(DECK)
    assert out["unsaved_before_our_edits"] is True
    assert "user's own unsaved work" in out["save_note"]


@windows_only
def test_the_answer_comes_from_the_first_edit_and_holds_for_later_ones(com):
    c = com(saved=False)
    c.edit()
    c.edit()
    c.edit()
    assert live_ops.live_save(DECK)["unsaved_before_our_edits"] is True


@windows_only
def test_a_save_with_no_edit_of_ours_does_not_know(com):
    c = com(saved=False)
    out = live_ops.live_save(DECK)
    assert out["unsaved_before_our_edits"] is None
    assert "cannot say whose" in out["save_note"]
    # That save recorded nothing: the next edit starts a fresh record.
    c.pres.Saved = True
    c.edit()
    assert live_ops.live_save(DECK)["unsaved_before_our_edits"] is False


@windows_only
def test_a_clean_deck_and_no_edit_of_ours_is_null_without_a_note(com):
    com(saved=True)
    out = live_ops.live_save(DECK)
    assert out["unsaved_before_our_edits"] is None
    assert "save_note" not in out


@windows_only
def test_after_a_save_the_next_edit_starts_a_new_record(com):
    c = com(saved=True)
    c.edit()
    assert live_ops.live_save(DECK)["unsaved_before_our_edits"] is False
    c.pres.Saved = False  # the user typed after the save
    c.edit()
    assert live_ops.live_save(DECK)["unsaved_before_our_edits"] is True


@windows_only
def test_a_deck_the_user_saved_in_between_starts_a_new_record(com):
    c = com(saved=False)
    c.edit()
    c.pres.Saved = True  # the user saved it themselves
    c.edit()
    assert live_ops.live_save(DECK)["unsaved_before_our_edits"] is False


@windows_only
def test_an_edit_that_changed_nothing_records_nothing(com):
    c = com(saved=True)
    c.edit(dirties=False)  # e.g. a replace that found no match
    c.pres.Saved = False  # the user starts typing
    c.edit()
    assert live_ops.live_save(DECK)["unsaved_before_our_edits"] is True


@windows_only
def test_a_failed_edit_that_changed_nothing_records_nothing(com):
    c = com(saved=True)
    with pytest.raises(RuntimeError):
        c.edit(fail=True)
    assert live_ops.live_save(DECK)["unsaved_before_our_edits"] is None


@windows_only
def test_a_failed_edit_that_left_the_deck_unsaved_counts_as_our_edit(com):
    c = com(saved=True)

    def body(session):
        c.pres.Saved = False
        raise RuntimeError("half done")

    with pytest.raises(RuntimeError):
        live.run_live(DECK, "live_set_text", body)
    assert live_ops.live_save(DECK)["unsaved_before_our_edits"] is False


@windows_only
def test_reads_never_record(com):
    com(saved=False)
    live.run_live(DECK, "live_get_text", lambda s: {}, mutating=False)
    assert live_ops.live_save(DECK)["unsaved_before_our_edits"] is None


@windows_only
def test_each_deck_has_its_own_record(com, monkeypatch):
    c = com(saved=False)
    c.edit()
    other = _Pres(path="C:/decks/Other.pptx", saved=True)
    monkeypatch.setattr(
        live, "_resolve_presentation",
        lambda pc, pw, w, app, path: (app, other),
    )

    def body(session):
        other.Saved = False
        return {}

    live.run_live("C:/decks/Other.pptx", "live_set_text", body)
    assert live_ops.live_save("C:/decks/Other.pptx")["unsaved_before_our_edits"] is False
    monkeypatch.setattr(
        live, "_resolve_presentation",
        lambda pc, pw, w, app, path: (app, c.pres),
    )
    assert live_ops.live_save(DECK)["unsaved_before_our_edits"] is True


@windows_only
def test_a_user_save_between_our_edits_starts_a_new_record(com):
    c = com(saved=True)
    c.edit()                  # our edit makes the clean deck unsaved: record false
    c.user_saves()            # the user saves it ...
    c.pres.Saved = False      # ... and types something new
    c.edit()                  # our next edit finds the deck unsaved
    out = live_ops.live_save(DECK)
    assert out["unsaved_before_our_edits"] is True
    assert "user's own unsaved work" in out["save_note"]


@windows_only
def test_a_user_save_and_typing_before_a_save_of_ours_cannot_be_told_apart(com):
    c = com(saved=True)
    c.edit()
    c.user_saves()
    c.pres.Saved = False      # typed after the save, and no edit of ours since
    out = live_ops.live_save(DECK)
    assert out["unsaved_before_our_edits"] is None
    assert "cannot say whose" in out["save_note"]


@windows_only
def test_edits_with_no_save_between_keep_the_first_answer(com):
    c = com(saved=True)
    c.edit()
    c.edit()
    c.edit()
    assert live_ops.live_save(DECK)["unsaved_before_our_edits"] is False


@windows_only
def test_a_deck_file_that_cannot_be_read_gives_null_not_false(com):
    c = com(saved=True)
    c.stamp = None            # no file to compare, so a save by the user would go unseen
    c.edit()
    c.edit()
    out = live_ops.live_save(DECK)
    assert out["unsaved_before_our_edits"] is None
    assert "cannot say whose" in out["save_note"]


@windows_only
def test_the_file_stamp_follows_a_real_write(tmp_path):
    f = tmp_path / "Stamp.pptx"
    f.write_bytes(b"one")
    before = live._file_stamp(_Pres(path=str(f)))
    assert before is not None
    assert live._file_stamp(_Pres(path=str(f))) == before
    f.write_bytes(b"two!")
    assert live._file_stamp(_Pres(path=str(f))) != before
    assert live._file_stamp(_Pres(path=str(tmp_path / "Missing.pptx"))) is None


def test_the_server_instructions_and_the_save_text_leave_the_save_to_the_user():
    text = ks_copy.INSTRUCTIONS
    assert "for the user to review and save" in text
    assert "only when the user asks" in text
    assert "call it last" not in text
    assert len(text) <= 2048
    sentence = ks_copy.OPEN_DECK_SENTENCE
    assert "only when the user asks" in sentence
    assert "call live_save last" not in sentence
    line = ks_copy.OP_LINES["live_save"]
    assert "only when the user asks" in line
    assert "unsaved" in line


def test_the_save_tool_help_names_the_new_field():
    from kitchensink4ppt import server

    assert "unsaved_before_our_edits" in (server.live_save.__doc__ or "")
    assert "unsaved_before_our_edits" in (live_ops.live_save.__doc__ or "")


# ---- saved_user_work: the fact that the save wrote the user's own work comes first ----

SAVED_USER_WORK_TRUE = (
    "This save also wrote unsaved changes the user had already made to this deck, "
    "besides this server's edits. Tell the user."
)
SAVED_USER_WORK_NULL = (
    "This save may also have written unsaved changes the user had made to this deck. "
    "Tell the user."
)


@windows_only
def test_saved_user_work_comes_first_with_the_exact_text_when_the_user_had_unsaved_work(com):
    c = com(saved=False)  # the user typed something first
    c.edit()
    out = live_ops.live_save(DECK)
    assert out["unsaved_before_our_edits"] is True
    assert list(out)[0] == "saved_user_work"
    assert out["saved_user_work"] == SAVED_USER_WORK_TRUE
    assert "save_note" in out and "unsaved_before_our_edits" in out


@windows_only
def test_saved_user_work_comes_first_with_the_exact_text_when_it_cannot_tell(com):
    com(saved=False)  # unsaved changes, and no edit of ours to compare with
    out = live_ops.live_save(DECK)
    assert out["unsaved_before_our_edits"] is None
    assert list(out)[0] == "saved_user_work"
    assert out["saved_user_work"] == SAVED_USER_WORK_NULL
    assert "cannot say whose" in out["save_note"]


@windows_only
def test_no_saved_user_work_when_the_deck_was_clean_before_our_edit(com):
    c = com(saved=True)
    c.edit()
    out = live_ops.live_save(DECK)
    assert out["unsaved_before_our_edits"] is False
    assert "saved_user_work" not in out


@windows_only
def test_no_saved_user_work_when_a_clean_deck_is_saved_with_no_edit_of_ours(com):
    com(saved=True)
    out = live_ops.live_save(DECK)
    assert out["unsaved_before_our_edits"] is None
    assert "saved_user_work" not in out


@windows_only
def test_the_server_answer_puts_ok_then_saved_user_work(com):
    from kitchensink4ppt import server

    c = com(saved=False)
    c.edit()
    out = server.live_save(DECK)
    assert list(out)[:2] == ["ok", "saved_user_work"]
    assert out["saved_user_work"] == SAVED_USER_WORK_TRUE


def test_the_save_help_names_saved_user_work():
    from kitchensink4ppt import server
    from kitchensink4ppt.ops import workflows

    assert "saved_user_work" in (server.live_save.__doc__ or "")
    assert "saved_user_work" in (live_ops.live_save.__doc__ or "")
    assert "saved_user_work" in ks_copy.OP_LINES["live_save"]
