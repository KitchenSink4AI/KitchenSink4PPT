"""Who a comment or reply is signed by when the call names no author.

The user's own name is the default: the name their own PowerPoint signs
with (the Office user name and initials), else the Windows sign-in name,
else "Author", and never an email address. KS4P_AUTHOR=assistant signs
"AI Assistant via KS4PPT" instead. KS4P_COMMENT_AUTHOR names an author
outright and wins over both; a call's own author wins over everything.
Every answer says who signed (author) and why (author_source), and an
unknown KS4P_AUTHOR stops the server from starting with a plain message.
"""

from __future__ import annotations

import asyncio
import importlib

import pytest

from kitchensink4ppt import server  # noqa: F401  (registers the operations)
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import comments as cm
from kitchensink4ppt.settings import Settings, SettingsError
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.plan import build

ASSISTANT_LABEL = "AI Assistant via KS4PPT"


@pytest.fixture
def office(monkeypatch):
    """The Office user name and initials the user's own PowerPoint signs
    with (the registry is never read in these tests), with the author
    settings cleared."""
    for var in ("KS4P_AUTHOR", "KS4P_COMMENT_AUTHOR"):
        monkeypatch.delenv(var, raising=False)
    monkeypatch.setenv("USERNAME", "jdoe")
    values = {"UserName": "Jane Doe", "UserInitials": "JD"}
    try:
        module = importlib.import_module("kitchensink4ppt.core.authorship")
    except ImportError:
        module = None
    if module is not None:
        monkeypatch.setattr(module, "_office_value", lambda name: values.get(name))
    return values


def _call(op: str, args: dict, settings: Settings | None = None):
    plan = build(settings or Settings())
    payload, is_error = asyncio.run(
        dispatch(plan, plan.specs[op].group, {"op": op, "args": args}))
    assert is_error is False, payload
    return payload["changed"] if "changed" in payload else payload


def _author_entry(deck, name: str) -> dict:
    pkg = PptxPackage(deck)
    for entry in pkg.root("ppt/authors.xml").findall(f"{{{cm.NS_P188}}}author"):
        if entry.get("name") == name:
            return dict(entry.attrib)
    raise AssertionError(f"no author entry named {name!r}")


def test_a_comment_is_signed_with_the_users_own_office_name_by_default(make_deck, office):
    deck = make_deck("signed_user.pptx")
    res = _call("add_comment", {"file_path": str(deck), "slide": 0, "text": "Check this"})
    assert res["author"] == "Jane Doe", "the default must be the user's own name"
    assert res["author_source"] == "office_username"
    assert _author_entry(deck, "Jane Doe")["initials"] == "JD"
    listed = _call("list_comments", {"file_path": str(deck)})
    assert listed["slides"][0]["comments"][0]["author"] == "Jane Doe"


def test_a_reply_is_signed_with_the_users_own_name_by_default(make_deck, office):
    deck = make_deck("signed_reply.pptx")
    root = _call("add_comment", {"file_path": str(deck), "slide": 0, "text": "Root",
                                 "author": "Alex Kim"})
    res = _call("reply_to_comment", {"file_path": str(deck), "slide": 0,
                                     "comment_id": root["comment_id"], "text": "Agreed"})
    assert res["author"] == "Jane Doe" and res["author_source"] == "office_username"
    assert root["author_source"] == "explicit"


def test_without_an_office_name_the_windows_sign_in_name_signs(make_deck, office):
    office.clear()
    deck = make_deck("signed_windows.pptx")
    res = _call("add_comment", {"file_path": str(deck), "slide": 0, "text": "Check this"})
    assert res["author"] == "jdoe" and res["author_source"] == "windows_signin"


def test_an_email_address_never_signs_a_comment(make_deck, office):
    office["UserName"] = "jane.doe@example.com"
    deck = make_deck("signed_no_email.pptx")
    res = _call("add_comment", {"file_path": str(deck), "slide": 0, "text": "Check this"})
    assert "@" not in res["author"]
    assert res["author"] == "jdoe" and res["author_source"] == "windows_signin"


def test_with_no_name_anywhere_the_comment_is_signed_author(make_deck, office, monkeypatch):
    office.clear()
    for var in ("USERNAME", "USER", "LOGNAME"):
        monkeypatch.delenv(var, raising=False)
    deck = make_deck("signed_fallback.pptx")
    res = _call("add_comment", {"file_path": str(deck), "slide": 0, "text": "Check this"})
    assert res["author"] == "Author" and res["author_source"] == "fallback"


def test_the_assistant_setting_signs_the_assistant_label(make_deck, office, monkeypatch):
    monkeypatch.setenv("KS4P_AUTHOR", "assistant")
    deck = make_deck("signed_assistant.pptx")
    res = _call("add_comment", {"file_path": str(deck), "slide": 0, "text": "Check this"},
                Settings.from_env())
    assert res["author"] == ASSISTANT_LABEL and res["author_source"] == "setting"
    assert _author_entry(deck, ASSISTANT_LABEL)["initials"] == "AI"
    reply = _call("reply_to_comment", {"file_path": str(deck), "slide": 0,
                                       "comment_id": res["comment_id"], "text": "Done"})
    assert reply["author"] == ASSISTANT_LABEL and reply["author_source"] == "setting"


def test_a_named_comment_author_wins_over_both_defaults(make_deck, office, monkeypatch):
    monkeypatch.setenv("KS4P_AUTHOR", "assistant")
    monkeypatch.setenv("KS4P_COMMENT_AUTHOR", "Review Desk")
    deck = make_deck("signed_named.pptx")
    res = _call("add_comment", {"file_path": str(deck), "slide": 0, "text": "Check this"})
    assert res["author"] == "Review Desk" and res["author_source"] == "named_setting"


def test_a_calls_own_author_wins_over_everything(make_deck, office, monkeypatch):
    monkeypatch.setenv("KS4P_AUTHOR", "assistant")
    monkeypatch.setenv("KS4P_COMMENT_AUTHOR", "Review Desk")
    deck = make_deck("signed_explicit.pptx")
    res = _call("add_comment", {"file_path": str(deck), "slide": 0, "text": "Check this",
                                "author": "Sam Lee"})
    assert res["author"] == "Sam Lee" and res["author_source"] == "explicit"


@pytest.mark.parametrize("raw, mode", [("user", "user"), ("assistant", "assistant"),
                                       (" Assistant ", "assistant"), ("", "user")])
def test_the_author_setting_reads_user_or_assistant(raw, mode):
    assert Settings.from_env({"KS4P_AUTHOR": raw}).author_mode == mode


def test_an_unknown_author_setting_stops_the_start_with_a_plain_message():
    with pytest.raises(SettingsError) as refused:
        Settings.from_env({"KS4P_AUTHOR": "robot"})
    message = str(refused.value)
    assert "KS4P_AUTHOR" in message and "'robot'" in message
    assert "user" in message and "assistant" in message and ASSISTANT_LABEL in message


@pytest.mark.parametrize("op", ["add_comment", "reply_to_comment"])
def test_the_help_names_who_signs_when_no_author_is_given(op):
    from kitchensink4ppt.surface.help import help_payload

    author = help_payload(build(Settings()), op)["args_schema"]["properties"]["author"]
    text = author.get("description", "")
    assert "KS4P_AUTHOR" in text and "own name" in text and ASSISTANT_LABEL in text


def test_an_empty_author_means_the_default_and_a_blank_one_is_refused(make_deck, office):
    deck = make_deck("signed_empty.pptx")
    res = _call("add_comment", {"file_path": str(deck), "slide": 0, "text": "Check this",
                                "author": ""})
    assert res["author"] == "Jane Doe" and res["author_source"] == "office_username"
    plan = build(Settings())
    payload, is_error = asyncio.run(dispatch(plan, "ppt_review", {"op": "add_comment", "args": {
        "file_path": str(deck), "slide": 0, "text": "Check this", "author": "   "}}))
    assert is_error is True and "non-empty name" in str(payload), payload


# An email address never signs a comment or a reply, whoever supplies it: a
# model that passed the user's email as the author would otherwise write it
# into the deck for everyone who opens it.


def _answer(op: str, args: dict) -> dict:
    """The whole answer, warnings included."""
    plan = build(Settings())
    payload, is_error = asyncio.run(
        dispatch(plan, plan.specs[op].group, {"op": op, "args": args}))
    assert is_error is False, payload
    return payload


def _signed_names(deck) -> list[str]:
    pkg = PptxPackage(deck)
    return [entry.get("name") for entry in
            pkg.root("ppt/authors.xml").findall(f"{{{cm.NS_P188}}}author")]


def _email_notes(answer: dict) -> list[str]:
    return [w for w in answer.get("warnings") or [] if "email address" in w]


@pytest.mark.parametrize("sent", ["jane.doe@example.com", "Jane Doe <jane.doe@example.com>",
                                  " jd@example.com "])
def test_an_email_sent_as_the_author_never_signs_a_comment(make_deck, office, sent):
    deck = make_deck("signed_sent_email.pptx")
    answer = _answer("add_comment", {"file_path": str(deck), "slide": 0,
                                     "text": "Check this", "author": sent})
    res = answer["changed"]
    assert res["author"] == "Jane Doe" and res["author_source"] == "office_username"
    assert _signed_names(deck) == ["Jane Doe"]
    notes = _email_notes(answer)
    assert len(notes) == 1, answer.get("warnings")
    assert '"Jane Doe"' in notes[0] and "@" not in notes[0]


def test_an_email_sent_as_the_author_never_signs_a_reply(make_deck, office):
    deck = make_deck("signed_reply_email.pptx")
    first = _call("add_comment", {"file_path": str(deck), "slide": 0, "text": "Check this"})
    answer = _answer("reply_to_comment", {
        "file_path": str(deck), "slide": 0, "comment_id": first["comment_id"],
        "text": "Done", "author": "jane.doe@example.com"})
    res = answer["changed"]
    assert res["author"] == "Jane Doe" and res["author_source"] == "office_username"
    assert all("@" not in name for name in _signed_names(deck))
    assert len(_email_notes(answer)) == 1


def test_an_email_author_falls_back_through_the_settings_as_usual(make_deck, office,
                                                                  monkeypatch):
    deck = make_deck("signed_email_settings.pptx")
    args = {"file_path": str(deck), "slide": 0, "text": "Check this",
            "author": "jane.doe@example.com"}
    monkeypatch.setenv("KS4P_COMMENT_AUTHOR", "Review Desk")
    res = _answer("add_comment", args)["changed"]
    assert res["author"] == "Review Desk" and res["author_source"] == "named_setting"
    monkeypatch.delenv("KS4P_COMMENT_AUTHOR")
    monkeypatch.setenv("KS4P_AUTHOR", "assistant")
    answer = _answer("add_comment", args)
    assert answer["changed"]["author"] == ASSISTANT_LABEL
    assert answer["changed"]["author_source"] == "setting"
    assert ASSISTANT_LABEL in _email_notes(answer)[0]
    assert all("@" not in name for name in _signed_names(deck))


def test_an_email_in_the_named_setting_never_signs_either(make_deck, office, monkeypatch):
    monkeypatch.setenv("KS4P_COMMENT_AUTHOR", "desk@example.com")
    deck = make_deck("signed_setting_email.pptx")
    answer = _answer("add_comment", {"file_path": str(deck), "slide": 0, "text": "Check this"})
    res = answer["changed"]
    assert res["author"] == "Jane Doe" and res["author_source"] == "office_username"
    assert _signed_names(deck) == ["Jane Doe"]
    notes = _email_notes(answer)
    assert len(notes) == 1 and "KS4P_COMMENT_AUTHOR" in notes[0]


def test_a_name_signs_without_a_note(make_deck, office):
    deck = make_deck("signed_plain.pptx")
    answer = _answer("add_comment", {"file_path": str(deck), "slide": 0,
                                     "text": "Check this", "author": "Sam Lee"})
    assert answer["changed"]["author"] == "Sam Lee"
    assert _email_notes(answer) == []


@pytest.mark.parametrize("op", ["add_comment", "reply_to_comment"])
def test_the_help_says_an_email_address_never_signs(op):
    from kitchensink4ppt.surface.help import help_payload

    author = help_payload(build(Settings()), op)["args_schema"]["properties"]["author"]
    assert "email address" in author.get("description", "")


# author="assistant" is reserved (the family rule, as in KitchenSink4XL): it
# signs exactly the assistant label, whatever the settings are, so a user who
# asks for a comment "from the assistant" gets the label and not the word.


@pytest.mark.parametrize("sent", ["assistant", "Assistant", "  ASSISTANT  "])
def test_author_assistant_signs_the_assistant_label(make_deck, office, monkeypatch, sent):
    monkeypatch.setenv("KS4P_COMMENT_AUTHOR", "Review Desk")
    deck = make_deck("signed_asks_assistant.pptx")
    answer = _answer("add_comment", {"file_path": str(deck), "slide": 0,
                                     "text": "Check this", "author": sent})
    res = answer["changed"]
    assert res["author"] == ASSISTANT_LABEL, res
    assert res["author_source"] == "explicit_assistant"
    assert _author_entry(deck, ASSISTANT_LABEL)["initials"] == "AI"
    assert not answer.get("warnings")
    reply = _call("reply_to_comment", {"file_path": str(deck), "slide": 0,
                                       "comment_id": res["comment_id"], "text": "Done",
                                       "author": sent})
    assert reply["author"] == ASSISTANT_LABEL
    assert reply["author_source"] == "explicit_assistant"
    assert _signed_names(deck) == [ASSISTANT_LABEL]


@pytest.mark.parametrize("sent", ["the assistant", "Assistant Editor", "assistants"])
def test_only_the_reserved_word_asks_for_the_label(make_deck, office, sent):
    deck = make_deck("signed_near_assistant.pptx")
    res = _call("add_comment", {"file_path": str(deck), "slide": 0, "text": "Check this",
                                "author": sent})
    assert res["author"] == sent and res["author_source"] == "explicit"


@pytest.mark.parametrize("op", ["add_comment", "reply_to_comment"])
def test_the_op_line_and_help_say_how_to_sign_as_the_assistant(op):
    from kitchensink4ppt.surface import copy
    from kitchensink4ppt.surface.help import help_payload

    wanted = ('Leave author out to sign with the user\'s own name, or pass '
              'author="assistant" to sign as "AI Assistant via KS4PPT".')
    assert wanted in copy.OP_LINES[op]
    author = help_payload(build(Settings()), op)["args_schema"]["properties"]["author"]
    assert 'author="assistant"' in " ".join(author.get("description", "").split())
