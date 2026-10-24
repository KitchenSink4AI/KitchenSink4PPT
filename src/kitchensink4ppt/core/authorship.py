"""Who the comments and replies this server writes are signed by.

The user's own name is the default, and the assistant label is a choice:

- A call's own `author` argument always wins. The value "assistant" (any
  case, spaces around it ignored) is reserved: it signs ASSISTANT_LABEL,
  whatever the settings are. Any other name is used as given.
- Otherwise KS4P_COMMENT_AUTHOR, when set, is the name to sign with.
- Otherwise the launch setting KS4P_AUTHOR decides. `user` (the default)
  signs with the user's Office user name, under
  HKCU\\Software\\Microsoft\\Office\\Common\\UserInfo, with its initials.
  `assistant` signs ASSISTANT_LABEL.
- With no Office user name set, the Windows sign-in name stands in; without
  that either, FALLBACK_NAME.

A value that is an email address is never used as a name, whoever supplies
it: a call's own author, KS4P_COMMENT_AUTHOR, or the Office and Windows
names. The next place in that order stands in for it, and when the call or
KS4P_COMMENT_AUTHOR supplied the email, the answer says so in one sentence.

Every answer that signs a comment says who signed it (`author`) and why
(`author_source`, one of SOURCES).

The deck's own author properties (dc:creator and cp:lastModifiedBy, which
set_document_properties writes) keep the email rule too: an address sent
for either is set aside and the name a comment would be signed with stands
in (property_name_with_note). They are properties, not signatures, so the
word "assistant" is written there as given.
"""

from __future__ import annotations

import os
import sys

#: The label the assistant setting signs with, and its initials.
ASSISTANT_LABEL = "AI Assistant via KS4PPT"
ASSISTANT_INITIALS = "AI"

#: The name used when neither Office nor Windows names the user.
FALLBACK_NAME = "Author"

#: The launch setting that chooses between the user's name and the label,
#: and the setting that names the author outright.
ENV_AUTHOR = "KS4P_AUTHOR"
ENV_COMMENT_AUTHOR = "KS4P_COMMENT_AUTHOR"

#: KS4P_AUTHOR's values; `user` is the default.
USER, ASSISTANT = "user", "assistant"
MODES = (USER, ASSISTANT)

#: Where a signing name came from, reported as `author_source`.
EXPLICIT = "explicit"                # the call named an author
EXPLICIT_ASSISTANT = "explicit_assistant"  # the call asked for the label: author="assistant"
NAMED_SETTING = "named_setting"      # KS4P_COMMENT_AUTHOR named one
SETTING = "setting"                  # KS4P_AUTHOR=assistant chose the label
OFFICE_USERNAME = "office_username"  # the user's Office user name
WINDOWS_SIGNIN = "windows_signin"    # the Windows sign-in name stood in
FALLBACK = "fallback"                # neither named the user
SOURCES = (EXPLICIT, EXPLICIT_ASSISTANT, NAMED_SETTING, SETTING, OFFICE_USERNAME,
           WINDOWS_SIGNIN, FALLBACK)

_USER_INFO = r"Software\Microsoft\Office\Common\UserInfo"


def _office_value(name: str) -> str | None:
    """One value under the Office UserInfo key, or None."""
    if sys.platform != "win32":
        return None
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _USER_INFO) as key:
            value, vtype = winreg.QueryValueEx(key, name)
    except OSError:
        return None
    if vtype not in (1, 2) or not isinstance(value, str):  # REG_SZ, REG_EXPAND_SZ
        return None
    value = value.strip()
    return value or None


def asks_for_assistant(author: object) -> bool:
    """True for the reserved author value: "assistant", in any case, with
    spaces around it ignored."""
    return isinstance(author, str) and author.strip().casefold() == ASSISTANT


def is_email(value: object) -> bool:
    """True when ``value`` would put an email address into a deck: it holds
    an "@". The same test covers every name this server writes: a comment's
    or a reply's author, anonymize_deck's replacement and the deck's author
    properties."""
    return "@" in str(value or "")


def _a_name(value: str | None) -> str | None:
    """``value`` when it can sign a comment: not empty and not an email
    address."""
    value = (value or "").strip()
    if not value or is_email(value):
        return None
    return value


def user_identity() -> tuple[str, str | None, str]:
    """(name, initials or None, source): the user's own name, and where it
    was read from."""
    name = _a_name(_office_value("UserName"))
    if name:
        return name, _a_name(_office_value("UserInitials")), OFFICE_USERNAME
    for var in ("USERNAME", "USER", "LOGNAME"):
        value = _a_name(os.environ.get(var))
        if value:
            return value, None, WINDOWS_SIGNIN
    return FALLBACK_NAME, None, FALLBACK


def mode_from(raw: str | None) -> str:
    """KS4P_AUTHOR's value as one of MODES. Raises ValueError for any other
    value; the launch settings turn that into a refusal to start."""
    value = (raw or "").strip().lower() or USER
    if value not in MODES:
        raise ValueError(value)
    return value


def mode() -> str:
    """The KS4P_AUTHOR this server runs with. The launch settings refuse to
    start with any other value, so one seen here reads as the default."""
    try:
        return mode_from(os.environ.get(ENV_AUTHOR))
    except ValueError:
        return USER


#: How each source reads in the sentence that says an email was not used.
_SOURCE_WORDS = {
    NAMED_SETTING: "the name KS4P_COMMENT_AUTHOR sets",
    SETTING: "the assistant label KS4P_AUTHOR=assistant chooses",
    OFFICE_USERNAME: "the user's Office user name",
    WINDOWS_SIGNIN: "the user's Windows sign-in name",
    FALLBACK: "the name used when neither Office nor Windows names the user",
}


def _email_note(sent: bool, setting: bool, name: str, source: str) -> str | None:
    """The one sentence an answer carries when an email address was not
    used as a name, or None when none was set aside."""
    if not (sent or setting):
        return None
    if sent and setting:
        what = (f"You sent an email address as the author and {ENV_COMMENT_AUTHOR} "
                "is one too, and an email address")
    elif sent:
        what = "You sent an email address as the author, which"
    else:
        what = f"{ENV_COMMENT_AUTHOR} is an email address, which"
    return (f'{what} is never written into a deck, so this is signed "{name}", '
            f"{_SOURCE_WORDS.get(source, source)}, instead.")


def signer_with_note(author: str | None = None) -> tuple[str, str | None, str, str | None]:
    """(name, initials or None, source, note or None) for one signed comment
    or reply. The note is the sentence the answer carries when the call's
    author or KS4P_COMMENT_AUTHOR was an email address and was not used.
    author="assistant" signs ASSISTANT_LABEL whatever the settings are."""
    if asks_for_assistant(author):
        return ASSISTANT_LABEL, ASSISTANT_INITIALS, EXPLICIT_ASSISTANT, None
    sent = author is not None and bool(str(author).strip())
    if sent:
        value = _a_name(str(author))
        if value:
            return value, None, EXPLICIT, None
    named_raw = (os.environ.get(ENV_COMMENT_AUTHOR) or "").strip()
    named = _a_name(named_raw)
    setting = bool(named_raw) and not named
    if named:
        return named, None, NAMED_SETTING, _email_note(sent, setting, named, NAMED_SETTING)
    if mode() == ASSISTANT:
        return (ASSISTANT_LABEL, ASSISTANT_INITIALS, SETTING,
                _email_note(sent, setting, ASSISTANT_LABEL, SETTING))
    name, initials, source = user_identity()
    return name, initials, source, _email_note(sent, setting, name, source)


def signer(author: str | None = None) -> tuple[str, str | None, str]:
    """(name, initials or None, source) for one signed comment or reply."""
    return signer_with_note(author)[:3]


#: How each document property reads in the sentence about an email address.
_PROPERTY_WORDS = {"author": "the author", "last_modified_by": "last_modified_by"}


def property_name_with_note(field: str) -> tuple[str, str, str]:
    """(name, source, note) for a document author property (`author` or
    `last_modified_by`) that a call sent as an email address. The name a
    comment would be signed with stands in, and the note is the one
    sentence the answer carries; it never repeats the address."""
    name, _initials, source, _note = signer_with_note(None)
    what = _PROPERTY_WORDS.get(field, field)
    note = (f"You sent an email address as {what}, which is never written into a "
            f'deck, so {what} is set to "{name}", {_SOURCE_WORDS.get(source, source)}, '
            "instead.")
    return name, source, note


__all__ = [
    "ASSISTANT", "ASSISTANT_INITIALS", "ASSISTANT_LABEL", "ENV_AUTHOR",
    "ENV_COMMENT_AUTHOR", "EXPLICIT", "EXPLICIT_ASSISTANT", "FALLBACK", "FALLBACK_NAME",
    "MODES",
    "NAMED_SETTING", "OFFICE_USERNAME", "SETTING", "SOURCES", "USER",
    "WINDOWS_SIGNIN", "asks_for_assistant", "is_email", "mode", "mode_from",
    "property_name_with_note", "signer", "signer_with_note", "user_identity",
]
