"""Reading a deck while a save swaps it in.

A save puts the new version of a deck in place with ReplaceFileW (see
``promote`` in the save module). ReplaceFileW moves the deck's old file
aside, under a name beside it, and then moves the new file in, so for a few
milliseconds the deck's path is missing. A reader that looks in that instant
waits here instead of reporting the deck as not found.

This module only looks: it lists folders, asks whether a path exists and
reads the swap journal.
Every route that reads a deck reaches it, the read-only tools included, so
nothing here may change a file.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import tempfile
import time
from pathlib import Path

from .errors import DocumentNotFound

#: Where a server writes down each swap before it starts (tests point it
#: at a folder of their own).
JOURNAL_ENV = "KS4P_SAVE_JOURNAL_DIR"

#: Pauses while a save has the old deck aside and the path is missing for an
#: instant (about 0.9 s at most).
_SWAP_POLL = (0.002, 0.005, 0.01, 0.02, 0.05, 0.1, 0.2, 0.5)


def aside_prefix(deck_name: str) -> str:
    """The start of the name a save gives a deck's old file while the new
    one is moved in, for this deck name only."""
    tag = hashlib.sha1(
        os.path.normcase(deck_name).encode("utf-8", "surrogatepass")
    ).hexdigest()[:8]
    return f".ks4p-replaced-{tag}-"


def _utf16_head(text: str, units: int) -> str:
    """The longest start of ``text`` that fits in ``units`` UTF-16 code units
    (Windows counts a name's length in those)."""
    out, used = [], 0
    for ch in text:
        used += 2 if ord(ch) > 0xFFFF else 1
        if used > units:
            break
        out.append(ch)
    return "".join(out)


def aside_name(deck_name: str, ident: str) -> str:
    """The name a save gives the deck's old file while the new one is moved
    in. It starts with the prefix that marks it as this deck's, and then says
    in words what it holds, because a save that is stopped at that instant
    leaves the deck under this name: ".ks4p-replaced-<tag>-<id> Budget.pptx
    (as it was before a save).tmp". A long deck name is cut short so that the
    whole name stays inside Windows' 255-unit limit."""
    return (f"{aside_prefix(deck_name)}{ident} {_utf16_head(deck_name, 120)} "
            "(as it was before a save).tmp")


def _extended(path: str) -> str:
    """``path`` in the form Windows needs past 260 characters."""
    if os.name != "nt" or len(path) <= 240 or path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):
        return "\\\\?\\UNC\\" + path[2:]
    return "\\\\?\\" + path


def asides(path: str | os.PathLike) -> list[str]:
    """This deck's old files that a save has aside, newest first. Normally
    none, or one for the few milliseconds of a swap. One that stays means a
    save did not finish: it holds the deck as it was before that save."""
    folder, name = os.path.split(os.path.abspath(os.fspath(path)))
    prefix = aside_prefix(name)
    try:
        found = [os.path.join(folder, entry) for entry in os.listdir(folder or ".")
                 if entry.startswith(prefix)]
    except OSError:
        return []

    def stamp(candidate: str) -> int:
        try:
            return os.stat(_extended(candidate)).st_mtime_ns
        except OSError:
            return 0

    return sorted(found, key=lambda c: (stamp(c), c), reverse=True)


#: A path component that looks like an 8.3 short name (DECKWI~1.PPT): only
#: such paths are asked for their long names.
_SHORT_COMPONENT = re.compile(r"~[0-9]")


def _looks_short(component: str) -> bool:
    return len(component) <= 12 and _SHORT_COMPONENT.search(component) is not None


def _plain(path: str) -> str:
    """``path`` absolute, without the extended-length prefix."""
    path = os.path.abspath(path)
    if path.startswith("\\\\?\\UNC\\"):
        return "\\\\" + path[8:]
    if path.startswith("\\\\?\\") and path[5:6] == ":":
        return path[4:]
    return path


def _long_names(path: str) -> str:
    """``path`` with its 8.3 short components written out long, as far as
    they name files and folders that are there."""
    if os.name != "nt":
        return path
    import ctypes

    head, tail = path, []
    buf = ctypes.create_unicode_buffer(32768)
    while True:
        if ctypes.windll.kernel32.GetLongPathNameW(_extended(head), buf, 32768):
            long_head = buf.value
            if long_head.startswith("\\\\?\\"):
                long_head = _plain(long_head)
            return os.path.join(long_head, *reversed(tail)) if tail else long_head
        parent, name = os.path.split(head)
        if not name or parent == head:
            return path
        tail.append(name)
        head = parent


def stored_name(path: str) -> tuple[str, str] | None:
    """(the name a file or folder is stored under, with its own case and
    written out long; its 8.3 short name, or ""), or None when nothing is
    there. Only reads."""
    if os.name != "nt":
        return None
    import ctypes
    from ctypes import wintypes

    class FindData(ctypes.Structure):
        _fields_ = [("attributes", wintypes.DWORD), ("created", wintypes.FILETIME),
                    ("accessed", wintypes.FILETIME), ("written", wintypes.FILETIME),
                    ("size_high", wintypes.DWORD), ("size_low", wintypes.DWORD),
                    ("reserved0", wintypes.DWORD), ("reserved1", wintypes.DWORD),
                    ("name", ctypes.c_wchar * 260), ("short", ctypes.c_wchar * 14),
                    ("unused0", wintypes.DWORD), ("unused1", wintypes.DWORD),
                    ("unused2", wintypes.WORD)]

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.FindFirstFileW.restype = wintypes.HANDLE
    k32.FindFirstFileW.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(FindData)]
    k32.FindClose.argtypes = [wintypes.HANDLE]
    if any(ch in os.path.basename(path) for ch in "*?"):
        return None
    found = FindData()
    handle = k32.FindFirstFileW(_extended(path), ctypes.byref(found))
    if not handle or handle == wintypes.HANDLE(-1).value:
        return None
    k32.FindClose(handle)
    return found.name, found.short


def deck_spelling(path: str | os.PathLike) -> str:
    """The deck's path as Windows resolves it, so that every spelling of one
    deck reads the same: absolute, with separators and the trailing dots and
    spaces Windows drops normalized, 8.3 short names written out long, and
    the deck's own name in the case it is stored with. A deck that is not
    there (a save that was stopped has it aside) and is named by its short
    name is looked up in the swap journal, which records each deck's short
    name. Only reads; a share is asked only about the deck's own path."""
    target = _plain(os.fspath(path))
    if os.name != "nt":
        return target
    if any(_looks_short(part) for part in Path(target).parts):
        target = _long_names(target)
    stored = stored_name(target)
    if stored is not None:
        return os.path.join(os.path.dirname(target), stored[0])
    folder, name = os.path.split(target)
    if _looks_short(name):
        return _journal_alias(folder, name) or target
    return target


def _journal_alias(folder: str, short: str) -> str | None:
    """The deck a journal entry recorded under this short name in this
    folder, or None. Only reads."""
    try:
        names = [n for n in os.listdir(journal_dir()) if n.endswith(".json")]
    except OSError:
        return None
    where = os.path.normcase(folder)
    for name in names:
        try:
            with open(os.path.join(journal_dir(), name), encoding="utf-8") as fh:
                record = json.loads(fh.read())
        except (OSError, ValueError):
            continue
        if not isinstance(record, dict):
            continue
        deck, alias = record.get("deck"), record.get("deck_short")
        if (isinstance(deck, str) and isinstance(alias, str) and alias
                and os.path.normcase(alias) == os.path.normcase(short)
                and os.path.normcase(os.path.dirname(os.path.abspath(deck))) == where):
            return os.path.abspath(deck)
    return None


def journal_dir() -> Path:
    """The folder of the swap journal: one small entry per swap that is
    running, or that a stopped server left behind."""
    override = os.environ.get(JOURNAL_ENV)
    if override:
        return Path(override)
    base = (os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_STATE_HOME")
            or tempfile.gettempdir())
    return Path(base) / "pptx-mcp" / "saves-in-progress"


def _copy_from_a_finished_save(path: str, aside: str) -> bool:
    """Whether the journal says ``aside`` is the old file of a save that
    went through: its new version had been moved in, so the deck was in
    its place after that save. Only reads."""
    folder = journal_dir()
    try:
        names = [n for n in os.listdir(folder) if n.endswith(".json")]
    except OSError:
        return False
    decks = {os.path.normcase(os.path.abspath(path)),
             os.path.normcase(deck_spelling(path))}
    wanted = os.path.normcase(os.path.abspath(aside))
    for name in names:
        try:
            with open(os.path.join(folder, name), encoding="utf-8") as fh:
                record = json.loads(fh.read())
        except (OSError, ValueError):
            continue
        if not isinstance(record, dict):
            continue
        fields = [record.get(k) for k in ("deck", "aside", "new")]
        if not all(isinstance(f, str) for f in fields):
            continue
        if (os.path.normcase(os.path.abspath(fields[0])) not in decks
                or os.path.normcase(os.path.abspath(fields[1])) != wanted):
            continue
        if record.get("swap") in ("made", "not-made"):
            return record["swap"] == "made"
        staged = [pair[0] for pair in record.get("staged") or []
                  if isinstance(pair, list) and pair and isinstance(pair[0], str)]
        return (not os.path.lexists(_extended(fields[2]))
                and any(os.path.lexists(_extended(s)) for s in staged))
    return False


def deck_not_found(path: str | os.PathLike, first: str) -> DocumentNotFound:
    """The refusal for a deck that is not at its path. ``first`` is the
    route's own sentence ("No file at ..."). When a save that did not finish
    left the deck aside, the answer says where it is and how it comes back;
    a read never moves it (reads change nothing), a change or a restart
    does."""
    target = os.path.abspath(os.fspath(path))
    found = asides(target)
    if not found:
        # The deck may be named by another spelling than the one its save
        # used (an 8.3 short name): look under the name as Windows resolves it.
        spelled = deck_spelling(target)
        if os.path.normcase(spelled) != os.path.normcase(target):
            found = asides(spelled)
    if not found:
        return DocumentNotFound(first)
    if _copy_from_a_finished_save(target, found[0]):
        exc = DocumentNotFound(
            f"{first}. The deck was renamed, moved or deleted after a save "
            f"of it went through. An older copy from before that save is at "
            f"{found[0]}. It is not put back, so it never replaces the deck; "
            "open it, or delete it if you do not need it.")
        exc.hint = ("Use the deck's new name if it was renamed or moved. The "
                    "file named in the message is only an older copy.")
        return exc
    exc = DocumentNotFound(
        f"{first}. A save that did not finish left the deck as it was at "
        f"{found[0]}; its content is safe there. The next change to the deck "
        "through this server, or the next start of the server, puts it back; "
        f"you can also rename that file to {os.path.basename(target)}.")
    exc.hint = ("Nothing in the deck was lost. Make the change you wanted (it "
                "puts the deck back first), restart the server, or rename the "
                "file named in the message back to the deck's name.")
    return exc


def require_deck(path: str | os.PathLike, first: str) -> None:
    """The one check every route that reads a deck makes before it reads it:
    the deck is at its path, or comes back from another save's swap within
    the wait. Otherwise DocumentNotFound, with ``first`` as its first
    sentence and, when a save that did not finish left the deck aside, where
    it is. A folder is not a deck."""
    if exists_after_swap(path) and not os.path.isdir(path):
        return
    raise deck_not_found(path, first)


def through_swap(path: str | os.PathLike, read_once):
    """``read_once()``, and once more if the deck went missing under it
    because another save swapped it in at that very instant and it came
    back."""
    try:
        return read_once()
    except FileNotFoundError:
        if not exists_after_swap(path):
            raise
    return read_once()


def aside_in_place(path: str | os.PathLike) -> bool:
    """Whether a save has this deck's old file aside right now.

    ReplaceFileW moves the old file to a backup name and then moves the new
    one into its place, so the path itself is missing for an instant. The
    backup name is there for exactly that instant (and until it is removed),
    which is how a reader tells a deck that is being swapped from one that is
    not there."""
    target = os.path.abspath(os.fspath(path))
    folder, name = os.path.split(target)
    prefixes = {aside_prefix(name)}
    if _looks_short(name) or any(_looks_short(part) for part in Path(folder).parts):
        # A deck named by its 8.3 short name: its save named the old file
        # after the long name.
        spelled_folder, spelled_name = os.path.split(deck_spelling(target))
        prefixes.add(aside_prefix(spelled_name))
        folder = spelled_folder or folder
    try:
        return any(entry.startswith(tuple(prefixes)) for entry in os.listdir(folder or "."))
    except OSError:
        return False


def exists_after_swap(path: str | os.PathLike) -> bool:
    """Whether the deck is at its path, waiting out another save's swap.

    True at once for a deck that is there. For a missing deck, True once it
    is back while a save has its old file aside; False at once, with no
    wait, when no save has it aside, so a deck that is not there is reported
    as such straight away."""
    for pause in (0.0, *_SWAP_POLL):
        if pause:
            time.sleep(pause)
        if os.path.exists(path):
            return True
        if not aside_in_place(path):
            return os.path.exists(path)
    return os.path.exists(path)
