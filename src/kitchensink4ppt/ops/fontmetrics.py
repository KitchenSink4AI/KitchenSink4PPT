"""Real advance-width measurement, behind the optional `metrics` extra.

Why this module exists: the fit model in ops/text.py was an average-glyph
width (0.5 x point size per character). On a 2026-09-21 field run it predicted
a two-line wrap for a Calibri 12 pt label that PowerPoint rendered as three
lines, because Calibri's per-character width swings between roughly 0.40 and
0.51 of the point size depending on the string. An estimator that passes a
real overflow is worse than no check, so when the font file can be found on
this machine the fit model measures the font's own advance widths instead.

Contract:
- fontTools is OPTIONAL and NEVER a runtime requirement. Every import of it
  is lazy and inside a try; without it `available()` is False and the caller
  keeps the old estimate, saying so.
- Advance widths only: the sum of each glyph's `hmtx` advance, scaled by
  `unitsPerEm`. No kerning, no ligatures, no justification, no shaping.
- The raw sum is NARROWER than what PowerPoint lays out, and narrow is the
  unsafe direction for an overflow check, so the raw sum is never what a fit
  decision sees. `text_width_pt` multiplies by WIDTH_SAFETY_FACTOR before
  anyone can act on it; `raw_text_width_pt` is the uncalibrated model, kept
  public so the calibration stays measurable. See WIDTH_SAFETY_FACTOR for
  the derivation.
- A face is a FILE PLUS AN INDEX. A .ttc/.otc collection holds several
  unrelated families (cambria.ttc is Cambria at 0 and Cambria Math at 1),
  the order is not a contract, and every face is indexed and cached by its
  own name table. PowerPoint's renderer remains the authority.
- Font lookup is by family plus bold/italic, through the OS font directories
  and, on Windows, the registry Fonts keys. Everything is cached per process.
- Nothing here raises for a missing font or an unreadable file: the caller
  gets None and falls back to the estimate.

fontTools over Pillow, recorded with the reasons: fontTools is pure Python
(one universal `py3-none-any` wheel, so 3.12 through 3.14 and every platform
are covered the day a new interpreter ships, where Pillow needs a compiled
wheel per interpreter), MIT licensed, and it reads the `hmtx` advance widths
this model wants directly. Pillow's `ImageFont.getlength` goes through
FreeType and a raster context to answer the same question, and Pillow is
already carried here for a different job (`optimize`), which would have made
one extra do two unrelated things.
"""

from __future__ import annotations

import contextlib
import os
import re
import sys
from pathlib import Path
from typing import NamedTuple


class FontFace(NamedTuple):
    """One measurable face: a font file plus the index inside it. The index
    is 0 for an ordinary .ttf/.otf and is load-bearing for a collection."""

    path: Path
    index: int = 0


#: Conservative calibration, applied ONCE PER LINE before any fit decision
#: is made. DERIVED FROM MEASUREMENT, not chosen.
#:
#: On 2026-09-22 the raw advance sum was compared against PowerPoint's own
#: TextRange.BoundWidth for twelve single-line cases, ten of them built to
#: land in the 0.96 to 1.15 band around the frame edge. The raw sum was
#: narrow in every case, but the shortfall was NOT proportional: it sat
#: between 1.23 and 3.35 POINTS regardless of string length (0.78 in to
#: 1.91 in of text), which is a constant trailing allowance in PowerPoint's
#: layout rather than a scale error in the advance widths. Expressed as a
#: ratio the same data reads 1.011x to 1.055x, and correcting it as a ratio
#: therefore over-corrects long strings: a 1.05 factor produced a false
#: positive on a label at 98.7% of its frame while an additive 3.5 pt did
#: not. A sweep of pads 0 to 5 pt against multipliers 1.00 to 1.02 gives
#: 3.5 pt with NO multiplier as the only setting with zero false negatives
#: AND zero false positives across all twelve.
#:
#: 3.5 pt covers the worst observed shortfall (3.35 pt) with margin.
#:
#: Why the calibration lives HERE and not in a downstream threshold: the
#: overflow boolean is decided by comparing a ratio to 1.0 deep inside the
#: fit model, so a threshold applied later cannot rescue a case already
#: called "fits" (review finding B1, 2026-09-22). Calibrating the
#: measurement makes every decision inherit the margin, including the WRAP
#: decision, where a narrow measurement silently drops a line PowerPoint
#: renders.
WIDTH_SAFETY_PAD_PT = 3.5

#: Trailing style words in a font's face name, longest first so "Bold Italic"
#: is consumed before "Bold".
_STYLE_WORDS = (
    ("bold italic", True, True),
    ("bolditalic", True, True),
    ("bold oblique", True, True),
    ("italic", False, True),
    ("oblique", False, True),
    ("bold", True, False),
    ("regular", False, False),
    ("normal", False, False),
)

_FONT_SUFFIXES = (".ttf", ".otf", ".ttc", ".otc")

#: (family_casefold, bold, italic) -> Path. Built once per process.
_INDEX: dict[tuple[str, bool, bool], Path] | None = None
#: Path -> measurement table, or None when the file could not be read.
_TABLES: dict[str, dict | None] = {}
#: (family, bold, italic) -> resolved Path or None, including misses.
_RESOLVED: dict[tuple[str, bool, bool], Path | None] = {}

_PAREN_TAIL = re.compile(r"\s*\((?:[^)]*)\)\s*$")


#: Memoized answer for available(); None until the first ask. Measurement
#: runs per word and, for an over-long word, per character, so the import
#: probe has to be a dict lookup rather than an import statement.
_AVAILABLE: bool | None = None


def available() -> bool:
    """Is the optional metrics extra installed? Lazy, never raises."""
    global _AVAILABLE
    if _AVAILABLE is None:
        try:
            import fontTools  # noqa: F401

            _AVAILABLE = True
        except Exception:
            _AVAILABLE = False
    return _AVAILABLE


def unavailable_reason() -> str:
    """Why measurement is off, in the words the results carry."""
    return "the optional metrics extra (fonttools) is not installed"


# ------------------------------------------------------------- font lookup


def _split_style(face: str) -> tuple[str, bool, bool]:
    """('Calibri Bold Italic') -> ('Calibri', True, True)."""
    name = _PAREN_TAIL.sub("", face or "").strip()
    low = name.casefold()
    for word, bold, italic in _STYLE_WORDS:
        if low.endswith(" " + word):
            return name[: -(len(word) + 1)].strip(), bold, italic
        if low == word:
            return "", bold, italic
    return name, False, False


def _font_dirs() -> list[Path]:
    dirs: list[Path] = []
    if sys.platform == "win32":
        win = os.environ.get("WINDIR") or r"C:\Windows"
        dirs.append(Path(win) / "Fonts")
        local = os.environ.get("LOCALAPPDATA")
        if local:
            dirs.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
    elif sys.platform == "darwin":
        dirs += [
            Path("/System/Library/Fonts"),
            Path("/Library/Fonts"),
            Path.home() / "Library" / "Fonts",
        ]
    else:
        dirs += [
            Path("/usr/share/fonts"),
            Path("/usr/local/share/fonts"),
            Path.home() / ".fonts",
            Path.home() / ".local" / "share" / "fonts",
        ]
    return [d for d in dirs if d.is_dir()]


def _registry_entries() -> list[tuple[str, str]]:
    """(face name, file) pairs from both Windows Fonts registry keys."""
    if sys.platform != "win32":
        return []
    try:
        import winreg
    except Exception:
        return []
    out: list[tuple[str, str]] = []
    key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, key_path) as key:
                count = winreg.QueryInfoKey(key)[1]
                for i in range(count):
                    try:
                        name, value, _kind = winreg.EnumValue(key, i)
                    except OSError:
                        continue
                    if isinstance(value, str) and value:
                        out.append((name, value))
        except OSError:
            continue
    return out


def _add(index: dict, family: str, bold: bool, italic: bool,
         face: FontFace) -> None:
    if not family:
        return
    key = (family.casefold(), bold, italic)
    index.setdefault(key, face)


def _is_collection(path: Path) -> bool:
    return path.suffix.lower() in (".ttc", ".otc")


def _face_identity(font) -> tuple[str, bool, bool] | None:
    """(family, bold, italic) read from a face's OWN name table.

    A collection's faces are not interchangeable and their order is not a
    contract: `cambria.ttc` carries Cambria at index 0 and Cambria Math at
    index 1, and picking index 0 for both is how Cambria Math silently gets
    measured as Cambria (review finding M2, 2026-09-22)."""
    try:
        names = font["name"]
        family = names.getBestFamilyName()
        subfamily = names.getBestSubFamilyName() or "Regular"
    except Exception:
        return None
    if not family:
        return None
    low = str(subfamily).casefold()
    bold = "bold" in low
    italic = "italic" in low or "oblique" in low
    try:  # fsSelection is the authoritative answer where the name is vague
        selection = font["OS/2"].fsSelection
        bold = bold or bool(selection & 0x20)
        italic = italic or bool(selection & 0x01)
    except Exception:
        pass
    return str(family), bold, italic


def _index_collection(index: dict, path: Path) -> bool:
    """Index every face in a .ttc/.otc by its own name table. Returns False
    when the collection could not be read, so the caller can fall back."""
    try:
        from fontTools.ttLib import TTCollection

        coll = TTCollection(str(path), lazy=True)
    except Exception:
        return False
    found = False
    try:
        for i, font in enumerate(coll.fonts):
            ident = _face_identity(font)
            if ident is None:
                continue
            family, bold, italic = ident
            _add(index, family, bold, italic, FontFace(path, i))
            found = True
    finally:
        with contextlib.suppress(Exception):
            coll.close()
    return found


def _index_registry(index: dict) -> None:
    win = Path(os.environ.get("WINDIR") or r"C:\Windows") / "Fonts"
    for face, value in _registry_entries():
        p = Path(value)
        if not p.is_absolute():
            p = win / value
        if not p.is_file():
            continue
        if _is_collection(p):
            # The registry records a collection as
            # "Cambria & Cambria Math (TrueType)", but the ampersand order is
            # not a promise about face order. Read the faces themselves.
            if _index_collection(index, p):
                continue
        base = _PAREN_TAIL.sub("", face).strip()
        for part in base.split("&"):
            family, bold, italic = _split_style(part.strip())
            _add(index, family, bold, italic, FontFace(p, 0))


def _index_directories(index: dict) -> None:
    for d in _font_dirs():
        try:
            entries = sorted(d.rglob("*"))
        except OSError:
            continue
        for p in entries:
            if p.suffix.lower() not in _FONT_SUFFIXES or not p.is_file():
                continue
            if _is_collection(p) and _index_collection(index, p):
                continue
            stem = p.stem
            # "DejaVuSans-BoldOblique" / "NotoSans_Bold" / "Calibri Bold"
            spaced = re.sub(r"[-_]+", " ", stem)
            spaced = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", spaced)
            family, bold, italic = _split_style(spaced)
            _add(index, family, bold, italic, FontFace(p, 0))


def _build_index() -> dict:
    index: dict[tuple[str, bool, bool], FontFace] = {}
    _index_registry(index)
    _index_directories(index)
    return index


def _index() -> dict:
    global _INDEX
    if _INDEX is None:
        try:
            _INDEX = _build_index()
        except Exception:
            _INDEX = {}
    return _INDEX


def find_font_face(family: str, *, bold: bool = False,
                   italic: bool = False) -> FontFace | None:
    """The FACE for a family and style, or None. A face is a file plus an
    index inside it, because a collection holds several unrelated families
    and the index is part of the identity. Bold/italic fall back to the
    nearest available face rather than failing outright."""
    if not family:
        return None
    key = (family.casefold(), bool(bold), bool(italic))
    if key in _RESOLVED:
        return _RESOLVED[key]
    idx = _index()
    fam = family.casefold()
    hit = None
    for want in (
        (fam, bool(bold), bool(italic)),
        (fam, bool(bold), False),
        (fam, False, bool(italic)),
        (fam, False, False),
    ):
        hit = idx.get(want)
        if hit is not None:
            break
    _RESOLVED[key] = hit
    return hit


def find_font_file(family: str, *, bold: bool = False,
                   italic: bool = False) -> Path | None:
    """The font FILE for a family and style, or None. Prefer
    find_font_face: inside a collection the file alone is ambiguous."""
    face = find_font_face(family, bold=bold, italic=italic)
    return face.path if face is not None else None


# ------------------------------------------------------------- measurement


def _load_table(face: FontFace) -> dict | None:
    """{'upem': int, 'widths': {codepoint: advance}, 'default': advance} or
    None when the face cannot be read.

    Keyed by (path, index): two faces of one collection are different fonts
    and caching them under the shared path served the second one the first
    one's widths."""
    key = (str(face.path), face.index)
    if key in _TABLES:  # a cached None is a face we already failed to read
        return _TABLES[key]
    path = face.path
    table = None
    try:
        from fontTools.ttLib import TTCollection, TTFont

        if _is_collection(path):
            coll = TTCollection(str(path), lazy=True)
            font = coll.fonts[face.index]
        else:
            font = TTFont(str(path), fontNumber=face.index, lazy=True)
        upem = int(font["head"].unitsPerEm) or 1000
        cmap = font.getBestCmap() or {}
        hmtx = font["hmtx"]
        widths: dict[int, int] = {}
        for cp, gname in cmap.items():
            try:
                widths[int(cp)] = int(hmtx[gname][0])
            except Exception:
                continue
        default = widths.get(ord(" "))
        if default is None:
            default = int(round(upem * 0.5))
        table = {"upem": upem, "widths": widths, "default": default}
        try:
            font.close()
        except Exception:
            pass
    except Exception:
        table = None
    _TABLES[key] = table
    return table


def raw_text_width_pt(text: str, family: str, size_pt: float, *,
                      bold: bool = False, italic: bool = False) -> float | None:
    """The UNCALIBRATED advance sum in points, or None when it cannot be
    measured. This is the raw model; callers making a fit decision want
    text_width_pt, which carries the safety factor. Kept public because
    the calibration itself has to be measurable against it."""
    if not available():
        return None
    face = find_font_face(family, bold=bold, italic=italic)
    if face is None:
        return None
    table = _load_table(face)
    if table is None:
        return None
    if not text:
        return 0.0
    widths = table["widths"]
    default = table["default"]
    total = 0
    for ch in text:
        total += widths.get(ord(ch), default)
    return total / table["upem"] * float(size_pt)


def text_width_pt(text: str, family: str, size_pt: float, *,
                  bold: bool = False, italic: bool = False) -> float | None:
    """Advance width of a STRETCH of text in points, uncalibrated.

    This is the per-token measure the wrapper works in, so it must NOT
    carry the calibration: the pad is a per-LINE allowance, and adding it to
    every token would over-count a line by the number of words on it. A fit
    decision uses `calibrated_line_pt` (or `wrap_styled`, which already
    accounts for it)."""
    return raw_text_width_pt(text, family, size_pt, bold=bold, italic=italic)


def calibrated_line_pt(raw_line_pt: float) -> float:
    """One laid-out line's width as PowerPoint will render it: the advance
    sum plus the measured trailing allowance. This is the number a fit
    decision is allowed to see."""
    return raw_line_pt + WIDTH_SAFETY_PAD_PT


def metrics_readable(family: str, *, bold: bool = False,
                     italic: bool = False) -> bool | None:
    """True when the face resolves AND its metrics parse, False when a file
    was found but could not be read, None when nothing resolved. The three
    outcomes are three different honest reasons to report."""
    if not available():
        return None
    face = find_font_face(family, bold=bold, italic=italic)
    if face is None:
        return None
    return _load_table(face) is not None


def font_file_name(family: str, *, bold: bool = False,
                   italic: bool = False) -> str | None:
    """The face's display name: the file, plus the index inside it when the
    file is a collection, because "cambria.ttc" alone names two fonts."""
    face = find_font_face(family, bold=bold, italic=italic)
    if face is None:
        return None
    if _is_collection(face.path):
        return f"{face.path.name}#{face.index}"
    return face.path.name


class StyledRun(NamedTuple):
    """A stretch of text in ONE resolved format. A paragraph is a sequence
    of these; measuring the whole paragraph in the first run's format is
    what review finding M1 was about.

    `spc_pt` is PowerPoint's character spacing (a:rPr/@spc, hundredths of a
    point) in points, added once per character. Ignoring it measured an
    expanded label at its unexpanded width and called the overflow a fit
    (second review, G1, 2026-09-22)."""

    text: str
    family: str
    size_pt: float
    bold: bool = False
    italic: bool = False
    spc_pt: float = 0.0


class Line(NamedTuple):
    """One laid-out line: how wide it came out, and the largest point size
    on it (which is what sets its height)."""

    width_pt: float
    max_size_pt: float


def _chunk_width(chunk: list[tuple[str, StyledRun]]) -> float | None:
    """Width of a run of (char, style) pairs, measured a STYLE AT A TIME so
    a word split across a format change is still measured correctly."""
    total = 0.0
    i = 0
    while i < len(chunk):
        style = chunk[i][1]
        j = i
        text = []
        while j < len(chunk) and chunk[j][1] is style:
            text.append(chunk[j][0])
            j += 1
        w = text_width_pt(
            "".join(text), style.family, style.size_pt,
            bold=style.bold, italic=style.italic,
        )
        if w is None:
            return None
        # Character spacing is an advance PowerPoint adds per character, so
        # it belongs here with the advances and nowhere else: every width
        # the wrapper works in comes through this function.
        total += w + style.spc_pt * len(text)
        i = j
    return total


def _max_size(chunk: list[tuple[str, StyledRun]], floor: float) -> float:
    return max([floor] + [s.size_pt for _c, s in chunk])


def wrap_styled(
    runs: list[StyledRun], first_width_pt: float,
    rest_width_pt: float | None = None,
) -> list[Line] | None:
    """Lay out a paragraph of mixed-format runs and report each line.

    Greedy word wrap on measured widths, with a word wider than the line
    broken inside itself the way PowerPoint breaks it. A "\\n" in a run's
    text is a hard break. Returns None when anything could not be measured,
    which is the caller's signal to fall back honestly.

    The calibration is applied HERE, by taking WIDTH_SAFETY_PAD_PT off the
    usable width once per line, which is where a constant trailing allowance
    belongs. That makes the wrap itself conservative: a line that PowerPoint
    would break one word earlier breaks one word earlier here too, instead
    of silently costing the caller a line."""
    if first_width_pt <= 0:
        return None
    if rest_width_pt is None or rest_width_pt <= 0:
        rest_width_pt = first_width_pt
    first_width_pt = max(1.0, first_width_pt - WIDTH_SAFETY_PAD_PT)
    rest_width_pt = max(1.0, rest_width_pt - WIDTH_SAFETY_PAD_PT)

    items: list[tuple[str, StyledRun]] = []
    for run in runs:
        for ch in run.text:
            items.append((ch, run))
    if not items:
        return None

    # Hard breaks first: each becomes its own laid-out block.
    blocks: list[list[tuple[str, StyledRun]]] = [[]]
    for ch, style in items:
        if ch == "\n":
            blocks.append([])
        else:
            blocks[-1].append((ch, style))

    default_size = max(r.size_pt for r in runs)
    out: list[Line] = []
    for block in blocks:
        tokens: list[list[tuple[str, StyledRun]]] = []
        current: list[tuple[str, StyledRun]] = []
        for ch, style in block:
            if ch.isspace():
                if current:
                    tokens.append(current)
                    current = []
                tokens.append([(ch, style)])
            else:
                current.append((ch, style))
        if current:
            tokens.append(current)
        if not tokens:
            out.append(Line(0.0, default_size))
            continue
        laid = _wrap_tokens(
            tokens, first_width_pt, rest_width_pt, default_size
        )
        if laid is None:
            return None
        out.extend(laid)
    return out


def _wrap_tokens(tokens, first_width_pt, rest_width_pt, default_size):
    lines: list[Line] = []
    limit = first_width_pt
    used = 0.0
    biggest = 0.0
    pending_space = 0.0

    def close_line():
        nonlocal used, biggest, limit, pending_space
        lines.append(Line(used, biggest or default_size))
        used = 0.0
        biggest = 0.0
        pending_space = 0.0
        limit = rest_width_pt

    for token in tokens:
        is_space = token[0][0].isspace()
        width = _chunk_width(token)
        if width is None:
            return None
        if is_space:
            if used > 0:
                pending_space += width
                biggest = _max_size(token, biggest)
            continue
        if used > 0 and used + pending_space + width > limit:
            close_line()
        if used == 0 and width > limit:
            # A single word wider than the line: PowerPoint breaks inside it
            # rather than letting it run out of the frame.
            for ch, style in token:
                cw = _chunk_width([(ch, style)])
                if cw is None:
                    return None
                if used > 0 and used + cw > limit:
                    close_line()
                used += cw
                biggest = max(biggest, style.size_pt)
            continue
        used += pending_space + width
        pending_space = 0.0
        biggest = _max_size(token, biggest)
    lines.append(Line(used, biggest or default_size))
    return lines


def _word_pieces(text: str) -> list[str]:
    """Split into wrap units, keeping the trailing space with each word the
    way a line breaker does (trailing spaces do not force a break)."""
    return re.findall(r"\S+\s*|\s+", text)


def wrapped_line_count(
    text: str,
    family: str,
    size_pt: float,
    first_width_pt: float,
    rest_width_pt: float | None = None,
    *,
    bold: bool = False,
    italic: bool = False,
) -> int | None:
    """Lines this string occupies when wrapped at measured word widths, or
    None when it cannot be measured. An empty string is one line."""
    if first_width_pt <= 0:
        return None
    if rest_width_pt is None or rest_width_pt <= 0:
        rest_width_pt = first_width_pt
    probe = text_width_pt("n", family, size_pt, bold=bold, italic=italic)
    if probe is None:
        return None
    if not text.strip():
        return 1

    def w(s: str) -> float:
        return text_width_pt(s, family, size_pt, bold=bold, italic=italic) or 0.0

    lines = 1
    limit = first_width_pt
    used = 0.0
    for piece in _word_pieces(text):
        bare = piece.rstrip()
        if not bare:
            continue  # run of spaces: never forces a break on its own
        width = w(bare)
        if used == 0.0:
            used = width
            if used > limit:
                # A single word wider than the line: it breaks inside itself.
                extra, used = _break_long_word(
                    bare, family, size_pt, limit, rest_width_pt,
                    bold=bold, italic=italic,
                )
                lines += extra
                limit = rest_width_pt
            continue
        space = w(" ")
        if used + space + width <= limit:
            used += space + width
            continue
        lines += 1
        limit = rest_width_pt
        used = width
        if used > limit:
            extra, used = _break_long_word(
                bare, family, size_pt, limit, rest_width_pt,
                bold=bold, italic=italic,
            )
            lines += extra
    return lines


def _break_long_word(word: str, family: str, size_pt: float,
                     limit: float, rest: float, *, bold: bool,
                     italic: bool) -> tuple[int, float]:
    """(extra lines, width used on the last line) for a word that does not
    fit one line. PowerPoint breaks inside such a word rather than letting
    it run out of the frame."""
    extra = 0
    used = 0.0
    cap = limit
    for ch in word:
        cw = text_width_pt(ch, family, size_pt, bold=bold, italic=italic) or 0.0
        if used > 0 and used + cw > cap:
            extra += 1
            cap = rest
            used = cw
        else:
            used += cw
    return extra, used
