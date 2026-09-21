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
  `unitsPerEm`. No kerning, no ligatures, no justification, no shaping. That
  is deliberate, and the residual was measured rather than assumed: against
  PowerPoint's own TextRange.BoundWidth on ten cases, these widths run
  between 1.9% and 6.1% NARROW, mean 3.5%. Narrow is the UNSAFE direction
  for an overflow check, which is why the caller's suppression threshold for
  this path sits at 1.05 rather than 1.0. Line counts were exact on all
  twelve cases measured. PowerPoint's renderer remains the authority.
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

import os
import re
import sys
from pathlib import Path

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


def _add(index: dict, family: str, bold: bool, italic: bool, path: Path) -> None:
    if not family:
        return
    key = (family.casefold(), bold, italic)
    index.setdefault(key, path)


def _index_registry(index: dict) -> None:
    win = Path(os.environ.get("WINDIR") or r"C:\Windows") / "Fonts"
    for face, value in _registry_entries():
        p = Path(value)
        if not p.is_absolute():
            p = win / value
        if not p.exists():
            continue
        # "Cambria & Cambria Math (TrueType)" registers two families on one
        # file; the ampersand form is how Windows records a collection.
        base = _PAREN_TAIL.sub("", face).strip()
        for part in base.split("&"):
            family, bold, italic = _split_style(part.strip())
            _add(index, family, bold, italic, p)


def _index_directories(index: dict) -> None:
    for d in _font_dirs():
        try:
            entries = sorted(d.rglob("*"))
        except OSError:
            continue
        for p in entries:
            if p.suffix.lower() not in _FONT_SUFFIXES or not p.is_file():
                continue
            stem = p.stem
            # "DejaVuSans-BoldOblique" / "NotoSans_Bold" / "Calibri Bold"
            spaced = re.sub(r"[-_]+", " ", stem)
            spaced = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", spaced)
            family, bold, italic = _split_style(spaced)
            _add(index, family, bold, italic, p)


def _build_index() -> dict:
    index: dict[tuple[str, bool, bool], Path] = {}
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


def find_font_file(family: str, *, bold: bool = False,
                   italic: bool = False) -> Path | None:
    """The font file for a family and style, or None. Bold/italic fall back
    to the nearest available face rather than failing outright."""
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


# ------------------------------------------------------------- measurement


def _load_table(path: Path) -> dict | None:
    """{'upem': int, 'widths': {codepoint: advance}, 'default': advance} or
    None when the file cannot be read as a font."""
    key = str(path)
    if key in _TABLES:  # a cached None is a file we already failed to read
        return _TABLES[key]
    table = None
    try:
        from fontTools.ttLib import TTCollection, TTFont

        if path.suffix.lower() in (".ttc", ".otc"):
            coll = TTCollection(str(path), lazy=True)
            font = coll.fonts[0]
        else:
            font = TTFont(str(path), fontNumber=0, lazy=True)
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


def text_width_pt(text: str, family: str, size_pt: float, *,
                  bold: bool = False, italic: bool = False) -> float | None:
    """Advance width of `text` in points, or None when it cannot be
    measured. Sum of per-glyph advances; no kerning."""
    if not available():
        return None
    path = find_font_file(family, bold=bold, italic=italic)
    if path is None:
        return None
    table = _load_table(path)
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


def font_file_name(family: str, *, bold: bool = False,
                   italic: bool = False) -> str | None:
    path = find_font_file(family, bold=bold, italic=italic)
    return path.name if path is not None else None


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
