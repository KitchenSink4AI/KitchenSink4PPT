"""Checks on a ZIP package's directory before any part of it is inflated.

Reading a package inflates its parts in full, and a part can claim far more
than it stores: a megabyte of compressed zeros inflates to a gigabyte, and
reading it asks for about twice that in memory. The sizes the package
declares are known before anything is inflated, and Python's zipfile never
returns more than the declared size of a part, so they are a safe bound:
the parts that inflate more than MAX_RATIO times their stored size are
added up, and past BOMB_FLOOR the package is refused. A part stored as it
is (media, a large video) never counts, whatever its size.

A package that names one part twice (or names it twice with different
capital letters: package part names ignore case) is refused too: a reader
keeps one copy, PowerPoint may show the other.
"""

from __future__ import annotations

from typing import Iterable
from zipfile import ZipInfo

from .errors import PackageRefused

#: A part that inflates to more than this many times its stored size counts
#: toward the limit below. Presentation XML inflates about 5 to 20 times;
#: pictures, video and fonts about once.
MAX_RATIO = 100

#: The most that highly inflating parts may add up to, in bytes.
BOMB_FLOOR = 256 * 1024 * 1024

_ASCII_LOWER = str.maketrans("ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz")


def _size(n: int) -> str:
    if n >= 100 * 1024 * 1024:
        return f"{n / (1024 * 1024):,.0f} MB"
    if n >= 1024 * 1024:
        return f"{n / (1024 * 1024):,.1f} MB"
    return f"{max(n, 1) / 1024:,.0f} KB"


def _twice(doc: str, first: str, second: str) -> PackageRefused:
    if first == second:
        what = f"lists the part {first} twice"
    else:
        what = (f"holds {first} and {second}, which name the same part (package part names "
                "ignore capital letters)")
    return PackageRefused(
        f"{doc}: its ZIP package {what}. PowerPoint repairs a file like that when it opens "
        "it, and the two copies can differ, so there is no telling which one PowerPoint "
        "shows; the file was not read. Nothing was changed. Open the file in PowerPoint, let "
        "it repair the file and save it, then try again.")


def check_members(infos: Iterable[ZipInfo], *, doc: str) -> None:
    """Refuse (PackageRefused, UNSUPPORTED_CONTENT) a package whose
    directory names a part twice, or whose highly inflating parts would add
    up past BOMB_FLOOR. Reads only the directory entries."""
    seen: dict[str, str] = {}
    inflating = 0
    worst: ZipInfo | None = None
    for info in infos:
        name = info.filename
        if not name.endswith("/"):
            key = name.translate(_ASCII_LOWER)
            if key in seen:
                raise _twice(doc, seen[key], name)
            seen[key] = name
        if info.file_size > MAX_RATIO * max(info.compress_size, 1):
            inflating += info.file_size
            if worst is None or info.file_size > worst.file_size:
                worst = info
    if inflating > BOMB_FLOOR and worst is not None:
        also = ("" if inflating == worst.file_size else
                f", and its highly compressed parts would take {_size(inflating)} in all")
        raise PackageRefused(
            f"{doc}: its part {worst.filename} would expand from {_size(worst.compress_size)} "
            f"in the file to {_size(worst.file_size)} when read{also}. A part like that is "
            "built to exhaust memory and no presentation needs one, so the file was not "
            "read. Nothing was changed. If you expected this file to open, ask whoever sent "
            "it for a new copy.")


def too_large_to_read(info: ZipInfo) -> bool:
    """True when reading this one part in full would ask for more than the
    limit, whatever it stores (for the readers that take one part straight
    from a package without opening it)."""
    return info.file_size > BOMB_FLOOR
