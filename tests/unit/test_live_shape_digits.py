"""The live route reads a shape given as text the way the file route does.

Text naming a shape is a name first: digits name the shape with that id
only when no shape has them as its name. The live lookup took any text that
str.isdigit() accepts as an id, so a shape named with digits could not be
reached by name there, and a superscript two (which isdigit accepts) raised
a ValueError from int() instead of a not-found answer.
"""

from __future__ import annotations

import pytest

from kitchensink4ppt.com import live_ops
from kitchensink4ppt.core.errors import TargetNotFound


class _Shape:
    Type = 1

    def __init__(self, sid, name):
        self.Id, self.Name = sid, name


class _Shapes:
    def __init__(self, items):
        self._items, self.Count = items, len(items)

    def Item(self, i):  # noqa: N802 - COM spelling
        return self._items[i - 1]


class _Slide:
    SlideIndex = 1

    def __init__(self, shapes):
        self.Shapes = _Shapes(shapes)


SLIDE = _Slide([_Shape(3, "12"), _Shape(12, "Title"), _Shape(5, "Note")])


def test_digits_that_name_a_shape_find_that_shape():
    assert live_ops._find_shape_live(SLIDE, "12").Id == 3


def test_digits_that_name_no_shape_are_an_id():
    assert live_ops._find_shape_live(SLIDE, "5").Id == 5
    assert live_ops._find_shape_live(SLIDE, 12).Id == 12


def test_a_superscript_digit_is_a_name_and_not_found():
    with pytest.raises(TargetNotFound):
        live_ops._find_shape_live(SLIDE, chr(0xB2))
