"""Served text about reading equations agrees with the reads: get_text and
find_text show an equation as its characters, so no recipe or help line may
say that text reads leave math out."""

import json
import re

from kitchensink4ppt.ops import workflows

_STALE = re.compile(r"never appear in\s+get_text|only read path for math", re.IGNORECASE)


def test_equation_recipe_does_not_say_text_reads_leave_math_out():
    recipe = workflows.WORKFLOWS["equation-authoring"]
    assert not _STALE.search(json.dumps(recipe))
    step = next(s for s in recipe["steps"] if s["op"] == "list_equations")
    assert "get_text" in step["why"] and "characters" in step["why"]


def test_equation_recipe_does_not_say_a_box_powerpoint_saved_is_left_out():
    """When PowerPoint saves a text box that holds an equation it keeps the
    whole box in a compatibility wrapper at the slide level. The reads now
    read that box, its equation included (test_wrapped_shape_reads.py), so
    the recipe must not say such a box is not read. list_equations also
    lists the equations in table cells and in the speaker notes
    (test_list_equations_tables_notes.py), and the recipe says where it
    looks."""
    recipe = workflows.WORKFLOWS["equation-authoring"]
    step = next(s for s in recipe["steps"] if s["op"] == "list_equations")
    assert "not read" not in step["why"] and "leave" not in step["why"]
    assert "table cell" in step["why"] and "notes" in step["why"]


def test_no_workflow_says_text_reads_leave_math_out():
    assert not _STALE.search(json.dumps(workflows.get_workflows()))
