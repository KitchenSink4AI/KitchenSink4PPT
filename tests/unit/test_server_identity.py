"""Guard: a connected agent can find out what it is connected to.

This server once shipped no way to answer "what am I running". The name field
is lowercase, clients display the user's own config alias, the injected
instructions carried no product name, and there was no get_server_info tool
at all, so an agent holding a KitchenSink4PPT session could not name the
product, the package, or the version it was talking to. Both halves of the
fix are checked here: the instructions lead line and the tool.

The figure tests are the reason this file matters. Every count
get_server_info reports has to be READ, not remembered: the totals off the
live pack registry, the test figure off the stamped measurement in
shipped.py. A hand-typed number in a self-description is how a server ends
up telling an agent it has 138 tools while shipping 142.
"""

from __future__ import annotations

import sys

from kitchensink4ppt import __version__, server, shipped
from kitchensink4ppt.registry import GROUP_OPS, OP_TO_GROUP


def _info() -> dict:
    return server.get_server_info()


def test_names_the_product_the_package_and_the_version():
    out = _info()
    assert out["product"] == "KitchenSink4PPT"
    assert out["package"] == "kitchensink4ppt"
    assert out["name"] == "kitchensink4ppt"
    assert out["version"] == __version__
    assert out["platform"].startswith(sys.platform + "-")


def test_points_at_the_landing_page_and_the_three_siblings():
    out = _info()
    assert out["homepage"] == "https://community.kitchensink4.ai/ppt/"
    assert out["family"] == [
        "kitchensink4word", "kitchensink4xl", "kitchensink4web",
    ]
    assert "kitchensink4ppt" not in out["family"]


def test_tool_totals_are_read_off_the_live_registry():
    """Not a remembered number: change the surface and this moves with it."""
    out = _info()
    assert out["tools_registered"] == len(GROUP_OPS) + 1
    assert out["operations_registered"] == len(OP_TO_GROUP)


def test_the_test_figure_comes_from_the_stamped_measurement():
    assert _info()["tests"] == shipped.TESTS


def test_reports_the_fixed_surface_and_group_counts():
    out = _info()
    assert out["surface"]["fixed"] is True
    assert out["surface"]["operations"] == 145
    assert out["surface"]["tools"] == {
        name: len(ops) for name, ops in GROUP_OPS.items()
    }


def test_needs_no_deck_and_no_arguments():
    """The orient call has to work on a fresh session with nothing open."""
    import inspect

    required = [
        parameter
        for parameter in inspect.signature(server.get_server_info).parameters.values()
        if parameter.default is inspect.Parameter.empty
        and parameter.kind not in {
            inspect.Parameter.VAR_POSITIONAL,
            inspect.Parameter.VAR_KEYWORD,
        }
    ]
    assert not required


def test_it_is_in_the_fixed_read_group_so_a_fresh_session_can_call_it():
    assert "get_server_info" in GROUP_OPS["ppt_read"]


def test_the_instructions_lead_with_the_brand():
    """The injected instructions are the only text an agent reads before it
    calls anything. If the product name is not in the first sentence, the
    tool above is the only path to it, and nothing tells the agent to look."""
    text = server.create_server().instructions
    assert text.startswith(
        "KitchenSink4PPT (kitchensink4ppt on PyPI), part of the "
        "KitchenSink4AI suite. ")
    assert "PowerPoint (.pptx) editor" in text
