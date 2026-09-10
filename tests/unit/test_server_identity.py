"""Guard: a connected agent can find out what it is connected to.

Until 2026-09-10 this server shipped no way to answer "what am I running".
The name field is lowercase, clients display the user's own config alias,
the injected instructions carried no product name, and there was no
get_server_info tool at all, so an agent holding a KitchenSink4PPT session
could not name the product, the package, or the version it was talking to.
Both halves of the fix are checked here: the instructions lead line and the
tool.

The figure tests are the reason this file matters after the copy wave. Every
count get_server_info reports has to be READ, not remembered: the totals off
the live pack registry, the test figure off the stamped measurement in
shipped.py. A hand-typed number in a self-description is how a server ends
up telling an agent it has 138 tools while shipping 142.
"""

from __future__ import annotations

from kitchensink4ppt import __version__, packs, server, shipped


def _info() -> dict:
    return server.get_server_info()


def test_names_the_product_the_package_and_the_version():
    out = _info()
    assert out["product"] == "KitchenSink4PPT"
    assert out["package"] == "kitchensink4ppt"
    assert out["name"] == "kitchensink4ppt"
    assert out["version"] == __version__


def test_points_at_the_landing_page_and_the_three_siblings():
    out = _info()
    assert out["homepage"] == "https://kitchensink4.ai/KitchenSink4PPT/"
    assert out["family"] == [
        "kitchensink4word", "kitchensink4xl", "kitchensink4web",
    ]
    assert "kitchensink4ppt" not in out["family"]


def test_tool_totals_are_read_off_the_live_registry():
    """Not a remembered number: change the surface and this moves with it."""
    names = packs.tool_names()
    out = _info()
    assert out["tools_registered"] == sum(len(v) for v in names.values())
    assert out["tools_lite"] == len(names["lite"])


def test_the_test_figure_comes_from_the_stamped_measurement():
    assert _info()["tests"] == shipped.TESTS


def test_reports_the_surface_and_the_packs_that_can_be_switched_on():
    out = _info()
    assert out["surface"] == packs.surface_report()
    assert out["packs_available"] == packs.pack_names()


def test_needs_no_deck_and_no_arguments():
    """The orient call has to work on a fresh session with nothing open."""
    import inspect

    assert not inspect.signature(server.get_server_info).parameters


def test_it_is_in_the_lite_core_so_a_fresh_session_can_call_it():
    assert "get_server_info" in packs.tool_names()["lite"]


def test_the_instructions_lead_with_the_brand():
    """The injected instructions are the only text an agent reads before it
    calls anything. If the product name is not in the first sentence, the
    tool above is the only path to it, and nothing tells the agent to look."""
    text = server.mcp.instructions
    assert text.startswith(
        "KitchenSink4PPT (kitchensink4ppt on PyPI), part of the "
        "KitchenSink4AI suite. ")
    assert "PowerPoint (.pptx) editor" in text
