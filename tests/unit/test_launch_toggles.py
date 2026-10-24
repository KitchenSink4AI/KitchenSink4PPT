"""Launch settings for the immutable nine-tool PowerPoint surface.
These are the replacements for the retired startup mode, tool-pack toggle,
and runtime enable/disable tests. Settings are parsed once and frozen into a
SurfacePlan; no call can change the advertised tool list.
"""

from __future__ import annotations

import pytest

from kitchensink4ppt.settings import (
    ENV_CALL_TIMEOUT,
    ENV_FORBID,
    ENV_MAX_OUTPUT,
    ENV_READ_ONLY,
    LEGACY_ENV_NAMES,
    LEGACY_REFUSAL,
    Settings,
    SettingsError,
    legacy_notice,
)
from kitchensink4ppt.surface.plan import ALL_TOOLS, build


@pytest.mark.parametrize("value", ["1", "true", "YES", "on"])
def test_read_only_truthy_values(value):
    settings = Settings.from_env({ENV_READ_ONLY: value})
    assert settings.read_only is True


@pytest.mark.parametrize("value", ["", "0", "false", "NO", "off"])
def test_read_only_false_values(value):
    settings = Settings.from_env({ENV_READ_ONLY: value})
    assert settings.read_only is False


def test_read_only_rejects_ambiguous_value():
    with pytest.raises(SettingsError, match=ENV_READ_ONLY):
        Settings.from_env({ENV_READ_ONLY: "sometimes"})


def test_forbid_is_trimmed_deduplicated_and_sorted():
    settings = Settings.from_env({
        ENV_FORBID: " get_text, ppt_graphics,get_text, ppt_help "
    })
    assert settings.forbid == ("get_text", "ppt_graphics", "ppt_help")


def test_unknown_forbid_fails_startup():
    with pytest.raises(SettingsError, match="not a tool or op"):
        build(Settings.from_env({ENV_FORBID: "not-a-real-op"}))


def test_read_only_surface_contains_only_read_and_inspect_plus_help():
    plan = build(Settings.from_env({ENV_READ_ONLY: "1"}))
    assert tuple(plan.tools) == ("ppt_read", "ppt_inspect", "ppt_help")
    assert all(tool.read_only for tool in plan.tools.values())


def test_forbid_can_remove_a_tool_and_one_operation():
    plan = build(Settings.from_env({
        ENV_FORBID: "ppt_graphics,get_text,ppt_help"
    }))
    assert "ppt_graphics" not in plan.tools
    assert "ppt_help" not in plan.tools
    assert "get_text" not in plan.tools["ppt_read"].ops
    assert plan.forbidden_ops == frozenset({"get_text"})
    assert plan.forbidden_tools == frozenset({"ppt_graphics", "ppt_help"})


LEGACY_MODE = "KS4P_MODE"
LEGACY_POLICY = "KS4P_PACK_POLICY"
LEGACY_ALL = "KS4P_ALL_TOOLS"
LEGACY_LOCK = "KS4P_LOCK_TOOLS"

ACCEPTED_LEGACY = (
    [(LEGACY_MODE, v) for v in ("lite", "full", "everything", "LITE", " Full ", "\tEverything\n")]
    + [(LEGACY_POLICY, v) for v in ("auto", "AUTO", "  Auto ")]
    + [(LEGACY_ALL, v) for v in (
        "1", "true", "on", "yes", "0", "false", "off", "no", "TRUE", " False ", "No",
    )]
    + [(LEGACY_LOCK, v) for v in ("0", "false", "off", "no", "FALSE", " Off ")]
)

REFUSED_LEGACY = [
    (LEGACY_MODE, "lite,graphics"),
    (LEGACY_MODE, "graphics"),
    (LEGACY_MODE, "full,typo-pack"),
    (LEGACY_MODE, "zzmarker"),
    (LEGACY_POLICY, "locked"),
    (LEGACY_POLICY, "LOCKED"),
    (LEGACY_POLICY, "zzmarker"),
    (LEGACY_ALL, "zzmarker"),
    (LEGACY_ALL, "2"),
    (LEGACY_LOCK, "1"),
    (LEGACY_LOCK, "true"),
    (LEGACY_LOCK, " ON "),
    (LEGACY_LOCK, "yes"),
    (LEGACY_LOCK, "zzmarker"),
]


def test_no_legacy_settings_means_nothing_ignored():
    assert Settings.from_env({}).ignored_legacy == ()
    assert Settings().has_legacy_notice is False


@pytest.mark.parametrize("name,value", ACCEPTED_LEGACY)
def test_old_defaults_are_accepted_and_recorded_as_ignored(name, value):
    settings = Settings.from_env({name: value})
    assert settings.ignored_legacy == (name,)
    assert settings.has_legacy_notice is True
    assert settings.read_only is False and settings.forbid == ()


def test_the_old_bundle_defaults_together_are_accepted():
    settings = Settings.from_env({
        LEGACY_MODE: "lite",
        LEGACY_POLICY: "auto",
        LEGACY_ALL: "false",
        LEGACY_LOCK: "false",
    })
    assert settings.ignored_legacy == LEGACY_ENV_NAMES


@pytest.mark.parametrize("name", LEGACY_ENV_NAMES)
@pytest.mark.parametrize("value", ["", "   ", "\t"])
def test_blank_old_settings_are_accepted_and_not_listed(name, value):
    settings = Settings.from_env({name: value})
    assert settings.ignored_legacy == ()
    assert settings.has_legacy_notice is False


@pytest.mark.parametrize("name,value", REFUSED_LEGACY)
def test_old_settings_that_limited_the_tools_stop_the_start(name, value):
    with pytest.raises(SettingsError) as caught:
        Settings.from_env({name: value})
    assert str(caught.value) == LEGACY_REFUSAL.format(name=name)
    assert "zzmarker" not in str(caught.value)
    assert ENV_READ_ONLY in str(caught.value) and ENV_FORBID in str(caught.value)


@pytest.mark.parametrize("name,value", REFUSED_LEGACY)
def test_a_refused_old_setting_still_stops_the_start_beside_mode_lite(name, value):
    env = {LEGACY_MODE: "lite", name: value}
    if name == LEGACY_MODE:
        env = {LEGACY_POLICY: "auto", LEGACY_ALL: "false", name: value}
    with pytest.raises(SettingsError) as caught:
        Settings.from_env(env)
    assert str(caught.value) == LEGACY_REFUSAL.format(name=name)


def test_mode_lite_does_not_loosen_the_present_day_restrictions():
    old = {LEGACY_MODE: "lite", LEGACY_POLICY: "auto", LEGACY_ALL: "false", LEGACY_LOCK: "false"}
    read_only = Settings.from_env({**old, ENV_READ_ONLY: "1"})
    plan = build(read_only)
    assert tuple(plan.tools) == ("ppt_read", "ppt_inspect", "ppt_help")
    forbidding = Settings.from_env({**old, ENV_FORBID: "get_text,ppt_graphics"})
    plan = build(forbidding)
    assert "ppt_graphics" not in plan.tools
    assert "get_text" not in plan.tools["ppt_read"].ops
    assert plan.forbidden_ops == frozenset({"get_text"})
    with pytest.raises(SettingsError, match=ENV_READ_ONLY):
        Settings.from_env({**old, ENV_READ_ONLY: "sometimes"})
    with pytest.raises(SettingsError, match=ENV_CALL_TIMEOUT):
        Settings.from_env({**old, ENV_CALL_TIMEOUT: "0"})
    with pytest.raises(SettingsError, match="not a tool or op"):
        build(Settings.from_env({**old, ENV_FORBID: "not-a-real-op"}))


def test_the_startup_notice_is_one_line_naming_only_the_ignored_settings():
    line = legacy_notice(("KS4P_MODE", "KS4P_ALL_TOOLS"))
    assert line == (
        "Ignored old settings: KS4P_MODE, KS4P_ALL_TOOLS. This version lists every "
        "tool from the start, and the server started normally. To limit the tools, "
        "use KS4P_READ_ONLY=1 or KS4P_FORBID=<tool or op names>. Remove these "
        "settings to stop this notice."
    )
    assert chr(10) not in line and chr(0x2014) not in line


@pytest.mark.parametrize(
    "name,value",
    [
        (ENV_CALL_TIMEOUT, "not-a-number"),
        (ENV_CALL_TIMEOUT, "0"),
        (ENV_CALL_TIMEOUT, "3601"),
        (ENV_MAX_OUTPUT, "1.5"),
        (ENV_MAX_OUTPUT, "1999"),
        (ENV_MAX_OUTPUT, "2000001"),
    ],
)
def test_numeric_setting_errors_fail_loudly(name, value):
    with pytest.raises(SettingsError, match=name):
        Settings.from_env({name: value})


def test_numeric_settings_are_frozen_into_plan():
    settings = Settings.from_env({
        ENV_CALL_TIMEOUT: "45.5",
        ENV_MAX_OUTPUT: "64000",
    })
    plan = build(settings)
    assert plan.settings.call_timeout_s == 45.5
    assert plan.settings.max_output_chars == 64000


def test_an_unknown_variable_name_is_ignored_and_the_readme_says_so():
    """The README said any setting the server does not understand stops
    startup. Only values it cannot use and retired names do; a misspelled
    name such as KS4P_READ_ONY is ignored (environment variables are shared
    with everything else on the machine), so the README now says that."""
    from pathlib import Path

    from kitchensink4ppt import server  # noqa: F401 - populates the registry

    settings = Settings.from_env({"KS4P_READ_ONY": "1"})
    assert settings.read_only is False
    assert "ppt_edit" in build(settings).tools
    readme = (Path(__file__).resolve().parents[2] / "README.md").read_text(encoding="utf-8")
    assert "A setting the server does not understand stops startup" not in readme
    assert "A variable name the server does not know is ignored" in readme
