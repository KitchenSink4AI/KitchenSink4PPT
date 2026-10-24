"""Frozen launch settings for the fixed KitchenSink4PPT surface."""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Mapping

from kitchensink4ppt.core.budget import DEFAULT_MAX_CHARS, MIN_MAX_CHARS


ENV_READ_ONLY = "KS4P_READ_ONLY"
ENV_FORBID = "KS4P_FORBID"
ENV_CALL_TIMEOUT = "KS4P_CALL_TIMEOUT_S"
ENV_MAX_OUTPUT = "KS4P_MAX_OUTPUT_CHARS"
#: Who comments and replies are signed by when a call names no author:
#: `user` (the user's own name, the default) or `assistant`
#: (core/authorship.py).
ENV_AUTHOR = "KS4P_AUTHOR"

LEGACY_ENV_NAMES = (
    "KS4P_MODE",
    "KS4P_PACK_POLICY",
    "KS4P_ALL_TOOLS",
    "KS4P_LOCK_TOOLS",
)

DEFAULT_CALL_TIMEOUT_S = 120.0
MIN_CALL_TIMEOUT_S = 1.0
MAX_CALL_TIMEOUT_S = 3600.0
DEFAULT_MAX_OUTPUT_CHARS = DEFAULT_MAX_CHARS
MIN_MAX_OUTPUT_CHARS = MIN_MAX_CHARS
MAX_MAX_OUTPUT_CHARS = 2_000_000

_TRUE = {"1", "true", "yes", "on"}
_FALSE = {"", "0", "false", "no", "off"}

# Launch settings from 1.x. Earlier versions let them limit or lock the tools;
# this version lists every tool from the start. A value that only restated the
# old default is ignored with a notice. A value that locked or narrowed the
# tools stops the start, so an upgrade never runs with less of a limit than the
# user set. Values are trimmed and compared without case.
_LEGACY_MODE_OK = {"lite", "full", "everything"}
_LEGACY_POLICY_OK = {"auto"}
_LEGACY_BOOL = {"1", "true", "on", "yes", "0", "false", "off", "no"}
_LEGACY_LOCK_OK = {"0", "false", "off", "no"}

LEGACY_REFUSAL = (
    "{name} is set to a value that limited the tools in an earlier version. "
    "This version does not read it, so the server stops here instead of "
    "starting without that limit. Remove {name} and use KS4P_READ_ONLY=1 "
    "(reading tools only) or KS4P_FORBID=<tool or op names> instead."
)

LEGACY_NOTICE = (
    "Ignored old settings: {names}. This version lists every tool from the "
    "start, and the server started normally. To limit the tools, use "
    "KS4P_READ_ONLY=1 or KS4P_FORBID=<tool or op names>. Remove these "
    "settings to stop this notice."
)


def legacy_notice(names: tuple[str, ...]) -> str:
    """The one startup line for old settings that were accepted and ignored."""
    return LEGACY_NOTICE.format(names=", ".join(names))


class SettingsError(ValueError):
    """A launch setting was not understood; startup must stop."""


@dataclass(frozen=True)
class Settings:
    read_only: bool = False
    forbid: tuple[str, ...] = ()
    call_timeout_s: float = DEFAULT_CALL_TIMEOUT_S
    max_output_chars: int = DEFAULT_MAX_OUTPUT_CHARS
    ignored_legacy: tuple[str, ...] = ()
    author_mode: str = "user"

    @classmethod
    def from_env(cls, env: Mapping[str, str] | None = None) -> "Settings":
        env = os.environ if env is None else env
        ignored_legacy = _classify_legacy(env)
        raw_read_only = (env.get(ENV_READ_ONLY) or "").strip().lower()
        if raw_read_only in _TRUE:
            read_only = True
        elif raw_read_only in _FALSE:
            read_only = False
        else:
            raise SettingsError(
                f"{ENV_READ_ONLY}={env.get(ENV_READ_ONLY)!r} is not understood; use 1 or 0"
            )

        forbid = tuple(
            sorted({item.strip() for item in (env.get(ENV_FORBID) or "").split(",") if item.strip()})
        )
        timeout = _number(
            env, ENV_CALL_TIMEOUT, DEFAULT_CALL_TIMEOUT_S,
            MIN_CALL_TIMEOUT_S, MAX_CALL_TIMEOUT_S, "seconds",
        )
        max_output = int(_number(
            env, ENV_MAX_OUTPUT, DEFAULT_MAX_OUTPUT_CHARS,
            MIN_MAX_OUTPUT_CHARS, MAX_MAX_OUTPUT_CHARS, "characters",
            whole=True,
        ))
        from kitchensink4ppt.core import authorship as _authorship

        try:
            author_mode = _authorship.mode_from(env.get(ENV_AUTHOR))
        except ValueError:
            raise SettingsError(
                f"{ENV_AUTHOR}={env.get(ENV_AUTHOR)!r} is not understood; use user (sign "
                "comments with the user's own name, the default) or assistant (sign them "
                f"as {_authorship.ASSISTANT_LABEL!r})"
            ) from None
        return cls(
            read_only=read_only,
            forbid=forbid,
            call_timeout_s=timeout,
            max_output_chars=max_output,
            ignored_legacy=ignored_legacy,
            author_mode=author_mode,
        )

    @property
    def has_legacy_notice(self) -> bool:
        return bool(self.ignored_legacy)


def _classify_legacy(env: Mapping[str, str]) -> tuple[str, ...]:
    """Sort the retired 1.x names into ignored (returned) and refused (raised).

    Every name is judged together, so a refused value stops the start even when
    another one, such as KS4P_MODE=lite, is accepted. A blank value is accepted
    and not listed, because there is nothing to tell the user. The raw value is
    never put into a message.
    """
    accepted_values = {
        "KS4P_MODE": _LEGACY_MODE_OK,
        "KS4P_PACK_POLICY": _LEGACY_POLICY_OK,
        "KS4P_ALL_TOOLS": _LEGACY_BOOL,
        "KS4P_LOCK_TOOLS": _LEGACY_LOCK_OK,
    }
    ignored: list[str] = []
    for name in LEGACY_ENV_NAMES:
        if name not in env:
            continue
        value = (env.get(name) or "").strip().lower()
        if not value:
            continue
        if value not in accepted_values[name]:
            raise SettingsError(LEGACY_REFUSAL.format(name=name))
        ignored.append(name)
    return tuple(ignored)


def _number(
    env: Mapping[str, str],
    name: str,
    default: float,
    lower: float,
    upper: float,
    unit: str,
    *,
    whole: bool = False,
) -> float:
    raw = (env.get(name) or "").strip()
    if not raw:
        return float(default)
    try:
        value = int(raw) if whole else float(raw)
    except ValueError:
        kind = "a whole number" if whole else "a number"
        raise SettingsError(f"{name}={raw!r} is not {kind} of {unit}") from None
    if not lower <= value <= upper:
        raise SettingsError(f"{name}={raw} is outside {lower:g} to {upper:g} {unit}")
    return float(value)


__all__ = [
    "DEFAULT_CALL_TIMEOUT_S", "DEFAULT_MAX_OUTPUT_CHARS", "ENV_AUTHOR", "ENV_CALL_TIMEOUT",
    "ENV_FORBID", "ENV_MAX_OUTPUT", "ENV_READ_ONLY", "LEGACY_ENV_NAMES",
    "LEGACY_NOTICE", "LEGACY_REFUSAL", "Settings", "SettingsError",
    "legacy_notice",
]
