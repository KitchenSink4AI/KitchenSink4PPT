"""The launcher's pin on the core wheel, as scripts/release_verify.py checks it.

The launcher declares the core once, unconditionally, at the release version.
It may also forward extras (metrics, optimize) to the same core; each of those
must carry the same exact pin. Anything else is refused by name.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]


def _load():
    path = ROOT / "scripts" / "release_verify.py"
    spec = importlib.util.spec_from_file_location("_release_verify_under_test", path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


rv = _load()
DIST = "kitchensink4ppt-core"
BASE = "kitchensink4ppt-core==2.0.0"
METRICS = 'kitchensink4ppt-core[metrics]==2.0.0; extra == "metrics"'
OPTIMIZE = 'kitchensink4ppt-core[optimize]==2.0.0; extra == "optimize"'
OTHER = "mcp>=1.2"


def test_the_base_pin_with_two_pinned_forwarding_extras_passes():
    rv.exact_pin([OTHER, BASE, METRICS, OPTIMIZE], DIST, "2.0.0")


def test_the_base_pin_alone_passes():
    rv.exact_pin([OTHER, BASE], DIST, "2.0.0")


def test_a_forwarded_extra_at_the_wrong_version_fails():
    wrong = 'kitchensink4ppt-core[metrics]==1.9.0; extra == "metrics"'
    with pytest.raises(rv.ReleaseCheckError, match="metrics"):
        rv.exact_pin([BASE, wrong, OPTIMIZE], DIST, "2.0.0")


def test_a_forwarded_extra_without_a_pin_fails():
    loose = 'kitchensink4ppt-core[metrics]; extra == "metrics"'
    with pytest.raises(rv.ReleaseCheckError, match="metrics"):
        rv.exact_pin([BASE, loose], DIST, "2.0.0")
    ranged = 'kitchensink4ppt-core[optimize]>=2.0.0; extra == "optimize"'
    with pytest.raises(rv.ReleaseCheckError, match="optimize"):
        rv.exact_pin([BASE, ranged], DIST, "2.0.0")


def test_a_second_unconditional_core_requirement_still_fails():
    with pytest.raises(rv.ReleaseCheckError, match="exactly once"):
        rv.exact_pin([BASE, "kitchensink4ppt-core[metrics]==2.0.0"], DIST, "2.0.0")
    with pytest.raises(rv.ReleaseCheckError, match="exactly once"):
        rv.exact_pin([BASE, BASE, METRICS], DIST, "2.0.0")


def test_no_core_requirement_still_fails():
    with pytest.raises(rv.ReleaseCheckError, match="exactly once"):
        rv.exact_pin([OTHER], DIST, "2.0.0")
    with pytest.raises(rv.ReleaseCheckError, match="exactly once"):
        rv.exact_pin([OTHER, METRICS], DIST, "2.0.0")
