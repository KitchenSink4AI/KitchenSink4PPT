"""The three deterministic build gates.

1. CONTENT. No Community guidance string reaches a Business artifact, and the
   notices that MUST ship are still there. Runtime suppression is not the
   requirement; physical absence is, so the gate fails in both directions.
2. LICENSE METADATA. No distribution ships a wheel granted under the other
   edition's license, which is what keeps two conflicting grants out of one
   customer's hands.
3. ADAPTER PROVENANCE. The license policy adapters being built are the pinned
   canonical shared source and not a fork. A fork is invisible in a built
   wheel (same name, same version, different payload), so the digest is the
   only mechanical evidence and the build refuses without it.

Each gate has a poisoned proof: an injected leak, a suppressed relabelling,
and a drifted adapter, each run against the real build rather than against
fixtures shaped like it.
"""
import importlib.util
import json
import subprocess
import sys
import tomllib
import zipfile
from datetime import date
from pathlib import Path

import pytest

import adapter_source

ROOT = Path(__file__).resolve().parents[2]

spec = importlib.util.spec_from_file_location(
    "split_build", ROOT / "scripts/build_editions.py")
build_script = importlib.util.module_from_spec(spec)
spec.loader.exec_module(build_script)


def _wheel(path: Path, members: dict) -> Path:
    with zipfile.ZipFile(path, "w") as archive:
        for name, text in members.items():
            archive.writestr(name, text)
    return path


NOTICES = "\n".join(build_script.REQUIRED_NOTICES)


# ------------------------------------------------------ gate 1: the content


def test_the_gate_knows_what_it_is_looking_for():
    needles = build_script.community_needles()
    assert len(needles) >= 5
    from ks4_license_community import notices
    assert any(n in notices.WARNING_13 for n in needles)
    # Short fragments are excluded on purpose; they collide with ordinary prose.
    assert all(len(n) >= 40 for n in needles)


def test_a_gate_with_no_needles_refuses_to_pass_anything(tmp_path):
    clean = _wheel(tmp_path / "clean.whl", {"a/LICENSE": NOTICES})
    with pytest.raises(build_script.BuildScanError, match="would pass anything"):
        build_script.scan_business_artifacts([clean], needles=())


def test_a_clean_artifact_set_passes(tmp_path):
    clean = _wheel(tmp_path / "clean.whl", {
        "ks4_license_business/policy.py": "CONFIGURED = True\n",
        "dist-info/licenses/LICENSE": NOTICES,
    })
    report = build_script.scan_business_artifacts([clean])
    assert report["members_scanned"] == 2
    assert report["needles"] >= 5


@pytest.mark.parametrize("member", [
    "kitchensink4ppt/server.py", "dist-info/METADATA", "INSTALL.md",
])
def test_injecting_the_guidance_into_any_member_fails_the_build(tmp_path, member):
    from ks4_license_community import notices
    poisoned = _wheel(tmp_path / "poisoned.whl", {
        member: "prose prose\n" + notices.WARNING_13 + "\nmore prose",
        "dist-info/licenses/LICENSE": NOTICES,
    })
    with pytest.raises(build_script.BuildScanError, match="Community guidance"):
        build_script.scan_business_artifacts([poisoned])


def test_the_gate_reads_inside_a_nested_bundle(tmp_path):
    from ks4_license_community import notices
    bundle = _wheel(tmp_path / "bundle.mcpb", {
        "manifest.json": json.dumps({"name": "kitchensink4ppt-business"}),
        "INSTALL.md": notices.COMMUNITY_GUIDANCE,
        "LICENSE": NOTICES,
    })
    with pytest.raises(build_script.BuildScanError, match="Community guidance"):
        build_script.scan_business_artifacts([bundle])


def test_stripping_a_required_notice_also_fails_the_build(tmp_path):
    """Zero Community guidance never means deleting a copyright notice."""
    stripped = _wheel(tmp_path / "stripped.whl", {"a/LICENSE": "nothing here"})
    with pytest.raises(build_script.BuildScanError, match="Required notices"):
        build_script.scan_business_artifacts([stripped])


def test_a_wheel_cannot_borrow_a_sibling_wheels_copyright_line(tmp_path):
    """A set-level check passes as long as SOMEONE in the set carries each
    notice, which lets a wheel whose notices went missing be excused by a
    sibling that still has them. Every wheel carries its own."""
    good = _wheel(tmp_path / "good.whl", {"a/LICENSE": NOTICES})
    bare = _wheel(tmp_path / "bare.whl", {"pkg/__init__.py": "x = 1\n"})
    with pytest.raises(build_script.BuildScanError, match="do not carry"):
        build_script.scan_business_artifacts([good, bare])


def test_the_gate_reports_which_artifact_carries_each_notice(tmp_path):
    """A count cannot tell you a disclaimer ended up in the wrong wheel."""
    wheel = _wheel(tmp_path / "one.whl", {"a/LICENSE": NOTICES})
    report = build_script.scan_business_artifacts([wheel])
    for notice in build_script.REQUIRED_NOTICES:
        assert report["carriers"][notice] == ["one.whl"]


@pytest.mark.business_sources
def test_the_allowlist_lets_the_commercial_notice_keep_its_agpl_sentences():
    """The commercial notice says the Community distribution remains
    available under AGPL-3.0. That sentence belongs in a Business artifact
    and a word-level scan would have deleted it."""
    notice = (ROOT / "business/COMMERCIAL_LICENSE_NOTICE.md").read_text("utf-8")
    assert "AGPL-3.0" in notice
    for needle in build_script.community_needles():
        assert needle not in notice


def test_the_business_install_doc_drops_only_the_community_section():
    text = build_script.business_install_doc()
    assert "## Community" not in text
    assert "## Business" in text
    assert "## Install and connect" in text
    for needle in build_script.community_needles():
        assert needle not in text


# ------------------------------------------------------ anchored rewrites

#: The build's version line, read from pyproject.toml rather than hard-coded,
#: so a dev bump does not break the anchor tests. The check itself is
#: unchanged.
_VERSION_ANCHOR = next(
    line for line in (ROOT / "pyproject.toml").read_text(encoding="utf-8").splitlines()
    if line.startswith("version = ")
)


@pytest.mark.business_sources
@pytest.mark.parametrize("anchor", [
    '    "kitchensink4ppt-core",',
    '    "kitchensink4-license-business",',
    _VERSION_ANCHOR,
], ids=["core", "license-business", "version"])
@pytest.mark.parametrize("drift", ["missing", "duplicate"])
def test_business_build_refuses_metadata_drift(tmp_path, anchor, drift):
    """The fail-closed property the whole split rests on: an anchor that is
    absent or ambiguous stops the build instead of silently skipping a
    rewrite."""
    source = ROOT / "packages/kitchensink4ppt-business/pyproject.toml"
    _assert_anchor_fails_closed(tmp_path, source, anchor, drift)


@pytest.mark.parametrize("anchor", [_VERSION_ANCHOR, 'license = "AGPL-3.0-only"'],
                         ids=["version", "license"])
@pytest.mark.parametrize("drift", ["missing", "duplicate"])
def test_core_relabelling_refuses_metadata_drift(tmp_path, anchor, drift):
    """The commercial core is produced by rewriting exactly these two lines.
    If either anchor drifts, the build stops rather than shipping a core
    wheel labelled with the wrong grant."""
    _assert_anchor_fails_closed(tmp_path, ROOT / "pyproject.toml", anchor, drift)


def _assert_anchor_fails_closed(tmp_path, source, anchor, drift):
    text = source.read_text("utf-8")
    assert text.count(anchor) == 1, f"anchor drifted out of {source}"
    # Duplicate anchors in comments keep TOML valid but expose ambiguity.
    drifted = (text + "\n# " + anchor + "\n" if drift == "duplicate"
               else text.replace(anchor, "", 1))
    staged = tmp_path / "pyproject.toml"
    staged.write_text(drifted, encoding="utf-8")
    with pytest.raises(ValueError, match="exactly one build metadata anchor"):
        build_script.staged_config(staged, [(anchor, anchor + "  # rewritten")])


@pytest.mark.parametrize("anchor", [
    '    "fastmcp>=3.4,<4",',
    '    "kitchensink4-license-community==0.2.0",',
    'metrics = ["fonttools>=4.0"]',
    'optimize = ["Pillow"]',
])
def test_community_launcher_anchors_are_present(anchor):
    source = ROOT / "packages/kitchensink4ppt-community/pyproject.toml"
    assert source.read_text("utf-8").count(anchor) == 1


def test_the_core_declares_no_console_scripts():
    """The public command names moved to the launchers, so nothing routes to
    the server without a policy being injected."""
    import tomllib
    config = tomllib.loads((ROOT / "pyproject.toml").read_text("utf-8"))
    assert "scripts" not in config["project"]
    assert config["project"]["name"] == "kitchensink4ppt-core"


def test_staged_metadata_that_stops_being_valid_toml_is_rejected(tmp_path):
    staged = tmp_path / "pyproject.toml"
    staged.write_text('[project]\nname = "x"\n', encoding="utf-8")
    with pytest.raises(Exception):
        build_script.staged_config(staged, [('name = "x"', 'name = "x\n[project]')])


# --------------------------------------------------- gate 2: the license


def _licensed_wheel(path, expression):
    members = {"pkg/__init__.py": "x = 1\n"}
    if expression is not None:
        members["pkg-1.0.dist-info/METADATA"] = (
            "Metadata-Version: 2.4\nName: pkg\nVersion: 1.0\n"
            f"License-Expression: {expression}\n\nbody\n")
    else:
        members["pkg-1.0.dist-info/METADATA"] = (
            "Metadata-Version: 2.4\nName: pkg\nVersion: 1.0\n\nbody\n")
    return _wheel(path, members)


def test_an_agpl_wheel_in_a_business_distribution_fails_the_build(tmp_path):
    """The dual-grant defect: a core wheel saying AGPL inside a commercial
    bundle puts two conflicting grants in one customer's hands."""
    wheel = _licensed_wheel(tmp_path / "core.whl", "AGPL-3.0-only")
    with pytest.raises(build_script.BuildScanError, match="Conflicting license"):
        build_script.scan_license_metadata([wheel], "business")


def test_a_commercial_wheel_in_a_community_distribution_fails_the_build(tmp_path):
    """Same defect wearing the other hat."""
    wheel = _licensed_wheel(tmp_path / "core.whl", "LicenseRef-KitchenSink-Commercial")
    with pytest.raises(build_script.BuildScanError, match="Conflicting license"):
        build_script.scan_license_metadata([wheel], "community")


@pytest.mark.parametrize("edition,expression", [
    ("business", "LicenseRef-KitchenSink-Commercial"),
    ("community", "AGPL-3.0-only"),
])
def test_a_correctly_labelled_distribution_passes(tmp_path, edition, expression):
    wheel = _licensed_wheel(tmp_path / "core.whl", expression)
    assert build_script.scan_license_metadata([wheel], edition)


def test_a_wheel_with_no_license_expression_fails_the_build(tmp_path):
    wheel = _licensed_wheel(tmp_path / "core.whl", None)
    with pytest.raises(build_script.BuildScanError, match="no license expression"):
        build_script.scan_license_metadata([wheel], "business")


def test_a_license_gate_with_no_wheels_refuses_to_pass_anything(tmp_path):
    with pytest.raises(build_script.BuildScanError, match="would pass anything"):
        build_script.scan_license_metadata([], "business")


# ------------------------------------------------- gate 3: the adapter pin


def test_the_adapter_gate_accepts_the_canonical_source():
    assert build_script.scan_adapter_provenance() == adapter_source.expected_digest()


def test_a_forked_adapter_fails_the_adapter_gate(tmp_path):
    """A fork is invisible in the built wheel and fatal in the index."""
    import shutil
    fork = tmp_path / "forked"
    shutil.copytree(adapter_source.ensure(), fork)
    # The last package the selected edition resolves: Business on the private
    # path, Community on the public Community-only path.
    target = next((fork / adapter_source.required_dists()[-1]).rglob("policy.py"))
    target.write_text(target.read_text(encoding="utf-8")
                      + '\nPRODUCT_ID = "kitchensink4ppt-business"\n',
                      encoding="utf-8")
    with pytest.raises(build_script.BuildScanError, match="not the pinned"):
        build_script.scan_adapter_provenance(fork)


def test_vendoring_the_adapters_into_this_repository_fails_the_gate(tmp_path,
                                                                    monkeypatch):
    vendored = tmp_path / "packages"
    (vendored / "kitchensink4-license-community").mkdir(parents=True)
    monkeypatch.setattr(build_script, "PACKAGES", vendored)
    with pytest.raises(build_script.BuildScanError, match="copied into this"):
        build_script.scan_adapter_provenance()


# --------------------------------------------------------- the real builds


@pytest.mark.business_sources
@pytest.mark.timeout(900)
def test_a_real_business_build_ships_a_commercially_labelled_core(tmp_path):
    """End to end: the Business core wheel is a distinct artifact identity
    with its own grant, built from the same source revision."""
    artifacts = build_script.build("business", tmp_path / "b", date(2026, 9, 13),
                                   None, "primary")
    declared = build_script.scan_license_metadata(artifacts, "business")
    assert all("AGPL" not in expression for expression in declared.values())
    core, = [n for n in declared if n.startswith("kitchensink4ppt_core-")]
    assert "+commercial" in core
    assert declared[core] == build_script.COMMERCIAL


@pytest.mark.timeout(900)
def test_a_real_community_build_ships_an_agpl_core_at_the_public_version(tmp_path):
    artifacts = build_script.build("community", tmp_path / "c", date(2026, 9, 13),
                                   None, "primary")
    declared = build_script.scan_license_metadata(artifacts, "community")
    assert all("LicenseRef" not in expression for expression in declared.values())
    app, = [a for a in artifacts if a.name.startswith("kitchensink4ppt-") and a.suffix == ".whl"]
    assert declared[app.name] == build_script.AGPL
    assert len([a for a in artifacts if a.suffix == ".whl"]) == 2
    assert not any(a.name.startswith("kitchensink4ppt_core-") for a in artifacts)
    with zipfile.ZipFile(app) as archive:
        names = archive.namelist()
        assert "kitchensink4ppt/__init__.py" in names
        assert "kitchensink4ppt/server.py" in names
        assert "kitchensink4ppt/templates/default.pptx" in names
        assert "ks4ppt_community_launch/__init__.py" in names
        metadata = archive.read(next(n for n in names if n.endswith(".dist-info/METADATA"))).decode()
        from email.parser import HeaderParser
        from packaging.requirements import Requirement
        headers = HeaderParser().parsestr(metadata)
        requirements = [Requirement(r) for r in headers.get_all("Requires-Dist", [])]
        policy, = [r for r in requirements if r.name == "kitchensink4-license-community"]
        assert str(policy.specifier) == "==0.2.0" and policy.marker is None
        assert not any(r.name == "kitchensink4ppt-core" for r in requirements)
        assert any(r.name == "fastmcp" for r in requirements)
        root_config = tomllib.loads((ROOT / "pyproject.toml").read_text("utf-8"))
        app_config = tomllib.loads((ROOT / "packages/kitchensink4ppt-community/pyproject.toml").read_text("utf-8"))
        assert set(app_config["project"]["dependencies"]) == {
            *root_config["project"]["dependencies"], "kitchensink4-license-community==0.2.0"}
        for extra in ("metrics", "optimize"):
            assert app_config["project"]["optional-dependencies"][extra] == root_config["project"]["optional-dependencies"][extra]
        assert any(r.name == "fonttools" and str(r.marker) == 'extra == "metrics"' for r in requirements)
        scripts = archive.read(next(n for n in names if n.endswith(".dist-info/entry_points.txt"))).decode()
        assert "ppt-mcp = ks4ppt_community_launch:main" in scripts
        assert "kitchensink4ppt = ks4ppt_community_launch:main" in scripts


@pytest.mark.business_sources
@pytest.mark.timeout(900)
def test_a_poisoned_core_label_fails_the_real_business_build(tmp_path, monkeypatch):
    """Skip the relabelling and the build must refuse to ship the bundle."""
    real = build_script.build_core

    def unlabelled(edition, output, scratch):
        return real("community", output, scratch)
    monkeypatch.setattr(build_script, "build_core", unlabelled)
    with pytest.raises((build_script.BuildScanError, ValueError)):
        build_script.build("business", tmp_path / "poison", date(2026, 9, 13),
                           None, "primary")


@pytest.mark.business_sources
@pytest.mark.timeout(900)
def test_the_required_notices_survive_the_commercial_relabelling(tmp_path):
    """Relabelling a distribution never means dropping a notice."""
    artifacts = build_script.build("business", tmp_path / "n", date(2026, 9, 13),
                                   None, "primary")
    assert build_script.scan_business_artifacts(artifacts)["required_notices"] == 2


@pytest.mark.business_sources
@pytest.mark.timeout(900)
def test_a_real_business_build_passes_the_gate_and_ships_no_community_module(tmp_path):
    """End to end: build the Business distribution and prove the gate ran on
    real artifacts rather than on fixtures shaped like them."""
    output = tmp_path / "business"
    artifacts = build_script.build("business", output, date(2026, 9, 13), None, "primary")
    names = [a.name for a in artifacts]
    assert any(n.startswith("kitchensink4ppt_core-") for n in names)
    assert any(n.startswith("kitchensink4_license_business-") for n in names)
    assert any(n.startswith("kitchensink4ppt_business-") for n in names)
    assert any(n.endswith(".mcpb") for n in names)
    for artifact in artifacts:
        with zipfile.ZipFile(artifact) as archive:
            assert not [n for n in archive.namelist() if "ks4_license_community" in n]
    # And the gate itself is not a no-op on this artifact set.
    assert build_script.scan_business_artifacts(artifacts)["members_scanned"] > 20


@pytest.mark.business_sources
@pytest.mark.timeout(900)
def test_a_poisoned_build_input_fails_the_real_build(tmp_path, monkeypatch):
    """Inject the guidance into a build input and the build must refuse."""
    from ks4_license_community import notices
    poisoned = build_script.business_install_doc() + "\n\n" + notices.WARNING_13
    monkeypatch.setattr(build_script, "business_install_doc", lambda: poisoned)
    with pytest.raises(build_script.BuildScanError, match="Community guidance"):
        build_script.build("business", tmp_path / "poisoned", date(2026, 9, 13),
                           None, "primary")


@pytest.mark.timeout(900)
def test_a_drifted_adapter_stops_the_real_build_before_any_wheel(tmp_path,
                                                                 monkeypatch):
    """The adapter gate runs first, so a fork never reaches a wheel."""
    import shutil
    fork = tmp_path / "forked"
    shutil.copytree(adapter_source.ensure(), fork)
    target = (fork / adapter_source.COMMUNITY_POLICY_DIST
              / "src/ks4_license_community/notices.py")
    target.write_text(target.read_text(encoding="utf-8") + "\n# fork\n",
                      encoding="utf-8")
    monkeypatch.setenv(adapter_source.SOURCE_ENV, str(fork))

    def forbidden(*args, **kwargs):
        pytest.fail("Build must stop before invoking the wheel builder")
    monkeypatch.setattr(build_script.subprocess, "run", forbidden)
    with pytest.raises((build_script.BuildScanError,
                        adapter_source.AdapterSourceError), match="not the pinned"):
        build_script.build("business", tmp_path / "drift", date(2026, 9, 13),
                           None, "primary")


def test_the_build_script_runs_headless_without_a_console_prompt():
    done = subprocess.run([sys.executable, "-X", "utf8",
                           str(ROOT / "scripts/build_editions.py"), "--help"],
                          capture_output=True, text=True, stdin=subprocess.DEVNULL)
    assert done.returncode == 0
    assert "--release-date" in done.stdout


# ------------------------------------------- the published long description
#
# kitchensink4ppt is the distribution this repository publishes to PyPI, and
# it is built from packages/kitchensink4ppt-community, which declared no
# readme. A release cut from here would have published the blessed one-line
# description and a project page with nothing on it, while the same release
# cut from the community branch carried the whole README, because there the
# published package is the root project and the root project declares one.
# build_launcher already stages the repository README beside the staged
# pyproject, so the declaration is the whole fix, and the assertion is against
# the built wheel rather than against the toml.


def test_the_community_wheel_carries_the_readme_as_its_long_description(
        tmp_path):
    artifacts = build_script.build("community", tmp_path / "c",
                                   date(2026, 9, 16), None)
    wheels = [a for a in artifacts
              if a.name.startswith("kitchensink4ppt-")
              and a.suffix == ".whl"]
    assert len(wheels) == 1, [a.name for a in artifacts]
    with zipfile.ZipFile(wheels[0]) as archive:
        name = next(n for n in archive.namelist()
                    if n.endswith(".dist-info/METADATA"))
        metadata = archive.read(name).decode("utf-8").replace(
            "\r\n", "\n")
    headers, _, body = metadata.partition("\n\n")
    assert "Description-Content-Type: text/markdown" in headers
    readme = (ROOT / "README.md").read_text("utf-8")
    first = readme.splitlines()[0]
    assert first and first in body
    # Not a truncated one: the README's own last non-empty line is in there.
    last = [line for line in readme.splitlines() if line.strip()][-1]
    assert last in body
    assert len(body) > len(readme) * 0.9
