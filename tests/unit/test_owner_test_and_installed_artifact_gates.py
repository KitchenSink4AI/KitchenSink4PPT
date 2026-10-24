"""Regressions for the owner-test and installed-artifact gates."""

from __future__ import annotations

import importlib.util
import zipfile
from datetime import date
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[2]


def _load_script(name: str):
    path = ROOT / "scripts" / name
    spec = importlib.util.spec_from_file_location(f"_phase_d_{path.stem}", path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _wheel(path: Path, copy_source: str, *, module: str = "surface/copy.py") -> Path:
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(f"kitchensink4ppt/{module}", copy_source)
    return path


def test_owner_test_gate_catches_runtime_assembled_slot_call(tmp_path):
    build = _load_script("build_editions.py")
    artifact = _wheel(
        tmp_path / "slot.whl",
        "def _slot(name):\n    return name\n\nTITLE = _slot('TITLE')\n",
    )
    with pytest.raises(SystemExit, match="unfilled copy slot"):
        build.scan_owner_test([artifact])


def test_owner_test_gate_catches_split_slot_sentinel_used_by_current_copy(tmp_path):
    build = _load_script("build_editions.py")
    artifact = _wheel(
        tmp_path / "split-slot.whl",
        'SLOT = "COPY" "_SLOT:"\nTITLE = f"{SLOT} TITLE"\n',
    )
    with pytest.raises(SystemExit, match="unfilled copy slot"):
        build.scan_owner_test([artifact])


def test_owner_test_gate_accepts_filled_copy_module(tmp_path):
    build = _load_script("build_editions.py")
    artifact = _wheel(tmp_path / "filled.whl", "TITLE = 'PowerPoint tools'\n")
    report = build.scan_owner_test([artifact])
    assert report["members_scanned"] == 1


def test_owner_test_gate_catches_slot_call_after_messages_move(tmp_path):
    build = _load_script("build_editions.py")
    artifact = _wheel(
        tmp_path / "messages.whl",
        "def _slot(name):\n    return name\n\nHINT = _slot('HINT')\n",
        module="core/messages.py",
    )
    with pytest.raises(SystemExit, match="unfilled copy slot"):
        build.scan_owner_test([artifact])


def test_owner_test_gate_catches_split_slot_after_messages_move(tmp_path):
    build = _load_script("build_editions.py")
    artifact = _wheel(
        tmp_path / "messages-split.whl",
        'SLOT = "COPY" "_SLOT:"\nHINT = f"{SLOT} HINT"\n',
        module="core/messages.py",
    )
    with pytest.raises(SystemExit, match="unfilled copy slot"):
        build.scan_owner_test([artifact])


def test_refused_build_never_publishes_and_cleans_only_its_stage(
    tmp_path, monkeypatch
):
    build = _load_script("build_editions.py")
    output = tmp_path / "community"
    unrelated = tmp_path / "keep.txt"
    unrelated.write_text("mine", encoding="utf-8")

    def refuse_in_stage(_edition, staging, *_args, **_kwargs):
        assert not output.exists()
        staging.mkdir()
        (staging / "refused.whl").write_bytes(b"candidate")
        raise SystemExit("owner-test ship gate refused the build")

    monkeypatch.setattr(build, "_build_staged", refuse_in_stage)
    with pytest.raises(SystemExit, match="owner-test ship gate refused"):
        build.build(
            "community", output, date(2026, 10, 1), owner_test=True
        )

    assert not output.exists()
    assert unrelated.read_text("utf-8") == "mine"
    assert not list(tmp_path.glob(".community.staging-*"))


def test_build_publishes_only_after_staged_gates_pass(tmp_path, monkeypatch):
    build = _load_script("build_editions.py")
    output = tmp_path / "community"
    observed = []

    def pass_in_stage(_edition, staging, *_args, **_kwargs):
        observed.append(output.exists())
        staging.mkdir()
        artifact = staging / "passed.whl"
        artifact.write_bytes(b"candidate")
        observed.append(output.exists())
        return [artifact]

    monkeypatch.setattr(build, "_build_staged", pass_in_stage)
    artifacts = build.build(
        "community", output, date(2026, 10, 1), owner_test=True
    )

    assert observed == [False, False]
    assert artifacts == [output / "passed.whl"]
    assert artifacts[0].read_bytes() == b"candidate"
    assert not list(tmp_path.glob(".community.staging-*"))


def test_both_editions_are_one_atomic_candidate(tmp_path, monkeypatch):
    build = _load_script("build_editions.py")
    output = tmp_path / "edition-dist"
    unrelated = tmp_path / "keep.txt"
    unrelated.write_text("mine", encoding="utf-8")

    def fail_second(edition, destination, *_args, **_kwargs):
        destination.mkdir(parents=True)
        artifact = destination / f"{edition}.whl"
        artifact.write_bytes(edition.encode())
        if edition == "business":
            raise SystemExit("business gate refused")
        return [artifact]

    monkeypatch.setattr(build, "build", fail_second)
    with pytest.raises(SystemExit, match="business gate refused"):
        build.build_set(
            ("community", "business"), output, date(2026, 10, 1),
            owner_test=True,
        )

    assert not output.exists()
    assert unrelated.read_text("utf-8") == "mine"
    assert not list(tmp_path.glob(".edition-dist.set-staging-*"))


def test_both_editions_publish_together_after_all_pass(tmp_path, monkeypatch):
    build = _load_script("build_editions.py")
    output = tmp_path / "edition-dist"
    observed = []

    def pass_edition(edition, destination, *_args, **_kwargs):
        observed.append(output.exists())
        destination.mkdir(parents=True)
        artifact = destination / f"{edition}.whl"
        artifact.write_bytes(edition.encode())
        return [artifact]

    monkeypatch.setattr(build, "build", pass_edition)
    artifacts = build.build_set(
        ("community", "business"), output, date(2026, 10, 1),
        owner_test=True,
    )

    assert observed == [False, False]
    assert artifacts == {
        "community": [output / "community" / "community.whl"],
        "business": [output / "business" / "business.whl"],
    }
    assert not list(tmp_path.glob(".edition-dist.set-staging-*"))


def test_smoke_has_owner_test_mode_and_no_python_pptx_oracle():
    smoke = (ROOT / "scripts" / "smoke_editions.py").read_text("utf-8")
    assert '"--owner-test"' in smoke
    assert "from pptx import Presentation" not in smoke
    assert '"ppt_deck", {"op": "create_presentation"' in smoke


def test_owner_test_release_build_requires_git_archive_marker(tmp_path, monkeypatch):
    build = _load_script("build_editions.py")
    marker = tmp_path / "ARCHIVE_REVISION.txt"
    marker.write_text("$Format:%H$\n", encoding="ascii")
    monkeypatch.setattr(build, "ARCHIVE_REVISION_FILE", marker)
    with pytest.raises(SystemExit, match="git archive HEAD"):
        build.require_archive_source()

    marker.write_text("a" * 40 + "\n", encoding="ascii")
    assert build.require_archive_source() == "a" * 40


def test_artifact_audit_allows_only_the_pinned_bundled_template():
    smoke = _load_script("smoke_editions.py")
    assert smoke.TEMPLATE_MEMBER == "kitchensink4ppt/templates/default.pptx"
    assert (
        smoke.TEMPLATE_SHA256
        == "eb33715284d0a4838d3c1ffaa1245df153fef590ea5e77fa8f9d2fe4bc6867e6"
    )


def test_smoke_launch_environment_removes_retired_policy_controls():
    smoke = _load_script("smoke_editions.py")
    source = {
        "PATH": "kept",
        "KS4P_MODE": "lite",
        "KS4P_PACK_POLICY": "auto",
        "KS4P_ALL_TOOLS": "false",
        "KS4P_LOCK_TOOLS": "false",
        "KS4P_LICENSE_FILE": "stale-license.json",
    }

    clean = smoke.clean_runtime_env(source)

    assert clean["PATH"] == "kept"
    assert clean["KS4P_UPDATE_CHECK"] == "off"
    assert clean["PYTHONDONTWRITEBYTECODE"] == "1"
    assert not (set(smoke.RETIRED_POLICY_ENV) & set(clean))
    assert "KS4P_LICENSE_FILE" not in clean


def test_publish_workflow_runs_release_gate_and_builds_the_community_edition():
    """The public repository carries the Community edition only, so its release
    builds Community alone. The two-edition build and smoke (with their
    Community/Business parity checks) run in private staging, where the
    Business sources are."""
    workflow = (ROOT / ".github" / "workflows" / "publish.yml").read_text(
        "utf-8"
    )
    assert 'KS4P_RELEASE_GATE: "1"' in workflow
    assert "python -m pytest tests/unit -q -m \"not live\"" in workflow
    assert "python -X utf8 scripts/check_site.py" in workflow
    assert "scripts/build_editions.py" in workflow
    assert "--edition community" in workflow
    assert "--edition both" not in workflow
    assert "--owner-test" in workflow
    assert "python -m build" not in workflow
    assert "edition-dist/community/kitchensink4ppt_core-*.whl" not in workflow
    assert "edition-dist/community/kitchensink4ppt-[0-9]*.whl" in workflow
    assert "release_verify.py check-build --product ppt" in workflow
    # No Business build and no edition artifact in a public GitHub run.
    assert "name: edition-artifacts" not in workflow
    assert "path: edition-dist/" not in workflow
    assert "plan:\n    needs: build" in workflow


def test_staging_ci_is_reduced_and_cancels_superseded_runs():
    workflow = (ROOT / ".github" / "workflows" / "tests.yml").read_text(
        "utf-8"
    )
    assert "cancel-in-progress: true" in workflow
    assert "fail-fast: false" in workflow
    assert "os: [ubuntu-latest, windows-latest]" in workflow
    assert 'python: ["3.14"]' in workflow
    assert 'python: ["3.12", "3.13", "3.14"]' in workflow
    assert "macos" not in workflow
    assert "ref: ci/license-adapters-9c7dc45" in workflow
    assert "KS4PPT_ADAPTER_SOURCE=" in workflow
    assert 'pip install -e ".[dev,metrics,optimize]"' in workflow
    assert "release-gate:" in workflow
    assert "needs: full-matrix" in workflow
    assert "tests/unit/test_public_copy_gate.py" in workflow
    assert workflow.count("python -X utf8 scripts/check_site.py") == 1
