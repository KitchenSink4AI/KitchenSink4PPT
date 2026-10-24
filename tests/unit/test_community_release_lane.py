"""Community publishing retains one app distribution and its license adapter."""
import importlib.util
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
_spec = importlib.util.spec_from_file_location("community_release_verify", ROOT / "scripts/release_verify.py")
rv = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(rv)


def _wheel(dist, kind, name, version, requires=()):
    directory = dist / kind
    directory.mkdir()
    normalized = name.replace("-", "_")
    path = directory / f"{normalized}-{version}-py3-none-any.whl"
    with zipfile.ZipFile(path, "w") as archive:
        headers = [f"Name: {name}", f"Version: {version}"]
        headers.extend(f"Requires-Dist: {r}" for r in requires)
        archive.writestr(f"{normalized}-{version}.dist-info/METADATA", "\n".join(headers) + "\n\n")
    return path


def _dist(tmp_path, requires):
    _wheel(tmp_path, "license", rv.LICENSE_DIST, "0.2.0")
    _wheel(tmp_path, "app", "kitchensink4ppt", "2.0.0", requires)
    return tmp_path


def test_community_lane_accepts_app_and_license(tmp_path):
    dist = _dist(tmp_path, ["kitchensink4-license-community==0.2.0", "fastmcp>=3.4,<4"])
    assert set(rv.check_build(ROOT, "ppt", "v2.0.0", dist)) == {"app", "license"}


def test_community_lane_rejects_core_dependency(tmp_path):
    dist = _dist(tmp_path, ["kitchensink4-license-community==0.2.0", "kitchensink4ppt-core==2.0.0"])
    with pytest.raises(rv.ReleaseCheckError, match="separate core"):
        rv.check_build(ROOT, "ppt", "v2.0.0", dist)


def test_community_lane_rejects_core_artifact(tmp_path):
    dist = _dist(tmp_path, ["kitchensink4-license-community==0.2.0"])
    _wheel(dist, "core", "kitchensink4ppt-core", "2.0.0")
    with pytest.raises(rv.ReleaseCheckError, match="unexpected entries"):
        rv.local_set(dist, "ppt")


def test_publish_workflow_has_only_existing_app_publisher():
    import yaml
    jobs = yaml.safe_load((ROOT / ".github/workflows/publish.yml").read_text("utf-8"))["jobs"]
    assert "publish-core" not in jobs
    assert jobs["verify-dependencies"]["needs"] == "plan"
    assert set(jobs["plan"]["outputs"]) == {"app", "license"}
    publishers = [step for job in jobs.values() for step in job.get("steps", [])
                  if step.get("uses", "").startswith("pypa/gh-action-pypi-publish@")]
    assert len(publishers) == 1
    assert publishers[0]["with"]["packages-dir"] == "dist/app/"
    assert jobs["registry"]["needs"] == "post-verify"
