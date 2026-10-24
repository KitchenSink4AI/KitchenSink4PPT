"""Regression coverage for the bundled from-scratch presentation template."""

from __future__ import annotations

import builtins
import hashlib
import tomllib
import zipfile
from importlib import resources
from pathlib import Path

from kitchensink4ppt.ops import slides


ROOT = Path(__file__).resolve().parents[2]
RESOURCE_NAME = "templates/default.pptx"
EXPECTED_SIZE = 25_721
EXPECTED_SHA256 = "eb33715284d0a4838d3c1ffaa1245df153fef590ea5e77fa8f9d2fe4bc6867e6"


def _pyproject() -> dict:
    with (ROOT / "pyproject.toml").open("rb") as stream:
        return tomllib.load(stream)


def test_default_template_is_declared_package_data_and_pinned():
    project = _pyproject()
    package_data = project["tool"]["setuptools"]["package-data"]
    assert RESOURCE_NAME in package_data["kitchensink4ppt"]

    template = resources.files("kitchensink4ppt").joinpath("templates", "default.pptx")
    assert template.is_file()
    payload = template.read_bytes()
    assert len(payload) == EXPECTED_SIZE
    assert hashlib.sha256(payload).hexdigest() == EXPECTED_SHA256


def test_default_template_carries_no_generator_identity_dates_or_thumbnail():
    template = resources.files("kitchensink4ppt").joinpath(
        "templates", "default.pptx"
    )

    with zipfile.ZipFile(template) as package:
        names = set(package.namelist())
        core = package.read("docProps/core.xml").decode("utf-8")
        root_rels = package.read("_rels/.rels").decode("utf-8")

    assert "generated using python-pptx" not in core
    assert "Steve Canny" not in core
    assert "2013-01-27" not in core
    assert "docProps/thumbnail.jpeg" not in names
    assert "relationships/metadata/thumbnail" not in root_rels


def test_python_pptx_remains_test_only():
    project = _pyproject()["project"]
    runtime = [requirement.lower() for requirement in project["dependencies"]]
    dev = [requirement.lower() for requirement in project["optional-dependencies"]["dev"]]
    assert not any(requirement.startswith("python-pptx") for requirement in runtime)
    assert any(requirement.startswith("python-pptx") for requirement in dev)


def test_create_presentation_uses_bundled_bytes_without_importing_pptx(
    monkeypatch, tmp_path
):
    real_import = builtins.__import__

    def refuse_python_pptx(name, *args, **kwargs):
        if name == "pptx" or name.startswith("pptx."):
            raise AssertionError("create_presentation attempted a runtime python-pptx import")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", refuse_python_pptx)
    output = tmp_path / "from-bundled-template.pptx"
    result = slides.create_presentation(output)

    assert output.is_file()
    assert result["slides_kept"] == 0
    assert result["template"].endswith(RESOURCE_NAME)
    with zipfile.ZipFile(output) as package:
        assert package.testzip() is None
        assert "ppt/presentation.xml" in package.namelist()
