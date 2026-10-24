"""Release-only joins between the fixed surface and public product files.

These checks run only when ``KS4P_RELEASE_GATE=1`` is set; ordinary runs
collect them and skip them.  publish.yml sets it before any artifact is built.
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
import tomllib
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[2]
PUBLIC = (
    ROOT / "README.md",
    ROOT / "docs" / "llms.txt",
    ROOT / "docs" / "index.html",
)
#: The files that carry the published figures. docs/index.html is a redirect
#: stub to the community site and carries none.
FIGURE_FILES = (
    ROOT / "README.md",
    ROOT / "docs" / "llms.txt",
)
public_copy = pytest.mark.public_copy

#: The figure that counts each action an operation can take, as the public
#: copy states it. ``scripts/count_operations.py`` derives it from committed
#: source and ``scripts/operations_snapshot.json`` records the per-operation
#: breakdown that produced it. Adding a dispatch value to any operation moves
#: the figure, which is a publishing decision: re-record the snapshot with
#: ``count_operations.py --json scripts/operations_snapshot.json``, restamp
#: every published figure and change this number in the same commit.
PUBLISHED_ACTIONS = 193
SNAPSHOT = ROOT / "scripts" / "operations_snapshot.json"


def _count_operations(*args: str) -> str:
    out = subprocess.run(
        [sys.executable, "-X", "utf8",
         str(ROOT / "scripts" / "count_operations.py"), *args],
        capture_output=True, text=True, encoding="utf-8", cwd=ROOT,
        stdin=subprocess.DEVNULL, timeout=300,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    assert out.returncode == 0, f"count_operations.py failed: {out.stderr}"
    return out.stdout


def _measurement() -> dict:
    sys.path.insert(0, str(ROOT / "scripts"))
    import measure_surface

    return measure_surface.measure()


@public_copy
def test_public_headlines_match_the_fixed_surface():
    measured = _measurement()
    assert (measured["tools"], measured["ops"]) == (9, 145)
    tokens = measured["approx_tokens"]
    token_forms = (str(tokens), f"{tokens:,}", f"{tokens / 1000:.1f}k")
    expected = {
        "tool count": re.compile(
            rf"\b{measured['tools']}(?:-tool|\s+tools?)\b", re.IGNORECASE
        ),
        "operation count": re.compile(
            rf"\b{measured['ops']}\s+(?:operations?|ops)\b", re.IGNORECASE
        ),
        "startup token estimate": re.compile(
            rf"(?:{'|'.join(map(re.escape, token_forms))})\s+"
            rf"(?:approximate\s+)?tokens?\b",
            re.IGNORECASE,
        ),
    }
    missing = []
    for path in FIGURE_FILES:
        text = path.read_text(encoding="utf-8")
        for label, pattern in expected.items():
            if pattern.search(text) is None:
                missing.append(
                    f"{path.relative_to(ROOT)}: {label} /{pattern.pattern}/"
                )
    assert not missing, "\n".join(missing)
    # The action count: the script's figure is the snapshot's and the copy's.
    report = _count_operations("--check", str(SNAPSHOT))
    assert "snapshot MATCHES" in report, (
        "count_operations.py disagrees with scripts/operations_snapshot.json. "
        "Re-record it with --json, then update the published figures and "
        "PUBLISHED_ACTIONS in this file.\n" + report
    )
    counted = int(
        re.search(r"TOTAL DISTINCT OPERATIONS:\s+(\d+)", report).group(1)
    )
    stored = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
    assert counted == stored["total_operations"] == PUBLISHED_ACTIONS
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert f"Your AI gets {PUBLISHED_ACTIONS} actions in PowerPoint" in readme
    assert (
        f"{measured['ops']} operations in {measured['tools']} tools, "
        f"{PUBLISHED_ACTIONS} counting each action"
    ) in readme
    for path in FIGURE_FILES:
        assert str(PUBLISHED_ACTIONS) in path.read_text(encoding="utf-8"), (
            f"{path.relative_to(ROOT)} is missing the actions count"
        )


@public_copy
def test_readme_and_llms_name_every_fixed_tool():
    from kitchensink4ppt.surface.plan import ALL_TOOLS

    for path in (ROOT / "README.md", ROOT / "docs" / "llms.txt"):
        text = path.read_text(encoding="utf-8")
        missing = [tool for tool in ALL_TOOLS if tool not in text]
        assert not missing, f"{path.relative_to(ROOT)} is missing {missing}"


@public_copy
def test_public_release_identity_matches_package_metadata():
    version = tomllib.loads(
        (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    )["project"]["version"]
    manifest = json.loads(
        (ROOT / "bundle" / "manifest.json").read_text(encoding="utf-8")
    )
    server = json.loads((ROOT / "server.json").read_text(encoding="utf-8"))
    assert manifest["version"] == version
    assert server["version"] == version
    assert server["packages"][0]["version"] == version
    assert len(server["description"]) <= 100
    llms = (ROOT / "docs" / "llms.txt").read_text(encoding="utf-8")
    assert [line for line in llms.splitlines() if line.startswith("Release:")] == [
        f"Release: v{version}"
    ]
    changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    assert f"### {version}" in changelog


@public_copy
@pytest.mark.parametrize(
    "path",
    (*PUBLIC, ROOT / "CHANGELOG.md", ROOT / "bundle" / "manifest.json",
     ROOT / "server.json"),
    ids=lambda path: str(path.relative_to(ROOT)),
)
def test_public_files_carry_no_writer_markers(path: Path):
    text = path.read_text(encoding="utf-8")
    for marker in ("COPY_SLOT:", "COPY PENDING", "FACTS TO CONVEY", "PLACEHOLDER"):
        assert marker not in text, f"{path.relative_to(ROOT)} carries {marker!r}"


def test_the_contributor_agreement_says_nothing_of_credits_on_our_own_commits():
    """The agreement is public text. It says what a tool credit means for
    authorship; it must not say that this project's own commits carry one."""
    text = (ROOT / "CLA.md").read_text(encoding="utf-8")
    assert "co-authored-by" not in text.lower()
    assert "crediting the tools" not in text


@public_copy
def test_llms_txt_token_figure_is_the_measured_figure():
    """llms.txt states the whole surface's cost as an exact number; it is the measured one."""
    measured = _measurement()["approx_tokens"]
    text = (ROOT / "docs" / "llms.txt").read_text(encoding="utf-8")
    written = [
        int(raw.replace(",", ""))
        for raw in re.findall(r"About ([\d,]+) approximate tokens", text)
    ]
    assert written, "docs/llms.txt no longer states the token figure"
    assert all(figure == measured for figure in written), (written, measured)
