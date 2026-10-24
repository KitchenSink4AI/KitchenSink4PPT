"""Derive every figure this repo publishes, and stamp the one that cannot
be derived at runtime.

Three figures reach a stranger: the fixed tool total, operation total, and
the collected test count. The first two come from the same immutable startup
plan passed to ``create_server``, which is what `measure_surface.py` and
`check_site.py` also read. The test count cannot: the suite is not in the
wheel. This script counts it for real
(`pytest tests/ --collect-only -q`, run headless with no console window)
and writes it into `src/kitchensink4ppt/shipped.py`, which is what
`get_server_info` reports and what the ship gate holds the published
surfaces to.

Run:
    .venv/Scripts/python.exe -X utf8 scripts/stamp_figures.py           # report
    .venv/Scripts/python.exe -X utf8 scripts/stamp_figures.py --write   # stamp
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

# The core does not choose an edition. A script that starts it is doing a
# launcher's job and has to pick one on purpose.
import _bootstrap  # noqa: E402

SHIPPED = ROOT / "src" / "kitchensink4ppt" / "shipped.py"

#: Windows: keep the pytest collection off the author's desktop entirely.
_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def tool_figures() -> dict[str, object]:
    _bootstrap.use_community()
    from kitchensink4ppt import registry, server
    from kitchensink4ppt.settings import Settings
    from kitchensink4ppt.surface import plan

    settings = Settings()
    surface = plan.build(settings, registry.public_ops())
    mcp = server.create_server(settings)
    approximate_tokens = 0
    for item in surface.tools.values():
        wire = {
            "name": item.name,
            "title": item.title,
            "description": item.description,
            "inputSchema": item.input_schema,
            "annotations": item.annotations,
        }
        approximate_tokens += round(
            len(json.dumps(wire, ensure_ascii=False, separators=(",", ":"))) / 4
        )
    return {
        "tools": len(surface.tools),
        "operations": len(surface.op_tool),
        "approx_tokens": approximate_tokens,
        "instructions_chars": len(mcp.instructions or ""),
    }


def test_count() -> int:
    """The real collection, counted rather than remembered."""
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", "-m", "pytest", "tests/",
         "--collect-only", "-q"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        stdin=subprocess.DEVNULL,
        timeout=600,
        creationflags=_NO_WINDOW,
    )
    tail = proc.stdout.strip().splitlines()
    for line in reversed(tail):
        hit = re.search(r"(\d+)\s+tests?\s+collected", line)
        if hit:
            return int(hit.group(1))
    raise SystemExit(
        "could not read a collection count out of pytest:\n"
        + "\n".join(tail[-10:]) + "\n" + proc.stderr[-2000:]
    )


def main() -> int:
    figures = tool_figures()
    tests = test_count()
    print("derived from the fixed startup plan:")
    for key in ("tools", "operations", "approx_tokens", "instructions_chars"):
        print(f"  {key:<12} {figures[key]}")
    print("derived from pytest --collect-only:")
    print(f"  {'tests':<12} {tests:,}")

    current = SHIPPED.read_text(encoding="utf-8")
    stamped = int(re.search(r"^TESTS = (\d+)$", current, re.M).group(1))
    if stamped == tests:
        print(f"\nshipped.TESTS is current ({tests:,}).")
        return 0
    if "--write" not in sys.argv:
        print(f"\nshipped.TESTS says {stamped:,}, the suite collects "
              f"{tests:,}. Re-run with --write to stamp it, then restamp "
              f"README.md and docs/llms.txt to match.")
        return 1
    # LF on every platform, as the repository keeps the file.
    SHIPPED.write_text(
        re.sub(r"^TESTS = \d+$", f"TESTS = {tests}", current, flags=re.M),
        encoding="utf-8",
        newline="\n",
    )
    print(f"\nstamped shipped.TESTS: {stamped:,} -> {tests:,}. Now restamp "
          f"README.md and docs/llms.txt.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
