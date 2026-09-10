"""Derive every figure this repo publishes, and stamp the one that cannot
be derived at runtime.

Three figures reach a stranger: the registered tool total, the lite core's
token bill, and the collected test count. The first two come off the live
pack registry, which is what `measure_surface.py` and `check_site.py`
already read, so a running server can never disagree with them. The test
count cannot: the suite is not in the wheel. This script counts it for real
(`pytest tests/ --collect-only -q`, run headless with no console window)
and writes it into `src/kitchensink4ppt/shipped.py`, which is what
`get_server_info` reports and what the ship gate holds the published
surfaces to.

Run:
    .venv/Scripts/python.exe -X utf8 scripts/stamp_figures.py           # report
    .venv/Scripts/python.exe -X utf8 scripts/stamp_figures.py --write   # stamp
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

SHIPPED = ROOT / "src" / "kitchensink4ppt" / "shipped.py"

#: Windows: keep the pytest collection off the author's desktop entirely.
_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def tool_figures() -> dict[str, object]:
    from kitchensink4ppt import packs, server  # noqa: F401

    names = packs.tool_names()
    lite_cost = sum(
        packs.approx_tokens(packs._REGISTRY["lite"][n]) for n in names["lite"]
    )
    full_cost = sum(
        packs.approx_tokens(packs._REGISTRY[pack][n])
        for pack, tools in names.items()
        for n in tools
    )
    return {
        "tools": sum(len(v) for v in names.values()),
        "lite": len(names["lite"]),
        "disabled": sum(len(v) for k, v in names.items() if k != "lite"),
        "lite_tokens": f"{lite_cost / 1000:.1f}k",
        "full_tokens": f"{full_cost / 1000:.1f}k",
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
    print("derived from the live pack registry:")
    for key in ("tools", "lite", "disabled", "lite_tokens", "full_tokens"):
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
              f"README.md, docs/llms.txt and docs/index.html to match.")
        return 1
    SHIPPED.write_text(
        re.sub(r"^TESTS = \d+$", f"TESTS = {tests}", current, flags=re.M),
        encoding="utf-8",
    )
    print(f"\nstamped shipped.TESTS: {stamped:,} -> {tests:,}. Now restamp "
          f"README.md, docs/llms.txt and docs/index.html.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
