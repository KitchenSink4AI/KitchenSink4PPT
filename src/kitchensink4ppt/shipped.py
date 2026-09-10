"""Measured figures this build publishes. Generated, never hand-typed.

Two rules keep this file honest:

- Anything derivable at runtime is NOT here. get_server_info reads the tool
  totals straight off the pack registry (``packs.tool_names()``), the same
  source ``scripts/measure_surface.py`` and ``scripts/check_site.py`` read,
  so no tool count can ever go stale in a running server.
- The test count CANNOT be derived at runtime: the suite does not ship in
  the wheel. So it is stamped here by ``scripts/stamp_figures.py --write``,
  which counts the real collection, and it is guarded from both sides by
  ``tests/unit/test_ship_gate.py``: the number below has to match the test
  figure published on README.md, docs/llms.txt, and docs/index.html.

Re-stamp with::

    .venv/Scripts/python.exe -X utf8 scripts/stamp_figures.py --write
"""

from __future__ import annotations

#: Tests collected by ``pytest tests/ --collect-only -q``.
TESTS = 1251
