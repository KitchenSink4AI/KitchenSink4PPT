"""Workflow and diagnostic contracts on the fixed operation surface."""

from __future__ import annotations

import re

import pytest

from kitchensink4ppt import server
from kitchensink4ppt.core.errors import PptMcpError
from kitchensink4ppt.ops import workflows as _wf
from kitchensink4ppt.ops.diagnostics import diagnose
from kitchensink4ppt.registry import GROUP_OPS, OP_TO_GROUP, public_ops
from kitchensink4ppt.surface import copy


def test_index_lists_all_workflows_without_retired_surface_controls():
    out = _wf.get_workflows()
    assert set(out["workflows"]) == set(_wf.WORKFLOWS)
    for entry in out["workflows"].values():
        assert entry["summary"]
        assert "packs" not in entry


def test_every_step_names_a_registered_operation():
    registered = set(public_ops())
    for name, workflow in _wf.WORKFLOWS.items():
        assert "packs" not in workflow, name
        for step in workflow["steps"]:
            assert step["tool"] in GROUP_OPS, (
                f"workflow {name} references unknown tool {step['tool']}"
            )
            assert step["op"] in registered, (
                f"workflow {name} references unknown operation {step['op']}"
            )
            assert OP_TO_GROUP[step["op"]] == step["tool"], (
                f"workflow {name} puts {step['op']} under {step['tool']}"
            )
            assert step["why"]


def test_every_surface_group_is_represented_by_a_workflow_step():
    named = {
        step["tool"]
        for workflow in _wf.WORKFLOWS.values()
        for step in workflow["steps"]
    }
    assert named == set(GROUP_OPS)


def test_every_customer_workflow_string_is_filled_copy():
    # The recipe prose is filled (the released wording): every string is
    # present and carries no slot marker, no pack word and no retired
    # toggle. The release gate scans the same strings in the built wheel.
    values: list[str] = []
    for workflow in _wf.WORKFLOWS.values():
        values.append(workflow["summary"])
        values.extend(step["why"] for step in workflow["steps"])
        values.extend(workflow["notes"])
    assert values and all(value.strip() for value in values)
    for value in values:
        assert copy.SLOT not in value and "FACTS TO CONVEY" not in value, value
        assert "enable_tools" not in value and "disable_tools" not in value, value
        assert not re.search(r"packs?", value.lower()), value


def test_unknown_task_refuses_with_menu():
    with pytest.raises(PptMcpError) as exc:
        _wf.get_workflows("make-it-pop")
    assert "build-a-diagram" in str(exc.value)


def test_diagnose_environment():
    out = diagnose()
    assert out["server"] == "kitchensink4ppt"
    assert "engines" in out
    assert "sandbox" in out
    assert "active" in out["sandbox"]


def test_diagnose_file(make_deck):
    deck = make_deck("diag.pptx")
    out = diagnose(str(deck))
    f = out["file"]
    assert f["exists"] is True
    assert f["opens_as_package"] is True
    assert f["slide_count"] >= 1
    assert f["writable"] is True


def test_diagnose_missing_file(tmp_path):
    out = diagnose(str(tmp_path / "ghost.pptx"))
    assert out["file"]["exists"] is False
    assert "problem" in out["file"]


def test_diagnose_server_operation_adds_fixed_surface():
    out = server.diagnose()
    assert out["surface"]["fixed"] is True
    assert out["surface"]["operations"] == 145
    assert out["surface"]["tools"] == {
        name: len(ops) for name, ops in GROUP_OPS.items()
    }
