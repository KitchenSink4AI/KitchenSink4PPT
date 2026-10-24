"""The COM-gated disconnect test ends only the PowerPoint its own server started.

test_family_surface_guards.test_v13 ended every POWERPNT.EXE that appeared
during its run, by pid, so on a machine where someone else's PowerPoint
starts meanwhile (the user's, another session's) it ended that one too, and
a recycled pid could name an unrelated process. Its cleanup must name only
the processes its server's own ownership journal records, each matched by
pid and creation time. COM starts PowerPoint from its own service process,
so the parent chain cannot name the server's PowerPoint.

This guard reads the test's source, so it runs everywhere; the test itself
needs KS4_VERIFY_PPT_COM=1 and a machine where no PowerPoint is running.
"""

from __future__ import annotations

import inspect

import test_family_surface_guards as guards


def test_v13_ends_powerpoint_only_by_its_servers_journal():
    source = inspect.getsource(
        guards.test_v13_disconnect_during_a_com_call_leaves_no_powerpoint_this_server_started
    )
    assert "taskkill" not in source
    assert "- before" not in source, "a PowerPoint that merely appeared is not ours"
    assert "_end_if_same(" in source and "KS4_V13_OWNED" in source


def test_the_cleanup_checks_pid_creation_time_and_image():
    source = inspect.getsource(guards._end_if_same)
    assert "_creation_time(pid, handle=handle) != created" in source
    assert "QueryFullProcessImageNameW" in source
    assert "taskkill" not in source and "/IM" not in source
