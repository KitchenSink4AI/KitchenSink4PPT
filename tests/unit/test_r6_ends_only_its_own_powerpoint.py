"""The COM-gated exit-ownership test ends only the PowerPoint its own server started.

test_output_and_restore_guards.test_r6 took every POWERPNT.EXE that appeared
after its server started as the server's, and its cleanup ended each one
still alive by pid, so a PowerPoint that someone else started meanwhile (the
user's, another session's) was ended too, and a recycled pid could name an
unrelated process. It now names the server's PowerPoint from the server's
own ownership journal (pid and creation time, written out by a hook in the
server, as test_v13 does), attaches only while that is the only PowerPoint,
quits it only while it is still invisible with no presentations, and ends
a journalled process only through a handle that shows the recorded creation
time and POWERPNT.EXE.

This guard reads the test's source, so it runs everywhere; the test itself
needs KS4_VERIFY_PPT_COM=1 and a machine where no PowerPoint is running.
"""

from __future__ import annotations

import inspect

import test_output_and_restore_guards as guards


def _source() -> str:
    return inspect.getsource(
        guards.test_r6_exit_never_ends_an_owned_instance_that_now_holds_someone_elses_presentation)


def test_r6_names_the_servers_powerpoint_from_its_journal():
    source = _source()
    assert "KS4_R6_OWNED" in source and "KS4_R6_OWNED" in guards._SLOW_OPEN
    assert "_OWNED_JOURNAL" in guards._SLOW_OPEN
    assert "started = bridge.powerpnt_pids()" not in source, "a PowerPoint that merely appeared is not ours"


def test_r6_ends_nothing_but_a_journalled_process_matched_by_creation_time():
    source = _source()
    assert "taskkill" not in source and "/IM" not in source
    assert "_end_if_same(" in source
    assert "& started" not in source
