"""The CHANGELOG names the restore operation by its real name.

The 2.0 entry's memory-bomb bullet said "`restore_backup` checks a backup or
a restore source". There is no operation named restore_backup (ppt_help
answers that it is not an operation); a restore is manage_backups with
action "restore". The bullet also said validate checks its file the same
way, while validate then still asked PowerPoint to open a file it had
refused; it now answers the refusal itself (test_validate_refused_files.py).
"""

from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def _current_entry() -> str:
    text = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    return text.split("\n### ")[1]


def test_the_changelog_names_no_restore_backup_operation():
    entry = _current_entry()
    assert "`restore_backup`" not in entry
    assert "A `manage_backups` restore checks its backup or source file" in entry


def test_the_changelog_says_validate_does_not_hand_a_refused_file_to_powerpoint():
    assert "does not ask PowerPoint to open the file" in _current_entry()
