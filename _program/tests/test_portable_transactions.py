from __future__ import annotations

from pathlib import Path
import os
import shutil

import pytest

from red_onion_integrity import sha256_file
from red_onion_transactions import ReportStore, RecoveryRequired, atomic_bytes, json_read
from red_onion_runtime import write_json_atomic


class PowerLoss(BaseException):
    pass


def workspace(tmp_path: Path):
    incoming, output, archive = [tmp_path / name for name in ("Daily Reports", "Finished Reports", "Archive")]
    for directory in (incoming, output, archive):
        directory.mkdir()
    store = ReportStore(archive, output, incoming)
    (output / "Richmond.xlsx").write_bytes(b"previous Richmond")
    (output / "Virginia Beach.xlsx").write_bytes(b"previous Virginia Beach")
    stage = tmp_path / "test staging"
    stage.mkdir()
    staged = []
    for name in ("Richmond.xlsx", "Virginia Beach.xlsx"):
        path = stage / name
        path.write_bytes(("new " + name).encode())
        staged.append(path)
    source = incoming / "renamed daily.xlsx"
    source.write_bytes(b"captured daily")
    saved = archive / "saved input.xlsx"
    saved.write_bytes(source.read_bytes())
    inputs = [{"name": source.name, "archive": saved.name, "sha256": sha256_file(source), "quarantine": ".input-cleanup"}]
    expected = {p.name: sha256_file(output / p.name) for p in staged}
    return store, staged, inputs, expected


@pytest.mark.parametrize("boundary", [
    "journal_saved", "output_0_displaced", "output_0_published", "output_1_displaced",
    "output_1_published", "before_commit", "after_commit",
])
def test_power_loss_recovers_previous_or_new_complete_set(tmp_path, boundary):
    store, staged, inputs, expected = workspace(tmp_path)
    old = {p.name: (store.output_dir / p.name).read_bytes() for p in staged}
    def crash(stage):
        if stage == boundary:
            raise PowerLoss(stage)
    with pytest.raises(PowerLoss):
        store.publish("test-run", staged, {"previous_run_id": None}, inputs, expected, checkpoint=crash)
    store.recover()
    after_commit = boundary == "after_commit"
    for path in staged:
        assert (store.output_dir / path.name).read_bytes() == (path.read_bytes() if after_commit else old[path.name])
    assert (store.input_dir / inputs[0]["name"]).exists() is not after_commit
    assert (store.current() is not None) is after_commit
    assert not list(store.output_dir.glob("*.previous"))


def test_exception_before_commit_rolls_back_immediately(tmp_path):
    store, staged, inputs, expected = workspace(tmp_path)
    def fail(stage):
        if stage == "output_1_displaced":
            raise OSError("disk became unavailable")
    with pytest.raises(OSError, match="unavailable"):
        store.publish("failed-run", staged, {"previous_run_id": None}, inputs, expected, checkpoint=fail)
    assert {p.name: sha256_file(store.output_dir / p.name) for p in staged} == expected
    assert store.current() is None
    assert (store.input_dir / inputs[0]["name"]).exists()


def test_late_error_preserves_committed_publication_and_retries_only_cleanup(tmp_path):
    store, staged, inputs, expected = workspace(tmp_path)
    def fail(stage):
        if stage == "after_commit":
            raise OSError("status unavailable")
    outputs, warnings = store.publish("published-run", staged, {"previous_run_id": None}, inputs, expected, checkpoint=fail)
    assert warnings and len(outputs) == 2
    assert store.current()["run_id"] == "published-run"
    store.recover()
    assert not (store.input_dir / inputs[0]["name"]).exists()
    assert json_read(store.runs / "published-run/journal.json")["phase"] == "complete"


def test_recovery_preserves_external_edit_and_previous_good_copy(tmp_path):
    store, staged, inputs, expected = workspace(tmp_path)
    def crash(stage):
        if stage == "output_0_published":
            raise PowerLoss()
    with pytest.raises(PowerLoss):
        store.publish("interrupted", staged, {"previous_run_id": None}, inputs, expected, checkpoint=crash)
    target = store.output_dir / "Richmond.xlsx"
    target.write_bytes(b"new manager edit")
    with pytest.raises(RecoveryRequired, match="newer edit"):
        store.recover()
    assert target.read_bytes() == b"new manager edit"
    assert sha256_file(store.runs / "interrupted/previous/Richmond.xlsx") == expected["Richmond.xlsx"]


def test_cleanup_does_not_delete_new_input_with_old_name(tmp_path):
    store, staged, inputs, expected = workspace(tmp_path)
    def replace(stage):
        if stage == "after_commit":
            (store.input_dir / inputs[0]["name"]).write_bytes(b"corrected later")
    _, warnings = store.publish("changed-input", staged, {"previous_run_id": None}, inputs, expected, checkpoint=replace)
    assert any("changed input" in w for w in warnings)
    assert (store.input_dir / inputs[0]["name"]).read_bytes() == b"corrected later"


def test_second_computer_needs_only_portable_workspace(tmp_path):
    source = tmp_path / "PC A"
    source.mkdir()
    store, staged, inputs, expected = workspace(source)
    store.publish("first", staged, {"previous_run_id": None}, inputs, expected)
    copied = tmp_path / "PC B different username and path"
    shutil.copytree(source, copied)
    second = ReportStore(copied / "Archive", copied / "Finished Reports", copied / "Daily Reports")
    assert second.current()["run_id"] == "first"
    (copied / "test staging/Richmond.xlsx").write_bytes(b"PC B corrected report")
    new_paths = [copied / "test staging" / p.name for p in staged]
    second.publish("second", new_paths, {"previous_run_id": "first"}, [], {p.name: sha256_file(second.output_dir / p.name) for p in new_paths})
    assert second.current()["run_id"] == "second"
    # Sequential sync back to A carries the new head; no stale machine state.
    synced = tmp_path / "PC A after sync"
    shutil.copytree(copied, synced)
    third = ReportStore(synced / "Archive", synced / "Finished Reports", synced / "Daily Reports")
    assert third.current()["run_id"] == "second"


def test_formatted_old_output_and_removed_old_snapshot_do_not_block_next_run(tmp_path):
    store, staged, inputs, expected = workspace(tmp_path)
    store.publish("first", staged, {"previous_run_id": None}, inputs, expected)
    (store.output_dir / "Richmond.xlsx").write_bytes(b"Excel saved this differently")
    (store.runs / "first/outputs/Richmond.xlsx").unlink()
    latest = {p.name: sha256_file(store.output_dir / p.name) for p in staged}
    store.publish("second", staged, {"previous_run_id": "first"}, [], latest)
    assert store.current()["run_id"] == "second"


def test_missing_pointer_recovers_from_completed_journal(tmp_path):
    store, staged, inputs, expected = workspace(tmp_path)
    store.publish("completed", staged, {"previous_run_id": None}, inputs, expected)
    store.pointer.unlink()
    warnings = store.recover()
    assert warnings and store.current()["run_id"] == "completed"


def test_atomic_exclusive_copy_does_not_overwrite_new_file(tmp_path):
    path = tmp_path / "report.xlsx"
    path.write_bytes(b"new unrelated edit")
    with pytest.raises(FileExistsError):
        atomic_bytes(path, b"replacement", exclusive=True)
    assert path.read_bytes() == b"new unrelated edit"


def test_rollback_preserves_manager_save_between_check_and_removal(tmp_path, monkeypatch):
    store, staged, inputs, expected = workspace(tmp_path)
    def crash(stage):
        if stage == "output_0_published":
            raise PowerLoss()
    with pytest.raises(PowerLoss):
        store.publish("racing-edit", staged, {"previous_run_id": None}, inputs, expected, checkpoint=crash)
    target = store.output_dir / "Richmond.xlsx"
    rename = os.rename
    def edit_then_rename(source, destination):
        if Path(source) == target and str(destination).endswith(".rollback"):
            target.write_bytes(b"manager saved during rollback")
        return rename(source, destination)
    monkeypatch.setattr(os, "rename", edit_then_rename)
    with pytest.raises(RecoveryRequired, match="newer edit"):
        store.recover()
    assert target.read_bytes() == b"manager saved during rollback"
    assert sha256_file(store.runs / "racing-edit/previous/Richmond.xlsx") == expected["Richmond.xlsx"]


def test_interrupted_rollback_resumes_after_quarantine_rename(tmp_path, monkeypatch):
    store, staged, inputs, expected = workspace(tmp_path)
    def crash(stage):
        if stage == "output_0_published":
            raise PowerLoss()
    with pytest.raises(PowerLoss):
        store.publish("rollback-crash", staged, {"previous_run_id": None}, inputs, expected, checkpoint=crash)
    rename = os.rename
    def crash_after_rename(source, destination):
        rename(source, destination)
        if str(destination).endswith(".rollback"):
            raise PowerLoss()
    monkeypatch.setattr(os, "rename", crash_after_rename)
    with pytest.raises(PowerLoss):
        store.recover()
    monkeypatch.setattr(os, "rename", rename)
    store.recover()
    assert {p.name: sha256_file(store.output_dir / p.name) for p in staged} == expected


def test_error_after_pointer_install_remains_committed(tmp_path, monkeypatch):
    store, staged, inputs, expected = workspace(tmp_path)
    write_pointer = store._write_pointer
    def install_then_fail(run_id):
        write_pointer(run_id)
        raise OSError("late pointer-wrapper failure")
    monkeypatch.setattr(store, "_write_pointer", install_then_fail)
    _, warnings = store.publish("late-install-error", staged, {"previous_run_id": None}, inputs, expected)
    assert warnings
    assert store.current()["run_id"] == "late-install-error"
    assert all((store.output_dir / p.name).read_bytes() == p.read_bytes() for p in staged)
    store.recover()
    assert not (store.input_dir / inputs[0]["name"]).exists()


def test_missing_receipt_uses_successor_order_not_another_computers_clock(tmp_path):
    store, staged, inputs, expected = workspace(tmp_path)
    store.publish("first-pc", staged, {"previous_run_id": None}, inputs, expected)
    new_expected = {p.name: sha256_file(store.output_dir / p.name) for p in staged}
    store.publish("second-pc", staged, {"previous_run_id": "first-pc"}, [], new_expected)
    first_journal = store.runs / "first-pc/journal.json"
    first = json_read(first_journal)
    first["started_at_utc"] = "2099-01-01T00:00:00+00:00"
    write_json_atomic(first_journal, first)
    store.pointer.unlink()
    store.recover()
    assert store.current()["run_id"] == "second-pc"
