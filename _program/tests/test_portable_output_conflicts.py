"""Preserve external outputs and reject reserved master filenames before writing."""
import pytest

import red_onion_weekly_metrics as weekly
import red_onion_workflow as workflow
from red_onion_transactions import RecoveryRequired
from test_portable_workflow import (
    arguments, output_hashes, published_workspace, published_copy, store_for,
)


def test_public_report_changed_during_build_is_preserved(published_copy, monkeypatch):
    root = published_copy
    previous = store_for(root).current()
    destination = root / "Finished Reports/Check_Wine_RVA101126.xlsx"
    master = root / "Finished Reports/Red_Onion_Server_Master.xlsx"
    master_before = master.read_bytes()
    replacement = b"a newer externally saved public report"
    real_writer = weekly.write_public_workbook

    def replace_during_generation(*args, **kwargs):
        result = real_writer(*args, **kwargs)
        destination.write_bytes(replacement)
        return result

    monkeypatch.setattr(weekly, "write_public_workbook", replace_during_generation)
    with pytest.raises(RecoveryRequired, match="changed while the report was being built"):
        workflow.run(arguments(root, rebuild_from_history=True))

    assert destination.read_bytes() == replacement
    assert master.read_bytes() == master_before
    assert store_for(root).current() == previous


@pytest.mark.parametrize("master_name", ["Check_Wine_RVA101126.xlsx", "check_wine_rva101826.xlsx"])
def test_master_cannot_use_current_or_future_restaurant_report_filename(
    published_copy, master_name, monkeypatch,
):
    root = published_copy
    original = root / "Finished Reports/Red_Onion_Server_Master.xlsx"
    renamed = root / "Finished Reports" / master_name
    if renamed.exists():
        renamed.unlink()
    original.rename(renamed)
    before = output_hashes(root)
    previous = store_for(root).current()

    def unexpected_generation(*args, **kwargs):
        pytest.fail("A filename conflict must be reported before generation.")

    monkeypatch.setattr(weekly, "write_public_workbook", unexpected_generation)
    with pytest.raises(workflow.WorkflowAttention, match="name conflicts"):
        workflow.run(arguments(root, rebuild_from_history=True))

    assert output_hashes(root) == before
    assert store_for(root).current() == previous
