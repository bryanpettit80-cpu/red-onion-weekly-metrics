from __future__ import annotations

from argparse import Namespace
from datetime import date
from pathlib import Path
import shutil

from openpyxl import load_workbook
import pytest

import red_onion_integrity as integrity
import red_onion_weekly_metrics as metrics


ACTIVE_DAY = date(2026, 7, 21)  # Tuesday
ACTIVE_WEEK_END = date(2026, 7, 26)  # Sunday


def workflow_args(tmp_path: Path, *, initialize_baseline: bool = False) -> Namespace:
    return Namespace(
        input_dir=str(tmp_path / "01 Daily Reports - Drop Here"),
        output_dir=str(tmp_path / "02 Finished Reports"),
        archive_dir=str(tmp_path / "03 Archive"),
        config=str(tmp_path / "missing-config.json"),
        week_start=None,
        week_end=None,
        migrate_history_from=[],
        migrate_history_only=False,
        initialize_integrity_baseline=initialize_baseline,
    )


def minimal_records(day: date, source_file: str) -> list[metrics.MetricRecord]:
    common = {
        "source_file": source_file,
        "report_date": day,
        "location": "RC Richmond",
        "gross_sales": 1000.0,
        "guest_count": 50.0,
        "check_average": 20.0,
        "wine_sales": 100.0,
        "wine_pct": 0.10,
        "rate_of_sale_by_guest_count": 0.20,
        "average_ticket_time_seconds": 600.0,
    }
    return [
        metrics.MetricRecord(
            raw_user_name="",
            display_name="",
            is_location_total=True,
            **common,
        ),
        metrics.MetricRecord(
            raw_user_name="Alex Server",
            display_name="Alex Server",
            is_location_total=False,
            **common,
        ),
    ]


@pytest.fixture(scope="module")
def valid_master_template(tmp_path_factory: pytest.TempPathFactory) -> Path:
    root = tmp_path_factory.mktemp("workflow-integrity-master")
    path = root / "Red_Onion_Server_Master.xlsx"
    metrics.write_master_workbook(
        minimal_records(ACTIVE_DAY, "Daily Report template.xlsx"),
        path,
        metrics.load_config(root / "missing-config.json"),
        root / "source",
        ACTIVE_DAY,
        ACTIVE_DAY,
    )
    assert metrics.verify_existing_management_workbook_integrity(path)
    return path


def manifest_paths(archive_dir: Path) -> list[Path]:
    root = metrics.integrity_manifest_dir(archive_dir)
    return sorted(root.glob("*.json"), key=lambda path: path.name.casefold())


def manifest_by_kind(archive_dir: Path, kind: str) -> Path:
    matches = [
        path
        for path in manifest_paths(archive_dir)
        if integrity.read_json_manifest(path).get("kind") == kind
    ]
    assert len(matches) == 1
    return matches[0]


def install_synthetic_parser(monkeypatch: pytest.MonkeyPatch) -> None:
    def parse(path: Path, config: dict) -> list[metrics.MetricRecord]:
        return minimal_records(ACTIVE_DAY, path.name)

    monkeypatch.setattr(metrics, "parse_daily_report", parse)


def test_explicit_baseline_succeeds_without_active_input_and_is_idempotent(
    tmp_path: Path,
) -> None:
    args = workflow_args(tmp_path, initialize_baseline=True)
    archive_dir = Path(args.archive_dir)

    first = metrics.run(args)

    assert len(first) == 1
    baseline = first[0]
    assert baseline.parent == metrics.integrity_manifest_dir(archive_dir)
    payload = integrity.read_json_manifest(baseline)
    assert payload["schema_version"] == 1
    assert payload["kind"] == "integrity-baseline"
    assert payload["raw_inventory"] == []
    assert payload["derived_archive_inventory"] == []
    assert payload["published_output_inventory"] == []
    assert payload["master_generated_content_sha256"] is None
    assert "previous_manifest" not in payload
    assert [entry.path for entry in integrity.verify_manifest_chain(baseline, baseline.parent)] == [
        baseline.resolve()
    ]
    assert not Path(args.input_dir).exists()

    second = metrics.run(args)

    assert second == [baseline]
    assert manifest_paths(archive_dir) == [baseline]


def test_successful_staged_run_chains_manifest_snapshots_outputs_and_deletes_source_last(
    tmp_path: Path,
    valid_master_template: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    args = workflow_args(tmp_path)
    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)
    archive_dir = Path(args.archive_dir)
    input_dir.mkdir(parents=True)
    source = input_dir / "Daily Report - TM - 07-21-2026.xlsx"
    source_bytes = b"unaltered-toast-source"
    source.write_bytes(source_bytes)
    install_synthetic_parser(monkeypatch)

    def write_valid_master(
        records: list[metrics.MetricRecord],
        output_path: Path,
        config: dict,
        source_dir: Path,
        public_start: date,
        public_end: date,
    ) -> Path:
        shutil.copy2(valid_master_template, output_path)
        return output_path

    monkeypatch.setattr(metrics, "write_master_workbook", write_valid_master)
    metrics.run(workflow_args(tmp_path, initialize_baseline=True))
    original_delete = metrics.quarantine_and_delete_captured_inputs
    deletion_observation: dict[str, object] = {}

    def assert_delete_is_last(
        captures: list[metrics.CapturedActiveInput],
        copies: list[metrics.VerifiedArchiveCopy],
        managed_input_dir: Path,
        run_id: str,
    ) -> None:
        captures = list(captures)
        copies = list(copies)
        weekly_manifest = manifest_by_kind(archive_dir, "weekly-run")
        weekly_payload = integrity.read_json_manifest(weekly_manifest)
        snapshot = (
            metrics.generated_workbook_archive_dir(archive_dir)
            / f"week-ending-{ACTIVE_WEEK_END.isoformat()}"
            / weekly_payload["run_id"]
        )
        assert source.exists()
        assert copies == [
            metrics.VerifiedArchiveCopy(
                source=source.resolve(),
                destination=(
                    metrics.canonical_daily_archive_dir(archive_dir)
                    / f"week-ending-{ACTIVE_WEEK_END.isoformat()}"
                    / source.name
                ).resolve(),
                created=True,
                sha256=integrity.sha256_file(source),
            )
        ]
        assert copies[0].destination.read_bytes() == source_bytes
        assert (snapshot / "published" / "Red_Onion_Server_Master.xlsx").is_file()
        assert (output_dir / "Red_Onion_Server_Master.xlsx").is_file()
        assert len(list((snapshot / "published").glob("*.xlsx"))) == 2
        assert len(list(output_dir.glob("*.xlsx"))) == 2
        deletion_observation.update(
            source_existed=True,
            manifest=weekly_manifest,
            snapshot=snapshot,
        )
        assert captures[0].fingerprint.sha256 == copies[0].sha256
        original_delete(captures, copies, managed_input_dir, run_id)

    monkeypatch.setattr(
        metrics, "quarantine_and_delete_captured_inputs", assert_delete_is_last
    )

    generated = metrics.run(args)

    assert deletion_observation["source_existed"] is True
    assert not source.exists()
    assert {path.name for path in generated} == {
        f"Check_Wine_RVA{ACTIVE_WEEK_END:%m%d%y}.xlsx",
        "Red_Onion_Server_Master.xlsx",
    }
    raw_root = metrics.canonical_daily_archive_dir(archive_dir)
    archived = raw_root / f"week-ending-{ACTIVE_WEEK_END.isoformat()}" / source.name
    assert archived.read_bytes() == source_bytes
    expected_raw = integrity.fingerprint_file(raw_root, archived).to_dict()

    baseline = manifest_by_kind(archive_dir, "integrity-baseline")
    weekly = manifest_by_kind(archive_dir, "weekly-run")
    chain = integrity.verify_manifest_chain(weekly, weekly.parent)
    assert [entry.path for entry in chain] == [weekly.resolve(), baseline.resolve()]
    weekly_payload = dict(chain[0].payload)
    assert weekly_payload["raw_inventory"] == [expected_raw]
    assert weekly_payload["details"]["archived_destinations"] == [
        f"week-ending-{ACTIVE_WEEK_END.isoformat()}/{source.name}"
    ]
    assert weekly_payload["details"]["active_input_inventory"] == [
        {
            "path": source.name,
            "size": len(source_bytes),
            "sha256": integrity.sha256_file(archived),
        }
    ]
    snapshot = Path(deletion_observation["snapshot"])
    expected_derived = integrity.build_raw_inventory(
        metrics.generated_workbook_archive_dir(archive_dir)
    )
    assert weekly_payload["derived_archive_inventory"] == [
        item.to_dict() for item in expected_derived
    ]
    assert {path.name for path in (snapshot / "published").glob("*.xlsx")} == {
        path.name for path in generated
    }


def test_raw_archive_tampering_fails_preflight_before_writers_or_source_deletion(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    baseline_args = workflow_args(tmp_path, initialize_baseline=True)
    archive_dir = Path(baseline_args.archive_dir)
    raw = (
        metrics.canonical_daily_archive_dir(archive_dir)
        / "week-ending-2026-07-19"
        / "Daily Report archived.xlsx"
    )
    raw.parent.mkdir(parents=True)
    raw.write_bytes(b"original")
    metrics.run(baseline_args)
    baseline = manifest_by_kind(archive_dir, "integrity-baseline")
    raw.write_bytes(b"tampered")  # Same size; only the SHA-256 changes.

    args = workflow_args(tmp_path)
    active = Path(args.input_dir) / "Daily Report active.xlsx"
    active.parent.mkdir(parents=True)
    active_bytes = b"active-source"
    active.write_bytes(active_bytes)
    calls: list[str] = []

    def should_not_run(*args, **kwargs):
        calls.append("unexpected")
        raise AssertionError("Parsing, writing, and deletion must not run after preflight failure")

    monkeypatch.setattr(metrics, "parse_daily_report", should_not_run)
    monkeypatch.setattr(metrics, "write_public_workbook", should_not_run)
    monkeypatch.setattr(metrics, "write_master_workbook", should_not_run)
    monkeypatch.setattr(metrics, "delete_verified_active_sources", should_not_run)

    with pytest.raises(
        integrity.IntegrityError,
        match=(
            r"Canonical raw archive verification failed\.[\s\S]*"
            r"expected\(size=8, sha256=.*\); actual\(size=8, sha256=.*\)"
        ),
    ):
        metrics.run(args)

    assert calls == []
    assert active.read_bytes() == active_bytes
    assert raw.read_bytes() == b"tampered"
    assert manifest_paths(archive_dir) == [baseline]
    assert not metrics.generated_workbook_archive_dir(archive_dir).exists()
    assert not Path(args.output_dir).exists()


def test_master_preflight_allows_management_inputs_but_rejects_generated_tampering(
    tmp_path: Path,
    valid_master_template: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    args = workflow_args(tmp_path)
    output_dir = Path(args.output_dir)
    archive_dir = Path(args.archive_dir)
    output_dir.mkdir(parents=True)
    master = output_dir / "Red_Onion_Server_Master.xlsx"
    shutil.copy2(valid_master_template, master)

    baseline_args = workflow_args(tmp_path, initialize_baseline=True)
    [baseline] = metrics.run(baseline_args)

    workbook = load_workbook(master, data_only=False)
    workbook["Management Setup"]["B6"] = 1234.0
    workbook["Management Setup"]["A21"] = "Pat Manager"
    workbook["Management Setup"]["B21"] = "Yes"
    workbook.save(master)
    workbook.close()

    same_manifest, _, _ = metrics.ensure_integrity_preflight(
        archive_dir,
        output_dir,
        Path(args.config),
        metrics.load_config(Path(args.config)),
    )
    assert same_manifest == baseline
    assert manifest_paths(archive_dir) == [baseline]
    assert metrics.verify_existing_management_workbook_integrity(master)

    workbook = load_workbook(master, data_only=False)
    workbook["Dashboard"]["A1"] = "TAMPERED GENERATED TITLE"
    workbook.save(master)
    workbook.close()

    active = Path(args.input_dir) / "Daily Report active.xlsx"
    active.parent.mkdir(parents=True)
    active_bytes = b"active-source"
    active.write_bytes(active_bytes)
    calls: list[str] = []

    def should_not_run(*args, **kwargs):
        calls.append("unexpected")
        raise AssertionError("Parsing, writing, and deletion must not run after preflight failure")

    monkeypatch.setattr(metrics, "parse_daily_report", should_not_run)
    monkeypatch.setattr(metrics, "write_public_workbook", should_not_run)
    monkeypatch.setattr(metrics, "write_master_workbook", should_not_run)
    monkeypatch.setattr(metrics, "delete_verified_active_sources", should_not_run)

    with pytest.raises(
        integrity.IntegrityError,
        match="Master workbook generated-content verification failed",
    ):
        metrics.run(args)

    assert calls == []
    assert active.read_bytes() == active_bytes
    assert manifest_paths(archive_dir) == [baseline]
    assert not metrics.generated_workbook_archive_dir(archive_dir).exists()
    workbook = load_workbook(master, data_only=False)
    assert workbook["Management Setup"]["B6"].value == 1234.0
    assert workbook["Management Setup"]["A21"].value == "Pat Manager"
    workbook.close()


def test_writer_failure_preserves_prior_outputs_and_active_source_without_weekly_artifacts(
    tmp_path: Path,
    valid_master_template: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    args = workflow_args(tmp_path)
    output_dir = Path(args.output_dir)
    archive_dir = Path(args.archive_dir)
    output_dir.mkdir(parents=True)
    master = output_dir / "Red_Onion_Server_Master.xlsx"
    prior_public = output_dir / f"Check_Wine_RVA{ACTIVE_WEEK_END:%m%d%y}.xlsx"
    shutil.copy2(valid_master_template, master)
    prior_public.write_bytes(b"prior-public-output")
    prior_master_bytes = master.read_bytes()
    prior_public_bytes = prior_public.read_bytes()
    [baseline] = metrics.run(workflow_args(tmp_path, initialize_baseline=True))

    active = Path(args.input_dir) / "Daily Report active.xlsx"
    active.parent.mkdir(parents=True)
    active_bytes = b"active-source"
    active.write_bytes(active_bytes)
    install_synthetic_parser(monkeypatch)

    def write_staged_public(
        location: str,
        selected_records: list[metrics.MetricRecord],
        output_path: Path,
        config: dict,
        public_start: date,
        public_end: date,
    ) -> Path:
        staged = output_path / prior_public.name
        staged.write_bytes(b"new-staged-public")
        return staged

    def fail_master_writer(*args, **kwargs):
        raise RuntimeError("simulated master writer failure")

    deletion_calls: list[object] = []
    monkeypatch.setattr(metrics, "write_public_workbook", write_staged_public)
    monkeypatch.setattr(metrics, "write_master_workbook", fail_master_writer)
    monkeypatch.setattr(
        metrics,
        "delete_verified_active_sources",
        lambda copies: deletion_calls.append(list(copies)),
    )

    with pytest.raises(RuntimeError, match="simulated master writer failure"):
        metrics.run(args)

    assert deletion_calls == []
    assert active.read_bytes() == active_bytes
    assert master.read_bytes() == prior_master_bytes
    assert prior_public.read_bytes() == prior_public_bytes
    assert manifest_paths(archive_dir) == [baseline]
    assert not metrics.generated_workbook_archive_dir(archive_dir).exists()
    assert not metrics.canonical_daily_archive_dir(archive_dir).exists()
    assert list(output_dir.glob(".weekly-run-*")) == []
