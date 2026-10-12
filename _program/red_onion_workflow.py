"""The ordinary weekly workflow: business validation and recoverable publication.

Legacy integrity manifests remain untouched for historical reference. The
portable current-run receipt contains the accepted daily facts and management
state; no account credentials, machine anchor or Git checkout is required.
"""
from __future__ import annotations

from argparse import Namespace
from dataclasses import asdict
from datetime import date
import math
from pathlib import Path
import re
import shutil
from typing import Any
import uuid

from red_onion_config import load_config
from red_onion_integrity import canonical_json_sha256, sha256_file
from red_onion_management import empty_state, find_master, json_value, read_state
from red_onion_runtime import RunAttemptRecorder, RunStage, utc_now, write_json_atomic
from red_onion_transactions import ReportStore, RecoveryRequired, atomic_bytes, child, json_read


class WorkflowAttention(ValueError):
    def __init__(self, outcome: str, message: str, details: dict[str, Any] | None = None):
        super().__init__(message)
        self.outcome = outcome
        self.details = details or {}


def business_digest(records: list[Any]) -> str:
    import red_onion_weekly_metrics as weekly
    return canonical_json_sha256(weekly.semantic_report_signature(records))


def _encode_records(records: list[Any]) -> list[dict[str, Any]]:
    return [json_value(asdict(record)) for record in records]


def _decode_records(rows: list[dict[str, Any]], day: str) -> list[Any]:
    import red_onion_weekly_metrics as weekly
    records = []
    for row in rows:
        data = dict(row)
        data["report_date"] = date.fromisoformat(data["report_date"])
        if data["report_date"].isoformat() != day:
            raise RecoveryRequired("An accepted daily record has the wrong business date.")
        for key, value in data.items():
            if isinstance(value, float) and not math.isfinite(value):
                raise RecoveryRequired(f"Non-finite {key} in accepted reporting history.")
        records.append(weekly.MetricRecord(**data))
    if not records:
        raise RecoveryRequired(f"Accepted business date {day} has no records.")
    return records


def accepted_days(commit: dict[str, Any] | None, config: dict[str, Any]) -> dict[str, dict[str, Any]]:
    import red_onion_weekly_metrics as weekly
    if commit is None:
        return {}
    days = commit.get("days")
    if not isinstance(days, dict):
        raise RecoveryRequired("The current reporting receipt has no accepted daily history.")
    records = []
    for day, item in days.items():
        date.fromisoformat(day)
        values = _decode_records(item["records"], day)
        if business_digest(values) != item["business_digest"]:
            raise RecoveryRequired(f"Accepted business data changed for {day}. Restore its run record.")
        records.extend(values)
    if records:
        weekly.validate_daily_location_reconciliation(records, config)
    return days


def _capture_reports(paths: list[Path], config: dict[str, Any]) -> tuple[list[Any], dict[Path, list[Any]]]:
    import red_onion_weekly_metrics as weekly
    captures = [weekly.capture_regular_file(path, fingerprint_path=path.name) for path in paths]
    return captures, weekly.read_captured_reports_by_path(captures, config) if captures else {}


def _bootstrap_history(store: ReportStore, config: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], list[Any]]:
    import red_onion_weekly_metrics as weekly
    paths = weekly.archived_daily_report_paths(store.archive_dir)
    captures, reports = _capture_reports(paths, config)
    hashes = {capture.source: capture.fingerprint.sha256 for capture in captures}
    if not reports:
        return {}, captures
    resolution = weekly.resolve_report_duplicates(reports)
    records = weekly.flatten_report_records(resolution.records_by_path)
    weekly.validate_daily_location_reconciliation(records, config)
    return {
        weekly.report_date_for_records(path, rows).isoformat(): {
            "business_digest": business_digest(rows), "records": _encode_records(rows),
            "source": path.relative_to(store.archive_dir).as_posix(),
            "source_sha256": hashes[path], "revision": 1,
        }
        for path, rows in resolution.records_by_path.items()
    }, captures


def _business_changes(before: list[Any], after: list[Any]) -> list[dict[str, Any]]:
    fields = (
        "gross_sales", "guest_count", "wine_sales", "check_count", "check_count_available",
        "rate_of_sale_by_guest_count", "rate_available", "average_ticket_time_seconds",
        "ticket_time_available",
    )
    def key(row: Any) -> tuple[str, str, bool]:
        return row.location, row.raw_user_name, row.is_location_total
    old, new = {key(row): row for row in before}, {key(row): row for row in after}
    changes = []
    for identity in sorted(old.keys() | new.keys()):
        left, right = old.get(identity), new.get(identity)
        for field in fields:
            a, b = getattr(left, field, None), getattr(right, field, None)
            if a != b:
                changes.append({
                    "location": identity[0], "person": "Restaurant total" if identity[2] else identity[1],
                    "field": field, "before": a, "after": b,
                })
    return changes


def _resolve_incoming(
    days: dict[str, dict[str, Any]], incoming: dict[Path, list[Any]],
    store: ReportStore, args: Namespace,
) -> tuple[dict[str, dict[str, Any]], dict[str, Path], list[dict[str, Any]]]:
    """Merge duplicates; a correction is one explicit, fact-bound choice."""
    import red_onion_weekly_metrics as weekly
    grouped: dict[str, dict[Path, list[Any]]] = {}
    for path, rows in incoming.items():
        day = weekly.report_date_for_records(path, rows).isoformat()
        grouped.setdefault(day, {})[path] = rows
    merged = dict(days)
    selected_paths: dict[str, Path] = {}
    conflicts: list[dict[str, Any]] = []
    for day, reports in sorted(grouped.items()):
        prior = days.get(day)
        old = _decode_records(prior["records"], day) if prior else []
        old_path = Path("__accepted__") / (day + ".xlsx")
        candidates = {**({old_path: old} if prior else {}), **reports}
        try:
            resolution = weekly.resolve_report_duplicates(candidates)
            path, rows = next(iter(resolution.records_by_path.items()))
            digest = business_digest(rows)
            if prior and digest == prior["business_digest"]:
                continue
            if path == old_path:
                continue
            merged[day] = {
                "business_digest": digest, "records": _encode_records(rows),
                "source": "", "source_sha256": rows[0].source_sha256,
                "revision": int(prior.get("revision", 0)) + 1 if prior else 1,
            }
            selected_paths[day] = path
        except ValueError:
            options: dict[str, dict[str, Any]] = {}
            if prior:
                options[prior["business_digest"]] = {
                    "digest": prior["business_digest"], "file": "Keep the accepted report", "changes": [],
                }
            for path, rows in reports.items():
                digest = business_digest(rows)
                options.setdefault(digest, {"digest": digest, "file": path.name, "changes": _business_changes(old, rows)})
            conflicts.append({
                "business_date": day, "accepted_digest": prior["business_digest"] if prior else None,
                "options": list(options.values()), "selected_digest": None,
            })
    if not conflicts:
        return merged, selected_paths, []
    signature = canonical_json_sha256([
        {"business_date": c["business_date"], "accepted_digest": c["accepted_digest"],
         "options": sorted(o["digest"] for o in c["options"])} for c in conflicts
    ])
    review_root = child(store.root, "corrections")
    review_path = child(review_root, signature + ".json")
    review = {"schema_version": 1, "review_id": signature, "corrections": conflicts}
    supplied = getattr(args, "corrections_file", None)
    prior_choice = json_read(Path(supplied)) if supplied else json_read(review_path) if review_path.exists() else None
    if prior_choice is not None and prior_choice.get("review_id") == signature:
        choices = {c["business_date"]: c.get("selected_digest") for c in prior_choice.get("corrections", [])}
        for item in conflicts:
            item["selected_digest"] = choices.get(item["business_date"])
    elif supplied:
        raise WorkflowAttention("needs_review", "The correction review is stale. Review the current business-value changes again.")
    if getattr(args, "review_corrections", False):
        for item in conflicts:
            if item["selected_digest"] in {o["digest"] for o in item["options"]}:
                continue
            print(f"\nCorrection for {item['business_date']}:")
            for index, option in enumerate(item["options"], 1):
                print(f"  {index}. {option['file']}")
                for change in option["changes"]:
                    print(f"     {change['location']} / {change['person']} / {change['field']}: {change['before']} -> {change['after']}")
            print("Choose one report number, or Enter to leave for review:", flush=True)
            answer = input().strip()
            if answer.isdigit() and 1 <= int(answer) <= len(item["options"]):
                item["selected_digest"] = item["options"][int(answer) - 1]["digest"]
    write_json_atomic(review_path, review)
    unresolved = [c for c in conflicts if c["selected_digest"] not in {o["digest"] for o in c["options"]}]
    if unresolved:
        summary = "; ".join(
            c["business_date"] + ": " + ", ".join(o["file"] for o in c["options"]) for c in unresolved
        )
        raise WorkflowAttention(
            "needs_review", "Conflicting business values require one correction choice. " + summary
            + ". Run with -ReviewCorrections to see the changes and choose. Previous reports are unchanged.",
            {"correction_review": str(review_path), "review_id": signature},
        )
    receipts = []
    for item in conflicts:
        day, choice = item["business_date"], item["selected_digest"]
        if choice != item["accepted_digest"]:
            path, rows = next((p, r) for p, r in grouped[day].items() if business_digest(r) == choice)
            merged[day] = {
                "business_digest": choice, "records": _encode_records(rows), "source": "",
                "source_sha256": rows[0].source_sha256,
                "revision": int(days[day].get("revision", 0)) + 1 if day in days else 1,
            }
            selected_paths[day] = path
        receipts.append({
            "business_date": day, "previous_digest": item["accepted_digest"],
            "selected_digest": choice, "review_id": signature, "reviewed_at_utc": utc_now(),
        })
    return merged, selected_paths, receipts


def _all_records(days: dict[str, dict[str, Any]]) -> list[Any]:
    return [row for day, item in sorted(days.items()) for row in _decode_records(item["records"], day)]


def _complete_weeks(records: list[Any], config: dict[str, Any], *, through: date | None = None) -> list[date]:
    import red_onion_weekly_metrics as weekly
    candidates = {weekly.week_period_for(row.report_date)[1] for row in records if weekly.is_operating_day(row.report_date)}
    candidates.update(
        weekly.week_period_for(date.fromisoformat(day))[1]
        for dates in config.get("closed_dates", {}).values() for day in dates
    )
    complete = []
    for end in sorted(candidates):
        if through is not None and end > through:
            continue
        start, _ = weekly.week_period_for(end)
        try:
            weekly.validate_complete_reporting_set(records, config, start, end)
        except ValueError:
            continue
        complete.append(end)
    return complete


def _result(
    recorder: RunAttemptRecorder, outcome: str, message: str, *, details: dict[str, Any] | None = None,
) -> None:
    """Post-commit status writes are advisory; the durable receipt is authoritative."""
    recorder.outcome = outcome
    recorder.stage = RunStage.COMPLETE if outcome in {"published", "published_with_warnings", "already_processed", "history_imported"} else RunStage.FAILED
    recorder.completed_at_utc = utc_now()
    recorder.message = message
    recorder.details.update(details or {})
    if outcome in {"published", "published_with_warnings"}:
        recorder.readiness.update({"integrity": "Ready", "workbook": "Ready", "distribution": "Ready"})
    elif outcome in {"waiting_for_inputs", "needs_review"}:
        recorder.readiness["integrity"] = "Attention"
    try:
        recorder.write()
    except OSError as exc:
        if outcome in {"published", "published_with_warnings"}:
            recorder.outcome = "published_with_warnings"
            recorder.details.setdefault("warnings", []).append(f"Status refresh failed after publication: {exc}")
            print(f"Publication succeeded; status refresh warning: {exc}")
        else:
            print(f"Status could not be saved: {exc}")


def run(args: Namespace) -> list[Path]:
    import red_onion_weekly_metrics as weekly
    config = load_config(Path(args.config).resolve())
    store = ReportStore(Path(args.archive_dir), Path(args.output_dir), Path(args.input_dir))
    for label, path in (("Daily Reports", store.input_dir), ("Finished Reports", store.output_dir), ("Archive", store.archive_dir)):
        if not path.is_dir():
            raise WorkflowAttention("waiting_for_inputs", f"{label} folder unavailable: {path}. Download the workspace or run setup; no reports were changed.")
    if len({store.input_dir, store.output_dir, store.archive_dir}) != 3 or any(
        a.is_relative_to(b) for a in (store.input_dir, store.output_dir, store.archive_dir)
        for b in (store.input_dir, store.output_dir, store.archive_dir) if a != b
    ):
        raise ValueError("Input, output and archive folders must be separate, non-nested directories.")
    run_id = str(uuid.uuid4())
    with weekly.workflow_run_lock(store.archive_dir):
        recorder = RunAttemptRecorder(
            run_id=run_id, operation="history-rebuild" if getattr(args, "rebuild_from_history", False) else "weekly-run",
            attempt_path=store.root / "attempts" / (run_id + ".json"),
            status_path=store.output_dir / weekly.LAST_RUN_STATUS_FILE,
        )
        recorder.details["workflow"] = "portable-v1"
        args._run_attempt_recorder = recorder
        try:
            warnings = store.recover()
            previous = store.current()
            days, bootstrap_captures = (accepted_days(previous, config), []) if previous else _bootstrap_history(store, config)
            recorder.update(RunStage.READING_INPUTS, "Reading accepted history and incoming reports.")
            return _run_locked(args, config, store, recorder, previous, days, warnings, bootstrap_captures)
        except WorkflowAttention as exc:
            _result(recorder, exc.outcome, str(exc), details=exc.details)
            if exc.outcome == "waiting_for_inputs":
                print(str(exc))
                return []
            raise
        except Exception as exc:
            # A receipt advanced before a late status or filesystem error remains
            # published. A retry must never rebuild it as a failed first attempt.
            try:
                current = store.current()
            except Exception:
                current = None
            if current is not None and current["run_id"] == run_id:
                paths = [child(store.output_dir, item["name"]) for item in current["outputs"]]
                outcome = {"history_import": "history_imported", "duplicate_receipt": "already_processed"}.get(current.get("operation"), "published_with_warnings")
                _result(recorder, outcome, "The reporting receipt committed; follow-up is needed.", details={"warnings": [str(exc)], "publication_id": current.get("report_publication_id")})
                return paths
            outcome = "needs_review" if isinstance(exc, (RecoveryRequired, ValueError)) else "failed_before_publication"
            _result(recorder, outcome, str(exc))
            raise


def _run_locked(
    args: Namespace, config: dict[str, Any], store: ReportStore, recorder: RunAttemptRecorder,
    previous: dict[str, Any] | None, days: dict[str, dict[str, Any]], warnings: list[str],
    bootstrap_captures: list[Any],
) -> list[Path]:
    import red_onion_weekly_metrics as weekly
    rebuild = bool(getattr(args, "rebuild_from_history", False))
    migration_only = bool(getattr(args, "migrate_history_only", False))
    migration_roots = [Path(p).resolve() for p in (getattr(args, "migrate_history_from", None) or [])]
    if migration_only and not migration_roots:
        raise ValueError("History import requires at least one --migrate-history-from folder.")
    active_paths = [] if rebuild or migration_only else weekly.daily_report_paths(store.input_dir)
    active_captures, incoming = _capture_reports(active_paths, config)
    migration_paths: list[Path] = []
    for root in migration_roots:
        if not root.is_dir():
            raise ValueError(f"History import folder unavailable: {root}")
        migration_paths.extend(weekly.find_daily_report_paths(root, recursive=True))
    migration_captures, imported = _capture_reports(migration_paths, config)
    if set(incoming) & set(imported):
        raise ValueError("The active drop folder cannot also be a history import source.")
    incoming.update(imported)
    for path, rows in incoming.items():
        weekly.validate_daily_location_reconciliation(rows, config)
        if any(not weekly.is_operating_day(row.report_date) for row in rows):
            raise ValueError(f"{path.name} is not a Tuesday-Sunday business-day report.")
    explicit_period = bool(args.week_start or args.week_end)
    bootstrap_only = bool(
        previous is None and days and not incoming and not rebuild
        and not migration_only and not migration_roots and not explicit_period
    )
    # The first ordinary legacy run can save canonical history and current
    # management without republishing an old week. An explicit import-only
    # command still requires its source folder, checked above.
    migration_only = migration_only or bootstrap_only
    pending_refresh = bool((previous or {}).get("pending_report_refresh"))
    pending_weeks = {date.fromisoformat(value) for value in (previous or {}).get("pending_report_weeks", [])}
    if not incoming and not rebuild and not explicit_period and not pending_refresh and not bootstrap_only:
        raise WorkflowAttention("waiting_for_inputs", "No new daily reports are available. The previous report is unchanged.")
    if active_paths:
        # A wrong-period attachment must not silently select a different week.
        weekly.active_week_for_paths({p: incoming[p] for p in active_paths})
    updated, selected, correction_receipts = _resolve_incoming(days, incoming, store, args)
    records = _all_records(updated)
    weekly.validate_daily_location_reconciliation(records, config)
    if records:
        weekly.validated_data_quality_coverage(records, min(r.report_date for r in records), max(r.report_date for r in records))
    affected_weeks = {weekly.week_period_for(date.fromisoformat(day))[1] for day in selected}
    if not migration_only:
        chosen_period = None
        if explicit_period:
            chosen_period = weekly.selected_public_dates(records, args.week_start, args.week_end)
            if chosen_period != weekly.week_period_for(chosen_period[1]):
                raise ValueError("Public reporting requires a complete Tuesday-Sunday period; a partial date range cannot be published.")
            weekly.validate_complete_reporting_set(records, config, *chosen_period)
        # Future configured closures do not advance the report by themselves.
        # An observed, previously published, or explicitly selected period sets
        # the horizon; closures only satisfy coverage within that horizon.
        horizon_candidates = [weekly.week_period_for(r.report_date)[1] for r in records]
        if chosen_period:
            horizon_candidates.append(chosen_period[1])
        if (previous or {}).get("period"):
            horizon_candidates.append(date.fromisoformat(previous["period"]["end"]))
        for end in sorted(affected_weeks):
            start, _ = weekly.week_period_for(end)
            try:
                weekly.validate_complete_reporting_set(records, config, start, end)
            except ValueError as exc:
                raise WorkflowAttention("waiting_for_inputs", str(exc)) from exc
        complete_weeks = _complete_weeks(records, config, through=max(horizon_candidates)) if horizon_candidates else []
        if not complete_weeks:
            raise WorkflowAttention("waiting_for_inputs", "No complete Tuesday-Sunday reporting set is available for both restaurants.")
        latest_end = max(complete_weeks)
        latest_start, _ = weekly.week_period_for(latest_end)
        # A copy-only import can accept historical facts without publication.
        # Refresh complete affected periods on the next ordinary invocation;
        # incomplete historical periods remain visible in master detail only.
        affected_weeks.update(pending_weeks.intersection(complete_weeks))
        if chosen_period:
            affected_weeks.add(chosen_period[1])
        if rebuild or previous is None:
            affected_weeks.add(latest_end)
    else:
        latest_start = latest_end = None
    expected_master = (
        (previous or {}).get("published_outputs", {}).get("master", {}).get("name")
        or (previous or {}).get("management_workbook_name")
    )
    current_master = find_master(store.output_dir, getattr(args, "management_workbook", None), expected_name=expected_master)
    old_management = previous.get("management_state") if previous else None
    current_master_hash = sha256_file(current_master) if current_master else None
    management = read_state(current_master, old_management) if current_master else old_management or empty_state()
    if current_master is not None and sha256_file(current_master) != current_master_hash:
        raise RecoveryRequired("Management inputs changed while being read. Retry after the workbook has finished saving.")
    if old_management is not None and current_master is None and (
        expected_master or (previous or {}).get("report_publication_id")
    ):
        if not rebuild:
            raise WorkflowAttention("needs_review", "The current management workbook is missing. Finish downloading the workspace or use --rebuild-from-history to restore the last saved management state.")
        warnings.append("The master was not present; its last saved management state was restored.")
    if current_master is not None and any(
        re.fullmatch(r"Check_Wine_" + re.escape(info["short_code"]) + r"\d{6}\.xlsx", current_master.name, re.IGNORECASE)
        for info in config["locations"].values()
    ):
        raise WorkflowAttention(
            "needs_review", "The management workbook name conflicts with a restaurant report filename. "
            "Rename the management workbook before retrying; no report was overwritten."
        )
    run_id = recorder.run_id
    run_root = child(store.runs, run_id)
    run_root.mkdir(parents=True, exist_ok=False)
    cleanup_inputs = []
    source_refs: dict[Path, str] = {}
    for index, capture in enumerate([*active_captures, *migration_captures]):
        relative = f"Reporting State/runs/{run_id}/sources/{index:04d}-{capture.source.name}"
        destination = child(store.archive_dir, relative)
        if atomic_bytes(destination, capture.content, exclusive=True) != capture.fingerprint.sha256:
            raise OSError(f"Could not verify saved input {capture.source.name}.")
        source_refs[capture.source] = relative
        if capture in active_captures:
            cleanup_inputs.append({
                "name": capture.source.name, "sha256": capture.fingerprint.sha256, "archive": relative,
                "quarantine": f".report-{run_id}-{index:04d}.input",
            })
    for day, path in selected.items():
        updated[day]["source"] = source_refs[path]
        updated[day]["source_sha256"] = sha256_file(child(store.archive_dir, source_refs[path]))
    config["_report_audit"] = {
        "canonical_archive": str(store.root), "unique_business_days": len(updated),
        "duplicate_files_ignored": max(0, len(incoming) - len(selected)),
        "duplicate_file_names": [], "conflicts": 0,
    }
    config["_portable_management_state"] = management
    staged_paths: list[Path] = []
    output_expected: dict[str, str | None] = {}
    stage = child(run_root, "staging")
    stage.mkdir()
    current_config = {k: v for k, v in config.items() if not k.startswith("_")}
    current_methodology = {"management": weekly.MANAGEMENT_METHODOLOGY_VERSION, "descriptive": weekly.PERFORMANCE_CONSISTENCY_METHODOLOGY_VERSION}
    no_business_change = (
        not selected and previous is not None and not rebuild and not migration_only and not explicit_period and not pending_refresh
        and previous.get("config") == current_config and previous.get("methodology") == current_methodology
        and json_value(management) == previous.get("management_state")
    )
    if no_business_change:
        # Only current-period reports are needed; old derivatives may be removed.
        required = [store.output_dir / f"Check_Wine_{info['short_code']}{latest_end:%m%d%y}.xlsx" for info in config["locations"].values()]
        if any(not path.is_file() for path in required):
            no_business_change = False
    if not no_business_change and not migration_only and not affected_weeks:
        affected_weeks.add(latest_end)
    try:
        if not no_business_change and not migration_only:
            master_name = current_master.name if current_master is not None else "Red_Onion_Server_Master.xlsx"
            planned_names = [
                f"Check_Wine_{info['short_code']}{end:%m%d%y}.xlsx"
                for end in sorted(affected_weeks) for info in config["locations"].values()
            ] + [master_name]
            if len({name.casefold() for name in planned_names}) != len(planned_names):
                raise ValueError("Generated report filenames collide; check the restaurant short codes and master filename.")
            # Record the destination baseline before any generation. A manager
            # save or cloud replacement during the build must not be overwritten.
            output_expected = {
                name: sha256_file(store.output_dir / name) if (store.output_dir / name).exists() else None
                for name in planned_names
            }
            weekly.assert_staging_capacity(store.output_dir, store.archive_dir, tuple(active_captures))
            recorder.update(RunStage.BUILDING_WORKBOOKS, "Building a complete reporting set and preserving management input.")
            for end in sorted(affected_weeks):
                start, _ = weekly.week_period_for(end)
                subset = [r for r in records if start <= r.report_date <= end]
                for location in config["locations"]:
                    staged_paths.append(weekly.write_public_workbook(location, subset, stage, config, start, end))
            staged_master = stage / "Red_Onion_Server_Master.xlsx"
            history_records = [r for r in records if r.report_date <= latest_end]
            staged_paths.append(weekly.write_master_workbook(history_records, staged_master, config, store.input_dir, latest_start, latest_end))
            weekly.validate_staged_outputs(stage, staged_paths)
            management = read_state(staged_master)
            if current_master is not None and current_master.name != staged_master.name:
                renamed = stage / current_master.name
                staged_master.rename(renamed)
                staged_paths = [renamed if p == staged_master else p for p in staged_paths]
        if current_master is not None and sha256_file(current_master) != current_master_hash:
            raise RecoveryRequired("Management inputs changed while the report was building. Rerun to preserve the latest edits.")
        weekly.verify_captured_active_inputs(active_captures, store.input_dir)
        for capture in [*migration_captures, *bootstrap_captures]:
            if sha256_file(capture.source) != capture.fingerprint.sha256:
                raise RecoveryRequired(f"History input changed during import: {capture.source.name}")
        published_outputs = dict(previous.get("published_outputs", {})) if previous else {}
        for path in staged_paths:
            role = "master" if (current_master is not None and path.name == current_master.name) or path.name == "Red_Onion_Server_Master.xlsx" else path.name
            published_outputs[role] = {"name": path.name, "snapshot": f"Reporting State/runs/{run_id}/outputs/{path.name}"}
        commit = {
            "schema_version": 1, "run_id": run_id, "created_at_utc": utc_now(),
            "previous_run_id": previous["run_id"] if previous else None,
            "operation": "history_import" if migration_only else "duplicate_receipt" if no_business_change else "history_rebuild" if rebuild else "weekly_report",
            "days": updated, "management_state": json_value(management),
            "management_workbook_name": current_master.name if current_master is not None else expected_master,
            "corrections": correction_receipts,
            "period": {"start": latest_start.isoformat(), "end": latest_end.isoformat()} if latest_end else (previous or {}).get("period"),
            "changed_weeks": [d.isoformat() for d in sorted(affected_weeks)],
            "config": current_config, "methodology": current_methodology,
            "published_outputs": published_outputs,
            "report_publication_id": (previous or {}).get("report_publication_id") if no_business_change or migration_only else run_id,
            "pending_report_refresh": bool(
                selected or pending_refresh or previous is None
                or previous.get("config") != current_config
                or previous.get("methodology") != current_methodology
                or previous.get("management_state") != json_value(management)
            ) if migration_only else False,
            "pending_report_weeks": [d.isoformat() for d in sorted(affected_weeks | pending_weeks)] if migration_only else [],
        }
        recorder.update(RunStage.PUBLISHING, "Validated reports are being published with a recovery journal.")
        final, cleanup_warnings = store.publish(
            run_id, staged_paths, commit, cleanup_inputs, output_expected,
        )
        warnings.extend(cleanup_warnings)
        outcome = "already_processed" if no_business_change else "history_imported" if migration_only else "published_with_warnings" if warnings else "published"
        message = "Inputs were already processed; no reports were regenerated." if no_business_change else "History imported without generating reports." if migration_only else "Complete reports published successfully."
        _result(recorder, outcome, message, details={
            "publication_id": commit["report_publication_id"], "period": commit["period"], "changed_weeks": commit["changed_weeks"],
            "generated_files": [p.name for p in final], "warnings": warnings,
        })
        print(f"Outcome: {outcome}. {message}")
        for warning in warnings:
            print("Warning: " + warning)
        return final
    finally:
        try:
            shutil.rmtree(stage)
        except OSError:
            # Staging contains no unique accepted data; a transient cleanup lock
            # does not invalidate the committed run or its previous-good copies.
            pass


def health_check(args: Namespace) -> dict[str, Any]:
    """Environment/storage health only, never authorization to process a week."""
    checks = []
    def add(name: str, status: str, message: str) -> None:
        checks.append({"name": name, "status": status, "message": message})
    try:
        config = load_config(Path(args.config).resolve())
        add("configuration", "Ready", f"Configuration valid for {len(config['locations'])} restaurants.")
    except Exception as exc:
        add("configuration", "Attention", str(exc))
        config = None
    store = ReportStore(Path(args.archive_dir), Path(args.output_dir), Path(args.input_dir))
    for name, path in (("input", store.input_dir), ("output", store.output_dir), ("archive", store.archive_dir)):
        add(name, "Ready" if path.is_dir() else "Attention", str(path))
    try:
        commit = store.current()
        if commit and config:
            accepted_days(commit, config)
        add("reporting_state", "Ready", "Portable reporting receipt is readable." if commit else "First portable run will import existing source history without changing old manifests.")
        pending = [p for p in store.runs.glob("*/journal.json") if json_read(p).get("phase") not in {"complete", "rolled_back"}]
        if pending:
            add("recovery", "Attention", "An interrupted run needs recovery on the next ordinary invocation.")
    except Exception as exc:
        add("reporting_state", "Attention", str(exc))
    return {
        "schema_version": 2, "overall": "Ready" if all(c["status"] == "Ready" for c in checks) else "Attention",
        "workflow": "portable-v1", "checks": checks,
        "scope": "Read-only local health. Does not establish incoming-week completeness, cloud sync or recipient access.",
    }


def evidence_source(args: Namespace) -> tuple[dict[str, Any], list[dict[str, Any]], Path]:
    """Existing local evidence export using accepted facts plus current decisions."""
    import red_onion_weekly_metrics as weekly
    store = ReportStore(Path(args.archive_dir), Path(args.output_dir), Path(args.input_dir))
    commit = store.current()
    if commit is None:
        raise ValueError("Publish a reporting set before exporting its management evidence.")
    master = find_master(store.output_dir, getattr(args, "management_workbook", None), expected_name=commit.get("published_outputs", {}).get("master", {}).get("name"))
    if master is None:
        raise ValueError("The editable management workbook is unavailable.")
    before = sha256_file(master)
    state = read_state(master, commit["management_state"])
    if sha256_file(master) != before:
        raise RecoveryRequired("The management workbook changed while its decisions were read.")
    rows = []
    for action in state["active_actions"]:
        if str(action.get("Status") or "").casefold() in {"complete", "dismissed"}:
            continue
        evidence = state["evidence_by_action_id"].get(str(action.get("Action ID")))
        if evidence is None:
            raise ValueError("A saved action has no supporting evidence record.")
        rows.append({**evidence, **action})
    commit_path = child(store.runs, commit["run_id"] + "/commit.json")
    source = {
        "manifest_path": commit_path.relative_to(store.archive_dir).as_posix(),
        "manifest_sha256": sha256_file(commit_path), "manifest_run_id": commit["run_id"],
        "manifest_created_at_utc": commit["created_at_utc"], "workbook_file": master.name,
        "workbook_generated_content_sha256": canonical_json_sha256(json_value(rows)),
        "workbook_generated_content_digest_scheme": "portable-business-v1",
        "generator_commit": None,
        "effective_config_sha256": canonical_json_sha256(commit["config"]),
        "methodology_version": weekly.MANAGEMENT_METHODOLOGY_VERSION,
    }
    return source, rows, master
