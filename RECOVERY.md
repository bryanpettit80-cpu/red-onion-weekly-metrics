# Recovery runbook

## Portable workspace backup

Keep an independently retained copy of the **whole operations root**, including
the program source/configuration, launcher, all three numbered folders, and the
latest editable master workbook. Finish reporting, close Excel, and let Dropbox
sync before taking the copy. A sync replica alone is not an independent backup.

`03 Archive\Reporting State` contains accepted daily facts, saved management
state, source copies, publication receipts, journals, and previous-good
outputs. Copy it together with the rest of the workspace. The latest master
may contain management edits newer than the last run and is also needed.

Store operational copies in the approved restricted backup location, outside
the live synchronized tree. They contain Restricted Employee Performance
Information. Retain the newest 13 weekly and 12 month-end recovery copies
under the existing retention policy. Record and exercise a restore quarterly.

Do not copy the local Python virtual environment, credentials, or application
sign-in state. They are not reporting history. The Python environment is
rebuilt on the replacement computer.

`Build-RecoveryBundle.ps1` remains useful for a source/release bundle.
Its `-IncludeOperationalData` option is a **legacy bundle format** that copies
the historical manifest-based layout and machine anchors. It does not include
the new Reporting State tree and is not a complete portable-workspace backup.

## Restore or move to another Windows computer

1. Restore the complete operations-root copy to an isolated local folder.
   Keep the numbered folders together; the program folder may be renamed.
2. Verify the copy against the backup inventory/checksums if provided, and
   confirm that the latest master, source reports, and Reporting State are
   present and downloaded.
3. Install a supported Python version and follow [SETUP.md](SETUP.md).
   From the program folder, run:

   ```powershell
   .\_program\Run-WeeklySnapshot.ps1 -OperationsRoot "D:\Restore\Red Onion Metrics" -RebuildEnvironment -HealthCheck
   ```

4. Inspect the latest workbook's management inputs and run receipt. Health is
   local configuration/storage/state verification; it does not prove Dropbox
   sync or complete incoming data.
5. If a journal is pending, an ordinary invocation in this isolated copy
   reconciles it first. Do not process a live workspace while another computer
   is running or while the restore is incomplete.
6. Verify the reporting behavior on the copy using a known input set. Check
   business totals, date coverage, management carry-forward, and output
   readability. Record any discrepancy before returning the restored folder
   to normal use.
7. Use only one live copy. Before switching PCs, finish the run, close the
   workbook, and wait for Dropbox to finish uploading on one and downloading
   on the other.

No Git checkout, old generated-workbook identity, baseline initialization, or
machine-anchor rebind is required for portable reporting. Folder-copy tests
support the design; they do not establish that a physical second Windows
computer and real Dropbox sync have been tested.

## Interrupted runs and ordinary failures

Publication is recoverable, not a distributed or multi-file atomic transaction.
Until commit, previous-good reports and a journal support rollback. After
commit, cleanup failures remain publication warnings.

- Before publication, invalid or missing inputs leave prior reports in place.
- An interrupted uncommitted publication is rolled back during recovery.
- A committed run resumes safe input/temporary-file cleanup on the next
  invocation without regenerating the accepted report.
- A changed input is preserved. A conflicting changed output during recovery
  stops rather than overwriting newer external work.
- If the master is absent, the saved management snapshot can regenerate it.
  Recover the latest workbook separately when it contains more recent edits.
- If current state is damaged, preserve it and restore the matching saved
  receipt/run record from backup. Do not edit JSON hashes or delete journals
  to force progress.

A missing current pointer can be reconstructed by the recovery code from a
completed journal. An unreadable or inconsistent run record needs explicit
reconciliation. See [INCIDENT_RESPONSE.md](INCIDENT_RESPONSE.md).

## Legacy archive investigation

Keep a preserved pre-migration copy when historical manifest-based evidence
needs examination. The old operational bundle includes its exact manifest
chain, raw/generated inventories, finished outputs, and machine-local anchor.

Verify the bundle and its entry checksums, restore those files to an isolated
legacy workspace, and rebuild the local Python environment. For that legacy
copy only, `-RebindRestoredIntegrityAnchor <saved-anchor.json>` verifies the
backed-up head before creating a path-bound anchor and receipt. The legacy
Python health check requires `--legacy-integrity --health-check` and explicit
input/output/archive paths. Do not run legacy writers against portable state.

## Restore record

Record date, tester, backup identity/hash or inventory, source release,
restored portable run/publication ID, health result, tested input set, business
totals, management-input check, discrepancies, and the decision to resume.
For a real second-PC test also record both machines and observed sync
completion. A source-only bundle or untested copy is not verified operational
recovery.
