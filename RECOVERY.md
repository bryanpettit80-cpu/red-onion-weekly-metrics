# Recovery Runbook

## Ownership and Storage

- The stable business owner approves recovery access and data use.
- The technical maintainer creates and tests bundles.
- A separately administered Google Drive folder is the independent recovery
  location; it must not be shared broadly or synchronized into the live
  Dropbox operator tree.
- Weekly bundles: retain the newest 13.
- Month-end bundles: retain the newest 12.
- Run and record a restore exercise at least quarterly.

Do not put credentials in a bundle. An operational-data bundle is classified
as Restricted Employee Performance Information.

## Create Recovery Material

Source/release-only:

```powershell
.\Build-RecoveryBundle.ps1 `
  -DestinationDirectory C:\Recovery\RedOnion `
  -RetentionClass Weekly
```

Restricted operational recovery, only to the approved private location:

```powershell
.\Build-RecoveryBundle.ps1 `
  -DestinationDirectory C:\Recovery\RedOnion `
  -RetentionClass Monthly `
  -IncludeOperationalData `
  -OperationsRoot "C:\...\Red Onion Metrics"
```

The restricted bundle includes finished reports, canonical processed reports,
generated-workbook snapshots, integrity manifests, and the machine-local
trusted-head anchor. It excludes the Git checkout because released source and a
Git bundle are already included separately.

## Replacement-Machine Restore

1. Work in an isolated folder, not the live Dropbox path.
2. Verify the bundle against its `.sha256.txt` sidecar, then verify every entry
   in `SHA256SUMS.txt`.
3. Restore `repository.bundle`, check out the recorded release tag, and confirm
   its commit matches `release-metadata.json`.
4. Restore the numbered operator folders to their normal parent layout.
5. Extract the backed-up integrity-anchor folder to a restricted temporary
   location. Do not rename or edit the source anchor JSON.
6. Rebuild the environment and rebind the backed-up head to the replacement
   archive path. Supply the source anchor whose recorded manifest is the
   restored head; the command verifies the full chain and managed outputs before
   creating the new machine-local anchor and audit receipt:

   ```powershell
   .\_program\Run-WeeklySnapshot.ps1 `
     -RebuildEnvironment `
     -RebindRestoredIntegrityAnchor C:\SecureRestore\<old-anchor>.json
   ```

7. Run `.\_program\Run-WeeklySnapshot.ps1 -HealthCheck` and review the result.
   Do not initialize a new baseline to bypass missing history.
8. In an isolated copy, open and inspect the latest master workbook and run the
   full test suite.
9. Obtain business-owner approval before making the restored location live.

If the manifest chain, raw inventory, generated archive, published workbooks,
master digest, or anchor disagrees, stop and follow `INCIDENT_RESPONSE.md`.

## Quarterly Restore Record

Record date, tester, bundle name/hash, release tag/commit, restored manifest
head/hash, health-check result, workbook spot-check result, discrepancies,
corrective actions, and business-owner sign-off. Never record credentials.
