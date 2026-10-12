# Set up Red Onion on Windows

Use one Windows computer at a time. Finish the current run, close the workbooks,
and let Dropbox finish syncing before switching computers. Files needed for a
run must be downloaded locally; an unavailable cloud folder is not an empty week.

1. Download the program and the existing reporting workspace, including its
   archive and latest master workbook. Keep the plain operator folders:
   `01 Daily Reports - Drop Here`, `02 Finished Reports`, and `03 Archive`.
   For a new workspace with no reporting history, create those three
   folders once in File Explorer before setup. For an existing workspace,
   download or restore any missing folder. Ordinary runs stop when a required
   folder is unavailable and never replace it with an empty folder.
2. Install Python 3.10, 3.11 or 3.12 with the Python launcher or Python on PATH.
   Python 3.12 is recommended. Windows PowerShell is included with Windows.
3. From the program folder, run this setup/check command:

   ```powershell
   .\_program\Run-WeeklySnapshot.ps1 -RebuildEnvironment -HealthCheck
   ```

   This builds the disposable environment under
   `%LOCALAPPDATA%\RedOnionMetrics\.venv` from the locked requirements and checks
   the configuration and workspace. It does not generate reports, move intake,
   send mail or activate a schedule. Internet access is needed when dependencies
   must be downloaded. Normal runs reuse and verify the installed environment.

Git, a particular branch, a clean checkout and machine-local integrity anchors
are not requirements for ordinary reporting. Excel is useful for reviewing the
workbooks but is not required to calculate or generate them. Red Onion needs no
mail or Google Drive credentials. Keep Dropbox authentication in the Dropbox app.

## Choose the workspace

The program uses an explicit `-OperationsRoot` first. Otherwise, it looks for
all three numbered operator folders beside the program's `_program` folder, then in the
parent folder. A new standalone copy defaults to its own folder. A program-folder
rename does not change these rules. If there are several possible workspaces,
choose one explicitly:

```powershell
.\_program\Run-WeeklySnapshot.ps1 -OperationsRoot "D:\Work\Red Onion Metrics" -HealthCheck
```

The tracked `Run Weekly Snapshot.cmd` works either inside the program folder or
beside the numbered operator folders with one program subfolder. If more than
one subfolder contains a runner, it stops with a clear message instead of choosing
an arbitrary copy. Keep only the desired launcher in the operator-facing root.

## Run and recover

Close the source and finished workbooks in Excel, put daily reports in the intake
folder, and double-click `Run Weekly Snapshot.cmd`. Read the reported outcome and
the current `LAST RUN STATUS.txt` in finished reports. A warning after publication
does not mean the previous report should be rerun or the input should be restored.

For an already authorized noninteractive caller, the same command accepts:

```powershell
& '.\Run Weekly Snapshot.cmd' -NonInteractive
```

This only removes the final keyboard pause; it does not install or activate a
schedule. The PowerShell runner itself never pauses for input. Use its process
exit code and recorded run outcome together.

Conflicting corrected reports pause with `needs_review`. Review them explicitly
with `-ReviewCorrections`; this asks for one choice per conflicting business date.
An unattended caller can use a previously reviewed `-CorrectionsFile` instead.
`-ReviewCorrections` cannot be combined with `-NonInteractive`. When there are
ambiguous editable master copies, `-ManagementWorkbook` selects the intended
workbook. Ordinary complete, unambiguous weeks need none of these options.

If the local Python environment is damaged, rerun the setup/check command. There
is no need to copy a virtual environment or recreate a computer-specific trust
history. Keep a separate backup of the operational folders and verify restoration
on a copy. Historical forensic integrity/baseline/rebind commands are maintenance
operations, not normal setup steps.
