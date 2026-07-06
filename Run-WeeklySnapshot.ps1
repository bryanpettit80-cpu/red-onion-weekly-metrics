param(
    [string]$InputDir,
    [string]$OutputDir,
    [string]$ArchiveDir
)

$ErrorActionPreference = "Stop"

$Runner = Join-Path $PSScriptRoot "_program\Run-WeeklySnapshot.ps1"
& $Runner @PSBoundParameters
exit $LASTEXITCODE
