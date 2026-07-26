param(
    [Parameter(Mandatory = $true)][string]$DestinationDirectory,
    [ValidateSet("Weekly", "Monthly")][string]$RetentionClass = "Weekly",
    [string]$OperationsRoot,
    [string]$IntegrityAnchorDirectory = (
        Join-Path $env:LOCALAPPDATA "RedOnionMetrics\integrity-anchors"
    ),
    [switch]$IncludeOperationalData
)

$ErrorActionPreference = "Stop"
$RepositoryRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepositoryRoot = [System.IO.Path]::GetFullPath($RepositoryRoot)
$DestinationDirectory = [System.IO.Path]::GetFullPath($DestinationDirectory)
$DestinationRoot = [System.IO.Path]::GetPathRoot($DestinationDirectory)
if ($DestinationDirectory.TrimEnd("\") -eq $DestinationRoot.TrimEnd("\")) {
    throw "Refusing to use a drive root as the recovery-bundle destination."
}
$RepositoryPrefix = $RepositoryRoot.TrimEnd("\") + "\"
if (
    $DestinationDirectory.TrimEnd("\").Equals(
        $RepositoryRoot.TrimEnd("\"), [System.StringComparison]::OrdinalIgnoreCase
    ) -or $DestinationDirectory.StartsWith(
        $RepositoryPrefix, [System.StringComparison]::OrdinalIgnoreCase
    )
) {
    throw "Recovery bundles must be written outside the Git repository."
}

function Invoke-Git {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)

    $Output = @(& git -C $RepositoryRoot @Arguments 2>&1)
    if ($LASTEXITCODE -ne 0) {
        throw "Git failed: git $($Arguments -join ' ')`n$($Output -join "`n")"
    }
    return (($Output | ForEach-Object { $_.ToString() }) -join "`n").Trim()
}

function Get-FileSha256 {
    param([Parameter(Mandatory = $true)][string]$Path)

    $Stream = [System.IO.File]::OpenRead($Path)
    $Hasher = [System.Security.Cryptography.SHA256]::Create()
    try {
        $HashBytes = $Hasher.ComputeHash($Stream)
        return (($HashBytes | ForEach-Object { $_.ToString("x2") }) -join "")
    } finally {
        $Hasher.Dispose()
        $Stream.Dispose()
    }
}

function Get-TextSha256 {
    param([Parameter(Mandatory = $true)][string]$Text)

    $Hasher = [System.Security.Cryptography.SHA256]::Create()
    try {
        $Bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        return (($Hasher.ComputeHash($Bytes) |
            ForEach-Object { $_.ToString("x2") }) -join "")
    } finally {
        $Hasher.Dispose()
    }
}

function Test-SafeCloudReparsePoint {
    param([Parameter(Mandatory = $true)]$Entry)

    if (
        ($Entry.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq 0
    ) {
        return $false
    }

    # PowerShell's LinkType/Target convenience properties are not exhaustive for
    # all reparse providers. Fail closed for recognized links, then allow only
    # Microsoft cloud-placeholder tags (0x9000001A through 0x9000F01A), which do
    # not set the name-surrogate bit used by redirecting links.
    $LinkType = $Entry.PSObject.Properties["LinkType"]
    $Target = $Entry.PSObject.Properties["Target"]
    $HasLinkType = $null -ne $LinkType -and -not [string]::IsNullOrWhiteSpace(
        [string]$LinkType.Value
    )
    $HasTarget = $null -ne $Target -and @($Target.Value).Count -gt 0 -and (
        @($Target.Value) | Where-Object {
            -not [string]::IsNullOrWhiteSpace([string]$_)
        }
    ).Count -gt 0
    if ($HasLinkType -or $HasTarget) {
        return $false
    }

    $Principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    $IsElevated = $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $IsElevated) {
        throw "fsutil reparsepoint query requires an elevated (Administrator) shell. Re-run Build-RecoveryBundle.ps1 from an elevated prompt to validate reparse points."
    }
    $FsutilStderr = $null
    $Query = @(& fsutil reparsepoint query $Entry.FullName 2>&1 | ForEach-Object {
        if ($_ -is [System.Management.Automation.ErrorRecord]) {
            $FsutilStderr = $_.ToString()
        } else {
            $_
        }
    })
    if ($LASTEXITCODE -ne 0) {
        $WarningDetail = if ($FsutilStderr) { ": $FsutilStderr" } else { "" }
        Write-Warning "fsutil reparsepoint query failed for '$($Entry.FullName)'$WarningDetail"
        return $false
    }
    $TagLine = $Query | Where-Object { $_ -match "Reparse Tag Value" } |
        Select-Object -First 1
    $TagMatch = [regex]::Match([string]$TagLine, "0x([0-9A-Fa-f]+)")
    if (-not $TagLine -or -not $TagMatch.Success) {
        return $false
    }
    $Tag = [Convert]::ToUInt32($TagMatch.Groups[1].Value, 16)
    $NameSurrogateBit = [uint32]0x20000000
    if (($Tag -band $NameSurrogateBit) -ne 0) {
        return $false
    }
    return ($Tag -band [uint32]0xFFFF0FFF) -eq [uint32]0x9000001A
}

function Test-UnsafeReparsePoint {
    param([Parameter(Mandatory = $true)]$Entry)

    if (
        ($Entry.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq 0
    ) {
        return $false
    }
    return -not (Test-SafeCloudReparsePoint -Entry $Entry)
}

function Assert-NormalTree {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "$Label is missing: $Path"
    }
    $RootEntry = Get-Item -LiteralPath $Path -Force
    if (Test-UnsafeReparsePoint -Entry $RootEntry) {
        throw "$Label is a link or reparse point: $Path"
    }
    $Entries = @(Get-ChildItem -LiteralPath $Path -Recurse -Force)
    foreach ($Entry in $Entries) {
        if (Test-UnsafeReparsePoint -Entry $Entry) {
            throw "$Label contains a link or reparse point: $($Entry.FullName)"
        }
    }
}

function Get-TreeFingerprint {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Label
    )

    Assert-NormalTree -Path $Path -Label $Label
    $ResolvedRoot = [System.IO.Path]::GetFullPath($Path).TrimEnd("\")
    $Lines = @()
    foreach ($File in Get-ChildItem -LiteralPath $ResolvedRoot -Recurse -File -Force |
        Where-Object { $_.Name -ne ".weekly-snapshot.lock" } |
        Sort-Object FullName) {
        $Relative = $File.FullName.Substring($ResolvedRoot.Length).TrimStart("\")
        $Lines += (
            "$($Relative -replace '\\','/')|$($File.Length)|" +
            "$(Get-FileSha256 -Path $File.FullName)"
        )
    }
    return Get-TextSha256 -Text (($Lines -join "`n") + "`n")
}

function Get-OperationalSourceFingerprint {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$AnchorDirectory
    )

    $Parts = @()
    foreach ($RelativePath in @(
        "02 Finished Reports",
        "03 Archive\processed-daily-reports",
        "03 Archive\generated-workbooks",
        "03 Archive\run-manifests"
    )) {
        $Source = Join-Path $Root $RelativePath
        $Parts += "$RelativePath=$(Get-TreeFingerprint -Path $Source -Label $RelativePath)"
    }
    foreach ($RelativePath in @(
        "03 Archive\run-attempts",
        "03 Archive\approved-management-evidence"
    )) {
        $Source = Join-Path $Root $RelativePath
        if (Test-Path -LiteralPath $Source) {
            $Parts += "$RelativePath=$(Get-TreeFingerprint -Path $Source -Label $RelativePath)"
        } else {
            $Parts += "$RelativePath=ABSENT"
        }
    }
    $Parts += (
        "integrity-anchors=" +
        (Get-TreeFingerprint -Path $AnchorDirectory -Label "Integrity anchor")
    )
    return Get-TextSha256 -Text (($Parts -join "`n") + "`n")
}

function Copy-NormalTree {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination
    )

    New-Item -ItemType Directory -Path $Destination | Out-Null
    foreach ($Entry in Get-ChildItem -LiteralPath $Source -Force) {
        if ($Entry.Name -eq ".weekly-snapshot.lock") {
            continue
        }
        Copy-Item -LiteralPath $Entry.FullName -Destination $Destination -Recurse
    }
}

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    throw "Git is required to build a recovery bundle."
}
$Status = Invoke-Git @("status", "--porcelain=v1", "--untracked-files=all")
if ($Status) {
    throw "Recovery bundles require a clean released checkout. Uncommitted state: $Status"
}
$Branch = Invoke-Git @("symbolic-ref", "--quiet", "--short", "HEAD")
if ($Branch -ne "main") {
    throw "Recovery bundles must be built from branch main; current branch is $Branch."
}
$Commit = Invoke-Git @("rev-parse", "--verify", "HEAD^{commit}")
$OriginMain = Invoke-Git @(
    "rev-parse", "--verify", "refs/remotes/origin/main^{commit}"
)
if ($Commit -ne $OriginMain) {
    throw "HEAD does not match the verified local origin/main release reference."
}
$Tags = @(
    (Invoke-Git @("tag", "--points-at", "HEAD", "--list", "v*")) -split "`n" |
        Where-Object { $_ }
)
if ($Tags.Count -eq 0) {
    throw "The current commit has no v* release tag."
}
$ReleaseTag = $Tags | Sort-Object | Select-Object -Last 1

if ($IncludeOperationalData) {
    if (-not $OperationsRoot) {
        throw "-IncludeOperationalData requires -OperationsRoot."
    }
    $OperationsRoot = [System.IO.Path]::GetFullPath($OperationsRoot)
    if ($OperationsRoot -eq $RepositoryRoot) {
        throw "Operational data root must be the operator workspace, not the Git repository."
    }
}

New-Item -ItemType Directory -Force -Path $DestinationDirectory | Out-Null
$TemporaryRoot = Join-Path (
    [System.IO.Path]::GetTempPath()
) ("RedOnion-Recovery-" + [guid]::NewGuid().ToString("N"))
$TemporaryRoot = [System.IO.Path]::GetFullPath($TemporaryRoot)
$ExpectedTemporaryPrefix = [System.IO.Path]::GetFullPath(
    [System.IO.Path]::GetTempPath()
).TrimEnd("\") + "\"
if (-not $TemporaryRoot.StartsWith(
    $ExpectedTemporaryPrefix, [System.StringComparison]::OrdinalIgnoreCase
)) {
    throw "Refusing an unsafe recovery staging path: $TemporaryRoot"
}

$Timestamp = [DateTime]::UtcNow.ToString("yyyyMMddTHHmmssZ")
$SafeTag = $ReleaseTag -replace "[^A-Za-z0-9._-]", "-"
$BundleBase = "Red-Onion-Recovery-$SafeTag-$Timestamp-$($RetentionClass.ToLowerInvariant())"
$FinalBundle = Join-Path $DestinationDirectory "$BundleBase.zip"
$FinalSidecar = "$FinalBundle.sha256.txt"
if (
    (Test-Path -LiteralPath $FinalBundle) `
    -or (Test-Path -LiteralPath $FinalSidecar)
) {
    throw "A recovery artifact with this timestamp already exists."
}

$OperationalLockStream = $null
$OperationalLockAcquired = $false
try {
    New-Item -ItemType Directory -Path $TemporaryRoot | Out-Null
    $PayloadRoot = Join-Path $TemporaryRoot "payload"
    New-Item -ItemType Directory -Path $PayloadRoot | Out-Null
    $SourceArchive = Join-Path $PayloadRoot "released-source.zip"
    & git -C $RepositoryRoot archive --format=zip --output=$SourceArchive $Commit
    if ($LASTEXITCODE -ne 0) {
        throw "Could not create the released-source archive."
    }
    $GitBundle = Join-Path $PayloadRoot "repository.bundle"
    & git -C $RepositoryRoot bundle create $GitBundle --all
    if ($LASTEXITCODE -ne 0) {
        throw "Could not create the offline Git recovery bundle."
    }

    $MetadataRoot = Join-Path $PayloadRoot "release-metadata"
    New-Item -ItemType Directory -Path $MetadataRoot | Out-Null
    foreach ($RelativePath in @(
        "_program\red_onion_config.json",
        "_program\requirements.txt",
        "_program\requirements.lock",
        "_program\requirements-constraints.txt",
        "_program\pyproject.toml",
        "MAINTAINER.md",
        "RECOVERY.md",
        "INCIDENT_RESPONSE.md",
        "DATA_GOVERNANCE.md"
    )) {
        $Source = Join-Path $RepositoryRoot $RelativePath
        if (Test-Path -LiteralPath $Source -PathType Leaf) {
            $Destination = Join-Path $MetadataRoot (
                $RelativePath -replace "[\\/]", "__"
            )
            Copy-Item -LiteralPath $Source -Destination $Destination
        }
    }

    $OperationalDataIncluded = $false
    $OperationalSourceFingerprint = $null
    if ($IncludeOperationalData) {
        $WorkflowLockPath = Join-Path $OperationsRoot (
            "03 Archive\run-manifests\.weekly-snapshot.lock"
        )
        if (-not (Test-Path -LiteralPath $WorkflowLockPath -PathType Leaf)) {
            throw "The weekly workflow lock is missing: $WorkflowLockPath"
        }
        $WorkflowLockEntry = Get-Item -LiteralPath $WorkflowLockPath -Force
        if (Test-UnsafeReparsePoint -Entry $WorkflowLockEntry) {
            throw "The weekly workflow lock is a link or reparse point."
        }
        $OperationalLockStream = [System.IO.File]::Open(
            $WorkflowLockPath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::ReadWrite
        )
        try {
            if ($OperationalLockStream.Length -eq 0) {
                $OperationalLockStream.WriteByte(0)
                $OperationalLockStream.Flush()
            }
            $OperationalLockStream.Position = 0
            $OperationalLockStream.Lock(0, 1)
            $OperationalLockAcquired = $true
        } catch {
            $OperationalLockStream.Dispose()
            $OperationalLockStream = $null
            throw (
                "Another weekly snapshot or integrity operation is already running. " +
                "Let it finish before building an operational recovery bundle."
            )
        }
        $BeforeFingerprint = Get-OperationalSourceFingerprint `
            -Root $OperationsRoot -AnchorDirectory $IntegrityAnchorDirectory
        $OperationalStage = Join-Path $TemporaryRoot "operational-data"
        New-Item -ItemType Directory -Path $OperationalStage | Out-Null
        foreach ($RelativePath in @(
            "02 Finished Reports",
            "03 Archive\processed-daily-reports",
            "03 Archive\generated-workbooks",
            "03 Archive\run-manifests"
        )) {
            $Source = Join-Path $OperationsRoot $RelativePath
            Assert-NormalTree -Path $Source -Label $RelativePath
            $Destination = Join-Path $OperationalStage (
                $RelativePath -replace "[\\/]", "__"
            )
            Copy-NormalTree -Source $Source -Destination $Destination
        }
        foreach ($RelativePath in @(
            "03 Archive\run-attempts",
            "03 Archive\approved-management-evidence"
        )) {
            $Source = Join-Path $OperationsRoot $RelativePath
            if (Test-Path -LiteralPath $Source) {
                Assert-NormalTree -Path $Source -Label $RelativePath
                $Destination = Join-Path $OperationalStage (
                    $RelativePath -replace "[\\/]", "__"
                )
                Copy-NormalTree -Source $Source -Destination $Destination
            }
        }
        Assert-NormalTree -Path $IntegrityAnchorDirectory -Label "Integrity anchor"
        Copy-NormalTree -Source $IntegrityAnchorDirectory `
            -Destination (Join-Path $OperationalStage "integrity-anchors")
        $AfterFingerprint = Get-OperationalSourceFingerprint `
            -Root $OperationsRoot -AnchorDirectory $IntegrityAnchorDirectory
        if ($BeforeFingerprint -ne $AfterFingerprint) {
            throw (
                "Operational source state changed during recovery capture. " +
                "No recovery bundle was published; wait for processing and Dropbox " +
                "sync to finish, then retry."
            )
        }
        $OperationalSourceFingerprint = $BeforeFingerprint
        $OperationalLockStream.Unlock(0, 1)
        $OperationalLockAcquired = $false
        $OperationalLockStream.Dispose()
        $OperationalLockStream = $null
        Compress-Archive -Path (
            Join-Path $OperationalStage "*"
        ) -DestinationPath (
            Join-Path $PayloadRoot "restricted-operational-data.zip"
        ) -CompressionLevel Optimal
        $OperationalDataIncluded = $true
    }

    $ReleaseMetadata = [ordered]@{
        schema_version = 1
        created_at_utc = [DateTime]::UtcNow.ToString("o")
        release_tag = $ReleaseTag
        commit = $Commit
        branch = $Branch
        retention_class = $RetentionClass
        retention_policy = if ($RetentionClass -eq "Weekly") {
            "Retain the newest 13 weekly bundles."
        } else {
            "Retain the newest 12 monthly bundles."
        }
        operational_data_included = $OperationalDataIncluded
        operational_source_fingerprint_sha256 = $OperationalSourceFingerprint
        operational_capture_lock = if ($OperationalDataIncluded) {
            "03 Archive/run-manifests/.weekly-snapshot.lock"
        } else {
            $null
        }
        data_classification = if ($OperationalDataIncluded) {
            "Restricted Employee Performance Information"
        } else {
            "Internal Technical Recovery Material"
        }
        automatic_upload = $false
        automatic_send = $false
        restore_test_cadence = "Quarterly"
    }
    $ReleaseMetadata | ConvertTo-Json -Depth 5 |
        Set-Content -LiteralPath (
            Join-Path $PayloadRoot "release-metadata.json"
        ) -Encoding UTF8

    $HashLines = @()
    foreach ($File in Get-ChildItem -LiteralPath $PayloadRoot -Recurse -File |
        Sort-Object FullName) {
        $Relative = $File.FullName.Substring($PayloadRoot.Length).TrimStart("\")
        $HashLines += "$(Get-FileSha256 -Path $File.FullName)  $($Relative -replace '\\','/')"
    }
    $HashLines | Set-Content -LiteralPath (
        Join-Path $PayloadRoot "SHA256SUMS.txt"
    ) -Encoding UTF8
    Compress-Archive -Path (
        Join-Path $PayloadRoot "*"
    ) -DestinationPath $FinalBundle -CompressionLevel Optimal
    $FinalHash = Get-FileSha256 -Path $FinalBundle
    "$FinalHash  $([System.IO.Path]::GetFileName($FinalBundle))" |
        Set-Content -LiteralPath $FinalSidecar -Encoding UTF8
    Write-Output $FinalBundle
    Write-Output $FinalSidecar
} finally {
    if ($null -ne $OperationalLockStream) {
        if ($OperationalLockAcquired) {
            try {
                $OperationalLockStream.Unlock(0, 1)
            } catch {
                Write-Warning "Could not explicitly unlock the recovery capture handle."
            }
        }
        $OperationalLockStream.Dispose()
    }
    if (
        (Test-Path -LiteralPath $TemporaryRoot) `
        -and $TemporaryRoot.StartsWith(
            $ExpectedTemporaryPrefix, [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        Remove-Item -LiteralPath $TemporaryRoot -Recurse -Force
    }
}
