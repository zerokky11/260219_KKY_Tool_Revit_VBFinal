param(
    [string]$Label = 'manual',
    [string]$Version = 'current'
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$backupRoot = Join-Path $repoRoot 'artifacts\test-build-backups'
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$safeLabel = (($Label -replace '[^A-Za-z0-9._-]', '_').Trim('_'))
if ([string]::IsNullOrWhiteSpace($safeLabel)) {
    $safeLabel = 'manual'
}

$snapshotName = "{0}_{1}" -f $timestamp, $safeLabel
$snapshotRoot = Join-Path $backupRoot $snapshotName
$filesRoot = Join-Path $snapshotRoot 'files'
$manifestPath = Join-Path $snapshotRoot 'manifest.json'
$listPath = Join-Path $snapshotRoot 'changed-files.txt'

$excludedPrefixes = @(
    '.git/',
    '.dotnet-cli/',
    'artifacts/',
    'Sever/Release/',
    'Compile/Compile/',
    '_build/',
    '_buildcheck/',
    '_temp_build/',
    '_temp_addins/',
    'tmp/',
    'Output/'
)

function Test-IsExcludedPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RelativePath
    )

    $normalized = $RelativePath.Replace('\', '/')
    foreach ($prefix in $excludedPrefixes) {
        if ($normalized.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Get-ChangedEntries {
    $lines = & git -C $repoRoot -c core.quotepath=false status --porcelain=v1 --untracked-files=all
    if ($LASTEXITCODE -ne 0) {
        throw 'git status failed while collecting changed files.'
    }

    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line) -or $line.Length -lt 4) {
            continue
        }

        $status = $line.Substring(0, 2)
        $rawPath = $line.Substring(3).Trim()
        $relativePath = $rawPath
        if ($rawPath.Contains(' -> ')) {
            $relativePath = $rawPath.Split(' -> ', 2)[1]
        }

        if ($relativePath.StartsWith('"') -and $relativePath.EndsWith('"') -and $relativePath.Length -ge 2) {
            $relativePath = $relativePath.Substring(1, $relativePath.Length - 2)
        }

        $relativePath = $relativePath.Replace('\', '/')
        $relativePath = $relativePath.Replace('\"', '"')
        if (Test-IsExcludedPath -RelativePath $relativePath) {
            continue
        }

        $absolutePath = Join-Path $repoRoot $relativePath
        $exists = Test-Path -LiteralPath $absolutePath -PathType Leaf

        $entries.Add([pscustomobject]@{
            Status = $status
            RelativePath = $relativePath
            AbsolutePath = $absolutePath
            Exists = $exists
        })
    }

    return $entries
}

function Copy-ChangedFiles {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[object]]$Entries
    )

    New-Item -ItemType Directory -Path $filesRoot -Force | Out-Null

    $copied = New-Object System.Collections.Generic.List[string]
    $missing = New-Object System.Collections.Generic.List[string]

    foreach ($entry in $Entries) {
        if (-not $entry.Exists) {
            $missing.Add($entry.RelativePath)
            continue
        }

        $destination = Join-Path $filesRoot $entry.RelativePath
        $destinationDir = Split-Path -Parent $destination
        if (-not (Test-Path -LiteralPath $destinationDir)) {
            New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
        }

        Copy-Item -LiteralPath $entry.AbsolutePath -Destination $destination -Force
        $copied.Add($entry.RelativePath)
    }

    return [pscustomobject]@{
        Copied = $copied
        Missing = $missing
    }
}

$headCommit = (& git -C $repoRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0) {
    throw 'git rev-parse HEAD failed.'
}

$branchName = (& git -C $repoRoot rev-parse --abbrev-ref HEAD).Trim()
if ($LASTEXITCODE -ne 0) {
    throw 'git rev-parse --abbrev-ref HEAD failed.'
}

$entries = Get-ChangedEntries
$copyResult = Copy-ChangedFiles -Entries $entries

$manifest = [pscustomobject]@{
    SnapshotName = $snapshotName
    Label = $Label
    Version = $Version
    CreatedAt = (Get-Date).ToString('s')
    RepoRoot = $repoRoot
    Branch = $branchName
    HeadCommit = $headCommit
    CopiedFileCount = $copyResult.Copied.Count
    MissingFileCount = $copyResult.Missing.Count
    CopiedFiles = $copyResult.Copied
    MissingFiles = $copyResult.Missing
    Entries = $entries
}

$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding utf8
$copyResult.Copied | Set-Content -LiteralPath $listPath -Encoding utf8

Write-Host "Backup created: $snapshotRoot"
Write-Host "Copied files : $($copyResult.Copied.Count)"
Write-Host "Missing files: $($copyResult.Missing.Count)"
Write-Host "Manifest     : $manifestPath"
Write-Output $snapshotRoot
