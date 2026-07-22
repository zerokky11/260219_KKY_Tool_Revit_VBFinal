param(
    [string]$StageRoot,
    [string]$LatestBuildMetadata,
    [string]$LatestMailPackageMetadata,
    [double]$MinimumMailPackageMB = 15.9,
    [string]$OutputDir,
    [switch]$VerifyInstalled
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
if (-not $StageRoot) {
    $StageRoot = Join-Path $repoRoot 'artifacts\family-browser\stage'
}
if (-not $LatestBuildMetadata) {
    $LatestBuildMetadata = Join-Path $repoRoot 'artifacts\family-browser\installers\latest-build.json'
}
if (-not $LatestMailPackageMetadata) {
    $LatestMailPackageMetadata = Join-Path $repoRoot 'artifacts\family-browser\mail-packages\latest-mail-package.json'
}
if (-not $OutputDir) {
    $OutputDir = Join-Path $repoRoot ('artifacts\family-browser-distribution-audit\' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

$checks = New-Object System.Collections.Generic.List[object]
$failures = New-Object System.Collections.Generic.List[string]

function Invoke-DistributionCheck {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$Script
    )

    try {
        $detail = & $Script
        $checks.Add([pscustomobject]@{ Name = $Name; Status = 'PASS'; Detail = (($detail | Out-String).Trim()) }) | Out-Null
    }
    catch {
        $message = $_.Exception.Message
        $checks.Add([pscustomobject]@{ Name = $Name; Status = 'FAIL'; Detail = $message }) | Out-Null
        $failures.Add("${Name}: $message") | Out-Null
    }
}

function Get-StreamSha256 {
    param([Parameter(Mandatory = $true)][System.IO.Stream]$Stream)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return [System.BitConverter]::ToString($sha.ComputeHash($Stream)).Replace('-', '')
    }
    finally {
        $sha.Dispose()
    }
}

function Assert-ChecksumFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ExpectedHash
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Checksum file is missing: $Path"
    }
    $line = (Get-Content -LiteralPath $Path -First 1).Trim()
    if (-not $line.StartsWith($ExpectedHash, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Checksum file does not contain the expected hash: $Path"
    }
}

$stageManifestPath = Join-Path $StageRoot 'stage-manifest.json'
$context = [ordered]@{
    StageManifest = $null
    BuildMetadata = $null
    MailMetadata = $null
}

Invoke-DistributionCheck -Name 'Stage manifest and payload hashes' -Script {
    if (-not (Test-Path -LiteralPath $stageManifestPath)) {
        throw "Stage manifest is missing: $stageManifestPath"
    }
    $context.StageManifest = Get-Content -LiteralPath $stageManifestPath -Raw | ConvertFrom-Json
    & (Join-Path $scriptRoot 'Verify-FamilyBrowserRecovered.ps1') -StageRoot $StageRoot | Out-Null
    "payload=$($context.StageManifest.payloadFileCount)"
}

Invoke-DistributionCheck -Name 'Installer metadata and freshness' -Script {
    if (-not (Test-Path -LiteralPath $LatestBuildMetadata)) {
        throw "Latest installer metadata is missing: $LatestBuildMetadata"
    }
    $context.BuildMetadata = Get-Content -LiteralPath $LatestBuildMetadata -Raw | ConvertFrom-Json
    $installerPath = [string]$context.BuildMetadata.Installer
    if (-not (Test-Path -LiteralPath $installerPath)) {
        throw "Installer is missing: $installerPath"
    }
    $installer = Get-Item -LiteralPath $installerPath
    $stream = [System.IO.File]::Open($installer.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
        if ($stream.Length -lt 2 -or $stream.ReadByte() -ne 0x4D -or $stream.ReadByte() -ne 0x5A) {
            throw 'Installer PE MZ header is invalid.'
        }
    }
    finally {
        $stream.Dispose()
    }
    $hash = (Get-FileHash -LiteralPath $installer.FullName -Algorithm SHA256).Hash
    if ($hash -ne [string]$context.BuildMetadata.Sha256) {
        throw 'Installer SHA256 does not match latest-build.json.'
    }
    if ($installer.Length -ne [int64]$context.BuildMetadata.Bytes) {
        throw 'Installer length does not match latest-build.json.'
    }
    if ($installer.LastWriteTimeUtc -lt (Get-Item -LiteralPath $stageManifestPath).LastWriteTimeUtc) {
        throw 'Installer is older than the current Stage manifest.'
    }
    Assert-ChecksumFile -Path ([string]$context.BuildMetadata.Sha256File) -ExpectedHash $hash
    "installer=$($installer.Name); sha256=$hash"
}

Invoke-DistributionCheck -Name 'Installer embeds current Stage revision' -Script {
    if ($null -eq $context.StageManifest -or $null -eq $context.BuildMetadata) {
        throw 'Stage or installer metadata was not available from the preceding checks.'
    }
    $stageEntries = @($context.StageManifest.payload)
    $buildEntries = @($context.BuildMetadata.StagePayload)
    if ($stageEntries.Count -eq 0 -or $stageEntries.Count -ne $buildEntries.Count) {
        throw "Stage snapshot count mismatch. stage=$($stageEntries.Count), installer=$($buildEntries.Count)"
    }
    $buildByPath = @{}
    foreach ($entry in $buildEntries) {
        $buildByPath[[string]$entry.relativePath] = $entry
    }
    foreach ($entry in $stageEntries) {
        $path = [string]$entry.relativePath
        if (-not $buildByPath.ContainsKey($path)) {
            throw "Installer metadata is missing Stage payload: $path"
        }
        $embedded = $buildByPath[$path]
        if ([int64]$embedded.bytes -ne [int64]$entry.bytes -or [string]$embedded.sha256 -ne [string]$entry.sha256) {
            throw "Installer Stage revision differs at: $path"
        }
    }
    "payload=$($stageEntries.Count)"
}

Invoke-DistributionCheck -Name 'Mail ZIP and embedded Setup integrity' -Script {
    if (-not (Test-Path -LiteralPath $LatestMailPackageMetadata)) {
        throw "Latest mail-package metadata is missing: $LatestMailPackageMetadata"
    }
    $context.MailMetadata = Get-Content -LiteralPath $LatestMailPackageMetadata -Raw | ConvertFrom-Json
    $packagePath = [string]$context.MailMetadata.Package
    if (-not (Test-Path -LiteralPath $packagePath)) {
        throw "Mail package is missing: $packagePath"
    }
    $package = Get-Item -LiteralPath $packagePath
    $minimumBytes = [int64][Math]::Ceiling($MinimumMailPackageMB * 1MB)
    if ($package.Length -le $minimumBytes) {
        throw "Mail package is not larger than $MinimumMailPackageMB MB."
    }
    $packageHash = (Get-FileHash -LiteralPath $package.FullName -Algorithm SHA256).Hash
    if ($packageHash -ne [string]$context.MailMetadata.Sha256) {
        throw 'Mail package SHA256 does not match latest-mail-package.json.'
    }
    Assert-ChecksumFile -Path ([string]$context.MailMetadata.Sha256File) -ExpectedHash $packageHash

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($package.FullName)
    try {
        $entries = @($archive.Entries | ForEach-Object { $_.FullName } | Sort-Object)
        $expectedEntries = @('README.txt', 'Setup.exe', 'mail_size_padding_do_not_run.bin') | Sort-Object
        if (($entries -join '|') -ne ($expectedEntries -join '|')) {
            throw "Mail ZIP entries differ: $($entries -join ', ')"
        }
        $setupEntry = $archive.Entries | Where-Object { $_.FullName -eq 'Setup.exe' } | Select-Object -First 1
        $setupStream = $setupEntry.Open()
        try {
            $setupHash = Get-StreamSha256 -Stream $setupStream
        }
        finally {
            $setupStream.Dispose()
        }
        if ($setupHash -ne [string]$context.BuildMetadata.Sha256 -or $setupHash -ne [string]$context.MailMetadata.InstallerSha256InsideZip) {
            throw 'Mail ZIP Setup.exe does not match the latest standalone installer.'
        }
    }
    finally {
        $archive.Dispose()
    }
    "package=$($package.Name); bytes=$($package.Length); sha256=$packageHash"
}

Invoke-DistributionCheck -Name 'No abandoned mail packaging work folders' -Script {
    if ($null -eq $context.MailMetadata) {
        throw 'Mail-package metadata was not available from the preceding check.'
    }
    $mailRoot = Split-Path -Parent ([string]$context.MailMetadata.Package)
    $abandoned = @(Get-ChildItem -LiteralPath $mailRoot -Directory -Filter 'mailpkg-*' -ErrorAction SilentlyContinue)
    if ($abandoned.Count -gt 0) {
        throw "Abandoned work folders: $($abandoned.Name -join ', ')"
    }
    'none'
}

if ($VerifyInstalled) {
    Invoke-DistributionCheck -Name 'ProgramData matches current Stage' -Script {
        & (Join-Path $scriptRoot 'Verify-FamilyBrowserRecovered.ps1') -Installed -StageRoot $StageRoot | Out-Null
        '2019/2021/2023/2025/2027 full payload match'
    }
}

$summary = [ordered]@{
    GeneratedAt = (Get-Date).ToString('o')
    Status = $(if ($failures.Count -eq 0) { 'PASS' } else { 'FAIL' })
    StageRoot = $StageRoot
    LatestBuildMetadata = $LatestBuildMetadata
    LatestMailPackageMetadata = $LatestMailPackageMetadata
    VerifyInstalled = $VerifyInstalled.IsPresent
    Checks = @($checks.ToArray())
    Failures = @($failures.ToArray())
}
$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDir 'distribution-audit-summary.json') -Encoding UTF8

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('# Family Browser Distribution Audit') | Out-Null
$lines.Add('') | Out-Null
$lines.Add("- Generated: $($summary.GeneratedAt)") | Out-Null
$lines.Add("- Status: $($summary.Status)") | Out-Null
$lines.Add("- Installed verified: $($summary.VerifyInstalled)") | Out-Null
$lines.Add('') | Out-Null
$lines.Add('| Check | Status | Detail |') | Out-Null
$lines.Add('|---|---:|---|') | Out-Null
foreach ($check in $checks) {
    $detail = ([string]$check.Detail).Replace('|', '/').Replace("`r", ' ').Replace("`n", ' ')
    $lines.Add("| $($check.Name) | $($check.Status) | $detail |") | Out-Null
}
$lines | Set-Content -LiteralPath (Join-Path $OutputDir 'distribution-audit-summary.md') -Encoding UTF8

if ($failures.Count -gt 0) {
    throw "Family Browser distribution audit failed. See $OutputDir"
}

Write-Host "Family Browser distribution audit passed: $OutputDir" -ForegroundColor Green
