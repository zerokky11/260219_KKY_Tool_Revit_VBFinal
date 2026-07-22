param(
    [string]$Version = '1.0.1',
    [string]$Label,
    [string]$InnoCompilerPath,
    [double]$MailPackageMinimumMB = 15.9,
    [switch]$PreserveExistingArtifacts,
    [switch]$SkipMailPackage
)

$ErrorActionPreference = 'Stop'

function Add-RandomBytes {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][Int64]$Bytes
    )

    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $buffer = New-Object byte[] 65536
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try {
            $remaining = $Bytes
            while ($remaining -gt 0) {
                $take = [int][Math]::Min($buffer.Length, $remaining)
                $rng.GetBytes($buffer)
                $stream.Write($buffer, 0, $take)
                $remaining -= $take
            }
        }
        finally {
            $stream.Close()
        }
    }
    finally {
        $rng.Dispose()
    }
}

function New-MailSizedInstallerPackage {
    param(
        [Parameter(Mandatory = $true)][string]$InstallerPath,
        [Parameter(Mandatory = $true)][string]$PackageRoot,
        [Parameter(Mandatory = $true)][string]$PackageLabel,
        [Parameter(Mandatory = $true)][double]$MinimumMB
    )

    New-Item -ItemType Directory -Path $PackageRoot -Force | Out-Null

    $dateStamp = Get-Date -Format 'yyyyMMdd'
    $packageIndex = 1
    do {
        $shortLabel = '{0}_{1:00}' -f $dateStamp, $packageIndex
        $zipPath = Join-Path $PackageRoot "$shortLabel.zip"
        $packageIndex++
    } while (Test-Path -LiteralPath $zipPath)

    $workDir = Join-Path $PackageRoot "mailpkg-$shortLabel"
    $packageRootFull = [System.IO.Path]::GetFullPath($PackageRoot).TrimEnd('\') + '\'
    $workDirFull = [System.IO.Path]::GetFullPath($workDir)
    if (-not $workDirFull.StartsWith($packageRootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsafe mail-package work directory: $workDirFull"
    }
    if (Test-Path -LiteralPath $workDir) {
        Remove-Item -LiteralPath $workDir -Recurse -Force
    }
    try {
        New-Item -ItemType Directory -Path $workDir -Force | Out-Null

        $installerName = 'Setup.exe'
        Copy-Item -LiteralPath $InstallerPath -Destination (Join-Path $workDir $installerName) -Force

        @"
KKY Family Browser installer mail package

Run this installer:
$installerName

Ignore this file:
mail_size_padding_do_not_run.bin

The padding file only makes this zip large enough for large-file mail attachment handling.
"@ | Set-Content -LiteralPath (Join-Path $workDir 'README.txt') -Encoding UTF8

        $paddingPath = Join-Path $workDir 'mail_size_padding_do_not_run.bin'
        New-Item -ItemType File -Path $paddingPath -Force | Out-Null

        $minimumBytes = [int64][Math]::Ceiling($MinimumMB * 1MB)
        $currentInputBytes = (Get-ChildItem -LiteralPath $workDir -File | Measure-Object -Property Length -Sum).Sum
        $initialPaddingBytes = [int64][Math]::Max(1MB, $minimumBytes - $currentInputBytes + 256KB)
        Add-RandomBytes -Path $paddingPath -Bytes $initialPaddingBytes

        for ($attempt = 0; $attempt -lt 4; $attempt++) {
            if (Test-Path -LiteralPath $zipPath) {
                Remove-Item -LiteralPath $zipPath -Force
            }
            Get-ChildItem -LiteralPath $workDir -File | Compress-Archive -DestinationPath $zipPath -CompressionLevel Optimal
            $zipItem = Get-Item -LiteralPath $zipPath
            if ($zipItem.Length -gt $minimumBytes) {
                return $zipItem.FullName
            }
            $extraBytes = [int64]($minimumBytes - $zipItem.Length + 512KB)
            Add-RandomBytes -Path $paddingPath -Bytes $extraBytes
        }

        $finalZipItem = Get-Item -LiteralPath $zipPath
        if ($finalZipItem.Length -le $minimumBytes) {
            throw "Mail package is still smaller than $MinimumMB MB: $($finalZipItem.FullName)"
        }
        return $finalZipItem.FullName
    }
    finally {
        if (Test-Path -LiteralPath $workDir) {
            $cleanupFull = [System.IO.Path]::GetFullPath($workDir)
            if (-not $cleanupFull.StartsWith($packageRootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "Unsafe mail-package cleanup directory: $cleanupFull"
            }
            Remove-Item -LiteralPath $workDir -Recurse -Force
        }
    }
}

function Remove-OldFamilyBrowserInstallerArtifacts {
    param(
        [Parameter(Mandatory = $true)][string[]]$Roots,
        [Parameter(Mandatory = $true)][datetime]$CutoffDate
    )

    $removed = 0
    foreach ($root in $Roots) {
        if (-not (Test-Path -LiteralPath $root)) {
            continue
        }

        $resolvedRoot = (Resolve-Path -LiteralPath $root).ProviderPath
        $items = Get-ChildItem -LiteralPath $resolvedRoot -Force |
            Where-Object { $_.LastWriteTime.Date -lt $CutoffDate.Date } |
            Sort-Object { $_.FullName.Length } -Descending

        foreach ($item in $items) {
            Remove-Item -LiteralPath $item.FullName -Recurse -Force
            $removed++
        }
    }

    if ($removed -gt 0) {
        Write-Host ("Removed old Family Browser installer artifacts before {0}: {1}" -f $CutoffDate.ToString('yyyy-MM-dd'), $removed) -ForegroundColor Yellow
    }
    else {
        Write-Host ("No old Family Browser installer artifacts before {0}." -f $CutoffDate.ToString('yyyy-MM-dd')) -ForegroundColor DarkGray
    }
}

function Assert-PortableExecutable {
    param([Parameter(Mandatory = $true)][string]$Path)

    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
        if ($stream.Length -lt 2 -or $stream.ReadByte() -ne 0x4D -or $stream.ReadByte() -ne 0x5A) {
            throw "Installer does not have a valid PE MZ header: $Path"
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Get-MailPackageValidation {
    param(
        [Parameter(Mandatory = $true)][string]$ZipPath,
        [Parameter(Mandatory = $true)][string]$ExpectedInstallerSha256
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $entries = @($archive.Entries | ForEach-Object { $_.FullName } | Sort-Object)
        foreach ($requiredEntry in @('Setup.exe', 'README.txt', 'mail_size_padding_do_not_run.bin')) {
            if ($entries -notcontains $requiredEntry) {
                throw "Mail package is missing $requiredEntry"
            }
        }
        $setupEntry = $archive.Entries | Where-Object { $_.FullName -eq 'Setup.exe' } | Select-Object -First 1
        if (-not $setupEntry) {
            throw 'Mail package Setup.exe entry was not found.'
        }
        $setupStream = $setupEntry.Open()
        $sha = [System.Security.Cryptography.SHA256]::Create()
        try {
            $setupHash = [System.BitConverter]::ToString($sha.ComputeHash($setupStream)).Replace('-', '')
        }
        finally {
            $sha.Dispose()
            $setupStream.Dispose()
        }
        if ($setupHash -ne $ExpectedInstallerSha256) {
            throw "Mail package Setup.exe hash mismatch. expected=$ExpectedInstallerSha256 actual=$setupHash"
        }
        return [pscustomobject]@{
            Entries = $entries
            InstallerSha256InsideZip = $setupHash
        }
    }
    finally {
        $archive.Dispose()
    }
}

function Write-JsonAtomic {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Value
    )

    $folder = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $temporaryPath = Join-Path $folder ('.' + [System.IO.Path]::GetFileName($Path) + '.' + [Guid]::NewGuid().ToString('N') + '.tmp')
    try {
        $Value | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $temporaryPath -Encoding UTF8
        Move-Item -LiteralPath $temporaryPath -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
    }
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
$productVersionSourcePath = Join-Path $repoRoot 'KKY_FamilyBrowser_SharedUi\FamilyBrowserProductUpdateService.cs'
if (-not (Test-Path -LiteralPath $productVersionSourcePath)) {
    throw "Family Browser product version source was not found: $productVersionSourcePath"
}
$productVersionSource = Get-Content -LiteralPath $productVersionSourcePath -Raw
$productVersionMatch = [regex]::Match(
    $productVersionSource,
    'public\s+const\s+string\s+CurrentProductVersion\s*=\s*"([^"]+)"\s*;'
)
if (-not $productVersionMatch.Success) {
    throw "CurrentProductVersion was not found in: $productVersionSourcePath"
}
$sourceProductVersion = $productVersionMatch.Groups[1].Value.Trim()
$requestedPackageVersion = if ($null -eq $Version) { '' } else { $Version.Trim() }
if ([string]::IsNullOrWhiteSpace($requestedPackageVersion)) {
    throw 'Installer version cannot be empty.'
}
try {
    [void][version]$requestedPackageVersion
}
catch {
    throw "Installer version is invalid: $requestedPackageVersion"
}
if ($sourceProductVersion -cne $requestedPackageVersion) {
    throw "Family Browser source/package version mismatch. source=$sourceProductVersion package=$requestedPackageVersion"
}
$Version = $requestedPackageVersion

if (-not $Label) {
    $Label = "recovered-{0}" -f (Get-Date -Format 'yyyyMMdd-HHmm')
}

if (-not $InnoCompilerPath) {
    $candidates = @(
        'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
        'C:\Program Files\Inno Setup 6\ISCC.exe',
        'C:\Program Files (x86)\Inno Setup 5\ISCC.exe',
        'C:\Program Files\Inno Setup 5\ISCC.exe'
    )
    $InnoCompilerPath = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
}

if (-not $InnoCompilerPath -or -not (Test-Path -LiteralPath $InnoCompilerPath)) {
    throw 'Inno Setup compiler was not found. Install Inno Setup 6 or pass -InnoCompilerPath.'
}

$buildScript = Join-Path $scriptRoot 'Build-FamilyBrowserRecovered.ps1'
$verifyScript = Join-Path $scriptRoot 'Verify-FamilyBrowserRecovered.ps1'
$issPath = Join-Path $scriptRoot 'KKY_FamilyBrowser_Compiler.iss'
$outputDir = Join-Path $repoRoot 'artifacts\family-browser\installers'
$mailPackageRoot = Join-Path $repoRoot 'artifacts\family-browser\mail-packages'
$outputBaseName = "KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v{0}_{1}_Setup" -f $Version, $Label
$outputExe = Join-Path $outputDir "$outputBaseName.exe"

New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
New-Item -ItemType Directory -Path $mailPackageRoot -Force | Out-Null
if ($PreserveExistingArtifacts) {
    Write-Host 'Preserving all existing Family Browser installer and mail-package artifacts.' -ForegroundColor DarkGray
}
else {
    Remove-OldFamilyBrowserInstallerArtifacts -Roots @($outputDir, $mailPackageRoot) -CutoffDate (Get-Date).Date
}

& $buildScript
& $verifyScript

$stageManifestPath = Join-Path $repoRoot 'artifacts\family-browser\stage\stage-manifest.json'
if (-not (Test-Path -LiteralPath $stageManifestPath)) {
    throw "Stage manifest was not created: $stageManifestPath"
}
$stageManifest = Get-Content -LiteralPath $stageManifestPath -Raw | ConvertFrom-Json

Write-Host "Compiling Family Browser installer..." -ForegroundColor Cyan
& $InnoCompilerPath `
    "/DMyAppVersion=$Version" `
    "/DMyOutputBaseName=$outputBaseName" `
    "/DMyOutputDir=$outputDir" `
    $issPath

if ($LASTEXITCODE -ne 0) {
    throw "Inno Setup compiler failed with exit code $LASTEXITCODE"
}

if (-not (Test-Path -LiteralPath $outputExe)) {
    throw "Installer was not created: $outputExe"
}

Assert-PortableExecutable -Path $outputExe
$hash = Get-FileHash -LiteralPath $outputExe -Algorithm SHA256
$hashPath = Join-Path $outputDir "$outputBaseName.sha256.txt"
"$($hash.Hash)  $([System.IO.Path]::GetFileName($outputExe))" | Set-Content -LiteralPath $hashPath -Encoding ASCII

$mailPackage = $null
$mailPackageHash = $null
if (-not $SkipMailPackage) {
    $mailPackage = New-MailSizedInstallerPackage `
        -InstallerPath $outputExe `
        -PackageRoot $mailPackageRoot `
        -PackageLabel $outputBaseName `
        -MinimumMB $MailPackageMinimumMB
    $mailPackageHash = (Get-FileHash -LiteralPath $mailPackage -Algorithm SHA256).Hash
}

$createdAt = (Get-Date).ToString('o')
$installerItem = Get-Item -LiteralPath $outputExe
$latestBuildPath = Join-Path $outputDir 'latest-build.json'
$latestBuild = [ordered]@{
    Installer = $installerItem.FullName
    Bytes = $installerItem.Length
    SizeMB = [Math]::Round($installerItem.Length / 1MB, 2)
    Sha256 = $hash.Hash
    Sha256File = $hashPath
    Label = $Label
    Version = $Version
    Created = $createdAt
    StageManifest = $stageManifestPath
    StageGeneratedAt = [string]$stageManifest.generatedAt
    StagePayloadFileCount = [int]$stageManifest.payloadFileCount
    StagePayload = @($stageManifest.payload)
}
Write-JsonAtomic -Path $latestBuildPath -Value $latestBuild

$latestMailPackagePath = $null
$mailPackageHashPath = $null
$mailValidation = $null
if ($mailPackage) {
    $mailValidation = Get-MailPackageValidation -ZipPath $mailPackage -ExpectedInstallerSha256 $hash.Hash
    $mailPackageHashPath = $mailPackage + '.sha256.txt'
    "$mailPackageHash  $([System.IO.Path]::GetFileName($mailPackage))" | Set-Content -LiteralPath $mailPackageHashPath -Encoding ASCII
    $mailPackageItem = Get-Item -LiteralPath $mailPackage
    $latestMailPackagePath = Join-Path $mailPackageRoot 'latest-mail-package.json'
    $latestMailPackage = [ordered]@{
        Package = $mailPackageItem.FullName
        Bytes = $mailPackageItem.Length
        SizeMB = [Math]::Round($mailPackageItem.Length / 1MB, 2)
        Sha256 = $mailPackageHash
        Sha256File = $mailPackageHashPath
        InstallerSha256InsideZip = $mailValidation.InstallerSha256InsideZip
        Installer = $installerItem.FullName
        InstallerSha256 = $hash.Hash
        Entries = @($mailValidation.Entries)
        Version = $Version
        Label = $Label
        Created = $createdAt
    }
    Write-JsonAtomic -Path $latestMailPackagePath -Value $latestMailPackage
}

$distributionAuditPath = $null
if ($mailPackage) {
    $distributionAuditScript = Join-Path $scriptRoot 'Test-FamilyBrowserDistribution.ps1'
    if (-not (Test-Path -LiteralPath $distributionAuditScript)) {
        throw "Distribution audit script was not found: $distributionAuditScript"
    }
    $safeAuditLabel = ($Label -replace '[^A-Za-z0-9._-]', '_')
    $distributionAuditPath = Join-Path $repoRoot ('artifacts\family-browser-distribution-audit\' + $safeAuditLabel)
    & $distributionAuditScript `
        -StageRoot (Join-Path $repoRoot 'artifacts\family-browser\stage') `
        -LatestBuildMetadata $latestBuildPath `
        -LatestMailPackageMetadata $latestMailPackagePath `
        -MinimumMailPackageMB $MailPackageMinimumMB `
        -OutputDir $distributionAuditPath
}

Write-Host ""
Write-Host "Family Browser installer created:" -ForegroundColor Green
Write-Host $outputExe
Write-Host "SHA256:"
Write-Host $hash.Hash
if ($mailPackage) {
    $mailPackageItem = Get-Item -LiteralPath $mailPackage
    Write-Host ""
    Write-Host "Mail-sized package created:" -ForegroundColor Green
    Write-Host $mailPackage
    Write-Host ("Size MB: {0}" -f ([Math]::Round($mailPackageItem.Length / 1MB, 2)))
    Write-Host "SHA256:"
    Write-Host $mailPackageHash
}

[pscustomobject]@{
    Installer = $outputExe
    Sha256 = $hash.Hash
    Sha256File = $hashPath
    LatestBuildMetadata = $latestBuildPath
    MailPackage = $mailPackage
    MailPackageSha256 = $mailPackageHash
    MailPackageSha256File = $mailPackageHashPath
    LatestMailPackageMetadata = $latestMailPackagePath
    DistributionAudit = $distributionAuditPath
}
