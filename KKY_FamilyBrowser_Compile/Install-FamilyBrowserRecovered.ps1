[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string[]]$Years = @('2019', '2021', '2023', '2025', '2027'),
    [string]$StageRoot,
    [switch]$Build
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
if (-not $StageRoot) {
    $StageRoot = Join-Path $repoRoot 'artifacts\family-browser\stage'
}

if ($Build) {
    & (Join-Path $scriptRoot 'Build-FamilyBrowserRecovered.ps1') -Years $Years
}

if (-not (Test-Path -LiteralPath $StageRoot)) {
    throw "Stage folder does not exist: $StageRoot"
}

$runningRevit = @(Get-Process -Name Revit -ErrorAction SilentlyContinue)
if ($runningRevit.Count -gt 0) {
    throw "Close every Revit process before installing Family Browser. Running process IDs: $($runningRevit.Id -join ', ')"
}

function Test-FileContentEqual {
    param(
        [Parameter(Mandatory = $true)][string]$Left,
        [Parameter(Mandatory = $true)][string]$Right
    )

    if (-not (Test-Path -LiteralPath $Left) -or -not (Test-Path -LiteralPath $Right)) {
        return $false
    }

    $leftFile = Get-Item -LiteralPath $Left
    $rightFile = Get-Item -LiteralPath $Right
    if ($leftFile.Length -ne $rightFile.Length) {
        return $false
    }

    $leftHash = (Get-FileHash -LiteralPath $Left -Algorithm SHA256).Hash
    $rightHash = (Get-FileHash -LiteralPath $Right -Algorithm SHA256).Hash
    return $leftHash -eq $rightHash
}

function Test-PayloadContentEqual {
    param(
        [Parameter(Mandatory = $true)][string]$SourceFolder,
        [Parameter(Mandatory = $true)][string]$DestinationFolder
    )

    if (-not (Test-Path -LiteralPath $SourceFolder) -or -not (Test-Path -LiteralPath $DestinationFolder)) {
        return $false
    }

    $sourceFiles = @(Get-ChildItem -LiteralPath $SourceFolder -Recurse -File)
    $destinationFiles = @(Get-ChildItem -LiteralPath $DestinationFolder -Recurse -File)
    if ($sourceFiles.Count -eq 0 -or $sourceFiles.Count -ne $destinationFiles.Count) {
        return $false
    }

    $sourceRoot = [System.IO.Path]::GetFullPath($SourceFolder).TrimEnd('\') + '\'
    foreach ($sourceFile in $sourceFiles) {
        $relativePath = $sourceFile.FullName.Substring($sourceRoot.Length)
        $destinationFile = Join-Path $DestinationFolder $relativePath
        if (-not (Test-FileContentEqual -Left $sourceFile.FullName -Right $destinationFile)) {
            return $false
        }
    }
    return $true
}

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$windowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($windowsIdentity)
$isAdministrator = $windowsPrincipal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $WhatIfPreference -and -not $isAdministrator) {
    foreach ($year in $Years) {
        $installRoot = [System.IO.Path]::GetFullPath("C:\ProgramData\Autodesk\Revit\Addins\$year").TrimEnd('\')
        $expectedRoot = [System.IO.Path]::GetFullPath("C:\ProgramData\Autodesk\Revit\Addins\$year").TrimEnd('\')
        if (-not $installRoot.Equals($expectedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Unsafe Revit addin root while checking ProgramData write access: $installRoot"
        }
        $existingPayload = Join-Path $installRoot 'KKY_FamilyBrowser'
        $yearStage = Join-Path $StageRoot "Rvt$year"
        $stagedManifest = Get-ChildItem -LiteralPath $yearStage -Filter '*.addin' -File | Select-Object -First 1
        if (-not $stagedManifest) {
            throw "Missing staged addin manifest while checking ProgramData write access: $yearStage"
        }
        $existingManifests = @(
            'KKY_FamilyBrowser_RevitHost.addin',
            'KKY_FamilyBrowser_RevitHost_2025.addin',
            'KKY_FamilyBrowser_RevitHost_2027.addin'
        ) | ForEach-Object { Join-Path $installRoot $_ } | Where-Object { Test-Path -LiteralPath $_ }
        New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
        $probeFolders = @($installRoot)
        if (Test-Path -LiteralPath $existingPayload) {
            $probeFolders += $existingPayload
        }
        foreach ($probeFolder in $probeFolders) {
            $probePath = Join-Path $probeFolder ('.kky-family-browser-install-probe-' + [Guid]::NewGuid().ToString('N') + '.tmp')
            $probeStream = $null
            try {
                $probeStream = [System.IO.File]::Open($probePath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
                $probeStream.WriteByte(75)
            }
            catch {
                throw "Family Browser ProgramData installation requires administrator rights or explicit write access to $probeFolder. No installed files were changed. $($_.Exception.Message)"
            }
            finally {
                if ($null -ne $probeStream) {
                    $probeStream.Dispose()
                }
                if (Test-Path -LiteralPath $probePath) {
                    [System.IO.File]::Delete($probePath)
                }
            }
        }
        foreach ($manifestPath in $existingManifests) {
            if ((Split-Path -Leaf $manifestPath) -eq $stagedManifest.Name -and
                (Test-FileContentEqual -Left $manifestPath -Right $stagedManifest.FullName)) {
                continue
            }
            $manifestStream = $null
            try {
                $manifestStream = [System.IO.File]::Open($manifestPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::ReadWrite)
            }
            catch {
                throw "Family Browser ProgramData installation cannot replace $manifestPath without administrator rights or explicit write access. No installed files were changed. $($_.Exception.Message)"
            }
            finally {
                if ($null -ne $manifestStream) {
                    $manifestStream.Dispose()
                }
            }
        }
    }
}

function Assert-ExactInstallPath {
    param(
        [Parameter(Mandatory = $true)][string]$Actual,
        [Parameter(Mandatory = $true)][string]$Expected,
        [Parameter(Mandatory = $true)][string]$Description
    )

    $actualFull = [System.IO.Path]::GetFullPath($Actual).TrimEnd('\')
    $expectedFull = [System.IO.Path]::GetFullPath($Expected).TrimEnd('\')
    if (-not $actualFull.Equals($expectedFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsafe $Description path: $actualFull"
    }
}

function Assert-CopiedPayload {
    param(
        [Parameter(Mandatory = $true)][string]$SourceFolder,
        [Parameter(Mandatory = $true)][string]$CopiedFolder
    )

    $sourceFiles = @(Get-ChildItem -LiteralPath $SourceFolder -Recurse -File)
    $copiedFiles = @(Get-ChildItem -LiteralPath $CopiedFolder -Recurse -File)
    if ($sourceFiles.Count -eq 0 -or $sourceFiles.Count -ne $copiedFiles.Count) {
        throw "Copied payload count mismatch. source=$($sourceFiles.Count), copied=$($copiedFiles.Count)"
    }
    $sourceRoot = [System.IO.Path]::GetFullPath($SourceFolder).TrimEnd('\') + '\'
    foreach ($sourceFile in $sourceFiles) {
        $relativePath = $sourceFile.FullName.Substring($sourceRoot.Length)
        $copiedPath = Join-Path $CopiedFolder $relativePath
        if (-not (Test-Path -LiteralPath $copiedPath)) {
            throw "Copied payload file is missing: $relativePath"
        }
        $sourceHash = (Get-FileHash -LiteralPath $sourceFile.FullName -Algorithm SHA256).Hash
        $copiedHash = (Get-FileHash -LiteralPath $copiedPath -Algorithm SHA256).Hash
        if ($sourceHash -ne $copiedHash) {
            throw "Copied payload hash mismatch: $relativePath"
        }
    }
}

foreach ($year in $Years) {
    $yearStage = Join-Path $StageRoot "Rvt$year"
    if (-not (Test-Path -LiteralPath $yearStage)) {
        throw "Missing staged Revit year folder: $yearStage"
    }

    $destinationRoot = "C:\ProgramData\Autodesk\Revit\Addins\$year"
    $destinationFolder = Join-Path $destinationRoot 'KKY_FamilyBrowser'
    $stagedFolder = Join-Path $yearStage 'KKY_FamilyBrowser'
    $stagedAddin = Get-ChildItem -LiteralPath $yearStage -Filter '*.addin' -File | Select-Object -First 1
    if (-not $stagedAddin) {
        throw "Missing staged addin manifest: $yearStage"
    }

    $knownManifestNames = @(
        'KKY_FamilyBrowser_RevitHost.addin',
        'KKY_FamilyBrowser_RevitHost_2025.addin',
        'KKY_FamilyBrowser_RevitHost_2027.addin'
    )
    $installedManifest = Join-Path $destinationRoot $stagedAddin.Name
    $obsoleteManifestExists = @($knownManifestNames | Where-Object {
        $_ -ne $stagedAddin.Name -and (Test-Path -LiteralPath (Join-Path $destinationRoot $_))
    }).Count -gt 0
    $installationUpToDate =
        -not $obsoleteManifestExists -and
        (Test-FileContentEqual -Left $stagedAddin.FullName -Right $installedManifest) -and
        (Test-PayloadContentEqual -SourceFolder $stagedFolder -DestinationFolder $destinationFolder)
    if ($installationUpToDate) {
        Write-Host "Family Browser Revit $year is already current -> $destinationRoot" -ForegroundColor Green
        continue
    }

    if ($PSCmdlet.ShouldProcess($destinationRoot, "Install Family Browser Revit $year addin")) {
        Assert-ExactInstallPath -Actual $destinationRoot -Expected "C:\ProgramData\Autodesk\Revit\Addins\$year" -Description 'Revit addin root'
        Assert-ExactInstallPath -Actual $destinationFolder -Expected "C:\ProgramData\Autodesk\Revit\Addins\$year\KKY_FamilyBrowser" -Description 'Family Browser payload'
        New-Item -ItemType Directory -Path $destinationRoot -Force | Out-Null
        $temporaryFolder = Join-Path $destinationRoot ('KKY_FamilyBrowser.installing-' + [Guid]::NewGuid().ToString('N'))
        Assert-ExactInstallPath -Actual (Split-Path -Parent $temporaryFolder) -Expected $destinationRoot -Description 'temporary install parent'
        try {
            New-Item -ItemType Directory -Path $temporaryFolder -Force | Out-Null
            Get-ChildItem -LiteralPath $stagedFolder -Force | ForEach-Object {
                Copy-Item -LiteralPath $_.FullName -Destination $temporaryFolder -Recurse -Force
            }
            Assert-CopiedPayload -SourceFolder $stagedFolder -CopiedFolder $temporaryFolder

            if (Test-Path -LiteralPath $destinationFolder) {
                Remove-Item -LiteralPath $destinationFolder -Recurse -Force
            }
            Move-Item -LiteralPath $temporaryFolder -Destination $destinationFolder

            foreach ($manifestName in $knownManifestNames) {
                $installedManifest = Join-Path $destinationRoot $manifestName
                if ($manifestName -eq $stagedAddin.Name) {
                    if (Test-FileContentEqual -Left $installedManifest -Right $stagedAddin.FullName) {
                        continue
                    }
                    if (Test-Path -LiteralPath $installedManifest) {
                        Remove-Item -LiteralPath $installedManifest -Force
                    }
                    Copy-Item -LiteralPath $stagedAddin.FullName -Destination $installedManifest -Force
                    continue
                }
                if (Test-Path -LiteralPath $installedManifest) {
                    Remove-Item -LiteralPath $installedManifest -Force
                }
            }
        }
        finally {
            if (Test-Path -LiteralPath $temporaryFolder) {
                Assert-ExactInstallPath -Actual (Split-Path -Parent $temporaryFolder) -Expected $destinationRoot -Description 'temporary cleanup parent'
                Remove-Item -LiteralPath $temporaryFolder -Recurse -Force
            }
        }
        Write-Host "Installed Family Browser Revit $year -> $destinationRoot" -ForegroundColor Green
    }
}
