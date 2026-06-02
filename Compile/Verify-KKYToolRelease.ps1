param(
    [string]$Version = '',
    [switch]$Allow2027,
    [switch]$SkipJsCheck,
    [switch]$StrictCacheVersion,
    [string]$ReportPath = '',
    [int]$BusyFileRetryCount = 5,
    [int]$BusyFileRetryDelayMs = 350
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$releaseDir = Join-Path $repoRoot 'Sever\Release'
$installerDir = Join-Path $releaseDir 'official'
$issPath = Join-Path $PSScriptRoot 'KKY_Tool_Compiler.iss'
$latestPath = Join-Path $releaseDir 'latest.json'
$historyPath = Join-Path $releaseDir 'release-history.json'
$siteIndexPath = Join-Path $releaseDir 'index.html'
$hubRoot = Join-Path $repoRoot 'KKY_Tool_Revit_2019-2023\Resources\HubUI'
$hubIndexPath = Join-Path $hubRoot 'index.html'
$hubMainPath = Join-Path $hubRoot 'js\main.js'
$topbarPath = Join-Path $hubRoot 'js\core\topbar.js'
$stableYearLabel = '2019,21,23,25,27'
$blockedYearLabel = ''

$failures = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[string]

function Add-Failure {
    param([string]$Message)
    $null = $script:failures.Add($Message)
}

function Add-Warn {
    param([string]$Message)
    $null = $script:warnings.Add($Message)
}

function Read-TextFile {
    param(
        [string]$Path,
        [switch]$Required
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        if ($Required) {
            Add-Failure "Required file not found: $Path"
        } else {
            Add-Warn "File not found, skipped: $Path"
        }
        return $null
    }

    $lastError = $null
    $attempts = [Math]::Max(1, $BusyFileRetryCount)

    foreach ($attempt in 1..$attempts) {
        $stream = $null
        $reader = $null
        try {
            $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
            return $reader.ReadToEnd()
        } catch {
            $lastError = $_
            if ($attempt -lt $attempts) {
                Start-Sleep -Milliseconds $BusyFileRetryDelayMs
            }
        } finally {
            if ($reader) {
                $reader.Close()
            } elseif ($stream) {
                $stream.Dispose()
            }
        }
    }

    $message = "Could not read file after $attempts attempt(s): $Path"
    if ($lastError) {
        $message = "$message - $($lastError.Exception.Message)"
    }

    if ($Required) {
        Add-Failure $message
    } else {
        Add-Warn $message
    }

    return $null
}

function Get-RequiredText {
    param([string]$Path)

    $text = Read-TextFile -Path $Path -Required
    if ($null -eq $text) {
        return ''
    }

    return $text
}

function Read-JsonFile {
    param([string]$Path)

    $raw = Read-TextFile -Path $Path -Required
    if ($null -eq $raw) {
        return $null
    }

    try {
        if ([string]::IsNullOrWhiteSpace($raw)) {
            Add-Failure "JSON file is empty: $Path"
            return $null
        }

        return $raw | ConvertFrom-Json
    } catch {
        Add-Failure "JSON parse failed: $Path - $($_.Exception.Message)"
        return $null
    }
}

function Get-UrlFileName {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return ''
    }

    try {
        $uri = [System.Uri]$Url
        return [System.Uri]::UnescapeDataString([System.IO.Path]::GetFileName($uri.AbsolutePath))
    } catch {
        return [System.IO.Path]::GetFileName($Url)
    }
}

function Assert-No2027ReleaseText {
    param(
        [string]$Path,
        [string]$Label
    )

    return
}

function Test-CacheVersions {
    $cacheWarnings = New-Object System.Collections.Generic.List[string]

    if (Test-Path -LiteralPath $hubIndexPath) {
        $lineNo = 0
        $hubIndexText = Read-TextFile -Path $hubIndexPath
        if ($null -ne $hubIndexText) {
            foreach ($line in ($hubIndexText -split "\r?\n")) {
                $lineNo += 1
                if ($line -match '<(?:link|script)\b' -and $line -match '(?:href|src)="(?<url>[^"]+\.(?:css|js)(?:\?v=[^"]*)?)"') {
                    $url = $Matches['url']
                    if ($url -match '^(css|js)/' -and $url -notmatch '\?v=') {
                        $null = $cacheWarnings.Add("Hub index cache version missing at line ${lineNo}: $url")
                    }
                }
            }
        }
    }

    if (Test-Path -LiteralPath $hubMainPath) {
        $lineNo = 0
        $hubMainText = Read-TextFile -Path $hubMainPath
        if ($null -ne $hubMainText) {
            foreach ($line in ($hubMainText -split "\r?\n")) {
                $lineNo += 1
                if ($line -match '^\s*import\b' -and $line -match "from\s+'(?<url>[^']+\.js(?:\?v=[^']*)?)'") {
                    $url = $Matches['url']
                    if ($url -notmatch '\?v=') {
                        $null = $cacheWarnings.Add("Hub main import cache version missing at line ${lineNo}: $url")
                    }
                }
            }
        }
    }

    foreach ($item in $cacheWarnings) {
        if ($StrictCacheVersion) {
            Add-Failure $item
        } else {
            Add-Warn $item
        }
    }
}

function Test-JsSyntax {
    if ($SkipJsCheck) {
        Add-Warn 'JS syntax check skipped by -SkipJsCheck.'
        return
    }

    $nodeCommand = Get-Command node -ErrorAction SilentlyContinue
    if (-not $nodeCommand) {
        Add-Warn 'Node.js was not found in PATH. JS syntax check skipped.'
        return
    }

    if (-not (Test-Path -LiteralPath $hubRoot)) {
        Add-Failure "Hub UI folder not found: $hubRoot"
        return
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $jsFiles = Get-ChildItem -LiteralPath $hubRoot -Recurse -File -Filter *.js |
        Where-Object { $_.Name -notmatch '\.bak-' }

    foreach ($file in $jsFiles) {
        $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("kky-js-check-" + [Guid]::NewGuid().ToString('N') + '.mjs')
        try {
            $source = Read-TextFile -Path $file.FullName
            if ($null -eq $source) {
                continue
            }
            [System.IO.File]::WriteAllText($tempPath, $source, $utf8NoBom)
            $output = & $nodeCommand.Source --check $tempPath 2>&1
            if ($LASTEXITCODE -ne 0) {
                Add-Failure "JS syntax check failed: $($file.FullName)`n$output"
            }
        } finally {
            if (Test-Path -LiteralPath $tempPath) {
                Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

$issText = Get-RequiredText -Path $issPath
$issVersionMatch = [regex]::Match($issText, '#define\s+MyAppVersion\s+"(?<value>[^"]+)"')
if (-not $issVersionMatch.Success) {
    Add-Failure "Could not parse MyAppVersion from: $issPath"
} else {
    $issVersion = $issVersionMatch.Groups['value'].Value
    if ([string]::IsNullOrWhiteSpace($Version)) {
        $Version = $issVersion
    } elseif ($Version -ne $issVersion) {
        Add-Failure "Requested version '$Version' does not match MyAppVersion '$issVersion'."
    }
}

if ([string]::IsNullOrWhiteSpace($Version)) {
    Add-Failure 'Release version could not be resolved.'
} else {
    $expectedZipName = "KKY_Tool_Revit($stableYearLabel)_v$Version.zip"
    $expectedExeName = "KKY_Tool_Revit($stableYearLabel)_v$Version.exe"

    $latest = Read-JsonFile -Path $latestPath
    if ($latest) {
        if ([string]$latest.version -ne $Version) {
            Add-Failure "latest.json version '$($latest.version)' does not match '$Version'."
        }
        if ([string]::IsNullOrWhiteSpace([string]$latest.url)) {
            Add-Failure 'latest.json url is empty.'
        } else {
            $latestFileName = Get-UrlFileName -Url ([string]$latest.url)
            if ($latestFileName -ne $expectedZipName) {
                Add-Failure "latest.json url should point to '$expectedZipName' but points to '$latestFileName'."
            }
        }
        if ([string]::IsNullOrWhiteSpace([string]$latest.publishedAt)) {
            Add-Failure 'latest.json publishedAt is empty.'
        }
        $latestInstallerFileName = Get-UrlFileName -Url ([string]$latest.installerUrl)
        if ($latestInstallerFileName -and $latestInstallerFileName -ne $expectedExeName) {
            Add-Failure "latest.json installerUrl should point to '$expectedExeName' but points to '$latestInstallerFileName'."
        }
        if ($latestInstallerFileName -and ([string]$latest.installerUrl) -notmatch '/official/') {
            Add-Failure "latest.json installerUrl should use the official installer folder."
        }
    }

    $history = Read-JsonFile -Path $historyPath
    if ($history) {
        $entries = @($history)
        if ($entries.Count -eq 0) {
            Add-Failure 'release-history.json has no entries.'
        } else {
            $first = $entries[0]
            if ([string]$first.version -ne $Version) {
                Add-Failure "release-history.json first version '$($first.version)' does not match '$Version'."
            }

            $packageFileName = Get-UrlFileName -Url ([string]$first.packageUrl)
            if ($packageFileName -and $packageFileName -ne $expectedZipName) {
                Add-Failure "release-history.json packageUrl should point to '$expectedZipName' but points to '$packageFileName'."
            }

            $installerFileName = Get-UrlFileName -Url ([string]$first.installerUrl)
            if ($installerFileName -and $installerFileName -ne $expectedExeName) {
                Add-Failure "release-history.json installerUrl should point to '$expectedExeName' but points to '$installerFileName'."
            }
            if ($installerFileName -and ([string]$first.installerUrl) -notmatch '/official/') {
                Add-Failure "release-history.json installerUrl should use the official installer folder."
            }
        }
    }

    $expectedZipPath = Join-Path $releaseDir $expectedZipName
    if (-not (Test-Path -LiteralPath $expectedZipPath)) {
        Add-Warn "Release artifact not found locally: $expectedZipPath"
    }

    $expectedExePath = Join-Path $installerDir $expectedExeName
    if (-not (Test-Path -LiteralPath $expectedExePath)) {
        Add-Warn "Release artifact not found locally: $expectedExePath"
    }

    $topbarText = Get-RequiredText -Path $topbarPath
    $topbarVersionMatch = [regex]::Match($topbarText, "APP_VERSION_FALLBACK\s*=\s*'v?(?<value>[^']+)';")
    if (-not $topbarVersionMatch.Success) {
        Add-Failure "Could not parse APP_VERSION_FALLBACK from: $topbarPath"
    } elseif ($topbarVersionMatch.Groups['value'].Value -ne $Version) {
        Add-Failure "APP_VERSION_FALLBACK '$($topbarVersionMatch.Groups['value'].Value)' does not match '$Version'."
    }
}

Assert-No2027ReleaseText -Path $latestPath -Label 'latest.json'
Assert-No2027ReleaseText -Path $historyPath -Label 'release-history.json'
Assert-No2027ReleaseText -Path $siteIndexPath -Label 'release homepage'

Test-CacheVersions
Test-JsSyntax

if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
    try {
        $reportFullPath = if ([System.IO.Path]::IsPathRooted($ReportPath)) {
            $ReportPath
        } else {
            Join-Path $repoRoot $ReportPath
        }
        $reportDir = Split-Path -Parent $reportFullPath
        if (-not [string]::IsNullOrWhiteSpace($reportDir) -and -not (Test-Path -LiteralPath $reportDir)) {
            New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
        }

        $status = if ($failures.Count -gt 0) { 'failed' } else { 'passed' }
        $payload = [ordered]@{
            generatedAt = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')
            status = $status
            version = $Version
            allow2027 = $Allow2027.IsPresent
            strictCacheVersion = $StrictCacheVersion.IsPresent
            skipJsCheck = $SkipJsCheck.IsPresent
            busyFileRetryCount = $BusyFileRetryCount
            busyFileRetryDelayMs = $BusyFileRetryDelayMs
            releaseDir = $releaseDir
            warnings = @($warnings)
            failures = @($failures)
        }

        $json = $payload | ConvertTo-Json -Depth 6
        [System.IO.File]::WriteAllText($reportFullPath, $json, [System.Text.Encoding]::UTF8)
        Write-Host "Verification report: $reportFullPath"
    } catch {
        Add-Warn "Could not write verification report: $($_.Exception.Message)"
    }
}

Write-Host ''
Write-Host 'KKY Tool release verification'
Write-Host "Version     : $Version"
Write-Host "Allow 2027  : $($Allow2027.IsPresent)"
Write-Host "Release dir : $releaseDir"

if ($warnings.Count -gt 0) {
    Write-Host ''
    Write-Host 'Warnings:'
    foreach ($warning in $warnings) {
        Write-Host " - $warning"
    }
}

if ($failures.Count -gt 0) {
    Write-Host ''
    Write-Host 'Failures:'
    foreach ($failure in $failures) {
        Write-Host " - $failure"
    }
    exit 1
}

Write-Host ''
Write-Host 'Verification passed.'
exit 0
