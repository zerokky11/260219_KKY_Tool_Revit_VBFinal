param(
    [string]$Notes = 'GitHub Pages release build'
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$issPath = Join-Path $PSScriptRoot 'KKY_Tool_Compiler.iss'
$feedScriptPath = Join-Path $PSScriptRoot 'New-UpdateFeed.ps1'
$releaseDir = Join-Path $repoRoot 'Sever\Release'
$indexPath = Join-Path $releaseDir 'index.html'
$feedPath = Join-Path $releaseDir 'latest.json'
$domainRoot = 'https://update.zerokky.com'
$isccPath = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe'

if (-not (Test-Path -LiteralPath $issPath)) {
    throw "Inno Setup script not found: $issPath"
}

if (-not (Test-Path -LiteralPath $feedScriptPath)) {
    throw "Feed generator script not found: $feedScriptPath"
}

if (-not (Test-Path -LiteralPath $isccPath)) {
    throw "ISCC.exe not found: $isccPath"
}

$issContent = Get-Content -Raw -LiteralPath $issPath
$versionMatch = [regex]::Match($issContent, '#define\s+MyAppVersion\s+"(?<value>[^"]+)"')
if (-not $versionMatch.Success) {
    throw 'Could not parse MyAppVersion from KKY_Tool_Compiler.iss'
}

$version = $versionMatch.Groups['value'].Value
$exeName = "KKY_Tool_Revit(2019,21,23,25)_v$version.exe"
$downloadUrl = "$domainRoot/$exeName"

& $isccPath $issPath

$exePath = Join-Path $releaseDir $exeName
if (-not (Test-Path -LiteralPath $exePath)) {
    throw "Compiled installer not found: $exePath"
}

& $feedScriptPath `
    -Version $version `
    -DownloadUrl $downloadUrl `
    -PublishedAt (Get-Date -Format 'yyyy-MM-dd') `
    -Notes $Notes `
    -OutputPath $feedPath

if (Test-Path -LiteralPath $indexPath) {
    $indexContent = Get-Content -Raw -LiteralPath $indexPath
    $indexContent = [regex]::Replace(
        $indexContent,
        'KKY_Tool_Revit\(2019,21,23,25\)_v[0-9.]+\.exe',
        [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $exeName },
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $indexContent = [regex]::Replace(
        $indexContent,
        'https://update\.zerokky\.com/KKY_Tool_Revit\(2019,21,23,25\)_v[0-9.]+\.exe',
        [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $downloadUrl },
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    [System.IO.File]::WriteAllText($indexPath, $indexContent, [System.Text.Encoding]::UTF8)
}

Write-Host ''
Write-Host 'Release build completed.'
Write-Host "Version      : $version"
Write-Host "Installer    : $exePath"
Write-Host "Feed file    : $feedPath"
Write-Host "Download URL : $downloadUrl"
