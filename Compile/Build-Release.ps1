param(
    [string]$Notes = 'GitHub Pages release build'
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$issPath = Join-Path $PSScriptRoot 'KKY_Tool_Compiler.iss'
$feedScriptPath = Join-Path $PSScriptRoot 'New-UpdateFeed.ps1'
$zipScriptPath = Join-Path $PSScriptRoot 'New-UpdateZip.ps1'
$historyScriptPath = Join-Path $PSScriptRoot 'Update-ReleaseHistory.ps1'
$releaseDir = Join-Path $repoRoot 'Sever\Release'
$stageRoot = Join-Path $repoRoot 'artifacts\release-stage\single'
$indexPath = Join-Path $releaseDir 'index.html'
$feedPath = Join-Path $releaseDir 'latest.json'
$historyPath = Join-Path $releaseDir 'release-history.json'
$domainRoot = 'https://update.zerokky.com'
$isccPath = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe'
$proj2019To2023 = Join-Path $repoRoot 'KKY_Tool_Revit_2019-2023\KKY_Tool_Revit.vbproj'
$proj2025 = Join-Path $repoRoot 'KKY_Tool_Revit_2025\KKY_Tool_Revit_2025.vbproj'

if (-not (Test-Path -LiteralPath $issPath)) {
    throw "Inno Setup script not found: $issPath"
}

if (-not (Test-Path -LiteralPath $feedScriptPath)) {
    throw "Feed generator script not found: $feedScriptPath"
}

if (-not (Test-Path -LiteralPath $zipScriptPath)) {
    throw "Zip generator script not found: $zipScriptPath"
}

if (-not (Test-Path -LiteralPath $historyScriptPath)) {
    throw "Release history script not found: $historyScriptPath"
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
$zipName = "KKY_Tool_Revit(2019,21,23,25)_v$version.zip"
$packageUrl = "$domainRoot/$zipName"
$installerUrl = "$domainRoot/$exeName"
if (Test-Path -LiteralPath $stageRoot) {
    Remove-Item -LiteralPath $stageRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $stageRoot -Force | Out-Null

foreach ($year in '2019', '2021', '2023') {
    $outputPath = Join-Path $stageRoot "Rvt$year\net48\"
    & dotnet build $proj2019To2023 -c Release -p:SkipDeployAllYears=true -p:AddinYear=$year -p:OutputPath=$outputPath
    if ($LASTEXITCODE -ne 0) {
        throw "Build failed for Revit $year"
    }
}

$outputPath2025 = Join-Path $stageRoot 'Rvt2025\net8.0-windows\'
& dotnet build $proj2025 -c Release -p:SkipCreateAddin=true -p:OutputPath=$outputPath2025
if ($LASTEXITCODE -ne 0) {
    throw 'Build failed for Revit 2025'
}

& $isccPath "/DMyAppVersion=$version" "/DMyBuildRoot=$stageRoot" "/DMyOutputDir=$releaseDir" $issPath
if ($LASTEXITCODE -ne 0) {
    throw 'Installer compile failed.'
}

$exePath = Join-Path $releaseDir $exeName
if (-not (Test-Path -LiteralPath $exePath)) {
    throw "Compiled installer not found: $exePath"
}

$zipPath = Join-Path $releaseDir $zipName
& $zipScriptPath `
    -Version $version `
    -BuildRoot $stageRoot `
    -OutputPath $zipPath

if (-not (Test-Path -LiteralPath $zipPath)) {
    throw "Compiled update zip not found: $zipPath"
}

& $feedScriptPath `
    -Version $version `
    -DownloadUrl $packageUrl `
    -PublishedAt (Get-Date -Format 'yyyy-MM-dd') `
    -Notes $Notes `
    -OutputPath $feedPath

& $historyScriptPath `
    -Version $version `
    -PublishedAt (Get-Date -Format 'yyyy-MM-dd') `
    -Notes $Notes `
    -PackageUrl $packageUrl `
    -InstallerUrl $installerUrl `
    -OutputPath $historyPath

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
        'KKY_Tool_Revit\(2019,21,23,25\)_v[0-9.]+\.zip',
        [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $zipName },
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $indexContent = [regex]::Replace(
        $indexContent,
        'https://update\.zerokky\.com/KKY_Tool_Revit\(2019,21,23,25\)_v[0-9.]+\.exe',
        [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $installerUrl },
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $indexContent = [regex]::Replace(
        $indexContent,
        'https://update\.zerokky\.com/KKY_Tool_Revit\(2019,21,23,25\)_v[0-9.]+\.zip',
        [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $packageUrl },
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    [System.IO.File]::WriteAllText($indexPath, $indexContent, [System.Text.Encoding]::UTF8)
}

Write-Host ''
Write-Host 'Release build completed.'
Write-Host "Version      : $version"
Write-Host "Installer    : $exePath"
Write-Host "Update zip   : $zipPath"
Write-Host "Feed file    : $feedPath"
Write-Host "Package URL  : $packageUrl"
