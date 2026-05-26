param(
    [string]$Version = '2.25',
    [string]$TestLabel = 'test',
    [switch]$SkipBackup
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$backupScriptPath = Join-Path $PSScriptRoot 'Backup-TestWorkingSet.ps1'
$issPath = Join-Path $PSScriptRoot 'KKY_Tool_Compiler.iss'
$releaseDir = Join-Path $repoRoot 'Sever\Release'
$testInstallerDir = Join-Path $releaseDir 'test'
$stageRoot = Join-Path $repoRoot ("artifacts\release-stage\test-{0}_{1}" -f $Version, $TestLabel)
$proj2019To2023 = Join-Path $repoRoot 'KKY_Tool_Revit_2019-2023\KKY_Tool_Revit.vbproj'
$proj2025 = Join-Path $repoRoot 'KKY_Tool_Revit_2025\KKY_Tool_Revit_2025.vbproj'
$proj2027 = Join-Path $repoRoot 'KKY_Tool_Revit_2027\KKY_Tool_Revit_2027.vbproj'
$isccPath = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe'
$outputBaseName = "KKY_Tool_Revit(2019,21,23,25,27)_v{0}_{1}" -f $Version, $TestLabel
$outputExePath = Join-Path $testInstallerDir ($outputBaseName + '.exe')

if (-not (Test-Path -LiteralPath $backupScriptPath)) {
    throw "Backup script not found: $backupScriptPath"
}

if (-not (Test-Path -LiteralPath $issPath)) {
    throw "Inno Setup script not found: $issPath"
}

if (-not (Test-Path -LiteralPath $isccPath)) {
    throw "ISCC.exe not found: $isccPath"
}

$backupPath = ''
if (-not $SkipBackup) {
    $backupPath = & $backupScriptPath -Label ("{0}_{1}" -f $Version, $TestLabel) -Version $Version | Select-Object -Last 1
    if ($LASTEXITCODE -ne 0) {
        throw 'Modified file backup failed before test build.'
    }
}

if (Test-Path -LiteralPath $stageRoot) {
    Remove-Item -LiteralPath $stageRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $stageRoot -Force | Out-Null
New-Item -ItemType Directory -Path $testInstallerDir -Force | Out-Null

foreach ($year in '2019', '2021', '2023') {
    $outputPath = Join-Path $stageRoot ("Rvt{0}\net48\" -f $year)
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

$outputPath2027 = Join-Path $stageRoot 'Rvt2027\net10.0-windows\'
& dotnet build $proj2027 -c Release -p:SkipCreateAddin=true -p:OutputPath=$outputPath2027
if ($LASTEXITCODE -ne 0) {
    throw 'Build failed for Revit 2027'
}

& $isccPath "/DMyAppVersion=$Version" "/DMyOutputBaseName=$outputBaseName" "/DMyBuildRoot=$stageRoot" "/DMyOutputDir=$testInstallerDir" $issPath
if ($LASTEXITCODE -ne 0) {
    throw 'Installer compile failed.'
}

if (-not (Test-Path -LiteralPath $outputExePath)) {
    throw "Compiled test installer not found: $outputExePath"
}

Write-Host ''
Write-Host 'Test build completed.'
Write-Host "Version   : $Version"
Write-Host "Label     : $TestLabel"
if (-not [string]::IsNullOrWhiteSpace($backupPath)) {
    Write-Host "Backup    : $backupPath"
}
Write-Host "Installer : $outputExePath"
