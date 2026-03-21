param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [Parameter(Mandatory = $true)]
    [string]$BuildRoot,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$stagingRoot = Join-Path $repoRoot "artifacts\update-package\$Version"
$outputFullPath = if ([System.IO.Path]::IsPathRooted($OutputPath)) { $OutputPath } else { Join-Path $repoRoot $OutputPath }

if (Test-Path -LiteralPath $stagingRoot) {
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $stagingRoot -Force | Out-Null

$yearMap = [ordered]@{
    '2019' = 'net48'
    '2021' = 'net48'
    '2023' = 'net48'
    '2025' = 'net8.0-windows'
}

foreach ($entry in $yearMap.GetEnumerator()) {
    $year = $entry.Key
    $tfm = $entry.Value

    $sourceDir = Join-Path $BuildRoot "Rvt$year\$tfm"
    if (-not (Test-Path -LiteralPath $sourceDir)) {
        throw "Build output not found for Revit ${year}: $sourceDir"
    }

    $yearStage = Join-Path $stagingRoot $year
    $payloadStage = Join-Path $yearStage 'KKY_Tool_Revit'
    New-Item -ItemType Directory -Path $payloadStage -Force | Out-Null

    Copy-Item -Path (Join-Path $sourceDir '*') -Destination $payloadStage -Recurse -Force

    $addinSource = Join-Path $PSScriptRoot ($year + "addin\KKY_Tool_Revit.addin")
    if (-not (Test-Path -LiteralPath $addinSource)) {
        throw "Addin file not found for Revit ${year}: $addinSource"
    }

    Copy-Item -LiteralPath $addinSource -Destination (Join-Path $yearStage 'KKY_Tool_Revit.addin') -Force
}

$manifest = [ordered]@{
    version = $Version
    createdAt = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')
    packageType = 'zip'
    years = @('2019', '2021', '2023', '2025')
}
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $stagingRoot 'package.json') -Encoding utf8

$outputDir = Split-Path -Parent $outputFullPath
if (-not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

if (Test-Path -LiteralPath $outputFullPath) {
    Remove-Item -LiteralPath $outputFullPath -Force
}

Compress-Archive -Path (Join-Path $stagingRoot '*') -DestinationPath $outputFullPath -Force
Write-Host "Generated update zip:" $outputFullPath
