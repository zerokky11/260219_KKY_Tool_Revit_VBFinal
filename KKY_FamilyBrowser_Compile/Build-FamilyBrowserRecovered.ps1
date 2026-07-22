param(
    [string[]]$Years = @('2019', '2021', '2023', '2025', '2027'),
    [string]$Configuration = 'Release',
    [switch]$SkipBuild
)

$ErrorActionPreference = 'Stop'

function Assert-SafeChildPath {
    param(
        [Parameter(Mandatory = $true)][string]$Parent,
        [Parameter(Mandatory = $true)][string]$Child
    )

    $parentFull = [System.IO.Path]::GetFullPath($Parent).TrimEnd('\') + '\'
    $childFull = [System.IO.Path]::GetFullPath($Child)
    if (-not $childFull.StartsWith($parentFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsafe path outside workspace: $childFull"
    }
}

function Invoke-DotNetBuild {
    param([Parameter(Mandatory = $true)][string]$ProjectPath)

    Write-Host "Building $ProjectPath" -ForegroundColor Cyan
    & dotnet build $ProjectPath -c $Configuration -v:minimal
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet build failed: $ProjectPath"
    }
}

function Copy-HostOutput {
    param(
        [Parameter(Mandatory = $true)][string]$OutputDirectory,
        [Parameter(Mandatory = $true)][string]$AssemblyName,
        [Parameter(Mandatory = $true)][string]$DestinationDirectory
    )

    New-Item -ItemType Directory -Path $DestinationDirectory -Force | Out-Null

    $requiredDll = Join-Path $OutputDirectory "$AssemblyName.dll"
    if (-not (Test-Path -LiteralPath $requiredDll)) {
        throw "Missing build output: $requiredDll"
    }

    foreach ($extension in @('dll', 'pdb', 'deps.json')) {
        $source = Join-Path $OutputDirectory "$AssemblyName.$extension"
        if (Test-Path -LiteralPath $source) {
            Copy-Item -LiteralPath $source -Destination $DestinationDirectory -Force
        }
    }
}

function Write-AddinManifest {
    param(
        [Parameter(Mandatory = $true)][string]$Year,
        [Parameter(Mandatory = $true)][string]$AddinPath,
        [Parameter(Mandatory = $true)][string]$AddinName,
        [Parameter(Mandatory = $true)][string]$AssemblyFileName,
        [Parameter(Mandatory = $true)][string]$AddInId,
        [Parameter(Mandatory = $true)][string]$FullClassName
    )

    $assemblyPath = "C:\ProgramData\Autodesk\Revit\Addins\$Year\KKY_FamilyBrowser\$AssemblyFileName"
    $xml = @"
<?xml version="1.0" encoding="utf-8" standalone="no"?>
<RevitAddIns>
  <AddIn Type="Application">
    <Name>$AddinName</Name>
    <Assembly>$assemblyPath</Assembly>
    <AddInId>$AddInId</AddInId>
    <FullClassName>$FullClassName</FullClassName>
    <VendorId>KKY</VendorId>
    <VendorDescription>KKY Family Browser recovered host</VendorDescription>
  </AddIn>
</RevitAddIns>
"@
    Set-Content -LiteralPath $AddinPath -Value $xml -Encoding UTF8
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
$stageRoot = Join-Path $repoRoot 'artifacts\family-browser\stage'

Assert-SafeChildPath -Parent $repoRoot -Child $stageRoot

$validYears = @('2019', '2021', '2023', '2025', '2027')
$selectedYears = @($Years | ForEach-Object { @(([string]$_) -split ',') } | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
foreach ($year in $selectedYears) {
    if ($validYears -notcontains $year) {
        throw "Unsupported Revit year: $year"
    }
}

$project2019 = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2019-2023\KKY_FamilyBrowser_RevitHost.csproj'
$project2025 = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\KKY_FamilyBrowser_RevitHost_2025.csproj'
$project2027 = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2027\KKY_FamilyBrowser_RevitHost_2027.csproj'

if (-not $SkipBuild) {
    if (@($selectedYears | Where-Object { @('2019', '2021', '2023') -contains $_ }).Count -gt 0) {
        Invoke-DotNetBuild -ProjectPath $project2019
    }
    if ($selectedYears -contains '2025') {
        Invoke-DotNetBuild -ProjectPath $project2025
    }
    if ($selectedYears -contains '2027') {
        Invoke-DotNetBuild -ProjectPath $project2027
    }
}

if (Test-Path -LiteralPath $stageRoot) {
    Remove-Item -LiteralPath $stageRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $stageRoot -Force | Out-Null

$out2019 = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2019-2023\bin\Release\net48'
$out2025 = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\bin\Release\net8.0-windows'
$out2027 = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2027\bin\Release\net10.0-windows'

foreach ($year in $selectedYears) {
    $yearStage = Join-Path $stageRoot "Rvt$year"
    $addinFolder = Join-Path $yearStage 'KKY_FamilyBrowser'
    New-Item -ItemType Directory -Path $yearStage -Force | Out-Null

    switch ($year) {
        { @('2019', '2021', '2023') -contains $_ } {
            Copy-HostOutput -OutputDirectory $out2019 -AssemblyName 'KKY_FamilyBrowser_RevitHost' -DestinationDirectory $addinFolder
            Write-AddinManifest `
                -Year $year `
                -AddinPath (Join-Path $yearStage 'KKY_FamilyBrowser_RevitHost.addin') `
                -AddinName 'KKY_FamilyBrowser_RevitHost' `
                -AssemblyFileName 'KKY_FamilyBrowser_RevitHost.dll' `
                -AddInId '7D1E58FC-5B4C-43A1-9121-6D7B6D640201' `
                -FullClassName 'KKY_FamilyBrowser_RevitHost_2019_2023.App'
            break
        }
        '2025' {
            Copy-HostOutput -OutputDirectory $out2025 -AssemblyName 'KKY_FamilyBrowser_RevitHost_2025' -DestinationDirectory $addinFolder
            Write-AddinManifest `
                -Year $year `
                -AddinPath (Join-Path $yearStage 'KKY_FamilyBrowser_RevitHost_2025.addin') `
                -AddinName 'KKY_FamilyBrowser_RevitHost_2025' `
                -AssemblyFileName 'KKY_FamilyBrowser_RevitHost_2025.dll' `
                -AddInId '9E6E4A65-38B8-4EE7-9A1D-0D621E335F41' `
                -FullClassName 'KKY_FamilyBrowser_RevitHost_2025.App'
            break
        }
        '2027' {
            Copy-HostOutput -OutputDirectory $out2027 -AssemblyName 'KKY_FamilyBrowser_RevitHost_2027' -DestinationDirectory $addinFolder
            Write-AddinManifest `
                -Year $year `
                -AddinPath (Join-Path $yearStage 'KKY_FamilyBrowser_RevitHost_2027.addin') `
                -AddinName 'KKY_FamilyBrowser_RevitHost_2027' `
                -AssemblyFileName 'KKY_FamilyBrowser_RevitHost_2027.dll' `
                -AddInId '8A0D0C84-8E3C-4A1B-B7D8-00C2CECB2027' `
                -FullClassName 'KKY_FamilyBrowser_RevitHost_2027.App'
            break
        }
    }
}

$payloadFiles = @(
    Get-ChildItem -Recurse -File -LiteralPath $stageRoot |
        Sort-Object FullName |
        ForEach-Object {
            [ordered]@{
                relativePath = $_.FullName.Substring($stageRoot.Length).TrimStart('\').Replace('\', '/')
                bytes = $_.Length
                sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
            }
        }
)

$manifest = [ordered]@{
    generatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss K')
    configuration = $Configuration
    years = $selectedYears
    stageRoot = $stageRoot
    payloadFileCount = $payloadFiles.Count
    payload = $payloadFiles
}
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $stageRoot 'stage-manifest.json') -Encoding UTF8

Write-Host ""
Write-Host "Family Browser stage created:" -ForegroundColor Green
Write-Host $stageRoot
Get-ChildItem -Recurse -File -LiteralPath $stageRoot |
    Select-Object FullName, Length |
    Format-Table -AutoSize
