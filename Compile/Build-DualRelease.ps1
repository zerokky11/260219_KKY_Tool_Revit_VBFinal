param(
    [string]$FirstVersion = '2.06',
    [string]$SecondVersion = '2.07'
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$assemblyInfoPath = Join-Path $repoRoot 'KKY_Tool_Revit_2019-2023\My Project\AssemblyInfo.vb'
$topbarPath = Join-Path $repoRoot 'KKY_Tool_Revit_2019-2023\Resources\HubUI\js\core\topbar.js'
$issPath = Join-Path $repoRoot 'Compile\KKY_Tool_Compiler.iss'
$feedScriptPath = Join-Path $repoRoot 'Compile\New-UpdateFeed.ps1'
$zipScriptPath = Join-Path $repoRoot 'Compile\New-UpdateZip.ps1'
$releaseDir = Join-Path $repoRoot 'Sever\Release'
$stageRoot = Join-Path $repoRoot 'artifacts\release-stage'
$proj2019To2023 = Join-Path $repoRoot 'KKY_Tool_Revit_2019-2023\KKY_Tool_Revit.vbproj'
$proj2025 = Join-Path $repoRoot 'KKY_Tool_Revit_2025\KKY_Tool_Revit_2025.vbproj'
$indexPath = Join-Path $releaseDir 'index.html'
$domainRoot = 'https://update.zerokky.com'
$isccPath = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe'

$assemblyOriginal = Get-Content -Raw -LiteralPath $assemblyInfoPath
$topbarOriginal = Get-Content -Raw -LiteralPath $topbarPath
$issOriginal = Get-Content -Raw -LiteralPath $issPath

function Set-FileContentWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Value,

        [switch]$AllowFailure
    )

    $lastError = $null
    foreach ($attempt in 1..5) {
        try {
            Set-Content -LiteralPath $Path -Value $Value -Encoding utf8
            return $true
        } catch {
            $lastError = $_
            Start-Sleep -Milliseconds 400
        }
    }

    if ($AllowFailure) {
        Write-Warning "Skipped updating locked file: $Path"
        if ($lastError) {
            Write-Warning $lastError.Exception.Message
        }
        return $false
    }

    throw $lastError
}

function Set-VersionFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Version
    )

    $parts = $Version.Split('.')
    $major = [int]$parts[0]
    $minor = [int]$parts[1]
    $assemblyVersion = "$major.$minor.0.0"

    $assemblyText = $assemblyOriginal
    $assemblyText = [regex]::Replace($assemblyText, 'AssemblyVersion\("[^"]+"\)', "AssemblyVersion(""$assemblyVersion"")")
    $assemblyText = [regex]::Replace($assemblyText, 'AssemblyFileVersion\("[^"]+"\)', "AssemblyFileVersion(""$assemblyVersion"")")
    $assemblyText = [regex]::Replace($assemblyText, 'AssemblyInformationalVersion\("[^"]+"\)', "AssemblyInformationalVersion(""$Version"")")
    Set-FileContentWithRetry -Path $assemblyInfoPath -Value $assemblyText | Out-Null

    if (-not [string]::IsNullOrWhiteSpace($topbarOriginal)) {
        $topbarText = $topbarOriginal
        $topbarText = [regex]::Replace($topbarText, "const APP_VERSION_FALLBACK = 'v[0-9.]+';", "const APP_VERSION_FALLBACK = 'v$Version';")
        Set-FileContentWithRetry -Path $topbarPath -Value $topbarText -AllowFailure | Out-Null
    }

    $issText = $issOriginal
    $issText = [regex]::Replace($issText, '#define MyAppVersion "[^"]+"', "#define MyAppVersion ""$Version""")
    Set-FileContentWithRetry -Path $issPath -Value $issText | Out-Null
}

function Build-StagedOutputs {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Version
    )

    $stageDir = Join-Path $stageRoot $Version
    if (Test-Path -LiteralPath $stageDir) {
        Remove-Item -LiteralPath $stageDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $stageDir -Force | Out-Null

    foreach ($year in '2019', '2021', '2023') {
        $outputPath = Join-Path $stageDir "Rvt$year\net48\"
        & dotnet build $proj2019To2023 -c Release -p:SkipDeployAllYears=true -p:AddinYear=$year -p:OutputPath=$outputPath
        if ($LASTEXITCODE -ne 0) {
            throw "Build failed for version $Version / Revit $year"
        }
    }

    $outputPath2025 = Join-Path $stageDir 'Rvt2025\net8.0-windows\'
    & dotnet build $proj2025 -c Release -p:SkipCreateAddin=true -p:OutputPath=$outputPath2025
    if ($LASTEXITCODE -ne 0) {
        throw "Build failed for version $Version / Revit 2025"
    }

    & $isccPath "/DMyAppVersion=$Version" "/DMyBuildRoot=$stageDir" "/DMyOutputDir=$releaseDir" $issPath
    if ($LASTEXITCODE -ne 0) {
        throw "Installer compile failed for version $Version"
    }

    & $zipScriptPath `
        -Version $Version `
        -BuildRoot $stageDir `
        -OutputPath (Join-Path $releaseDir "KKY_Tool_Revit(2019,21,23,25)_v$Version.zip")

    if ($LASTEXITCODE -ne 0) {
        throw "Zip package build failed for version $Version"
    }
}

Set-VersionFiles -Version $FirstVersion
Build-StagedOutputs -Version $FirstVersion

Set-VersionFiles -Version $SecondVersion
Build-StagedOutputs -Version $SecondVersion

$finalZipName = "KKY_Tool_Revit(2019,21,23,25)_v$SecondVersion.zip"
$finalExeName = "KKY_Tool_Revit(2019,21,23,25)_v$SecondVersion.exe"
$finalDownloadUrl = "$domainRoot/$finalZipName"

& $feedScriptPath `
    -Version $SecondVersion `
    -DownloadUrl $finalDownloadUrl `
    -PublishedAt (Get-Date -Format 'yyyy-MM-dd') `
    -Notes "Release build v$SecondVersion" `
    -OutputPath (Join-Path $releaseDir 'latest.json')

if (Test-Path -LiteralPath $indexPath) {
    $indexText = Get-Content -Raw -LiteralPath $indexPath
    $indexText = [regex]::Replace($indexText, 'KKY_Tool_Revit\(2019,21,23,25\)_v[0-9.]+\.exe', $finalExeName)
    $indexText = [regex]::Replace($indexText, 'KKY_Tool_Revit\(2019,21,23,25\)_v[0-9.]+\.zip', $finalZipName)
    $indexText = [regex]::Replace($indexText, 'https://update\.zerokky\.com/KKY_Tool_Revit\(2019,21,23,25\)_v[0-9.]+\.exe', "$domainRoot/$finalExeName")
    $indexText = [regex]::Replace($indexText, 'https://update\.zerokky\.com/KKY_Tool_Revit\(2019,21,23,25\)_v[0-9.]+\.zip', $finalDownloadUrl)
    Set-Content -LiteralPath $indexPath -Value $indexText -Encoding utf8
}

Write-Host ''
Write-Host 'Dual release build completed.'
Write-Host (Join-Path $releaseDir "KKY_Tool_Revit(2019,21,23,25)_v$FirstVersion.exe")
Write-Host (Join-Path $releaseDir "KKY_Tool_Revit(2019,21,23,25)_v$SecondVersion.exe")
Write-Host (Join-Path $releaseDir "KKY_Tool_Revit(2019,21,23,25)_v$FirstVersion.zip")
Write-Host (Join-Path $releaseDir "KKY_Tool_Revit(2019,21,23,25)_v$SecondVersion.zip")
Write-Host (Join-Path $releaseDir 'latest.json')
