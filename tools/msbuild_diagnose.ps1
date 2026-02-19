param(
    [string]$TargetPath,
    [string]$OutDir = "artifacts/msbuild-diagnose"
)

$ErrorActionPreference = "Continue"

function Resolve-RepoRoot {
    if ($PSScriptRoot) {
        return (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
    }
    return (Get-Location).Path
}

function Get-MsbuildPath {
    $cmd = Get-Command msbuild -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Path }

    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $installPath = & $vswhere -latest -requires Microsoft.Component.MSBuild -property installationPath
        if ($LASTEXITCODE -eq 0 -and $installPath) {
            $candidate = Join-Path $installPath "MSBuild\Current\Bin\MSBuild.exe"
            if (Test-Path $candidate) { return $candidate }
        }
    }

    throw "MSBuild not found. Install Visual Studio Build Tools or use Developer PowerShell."
}

function Invoke-MsbuildStep {
    param(
        [string]$MsbuildExe,
        [string]$Input,
        [string]$Label,
        [string]$PPPath,
        [string]$BinlogPath
    )

    Write-Host "\n=== [$Label] PREPROCESS (/pp) ==="
    Write-Host "Input : $Input"
    Write-Host "Output: $PPPath"
    try {
        & $MsbuildExe $Input "/nologo" "/pp:$PPPath"
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "PREPROCESS failed (exit=$LASTEXITCODE)."
        }
    }
    catch {
        Write-Warning "PREPROCESS threw an exception: $($_.Exception.Message)"
    }

    Write-Host "\n=== [$Label] BUILD BINLOG (/bl) ==="
    Write-Host "Input : $Input"
    Write-Host "Output: $BinlogPath"
    try {
        & $MsbuildExe $Input "/nologo" "/restore" "/t:Build" "/m:1" "/v:diag" "/bl:$BinlogPath"
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "BUILD failed (exit=$LASTEXITCODE). Binlog should still be generated if evaluation reached logger init."
        }
    }
    catch {
        Write-Warning "BUILD threw an exception: $($_.Exception.Message)"
    }
}

$repoRoot = Resolve-RepoRoot
Set-Location $repoRoot

$msbuild = Get-MsbuildPath
Write-Host "Using MSBuild: $msbuild"

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$runDir = Join-Path $repoRoot (Join-Path $OutDir $timestamp)
New-Item -ItemType Directory -Force -Path $runDir | Out-Null

$defaultTargets = @(
    "KKY_Tool_Revit_2019-2023/KKY_Tool_Revit.vbproj",
    "KKY_Tool_Revit_2025/KKY_Tool_Revit_2025.vbproj"
)

$targets = @()
if ($TargetPath) {
    $targets += $TargetPath
}
else {
    $targets += $defaultTargets
}

foreach ($t in $targets) {
    $resolved = $t
    if (-not [System.IO.Path]::IsPathRooted($resolved)) {
        $resolved = Join-Path $repoRoot $resolved
    }

    if (-not (Test-Path $resolved)) {
        Write-Warning "Target not found: $t"
        continue
    }

    $name = [System.IO.Path]::GetFileNameWithoutExtension($resolved)
    $pp = Join-Path $runDir ("$name.preprocessed.xml")
    $bl = Join-Path $runDir ("$name.build.binlog")

    Invoke-MsbuildStep -MsbuildExe $msbuild -Input $resolved -Label $name -PPPath $pp -BinlogPath $bl
}

Write-Host "\n=== Diagnose output directory ==="
Write-Host $runDir
Write-Host "\nOpen *.build.binlog with MSBuild Structured Log Viewer and inspect the Imports tree to find the exact file path that re-imports SDK targets and causes circular dependency."
