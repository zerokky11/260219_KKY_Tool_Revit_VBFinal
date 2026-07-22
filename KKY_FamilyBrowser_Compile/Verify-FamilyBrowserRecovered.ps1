param(
    [string[]]$Years = @('2019', '2021', '2023', '2025', '2027'),
    [string]$StageRoot,
    [switch]$Installed
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
if (-not $StageRoot) {
    $StageRoot = Join-Path $repoRoot 'artifacts\family-browser\stage'
}
$StageRoot = [System.IO.Path]::GetFullPath($StageRoot).TrimEnd('\')

$results = New-Object System.Collections.Generic.List[object]

function Get-RelativePayloadPath {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $rootFull = [System.IO.Path]::GetFullPath($Root).TrimEnd('\') + '\'
    $pathFull = [System.IO.Path]::GetFullPath($Path)
    if (-not $pathFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Payload path is outside its root: $pathFull"
    }
    return $pathFull.Substring($rootFull.Length).Replace('\', '/')
}

function Assert-StageManifestIntegrity {
    param([Parameter(Mandatory = $true)][string]$Root)

    $manifestPath = Join-Path $Root 'stage-manifest.json'
    $manifestPathFull = [System.IO.Path]::GetFullPath($manifestPath)
    if (-not (Test-Path -LiteralPath $manifestPath)) {
        throw "Stage manifest is missing: $manifestPath"
    }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    $manifestEntries = @($manifest.payload)
    if ($manifestEntries.Count -eq 0) {
        throw 'Stage manifest has no payload hashes.'
    }

    $actualFiles = @(Get-ChildItem -LiteralPath $Root -Recurse -File | Where-Object {
        -not [string]::Equals([System.IO.Path]::GetFullPath($_.FullName), $manifestPathFull, [System.StringComparison]::OrdinalIgnoreCase)
    })
    $actualPaths = @($actualFiles | ForEach-Object { Get-RelativePayloadPath -Root $Root -Path $_.FullName })
    if ($actualFiles.Count -ne [int]$manifest.payloadFileCount -or $actualFiles.Count -ne $manifestEntries.Count) {
        throw "Stage payload count mismatch. manifest=$($manifestEntries.Count), actual=$($actualFiles.Count)"
    }

    foreach ($entry in $manifestEntries) {
        $relativePath = ([string]$entry.relativePath).Replace('/', '\')
        $path = Join-Path $Root $relativePath
        if (-not (Test-Path -LiteralPath $path)) {
            throw "Stage payload file is missing: $relativePath"
        }
        $file = Get-Item -LiteralPath $path
        if ($file.Length -ne [int64]$entry.bytes) {
            throw "Stage payload length mismatch: $relativePath"
        }
        $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        if ($hash -ne [string]$entry.sha256) {
            throw "Stage payload hash mismatch: $relativePath"
        }
    }

    $manifestPaths = @($manifestEntries | ForEach-Object { ([string]$_.relativePath).Replace('\', '/') })
    $extraPaths = @($actualPaths | Where-Object { $manifestPaths -notcontains $_ })
    if ($extraPaths.Count -gt 0) {
        throw "Stage contains files outside its manifest: $($extraPaths -join ', ')"
    }
}

Assert-StageManifestIntegrity -Root $StageRoot

foreach ($year in $Years) {
    if (@('2019', '2021', '2023') -contains $year) {
        $expectedAddinName = 'KKY_FamilyBrowser_RevitHost.addin'
    }
    elseif ($year -eq '2025') {
        $expectedAddinName = 'KKY_FamilyBrowser_RevitHost_2025.addin'
    }
    elseif ($year -eq '2027') {
        $expectedAddinName = 'KKY_FamilyBrowser_RevitHost_2027.addin'
    }
    else {
        throw "Unsupported Revit year: $year"
    }

    if ($Installed) {
        $yearRoot = "C:\ProgramData\Autodesk\Revit\Addins\$year"
    }
    else {
        $yearRoot = Join-Path $StageRoot "Rvt$year"
    }

    $addinPath = Join-Path $yearRoot $expectedAddinName
    $addinFile = Get-Item -LiteralPath $addinPath -ErrorAction SilentlyContinue
    $status = 'OK'
    $details = New-Object System.Collections.Generic.List[string]

    if (-not $addinFile) {
        $status = 'FAIL'
        $details.Add('addin manifest missing')
        $results.Add([pscustomobject]@{ Year = $year; Scope = $(if ($Installed) { 'Installed' } else { 'Stage' }); Status = $status; Details = ($details -join '; ') })
        continue
    }

    try {
        [xml]$addinXml = Get-Content -LiteralPath $addinFile.FullName -Raw
        $node = $addinXml.RevitAddIns.AddIn
        $assemblyPath = [string]$node.Assembly
        $assemblyFile = Split-Path -Leaf $assemblyPath
        $className = [string]$node.FullClassName
        if (-not $assemblyFile.EndsWith('.dll', [System.StringComparison]::OrdinalIgnoreCase)) {
            $status = 'FAIL'
            $details.Add('assembly path is not a dll')
        }
        if ([string]::IsNullOrWhiteSpace($className)) {
            $status = 'FAIL'
            $details.Add('FullClassName missing')
        }

        if ($Installed) {
            if (-not (Test-Path -LiteralPath $assemblyPath)) {
                $status = 'FAIL'
                $details.Add("installed assembly missing: $assemblyPath")
            }
            $addinFolder = Split-Path -Parent $assemblyPath
        }
        else {
            $addinFolder = Join-Path $yearRoot 'KKY_FamilyBrowser'
            $stagedAssembly = Join-Path $addinFolder $assemblyFile
            if (-not (Test-Path -LiteralPath $stagedAssembly)) {
                $status = 'FAIL'
                $details.Add("staged assembly missing: $stagedAssembly")
            }
        }

        $revitApiCopies = Get-ChildItem -LiteralPath $addinFolder -Filter 'RevitAPI*.dll' -File -ErrorAction SilentlyContinue
        if ($revitApiCopies.Count -gt 0) {
            $status = 'FAIL'
            $details.Add('RevitAPI dlls should not be deployed in addin folder')
        }

        if ($details.Count -eq 0) {
            $details.Add("$($addinFile.Name) -> $assemblyFile / $className")
        }

        if ($Installed -and $status -eq 'OK') {
            $stageYearRoot = Join-Path $StageRoot "Rvt$year"
            if (-not (Test-Path -LiteralPath $stageYearRoot)) {
                $status = 'FAIL'
                $details.Add("staged year payload missing: $stageYearRoot")
            }
            else {
                $stageFiles = @(Get-ChildItem -LiteralPath $stageYearRoot -Recurse -File)
                $stageRelativePaths = @($stageFiles | ForEach-Object { Get-RelativePayloadPath -Root $stageYearRoot -Path $_.FullName })
                foreach ($stageFile in $stageFiles) {
                    $relativePath = Get-RelativePayloadPath -Root $stageYearRoot -Path $stageFile.FullName
                    $installedPath = Join-Path $yearRoot $relativePath.Replace('/', '\')
                    if (-not (Test-Path -LiteralPath $installedPath)) {
                        $status = 'FAIL'
                        $details.Add("installed payload missing: $relativePath")
                        continue
                    }
                    $installedFile = Get-Item -LiteralPath $installedPath
                    if ($stageFile.Length -ne $installedFile.Length) {
                        $status = 'FAIL'
                        $details.Add("installed payload length mismatch: $relativePath")
                        continue
                    }
                    $stageHash = (Get-FileHash -LiteralPath $stageFile.FullName -Algorithm SHA256).Hash
                    $installedHash = (Get-FileHash -LiteralPath $installedPath -Algorithm SHA256).Hash
                    if ($stageHash -ne $installedHash) {
                        $status = 'FAIL'
                        $details.Add("installed payload hash mismatch: $relativePath")
                    }
                }

                $installedPayloadFiles = @(Get-ChildItem -LiteralPath $addinFolder -Recurse -File -ErrorAction SilentlyContinue)
                $installedPayloadRelativePaths = @($installedPayloadFiles | ForEach-Object { 'KKY_FamilyBrowser/' + (Get-RelativePayloadPath -Root $addinFolder -Path $_.FullName) })
                $extraPayload = @($installedPayloadRelativePaths | Where-Object { $stageRelativePaths -notcontains $_ })
                if ($extraPayload.Count -gt 0) {
                    $status = 'FAIL'
                    $details.Add("obsolete installed payload remains: $($extraPayload -join ', ')")
                }

                $knownManifestNames = @(
                    'KKY_FamilyBrowser_RevitHost.addin',
                    'KKY_FamilyBrowser_RevitHost_2025.addin',
                    'KKY_FamilyBrowser_RevitHost_2027.addin'
                )
                $obsoleteManifests = @($knownManifestNames | Where-Object { $_ -ne $expectedAddinName -and (Test-Path -LiteralPath (Join-Path $yearRoot $_)) })
                if ($obsoleteManifests.Count -gt 0) {
                    $status = 'FAIL'
                    $details.Add("obsolete addin manifest remains: $($obsoleteManifests -join ', ')")
                }

                if ($status -eq 'OK') {
                    $details.Add("full payload hash match ($($stageFiles.Count) files)")
                }
            }
        }
    }
    catch {
        $status = 'FAIL'
        $details.Add($_.Exception.Message)
    }

    $results.Add([pscustomobject]@{
        Year = $year
        Scope = $(if ($Installed) { 'Installed' } else { 'Stage' })
        Status = $status
        Details = ($details -join '; ')
    })
}

$results | Format-Table -AutoSize

if (($results | Where-Object { $_.Status -ne 'OK' }).Count -gt 0) {
    throw 'Family Browser addin verification failed.'
}

Write-Host "Family Browser addin verification passed." -ForegroundColor Green
