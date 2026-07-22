param(
    [string]$BootstrapUrl = 'https://update.zerokky.com/family-browser/bootstrap.json',
    [string]$BootstrapJsonPath,
    [string[]]$Years = @('2019', '2021', '2023', '2025', '2027'),
    [string]$OutputDir,
    [switch]$SkipHashValidation,
    [switch]$TreatUnavailableAsFailure
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
if (-not $OutputDir) {
    $OutputDir = Join-Path $repoRoot ('artifacts\family-browser-managed-data-audit\' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
}
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$checks = New-Object System.Collections.Generic.List[object]
$candidateKeys = @{}
$candidates = New-Object System.Collections.Generic.List[object]
$bootstrap = $null
$selectedCandidate = $null
$dataRoot = ''
$policyPath = ''

function Add-Check {
    param(
        [string]$Component,
        [ValidateSet('OK', 'WARN', 'FAIL', 'INFO', 'SKIP')][string]$Status,
        [string]$Path,
        [string]$Message
    )
    $checks.Add([pscustomobject]@{
        Component = $Component
        Status = $Status
        Path = $Path
        Message = $Message
    }) | Out-Null
}

function Get-PropertyValue {
    param($Object, [string[]]$Names)
    if ($null -eq $Object) {
        return $null
    }
    foreach ($name in $Names) {
        $property = $Object.PSObject.Properties[$name]
        if ($null -ne $property) {
            return $property.Value
        }
    }
    return $null
}

function Expand-ManagedPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ''
    }
    return [Environment]::ExpandEnvironmentVariables($Path.Trim())
}

function Resolve-PolicyPathFromRoot([string]$RootPath) {
    $expanded = Expand-ManagedPath $RootPath
    if ([string]::IsNullOrWhiteSpace($expanded)) {
        return ''
    }
    if ([string]::Equals([IO.Path]::GetExtension($expanded), '.json', [StringComparison]::OrdinalIgnoreCase)) {
        return $expanded
    }
    return [IO.Path]::Combine($expanded, 'Config', 'standard-policy.json')
}

function Test-AllowedManagedRuntimePath([string]$Path) {
    $expanded = Expand-ManagedPath $Path
    if ([string]::IsNullOrWhiteSpace($expanded)) {
        return $false
    }
    if ($expanded.StartsWith('\\', [StringComparison]::Ordinal)) {
        return $true
    }
    try {
        $root = [IO.Path]::GetPathRoot($expanded)
        $windowsRoot = [IO.Path]::GetPathRoot([Environment]::GetFolderPath([Environment+SpecialFolder]::Windows))
        if ([string]::IsNullOrWhiteSpace($root) -or [string]::Equals($root, $windowsRoot, [StringComparison]::OrdinalIgnoreCase)) {
            return $false
        }
        return $true
    }
    catch {
        return $false
    }
}

function Add-ManagedCandidate {
    param([string]$DisplayPath, [string]$PolicyPath, [bool]$IsRoot)
    $display = Expand-ManagedPath $DisplayPath
    $policy = Expand-ManagedPath $PolicyPath
    if ([string]::IsNullOrWhiteSpace($policy)) {
        return
    }
    $key = $policy.ToLowerInvariant()
    if ($candidateKeys.ContainsKey($key)) {
        return
    }
    $candidateKeys[$key] = $true
    $candidates.Add([pscustomobject]@{
        DisplayPath = $display
        PolicyPath = $policy
        IsRoot = $IsRoot
    }) | Out-Null
}

function Read-JsonFile {
    param([string]$Path, [string]$Component)
    try {
        $value = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        Add-Check $Component 'OK' $Path 'JSON parsed successfully.'
        return $value
    }
    catch {
        Add-Check $Component 'FAIL' $Path ("JSON could not be read: " + $_.Exception.Message)
        return $null
    }
}

function Test-ReferencedJsonFile {
    param([string]$Path, [string]$Component, [bool]$Required = $true)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        Add-Check $Component $(if ($Required) { 'FAIL' } else { 'INFO' }) '' $(if ($Required) { 'Referenced path is empty.' } else { 'Optional path is empty.' })
        return $null
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        Add-Check $Component $(if ($Required) { 'FAIL' } else { 'WARN' }) $Path 'Referenced file does not exist.'
        return $null
    }
    return Read-JsonFile $Path $Component
}

function Test-ArtifactReference {
    param(
        [string]$ManifestFolder,
        $Reference,
        [string]$Component,
        [bool]$Required = $false
    )
    $relativePath = [string](Get-PropertyValue $Reference @('RelativePath', 'relativePath'))
    if ([string]::IsNullOrWhiteSpace($relativePath)) {
        Add-Check $Component $(if ($Required) { 'FAIL' } else { 'INFO' }) $ManifestFolder $(if ($Required) { 'Required artifact reference is empty.' } else { 'Optional artifact is not published.' })
        return $null
    }
    $manifestFull = [IO.Path]::GetFullPath($ManifestFolder).TrimEnd('\', '/')
    $artifactPath = [IO.Path]::GetFullPath((Join-Path $ManifestFolder $relativePath))
    if (-not ($artifactPath.Equals($manifestFull, [StringComparison]::OrdinalIgnoreCase) -or $artifactPath.StartsWith($manifestFull + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase))) {
        Add-Check $Component 'FAIL' $artifactPath 'Artifact path escapes its BrowserCacheV2 source folder.'
        return $null
    }
    if (-not (Test-Path -LiteralPath $artifactPath -PathType Leaf)) {
        Add-Check $Component 'FAIL' $artifactPath 'Referenced artifact is missing.'
        return $null
    }
    $file = Get-Item -LiteralPath $artifactPath
    $expectedLength = Get-PropertyValue $Reference @('Length', 'length')
    if ($null -ne $expectedLength -and [long]$expectedLength -gt 0 -and $file.Length -ne [long]$expectedLength) {
        Add-Check $Component 'FAIL' $artifactPath ("Length mismatch. expected=$expectedLength actual=$($file.Length)")
        return $null
    }
    $expectedHash = [string](Get-PropertyValue $Reference @('Sha256', 'sha256'))
    if (-not $SkipHashValidation -and -not [string]::IsNullOrWhiteSpace($expectedHash)) {
        $actualHash = (Get-FileHash -LiteralPath $artifactPath -Algorithm SHA256).Hash
        if (-not [string]::Equals($actualHash, $expectedHash, [StringComparison]::OrdinalIgnoreCase)) {
            Add-Check $Component 'FAIL' $artifactPath 'SHA256 mismatch.'
            return $null
        }
    }
    Add-Check $Component 'OK' $artifactPath 'Artifact exists and matches its manifest reference.'
    return $artifactPath
}

try {
    if ($BootstrapJsonPath) {
        $bootstrap = Read-JsonFile $BootstrapJsonPath 'Homepage bootstrap'
    }
    else {
        try {
            $response = Invoke-WebRequest -UseBasicParsing -Uri $BootstrapUrl -TimeoutSec 20
            $bootstrap = $response.Content | ConvertFrom-Json -ErrorAction Stop
            Add-Check 'Homepage bootstrap' 'OK' $BootstrapUrl ("HTTP " + [int]$response.StatusCode)
        }
        catch {
            Add-Check 'Homepage bootstrap' 'FAIL' $BootstrapUrl ("Bootstrap could not be downloaded: " + $_.Exception.Message)
        }
    }

    if ($null -ne $bootstrap) {
        $rootPath = [string](Get-PropertyValue $bootstrap @('ManagedRootPath', 'managedRootPath'))
        Add-ManagedCandidate $rootPath (Resolve-PolicyPathFromRoot $rootPath) $true
        $rootCandidates = @(Get-PropertyValue $bootstrap @('ManagedRootPathCandidates', 'managedRootPathCandidates'))
        foreach ($rootCandidate in $rootCandidates) {
            Add-ManagedCandidate ([string]$rootCandidate) (Resolve-PolicyPathFromRoot ([string]$rootCandidate)) $true
        }
        $directPolicyPath = [string](Get-PropertyValue $bootstrap @('ManagedPolicyPath', 'managedPolicyPath'))
        Add-ManagedCandidate $directPolicyPath $directPolicyPath $false
        $policyCandidates = @(Get-PropertyValue $bootstrap @('ManagedPolicyPathCandidates', 'managedPolicyPathCandidates'))
        foreach ($policyCandidate in $policyCandidates) {
            Add-ManagedCandidate ([string]$policyCandidate) ([string]$policyCandidate) $false
        }
    }

    foreach ($candidate in $candidates) {
		$allowed = Test-AllowedManagedRuntimePath $candidate.PolicyPath
        $rootReachable = $candidate.IsRoot -and (Test-Path -LiteralPath $candidate.DisplayPath -PathType Container)
        $policyExists = Test-Path -LiteralPath $candidate.PolicyPath -PathType Leaf
        $parent = Split-Path -Parent $candidate.PolicyPath
        $parentReachable = -not [string]::IsNullOrWhiteSpace($parent) -and (Test-Path -LiteralPath $parent -PathType Container)
		$reachable = $allowed -and ($rootReachable -or $policyExists -or $parentReachable)
		$message = if (-not $allowed) { 'Candidate is local to the Windows system drive and is rejected by the runtime managed-path policy.' } elseif ($reachable) { 'Candidate is reachable.' } else { 'Candidate root, policy file, and policy parent are unavailable.' }
		Add-Check 'Managed path candidate' $(if ($reachable) { 'OK' } else { 'WARN' }) $candidate.DisplayPath $message
        if ($null -eq $selectedCandidate -and $reachable) {
            $selectedCandidate = $candidate
        }
    }

    if ($null -ne $selectedCandidate) {
        $policyPath = $selectedCandidate.PolicyPath
        $policyFolder = Split-Path -Parent $policyPath
        if ([string]::Equals((Split-Path -Leaf $policyFolder), 'Config', [StringComparison]::OrdinalIgnoreCase)) {
            $dataRoot = Split-Path -Parent $policyFolder
        }
        else {
            $dataRoot = $policyFolder
        }
        Add-Check 'Managed data root' 'OK' $dataRoot 'First reachable homepage candidate selected using runtime candidate order.'

        $policy = $null
        if (Test-Path -LiteralPath $policyPath -PathType Leaf) {
            $policy = Read-JsonFile $policyPath 'Shared standard policy'
        }
        else {
            Add-Check 'Shared standard policy' 'WARN' $policyPath 'Managed folder is reachable but standard-policy.json has not been created yet.'
        }

        if ($null -ne $policy) {
            $slots = New-Object System.Collections.Generic.List[object]
            $integrated = Get-PropertyValue $policy @('IntegratedLibrary', 'integratedLibrary')
            if ($null -ne $integrated) {
                $slots.Add($integrated) | Out-Null
            }
            foreach ($slot in @(Get-PropertyValue $policy @('DisciplineLibraries', 'disciplineLibraries'))) {
                if ($null -ne $slot) {
                    $slots.Add($slot) | Out-Null
                }
            }
            foreach ($slot in $slots) {
                $enabled = Get-PropertyValue $slot @('Enabled', 'enabled')
                if ($null -ne $enabled -and -not [bool]$enabled) {
                    continue
                }
                $slotKey = [string](Get-PropertyValue $slot @('SlotKey', 'slotKey'))
                $standardListPath = Expand-ManagedPath ([string](Get-PropertyValue $slot @('StandardListPath', 'standardListPath')))
                if (-not [string]::IsNullOrWhiteSpace($standardListPath)) {
                    if (Test-Path -LiteralPath $standardListPath -PathType Leaf) {
                        Add-Check "Standard list [$slotKey]" 'OK' $standardListPath 'Configured approved-list source exists.'
                    }
                    else {
                        Add-Check "Standard list [$slotKey]" 'FAIL' $standardListPath 'Configured approved-list source is missing.'
                    }
                }
            }
        }

        foreach ($year in $Years) {
            $versionRoot = Join-Path $dataRoot ("RevitVersions\Revit$year")
            if (-not (Test-Path -LiteralPath $versionRoot -PathType Container)) {
                Add-Check "Revit $year data root" 'INFO' $versionRoot 'No managed data has been created for this Revit version.'
                continue
            }
            Add-Check "Revit $year data root" 'OK' $versionRoot 'Versioned managed data folder exists.'
            $registryFolder = Join-Path $versionRoot 'Registry'
            $registrations = @()
            if (Test-Path -LiteralPath $registryFolder -PathType Container) {
                $registrations = @(Get-ChildItem -LiteralPath $registryFolder -Filter '*.json' -File -ErrorAction SilentlyContinue)
            }
            Add-Check "Revit $year registry" 'INFO' $registryFolder ("Registration files: " + $registrations.Count)
            foreach ($registrationFile in $registrations) {
                $registration = Read-JsonFile $registrationFile.FullName "Revit $year registration"
                if ($null -eq $registration) {
                    continue
                }
                $snapshotPath = Expand-ManagedPath ([string](Get-PropertyValue $registration @('LastSnapshotPath', 'lastSnapshotPath')))
                $null = Test-ReferencedJsonFile $snapshotPath "Revit $year standard snapshot" $true
				$sourcePath = Expand-ManagedPath ([string](Get-PropertyValue $registration @('ResolvedPath', 'resolvedPath')))
				if ([string]::IsNullOrWhiteSpace($sourcePath)) {
					$sourcePath = Expand-ManagedPath ([string](Get-PropertyValue $registration @('Locator', 'locator')))
				}
                if (-not [string]::IsNullOrWhiteSpace($sourcePath)) {
                    Add-Check "Revit $year standard RVT" $(if (Test-Path -LiteralPath $sourcePath -PathType Leaf) { 'OK' } else { 'WARN' }) $sourcePath $(if (Test-Path -LiteralPath $sourcePath -PathType Leaf) { 'Registered RVT source exists.' } else { 'Registered RVT source is unavailable; browsing can use a valid snapshot, but load/apply will be blocked.' })
                }
            }

            $thumbnailFolder = Join-Path $versionRoot 'Thumbnails'
            $thumbnailCount = if (Test-Path -LiteralPath $thumbnailFolder -PathType Container) { @(Get-ChildItem -LiteralPath $thumbnailFolder -Filter '*.png' -File -Recurse -ErrorAction SilentlyContinue).Count } else { 0 }
            Add-Check "Revit $year thumbnails" 'INFO' $thumbnailFolder ("PNG files: $thumbnailCount")

            $projectsFolder = Join-Path $versionRoot 'Projects'
            $projectRecords = @()
            if (Test-Path -LiteralPath $projectsFolder -PathType Container) {
                $projectRecords = @(Get-ChildItem -LiteralPath $projectsFolder -Filter 'project-scan-latest*.json' -File -Recurse -ErrorAction SilentlyContinue)
            }
            Add-Check "Revit $year project cache" 'INFO' $projectsFolder ("Latest/alias records: " + $projectRecords.Count)
            foreach ($projectRecordFile in $projectRecords) {
                $record = Read-JsonFile $projectRecordFile.FullName "Revit $year project scan record"
                if ($null -eq $record) {
                    continue
                }
                $projectSnapshotPath = Expand-ManagedPath ([string](Get-PropertyValue $record @('ProjectSnapshotPath', 'projectSnapshotPath')))
                $comparisonPath = Expand-ManagedPath ([string](Get-PropertyValue $record @('ComparisonReportPath', 'comparisonReportPath')))
                $standardSnapshotPath = Expand-ManagedPath ([string](Get-PropertyValue $record @('StandardSnapshotPath', 'standardSnapshotPath')))
                $null = Test-ReferencedJsonFile $projectSnapshotPath "Revit $year cached project snapshot" $true
                $null = Test-ReferencedJsonFile $comparisonPath "Revit $year cached comparison report" $true
                $null = Test-ReferencedJsonFile $standardSnapshotPath "Revit $year cached standard snapshot" $true
            }
        }

        $browserCacheRoot = Join-Path $dataRoot 'BrowserCacheV2'
        $manifestFiles = @()
        if (Test-Path -LiteralPath $browserCacheRoot -PathType Container) {
            $manifestFiles = @(Get-ChildItem -LiteralPath $browserCacheRoot -Filter 'family-browser-manifest-v2.json' -File -Recurse -ErrorAction SilentlyContinue)
        }
        Add-Check 'BrowserCacheV2' 'INFO' $browserCacheRoot ("Manifest files: " + $manifestFiles.Count)
        foreach ($manifestFile in $manifestFiles) {
            $manifest = Read-JsonFile $manifestFile.FullName 'Browser V2 manifest'
            if ($null -eq $manifest) {
                continue
            }
            $schemaVersion = Get-PropertyValue $manifest @('SchemaVersion', 'schemaVersion')
            if ([int]$schemaVersion -ne 2) {
                Add-Check 'Browser V2 manifest schema' 'FAIL' $manifestFile.FullName ("Expected schema 2, found $schemaVersion.")
            }
            $manifestFolder = Split-Path -Parent $manifestFile.FullName
            $indexPath = Test-ArtifactReference $manifestFolder (Get-PropertyValue $manifest @('StandardIndex', 'standardIndex')) 'Browser V2 standard index' $true
            $detailsCatalogPath = Test-ArtifactReference $manifestFolder (Get-PropertyValue $manifest @('StandardDetails', 'standardDetails')) 'Browser V2 details catalog' $true
            $null = Test-ArtifactReference $manifestFolder (Get-PropertyValue $manifest @('ThumbnailIndex', 'thumbnailIndex')) 'Browser V2 thumbnail index' $true
            $null = Test-ArtifactReference $manifestFolder (Get-PropertyValue $manifest @('StandardList', 'standardList')) 'Browser V2 standard list' $false
            $projectStatePath = Test-ArtifactReference $manifestFolder (Get-PropertyValue $manifest @('ProjectState', 'projectState')) 'Browser V2 project state' $false
            if ($indexPath) {
                $null = Read-JsonFile $indexPath 'Browser V2 standard index JSON'
            }
            if ($detailsCatalogPath) {
                $detailsCatalog = Read-JsonFile $detailsCatalogPath 'Browser V2 details catalog JSON'
                foreach ($detail in @(Get-PropertyValue $detailsCatalog @('Items', 'items'))) {
                    if ($null -eq $detail) {
                        continue
                    }
                    $detailReference = [pscustomobject]@{
                        RelativePath = [string](Get-PropertyValue $detail @('DetailKey', 'detailKey'))
                        Sha256 = [string](Get-PropertyValue $detail @('Sha256', 'sha256'))
                        Length = Get-PropertyValue $detail @('Length', 'length')
                    }
                    $detailPath = Test-ArtifactReference $manifestFolder $detailReference 'Browser V2 detail record' $true
                    if ($detailPath) {
                        $null = Read-JsonFile $detailPath 'Browser V2 detail JSON'
                    }
                }
            }
            if ($projectStatePath) {
                $projectState = Read-JsonFile $projectStatePath 'Browser V2 project state JSON'
                if ($null -ne $projectState) {
                    $stateProjectSnapshot = Expand-ManagedPath ([string](Get-PropertyValue $projectState @('ProjectSnapshotPath', 'projectSnapshotPath')))
                    $stateComparison = Expand-ManagedPath ([string](Get-PropertyValue $projectState @('ComparisonReportPath', 'comparisonReportPath')))
                    $null = Test-ReferencedJsonFile $stateProjectSnapshot 'Browser V2 project-state snapshot' $true
                    $null = Test-ReferencedJsonFile $stateComparison 'Browser V2 project-state comparison' $true
                }
            }
        }
    }
}
finally {
    $checkArray = @($checks.ToArray())
    $failCount = @($checkArray | Where-Object { $_.Status -eq 'FAIL' }).Count
    $warningCount = @($checkArray | Where-Object { $_.Status -eq 'WARN' }).Count
    $managedAvailable = $null -ne $selectedCandidate
    $overallStatus = if (-not $managedAvailable) { 'UNAVAILABLE' } elseif ($failCount -gt 0) { 'FAIL' } elseif ($warningCount -gt 0) { 'WARN' } else { 'OK' }
    $localCacheRoot = Join-Path $env:LOCALAPPDATA 'KKY\FamilyBrowser\Cache\v2'
    $localSourceManifestCount = if (Test-Path -LiteralPath $localCacheRoot -PathType Container) { @(Get-ChildItem -LiteralPath $localCacheRoot -Filter 'family-browser-manifest-v2.json' -File -Recurse -ErrorAction SilentlyContinue).Count } else { 0 }
    $localRowCacheCount = if (Test-Path -LiteralPath $localCacheRoot -PathType Container) { @(Get-ChildItem -LiteralPath $localCacheRoot -Filter 'browser-row-cache-v2-*.json' -File -Recurse -ErrorAction SilentlyContinue).Count } else { 0 }
    $summary = [ordered]@{
        GeneratedAt = (Get-Date).ToString('o')
        BootstrapUrl = $BootstrapUrl
        BootstrapJsonPath = $BootstrapJsonPath
        BootstrapVersion = if ($null -ne $bootstrap) { [string](Get-PropertyValue $bootstrap @('Version', 'version')) } else { '' }
        Status = $overallStatus
        ManagedPathAvailable = $managedAvailable
        SelectedPolicyPath = $policyPath
        DataRoot = $dataRoot
        CandidateCount = $candidates.Count
        FailureCount = $failCount
        WarningCount = $warningCount
        LocalCacheRoot = $localCacheRoot
        LocalSourceManifestCount = $localSourceManifestCount
        LocalRowCacheCount = $localRowCacheCount
        Checks = $checkArray
    }
    $summary | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $OutputDir 'managed-data-audit.json') -Encoding UTF8

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('# Family Browser Managed Data Audit') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add("- Generated: $($summary.GeneratedAt)") | Out-Null
    $lines.Add("- Status: $overallStatus") | Out-Null
    $lines.Add("- Bootstrap version: $($summary.BootstrapVersion)") | Out-Null
    $lines.Add("- Reachable managed path: $managedAvailable") | Out-Null
    $lines.Add("- Selected policy: $policyPath") | Out-Null
    $lines.Add("- Data root: $dataRoot") | Out-Null
    $lines.Add("- Failures / warnings: $failCount / $warningCount") | Out-Null
    $lines.Add("- Local V2 source manifests / row caches: $localSourceManifestCount / $localRowCacheCount") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('| Component | Status | Path | Message |') | Out-Null
    $lines.Add('|---|---:|---|---|') | Out-Null
    foreach ($check in $checkArray) {
        $safePath = ([string]$check.Path).Replace('|', '/')
        $safeMessage = ([string]$check.Message).Replace('|', '/').Replace("`r", ' ').Replace("`n", ' ')
        $lines.Add("| $($check.Component) | $($check.Status) | $safePath | $safeMessage |") | Out-Null
    }
    $lines | Set-Content -LiteralPath (Join-Path $OutputDir 'managed-data-audit.md') -Encoding UTF8
}

if ($summary.Status -eq 'FAIL') {
    throw "Family Browser managed-data audit found broken references. See $OutputDir"
}
if ($summary.Status -eq 'UNAVAILABLE' -and $TreatUnavailableAsFailure) {
    throw "No homepage managed-folder candidate is reachable. See $OutputDir"
}

Write-Host "Family Browser managed-data audit: $($summary.Status) - $OutputDir" -ForegroundColor $(if ($summary.Status -eq 'OK') { 'Green' } elseif ($summary.Status -eq 'WARN') { 'Yellow' } else { 'DarkYellow' })
