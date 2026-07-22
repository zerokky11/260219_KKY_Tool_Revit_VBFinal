param(
    [string[]]$HostFolders = @(
        'KKY_FamilyBrowser_RevitHost_2019-2023',
        'KKY_FamilyBrowser_RevitHost_2025',
        'KKY_FamilyBrowser_RevitHost_2027'
    )
)

$ErrorActionPreference = 'Stop'
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$failures = New-Object System.Collections.Generic.List[string]

function Add-Failure([string]$message) {
    $script:failures.Add($message) | Out-Null
}

function Assert-Contains([string]$text, [string]$needle, [string]$label) {
    if ($text.IndexOf($needle, [StringComparison]::Ordinal) -lt 0) {
        Add-Failure "$label is missing"
    }
}

function Assert-NotContains([string]$text, [string]$needle, [string]$label) {
    if ($text.IndexOf($needle, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
        Add-Failure "$label is unexpectedly present"
    }
}

function Assert-Regex([string]$text, [string]$pattern, [string]$label) {
    if (-not [regex]::IsMatch($text, $pattern, [Text.RegularExpressions.RegexOptions]::Singleline)) {
        Add-Failure "$label is missing"
    }
}

foreach ($hostFolder in $HostFolders) {
    $applyPath = Join-Path $repoRoot "$hostFolder\SystemTypeApplyExecutionService.cs"
    $policyPath = Join-Path $repoRoot "$hostFolder\SystemTypeSupportPolicyService.cs"
    if (-not (Test-Path -LiteralPath $applyPath)) {
        Add-Failure "$hostFolder apply source is missing: $applyPath"
        continue
    }
    if (-not (Test-Path -LiteralPath $policyPath)) {
        Add-Failure "$hostFolder policy source is missing: $policyPath"
        continue
    }

    $apply = Get-Content -LiteralPath $applyPath -Raw
    $policy = Get-Content -LiteralPath $policyPath -Raw

    Assert-Regex $policy '\{\s*"stairstype"\s*,\s*"PreflightThenConfirm"\s*\}' "$hostFolder Stair apply policy"
    Assert-Regex $policy '\{\s*"railingtype"\s*,\s*"PreflightThenConfirm"\s*\}' "$hostFolder Railing apply policy"
    Assert-Regex $policy '\{\s*"paneltype"\s*,\s*"ReviewOnly"\s*\}' "$hostFolder Curtain Panel review-only policy"

    Assert-Contains $apply 'new string[2] { "railingtype", "stairstype" }' "$hostFolder authoritative Railing/Stair root set"
    Assert-Contains $apply 'PrepareAuthoritativeTypeLoadableDependencies(targetDocument, standardDocument, sourceType, dependencyItems, resultItem);' "$hostFolder dependency preparation"
    Assert-Contains $apply 'return IsSystemTypeRoutingRebuildAction(syncItem.Action);' "$hostFolder atomic create/overwrite group"
    Assert-Regex $apply 'CopyCanonicalType\([^;]+;\s*if\s*\(copiedType\s*!=\s*null\)\s*\{\s*ApplyStandardSystemTypeDefinition' "$hostFolder canonical-copy definition apply"

    Assert-Contains $apply 'ApplyAuthoritativeCompoundStructure(targetDocument, standardDocument, sourceType, targetType, resultItem);' "$hostFolder compound structure apply call"
    Assert-Contains $apply 'targetHost.SetCompoundStructure(mappedSource);' "$hostFolder compound structure mutation"
    Assert-Contains $apply '!mappedSource.IsEqual(applied)' "$hostFolder compound structure post-check"
    Assert-Contains $apply 'RevitElementIdCompat.CompatIntegerValue(sourceMaterialId) > 0' "$hostFolder By Category material guard"
    Assert-Contains $apply 'RevitElementIdCompat.CompatIntegerValue(sourceDeckProfileId) > 0' "$hostFolder empty deck-profile guard"

    Assert-Contains $apply 'FindMaterialByName(targetDocument, materialName)' "$hostFolder same-name material reuse"
    Assert-Contains $apply 'Material.Create(targetDocument, materialName)' "$hostFolder missing material creation"
    Assert-Contains $apply 'if (!material.ExactSignatureMatch)' "$hostFolder strict layer material post-check"
    Assert-Contains $apply 'system type apply was rolled back' "$hostFolder material mismatch rollback"
    Assert-Contains $apply '"AppearanceAssetId"' "$hostFolder appearance asset synchronization"
    Assert-Contains $apply '"StructuralAssetId"' "$hostFolder structural asset synchronization"
    Assert-Contains $apply '"ThermalAssetId"' "$hostFolder thermal asset synchronization"

    Assert-Contains $apply 'ReloadAuthoritativeComponentFamily' "$hostFolder referenced family authoritative reload"
    Assert-Contains $apply 'GuardDependencyLoadDidNotCreateDuplicateFamilies' "$hostFolder duplicate family guard"
    Assert-Contains $apply 'ConsolidateObsoleteTypes' "$hostFolder duplicate system type consolidation"

    Assert-Contains $apply 'ApplyAuthoritativeDetailedSystemTypeDefinition(targetDocument, standardDocument, sourceType, targetType, resultItem);' "$hostFolder detailed Railing/Stair apply"
    Assert-Contains $apply 'BuildOptionalDetailedComponentSignature(sourceRows)' "$hostFolder detailed source post-check"
    Assert-Contains $apply 'BuildOptionalDetailedComponentSignature(targetRows)' "$hostFolder detailed target post-check"
    Assert-Contains $apply 'independently of the comparison option' "$hostFolder comparison-option-independent apply diagnostic"
    Assert-NotContains $apply 'compareDetailedSystemTypeComponents' "$hostFolder comparison option in mutation path"
}

if ($failures.Count -gt 0) {
    Write-Host 'System type authoritative apply checks failed:' -ForegroundColor Red
    foreach ($failure in $failures) {
        Write-Host " - $failure" -ForegroundColor Red
    }
    exit 1
}

Write-Host "System type authoritative apply checks passed for $($HostFolders.Count) host source sets." -ForegroundColor Green
exit 0
