param()

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
$sourcePaths = @(
    (Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2019-2023\StandardNestedLoadableFamilySnapshotItem.cs'),
    (Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2019-2023\StandardFamilyParameterSnapshotItem.cs'),
    (Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2019-2023\LoadableFingerprintDifferenceDetailItem.cs'),
    (Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2019-2023\LoadableFamilyComparisonItem.cs'),
    (Join-Path $repoRoot 'KKY_FamilyBrowser_SharedUi\NestedLoadableFamilyDifferencePropagationService.cs')
)

foreach ($path in $sourcePaths) {
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Nested-family propagation source is missing: $path"
    }
}

if (-not ('NestedLoadableFamilyDifferencePropagationService' -as [type])) {
    $body = ($sourcePaths | ForEach-Object { (Get-Content -LiteralPath $_ -Raw) -replace '(?m)^using\s+[^;]+;\s*\r?\n', '' }) -join [Environment]::NewLine
    $source = "using System;`r`nusing System.Collections.Generic;`r`nusing System.Linq;`r`n" + $body
    Add-Type -TypeDefinition $source -Language CSharp
}

function New-NestedReference {
    param([string]$Name, [string]$Category = 'Mechanical Equipment')
    $item = [StandardNestedLoadableFamilySnapshotItem]::new()
    $item.FamilyName = $Name
    $item.CategoryName = $Category
    return $item
}

function New-ComparisonItem {
    param(
        [string]$Name,
        [string]$Status,
        [string]$Category = 'Mechanical Equipment'
    )
    $item = [LoadableFamilyComparisonItem]::new()
    $item.FamilyName = $Name
    $item.CategoryName = $Category
    $item.Status = $Status
    $item.StandardFingerprint = 'standard-' + $Name
    $item.ProjectFingerprint = 'project-' + $Name
    $item.StandardContentFingerprint = 'standard-content-' + $Name
    $item.ProjectContentFingerprint = 'project-content-' + $Name
    return $item
}

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

$parent = New-ComparisonItem -Name 'AUDIT_PARENT' -Status 'LoadedLatest'
$parent.Notes = 'Project content matches the current standard snapshot and tracked approval stamp.'
$parent.NestedLoadableFamilies.Add((New-NestedReference -Name 'AUDIT_CHILD'))
$child = New-ComparisonItem -Name 'AUDIT_CHILD' -Status 'DifferentFromStandard'
$child.FingerprintDifferenceSummary.Add('Parameter differs: Width.')
$childDetail = [LoadableFingerprintDifferenceDetailItem]::new()
$childDetail.Area = 'parameters/formulas'
$childDetail.DifferenceKind = 'modified'
$childDetail.StandardValue = 'Width=600'
$childDetail.ProjectValue = 'Width=500'
$childDetail.Details = 'Parameter value differs.'
$child.FingerprintDifferenceDetails.Add($childDetail)
$directItems = [System.Collections.Generic.List[LoadableFamilyComparisonItem]]::new()
$directItems.Add($parent)
$directItems.Add($child)
[NestedLoadableFamilyDifferencePropagationService]::Apply($directItems)

Assert-True $child.IsNestedLoadableChild 'Differing child was not marked as a nested family.'
Assert-True $child.IsNestedLoadableDifference 'Differing child was not marked for exception display.'
Assert-True ($child.NestedParentFamilyNames -contains 'AUDIT_PARENT') 'Differing child did not retain its parent family name.'
Assert-True ($parent.Status -eq 'DifferentFromStandard') 'Matching parent was not changed to DifferentFromStandard.'
Assert-True $parent.IsNestedLoadableDifference 'Parent did not retain its nested-difference display flag.'
Assert-True ($parent.NestedDifferenceFamilyNames.Count -eq 1) 'Parent did not record the differing nested family.'
Assert-True (($parent.FingerprintDifferenceSummary -join '|') -match 'AUDIT_CHILD') 'Parent summary does not identify the differing nested family.'
Assert-True (($parent.FingerprintDifferenceDetails | Where-Object { $_.Area -eq 'nested family' }).Count -ge 1) 'Parent detail table does not contain nested-family difference rows.'
Assert-True ($parent.Notes -notmatch 'matches the current standard snapshot') 'Parent retained a false basis-match note after nested propagation.'

$matchingParent = New-ComparisonItem -Name 'AUDIT_MATCHING_PARENT' -Status 'LoadedLatest'
$matchingParent.NestedLoadableFamilies.Add((New-NestedReference -Name 'AUDIT_MATCHING_CHILD'))
$matchingChild = New-ComparisonItem -Name 'AUDIT_MATCHING_CHILD' -Status 'LoadedLatest'
$matchingItems = [System.Collections.Generic.List[LoadableFamilyComparisonItem]]::new()
$matchingItems.Add($matchingParent)
$matchingItems.Add($matchingChild)
[NestedLoadableFamilyDifferencePropagationService]::Apply($matchingItems)

Assert-True $matchingChild.IsNestedLoadableChild 'Matching child was not recognized as nested.'
Assert-True (-not $matchingChild.IsNestedLoadableDifference) 'Matching nested child was incorrectly exposed as a difference.'
Assert-True ($matchingParent.Status -eq 'LoadedLatest') 'Matching parent was incorrectly changed.'

$grandParent = New-ComparisonItem -Name 'AUDIT_GRAND_PARENT' -Status 'LoadedLatest'
$nestedParent = New-ComparisonItem -Name 'AUDIT_NESTED_PARENT' -Status 'LoadedLatest'
$leaf = New-ComparisonItem -Name 'AUDIT_NESTED_LEAF' -Status 'LocallyModified'
$grandParent.NestedLoadableFamilies.Add((New-NestedReference -Name 'AUDIT_NESTED_PARENT'))
$nestedParent.NestedLoadableFamilies.Add((New-NestedReference -Name 'AUDIT_NESTED_LEAF'))
$transitiveItems = [System.Collections.Generic.List[LoadableFamilyComparisonItem]]::new()
$transitiveItems.Add($grandParent)
$transitiveItems.Add($nestedParent)
$transitiveItems.Add($leaf)
[NestedLoadableFamilyDifferencePropagationService]::Apply($transitiveItems)

Assert-True $nestedParent.IsNestedLoadableDifference 'Nested parent did not inherit its child difference.'
Assert-True ($grandParent.Status -eq 'DifferentFromStandard') 'Nested difference did not propagate to the top-level parent.'
Assert-True (($grandParent.FingerprintDifferenceSummary -join '|') -match 'AUDIT_NESTED_PARENT') 'Top-level parent does not identify the nested parent that differs.'

$projectParent = New-ComparisonItem -Name 'AUDIT_PROJECT_PARENT' -Status 'LoadedLatest'
$projectOnlyChild = New-ComparisonItem -Name 'AUDIT_PROJECT_ONLY_CHILD' -Status 'ProjectOnly'
$projectParent.ProjectNestedLoadableFamilies.Add((New-NestedReference -Name 'AUDIT_PROJECT_ONLY_CHILD'))
$projectOnlyItems = [System.Collections.Generic.List[LoadableFamilyComparisonItem]]::new()
$projectOnlyItems.Add($projectParent)
$projectOnlyItems.Add($projectOnlyChild)
[NestedLoadableFamilyDifferencePropagationService]::Apply($projectOnlyItems)

Assert-True $projectOnlyChild.IsNestedLoadableDifference 'Project-only nested family was not exposed for review.'
Assert-True ($projectOnlyChild.Status -eq 'NestedExtraInParent') 'Project-only nested family did not receive the explicit parent-composition status.'
Assert-True ($projectParent.Status -eq 'DifferentFromStandard') 'Project-only nested family did not mark its parent as different.'

$missingParent = New-ComparisonItem -Name 'AUDIT_MISSING_PARENT' -Status 'LoadedLatest'
$missingChild = New-ComparisonItem -Name 'AUDIT_MISSING_CHILD' -Status 'LoadAvailable'
$missingChild.ProjectFingerprint = ''
$missingChild.ProjectContentFingerprint = ''
$missingChild.ProjectContentFingerprintFailureReason = ''
$missingParent.NestedLoadableFamilies.Add((New-NestedReference -Name 'AUDIT_MISSING_CHILD'))
$missingItems = [System.Collections.Generic.List[LoadableFamilyComparisonItem]]::new()
$missingItems.Add($missingParent)
$missingItems.Add($missingChild)
[NestedLoadableFamilyDifferencePropagationService]::Apply($missingItems)

Assert-True ($missingChild.Status -eq 'NestedMissingFromParent') 'Missing nested child did not receive the explicit nested-missing status.'
Assert-True (($missingChild.FingerprintDifferenceSummary -join '|') -match 'missing from parent family') 'Missing child summary does not describe the parent-family relationship.'
Assert-True (($missingChild.Notes -match 'AUDIT_MISSING_PARENT') -and ($missingChild.Notes -match 'Update the parent family')) 'Missing child memo does not identify the parent update action.'
Assert-True (($missingChild.FingerprintDifferenceDetails | Where-Object { $_.Area -eq 'nested family' -and $_.DifferenceKind -eq 'missing' }).Count -eq 1) 'Missing child detail was not emitted as a nested-family table row.'
Assert-True ($missingParent.Status -eq 'DifferentFromStandard') 'Missing nested child did not mark its parent as different.'
Assert-True (($missingParent.FingerprintDifferenceSummary -join '|') -match 'AUDIT_MISSING_CHILD') 'Parent summary does not identify the missing nested child.'

$absentParent = New-ComparisonItem -Name 'AUDIT_ABSENT_PARENT' -Status 'LoadAvailable'
$absentChild = New-ComparisonItem -Name 'AUDIT_ABSENT_CHILD' -Status 'LoadAvailable'
$absentParent.NestedLoadableFamilies.Add((New-NestedReference -Name 'AUDIT_ABSENT_CHILD'))
$absentItems = [System.Collections.Generic.List[LoadableFamilyComparisonItem]]::new()
$absentItems.Add($absentParent)
$absentItems.Add($absentChild)
[NestedLoadableFamilyDifferencePropagationService]::Apply($absentItems)

Assert-True (-not $absentChild.IsNestedLoadableDifference) 'A child of a wholly absent parent was incorrectly exposed as a separate difference.'
Assert-True ($absentChild.Status -eq 'LoadAvailable') 'A child of a wholly absent parent lost its LoadAvailable state.'
Assert-True ($absentParent.Status -eq 'LoadAvailable') 'A wholly absent parent was incorrectly changed to DifferentFromStandard.'

$preDifferentGrandParent = New-ComparisonItem -Name 'AUDIT_PRE_DIFFERENT_GRAND_PARENT' -Status 'LoadedLatest'
$preDifferentParent = New-ComparisonItem -Name 'AUDIT_PRE_DIFFERENT_PARENT' -Status 'DifferentFromStandard'
$preDifferentLeaf = New-ComparisonItem -Name 'AUDIT_PRE_DIFFERENT_LEAF' -Status 'LocallyModified'
$preDifferentParent.FingerprintDifferenceSummary.Add('Own parameter differs: Height.')
$preDifferentGrandParent.NestedLoadableFamilies.Add((New-NestedReference -Name 'AUDIT_PRE_DIFFERENT_PARENT'))
$preDifferentParent.NestedLoadableFamilies.Add((New-NestedReference -Name 'AUDIT_PRE_DIFFERENT_LEAF'))
$preDifferentItems = [System.Collections.Generic.List[LoadableFamilyComparisonItem]]::new()
$preDifferentItems.Add($preDifferentGrandParent)
$preDifferentItems.Add($preDifferentParent)
$preDifferentItems.Add($preDifferentLeaf)
[NestedLoadableFamilyDifferencePropagationService]::Apply($preDifferentItems)

Assert-True ($preDifferentGrandParent.Status -eq 'DifferentFromStandard') 'A pre-different nested parent did not mark the top-level parent as different.'
Assert-True (($preDifferentGrandParent.FingerprintDifferenceDetails | Where-Object { $_.Details -match 'AUDIT_PRE_DIFFERENT_LEAF' }).Count -ge 1) 'The deepest nested-family reason did not reach the top-level parent detail rows.'

Write-Host 'Nested family difference propagation tests passed.' -ForegroundColor Green
