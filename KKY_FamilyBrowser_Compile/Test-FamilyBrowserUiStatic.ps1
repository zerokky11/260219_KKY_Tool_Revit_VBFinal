param(
    [string[]]$HostFolders = @(
        'KKY_FamilyBrowser_RevitHost_2019-2023',
        'KKY_FamilyBrowser_RevitHost_2025',
        'KKY_FamilyBrowser_RevitHost_2027'
    )
)

$ErrorActionPreference = 'Stop'
$root = Resolve-Path (Join-Path $PSScriptRoot '..')
$failures = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[string]

function Add-Failure([string]$message) {
    $script:failures.Add($message) | Out-Null
}

function Add-Warning([string]$message) {
    $script:warnings.Add($message) | Out-Null
}

function Assert-Contains([string]$path, [string]$needle, [string]$label) {
    $text = Get-Content -LiteralPath $path -Raw -Encoding UTF8
    if ($text.IndexOf($needle, [StringComparison]::Ordinal) -lt 0) {
        Add-Failure "$label missing in $path"
    }
}

function Assert-NotContains([string]$path, [string]$needle, [string]$label) {
    $text = Get-Content -LiteralPath $path -Raw -Encoding UTF8
    if ($text.IndexOf($needle, [StringComparison]::Ordinal) -ge 0) {
        Add-Failure "$label unexpectedly present in $path"
    }
}

function Assert-Before([string]$path, [string]$firstNeedle, [string]$secondNeedle, [string]$label) {
    $text = Get-Content -LiteralPath $path -Raw -Encoding UTF8
    $firstIndex = $text.IndexOf($firstNeedle, [StringComparison]::Ordinal)
    $secondIndex = $text.IndexOf($secondNeedle, [StringComparison]::Ordinal)
    if ($firstIndex -lt 0 -or $secondIndex -lt 0 -or $firstIndex -ge $secondIndex) {
        Add-Failure "$label order is invalid in $path"
    }
}

function Assert-MinimumOccurrences([string]$path, [string]$needle, [int]$minimum, [string]$label) {
    $text = Get-Content -LiteralPath $path -Raw -Encoding UTF8
    $count = ([regex]::Matches($text, [regex]::Escape($needle))).Count
    if ($count -lt $minimum) {
        Add-Failure "$label expected at least $minimum occurrence(s), found $count in $path"
    }
}

function Test-DocumentTextHasMeta([string]$path) {
    $lines = Get-Content -LiteralPath $path -Encoding UTF8
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -notmatch 'DocumentText\s*=') {
            continue
        }
        if ($lines[$i] -match 'DocumentText\s*=\s*(BuildHtml\(\)|html\b|startupHtml\b|FamilyBrowserMessageHtmlRenderer\.Build)') {
            continue
        }

        $start = [Math]::Max(0, $i - 90)
        $end = [Math]::Min($lines.Count - 1, $i + 20)
        $window = ($lines[$start..$end] -join "`n")
        if ($window -notmatch '<meta\s+charset|charset\s*=') {
            Add-Warning "DocumentText assignment may not be backed by UTF-8 meta near line $($i + 1) in $path"
        }
    }
}

function Assert-WindowContains([string]$path, [string]$anchor, [string]$needle, [string]$label, [int]$forwardLines = 24) {
    $lines = Get-Content -LiteralPath $path -Encoding UTF8
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i].IndexOf($anchor, [StringComparison]::Ordinal) -lt 0) {
            continue
        }
        $start = [Math]::Max(0, $i - 4)
        $end = [Math]::Min($lines.Count - 1, $i + $forwardLines)
        $window = ($lines[$start..$end] -join "`n")
        if ($window.IndexOf($needle, [StringComparison]::Ordinal) -lt 0) {
            Add-Failure "$label missing near '$anchor' in $path"
        }
        return
    }
    Add-Failure "$label anchor '$anchor' missing in $path"
}

function Assert-WindowNotContains([string]$path, [string]$anchor, [string]$needle, [string]$label) {
    $lines = Get-Content -LiteralPath $path -Encoding UTF8
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i].IndexOf($anchor, [StringComparison]::Ordinal) -lt 0) {
            continue
        }
        $start = [Math]::Max(0, $i - 4)
        $end = [Math]::Min($lines.Count - 1, $i + 24)
        $window = ($lines[$start..$end] -join "`n")
        if ($window.IndexOf($needle, [StringComparison]::Ordinal) -ge 0) {
            Add-Failure "$label unexpectedly present near '$anchor' in $path"
        }
        return
    }
}

function Assert-MethodContainsAny([string]$path, [string]$methodName, [string[]]$needles, [string]$label) {
    $lines = Get-Content -LiteralPath $path -Encoding UTF8
    $escapedMethod = [regex]::Escape($methodName)
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match ("^\s*(private|internal|public)\s+.*\b" + $escapedMethod + "\s*\(")) {
            $start = $i
            break
        }
    }
    if ($start -lt 0) {
        Add-Failure "$label method '$methodName' missing in $path"
        return
    }

    $end = $lines.Count - 1
    for ($i = $start + 1; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*(private|internal|public)\s+.*\w+\s*\(') {
            $end = $i - 1
            break
        }
    }
    $methodBody = ($lines[$start..$end] -join "`n")
    foreach ($needle in $needles) {
        if ($methodBody.IndexOf($needle, [StringComparison]::Ordinal) -ge 0) {
            return
        }
    }
    Add-Failure "$label method '$methodName' does not trigger an immediate dashboard refresh in $path"
}

function Assert-MethodNotContains([string]$path, [string]$methodName, [string]$needle, [string]$label) {
    $lines = Get-Content -LiteralPath $path -Encoding UTF8
    $escapedMethod = [regex]::Escape($methodName)
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match ("^\s*(private|internal|public)\s+.*\b" + $escapedMethod + "\s*\(")) {
            $start = $i
            break
        }
    }
    if ($start -lt 0) {
        Add-Failure "$label method '$methodName' missing in $path"
        return
    }
    $end = $lines.Count - 1
    for ($i = $start + 1; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*(private|internal|public)\s+.*\w+\s*\(') {
            $end = $i - 1
            break
        }
    }
    $methodBody = ($lines[$start..$end] -join "`n")
    if ($methodBody.IndexOf($needle, [StringComparison]::Ordinal) -ge 0) {
        Add-Failure "$label method '$methodName' unexpectedly contains '$needle' in $path"
    }
}

$standardRevisionServicePath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserStandardRevisionService.cs'
$standardRevisionDashboardPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserDashboardStandardRevision.cs'
$dataLoaderPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserDataLoader.cs'
$atomicFileServicePath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserAtomicFileService.cs'
$uniqueJsonReportStorePath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserUniqueJsonReportStore.cs'
$requestConcurrencyPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserRequestConcurrencyService.cs'
$requestFileTransactionPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserRequestFileTransactionService.cs'
$projectCatalogServicePath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserProjectCatalogService.cs'
$projectCatalogDashboardPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserDashboardProjectCatalog.cs'
$trackingPersistencePath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserTrackingPersistenceService.cs'
$elementChangeTrackingPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserElementChangeTrackingService.cs'
$elementTrackingTransitionPolicyPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserElementTrackingTransitionPolicy.cs'
$fileGuardPathMatcherPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserFileGuardPathMatcher.cs'
$fileGuardDisciplinePath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserFileGuardDisciplineService.cs'
$fileGuardExcelPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserFileGuardExcelService.cs'
$automaticModelCheckPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserAutomaticModelCheckService.cs'
$automaticModelCheckProgressPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserAutomaticModelCheckProgressWindow.cs'
$nestedOnlyRuntimePath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserNestedOnlyPlacementRuntimeService.cs'
if (Test-Path -LiteralPath $standardRevisionServicePath) {
	Assert-Contains $standardRevisionServicePath 'SHA256-SAMPLED-V1' 'same-stamp Standard RVT content revision hashing'
	Assert-Contains $standardRevisionServicePath 'GetFileInformationByHandle' 'file-identity path alias matching'
	Assert-Contains $standardRevisionServicePath 'WNetGetUniversalName' 'mapped-drive to UNC path normalization'
	Assert-Contains $standardRevisionServicePath 'FamilyBrowserAtomicFileService.Promote(temporary, path)' 'recoverable atomic Standard RVT revision manifest replacement'
	Assert-Contains $standardRevisionServicePath 'public bool BlocksStandardUse' 'Standard RVT stale-source blocking state'
} else {
	Add-Failure "Missing Standard RVT revision service: $standardRevisionServicePath"
}
if (Test-Path -LiteralPath $fileGuardPathMatcherPath) {
	Assert-Contains $fileGuardPathMatcherPath 'WNetGetConnection' 'shared File Guard expands mapped drives without enumerating the network'
	Assert-Contains $fileGuardPathMatcherPath 'BuildStablePolicyPathKey' 'File Guard settings deduplicate paths with the runtime identity model'
	Assert-Contains $fileGuardPathMatcherPath 'FamilyBrowserPathIdentityService.GetFileIdentity' 'File Guard aliases require physical file identity evidence'
	Assert-Contains $fileGuardPathMatcherPath 'FamilyBrowserPathIdentityService.GetComparableIdentity' 'shared File Guard uses final file identity as an alias fallback'
	Assert-Contains $fileGuardPathMatcherPath 'ConservativeDuplicatePathIdentity' 'duplicate physical targets preserve a conservative guard decision'
	Assert-Contains $fileGuardPathMatcherPath 'IdentityUncertain' 'transient mapped or UNC identity failure is represented explicitly'
	Assert-Contains $fileGuardPathMatcherPath 'ConservativeAmbiguousWorksharedNamePendingIdentity' 'ambiguous workshared identity uncertainty preserves the strictest guard'
	Assert-Contains $fileGuardPathMatcherPath 'MergeConservativeTargets' 'shared duplicate-policy merge never weakens restrictions'
	Assert-NotContains $fileGuardPathMatcherPath 'UNC-RELATIVE|' 'endpoint-neutral UNC strings must not identify unrelated servers'
	Assert-NotContains $fileGuardPathMatcherPath 'BuildEndpointNeutralUncKey' 'unsafe endpoint-neutral UNC helper is removed'
	Assert-Contains $fileGuardPathMatcherPath 'UniqueWorksharedNameFallback' 'workshared filename fallback is explicit and unique'
	Assert-Contains $fileGuardPathMatcherPath 'AmbiguousWorksharedName' 'ambiguous same-name policies fail without selecting the first target'
	Assert-Contains $fileGuardPathMatcherPath 'bool hasCentralPath = HasUsablePathCandidate(context.CentralPath)' 'File Guard uncertainty is based on central identity rather than an unrelated local path'
	Assert-Contains $fileGuardPathMatcherPath '!context.IsWorkshared || (HasUsablePathCandidate(context.CentralPath) && !identityUncertain)' 'filename fallback is limited to workshared identity uncertainty'
} else {
	Add-Failure "Missing shared File Guard path matcher: $fileGuardPathMatcherPath"
}
if (Test-Path -LiteralPath $fileGuardDisciplinePath) {
	Assert-Contains $fileGuardDisciplinePath 'ResolveAssignedDiscipline' 'File Guard resolves one assigned trade per RVT'
	Assert-Contains $fileGuardDisciplinePath 'GetSelectableSlots' 'File Guard trade selector uses registered standard slots'
} else {
	Add-Failure "Missing File Guard discipline service: $fileGuardDisciplinePath"
}
if (Test-Path -LiteralPath $fileGuardExcelPath) {
	Assert-Contains $fileGuardExcelPath '"discipline", "trade", "field", "공종", "분야"' 'File Guard Excel accepts English and Korean trade columns'
	Assert-Contains $fileGuardExcelPath 'FamilyBrowserFileGuardDisciplineService.ResolveSlot' 'File Guard Excel rejects unregistered trade assignments'
	Assert-Contains $fileGuardExcelPath 'Discipline is required for every RVT row.' 'File Guard Excel requires an explicit trade on every RVT row'
	Assert-Contains $fileGuardExcelPath 'BuildStablePolicyPathKey' 'File Guard Excel deduplicates physical path aliases'
	Assert-Contains $fileGuardExcelPath 'MergeConservativeTargets' 'File Guard Excel duplicate rows preserve the strictest policy'
} else {
	Add-Failure "Missing File Guard Excel service: $fileGuardExcelPath"
}
if (Test-Path -LiteralPath $automaticModelCheckPath) {
	Assert-Contains $automaticModelCheckPath 'FamilyBrowserFileGuardPathMatcher.FindMatchingTarget' 'automatic model check runs only for a matching File Guard RVT'
	Assert-Contains $automaticModelCheckPath 'FamilyBrowserFileGuardDisciplineService.ResolveAssignedDiscipline' 'automatic model check selects the assigned trade'
	Assert-Contains $automaticModelCheckPath 'allowLegacyFallback: false' 'automatic model check never borrows the currently selected trade for a legacy blank row'
	Assert-MinimumOccurrences $automaticModelCheckPath 'ProjectSnapshotStore.TryLoadLatestProjectScan' 2 'automatic model check reuses matching cached comparisons before scanning'
	Assert-Contains $automaticModelCheckPath 'ProjectSnapshotCaptureService.Capture' 'automatic model check captures stale or missing project snapshots'
	Assert-Contains $automaticModelCheckPath 'ProjectSnapshotStore.SaveLatestProjectScan' 'automatic model check publishes its comparison cache'
	Assert-Contains $automaticModelCheckPath 'ProjectSnapshotStore.TryAcquireProjectPublicationLock' 'automatic model check uses the shared manual/automatic publication lock'
	Assert-MinimumOccurrences $automaticModelCheckPath 'FamilyBrowserStandardRevisionService.Probe' 4 'automatic model check validates the Standard RVT before, after lock, after capture, and before publication'
	Assert-Contains $automaticModelCheckPath 'FamilyBrowserStandardRevisionService.IsSameCurrentRevision' 'automatic model check rejects a Standard RVT change during capture'
	Assert-Contains $automaticModelCheckPath 'ProjectSnapshotStore.CanPublishSharedProjectState' 'automatic model check rejects dirty, unsynchronized, or stale workshared writers'
	Assert-Contains $automaticModelCheckPath 'Last reason: ' 'automatic model check deferred status preserves the actual retry cause'
	Assert-WindowContains $automaticModelCheckPath 'public static bool ProcessPending' 'PruneInvalidRequests();' 'automatic model check removes closed or invalid document requests during idle processing' 12
	Assert-Contains $automaticModelCheckPath 'standardRevisionBeforePublication' 'automatic model check revalidates the Standard RVT after comparison construction'
	Assert-Contains $automaticModelCheckPath 'finalPublicationReason' 'automatic model check performs a final project publication readiness check'
	Assert-Contains $automaticModelCheckPath 'SeedFromProjectSnapshot' 'automatic model check seeds exact nested-only fingerprints from verified snapshots'
	Assert-Contains $automaticModelCheckPath 'RetryWindow = TimeSpan.FromMinutes(30.0)' 'automatic model check waits by elapsed time instead of a small idle-tick count'
	Assert-Contains $automaticModelCheckPath 'WriteStatus(statusWorkspaceRoot, document, statusToWrite)' 'automatic model check persists non-success and deferred outcomes'
	Assert-Contains $automaticModelCheckPath 'FamilyBrowserAutomaticModelCheckProgressWindow.Begin' 'automatic model check exposes visible progress during synchronous Revit API capture'
	Assert-MinimumOccurrences $automaticModelCheckPath 'dirtyMarkerAtCheckStart == null && cached != null && cached.Success' 2 'unresolved protected-content markers bypass automatic comparison cache reuse'
	Assert-Contains $automaticModelCheckPath 'ClearCurrentModelCheckRequired(document, dirtyMarkerAtCheckStart)' 'automatic check clears only the marker generation it actually inspected'
	Assert-Contains $automaticModelCheckPath 'MapProgress(current, total, 5, 80)' 'automatic model check maps capture progress into a stable end-to-end range'
} else {
	Add-Failure "Missing automatic model check service: $automaticModelCheckPath"
}
if (Test-Path -LiteralPath $automaticModelCheckProgressPath) {
	Assert-Contains $automaticModelCheckProgressPath 'RefreshVisibleSurface' 'automatic model check progress repaints without a background Revit API call'
	Assert-Contains $automaticModelCheckProgressPath 'Update();' 'automatic model check progress uses synchronous paint updates'
	Assert-NotContains $automaticModelCheckProgressPath 'Application.DoEvents' 'automatic model check progress does not pump nested Revit commands'
} else {
	Add-Failure "Missing automatic model check progress window: $automaticModelCheckProgressPath"
}
if (Test-Path -LiteralPath $nestedOnlyRuntimePath) {
	Assert-Contains $nestedOnlyRuntimePath 'ResolveSlots(policy, discipline)' 'nested-only placement guard scopes candidates to the assigned trade'
	Assert-Contains $nestedOnlyRuntimePath 'BuildCandidateCacheKey(discipline, slots)' 'nested-only placement cache is isolated by trade'
	Assert-Contains $nestedOnlyRuntimePath 'SeedFromProjectSnapshot' 'nested-only placement verification reuses a verified precise project snapshot'
	Assert-Contains $nestedOnlyRuntimePath 'RemovePendingFamily' 'seeded nested-only families leave no avoidable verification window'
} else {
	Add-Failure "Missing nested-only placement runtime: $nestedOnlyRuntimePath"
}
if (Test-Path -LiteralPath $elementChangeTrackingPath) {
	Assert-Contains $elementChangeTrackingPath 'public static class FamilyBrowserElementChangeTrackingService' 'shared project element change tracker'
	Assert-Contains $elementChangeTrackingPath 'e.GetAddedElementIds()' 'created element event capture'
	Assert-Contains $elementChangeTrackingPath 'e.GetModifiedElementIds()' 'modified element event capture'
	Assert-Contains $elementChangeTrackingPath 'e.GetDeletedElementIds()' 'deleted element event capture'
	Assert-Contains $elementChangeTrackingPath 'TransactionUndone' 'undo-aware activity reduction'
	Assert-Contains $elementChangeTrackingPath 'TransactionRedone' 'redo-aware activity reduction'
	Assert-Contains $elementChangeTrackingPath 'SynchronizingWithCentral' 'incoming synchronization change attribution guard'
	Assert-Contains $elementChangeTrackingPath 'StartReloadLatestBridge' 'version-tolerant Reload Latest event bridge'
	Assert-Contains $elementChangeTrackingPath 'DocumentReloadingLatest' 'Reload Latest begin event suppression'
	Assert-Contains $elementChangeTrackingPath 'RebaseSessionAfterExternalUpdate' 'incoming Reload Latest baseline rebase'
	Assert-Contains $elementChangeTrackingPath 'private static bool CaptureBaseline' 'baseline capture reports failure instead of returning a partial snapshot'
	Assert-Contains $elementChangeTrackingPath 'states.Clear();' 'failed baseline capture discards partial model state'
	Assert-Contains $elementTrackingTransitionPolicyPath 'baselineCapturedLate && hasActiveActivity && wasDeleted' 'late-baseline deletions remain visible with incomplete metadata'
	Assert-Contains $elementChangeTrackingPath 'FamilyName = display == null ? string.Empty : display.FamilyName' 'deleted history retains the last known Family name'
	Assert-Contains $elementChangeTrackingPath 'TypeName = display == null ? string.Empty : display.TypeName' 'deleted history retains the last known Type name'
	Assert-Contains $elementChangeTrackingPath 'ResolveDocumentPolicyEnabledCore' 'element tracking evaluates file-specific scope for every document'
	Assert-Contains $elementChangeTrackingPath 'IsProjectElementTrackingScopeEnabled' 'element tracking reuses normalized File Guard target matching'
	Assert-Contains $elementChangeTrackingPath 'PersistElementChangeCommits' 'successful commit persistence route'
	Assert-Contains $elementChangeTrackingPath 'Exact attribution requires the add-in on every editing workstation.' 'multi-client attribution boundary'
} else {
	Add-Failure "Missing project element change tracking service: $elementChangeTrackingPath"
}
if (Test-Path -LiteralPath $trackingPersistencePath) {
	Assert-Contains $trackingPersistencePath 'PersistElementChangeCommits' 'element change local write-ahead persistence'
	Assert-Contains $trackingPersistencePath 'BuildSpoolFolder("ElementChanges")' 'element change offline spool'
	Assert-Contains $trackingPersistencePath 'ElementChangeHistory' 'immutable per-project element change history'
	Assert-Contains $trackingPersistencePath 'LoadImmutableElementChangeCommits' 'immutable per-project element change history reader'
}
if (Test-Path -LiteralPath $standardRevisionDashboardPath) {
	Assert-Contains $standardRevisionDashboardPath 'QueueStandardRevisionProbe' 'automatic Standard RVT source probe'
	Assert-Contains $standardRevisionDashboardPath 'ApplyStandardRevisionBlockIfNeeded' 'automatic stale-source UI blocking'
	Assert-Contains $standardRevisionDashboardPath 'Confirmed Standard RVT Change History' 'detailed Standard RVT history UI'
	Assert-Contains $standardRevisionDashboardPath 'TryLoadRegistrationForSlot' 'active-slot revision state refresh'
	Assert-WindowContains $standardRevisionDashboardPath 'private static bool StandardRevisionStatesEquivalent' 'left.SnapshotPath' 'remote Standard snapshot replacement invalidates an open dashboard' 24
	Assert-WindowContains $standardRevisionDashboardPath 'private static bool StandardRevisionStatesEquivalent' 'left.SnapshotAtUtc' 'same-source Standard rescan capture time invalidates an open dashboard' 24
} else {
	Add-Failure "Missing Standard RVT revision dashboard integration: $standardRevisionDashboardPath"
}
if (Test-Path -LiteralPath $dataLoaderPath) {
	Assert-Contains $dataLoaderPath 'FamilyBrowserStandardRevisionService.Probe' 'load/apply revision guard uses shared source probe'
	Assert-Contains $dataLoaderPath 'FamilyBrowserAtomicFileService.Promote(tempPath, destinationPath)' 'V2 artifacts use recoverable same-folder promotion'
	Assert-Contains $dataLoaderPath 'AcquireManifestMutationLock' 'V2 manifest read-modify-write is serialized across processes'
	Assert-Contains $dataLoaderPath 'FileShare.None' 'V2 manifest lock is exclusive across Revit sessions'
} else {
	Add-Failure "Missing Family Browser data loader: $dataLoaderPath"
}
if (Test-Path -LiteralPath $atomicFileServicePath) {
	Assert-Contains $atomicFileServicePath 'File.Replace(temporaryPath, destinationPath, null, true)' 'atomic file service prefers native replacement'
	Assert-Contains $atomicFileServicePath 'File.Move(destinationPath, backupPath)' 'atomic file service preserves the committed file before fallback promotion'
	Assert-Contains $atomicFileServicePath 'File.Move(backupPath, destinationPath)' 'atomic file service restores the committed file after failed promotion'
	Assert-Contains $atomicFileServicePath 'CreateSiblingTemporaryPath' 'short sibling temporary paths avoid legacy MAX_PATH suffix expansion'
	Assert-Contains $atomicFileServicePath 'CreateSiblingBackupPath' 'short sibling backup paths avoid legacy MAX_PATH suffix expansion'
	Assert-Contains $atomicFileServicePath 'Substring(0, 8)' 'atomic auxiliary names remain short for .NET Framework hosts'
	Assert-MethodNotContains $atomicFileServicePath 'Promote' 'File.Copy(temporaryPath, destinationPath, true)' 'atomic fallback must not expose an in-place partial overwrite'
} else {
	Add-Failure "Missing shared atomic file service: $atomicFileServicePath"
}
if (Test-Path -LiteralPath $uniqueJsonReportStorePath) {
	Assert-Contains $uniqueJsonReportStorePath 'yyyyMMdd-HHmmss-fffffff' 'audit report filenames use sub-second UTC precision'
	Assert-Contains $uniqueJsonReportStorePath 'Guid.NewGuid()' 'simultaneous audit reports cannot share a filename'
	Assert-Contains $uniqueJsonReportStorePath 'FamilyBrowserAtomicFileService.Promote' 'audit reports are atomically promoted'
	Assert-Contains $uniqueJsonReportStorePath 'stream.Flush(true)' 'audit report bytes are flushed before publication'
} else {
	Add-Failure "Missing unique JSON report store: $uniqueJsonReportStorePath"
}
if (Test-Path -LiteralPath $requestConcurrencyPath) {
	Assert-Contains $requestConcurrencyPath 'FileShare.None' 'request-scoped lock excludes concurrent writers'
	Assert-Contains $requestConcurrencyPath 'EnsureExpectedRevision' 'request optimistic revision validator'
	Assert-Contains $requestConcurrencyPath 'ComputeFileToken' 'legacy request content change token'
	Assert-Contains $requestConcurrencyPath 'FamilyBrowserRequestConflictException' 'typed request conflict result'
} else {
	Add-Failure "Missing request concurrency service: $requestConcurrencyPath"
}
if (Test-Path -LiteralPath $requestFileTransactionPath) {
	Assert-Contains $requestFileTransactionPath 'CopyContentAddressed' 'content-addressed request attachment copy'
	Assert-Contains $requestFileTransactionPath 'CopyAndHash' 'single-pass request attachment copy and hash'
	Assert-Contains $requestFileTransactionPath 'File.Move(temporaryPath, storedPath)' 'same-folder attachment promotion'
	Assert-Contains $requestFileTransactionPath 'WriteImmutableText' 'immutable request deletion audit writer'
	Assert-Contains $requestFileTransactionPath 'File.Move(temporaryPath, path)' 'immutable audit same-folder promotion'
	Assert-Contains $requestFileTransactionPath 'FileMode.CreateNew' 'immutable audit cannot overwrite a prior event'
} else {
	Add-Failure "Missing request attachment transaction service: $requestFileTransactionPath"
}
if (Test-Path -LiteralPath $projectCatalogServicePath) {
	Assert-Contains $projectCatalogServicePath 'GetFamilySymbolIds()' 'project catalog family type name capture'
	Assert-Contains $projectCatalogServicePath 'WhereElementIsElementType()' 'project catalog system type name capture'
	Assert-Contains $projectCatalogServicePath 'ProjectCatalogs' 'managed project catalog storage'
	Assert-Contains $projectCatalogServicePath 'AcceptedSnapshotPath' 'accepted project catalog baseline separation'
	Assert-Contains $projectCatalogServicePath 'LastObservedSnapshotPath' 'last observed project catalog separation'
	Assert-Contains $projectCatalogServicePath 'OperationLogs' 'Family Browser operation attribution'
	Assert-Contains $projectCatalogServicePath 'LoadImmutableOperationEntries(workspaceRoot, 10000)' 'immutable Family Browser operation attribution'
	Assert-Contains $projectCatalogServicePath 'FamilyBrowserAtomicFileService.Promote(temporary, path)' 'project catalog snapshots use recoverable atomic promotion'
	Assert-Contains $projectCatalogServicePath '!string.IsNullOrWhiteSpace(operation.TypeName)' 'exact Family Type operation attribution'
	Assert-Contains $projectCatalogServicePath 'ExternalUntracked' 'external or untracked change classification'
	Assert-WindowContains $projectCatalogServicePath 'public static FamilyBrowserProjectCatalogState Observe' 'ProjectSnapshotStore.CanPublishSharedProjectState' 'automatic project catalog observation rejects uncommitted or stale workshared state' 28
	Assert-WindowContains $projectCatalogServicePath 'public static FamilyBrowserProjectCatalogState AcceptCurrent' 'ProjectSnapshotStore.CanPublishSharedProjectState' 'manual project catalog acceptance rejects uncommitted or stale workshared state' 28
	Assert-WindowContains $projectCatalogServicePath 'private static FamilyBrowserProjectCatalogState PersistSnapshot' 'ProjectSnapshotStore.CanPublishSharedProjectState' 'project catalog persistence revalidates worksharing state inside its publication lock' 32
	Assert-MinimumOccurrences $projectCatalogServicePath 'ProjectSnapshotStore.CanPublishSharedProjectState' 5 'project catalog checks before capture and at both persistence boundaries'
	Assert-WindowContains $projectCatalogServicePath 'public static FamilyBrowserProjectCatalogState LoadLatestState' 'AcquireCatalogLock(folder)' 'project catalog state reads share the cross-process publication lock' 40
	Assert-Contains $projectCatalogServicePath 'public static bool IsPublishedObservationState' 'project catalog distinguishes durable observations from deferred or failed attempts'
	Assert-Contains $projectCatalogServicePath 'PublicationDeferred' 'deferred shared publication is represented explicitly'
	Assert-NotContains $projectCatalogServicePath 'EditFamily(' 'lightweight project catalog must not open family documents'
} else {
	Add-Failure "Missing project catalog tracking service: $projectCatalogServicePath"
}
if (Test-Path -LiteralPath $trackingPersistencePath) {
	Assert-Contains $trackingPersistencePath 'PendingTracking' 'local write-ahead tracking spool'
	Assert-Contains $trackingPersistencePath 'PersistOperationEntries' 'durable Family Browser operation persistence'
	Assert-Contains $trackingPersistencePath 'PersistStandardCandidateEntries' 'durable Standard RVT candidate persistence'
	Assert-Contains $trackingPersistencePath 'LoadImmutableOperationEntries' 'immutable Family Browser operation reader'
	Assert-NotContains $trackingPersistencePath 'DateTimeStyles.RoundtripKind |' 'invalid RoundtripKind DateTimeStyles combination'
} else {
	Add-Failure "Missing durable tracking persistence service: $trackingPersistencePath"
}
if (Test-Path -LiteralPath $projectCatalogDashboardPath) {
	Assert-Contains $projectCatalogDashboardPath 'project-catalog-observe-auto' 'automatic project catalog observation route'
	Assert-Contains $projectCatalogDashboardPath 'project-catalog-check' 'manual project catalog check route'
	Assert-Contains $projectCatalogDashboardPath 'project-catalog-accept' 'project catalog baseline acceptance route'
	Assert-Contains $projectCatalogDashboardPath 'AppendHomeProjectCatalogBoard' 'project catalog Home board'
	Assert-Contains $projectCatalogDashboardPath 'ShowDashboardResultWithExcelExport' 'project catalog HTML result and user-triggered Excel export'
	Assert-Contains $projectCatalogDashboardPath 'ShowProjectElementChangeHistory' 'project element change history viewer'
	Assert-Contains $projectCatalogDashboardPath '권한 / 차단에 등록되고 요소 변경 추적이 체크된 RVT 파일만 새 이력이 기록됩니다.' 'history windows disclose the per-RVT tracking scope'
	Assert-Contains $projectCatalogDashboardPath 'activeProjectTrackingEnabled' 'all-project history does not synthesize an unregistered active project'
	Assert-Contains $projectCatalogDashboardPath 'LoadImmutableElementChangeCommitResult' 'project element change history reader integration'
	Assert-Contains $projectCatalogDashboardPath 'Excel is created only when Export Excel is selected.' 'element history does not create automatic diagnostics workbooks'
	Assert-Contains $projectCatalogDashboardPath 'previewIndex >= 20' 'recent element changes are previewed in the result window with a bounded row count'
	Assert-Contains $projectCatalogDashboardPath 'pendingTrackingPill' 'offline tracking queue header warning'
	Assert-Contains $projectCatalogDashboardPath 'pendingTrackingQueue' 'offline tracking queue Home board'
	Assert-Contains $projectCatalogDashboardPath 'FamilyBrowserTrackingPersistenceService.GetPendingCount()' 'offline tracking queue count source'
	Assert-Contains $projectCatalogDashboardPath '_pendingProjectCatalogTrigger = string.Empty;' 'rejected project-catalog dispatch clears the pending trigger'
	Assert-Before $projectCatalogDashboardPath '_modelessActionDispatcher("project-catalog-observe-auto")' '_lastProjectCatalogObservationUtc = DateTime.UtcNow;' 'project catalog throttle begins only after successful ExternalEvent dispatch'
} else {
	Add-Failure "Missing project catalog dashboard integration: $projectCatalogDashboardPath"
}

$detailedSystemTypeCapturePath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\SystemTypeDetailedComponentSnapshot.cs'
if (Test-Path -LiteralPath $detailedSystemTypeCapturePath) {
	Assert-Contains $detailedSystemTypeCapturePath '"RailingType"' 'Railing detailed component capture support'
	Assert-Contains $detailedSystemTypeCapturePath '"StairsType"' 'Stair detailed component capture support'
	Assert-Contains $detailedSystemTypeCapturePath '"BalusterPlacement"' 'Railing baluster placement capture'
	Assert-Contains $detailedSystemTypeCapturePath 'GetProperties(BindingFlags.Instance | BindingFlags.Public)' 'Railing and Stair public reference-property capture'
	Assert-Contains $detailedSystemTypeCapturePath 'CaptureElementReference' 'Railing and Stair referenced type resolution'
	Assert-Contains $detailedSystemTypeCapturePath '"GetBalusterCount"' 'Railing indexed baluster capture'
	Assert-Contains $detailedSystemTypeCapturePath 'BuiltInParameter.AUTO_PANEL_WALL' 'curtain wall default panel reference capture'
	Assert-Contains $detailedSystemTypeCapturePath 'BuiltInParameter.AUTO_PANEL' 'curtain system default panel reference capture'
	Assert-Contains $detailedSystemTypeCapturePath 'CurtainPanelDependencies' 'mandatory curtain panel dependency group'
	Assert-Contains $detailedSystemTypeCapturePath '"PanelType"' 'direct Curtain Panel system type capture support'
	Assert-Contains $detailedSystemTypeCapturePath 'CaptureElement(context, directPanelType' 'direct Curtain Panel content fingerprint capture'
	Assert-Contains $detailedSystemTypeCapturePath 'CaptureFamilySymbolTypeParameters' 'Railing, Stair, and curtain dependent family type parameter capture'
	Assert-Contains $detailedSystemTypeCapturePath 'ResolveParameterValueKind' 'dependent component parameter spec classification'
	Assert-Contains $detailedSystemTypeCapturePath 'ToString("G17", CultureInfo.InvariantCulture)' 'dependent component raw internal-unit precision'
	Assert-Contains $detailedSystemTypeCapturePath 'BuildRequiredCurtainPanelSignature' 'mandatory curtain panel dependency signature'
	Assert-Contains $detailedSystemTypeCapturePath 'captureCompleted = includeDeepLoadableContent;' 'precise-scan-only component fingerprint certification'
} else {
	Add-Failure "Missing detailed System Type capture service: $detailedSystemTypeCapturePath"
}

foreach ($folder in $HostFolders) {
    $hostPath = Join-Path $root "$folder\FamilyBrowserDashboardHtmlForm.cs"
    $assetPath = Join-Path $root "$folder\KKY.FamilyBrowser.DashboardAssets.family-browser-shell.css"
	$comparisonPath = Join-Path $root "$folder\ProjectStandardComparisonService.cs"
	$comparisonModelPath = Join-Path $root "$folder\LoadableFamilyComparisonItem.cs"
	$auditScenarioPath = Join-Path $root "$folder\FamilyBrowserDashboardAuditScenario.cs"
	$standardPolicyPath = Join-Path $root "$folder\FamilyBrowserStandardPolicy.cs"
	$standardPolicyStorePath = Join-Path $root "$folder\FamilyBrowserStandardPolicyStore.cs"
	$userSettingsPath = Join-Path $root "$folder\FamilyBrowserUserSettingsStore.cs"
	$bridgeHandlerPath = Join-Path $root "$folder\FamilyBrowserRevitBridgeExternalEventHandler.cs"
	$semanticCapturePath = Join-Path $root "$folder\SystemTypeSemanticCaptureService.cs"
	$registrationPath = Join-Path $root "$folder\StandardLibraryRegistrationService.cs"
	$projectCapturePath = Join-Path $root "$folder\ProjectSnapshotCaptureService.cs"
	$changeCandidatePath = Join-Path $root "$folder\StandardRvtChangeCandidateService.cs"
	$changeCandidateModelPath = Join-Path $root "$folder\StandardRvtChangeCandidateEntry.cs"
	$preflightPath = Join-Path $root "$folder\SystemTypePreflightBuilderService.cs"
	$applyPath = Join-Path $root "$folder\SystemTypeApplyExecutionService.cs"
	$supportPolicyPath = Join-Path $root "$folder\SystemTypeSupportPolicyService.cs"
	$systemFingerprintPath = Join-Path $root "$folder\SystemTypeFingerprintService.cs"
	$systemDetailSummaryPath = Join-Path $root "$folder\SystemTypeDetailSummaryService.cs"
	$standardSnapshotPath = Join-Path $root "$folder\StandardLibrarySnapshot.cs"
	$standardRegistryPath = Join-Path $root "$folder\StandardLibraryRegistryStore.cs"
	$projectSnapshotStorePath = Join-Path $root "$folder\ProjectSnapshotStore.cs"
	$requestStorePath = Join-Path $root "$folder\FamilyBrowserRequestStore.cs"
	$requestRecordPath = Join-Path $root "$folder\FamilyBrowserRequestRecord.cs"
	$requestAttachmentRecordPath = Join-Path $root "$folder\FamilyBrowserRequestAttachmentFile.cs"
	$standardListServicePath = Join-Path $root "$folder\FamilyBrowserStandardListService.cs"
	$fileGuardUiPath = Join-Path $root "$folder\FileGuardHtmlConfigurationForm.cs"
	$fileGuardTargetPath = Join-Path $root "$folder\FamilyBrowserFileGuardTarget.cs"
	$permissionExcelExportPath = Join-Path $root "$folder\FamilyBrowserPermissionExcelExportService.cs"
	$sheetSelectionUiPath = Join-Path $root "$folder\StandardListSheetSelectionHtmlForm.cs"
	$bootstrapPath = Join-Path $root "$folder\FamilyBrowserDeploymentBootstrapService.cs"
	$appPath = Get-ChildItem -LiteralPath (Join-Path $root $folder) -Recurse -Filter 'App.cs' | Select-Object -First 1 -ExpandProperty FullName
    if (-not (Test-Path -LiteralPath $hostPath)) {
        Add-Failure "Missing host file: $hostPath"
        continue
    }

	Assert-Contains $hostPath 'case "lang-ko":' 'Korean language route'
	Assert-Contains $hostPath 'DashboardProductVersion = FamilyBrowserProductUpdateService.CurrentProductVersion' 'dashboard display uses the canonical product version'
	Assert-Contains $hostPath 'permission-diagnostic-grid' 'Permission/Guard diagnostics use a dedicated full-width row layout'
	Assert-Contains $hostPath 'class=\"diagnostic-detail\" title=\"' 'Permission/Guard long paths expose their full text as a tooltip'
	Assert-Contains $hostPath 'system-type-detail-components/' 'detailed System Type component comparison route'
	Assert-NotContains $hostPath 'project-element-change-tracking/' 'removed global project element tracking route'
	Assert-Contains $hostPath 'case "project-element-change-history":' 'project element change history route'
	Assert-MinimumOccurrences $hostPath 'project-element-change-history' 2 'project element change history route and action'
	Assert-Contains $hostPath 'fbCurrentProjectHistoryTool' 'current-project history sidebar tool'
	Assert-Contains $hostPath 'fbAllProjectHistoryTool' 'all-project history sidebar tool'
	Assert-Contains $hostPath 'T("History", "이력 관리")' 'dedicated history sidebar group'
	Assert-NotContains $hostPath 'AppendProjectElementChangeTrackingSetting' 'removed global tracking settings checkbox'
	Assert-Contains $hostPath 'System Type 상세 구성 요소 비교 (Railing, Stair 등)' 'detailed System Type component comparison setting label'
	Assert-Contains $hostPath 'admin-check-leading' 'IE-safe detailed component checkbox spacing cell'
	Assert-Contains $hostPath 'admin-check-state' 'detailed component enabled/disabled state label'
	Assert-Contains $hostPath 'data-system-component-table=\"true\"' 'detailed System Type component table renderer'
	Assert-Contains $hostPath 'data-system-curtain-panel-table=\\\"true\\\"' 'mandatory curtain panel dependency table renderer'
	Assert-MinimumOccurrences $hostPath 'FamilyBrowserSystemDetailedComponentUnitUi.Script(!IsKorean())' 2 'main and detached dependent-component unit UI injection'
	Assert-Contains $hostPath 'Curtain Panel Dependencies' 'curtain panel dependency localized detail heading'
	Assert-Contains $hostPath 'renderDetachedSystemDetailWithoutComponents' 'detached detail window component table renderer'
	Assert-Contains $hostPath 'TranslateNote(ApplySystemTypeDetailPolicy(row2.ParameterSummary))' 'final System row render enforces detailed-component policy'
	Assert-Contains $standardPolicyPath 'CompareDetailedSystemTypeComponents' 'detailed System Type policy field'
	Assert-Contains $standardPolicyPath 'TrackProjectElementChanges' 'project element change tracking policy field'
	Assert-Contains $standardPolicyStorePath 'IsDetailedSystemTypeComparisonEnabled' 'detailed System Type policy normalization'
	Assert-Contains $standardPolicyStorePath 'IsProjectElementChangeTrackingEnabled' 'project element change tracking policy normalization'
	Assert-NotContains $standardPolicyStorePath 'SetProjectElementChangeTracking' 'removed bulk tracking policy mutation'
	Assert-NotContains $standardPolicyStorePath 'return policy.TrackProjectElementChanges == true;' 'legacy global tracking flag cannot enable unregistered RVTs'
	Assert-Contains $standardPolicyStorePath 'target.TrackElementChangesConfigured' 'legacy File Guard targets migrate to element tracking enabled'
	Assert-Contains $standardPolicyStorePath 'fileScopedTrackingEnabled' 'File Guard save synchronizes the tracking master state'
	Assert-Contains $standardPolicyStorePath 'Discipline = (target.Discipline ?? string.Empty)' 'File Guard policy clones preserve assigned trade'
	Assert-Contains $standardPolicyStorePath 'target.Discipline = (target.Discipline ?? string.Empty).Trim()' 'File Guard policy normalizes assigned trade'
	Assert-Contains $userSettingsPath 'system-type-detail-components.txt' 'per-user detailed System Type comparison preference file'
	Assert-Contains $userSettingsPath 'ResolveDetailedSystemTypeComparisonEnabled' 'per-user detailed System Type comparison resolver'
	Assert-Contains $userSettingsPath 'SaveDetailedSystemTypeComparisonEnabled' 'per-user detailed System Type comparison persistence'
	Assert-Contains $userSettingsPath 'ClearDetailedSystemTypeComparisonPreference' 'per-user detailed System Type comparison reset'
	Assert-Contains $hostPath 'FamilyBrowserUserSettingsStore.SaveDetailedSystemTypeComparisonEnabled(enabled);' 'detailed System Type setting is persisted before policy refresh'
	Assert-Contains $hostPath '_auditDetailedSystemTypeComparisonEnabled.HasValue' 'UI audit detailed comparison preference isolation'
	Assert-Contains $auditScenarioPath '_auditDetailedSystemTypeComparisonEnabled = scenario.CompareDetailedSystemTypeComponents;' 'UI audit detailed comparison preference initialization'
	Assert-Contains $bridgeHandlerPath 'FamilyBrowserUserSettingsStore.ResolveDetailedSystemTypeComparisonEnabled' 'Revit bridge uses persisted detailed comparison preference'
	Assert-Contains $semanticCapturePath 'SystemTypeDetailedComponentSnapshotService.Capture' 'Railing and Stair detailed component capture'
	Assert-Contains $semanticCapturePath '"PanelType"' 'direct Curtain Panel semantic catalog support'
	Assert-MinimumOccurrences $semanticCapturePath 'SupportsRequiredCurtainPanelComponents' 2 'direct Curtain Panel full and selected semantic capture'
	Assert-Contains $registrationPath '"PanelType"' 'direct Curtain Panel standard registration catalog support'
	Assert-Contains $registrationPath 'SupportsRequiredCurtainPanelComponents' 'direct Curtain Panel standard registration collector'
	Assert-MinimumOccurrences $projectCapturePath 'SupportsRequiredCurtainPanelComponents' 3 'direct Curtain Panel project full, light, and selected capture'
	Assert-MinimumOccurrences $changeCandidatePath 'SupportsRequiredCurtainPanelComponents' 3 'direct Curtain Panel change tracking'
	Assert-Contains $hostPath 'prepared.StandardRevisionState = prepared.Registration == null ? null : FamilyBrowserStandardRevisionService.Probe' 'startup Standard RVT revision preload'
	Assert-MinimumOccurrences $hostPath 'QueueStandardRevisionProbe(true' 2 'manual and periodic Standard RVT revision refresh'
	Assert-Contains $hostPath 'ApplyStandardRevisionBlockIfNeeded' 'stale Standard RVT dashboard block'
	Assert-Contains $hostPath 'QueueProjectCatalogObservation("BrowserStartup", false)' 'automatic project catalog observation on browser startup'
	Assert-Contains $hostPath 'QueueProjectCatalogObservation("HeaderRefresh", true)' 'manual project catalog observation on header refresh'
	Assert-Contains $hostPath 'AcceptProjectCatalogAfterCurrentModelCheck(doc, projectSnapshot)' 'Current Model Check accepts project catalog baseline'
	Assert-Contains $hostPath 'ReloadProjectCatalogStateAfterCommit' 'save and sync refresh project catalog UI state'
	Assert-Contains $auditScenarioPath 'StandardRvtChanged' 'changed Standard RVT audit scenario input'
	Assert-Contains $auditScenarioPath 'StandardRvtUnavailable' 'unavailable Standard RVT audit scenario input'
	Assert-Contains $auditScenarioPath 'ProjectCatalogBaselineMissing' 'missing project catalog baseline audit scenario input'
	Assert-Contains $auditScenarioPath 'ProjectCatalogChanged' 'changed project catalog audit scenario input'
	Assert-Contains $auditScenarioPath 'ProjectCatalogUntracked' 'external project catalog change audit scenario input'
	if ([string]::IsNullOrWhiteSpace($appPath)) {
		Add-Failure "Missing Revit host App.cs under $folder"
	} else {
		Assert-Contains $appPath 'FamilyBrowserAutomaticModelCheckService.Schedule(document, "DocumentOpened", force: true)' 'registered RVT schedules automatic model check on first open'
		Assert-Contains $appPath 'FamilyBrowserAutomaticModelCheckService.ProcessPending(uiApplication)' 'automatic model check advances on safe Revit idle passes'
		Assert-Contains $appPath 'FamilyBrowserAutomaticModelCheckService.Remove(document)' 'automatic model check releases closed document state'
		Assert-MinimumOccurrences $appPath 'ObserveProjectCatalogAfterCommit(document,' 3 'Save, SaveAs, and Sync conditional project catalog decisions'
		Assert-Contains $appPath 'FamilyBrowserElementChangeTrackingService.ConsumeProjectCatalogObservationRequired(document)' 'tracking-driven project catalog decision'
		Assert-Contains $appPath 'FamilyBrowserProjectCatalogService.Observe(' 'conditional project catalog observation'
		Assert-MinimumOccurrences $appPath 'FamilyBrowserProjectCatalogService.IsSuccessfulRevitEventStatus' 3 'successful Revit commit status guard for project catalog observation'
		Assert-Contains $appPath 'DocumentOpened += HandleDocumentOpened' 'element tracking baseline begins on document open'
		Assert-Contains $appPath 'StartReloadLatestBridge(application.ControlledApplication' 'element tracker attaches the version-tolerant Reload Latest bridge'
		Assert-Contains $appPath 'DocumentSynchronizingWithCentral += HandleDocumentSynchronizingWithCentral' 'element tracker suppresses incoming synchronization changes'
		Assert-Contains $appPath 'FamilyBrowserElementChangeTrackingService.HandleDocumentChanged' 'element tracker receives Revit document changes'
		Assert-MinimumOccurrences $appPath 'FamilyBrowserElementChangeTrackingService.HandleDocumentCommitted' 3 'Save, SaveAs, and Sync element tracking commit boundaries'
		Assert-Contains $appPath 'FamilyBrowserElementChangeTrackingService.HandleDocumentClosing' 'close-start preserves element activity until closure is final'
		Assert-Contains $appPath 'FamilyBrowserElementChangeTrackingService.HandleDocumentClosed' 'uncommitted element activity discarded only after actual close'
		Assert-Contains $appPath 'RunEventHandlerSafely' 'independent lifecycle services cannot block the element ledger'
	}
	Assert-Contains $standardRegistryPath 'FamilyBrowserStandardRevisionService.RecordBaseline' 'accepted scan records a source revision baseline'
	Assert-Contains $standardRegistryPath 'PublishSnapshotAndActiveRegistration' 'standard snapshot and active registration publish as one coordinated unit'
	Assert-Contains $standardRegistryPath '.standard-library-publication.lock' 'standard publication is serialized across Revit processes'
	Assert-Contains $standardRegistryPath 'WriteTextAtomic' 'standard snapshot and active registration use atomic file promotion'
	Assert-Contains $standardRegistryPath 'yyyyMMdd-HHmmss-fffffff' 'standard snapshot filenames have sub-second uniqueness'
	Assert-Contains $standardRegistryPath 'Guid.NewGuid()' 'standard snapshot filenames cannot collide across processes'
	Assert-WindowContains $standardRegistryPath 'PublishSnapshotAndActiveRegistration' 'FamilyBrowserDataLoader.PublishStandardArtifacts(workspaceRoot, snapshotPath, snapshot);' 'coordinated Standard publication prepares V2 artifacts before switching the active pointer' 34
	Assert-MinimumOccurrences $registrationPath 'PublishSnapshotAndActiveRegistration' 3 'full and partial Standard scans use coordinated publication'
	Assert-NotContains $registrationPath 'StandardLibraryRegistryStore.SaveSnapshot(workspaceRoot, snapshot)' 'scan workflows do not split snapshot and registration publication'
	Assert-Contains $changeCandidatePath 'PendingCandidateSaveBatchesByDocument' 'pending Standard RVT change batch'
	Assert-Contains $changeCandidatePath 'CommitPendingCandidateEntries' 'save-success Standard RVT change commit'
	Assert-Contains $changeCandidatePath 'AppendImmutableCandidates' 'immutable Standard RVT change history'
	Assert-Contains $changeCandidatePath 'FamilyBrowserTrackingPersistenceService.PersistStandardCandidateEntries' 'Standard RVT local write-ahead persistence'
	Assert-Contains $changeCandidatePath 'FamilyBrowserTrackingPersistenceService.PersistOperationEntries' 'Family Browser operation local write-ahead persistence'
	Assert-Contains $changeCandidatePath 'RestorePendingCandidateBatch' 'failed Standard RVT persistence batch restoration'
	Assert-Contains $changeCandidatePath 'RestorePendingOperationBatch' 'failed operation persistence batch restoration'
	Assert-Contains $changeCandidatePath 'ResolveLoadedFamilyTypeNames' 'exact loaded Family Type operation capture'
	Assert-Contains $changeCandidatePath 'typeEntry.TypeName = typeName' 'Family Type operation entry assignment'
	Assert-Contains $requestStorePath 'FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path)' 'request store uses a short same-folder temporary path'
	Assert-Contains $requestStorePath 'FamilyBrowserAtomicFileService.Promote(tempPath, path)' 'request store uses recoverable network-share promotion'
	Assert-MethodNotContains $requestStorePath 'WriteAllTextAtomic' 'File.Delete(path);' 'request-store fallback must not delete the only committed request before promotion'
	Assert-MinimumOccurrences $requestStorePath 'FamilyBrowserRequestConcurrencyService.Acquire' 3 'request create/save, status, and delete share a request-scoped lock'
	Assert-MinimumOccurrences $requestStorePath 'FamilyBrowserRequestConcurrencyService.EnsureExpectedRevision' 3 'request mutations reject stale revisions and tokens'
	Assert-Contains $requestStorePath 'FamilyBrowserRequestFileTransactionService.CopyContentAddressed' 'request attachments use content-addressed idempotent copies'
	Assert-Contains $requestStorePath 'RollbackAttachmentMutation' 'failed pre-commit request save rolls back attachment state'
	Assert-Contains $requestStorePath 'requestCommitted = true;' 'request JSON is the authoritative attachment commit point'
	Assert-Contains $requestStorePath '"RequestAudit", "Deleted"' 'deleted request audit lives outside active requests'
	Assert-Contains $requestStorePath '"DeletePrepared"' 'request deletion records a durable prepared snapshot'
	Assert-Contains $requestStorePath '"DeleteCompleted"' 'request deletion records an immutable completion event'
	Assert-Before $requestStorePath 'WriteRequestDeletionAudit(rootPath, record, deletionId, "DeletePrepared"' 'DeleteFileInsideRoot(requestPath, rootPath);' 'request deletion snapshot is written before the active request file is removed'
	Assert-Contains $requestRecordPath 'public long Revision' 'request record optimistic revision'
	Assert-Contains $requestRecordPath 'public string RevisionToken' 'request record optimistic revision token'
	Assert-Contains $requestAttachmentRecordPath 'public string ContentSha256' 'request attachment metadata content hash'
	Assert-Contains $hostPath 'HandleRequestConflict' 'request conflict refresh and user guidance'
	Assert-Contains $hostPath 'expectedRevisionToken' 'request action carries rendered revision token'
	Assert-MinimumOccurrences $hostPath 'ShowRequestAuxiliaryWarning(' 3 'request create and status update surface committed auxiliary metadata warnings'
	Assert-Contains $hostPath 'FamilyBrowserErrorHelp.WriteLog(_workspaceRoot, "Request auxiliary metadata warning"' 'request auxiliary warning writes a diagnostic log'
	Assert-Contains $standardListServicePath 'FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(outputPath)' 'standard list uses a short same-folder temporary path'
	Assert-Contains $standardListServicePath 'FamilyBrowserAtomicFileService.Promote(tempPath, outputPath)' 'standard list replacement preserves the prior committed list'
	Assert-MethodNotContains $standardListServicePath 'WriteAllTextAtomically' 'File.Delete(outputPath)' 'standard list save must not delete the only committed list before promotion'
	Assert-Contains $changeCandidatePath '"History"' 'dated Standard RVT history folder'
	Assert-MethodNotContains $changeCandidatePath 'HandleDocumentSaving' 'AppendCandidates(' 'save-start must not commit Standard RVT history'
	Assert-Contains $changeCandidateModelPath 'RevitUserName' 'Revit user in Standard RVT history'
	Assert-Contains $changeCandidateModelPath 'MachineName' 'machine identity in Standard RVT history'
	Assert-Contains $changeCandidateModelPath 'BeforeFingerprint' 'before fingerprint summary in Standard RVT history'
	Assert-Contains $changeCandidateModelPath 'AfterFingerprint' 'after fingerprint summary in Standard RVT history'
	Assert-Contains $preflightPath 'SupportsRequiredCurtainPanelComponents' 'direct Curtain Panel review preflight lookup'
	Assert-Contains $applyPath 'SupportsRequiredCurtainPanelComponents' 'direct Curtain Panel identity map lookup'
	Assert-Contains $supportPolicyPath '{ "paneltype", "ReviewOnly" }' 'direct Curtain Panel remains review-only'
	Assert-Contains $systemFingerprintPath 'SYSFP|v4' 'detailed System Type fingerprint v4'
	Assert-Contains $systemFingerprintPath 'SYSFP|v5' 'mandatory curtain panel dependency fingerprint v5'
	Assert-Contains $comparisonPath 'BuildDetailedComponentDifferenceSummary' 'detailed System Type difference comparison'
	Assert-Contains $comparisonPath 'BuildCurtainPanelDifferenceSummary' 'mandatory curtain panel dependency difference comparison'
	Assert-NotContains $comparisonPath 'detail = FilterDetailedComponentSections(detail);' 'comparison report must retain optional component detail even when comparison is disabled'
	Assert-NotContains $comparisonPath 'private static string FilterDetailedComponentSections' 'destructive comparison-detail filtering helper'
	Assert-Contains $comparisonPath '"@component-diff"' 'structured dependent-component difference row'
	Assert-Contains $systemDetailSummaryPath '"@component"' 'structured dependent-component summary row'
	Assert-Contains $standardSnapshotPath 'SnapshotSchemaVersion = 9;' 'dependent component parameter metadata snapshot schema'
	Assert-Contains $standardRegistryPath 'SnapshotSchemaVersion < 9' 'legacy snapshot rescan requirement for component units'
	Assert-Contains $auditScenarioPath 'AUDIT_GUARDRAIL' 'Railing detailed component UI fixture'
	Assert-Contains $auditScenarioPath 'AUDIT_CURTAIN_WALL' 'curtain panel dependency UI fixture'
	Assert-Contains $auditScenarioPath 'AUDIT_SYSTEM_PANEL_TYPE' 'direct Curtain Panel type dependency UI fixture'
	Assert-Contains $auditScenarioPath '@component\tcomponents\tRail Height\tLength\t3\t914.4 mm' 'Railing component length unit fixture'
	Assert-Contains $auditScenarioPath '@component\tcurtain-components\tPanel Width\tLength\t4\t1219.2 mm' 'curtain wall panel length unit fixture'
	Assert-Contains $auditScenarioPath '@component\tcurtain-components\tPanel Thickness\tLength\t0.25\t76.2 mm' 'direct PanelType length unit fixture'
	Assert-Contains $auditScenarioPath '@section\tcomponent-differences' 'Railing detailed component difference fixture'
	Assert-Contains $auditScenarioPath '@section\tcurtain-component-differences' 'curtain panel dependency difference fixture'
	Assert-Contains $auditScenarioPath 'CompareDetailedSystemTypeComponents = scenario.CompareDetailedSystemTypeComponents' 'audit fixture supports enabled and disabled component comparison'
	Assert-MinimumOccurrences $hostPath 'FamilyBrowserOverflowTitleScript.Script()' 5 'main, request, selection, standard manager, and detached-detail overflow title service injection'
	Assert-Contains $hostPath 'FamilyBrowserManagedFolderSetupService.TryApplyPersistedOverride(Environment.UserName' 'persisted management-folder override applies before homepage bootstrap'
	Assert-Contains $hostPath 'RefreshManagedFolderAvailabilityState(queueOnboarding: true);' 'first-open management-folder availability check'
	Assert-MinimumOccurrences $hostPath 'QueueManagedFolderOnboardingIfNeeded();' 2 'startup-completion and browser-render management-folder onboarding queues'
	Assert-Contains $hostPath '이 PC에서 홈페이지에 등록된 네트워크 관리폴더 경로를 찾을 수 없습니다.' 'requested missing management-folder message'
	Assert-Contains $hostPath 'TEST 관리폴더로 사용할 내부 네트워크 공유폴더를 선택하세요.' 'TEST management-folder internal-share guidance'
	Assert-Contains $hostPath 'managedFolderRecovery' 'persistent Home management-folder recovery banner'
	Assert-Contains $hostPath 'kkyfb:managed-folder-retry' 'homepage management-folder retry action'
	Assert-Contains $hostPath 'kkyfb:managed-folder-test-setup' 'TEST management-folder setup action'
	Assert-Contains $hostPath 'ProbeHomepageReturnWhileTestFolderActive(showResult: true)' 'top refresh and recovery retry probe homepage while TEST override remains active'
	Assert-Contains $hostPath 'case "managed-folder-switch-homepage":' 'homepage management-folder switch route'
	Assert-Contains $hostPath 'case "managed-folder-migrate-homepage":' 'managed-folder migrate-and-switch route'
	Assert-Contains $bootstrapPath 'ProbeHomepageManagedFolder' 'non-mutating live homepage management-folder probe'
	Assert-Contains $bootstrapPath 'ResolveHomepageBootstrapUrlForProbe' 'homepage probe ignores only the locally persisted disabled token'
	Assert-Contains $bootstrapPath 'FetchText(AddNoCacheQuery(bootstrapUrl))' 'homepage management-folder probe requires a live no-cache response'
	Assert-Contains $hostPath '수동으로 수정, 이동, 이름 변경 또는 삭제하지 마세요.' 'managed-content no-manual-edit warning'
	Assert-Contains $hostPath 'FolderBrowserDialog dialog = new FolderBrowserDialog' 'internal network management-folder selector'
	Assert-Contains $hostPath 'FamilyBrowserManagedFolderSetupService.Configure(_workspaceRoot, dialog.SelectedPath, Environment.UserName)' 'selected management-folder activation route'
	if (Test-Path -LiteralPath $fileGuardUiPath) {
		Assert-Contains $fileGuardUiPath 'FamilyBrowserOverflowTitleScript.Script()' 'file Guard overflow title service injection'
		Assert-Contains $fileGuardUiPath 'BuildDisciplineOptionsHtml' 'File Guard renders a per-RVT trade selector'
		Assert-Contains $fileGuardUiPath 'FamilyBrowserFileGuardExcelService.Import' 'File Guard imports bulk RVT/trade policies from Excel'
		Assert-Contains $fileGuardUiPath 'FamilyBrowserPermissionExcelExportService.ExportFileGuardPolicy' 'File Guard saves an Excel template with current assignments'
		Assert-Contains $fileGuardUiPath 'Discipline = assignedSlot.Discipline' 'File Guard persists the normalized registered trade'
	} else {
		Add-Failure "Missing file Guard HTML form: $fileGuardUiPath"
	}
	Assert-Contains $fileGuardTargetPath 'public string Discipline' 'File Guard target stores its assigned trade'
	Assert-Contains $permissionExcelExportPath '"공종"' 'File Guard Korean Excel export includes the trade column'
	Assert-Contains $permissionExcelExportPath '"Discipline"' 'File Guard English Excel export includes the trade column'
	if (Test-Path -LiteralPath $sheetSelectionUiPath) {
		Assert-Contains $sheetSelectionUiPath 'FamilyBrowserOverflowTitleScript.Script()' 'worksheet selection overflow title service injection'
	} else {
		Add-Failure "Missing worksheet selection HTML form: $sheetSelectionUiPath"
	}
	Assert-Contains $hostPath 'Math.Min(1400, (int)Math.Floor((double)area.Width * 0.93))' 'KKY Tool-matched 1400px startup width with 93 percent working-area cap'
	Assert-Contains $hostPath 'Math.Min(900, (int)Math.Floor((double)area.Height * 0.93))' 'KKY Tool-matched 900px startup height with 93 percent working-area cap'
	Assert-Contains $hostPath 'new Size(Math.Min(1100, startSize.Width), Math.Min(720, startSize.Height))' 'KKY Tool-matched 1100x720 minimum window size'
	Assert-Contains $hostPath 'base.Location = new Point(area.Left + Math.Max(0' 'startup window remains centered on the resolved Revit screen'
	Assert-Contains $hostPath 'RefreshDocumentShellOnlyWithProgress(refreshTitle, refreshMessage, 36, 100, allowStartupPreload: true);' 'startup preload is explicitly allowed for the initial shell render only'
	Assert-Contains $hostPath 'bool allowStartupPreload = false' 'post-action shell refresh defaults to live persisted state'
	Assert-Contains $hostPath 'bool allowPreparedSlotData = false' 'post-action shell refresh does not reuse prepared slot data unless explicitly requested'
	Assert-Contains $hostPath '(allowStartupPreload && _startupPreloadResult != null && _startupPreloadResult.Policy != null) ? _startupPreloadResult.Policy : LoadStandardPolicy()' 'shell refresh reloads live policy after mutations'
	Assert-Contains $hostPath '_allowStartupPreloadDuringShellRefresh = allowStartupPreload;' 'startup preload permission spans the complete shell refresh'
	Assert-Contains $hostPath '_allowPreparedSlotDataDuringShellRefresh = allowPreparedSlotData;' 'trade switch can explicitly reuse prepared slot payloads without reusing stale policy'
	Assert-Contains $hostPath 'if ((!_allowStartupPreloadDuringShellRefresh && !_allowPreparedSlotDataDuringShellRefresh) || _startupPreloadResult == null' 'prepared slot payloads remain gated outside startup and explicit trade switching'
	Assert-Contains $hostPath '_allowStartupPreloadDuringShellRefresh = false;' 'startup preload permission is cleared after every shell refresh'
	Assert-Contains $hostPath '_allowPreparedSlotDataDuringShellRefresh = false;' 'trade-switch prepared-data permission is cleared after every shell refresh'
	Assert-Contains $hostPath 'shell-policy-source:' 'runtime diagnostics record startup versus live refresh source'
	Assert-NotContains $hostPath '_standardPolicy = (_startupPreloadResult == null || _startupPreloadResult.Policy == null) ? LoadStandardPolicy() : _startupPreloadResult.Policy;' 'unconditional stale startup policy reuse'
	$liveMutationMethods = @(
		'SetStandardMode',
		'SetActiveDiscipline',
		'AddStandardDisciplineTarget',
		'RenameActiveStandardDisciplineTarget',
		'DeleteActiveStandardDisciplineTarget',
		'ConfigureStandardListExcel',
		'ClearStandardListExcel',
		'ConfigureSecurityUsers',
		'ConfigurePermissionExcel',
		'ClearPermissionExcel',
		'ConfigureFileGuardPolicy',
		'ClearFileGuardPolicy',
		'ExportPermissionExcelFromRvtFolder',
		'AddCurrentProjectPolicyRule',
		'ClearProjectPolicyRules',
		'ResetRegisteredStandardRvt',
		'ConfigureNetworkRequestStore',
		'ConfigureConnectorRequestStore'
	)
	foreach ($methodName in $liveMutationMethods) {
		Assert-MethodContainsAny $hostPath $methodName @('RefreshDocumentShellOnly(', 'RefreshDashboard(', 'RenderDashboard(') 'persisted-state live refresh contract'
	}
	Assert-NotContains $hostPath '.Wait(' 'UI-thread Task.Wait call'
	Assert-NotContains $hostPath 'Task.Wait(' 'UI-thread static Task.Wait call'
	Assert-Contains $hostPath 'FamilyBrowserDeploymentProjectIdentity projectIdentity = BuildDeploymentProjectIdentity(GetActiveDocument());' 'startup captures local and central Revit project identity on the UI thread'
	Assert-Contains $hostPath 'Task.Factory.StartNew(() => PrepareStartupPreload(activeGeneration, projectIdentity)' 'two-stage background startup preload'
	Assert-Before $hostPath 'FamilyBrowserDeploymentBootstrapService.TryApply(_workspaceRoot, Environment.UserName, force: false, projectIdentity);' 'result.Policy = FamilyBrowserStandardPolicyStore.LoadOrCreate(_workspaceRoot, Environment.UserName);' 'homepage managed path applies before startup policy preload'
	Assert-Contains $hostPath 'prepared.ProjectScan = ProjectSnapshotStore.TryLoadLatestProjectScan(_workspaceRoot, projectIdentity' 'saved project comparison is preloaded off the Revit UI thread'
	Assert-Contains $hostPath 'prepared.ProjectScanCacheStamp = ProjectSnapshotStore.BuildProjectScanCacheStamp(prepared.ProjectScan);' 'prepared project comparison carries a stable cache revision'
	Assert-Contains $hostPath 'ProjectSnapshotStore.TryLoadLatestProjectScan(_workspaceRoot, doc, registration, standardSnapshot)' 'initial dashboard revalidates prepared comparison data against the live document'
	Assert-Contains $hostPath 'v7-trade-switch-project-restore' 'stale trade-switch standard-only row caches are invalidated'
	Assert-Contains $hostPath 'BuildPreparedProjectScanCacheStamp(browseSlots)' 'startup row cache key includes prepared project comparison revisions'
	Assert-Contains $hostPath 'preparedSlot.StandardListCacheStamp' 'startup row cache key includes the prepared standard-list revision'
	Assert-Contains $hostPath 'RefreshDocumentShellOnlyWithProgress(preserveBrowseTargetUiState ? T("Switching check target"' 'list trade switch uses visible progress instead of an unresponsive full refresh'
	Assert-Contains $hostPath 'allowStartupPreload: false, allowPreparedSlotData: true' 'list trade switch keeps live policy while reusing prepared target data'
	Assert-Contains $hostPath 'ProjectSnapshotStore.TryLoadLatestProjectScan(_workspaceRoot, doc, registration, standardSnapshot)' 'non-startup trade restores and validates its saved comparison against the live document'
	Assert-Contains $hostPath 'prepared.ProjectScanCacheStamp = ProjectSnapshotStore.BuildProjectScanCacheStamp(loaded);' 'on-demand trade comparison updates the row-cache revision'
	Assert-Contains $hostPath 'PrimePreparedProjectScanForSelectedSlot(doc, selectedSlot, selectedRegistration);' 'target comparison is primed before row-cache lookup'
	Assert-Before $hostPath 'PrimePreparedProjectScanForSelectedSlot(doc, selectedSlot, selectedRegistration);' 'string cacheKey = BuildModelerAllSlotsCacheKey(doc, browseSlots, null);' 'trade comparison revision is available before row-cache key calculation'
	Assert-Contains $hostPath "currentTreeDiscipline='All'" 'browse tree starts at the current target root instead of a stale trade node'
	Assert-Contains $hostPath 'function resetBrowseTreeFilters()' 'trade switch clears stale family and system tree filters'
	Assert-Contains $hostPath 'function beginBrowseDisciplineSwitch(filter){currentDiscipline=filter;' 'trade chip immediately updates visible rows while host data switches'
	Assert-Contains $hostPath 'function syncBrowseTreeTradeLabel(filter)' 'trade chip immediately synchronizes the visible left-tree trade label'
	Assert-Contains $hostPath 'syncBrowseTreeTradeLabel(filter);resetBrowseTreeFilters();' 'trade label synchronization runs before the tree state is reset'
	Assert-Contains $hostPath 'FamilyTreeRowsForCurrentBrowse(rows)' 'family tree is rebuilt from only the selected standard target'
	Assert-Contains $hostPath 'SystemTreeRowsForCurrentBrowse(rows)' 'system tree is rebuilt from only the selected standard target'
	Assert-Contains $hostPath 'private string ResolveCurrentBrowseTreeLabel(string fallbackLabel, string fallbackKey)' 'left-tree trade label resolves from the selected standard target'
	Assert-MinimumOccurrences $hostPath 'ResolveCurrentBrowseTreeLabel(firstRow.DisciplineLabel, disciplineKey)' 2 'family and system trees use the selected standard target label'
	Assert-Contains $hostPath 'FamilyBrowserStandardLibrarySlot slot = ResolvePolicySlotExact(_standardPolicy, candidate);' 'row trade label resolves the actual row slot before using a canonical fallback'
	Assert-NotContains $hostPath 'FamilyBrowserStandardLibrarySlot slot = ResolvePolicySlot(_standardPolicy, resolvedKey);' 'row trade label must not fall back to the policy-active trade'
	Assert-Contains $hostPath 'data-ready=\' 'trade chip keeps its registered readiness state during immediate feedback'
	Assert-Contains $hostPath 'string switchScript = "beginBrowseDisciplineSwitch(' 'trade chip builds immediate browser-side target feedback'
	Assert-Contains $hostPath 'onclick=\"" + Attr(switchScript)' 'trade chip safely emits the immediate switch handler'
	Assert-NotContains $hostPath 'Project comparison cache was not read on the UI thread. Persistent browser rows are used when available.' 'initial comparison restore skip path'
	Assert-Contains $projectSnapshotStorePath 'TryLoadLatestProjectScan(string workspaceRoot, FamilyBrowserDeploymentProjectIdentity projectIdentity' 'project scan store supports a Revit-free preload identity'
	Assert-Contains $projectSnapshotStorePath 'TryLoadLatestProjectScanCore(string workspaceRoot, ProjectLookupIdentity projectIdentity' 'document and background restore share the same validation core'
	Assert-Contains $projectSnapshotStorePath 'BuildProjectLookupIdentity(FamilyBrowserDeploymentProjectIdentity projectIdentity)' 'background identity preserves project paths for comparison lookup'
	Assert-Contains $hostPath '_startupPreloadResult = null;' 'deferred homepage path change invalidates startup preload data'
	Assert-Contains $hostPath 'private void ClearDashboardDataCaches()' 'standard registration and scan cache invalidation is instance-aware'
	Assert-MethodContainsAny $hostPath 'ClearDashboardDataCaches' @('_startupPreloadResult = null;') 'standard registration and scan changes invalidate prepared trade data'
	Assert-MethodContainsAny $hostPath 'ClearStandardListDashboardCaches' @('_startupPreloadResult = null;') 'approved-list changes invalidate prepared trade data'
	Assert-Contains $hostPath 'private void InvalidatePreparedDashboardDataAfterManagedPolicyChange(string reason)' 'shared managed-policy prepared-data invalidation helper'
	Assert-MinimumOccurrences $hostPath 'InvalidatePreparedDashboardDataAfterManagedPolicyChange(' 8 'all homepage path/profile/security refresh routes invalidate prepared data'
	Assert-Contains $hostPath 'FamilyBrowserDataLoader.Default.LoadSnapshotProjection' 'V2 compact snapshot preload'
	Assert-Contains $hostPath 'StandardListAllowsNestedDifference' 'standard-list exception for a differing nested family'
	Assert-Contains $hostPath 'item.IsNestedLoadableDifference' 'nested difference display gate'
	Assert-Contains $hostPath 'ShouldDisplayNestedLoadableDifference' 'matching nested rows stay hidden while differing nested rows remain visible'
	Assert-Contains $hostPath 'bool selectable = !row.IsNestedLoadableChild' 'nested helper rows cannot be loaded independently'
	Assert-Contains $hostPath 'nestedDifferenceException ? NestedLoadableDifferenceAction(item.Status)' 'nested difference row uses relation-aware parent action'
	Assert-Contains $hostPath 'return T("Nested Family Missing", "하위 패밀리 누락");' 'missing nested child has an explicit status label'
	Assert-Contains $hostPath 'return T("Update parent family", "상위 패밀리 업데이트 필요");' 'missing nested child directs update through its parent family'
	Assert-Contains $hostPath 'string.Equals(status, "nestedmissingfromparent", StringComparison.Ordinal)' 'missing nested child is excluded from fingerprint-capture-failure classification'
	Assert-Contains $hostPath 'bool hasProjectEvidence = !string.IsNullOrWhiteSpace(item.ProjectFingerprint)' 'fingerprint failure requires evidence that a current-project family existed'
	Assert-Contains $hostPath 'public bool HasNestedLoadableDifference' 'dashboard row retains nested difference state'
	Assert-Contains $hostPath 'HasNestedLoadableDifference = item.IsNestedLoadableDifference' 'comparison nested difference state reaches dashboard rows'
	Assert-Contains $hostPath 'row.IsNestedLoadableChild && row.HasNestedLoadableDifference' 'nested review action survives virtual and modeler row rendering'
	Assert-Contains $hostPath 'if (row.HasNestedLoadableDifference)' 'top-level parent remains visible to modelers when a nested family differs'
	Assert-Contains $comparisonPath 'NestedLoadableFamilyDifferencePropagationService.Apply(results);' 'nested-family difference propagation runs after loadable comparison'
	Assert-Contains $comparisonModelPath 'public bool IsNestedLoadableDifference' 'nested-family exception-display model flag'
	Assert-Contains $comparisonModelPath 'public List<string> NestedParentFamilyNames' 'nested-family parent model list'
	Assert-Contains $comparisonModelPath 'public List<string> NestedDifferenceFamilyNames' 'parent nested-difference model list'
	Assert-MinimumOccurrences $auditScenarioPath 'HasNestedLoadableDifference = true' 2 'audit parent and nested child retain nested difference state'
	Assert-Contains $auditScenarioPath 'CountVisibleAuditFamilyRows(rows)' 'performance fixture counts only list-visible family rows'
	Assert-Contains $auditScenarioPath '!row.IsNestedLoadableChild || row.HasNestedLoadableDifference' 'performance fixture excludes matching nested helpers from its target count'
	Assert-Contains $hostPath 'HydrateSelectedRowDetailFromV2' 'lazy selected-detail hydration'
	Assert-Contains $hostPath 'AppendSystemRowsFromComparison(ProjectStandardComparisonReport report, StandardLibrarySnapshot standardSnapshot' 'comparison System rows receive the standard detail source'
	Assert-Contains $hostPath 'DetailKey = standardSystemItem == null ? string.Empty : standardSystemItem.BrowserDetailKey' 'comparison System rows retain their V2 detail key'
	Assert-Contains $hostPath 'DetailSourceKey = standardSnapshot == null ? string.Empty : standardSnapshot.BrowserManifestSourceKey' 'comparison System rows retain their V2 detail source key'
	Assert-Contains $hostPath 'HasOptionalSystemTypeComponentDetail(currentDetailSummary)' 'legacy stripped comparison detail is repaired from V2 on selection'
	Assert-Contains $hostPath 'data-detailkey=' 'row carries stable detail key instead of full detail payload'
	Assert-Contains $hostPath 'window.rowWindowSize=150' '150-row browser window'
	Assert-Contains $hostPath 'window.changeRowWindow=function(delta)' 'browser row-window navigation'
	Assert-Contains $hostPath 'BuildVirtualRowPayloadJson' 'compact virtual row payload builder'
	Assert-Contains $hostPath '<tbody id=\"familiesBody\">' 'family virtual row injection body'
	Assert-Contains $hostPath '<tbody id=\"systemsBody\">' 'system virtual row injection body'
	Assert-Contains $hostPath 'AppendDashboardScriptAsset(sb, "family-browser-row-window.js")' 'shared virtual row runtime injection'
	Assert-Contains $hostPath '<div id=\"dashboardStatusBar\" class=\"statusbar\"><span id=\"dashboardStatusText\"' 'browser status text and paging use separate regions'
	Assert-Contains $hostPath '<span id=\"rowWindowPages\" class=\"row-window-pages\"></span>' 'numbered page navigation host'
	Assert-Contains $hostPath 'SetDashboardElementText("dashboardStatusText", _statusMessage)' 'status updates preserve paging controls'
	Assert-Contains $hostPath 'KKYFB.setRowsFromJson(tab,rowPayload)' 'partial pane virtual payload refresh'
	Assert-Contains $hostPath 'new object[3] { tab, pane.ToString(), rowPayload }' 'partial pane passes compact payload separately'
	Assert-Contains $hostPath 'TryReplaceDashboardPaneInPlace' 'active pane partial replacement'
	Assert-Contains $hostPath 'replaceDashboardPaneHtml' 'ES5 pane replacement API'
	Assert-Contains $hostPath 'requireCurrentRevision: true' 'model mutation requires current standard revision'
	Assert-Contains $hostPath 'FamilyBrowserDataLoader.ValidateCurrentSourceRevision' 'standard revision verification route'
    Assert-Contains $hostPath 'case "lang-en":' 'English language route'
    Assert-Contains $hostPath 'SetLanguage("ko")' 'Korean language handler'
    Assert-Contains $hostPath 'SetLanguage("en")' 'English language handler'
    Assert-Contains $hostPath '_refreshDetachedDetailAfterDocumentCompleted' 'deferred detached detail language refresh'
    Assert-Contains $hostPath 'return FamilyBrowserModernMessageDialog.Show(effectiveOwner, IsKorean(), message, caption, buttons, icon, defaultButton);' 'dashboard message shared modern dialog route'
    Assert-Contains $hostPath 'return FamilyBrowserModernMessageDialog.Show(effectiveOwner, IsKorean(), message, caption, MessageBoxButtons.YesNo, icon, MessageBoxDefaultButton.Button1, positiveButtonText, negativeButtonText);' 'dashboard choice shared modern dialog route'
    Assert-Contains $hostPath 'FamilyBrowserModernMessageDialog.Show(this, _isKorean, Tx("Select at least one family."' 'standard family selection shared modern dialog route'
    Assert-Contains $hostPath 'BuildSelectedSystemTypePreflightReport(registration, standardDoc, doc, selectedCategoryName, selectedSystemTypeName, selectedSystemFamilyKind, selectedSystemTypes' 'selected system type apply uses selected-only preflight path'
    Assert-Contains $hostPath 'MergeSelectedSystemTypePreflightReports' 'multi-selected system preflight merge helper'
    Assert-Contains $hostPath 'ReportSelectedSystemTypePreflightProgress' 'multi-selected system preflight progress mapper'
    Assert-Contains $hostPath 'FamilyBrowserOperationHtmlDialog.ShowFamilyLoadConfirmation' 'family load confirmation uses HTML operation dialog'
    Assert-Contains $hostPath 'FamilyBrowserOperationHtmlDialog.ShowFamilyLoadResult' 'family load result uses HTML operation dialog'
    Assert-Contains $hostPath 'FamilyBrowserOperationHtmlDialog.ShowSystemTypeApplyConfirmation' 'system type apply confirmation uses HTML operation dialog'
    Assert-Contains $hostPath 'FamilyBrowserOperationHtmlDialog.ShowSystemTypeApplyResult' 'system type apply result uses HTML operation dialog'
    Assert-Contains $hostPath 'resultBrowser.DocumentText = FamilyBrowserMessageHtmlRenderer.Build' 'current model check result uses structured HTML body'
    Assert-NotContains $hostPath 'TextBox messageBox = new TextBox' 'plain multiline current-model result body'
    Assert-Contains $hostPath 'using (FamilyBrowserHtmlDialogHost dialog = new FamilyBrowserHtmlDialogHost(IsKorean(), caption, html' 'current model check uses full HTML dialog shell'
    Assert-Contains $hostPath 'dialog.AuxiliaryActionRequested += delegate' 'current model check HTML Excel export action'
	Assert-Contains $hostPath 'ProjectComparisonReviewExcelExportService.CountReviewRows(report) + handledDialogs.Count' 'current model check exposes Excel when only auto-handled scan dialogs exist'
	Assert-Contains $hostPath 'SaveReviewList(dialog.FileName, report, disciplineLabel, _lastComparisonPath, handledDialogs, IsKorean())' 'current model check sends auto-handled dialogs to the workbook export'
    Assert-NotContains $hostPath '= new CurrentModelCheckResultDialog(' 'legacy current model result dialog activation'
    Assert-Contains $hostPath 'RefreshDocumentShellOnlyWithProgress(commandTitle, T("Refreshing browser rows after system type apply...' 'post-system-apply shell refresh has visible progress'
    Assert-Contains $hostPath 'ShowStandardRegistrationResultWithExcelExport' 'standard scan result exposes on-demand Excel export'
    Assert-Contains $hostPath 'ShowPartialRefreshResultWithExcelExport' 'selected family refresh result exposes on-demand Excel export'
    Assert-Contains $hostPath 'ShowThumbnailResultWithExcelExport' '3D image result exposes on-demand Excel export'
    Assert-Contains $hostPath 'T("Export Excel", "Excel 내보내기")' 'result dialogs expose localized Excel export action'
    Assert-NotContains $hostPath 'T("Diagnostic Excel: ", "진단 Excel: ")' 'result dialogs do not point to automatically generated diagnostic workbooks'
    Assert-NotContains $hostPath 'PromptSaveThumbnailDiagnosticReport' 'thumbnail result does not open a follow-up save prompt'
    Assert-NotContains $hostPath 'selectedSystemTypeCount > 1) ? SystemTypePreflightBuilderService.BuildReport' 'legacy multi-selected system type full BuildReport preflight'
    Assert-NotContains $hostPath 'FamilyLoadResultDialog resultDialog' 'legacy family load result dialog call'
    Assert-NotContains $hostPath 'MessageBox.Show(owner, message, caption' 'raw dashboard message box fallback'
    Assert-NotContains $hostPath 'using DashboardMessageDialog dialog' 'legacy dashboard message dialog route'
    Assert-NotContains $hostPath 'MessageBox.Show(this, Tx("Select at least one family."' 'raw standard selection message box'
    Assert-WindowContains $hostPath 'Export current model review list' 'Excel Workbook (*.xlsx)|*.xlsx' 'current model review Excel-only export filter'
    Assert-WindowContains $hostPath 'Export current model review list' 'Excel review workbook exported.' 'current model review Excel export message'
    Assert-WindowNotContains $hostPath 'Export current model review list' 'CSV UTF-8 (*.csv)' 'current model review unsupported CSV filter'
    Assert-WindowNotContains $hostPath 'Export file-specific guards to Excel' 'CSV UTF-8 (*.csv)' 'file guard export unsupported CSV filter'
    Assert-Contains $hostPath '<meta charset' 'dashboard/sub-window UTF-8 meta'
    Assert-Contains $hostPath 'replace(/\\r?\\n/g' 'escaped generated JavaScript newline regex'
    Assert-NotContains $hostPath 'replace(/\r?\n/g' 'unescaped generated JavaScript newline regex'
    Assert-Contains $hostPath 'function captureDashboardUiStateJson' 'dashboard UI state capture'
    Assert-Contains $hostPath 'bool canUseDebug = CanSeeInternalPaths();' 'debug UI and host permission gate alignment'
    Assert-Contains $hostPath 'private bool TryToggleDashboardDebugConsole()' 'debug toggle host verifies panel exists'
    Assert-Contains $hostPath 'if(!panel||typeof toggleDebug!=''function'')return false;' 'debug host rejects missing debug panel'
    Assert-Contains $hostPath 'var panel=byId(''fbDebug'');if(!b||!panel)return true;' 'debug JS detects missing debug panel before host route'
    Assert-Contains $hostPath 'if(toggleDebug()===false)' 'debug JS toggles visible debug panel first'
    Assert-Contains $hostPath 'window.location=''kkyfb:debug-log'';' 'debug JS routes missing-panel F12 to host'
    Assert-Contains $hostPath '#fbDebug{display:none;position:fixed;left:212px;right:0;top:auto;bottom:0;height:320px' 'debug panel inline CSS docks at bottom of browser window'
    Assert-NotContains $hostPath '<a id=\"fbDebugFab\"' 'floating debug FAB button is not rendered'
	Assert-Contains $hostPath 'rail-title rail-support-title' 'always-visible Support sidebar group'
	Assert-Contains $hostPath 'id=\"fbUpdateCheckTool\" class=\"tool\" href=\"kkyfb:update-check\"' 'Family Browser update-check sidebar action'
	Assert-Contains $hostPath 'id=\"fbHomepageTool\" class=\"tool\" href=\"kkyfb:open-homepage\"' 'Family Browser homepage sidebar action'
	Assert-Contains $hostPath 'id=\"fbManualTool\" class=\"tool\" href=\"kkyfb:open-manual\"' 'Family Browser website manual sidebar action'
	Assert-NotContains $hostPath 'ToolTabLink("help", T("Manual", "매뉴얼"))' 'obsolete in-app manual sidebar action'
	Assert-Contains $hostPath 'case "update-check":' 'Family Browser update-check host route'
	Assert-Contains $hostPath 'CheckFamilyBrowserProductUpdate();' 'Family Browser update-check handler call'
	Assert-Contains $hostPath 'case "open-homepage":' 'Family Browser homepage host route'
	Assert-Contains $hostPath 'OpenFamilyBrowserHomepage();' 'Family Browser homepage handler call'
	Assert-Contains $hostPath 'case "open-manual":' 'Family Browser website manual host route'
	Assert-Contains $hostPath 'OpenFamilyBrowserManual();' 'Family Browser website manual handler call'
	Assert-Contains $hostPath 'Family Browser Manual' 'Manual pane workflow content'
    Assert-NotContains $hostPath 'bool canUseDebug = showAdminUi && canViewAdmin;' 'legacy debug gate mismatch'
    Assert-NotContains $hostPath 'if (TryInvokeDashboardScript("toggleDebug"))' 'legacy debug invoke without visible-panel check'
    Assert-Contains $hostPath 'detachedPreviewInlineSource' 'detached detail inline 3D preview fallback'
    Assert-Contains $hostPath 'TryBuildInlinePreviewDataUri(previewPath' 'detached detail preview data-uri generation'
    Assert-Contains $hostPath 'private bool IsActiveStandardListRegistrationRequired()' 'registered RVT / missing standard-list state classifier'
    Assert-Contains $hostPath '등록된 표준 RVT의 표준 목록을 연결해주세요' 'registered RVT / missing standard-list Korean prompt'
    Assert-WindowContains $hostPath 'else if (IsActiveStandardListRegistrationRequired())' 'AppendStandardListRegistrationRequiredRow' 'standard-list missing branch renders list-registration CTA before RVT fallback'
    Assert-Contains $hostPath 'standard-action-layout baseline-actions' 'admin standard baseline action grouped layout'
    Assert-Contains $hostPath 'standard-action-layout visible-list-actions' 'admin standard visible-list action grouped layout'
    Assert-Contains $hostPath 'settings-action-grid admin-standard-action-grid' 'admin standard action grid has scoped layout class'
    Assert-Contains $hostPath 'admin-trade-control' 'admin trade selector and management actions share one control group'
    Assert-Contains $hostPath 'admin-trade-management-actions' 'admin trade management actions are next to the target selector'
    Assert-Contains $hostPath 'selectedDisciplineKey' 'admin selected trade state follows the current browser target'
    Assert-WindowNotContains $hostPath 'standard-action-layout baseline-actions' 'standard-action-row trade-row' 'baseline RVT card no longer contains trade management actions'
    Assert-Contains $hostPath 'settings-action-grid audit-action-grid' 'model-check action grid has scoped layout class'
    Assert-Contains $hostPath 'AppendAuditTargetSelector(sb, policy, reviewTarget);' 'model-check target selector is rendered near selected review target'
    Assert-Contains $hostPath 'private void AppendAuditTargetSelector' 'model-check target selector helper'
    Assert-Contains $hostPath 'audit-target-selector' 'model-check target selector markup'
    Assert-Contains $hostPath 'audit-target-chip' 'model-check target selector chip markup'
    Assert-Contains $hostPath 'Current Model Check target changed:' 'model-check target switch status message'
    Assert-WindowContains $hostPath 'private void SetBrowseDiscipline' 'bool preserveBrowseTargetUiState' 'model-check target switch identifies scroll-preserving workflow'
    Assert-WindowContains $hostPath 'private void SetBrowseDiscipline' 'CaptureDashboardUiState();' 'model-check target switch captures current scroll before refresh'
    Assert-WindowContains $hostPath 'private void SetBrowseDiscipline' 'if (preserveBrowseTargetUiState)' 'model-check target switch keeps UI state while browser tabs retain reset behavior'
    Assert-WindowContains $hostPath 'AppendCheckMaintenanceWorkspace' 'settings-action-grid audit-action-grid' 'model-check workspace action grid is scoped to audit layout' 64
    Assert-WindowContains $hostPath 'Baseline RVT' 'settings-action-grid admin-standard-action-grid' 'admin standard section action grid is scoped to admin layout'
    Assert-WindowContains $hostPath 'Last scan state' 'settings-action-grid standard-rvt-manager-action-grid' 'standard RVT manager does not borrow admin standard action-grid class'
    Assert-Contains $hostPath 'AppendStandardActionLink' 'admin standard grouped action link helper'
    Assert-Contains $hostPath 'private string DisplayProjectSubtitleHtml()' 'top project subtitle renders file tokens with tooltips'
    Assert-Contains $hostPath 'ProjectFileTokenHtml(T("Local",' 'top project subtitle local file token'
    Assert-Contains $hostPath 'ProjectFileTokenHtml(T("Central",' 'top project subtitle central file token'
    Assert-Contains $hostPath 'TryResolveDisplayCentralPath(localPath)' 'top project subtitle central path resolution'
    Assert-Contains $hostPath 'DisplayProjectSubtitleHtml()' 'dashboard top uses HTML subtitle with title attributes'
    Assert-Contains $hostPath 'project-file-stack' 'top project file tokens render as stacked rows'
    Assert-Contains $hostPath '<div class=\"top-actions\">' 'header action buttons are separate from status pills'
    Assert-Contains $hostPath '<div class=\"pills status-pills\">' 'status pills are separate from top action buttons'
    Assert-NotContains $hostPath 'parts.Add("<span class=\"project-title\"' 'duplicate current project title in file context'
    Assert-NotContains $hostPath 'Html(DisplayProjectSubtitle())' 'legacy plain project subtitle without path tooltips'
    if (Test-Path -LiteralPath $assetPath) {
        Assert-Contains $assetPath 'Header file context: keep actions top-right' 'header file context layout CSS block'
        Assert-Contains $assetPath 'body.fb-shell-20260507 .top-actions' 'top action button layout CSS'
        Assert-Contains $assetPath 'body.fb-shell-20260507 .project-context' 'right-aligned local/central context CSS'
        Assert-Contains $assetPath 'body.fb-shell-20260507 .pills.status-pills' 'separate status pill layout CSS'
        Assert-Contains $assetPath 'top: 178px !important;' 'body layout is pushed below enlarged header'
    } else {
        Add-Failure "Missing dashboard shell asset: $assetPath"
    }
    Assert-Contains $hostPath 'kkyAutoOpenDetachedDetailForCurrentTab' 'Family/System tab auto detached detail open helper'
    Assert-Contains $hostPath "kkyQueueAutoDetachedDetailOpen('setTab')" 'Family/System tab switch queues detached detail auto open'
    Assert-Contains $hostPath "kkyQueueAutoDetachedDetailOpen('onload')" 'initial Family/System render queues detached detail auto open'
    Assert-Contains $hostPath 'function kkySearchInputActive(search)' 'Family/System search focus active helper'
    Assert-Contains $hostPath 'kkyRestoreSearchFocus(search,restoreFocus)' 'Family/System queued search restores search focus'
    Assert-Contains $hostPath "filterRows('search')" 'Family/System search input calls quiet filter path'
    Assert-Contains $hostPath 'selectRow(first,false,quietDetail)' 'Family/System search filtering uses quiet detail selection'
    Assert-Contains $hostPath 'window.scheduleDetachedDetailSync=function(){return false;};' 'Family/System quiet detail selection suppresses detached sync'
    Assert-Contains $hostPath 'window.requestInlinePreview=function(path,detail)' 'Family/System quiet detail selection suppresses inline preview host action'
    Assert-Contains $hostPath 'string title = T("Selected Item Detail",' 'detached detail window title avoids duplicate family name'
    Assert-NotContains $hostPath 'string detailName = GetDashboardElementInnerText("detailName");' 'legacy detached detail duplicate title source'
    Assert-Contains $hostPath '.detached-content #detailParameterBlock{display:block;clear:both;width:100%;' 'detached detail parameter block uses full width'
    Assert-Contains $hostPath '.detached-content #detailParameterBlock .parameter-panel{max-height:none;overflow:visible;' 'detached detail parameter panel does not create nested vertical scroll'
    Assert-Contains $hostPath 'parameter-table-scroll detached-parameter-scroll' 'detached detail parameter tables are wrapped in an internal scroll container'
    Assert-Contains $hostPath 'min-width:1280px!important' 'detached detail parameter table has controlled horizontal scroll width'
    Assert-Contains $hostPath 'col.parameter-formula-col{width:620px!important;}' 'detached detail parameter formula column has dedicated width'
    Assert-Contains $hostPath 'parameter-formula-cell' 'detached detail parameter formula cell class'
    Assert-Contains $hostPath 'renderDetachedParameterRowsTable' 'detached detail parameter table renderer override'
    Assert-Contains $hostPath 'DeduplicatePreviewParameters(allParameters);' 'parameter preview includes all captured family and instance parameters'
    Assert-NotContains $hostPath 'DeduplicatePreviewParameters(allParameters.Where' 'parameter preview is not limited to shared parameters'
    Assert-NotContains $hostPath 'familyParameters.Take(6)' 'family parameter preview is not truncated'
    Assert-NotContains $hostPath 'instanceParameters.Take(8)' 'instance parameter preview is not truncated'
    Assert-NotContains $hostPath 'sampleTypeParameters.Take(12)' 'sample type parameter preview is not truncated'
    Assert-Contains $hostPath 'function mergeParameterRows(baseInfo,typeInfo,typeName)' 'main and detached detail merge common and selected-type parameters'
    Assert-Contains $hostPath 'data-unified-parameter-table=\"true\"' 'parameter detail renders one semantic unified table'
    Assert-Contains $hostPath 'id=\"paramUnifiedRows_' 'main detail unified parameter table target'
    Assert-Contains $hostPath 'id=\"detachedUnifiedRows_' 'detached detail unified parameter table target'
    Assert-Contains $hostPath 'if(hasTypes&&unifiedParameterScopeIsType(baseRows[i].scope)&&!unifiedParameterIsCsv(baseRows[i]))continue;' 'sample-type rows are removed before selected type rows are merged'
    Assert-Contains $hostPath "value.indexOf(' rows x ')>=0&&value.indexOf(' columns')>=0" 'lookup CSV parameter summary remains in the unified table'
    Assert-Contains $hostPath 'var scope=csv?paramCsvLabel:' 'lookup CSV rows use an explicit CSV scope'
    Assert-Contains $hostPath "var paramCsvLabel='CSV';" 'main dashboard defines the unified CSV parameter label'
    Assert-Contains $hostPath 'parseNestedChildRows(raw)' 'nested loadable family table parser'
    Assert-Contains $hostPath 'function addNestedChildRow(rows,map,category,family)' 'nested child family dedupe helper'
    Assert-Contains $hostPath "(existing.category||'-')=='-'&&category!='-'" 'nested child dedupe prefers categorized rows over dash rows'
    Assert-Contains $hostPath 'nestedCategoryLabel' 'nested child category column label'
    Assert-Contains $hostPath 'nestedFamilyNameLabel' 'nested child family-name column label'
    Assert-Contains $hostPath 'class=\"nested-child-table\"' 'nested child family table markup'
    Assert-Contains $hostPath 'class=\"family-type-table\"' 'family type list is rendered as a table'
    Assert-Contains $hostPath 'family-type-panel' 'family type list is wrapped in a styled panel'
    Assert-Contains $hostPath 'categoryName + "\t" + familyName' 'nested loadable summary preserves category and family'
    Assert-Contains $hostPath 'class=\"preview-open-chip\" data-src=\"' 'preview large-view chip has clickable modal source'
    Assert-Contains $hostPath '.detached-content #previewBlock .preview{height:330px;min-height:330px;background:#fff;}' 'detached detail stable preview sizing'
    Assert-Contains $hostPath '.detached-content .parameter-unified-head select{min-width:420px;max-width:none;}' 'detached unified parameter type dropdown has wide readable layout'
    Assert-Contains $hostPath '.detached-content .preview-image-cell{left:0!important;right:0!important;top:0!important;bottom:0!important;}' 'detached preview image cell uses IE-compatible bounds'
    Assert-Contains $hostPath 'Math.round((boxW-w)/2)' 'preview image is pixel-centered instead of percent-transform centered'
    Assert-Contains $hostPath "img.style.transform='none';" 'preview image fit does not rely on CSS transform'
    Assert-Contains $hostPath 'data-diffraw=' 'fingerprint diff detail button carries raw payload for detached detail'
    Assert-Contains $hostPath 'IsLoadAvailableComparisonItem(item) ? string.Empty : BuildFingerprintDifferenceTableText(item)' 'load-available family rows suppress fingerprint diff table'
    Assert-Contains $hostPath 'item == null || IsLoadAvailableComparisonItem(item)' 'load-available family detail uses standard snapshot even in admin mode'
    Assert-Contains $hostPath 'ResolveLoadableDetailNestedFamilies(item, standardFamily)' 'load-available family nested detail can resolve from standard snapshot'
    Assert-Contains $hostPath 'return standardFamily.NestedLoadableFamilies ?? new List<StandardNestedLoadableFamilySnapshotItem>();' 'standard snapshot nested children feed load-available detail'
    Assert-Contains $hostPath 'function summarizeFingerprintDiffRows(rows)' 'fingerprint diff concise summary classifier'
    Assert-Contains $hostPath 'function conciseFingerprintDiffText(raw)' 'basis status uses concise fingerprint diff text'
    Assert-Contains $hostPath '<div id=\"diffModalMask\" class=\"preview-modal-mask diff-modal-mask\"' 'detached/main fingerprint diff modal markup'
    Assert-Contains $hostPath "function openDiffModal(raw){return openFingerprintDiffModal(raw||'');}" 'detached fingerprint diff modal route'
    Assert-NotContains $hostPath 'function openDiffModal(){return false;}' 'legacy detached fingerprint diff no-op'
    Assert-NotContains $hostPath 'AppendFilterBar(sb);' 'duplicate search-area status summary filters'
    Assert-Contains $hostPath '.disciplinebar{margin-top:9px;white-space:normal;overflow:visible;display:flex;flex-wrap:wrap;' 'family/system discipline filters wrap instead of clipping'
    Assert-Contains $hostPath 'body.fb-browser .pane-head.action-head{display:-ms-flexbox!important;display:flex!important;-ms-flex-wrap:wrap!important;' 'family/system action row wraps inline status filters'
    Assert-Contains $hostPath 'body.fb-browser .pane-head.action-head .family-kind-toggle,body.fb-browser .pane-head.action-head .inline-status-toggle{display:-ms-inline-flexbox!important;display:inline-flex!important;-ms-flex-wrap:wrap!important;' 'family/system inline status toggle wrap override'
    Assert-Contains $hostPath 'AppendFamilyInlineStatusFilterBar(sb, rows);' 'family action row keeps the remaining inline status filters'
    Assert-Contains $hostPath 'AppendSystemInlineStatusFilterBar(sb, systemRows);' 'system action row keeps the remaining inline status filters'
    Assert-NotContains $hostPath '.disciplinebar{margin-top:9px;white-space:nowrap;overflow:hidden;}' 'legacy one-line clipped discipline filters'
    Assert-NotContains $hostPath '.inline-status-toggle{display:inline-block;vertical-align:middle;margin-left:10px;max-width:520px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}' 'legacy clipped inline action status filters'
    Assert-Contains $hostPath 'about:kkyfb-request:' 'request draft about-scheme navigation guard'
    Assert-Contains $hostPath 'about:kkyfb-select:' 'standard-family selection about-scheme navigation guard'
    Assert-Contains $hostPath 'about:kkyfb:' 'dashboard and standard RVT manager about-scheme navigation guard'
    Assert-NotContains $hostPath 'selectRow(this,true)' 'row-click detached-detail navigation'
    Assert-NotContains $hostPath 'kkyOriginalSelectRow' 'legacy selectRow wrapper'
	Assert-Contains $hostPath 'ApplyPendingOperationOverlay(GetActiveDocument())' 'language refresh reapplies pending operation overlay'
	Assert-Contains $hostPath 'deepProjectScanDeferredForPendingCommit' 'precise project scan defers while document mutations await commit'
	Assert-Contains $hostPath 'FamilyPendingSaveOrSync' 'family pending save/sync state'
	Assert-Contains $hostPath 'SystemPendingSaveOrSync' 'system pending save/sync state'
	Assert-Contains $hostPath 'RefreshAfterDocumentCommit' 'post-save/sync live dashboard rebuild'
	Assert-Contains $hostPath 'kkyPendingBaseRenderBasisMemo' 'pending detail warning card renderer'
	Assert-Contains $hostPath 'kkyPendingDetailTone' 'pending detail warning tone'

	$operationServicePath = Join-Path $root "$folder\StandardRvtChangeCandidateService.cs"
	$operationEntryPath = Join-Path $root "$folder\FamilyBrowserOperationLogEntry.cs"
	$modelessRuntimePath = Join-Path $root "$folder\FamilyBrowserDashboardModelessRuntime.cs"
	$auditScenarioPath = Join-Path $root "$folder\FamilyBrowserDashboardAuditScenario.cs"
	$appPath = Get-ChildItem -LiteralPath (Join-Path $root $folder) -Recurse -Filter 'App.cs' | Select-Object -First 1 -ExpandProperty FullName
	Assert-Contains $operationServicePath 'public static bool HandleDocumentSaved' 'successful Save/Save As commits pending operations'
	Assert-Contains $operationServicePath 'public static bool HandleDocumentSynchronizedWithCentral' 'successful synchronization commits pending operations'
	Assert-Contains $operationServicePath 'GetPendingOperationEntries' 'dashboard pending-operation projection source'
	Assert-Contains $operationServicePath 'HandleDocumentClosing' 'document close registers pending-state cleanup'
	Assert-Contains $operationServicePath 'HandleDocumentClosed' 'document close discards uncommitted pending state'
	Assert-Contains $operationServicePath 'return "runtime:" + RuntimeHelpers.GetHashCode(doc)' 'pending key survives Save As path and managed-root changes'
	Assert-Contains $operationServicePath 'outcome.Contains("consolidated")' 'consolidated system apply is treated as a successful mutation'
	Assert-Contains $operationEntryPath 'public string StandardSourceKey' 'pending entry scopes overlay to its standard source'
	Assert-Contains $modelessRuntimePath 'NotifyDocumentCommitFinalized' 'post-commit modeless dashboard notification'
	Assert-Contains $modelessRuntimePath 'RefreshForActiveDocumentChanged(key)' 'active document switches refresh the already-open dashboard'
	Assert-Contains $auditScenarioPath 'IncludePendingRows' 'pending state audit scenario seam'
	if ($appPath) {
		Assert-Contains $appPath 'DocumentSaved += HandleDocumentSaved' 'DocumentSaved subscription'
		Assert-Contains $appPath 'DocumentSavedAs += HandleDocumentSavedAs' 'DocumentSavedAs subscription'
		Assert-Contains $appPath 'DocumentClosing += HandleDocumentClosing' 'DocumentClosing subscription'
		Assert-Contains $appPath 'DocumentClosed += HandleDocumentClosed' 'DocumentClosed subscription'
		Assert-Contains $appPath 'NotifyDocumentCommitFinalized' 'save/sync event refreshes browser after commit'
	} else {
		Add-Failure "Missing App.cs under $folder"
	}
    Test-DocumentTextHasMeta $hostPath

    $operationDialogPath = Join-Path $root "$folder\FamilyBrowserOperationHtmlDialog.cs"
    if (Test-Path -LiteralPath $operationDialogPath) {
        Assert-Contains $operationDialogPath 'internal static class FamilyBrowserOperationHtmlDialog' 'shared operation HTML dialog class'
        Assert-Contains $operationDialogPath 'FamilyBrowserHtmlDialogHost' 'operation result/confirmation uses full HTML dialog shell'
        Assert-NotContains $operationDialogPath 'private sealed class OperationHtmlDialog : Form' 'legacy operation WinForms shell'
        Assert-NotContains $operationDialogPath 'Button okButton' 'legacy operation WinForms action button'
        Assert-Contains $operationDialogPath 'ShowFamilyLoadConfirmation' 'family load HTML confirmation entry point'
        Assert-Contains $operationDialogPath 'ShowSystemTypeApplyConfirmation' 'system type HTML confirmation entry point'
        Assert-Contains $operationDialogPath 'ShowFamilyLoadResult' 'family load HTML result entry point'
        Assert-Contains $operationDialogPath 'ShowSystemTypeApplyResult' 'system type HTML result entry point'
        Assert-Contains $operationDialogPath 'FamilyBrowserResultExcelExportUi.SaveRows' 'family/system result workbook is created only by the result export action'
        Assert-Contains $operationDialogPath 'Tx(isKorean, "Export Excel", "Excel 내보내기")' 'family/system result dialog exposes localized Excel export'
        Assert-Contains $operationDialogPath 'dialog.AuxiliaryActionRequested += delegate' 'family/system result Excel export is button-driven'
        Assert-Contains $operationDialogPath 'Item result table' 'HTML result table section'
        Assert-Contains $operationDialogPath 'Post-apply review and diagnostic JSON paths are saved' 'system result explains post-apply diagnostic save'
        Assert-Contains $operationDialogPath 'background:#f5f7fb;color:#111827' 'operation dialog KKY blue light palette'
        Assert-Contains $operationDialogPath 'Tx(isKorean, "Create new", "신규 생성")' 'new system type action uses user-facing new-create terminology'
        Assert-Contains $operationDialogPath 'Tx(isKorean, "Load new", "신규 로드")' 'new family action uses user-facing new-load terminology'
        Assert-NotContains $operationDialogPath 'Tx(isKorean, "Create missing", "누락 생성")' 'legacy missing-create work item terminology'
        Assert-NotContains $operationDialogPath 'Tx(isKorean, "Load missing", "누락 로드")' 'legacy missing-load work item terminology'
        Assert-Contains $hostPath 'T("Create new supported types: ", "신규 지원 타입 생성: ")' 'system apply fallback summary uses new-create terminology'
        Assert-Contains $hostPath 'return T("Create new", "신규 생성");' 'dashboard action translation uses new-create terminology'
    } else {
        Add-Failure "Missing operation HTML dialog: $operationDialogPath"
    }

    $commandFolder = Get-ChildItem -LiteralPath (Join-Path $root $folder) -Directory | Where-Object { $_.Name -like 'KKY_FamilyBrowser_RevitHost_*' } | Select-Object -First 1
    if ($commandFolder) {
        foreach ($commandName in @('CmdApplyStandardLoadableFamilies.cs','CmdApplySystemTypes.cs','CmdCompareProjectToStandard.cs','CmdPreflightSystemTypes.cs','CmdRegisterStandardLibrary.cs','CmdRunFamilyBrowserDiagnostics.cs','CmdStampCurrentProjectState.cs')) {
            $commandPath = Join-Path $commandFolder.FullName $commandName
            Assert-Contains $commandPath 'FamilyBrowserResultDialog.' "rich result dialog route for $commandName"
            Assert-NotContains $commandPath 'TaskDialog.Show(' "plain TaskDialog result route for $commandName"
        }
        Assert-Contains (Join-Path $commandFolder.FullName 'CmdApplySystemTypes.cs') 'FamilyBrowserResultDialog.Confirm' 'system type confirmation uses rich HTML choice dialog'
        Assert-NotContains (Join-Path $commandFolder.FullName 'CmdApplySystemTypes.cs') 'new TaskDialog(' 'native system type confirmation dialog'
    } else {
        Add-Failure "Missing ribbon command folder under $folder"
    }

    $scanDialogRecordPath = Join-Path $root "$folder\FamilyThumbnailAutoConfirmedDialogRecord.cs"
    if (Test-Path -LiteralPath $scanDialogRecordPath) {
        Assert-Contains $scanDialogRecordPath 'public string ActionTaken { get; set; }' 'standard scan dialog record stores OK/Cancel action'
        Assert-Contains $scanDialogRecordPath 'public string AvailableButtons { get; set; }' 'standard scan dialog record stores available buttons'
    } else {
        Add-Failure "Missing scan dialog record: $scanDialogRecordPath"
    }

    $scanDialogGuardPath = Join-Path $root "$folder\FamilyThumbnailConstraintDialogGuard.cs"
    if (Test-Path -LiteralPath $scanDialogGuardPath) {
        Assert-Contains $scanDialogGuardPath 'IsDestructiveDeleteChoiceText' 'standard precise scan detects Delete Instance/Delete Type choices'
        Assert-Contains $scanDialogGuardPath 'HasOnlyDeleteOrCancelButtons' 'standard precise scan cancels delete-only native dialogs'
        Assert-Contains $scanDialogGuardPath 'BuildNativeButtonSummary' 'standard precise scan records native button summary'
        Assert-Contains $scanDialogGuardPath 'StandardFamilyScanOkDialog' 'standard precise scan generic OK dialog reason'
		Assert-Contains $scanDialogGuardPath 'OpeningNotCuttingAnything' 'family edit scan handles Opening not cutting anything warning'
		Assert-Contains $scanDialogGuardPath 'activeFamilyEditScope && hasOkButton' 'active family edit scan accepts any real enabled OK/Confirm/Continue button'
		Assert-Contains $scanDialogGuardPath 'activeFamilyEditScope && hasCancelButton && !hasOkButton && HasOnlyDeleteOrCancelButtons(buttons)' 'active family edit scan cancels only delete-only destructive choices'
		Assert-Contains $scanDialogGuardPath 'ResolveFamilyEditDialogActionForAudit' 'family edit dialog button-topology audit seam'
		Assert-Contains $scanDialogGuardPath 'if (IsNativeDeleteButtonLabel(button.Text))' 'destructive button text overrides misleading IDOK control id'
		Assert-Contains $scanDialogGuardPath 'label.StartsWith("idok:", StringComparison.OrdinalIgnoreCase)' 'family edit audit seam can model misleading native control ids'
		Assert-Contains $scanDialogGuardPath 'TryGetCurrentFamilyContext' 'family edit dialog automation is scoped to the active family'
        Assert-Contains $scanDialogGuardPath 'hasOkButton && (IsGeometryConstraintWarningText(text) || LooksLikeStandardFamilyScanDialog(text))' 'native scan dialog OK requires actual OK button'
        Assert-NotContains $scanDialogGuardPath 'clickedButtonText = "IDOK";' 'legacy native scan OK fallback without actual OK button'
    } else {
        Add-Failure "Missing scan dialog guard: $scanDialogGuardPath"
    }

    $registrationPath = Join-Path $root "$folder\StandardLibraryRegistrationService.cs"
    if (Test-Path -LiteralPath $registrationPath) {
        Assert-NotContains $registrationPath 'standard-scan-dialog-diagnostics-' 'standard scan does not create a diagnostic workbook automatically'
        Assert-NotContains $registrationPath 'WriteRegistrationDialogDiagnosticReport' 'standard scan service retains diagnostics in memory instead of writing a result workbook'
        Assert-Contains $registrationPath 'DiagnosticReportPath = string.Empty' 'legacy diagnostic path remains empty for compatibility'
    } else {
        Add-Failure "Missing standard registration service: $registrationPath"
    }

    $thumbnailServicePath = Join-Path $root "$folder\FamilyThumbnailPreviewService.cs"
    Assert-NotContains $thumbnailServicePath 'WriteBatchDiagnosticReport(result)' 'thumbnail scan does not create a diagnostic text file automatically'

    $detailPath = Join-Path $root "$folder\FamilyBrowserDetachedDetailWindow.cs"
    if (Test-Path -LiteralPath $detailPath) {
        Assert-Contains $detailPath 'about:kkyfb:' 'detached detail about:kkyfb navigation guard'
        Assert-Contains $detailPath 'Math.Min(980, workingArea.Width - 140)' 'detached detail wider default width'
        Assert-Contains $detailPath 'MinimumSize = new Size(640, 560)' 'detached detail stable minimum size'
    }

    $auditScenarioPath = Join-Path $root "$folder\FamilyBrowserDashboardAuditScenario.cs"
    if (Test-Path -LiteralPath $auditScenarioPath) {
        Assert-Contains $auditScenarioPath 'EnsureAuditPreviewImage(_workspaceRoot)' 'audit scenario 3D preview fixture use'
        Assert-Contains $auditScenarioPath 'audit-family-preview.png' 'audit scenario 3D preview fixture file'
        Assert-Contains $auditScenarioPath 'PreviewImagePath = previewImagePath' 'audit scenario preview path row data'
        Assert-Contains $auditScenarioPath 'TypeParameterSummary = "@type' 'audit scenario type parameter detail data'
        Assert-Contains $auditScenarioPath 'FlowOffset * DiversityFactor' 'audit scenario long parameter formula fixture'
        Assert-Contains $auditScenarioPath 'DifferenceSummaryTable = string.Empty' 'audit scenario load-available row has no diff detail data'
        Assert-Contains $auditScenarioPath 'DifferenceSummaryTable = T("Type Count\t2 types\t1 type' 'audit scenario fingerprint diff detail data'
        Assert-Contains $auditScenarioPath 'NestedSummary = "AUDIT_SUPPLY_DIFFUSER_FACE\nAir Terminals\tAUDIT_SUPPLY_DIFFUSER_FACE' 'audit scenario duplicate dash/category nested child row'
        Assert-Contains $auditScenarioPath 'AUDIT_FLOW_BOX\nMechanical Equipment\tAUDIT_FLOW_BOX' 'audit scenario duplicate dash/category nested child second row'
        Assert-Contains $auditScenarioPath 'UI_Audit_Local.rvt' 'audit scenario local project filename'
        Assert-Contains $auditScenarioPath 'UI_Audit_Central.rvt' 'audit scenario central project filename'
        Assert-Contains $auditScenarioPath 'AUDIT_REVIEW_FAMILY' 'audit scenario dense family inline status filters'
        Assert-Contains $auditScenarioPath 'AUDIT_POWER_SYSTEM' 'audit scenario dense system inline status filters'
    } else {
        Add-Failure "Missing dashboard audit scenario: $auditScenarioPath"
    }

    $assetPath = Join-Path $root "$folder\KKY.FamilyBrowser.DashboardAssets.family-browser-shell.js"
    if (Test-Path -LiteralPath $assetPath) {
        Assert-Contains $assetPath 'window.setTab' 'dashboard local tab switch'
        Assert-NotContains $assetPath 'kkyOriginalSelectRow' 'legacy asset selectRow wrapper'
        Assert-Contains $assetPath 'var headBaseH = tab === ''families'' ? 58 : 56;' 'browser action row base height before dynamic measurement'
        Assert-Contains $assetPath 'styleToggleLinks(statusToggle);' 'browser inline status links get no-ellipsis inline style'
        Assert-Contains $assetPath 'head.scrollHeight' 'browser action row height measures wrapped status filters'
        Assert-NotContains $assetPath 'max-width:'' + (tab === ''families'' ? ''520'' : ''620'') + ''px !important;white-space:nowrap !important;overflow:hidden !important;text-overflow:ellipsis' 'legacy inline status JS clipping'
    }

    $cssAssetPath = Join-Path $root "$folder\KKY.FamilyBrowser.DashboardAssets.family-browser-shell.css"
    if (Test-Path -LiteralPath $cssAssetPath) {
        Assert-Contains $cssAssetPath 'standard-action-layout' 'admin standard action grouped layout CSS'
        Assert-Contains $cssAssetPath 'standard-action-row.trade-row' 'admin standard trade action row CSS'
        Assert-Contains $cssAssetPath 'a.standard-action-link' 'admin standard action link CSS'
        Assert-Contains $cssAssetPath 'Admin standard/audit action cards: IE-safe two-column wrap.' 'admin/audit action card IE-safe final CSS block'
        Assert-Contains $cssAssetPath 'admin-standard-action-grid .settings-action-group' 'admin standard action cards use two-column scoped layout'
        Assert-Contains $cssAssetPath 'Admin target controls stay together, and both standard action cards use one button matrix.' 'admin trade control and standard action matrix final CSS block'
        Assert-Contains $cssAssetPath 'admin-trade-management-actions a.standard-action-link' 'admin trade management compact equal-width button CSS'
        Assert-Contains $cssAssetPath 'admin-standard-action-grid .standard-action-row a.standard-action-link' 'admin standard card equal button matrix CSS'
        Assert-Contains $cssAssetPath 'audit-action-grid .settings-action-group' 'model-check action cards use two-column scoped layout'
        Assert-Contains $cssAssetPath 'display: inline-block !important;' 'admin/audit action card layout does not depend on CSS grid'
        Assert-Contains $cssAssetPath 'audit-target-selector' 'model-check target selector CSS'
        Assert-Contains $cssAssetPath 'audit-target-chip.active' 'model-check selected target chip CSS'
        Assert-Contains $cssAssetPath 'body.fb-shell-20260507 .nested-child-table' 'nested child table shell CSS'
        Assert-Contains $cssAssetPath 'body.fb-shell-20260507 .family-type-table' 'family type table shell CSS'
        Assert-Contains $cssAssetPath 'transform: none !important;' 'preview shell CSS does not force transform centering'
        Assert-Contains $cssAssetPath 'Filter buttons must wrap after a scan' 'status/trade filter no-ellipsis CSS guard'
        Assert-Contains $cssAssetPath 'flex-wrap: wrap !important;' 'status/trade filters wrap in shell CSS'
        Assert-Contains $cssAssetPath 'text-overflow: clip !important;' 'status/trade filter labels are not ellipsized'
        Assert-Contains $cssAssetPath 'Inline action status filters sit beside load/apply buttons and must wrap instead of ellipsizing.' 'inline action status filter no-ellipsis CSS guard'
        Assert-Contains $cssAssetPath 'body.fb-shell-20260507.fb-browser .pane-head.action-head .inline-status-toggle' 'inline action status filter shell selector'
        Assert-Contains $cssAssetPath 'Debug log is opened from the left menu only and docks like a bottom console.' 'debug log bottom dock CSS guard'
        Assert-Contains $cssAssetPath 'body.fb-shell-20260507 #fbDebugFab' 'debug FAB CSS is explicitly hidden if legacy markup appears'
        Assert-Contains $cssAssetPath 'left: 212px !important;' 'debug dock aligns after left navigation'
        Assert-Contains $cssAssetPath 'bottom: 0 !important;' 'debug dock is attached to bottom edge'
    }

    $modernDialogPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserModernMessageDialog.cs'
    if (Test-Path -LiteralPath $modernDialogPath) {
        Assert-Contains $modernDialogPath 'internal static class FamilyBrowserModernMessageDialog' 'shared modern message dialog class'
        Assert-Contains $modernDialogPath 'FamilyBrowserHtmlDialogHost' 'shared modern message dialog full HTML shell host'
        Assert-NotContains $modernDialogPath 'private sealed class ModernMessageDialog : Form' 'legacy modern WinForms shell'
        Assert-NotContains $modernDialogPath 'Button okButton' 'legacy modern WinForms OK button'
        Assert-NotContains $modernDialogPath 'Label title = new Label' 'legacy modern WinForms title label'
        Assert-Contains $modernDialogPath 'BuildHtmlForAudit' 'shared modern message dialog audit render seam'
        Assert-NotContains $modernDialogPath 'TextBox messageBox = new TextBox' 'legacy plain-text modern message body'
    } else {
        Add-Failure "Missing shared modern message dialog: $modernDialogPath"
    }

    $htmlDialogHostPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserHtmlDialogHost.cs'
    if (Test-Path -LiteralPath $htmlDialogHostPath) {
        Assert-Contains $htmlDialogHostPath 'internal sealed class FamilyBrowserHtmlDialogHost : Form' 'single WinForms host for full HTML dialog'
		Assert-Contains $htmlDialogHostPath 'FamilyBrowserOverflowTitleScript.Script()' 'common HTML result and confirmation dialogs use overflow titles'
        Assert-Contains $htmlDialogHostPath 'FormBorderStyle = FormBorderStyle.None' 'native dialog chrome is hidden'
        Assert-Contains $htmlDialogHostPath 'private readonly WebBrowser _browser;' 'full HTML dialog has one visible WebBrowser surface'
        Assert-Contains $htmlDialogHostPath 'data-dialog-shell' 'full HTML dialog semantic shell marker'
        Assert-Contains $htmlDialogHostPath 'id=\"dialogTitle\"' 'full HTML dialog title element'
        Assert-Contains $htmlDialogHostPath 'id=\"dialogAccept\"' 'full HTML dialog accept action'
        Assert-Contains $htmlDialogHostPath 'id=\"dialogClose\"' 'full HTML dialog close action'
        Assert-Contains $htmlDialogHostPath 'case "copy-details"' 'full HTML dialog copy-details route'
        Assert-Contains $htmlDialogHostPath 'case "open-folder"' 'full HTML dialog report-folder route'
        Assert-Contains $htmlDialogHostPath 'id=\"dialogAuxiliary\"' 'full HTML dialog auxiliary action element'
        Assert-Contains $htmlDialogHostPath 'case "auxiliary"' 'full HTML dialog auxiliary action route'
        Assert-Contains $htmlDialogHostPath 'AuxiliaryActionRequested?.Invoke' 'full HTML dialog auxiliary action event'
        Assert-Contains $htmlDialogHostPath 'buttons == MessageBoxButtons.YesNo ? DialogResult.No : DialogResult.OK' 'full HTML choice close means cancel/no'
        Assert-NotContains $htmlDialogHostPath 'new Button' 'full HTML dialog has no visible WinForms buttons'
        Assert-NotContains $htmlDialogHostPath 'new Label' 'full HTML dialog has no visible WinForms labels'
    } else {
        Add-Failure "Missing full HTML dialog host: $htmlDialogHostPath"
    }

    $resultExcelUiPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserResultExcelExportUi.cs'
    if (Test-Path -LiteralPath $resultExcelUiPath) {
        Assert-Contains $resultExcelUiPath 'using (SaveFileDialog dialog = new SaveFileDialog())' 'result Excel export always asks the user for a destination'
        Assert-Contains $resultExcelUiPath 'if (dialog.ShowDialog(owner) != DialogResult.OK)' 'cancelled result export creates no workbook'
        Assert-Contains $resultExcelUiPath 'FamilyBrowserStandardListExcelExportService.SaveRows(dialog.FileName' 'result workbook write occurs only after save confirmation'
        Assert-Contains $resultExcelUiPath 'Excel Workbook (*.xlsx)|*.xlsx' 'result export is xlsx-only'
    } else {
        Add-Failure "Missing shared on-demand result Excel export UI: $resultExcelUiPath"
    }

    $messageRendererPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserMessageHtmlRenderer.cs'
    if (Test-Path -LiteralPath $messageRendererPath) {
        Assert-Contains $messageRendererPath 'data-message-structured' 'structured message semantic state'
        Assert-Contains $messageRendererPath 'messageSectionCause' 'structured message failure-reason card id'
        Assert-Contains $messageRendererPath 'messageSectionAction' 'structured message next-action card id'
        Assert-Contains $messageRendererPath 'messageSupportCode' 'structured message support-code row'
        Assert-Contains $messageRendererPath 'technical-scroll' 'structured message technical detail internal scroll'
        Assert-Contains $messageRendererPath 'overflow-x:hidden' 'structured message page horizontal overflow guard'
        Assert-Contains $messageRendererPath 'Content-Type' 'structured message explicit UTF-8 content type'
        Assert-Contains $messageRendererPath 'data-message-auto-result' 'automatic result semantic state'
        Assert-Contains $messageRendererPath 'id=\"messageMetricGrid\"' 'automatic result metric card grid'
        Assert-Contains $messageRendererPath 'id=\"messageContextTable\"' 'automatic result context table'
        Assert-Contains $messageRendererPath 'id=\"messageOutputList\"' 'automatic result output path table'
        Assert-Contains $messageRendererPath 'AnalyzeAutomaticResult' 'automatic result line classifier'
        Assert-Contains $messageRendererPath 'FindPrimaryOutputPath' 'automatic result report action path resolver'
    } else {
        Add-Failure "Missing shared message HTML renderer: $messageRendererPath"
    }

    $resultDialogPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserResultDialog.cs'
    if (Test-Path -LiteralPath $resultDialogPath) {
        Assert-Contains $resultDialogPath 'internal static class FamilyBrowserResultDialog' 'shared rich result dialog entry point'
        Assert-Contains $resultDialogPath 'FamilyBrowserModernMessageDialog.Show' 'shared rich result routes to HTML shell'
        Assert-Contains $resultDialogPath 'public static bool Confirm' 'shared rich choice dialog entry point'
    } else {
        Add-Failure "Missing shared rich result dialog: $resultDialogPath"
    }

    $fileGuardPath = Join-Path $root "$folder\FileGuardHtmlConfigurationForm.cs"
    if (Test-Path -LiteralPath $fileGuardPath) {
        Assert-Contains $fileGuardPath 'FamilyBrowserModernMessageDialog.Show(this, _isKorean, Tx("Remove every RVT from the file-specific guard list?"' 'file guard clear-all shared modern dialog route'
        Assert-Contains $fileGuardPath 'about:kkyfileguard://' 'file guard about-scheme navigation guard'
		Assert-Contains $fileGuardPath 'BlockNestedOnlyStandalonePlacement = row.BlockNestedOnly' 'file guard persists nested-only standalone placement option'
		Assert-Contains $fileGuardPath 'Block nested-only standalone placement' 'file guard exposes nested-only standalone placement column'
		Assert-Contains $fileGuardPath '하위 전용 패밀리 단독 모델링 금지' 'file guard exposes the requested Korean nested-only policy label'
		Assert-Contains $fileGuardPath 'new precise standard scan' 'file guard explains precise-scan prerequisite'
		Assert-Contains $fileGuardPath 'TrackElementChanges = row.TrackElements' 'file guard persists per-file element tracking scope'
		Assert-Contains $fileGuardPath 'TrackElementChangesConfigured = true' 'file guard distinguishes explicit tracking choices from legacy targets'
		Assert-Contains $fileGuardPath 'Track element changes' 'file guard exposes the element tracking column'
		Assert-Contains $fileGuardPath '요소 생성·수정·삭제 추적' 'file guard exposes the requested Korean element tracking label'
        Assert-NotContains $fileGuardPath 'MessageBox.Show(this, Tx("Remove every RVT from the file-specific guard list?"' 'raw file guard clear-all message box'
    } else {
        Add-Failure "Missing file guard HTML configuration form: $fileGuardPath"
    }

    $sheetSelectionPath = Join-Path $root "$folder\StandardListSheetSelectionHtmlForm.cs"
    if (Test-Path -LiteralPath $sheetSelectionPath) {
        Assert-Contains $sheetSelectionPath 'about:kkysheet://' 'standard list sheet-selection about-scheme navigation guard'
    } else {
        Add-Failure "Missing standard list sheet selection form: $sheetSelectionPath"
    }

    $reviewExportPath = Join-Path $root "$folder\ProjectComparisonReviewExcelExportService.cs"
    if (Test-Path -LiteralPath $reviewExportPath) {
        Assert-Contains $reviewExportPath 'EnsureWorkbookOutputPath(outputPath)' 'review-list workbook path normalization'
		Assert-Contains $reviewExportPath 'WriteWorkbook(workbookPath, sheetName, headers, rows, dialogSheetName, BuildAutoHandledDialogHeaders(korean), dialogRows)' 'review-list Excel workbook writer includes optional scan-dialog sheet'
		Assert-Contains $reviewExportPath '"ScanDialogs"' 'review-list workbook scan-dialog sheet name'
		Assert-Contains $reviewExportPath 'record.DialogText ?? string.Empty' 'review-list workbook exports the full Revit error or warning text'
        Assert-NotContains $reviewExportPath 'WriteCsv(' 'review-list unsupported CSV writer'
        Assert-NotContains $reviewExportPath 'EscapeCsv(' 'review-list unsupported CSV escaping'
        Assert-Contains $reviewExportPath 'BuildSystemDifferenceFields(item2, korean)' 'system type difference export fields'
        Assert-Contains $reviewExportPath '"차이 항목", "표준", "프로젝트", "차이 요약"' 'review export difference headers'
    } else {
        Add-Failure "Missing review export service: $reviewExportPath"
    }

    $signaturePath = Join-Path $root "$folder\LoadableFamilyContentSignatureService.cs"
    if (Test-Path -LiteralPath $signaturePath) {
        Assert-Contains $signaturePath 'BuildLookupTableSignature(familyDocument, lookupDisplayLines)' 'loadable-family normalized lookup signature and display metadata capture'
        Assert-Contains $signaturePath 'lookup-display-table=' 'lookup CSV original-case display metadata'
        Assert-Contains $signaturePath 'NormalizeLookupDisplayToken' 'lookup CSV display token preserves source casing'
        Assert-Contains $signaturePath 'BuildDebugMetadata(elementDebugLines.Concat(lookupDisplayLines))' 'lookup CSV display metadata is persisted outside the fingerprint source'
        Assert-Contains $signaturePath 'typeof(FamilySizeTableManager)' 'Revit family size table manager lookup'
        Assert-Contains $signaturePath 'ResolveFamilySizeTableOwnerFamilyId(familyDocument)' 'family size table owner family id resolution'
        Assert-Contains $signaturePath 'new object[2] { familyDocument, ownerFamilyId }' 'Revit family size table two-argument API call'
        Assert-Contains $signaturePath 'GetAllSizeTableNames' 'family size table name capture'
        Assert-Contains $signaturePath 'GetSizeTable' 'family size table body capture'
        Assert-Contains $signaturePath 'AsValueString' 'family size table cell value capture'
        Assert-Contains $signaturePath '|columns=' 'family size table column signature capture'
        Assert-Contains $signaturePath '|rows=' 'family size table row/cell signature capture'
    } else {
        Add-Failure "Missing signature service: $signaturePath"
    }

    $comparisonPath = Join-Path $root "$folder\ProjectStandardComparisonService.cs"
    if (Test-Path -LiteralPath $comparisonPath) {
        Assert-Contains $comparisonPath '"lookup tables"' 'lookup table difference classification'
        Assert-Contains $comparisonPath 'BuildLookupCsvDifferenceDetails' 'lookup CSV difference details builder'
        Assert-Contains $comparisonPath '"lookup csv"' 'lookup CSV concise difference area'
        Assert-Contains $comparisonPath 'Lookup CSV row/column count differs.' 'lookup CSV row-column difference message'
    } else {
        Add-Failure "Missing comparison service: $comparisonPath"
    }

    $hostPath = Join-Path $root "$folder\FamilyBrowserDashboardHtmlForm.cs"
    if (Test-Path -LiteralPath $hostPath) {
        Assert-Contains $hostPath 'BuildLookupCsvPreviewText' 'detail view lookup CSV preview builder'
        Assert-Contains $hostPath 'displayResult.Count > 0' 'detail view prefers original-case lookup CSV display metadata'
        Assert-Contains $hostPath 'lookup-display-table=' 'detail view parses original-case lookup CSV display metadata'
        Assert-NotContains $hostPath 'string line = Normalize(rawLine);' 'lookup CSV preview parser must not lowercase display names'
        Assert-Contains $hostPath 'int.TryParse(value.Trim(), NumberStyles.Integer' 'lookup CSV display metadata reads explicit row and column counts'
        Assert-Contains $hostPath 'paramLookupCsvScopeLabel' 'detached detail parameter parser supports lookup CSV section'
        Assert-Contains $hostPath 'BuildLookupCsvDifferenceLine' 'fingerprint diff table renders lookup CSV rows'
        Assert-Contains $hostPath '"CSV 테이블"' 'Korean lookup CSV label'
        Assert-Contains $hostPath 'renderSystemDetailTable' 'system type detail table renderer'
        Assert-Contains $hostPath 'data-system-routing-preferences' 'system type unified routing-preference surface'
        Assert-Contains $hostPath 'system-routing-preference-table' 'system type Revit-style routing table'
        Assert-Contains $hostPath 'parseSystemRoutingModel' 'system type structured routing model parser'
		Assert-MinimumOccurrences $hostPath 'FamilyBrowserSystemRoutingUnitUi.Script(!IsKorean(), InitialMeasurementDisplayUnit())' 2 'main and detached persisted measurement-unit renderer injection'
		Assert-Contains $hostPath 'actionKey.StartsWith("measurement-unit/", StringComparison.OrdinalIgnoreCase)' 'measurement unit UI-only action routing'
		Assert-Contains $hostPath 'SetMeasurementDisplayUnitPreference(action.Substring("measurement-unit/".Length))' 'measurement unit host action handler'
		Assert-Contains $hostPath 'FamilyBrowserMeasurementUnitPreferenceService.Load()' 'persisted measurement unit startup load'
		Assert-Contains $hostPath 'TryInvokeDashboardScript("setSystemDisplayUnitFromHost", unit)' 'measurement unit main dashboard synchronization'
		Assert-Contains $hostPath '_detachedDetailWindow.SyncMeasurementDisplayUnit(unit)' 'measurement unit detached detail synchronization'
		Assert-Contains $hostPath 'bool isSystemDetail = (parameterRaw ?? string.Empty).IndexOf("@system-detail-v1", StringComparison.OrdinalIgnoreCase) >= 0;' 'detached detail identifies system content explicitly'
		Assert-Contains $hostPath 'body.fb-system-detail .detached-content #detailNestedBlock,body.fb-system-detail .detached-content #previewBlock{display:none!important;}' 'detached system detail suppresses family composition and legacy preview blocks'
		Assert-Contains $hostPath '(isSystemDetail ? " fb-system-detail" : string.Empty)' 'detached system detail body class routing'
		Assert-MinimumOccurrences $hostPath 'T("Elbows", "엘보 (Elbows)")' 2 'main and detached Korean Elbows bilingual label'
		Assert-MinimumOccurrences $hostPath 'T("Junctions", "접합 (Junctions)")' 2 'main and detached Korean Junctions bilingual label'
		Assert-MinimumOccurrences $hostPath 'T("Transitions", "전이 (Transitions)")' 2 'main and detached Korean Transitions bilingual label'
		Assert-MinimumOccurrences $hostPath 'T("Unions", "유니온 (Unions)")' 2 'main and detached Korean Unions bilingual label'
		Assert-MinimumOccurrences $hostPath 'T("Caps", "캡 (Caps)")' 2 'main and detached Korean Caps bilingual label'
        Assert-Contains $hostPath "previewBlock.style.display=(currentTab=='systems')?'none':'block'" 'system type legacy bottom preview hidden'
        Assert-Contains $hostPath 'id=\"detailParameterTitle\"' 'system type detail section title target'
        Assert-Contains $hostPath '@system-detail-v1' 'system type detail raw marker support'
        Assert-Contains $hostPath "currentTab=='systems'" 'system tab detail rendering branch'
        Assert-Contains $hostPath 'BuildSystemPreflightDetailSummaryMap' 'system preflight detail summary resolver'
        Assert-Contains $hostPath 'ParameterSummary = ResolveSystemPreflightDetailSummary' 'system execution rows carry detail summary'
    } else {
        Add-Failure "Missing dashboard host: $hostPath"
    }

	$detachedDetailPath = Join-Path $root "$folder\FamilyBrowserDetachedDetailWindow.cs"
	if (Test-Path -LiteralPath $detachedDetailPath) {
		Assert-Contains $detachedDetailPath 'Action<string> _measurementUnitChanged' 'detached detail measurement-unit callback'
		Assert-Contains $detachedDetailPath 'SyncMeasurementDisplayUnit' 'detached detail measurement-unit synchronization API'
		Assert-Contains $detachedDetailPath 'action.StartsWith("measurement-unit/"' 'detached detail measurement-unit route handler'
		Assert-Contains $detachedDetailPath 'FamilyBrowserMeasurementUnitPreferenceService.Save(unit)' 'detached detail measurement-unit persistence'
	} else {
		Add-Failure "Missing detached detail window: $detachedDetailPath"
	}

    $auditScenarioPath = Join-Path $root "$folder\FamilyBrowserDashboardAuditScenario.cs"
    if (Test-Path -LiteralPath $auditScenarioPath) {
		Assert-Contains $auditScenarioPath 'AUDIT_COMPOSITE_PARENT' 'audit scenario nested-difference parent fixture'
		Assert-Contains $auditScenarioPath 'AUDIT_NESTED_FLOW_BOX' 'audit scenario differing nested-child fixture'
		Assert-Contains $auditScenarioPath 'AUDIT_MATCHING_NESTED_CHILD' 'audit scenario matching nested-child hidden fixture'
		Assert-Contains $auditScenarioPath 'IsNestedLoadableChild = true' 'audit scenario nested-child marker'
        Assert-Contains $auditScenarioPath 'Audit_SizeTable' 'audit scenario lookup CSV mixed-case table fixture'
		Assert-Contains $auditScenarioPath '@row\t600x600\tFamily\tType\tIsEnabled\tYes\t-' 'audit scenario Yes/No parameter Yes fixture'
		Assert-Contains $auditScenarioPath '@row\t1200x300\tFamily\tType\tIsEnabled\tNo\t-' 'audit scenario Yes/No parameter No fixture'
        Assert-Contains $auditScenarioPath '12 rows x 5 columns' 'audit scenario lookup CSV detail size fixture'
        Assert-Contains $auditScenarioPath 'Lookup CSV row/column count differs.' 'audit scenario lookup CSV diff fixture'
        Assert-Contains $auditScenarioPath 'AUDIT_DUCT_SEGMENT' 'audit scenario system detail segment fixture'
        Assert-Contains $auditScenarioPath 'AUDIT_DUCT_ELBOW' 'audit scenario system detail dependency fixture'
		Assert-Contains $auditScenarioPath '@route\tSegments\t0\tAUDIT_DUCT_SEGMENT' 'audit scenario structured routing fixture'
		Assert-Contains $auditScenarioPath '@layer\t1\tFinish 1 [4]\tAUDIT_BRICK' 'audit scenario structured exterior layer fixture'
		Assert-Contains $auditScenarioPath '@layer\t2\tStructure [1]\tAUDIT_CONCRETE' 'audit scenario structured core layer fixture'
		Assert-Contains $auditScenarioPath '@layer\t3\tFinish 2 [5]\tAUDIT_GYPSUM' 'audit scenario structured variable layer fixture'
		Assert-Contains $auditScenarioPath 'MeasurementUnitCode' 'audit scenario initial measurement unit control'
		Assert-Contains $auditScenarioPath 'min=0.3280839895013123 max=0.984251968503937' 'audit scenario internal-feet range fixture'
		Assert-Contains $auditScenarioPath 'size 1min=0.3280839895013123' 'audit scenario preserves a legacy missing-space minimum fixture'
		Assert-Contains $auditScenarioPath 'min=-1E+30 max=1E+30' 'audit scenario unbounded range sentinel fixture'
        Assert-Contains $auditScenarioPath '@system-detail-v1' 'audit scenario system detail raw fixture'
		Assert-Contains $auditScenarioPath 'SyntheticFamilyCount' '1,000-family performance fixture control'
		Assert-Contains $auditScenarioPath 'SyntheticSystemCount' '1,000-system performance fixture control'
		Assert-Contains $auditScenarioPath 'ExpandAuditFamilyRows' 'synthetic family row builder'
		Assert-Contains $auditScenarioPath 'ExpandAuditSystemRows' 'synthetic system row builder'
		Assert-Contains $auditScenarioPath '_activeStandardScanNeeded = false;' 'missing-RVT audit state stays distinct from scan-needed state'
		Assert-NotContains $auditScenarioPath '_activeStandardScanNeeded = !scenario.StandardRvtRegistered;' 'legacy audit fixture conflating missing RVT with scan-needed state'
    } else {
        Add-Failure "Missing audit scenario: $auditScenarioPath"
    }

	$familyCapturePath = Join-Path $root "$folder\FamilyDocumentParameterCaptureService.cs"
	$standardCapturePath = Join-Path $root "$folder\StandardLibraryRegistrationService.cs"
	$projectCapturePath = Join-Path $root "$folder\ProjectSnapshotCaptureService.cs"
	$systemCapturePath = Join-Path $root "$folder\SystemTypeSemanticCaptureService.cs"
	$systemLayerModelPath = Join-Path $root "$folder\StandardSystemTypeLayerSnapshotItem.cs"
	$comparisonClonePath = Join-Path $root "$folder\ProjectStandardComparisonService.cs"
	Assert-Contains $familyCapturePath 'FamilyBrowserYesNoParameterFormatter.FormatInteger(familyParameter' 'family-document Yes/No value formatting'
	Assert-Contains $standardCapturePath 'FamilyBrowserYesNoParameterFormatter.FormatInteger(familyParameter' 'standard registration family Yes/No value formatting'
	Assert-Contains $standardCapturePath 'FamilyBrowserYesNoParameterFormatter.FormatInteger(parameter, integerValue' 'standard RVT Yes/No value formatting'
	Assert-Contains $standardCapturePath 'GetFirstCoreLayerIndex' 'compound layer first core boundary capture'
	Assert-Contains $standardCapturePath 'GetLastCoreLayerIndex' 'compound layer last core boundary capture'
	Assert-Contains $standardCapturePath 'StructuralMaterialIndex' 'compound layer structural material capture'
	Assert-Contains $standardCapturePath 'VariableLayerIndex' 'compound layer variable index capture'
	Assert-Contains $standardCapturePath 'ResolveCompoundLayerIndex' 'cross-version compound layer metadata resolver'
	Assert-Contains $projectCapturePath 'FamilyBrowserYesNoParameterFormatter.FormatInteger(parameter, integerValue' 'project snapshot Yes/No value formatting'
	Assert-Contains $systemCapturePath 'FamilyBrowserYesNoParameterFormatter.FormatInteger(parameter, parameter.AsInteger())' 'system semantic Yes/No value formatting'
	Assert-Contains $systemLayerModelPath 'public bool IsCore' 'compound layer core metadata model'
	Assert-Contains $systemLayerModelPath 'public bool IsStructuralMaterial' 'compound layer structural metadata model'
	Assert-Contains $systemLayerModelPath 'public bool IsVariable' 'compound layer variable metadata model'
	Assert-Contains $comparisonClonePath 'IsCore = layer.IsCore' 'compound layer core metadata clone'
	Assert-Contains $comparisonClonePath 'IsStructuralMaterial = layer.IsStructuralMaterial' 'compound layer structural metadata clone'
	Assert-Contains $comparisonClonePath 'IsVariable = layer.IsVariable' 'compound layer variable metadata clone'

    $systemDetailPath = Join-Path $root "$folder\SystemTypeDetailSummaryService.cs"
    if (Test-Path -LiteralPath $systemDetailPath) {
        Assert-Contains $systemDetailPath 'RoutingPreferenceManager' 'system type detail routing preference capture'
        Assert-Contains $systemDetailPath 'GetSizes' 'system type detail size count capture'
        Assert-Contains $systemDetailPath 'GetMEPSizes' 'system type detail MEP size count capture'
        Assert-Contains $systemDetailPath '@system-detail-v1' 'system type detail raw marker writer'
        Assert-Contains $systemDetailPath 'AddRoutingPreferenceRow' 'system type structured routing row writer'
		Assert-Contains $systemDetailPath 'routingPreferenceRows' 'system type structured routing row capture'
		Assert-Contains $systemDetailPath 'criteria.Add(JoinClean(" ",' 'routing criterion writer separates size index, minimum, and maximum tokens'
		Assert-NotContains $systemDetailPath 'ToString(CultureInfo.InvariantCulture) + JoinClean(" ", string.IsNullOrWhiteSpace(min)' 'routing criterion writer must not emit size 1min without a delimiter'
        Assert-Contains $systemDetailPath 'routeIndexByRole' 'system type dependency routing priority counter'
        Assert-NotContains $systemDetailPath 'AddRow(lines, "identity", "Segment"' 'misleading system identity Segment row'
        Assert-NotContains $systemDetailPath 'AddRow(lines, "identity", "Material"' 'misleading system identity Material row'
        Assert-Contains $systemDetailPath 'AddDependencyRows' 'system type detail dependent loadable family rows'
		Assert-Contains $systemDetailPath 'AddLayerRows' 'system type detail layer rows'
		Assert-Contains $systemDetailPath '"@layer"' 'system type structured layer row marker'
		Assert-Contains $systemDetailPath 'layer.ThicknessFeet.ToString("G17"' 'system type structured raw-feet layer thickness'
		Assert-Contains $systemDetailPath 'layer.IsCore ? "true" : "false"' 'system type core-layer metadata writer'
    } else {
        Add-Failure "Missing system type detail summary service: $systemDetailPath"
    }
}

$nestedPropagationPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\NestedLoadableFamilyDifferencePropagationService.cs'
if (Test-Path -LiteralPath $nestedPropagationPath) {
	Assert-Contains $nestedPropagationPath 'BuildParentMap' 'nested-family parent relation map'
	Assert-Contains $nestedPropagationPath 'ProjectNestedLoadableFamilies' 'current-project nested relation source'
	Assert-Contains $nestedPropagationPath 'AppendNestedDifferenceToParent' 'nested difference parent propagation'
	Assert-Contains $nestedPropagationPath 'ParentExistsInCurrentProject' 'wholly absent parent nested-child false-positive guard'
	Assert-Contains $nestedPropagationPath 'parent.Status = string.Equals(Normalize(child.Status), "manualreview", StringComparison.Ordinal) ? "ManualReview" : "DifferentFromStandard";' 'matching parent becomes different when a child differs'
	Assert-Contains $nestedPropagationPath 'UpsertNestedDependencyDetail' 'nested child exact difference rows propagate recursively'
	Assert-Contains $nestedPropagationPath 'nested helper rows are not loaded independently' 'nested child direct-load guidance'
	Assert-Contains $nestedPropagationPath 'child.Status = "NestedMissingFromParent";' 'missing nested child keeps absence distinct from capture failure'
	Assert-Contains $nestedPropagationPath 'child.Status = "NestedExtraInParent";' 'extra nested child keeps parent-composition state'
	Assert-Contains $nestedPropagationPath 'Nested family missing from parent family:' 'missing nested child summary names the relationship failure'
} else {
	Add-Failure "Missing nested-family difference propagation service: $nestedPropagationPath"
}

$nestedPropagationTestPath = Join-Path $root 'KKY_FamilyBrowser_Compile\Test-NestedFamilyDifferencePropagation.ps1'
$qualityGatePath = Join-Path $root 'KKY_FamilyBrowser_Compile\Invoke-FamilyBrowserQualityGate.ps1'
if (Test-Path -LiteralPath $nestedPropagationTestPath) {
	Assert-Contains $nestedPropagationTestPath 'The deepest nested-family reason did not reach the top-level parent detail rows.' 'transitive nested difference behavior test'
	Assert-Contains $nestedPropagationTestPath 'A child of a wholly absent parent was incorrectly exposed as a separate difference.' 'absent-parent nested false-positive test'
	Assert-Contains $nestedPropagationTestPath 'Missing nested child did not receive the explicit nested-missing status.' 'nested child missing-state behavior test'
} else {
	Add-Failure "Missing nested-family propagation behavior test: $nestedPropagationTestPath"
}
Assert-Contains $qualityGatePath 'Test-NestedFamilyDifferencePropagation.ps1' 'quality gate invokes nested-family behavior test'

$yesNoFormatterPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserYesNoParameterFormatter.cs'
if (Test-Path -LiteralPath $yesNoFormatterPath) {
	Assert-Contains $yesNoFormatterPath 'GetDataType' 'new Revit API Yes/No data-type detection'
	Assert-Contains $yesNoFormatterPath 'ParameterType' 'legacy Revit API Yes/No parameter-type detection'
	Assert-Contains $yesNoFormatterPath '"YesNo"' 'legacy Yes/No enum name'
	Assert-Contains $yesNoFormatterPath '"boolean"' 'ForgeTypeId boolean detection'
	Assert-Contains $yesNoFormatterPath 'return value == 0 ? "No" : "Yes";' 'stable Yes/No display mapping'
} else {
	Add-Failure "Missing shared Yes/No formatter: $yesNoFormatterPath"
}

$systemRoutingUnitUiPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserSystemRoutingUnitUi.cs'
if (Test-Path -LiteralPath $systemRoutingUnitUiPath) {
	Assert-Contains $systemRoutingUnitUiPath 'parseSystemRoutingCriteria' 'system routing size criteria parser'
	Assert-Contains $systemRoutingUnitUiPath 'piece.match(/min\s*=\s*([^\s;]+)/i)' 'system routing parser reads minimum from current and legacy malformed criteria'
	Assert-NotContains $systemRoutingUnitUiPath 'piece.match(/\bmin\s*=' 'system routing parser does not reject legacy size 1min criteria'
	Assert-Contains $systemRoutingUnitUiPath 'system-routing-criteria-table' 'system routing nested range table'
	Assert-Contains $systemRoutingUnitUiPath 'system-routing-unit-select' 'system routing unit selector'
	Assert-Contains $systemRoutingUnitUiPath 'changeSystemRoutingUnit' 'system routing in/mm live unit switch'
	Assert-Contains $systemRoutingUnitUiPath 'setSystemDisplayUnitFromHost' 'host-to-main-and-detail measurement unit synchronization API'
	Assert-Contains $systemRoutingUnitUiPath 'kkyfb:measurement-unit/' 'measurement unit persistence action'
	Assert-Contains $systemRoutingUnitUiPath 'system-layer-composition-table' 'Revit-style compound layer table'
	Assert-Contains $systemRoutingUnitUiPath 'data-system-layer-composition' 'compound layer semantic surface'
	Assert-Contains $systemRoutingUnitUiPath 'parseSystemLayerModel' 'structured and legacy compound layer parser'
	Assert-Contains $systemRoutingUnitUiPath "p[0]!='@layer'" 'structured compound layer row parser marker'
	Assert-Contains $systemRoutingUnitUiPath 'system-layer-thickness-value' 'compound layer live unit value'
	Assert-Contains $systemRoutingUnitUiPath 'systemLayerCoreBoundaryHtml' 'compound layer core boundary renderer'
	Assert-Contains $systemRoutingUnitUiPath 'systemLayerExteriorLabel' 'compound layer exterior direction label'
	Assert-Contains $systemRoutingUnitUiPath 'systemLayerInteriorLabel' 'compound layer interior direction label'
	Assert-Contains $systemRoutingUnitUiPath '304.8' 'internal feet to millimetres conversion'
	Assert-Contains $systemRoutingUnitUiPath "unit=='in'?12:304.8" 'internal feet to inches conversion'
	Assert-Contains $systemRoutingUnitUiPath 'Math.abs(value)>=1e20' 'unbounded Revit size sentinel guard'
	Assert-Contains $systemRoutingUnitUiPath 'No limit' 'English unbounded range label'
	Assert-Contains $systemRoutingUnitUiPath '제한 없음' 'Korean unbounded range label'
} else {
	Add-Failure "Missing shared system routing unit UI: $systemRoutingUnitUiPath"
}

$systemComponentUnitUiPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserSystemDetailedComponentUnitUi.cs'
if (Test-Path -LiteralPath $systemComponentUnitUiPath) {
	Assert-Contains $systemComponentUnitUiPath "p[0]=='@component'" 'structured dependent-component parser'
	Assert-Contains $systemComponentUnitUiPath "p[0]=='@component-diff'" 'structured dependent-component difference parser'
	Assert-Contains $systemComponentUnitUiPath 'parseCurtainPanelComponentRows=function' 'curtain panel structured parser override'
	Assert-Contains $systemComponentUnitUiPath 'system-component-unit-select' 'Railing, Stair, and curtain component unit selector'
	Assert-Contains $systemComponentUnitUiPath 'formatSystemRoutingSize(raw,systemRoutingDisplayUnit)' 'shared raw-feet component conversion'
	Assert-Contains $systemComponentUnitUiPath 'changeSystemRoutingUnit(this)' 'shared persisted component unit switch'
} else {
	Add-Failure "Missing shared detailed component unit UI: $systemComponentUnitUiPath"
}

$measurementUnitPreferencePath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserMeasurementUnitPreferenceService.cs'
if (Test-Path -LiteralPath $measurementUnitPreferencePath) {
	Assert-Contains $measurementUnitPreferencePath 'measurement-unit.txt' 'persistent measurement unit settings file'
	Assert-Contains $measurementUnitPreferencePath 'Environment.SpecialFolder.LocalApplicationData' 'measurement unit LocalAppData settings root'
	Assert-Contains $measurementUnitPreferencePath 'LoadFromPathForAudit' 'measurement unit persisted-read audit seam'
	Assert-Contains $measurementUnitPreferencePath 'SaveToPathForAudit' 'measurement unit persisted-write audit seam'
	Assert-Contains $measurementUnitPreferencePath 'return string.Equals' 'measurement unit normalization with mm fallback'
} else {
	Add-Failure "Missing shared measurement unit preference service: $measurementUnitPreferencePath"
}

$sharedModelsPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserPerformanceModels.cs'
$sharedLoaderPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserDataLoader.cs'
if (Test-Path -LiteralPath $sharedModelsPath) {
    Assert-Contains $sharedModelsPath 'public interface IFamilyBrowserDataLoader' 'shared browser data-loader interface'
    Assert-Contains $sharedModelsPath 'class FamilyBrowserManifestV2' 'V2 manifest contract'
    Assert-Contains $sharedModelsPath 'class BrowserIndexItem' 'compact browser index item contract'
    Assert-Contains $sharedModelsPath 'class BrowserDetailRecord' 'lazy browser detail contract'
    Assert-Contains $sharedModelsPath 'class ThumbnailIndexEntry' 'thumbnail index contract'
    Assert-Contains $sharedModelsPath 'class FamilyBrowserLoadGeneration' 'stale async generation guard'
} else {
    Add-Failure "Missing shared performance models: $sharedModelsPath"
}
if (Test-Path -LiteralPath $sharedLoaderPath) {
    Assert-Contains $sharedLoaderPath 'family-browser-manifest-v2.json' 'V2 manifest artifact'
    Assert-Contains $sharedLoaderPath 'standard-browser-index-v2.json' 'V2 compact index artifact'
    Assert-Contains $sharedLoaderPath 'standard-browser-details-v2.json' 'V2 detail catalog artifact'
    Assert-Contains $sharedLoaderPath 'thumbnail-index-v2.json' 'V2 thumbnail artifact'
    Assert-Contains $sharedLoaderPath 'project-browser-state-v2.json' 'V2 project state artifact'
    Assert-Contains $sharedLoaderPath 'browser-row-cache-v2-' 'persistent browser row cache artifact'
    Assert-Contains $sharedLoaderPath 'LocalApplicationData' 'LocalAppData read-through cache root'
    Assert-Contains $sharedLoaderPath 'CopyFileAtomic' 'atomic remote-to-local cache copy'
    Assert-Contains $sharedLoaderPath 'ValidateCurrentSourceRevision' 'write-time source revision validator'
    Assert-Contains $sharedLoaderPath 'RunSyntheticPerformanceAudit' 'cold/warm/offline cache performance audit'
    Assert-Contains $sharedLoaderPath 'thumbnail-index-build' 'one-time thumbnail metadata indexing metric'
} else {
    Add-Failure "Missing shared data loader: $sharedLoaderPath"
}

$sharedRowWindowPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\family-browser-row-window.js'
if (Test-Path -LiteralPath $sharedRowWindowPath) {
	Assert-Contains $sharedRowWindowPath 'api.setRows = function' 'virtual row store API'
	Assert-Contains $sharedRowWindowPath 'data-kkyfb-virtualized' 'virtualized table DOM marker'
	Assert-Contains $sharedRowWindowPath 'Math.min(store.filtered.length, start + windowSize)' '150-row DOM slice'
	Assert-Contains $sharedRowWindowPath 'w.checkedFamilyRows = function' 'cross-page family checked-row adapter'
	Assert-Contains $sharedRowWindowPath 'w.checkedSystemRows = function' 'cross-page system checked-row adapter'
	Assert-Contains $sharedRowWindowPath 'w.goToRowWindowPage = function' 'direct numbered page navigation API'
	Assert-Contains $sharedRowWindowPath 'renderPageLinks(pages, page, totalPages)' 'numbered page link renderer'
	Assert-Contains $sharedRowWindowPath 'function ensureColumnResizers(tab)' 'family/system header resize handle runtime'
	Assert-Contains $sharedRowWindowPath 'w.resizeColumnForAudit = function' 'column resize audit seam'
	Assert-Contains $sharedRowWindowPath 'function lockTablePixelWidth(table, total)' 'exact table-width lock for isolated column resizing'
	Assert-Contains $sharedRowWindowPath 'rememberColumnWidths(tab, baseline);' 'explicit baseline width freeze before column dragging'
	Assert-Contains $sharedRowWindowPath "kkyfb-column-width-locked" 'isolated column-resize table marker'
	Assert-Contains $sharedRowWindowPath 'kkyfb-column-widths-v1:' 'persisted user column widths'
	Assert-MinimumOccurrences $sharedRowWindowPath 'refreshOverflowTitles' 2 'column drag and audit resize refresh generated overflow titles'
} else {
	Add-Failure "Missing shared virtual row runtime: $sharedRowWindowPath"
}

$overflowTitlePath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserOverflowTitleScript.cs'
if (Test-Path -LiteralPath $overflowTitlePath) {
	Assert-Contains $overflowTitlePath "var generated='data-kkyfb-overflow-title'" 'generated overflow title ownership marker'
	Assert-Contains $overflowTitlePath 'el.scrollWidth>el.clientWidth+1||el.scrollHeight>el.clientHeight+1' 'actual horizontal or vertical clipping test'
	Assert-Contains $overflowTitlePath 'if(hasAuthorTitle(el))return 2;' 'authored title preservation'
	Assert-Contains $overflowTitlePath 'if(!clipped){clearGenerated(el);return 0;}' 'stale generated title removal after widening'
	Assert-Contains $overflowTitlePath "d.addEventListener('mouseover',inspectEvent,true)" 'delegated hover inspection for dynamic browser text'
} else {
	Add-Failure "Missing shared overflow title service: $overflowTitlePath"
}

$managedFolderSetupPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserManagedFolderSetupService.cs'
$managedFolderTransitionPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserDashboardManagedFolderTransition.cs'
if (Test-Path -LiteralPath $managedFolderSetupPath) {
	Assert-Contains $managedFolderSetupPath 'managed-folder-override.txt' 'per-user persisted management-folder pointer'
	Assert-Contains $managedFolderSetupPath 'if (!IsInternalNetworkShare(root))' 'manual management-folder network-only validation'
	Assert-Contains $managedFolderSetupPath 'drive.DriveType == DriveType.Network' 'mapped network drive validation'
	Assert-Contains $managedFolderSetupPath 'CreateAndProbeManagedFolders(root);' 'selected share write permission probe'
	Assert-Contains $managedFolderSetupPath 'KKY_FAMILY_BROWSER_MANAGED_FOLDER_README.txt' 'managed-folder safety readme'
	Assert-Contains $managedFolderSetupPath '수동으로 수정, 이동, 이름 변경 또는 삭제하지 마세요.' 'managed-folder safety warning content'
	Assert-Contains $managedFolderSetupPath 'FamilyBrowserStandardPolicyStore.SetRequestStore' 'manual management folder becomes request store'
	Assert-Contains $managedFolderSetupPath 'AnalyzeMigration' 'management-folder migration preflight'
	Assert-Contains $managedFolderSetupPath 'MigrateToHomepage' 'TEST-to-homepage managed-data migration'
	Assert-Contains $managedFolderSetupPath '"ProjectCatalogs"' 'project catalog baseline is included in managed-folder setup and migration'
	Assert-Contains $managedFolderSetupPath '"StandardRevisionManifests"' 'Standard RVT revision baseline is included in managed-folder setup and migration'
	Assert-Contains $managedFolderSetupPath 'TryClearOverrideRoot' 'TEST override pointer is cleared only through a verified result'
	Assert-Contains $managedFolderSetupPath 'A managed file changed after preflight. Nothing was overwritten' 'migration refuses to overwrite a changed homepage file'
	Assert-Contains $managedFolderSetupPath 'RebaseJsonStringValues' 'structured JSON path rebasing during migration'
	Assert-Contains $managedFolderSetupPath 'FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(destinationFile)' 'managed-folder migration avoids long temporary suffixes'
	Assert-Contains $managedFolderSetupPath 'name.StartsWith(".kky-r-"' 'request lock files are excluded from managed-folder migration'
} else {
	Add-Failure "Missing shared managed-folder setup service: $managedFolderSetupPath"
}
if (Test-Path -LiteralPath $managedFolderTransitionPath) {
	Assert-Contains $managedFolderTransitionPath 'managed-folder-switch-homepage' 'homepage-folder switch UI action'
	Assert-Contains $managedFolderTransitionPath 'managed-folder-migrate-homepage' 'migrate-and-switch UI action'
	Assert-Contains $managedFolderTransitionPath 'The TEST source folder will not be deleted' 'TEST source retention confirmation'
	Assert-Contains $managedFolderTransitionPath 'TryApplyPersistedOverride' 'failed homepage activation restores the TEST override'
	Assert-Contains $managedFolderTransitionPath 'TryClearOverrideRoot' 'homepage activation verifies TEST pointer removal'
	Assert-Contains $managedFolderTransitionPath 'FamilyBrowserTrackingPersistenceService.FlushPending(_workspaceRoot)' 'homepage activation immediately flushes locally protected tracking records'
	Assert-Contains $managedFolderTransitionPath 'AppendManagedFolderTransitionPanel' 'Home and Admin managed-folder transition component'
} else {
	Add-Failure "Missing shared managed-folder transition UI: $managedFolderTransitionPath"
}

$sharedShellCssPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\family-browser-shell.css'
$sharedShellJsPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\family-browser-shell.js'
$nestedOnlyCatalogPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserNestedOnlyPlacementCatalog.cs'
Assert-Contains $sharedShellCssPath '.row-window-page.active' 'active numbered page visual state'
Assert-Contains $sharedShellCssPath 'table.kkyfb-column-width-locked' 'resizable table minimum-width override'
Assert-Contains $sharedShellCssPath '.column-resize-handle' 'column resize cursor handle visual'
Assert-Contains $sharedShellCssPath '.admin-check-leading' 'IE-safe admin checkbox leading cell style'
Assert-Contains $sharedShellCssPath 'display: table-cell !important;' 'IE-safe admin option table-cell layout'
Assert-Contains $sharedShellCssPath '.admin-check-state:before' 'admin option visual state indicator'
Assert-Contains $sharedShellCssPath '.managed-folder-transition.return-ready' 'homepage-return-ready managed-folder panel styling'
Assert-Contains $sharedShellCssPath '.managed-folder-transition-paths' 'managed-folder source and destination path layout'
Assert-Contains $sharedShellCssPath '.disciplinebar a.discipline-filter.ready' 'registered trade selector visual state'
Assert-Contains $sharedShellCssPath '.disciplinebar a.discipline-filter.unavailable' 'unregistered trade selector visual state'
Assert-Contains $sharedShellCssPath '.disciplinebar a.discipline-filter.active' 'selected trade selector emphasis'
Assert-Contains $sharedShellCssPath '.permission-diagnostic-grid .diagnostic-card' 'Permission/Guard diagnostics stack as full-width rows'
Assert-Contains $sharedShellCssPath '.permission-diagnostic-grid .diagnostic-detail' 'Permission/Guard path cell owns the remaining row width'
Assert-Contains $sharedShellJsPath 'styleBrowserPane(pane, tab, searchH, 50)' 'browser pane reserves expanded paging status bar'
Assert-Contains $sharedShellJsPath 'Math.ceil(search.scrollHeight || 0) + 2' 'browser search area measures the remaining search and trade controls'
Assert-Contains $sharedShellJsPath 'var minH = small ? 86 : 84;' 'browser search area releases the removed status-row height'
Assert-Contains $sharedShellJsPath 'function resetDetailedFilterState()' 'shared shell has a complete detailed-filter reset'
Assert-Contains $sharedShellJsPath 'isBrowser(previousTab) && isBrowser(nextTab) && previousTab !== nextTab' 'Family/System cross-tab switch resets detailed filters'
Assert-Contains $sharedShellJsPath "if (status) status.value = 'All';" 'cross-tab reset updates the detailed status control'
Assert-Contains $standardRevisionServicePath 'state.SnapshotPath = manifest.SnapshotPath' 'Standard revision state carries the published snapshot path'
Assert-Contains $standardRevisionServicePath 'state.SnapshotAtUtc = manifest.SnapshotAtUtc' 'Standard revision state carries the published snapshot capture time'
Assert-Contains $standardRevisionServicePath 'public static bool MatchesSnapshotGeneration' 'Standard revision service can prove the loaded snapshot generation'
Assert-WindowContains $standardRevisionServicePath 'public static string BuildCurrentRevisionToken' 'state.SnapshotAtUtc' 'Standard revision token changes when the registered snapshot generation changes' 30

foreach ($folder in $HostFolders) {
    $projectPath = Get-ChildItem -LiteralPath (Join-Path $root $folder) -Filter '*.csproj' | Select-Object -First 1 -ExpandProperty FullName
    if ($projectPath) {
        Assert-Contains $projectPath 'KKY_FamilyBrowser_SharedUi\*.cs' 'host links shared UI/performance sources'
		Assert-Contains $projectPath 'family-browser-row-window.js' 'host embeds shared virtual row runtime'
    }
    $registryPath = Join-Path $root "$folder\StandardLibraryRegistryStore.cs"
	$projectStorePath = Join-Path $root "$folder\ProjectSnapshotStore.cs"
	$projectTrackingStorePath = Join-Path $root "$folder\ProjectTrackingStoreService.cs"
	$loadableFamilySyncStorePath = Join-Path $root "$folder\LoadableFamilySyncStore.cs"
	$systemTypeApplyStorePath = Join-Path $root "$folder\SystemTypeApplyStore.cs"
	$systemTypePreflightStorePath = Join-Path $root "$folder\SystemTypePreflightStore.cs"
	$projectTrackingOutputStorePath = Join-Path $root "$folder\ProjectTrackingOutputStore.cs"
    $thumbnailPath = Join-Path $root "$folder\FamilyThumbnailPreviewService.cs"
	$userSettingsPath = Join-Path $root "$folder\FamilyBrowserUserSettingsStore.cs"
	$nativeGuardPath = Join-Path $root "$folder\FamilyBrowserNativeCommandGuardService.cs"
	$familySyncPath = Join-Path $root "$folder\LoadableFamilySyncExecutionService.cs"
	$systemApplyPath = Join-Path $root "$folder\SystemTypeApplyExecutionService.cs"
	$dashboardPath = Join-Path $root "$folder\FamilyBrowserDashboardHtmlForm.cs"
	$errorHelpPath = Join-Path $root "$folder\FamilyBrowserErrorHelp.cs"
	$diagnosticsOutputPath = Get-ChildItem -LiteralPath (Join-Path $root $folder) -Recurse -Filter 'DiagnosticsOutputStore.cs' | Select-Object -First 1 -ExpandProperty FullName
	$bootstrapPath = Join-Path $root "$folder\FamilyBrowserDeploymentBootstrapService.cs"
	$fileGuardFormPath = Join-Path $root "$folder\FileGuardHtmlConfigurationForm.cs"
	$projectCacheRecordPath = Join-Path $root "$folder\ProjectScanCacheRecord.cs"
	$appPath = if ($folder -eq 'KKY_FamilyBrowser_RevitHost_2019-2023') {
		Join-Path $root "$folder\KKY_FamilyBrowser_RevitHost_2019_2023\App.cs"
	} else {
		Join-Path $root "$folder\$folder\App.cs"
	}
    Assert-Contains $registryPath 'PublishStandardArtifacts' 'standard scan publishes V2 browser artifacts'
	Assert-Contains $registryPath 'FamilyBrowserNestedOnlyPlacementCatalogStore.SaveForSnapshot' 'standard scan publishes nested-only placement guard catalog'
    Assert-Contains $projectStorePath 'PublishProjectState' 'project scan publishes V2 browser state'
	Assert-Contains $projectStorePath 'string.IsNullOrWhiteSpace(GetProjectScanCacheFolder(workspaceRoot))' 'project cache lookup resolves the homepage managed root when workspaceRoot is empty'
	Assert-Contains $projectStorePath 'IsSafeProjectAliasRecord(record, projectIdentity)' 'project alias lookup rejects same-name collisions by file stamp'
	Assert-Contains $projectStorePath 'ProjectDocumentRevisionToken' 'project scan cache records a Revit document revision token'
	Assert-Contains $projectStorePath 'BasicFileInfo.Extract(documentPath)' 'workshared cache freshness reads the Revit central episode revision'
	Assert-Contains $projectStorePath 'projectIdentity.IsDocumentModified' 'unsaved live documents cannot reuse a stale project scan'
	Assert-Contains $projectStorePath 'public static bool CanPublishSharedProjectState' 'shared project publication has one authoritative readiness guard'
	Assert-Contains $projectStorePath 'AllLocalChangesSavedToCentral' 'workshared publication rejects locally saved but unsynchronized data'
	Assert-Contains $projectStorePath 'centralInfo.LatestCentralVersion > localInfo.LatestCentralVersion' 'workshared publication rejects a local file behind Central'
	Assert-WindowContains $projectStorePath 'public static ProjectScanCacheRecord SaveLatestProjectScan' 'CanPublishSharedProjectState(doc, out publicationReason)' 'latest project cache publication is guarded at the persistence boundary' 28
	Assert-WindowContains $projectStorePath 'public static ProjectScanCacheRecord SaveLatestProjectScan' 'FamilyBrowserStandardRevisionService.MatchesSnapshotGeneration' 'latest project cache publication rejects a replaced Standard snapshot at the persistence boundary' 36
	Assert-WindowContains $projectStorePath 'public static ProjectScanCacheLoadResult TryLoadLatestProjectScan(string workspaceRoot, Document doc' 'CanPublishSharedProjectState(doc, out publicationReason)' 'live project cache reuse is guarded by worksharing publication readiness' 28
	Assert-Contains $projectStorePath 'GetProjectCoordinationFolder' 'multi-PC scan locks use a dedicated physical-identity folder'
	Assert-Contains $projectStorePath 'TryAcquireProjectPublicationLock' 'manual and automatic project publications share one lock'
	Assert-Contains $projectStorePath 'FamilyBrowserAtomicFileService.Promote' 'project latest and alias records use atomic promotion'
	Assert-Contains $projectStorePath 'recordedIdentity' 'non-primary project aliases require a matching physical or canonical identity'
	Assert-Contains $projectTrackingStorePath 'AcquireDirtyMarkerMutationLock' 'protected-content marker mutations are serialized across Revit clients'
	Assert-Contains $projectTrackingStorePath 'MergeDirtyMarkerItems' 'concurrent protected-content findings are merged instead of overwritten'
	Assert-Contains $projectTrackingStorePath 'FamilyBrowserAtomicFileService.Promote' 'protected-content marker is atomically promoted'
	Assert-WindowContains $projectTrackingStorePath 'public static bool ClearCurrentModelCheckRequired' 'DirtyMarkerIdentityMatches' 'completed checks cannot clear a newer protected-content marker' 26
	Assert-MinimumOccurrences $dashboardPath 'ClearCurrentModelCheckRequired(doc, dirtyMarkerAtCheckStart)' 2 'manual checks and tracking stamps clear only their starting marker generation'
	Assert-Contains $loadableFamilySyncStorePath 'FamilyBrowserUniqueJsonReportStore.Save' 'Family load reports use collision-free atomic history storage'
	Assert-Contains $systemTypeApplyStorePath 'FamilyBrowserUniqueJsonReportStore.Save' 'System Type apply reports use collision-free atomic history storage'
	Assert-Contains $systemTypePreflightStorePath 'FamilyBrowserUniqueJsonReportStore.Save' 'System Type preflight reports use collision-free atomic history storage'
	Assert-Contains $projectTrackingOutputStorePath 'FamilyBrowserUniqueJsonReportStore.Save' 'tracking stamp reports use collision-free atomic history storage'
	Assert-Contains $projectCacheRecordPath 'SchemaVersion = 4' 'project scan cache schema records dirty-writer validation'
	Assert-Contains $projectCacheRecordPath 'CapturedFromModifiedDocument' 'project scan cache persists writer dirty state'
	Assert-Contains $dashboardPath 'ProjectSnapshotStore.TryAcquireProjectPublicationLock' 'manual Current Model Check uses the common project publication lock'
	Assert-MinimumOccurrences $dashboardPath 'base.CancelButton = cancelButton;' 2 'prompt and deployment-profile dialogs wire Esc to their visible cancel button'
	Assert-NotContains $dashboardPath 'cancelButton = cancelButton;' 'self-assignment does not silently disable dialog keyboard cancellation'
	Assert-MinimumOccurrences $dashboardPath 'ValidateStandardRevisionAfterOperation' 2 'manual Current Model Check revalidates the Standard RVT after capture and before publication'
	Assert-Contains $dashboardPath 'finalPublicationReason' 'manual Current Model Check performs a final project publication readiness check'
	Assert-WindowContains $appPath 'private void HandleDocumentOpened' 'RunEventHandlerSafely("Managed policy document-open preparation failed"' 'document-open managed-policy failure is isolated' 28
	Assert-WindowContains $appPath 'private void HandleDocumentOpened' 'RunEventHandlerSafely("Element tracking document-open baseline failed"' 'document-open tracking failure is isolated' 28
	Assert-WindowContains $appPath 'private void HandleDocumentOpened' 'RunEventHandlerSafely("Native guard document-open policy preload failed"' 'document-open native guard failure is isolated' 28
	Assert-WindowContains $appPath 'private void HandleDocumentOpened' 'RunEventHandlerSafely("Automatic Current Model Check document-open scheduling failed"' 'document-open automatic check failure is isolated' 28
	Assert-WindowContains $appPath 'private void HandleViewActivated' 'RunEventHandlerSafely("Native guard view-activation policy preload failed"' 'view-activation native guard preload failure is isolated' 28
	Assert-WindowContains $appPath 'private void HandleDocumentClosing' 'RunEventHandlerSafely("Automatic Current Model Check document-closing cleanup failed"' 'document-closing automatic check cleanup failure is isolated' 40
	Assert-WindowContains $appPath 'private static void ObserveProjectCatalogAfterCommit' 'RestoreProjectCatalogObservationRequired(document)' 'failed project catalog publication restores its retry decision' 40
	Assert-Contains $dashboardPath 'The project changed while Current Model Check was running.' 'manual Current Model Check rejects mid-capture project mutation'
	Assert-Contains $fileGuardFormPath 'BuildStablePolicyPathKey' 'File Guard UI deduplicates mapped and canonical paths consistently'
	Assert-WindowNotContains $fileGuardFormPath 'private FamilyBrowserFileGuardPolicy BuildExcelPolicyFromRows()' 'GetEffectiveSlot(_standardPolicy)' 'File Guard export does not silently replace a missing per-file trade' 80
	Assert-NotContains $projectStorePath 'string.IsNullOrWhiteSpace(workspaceRoot) || string.IsNullOrWhiteSpace(projectKey)' 'legacy empty-workspace project cache read guard'
    Assert-Contains $thumbnailPath 'ResolveThumbnailPath' 'thumbnail lookup uses shared single index'
	Assert-Contains $thumbnailPath 'FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path)' 'thumbnail metadata uses a same-folder temporary file'
	Assert-Contains $thumbnailPath 'stream.Flush(true);' 'thumbnail metadata is durably flushed before publication'
	Assert-Contains $thumbnailPath 'FamilyBrowserAtomicFileService.Promote(temporaryPath, path);' 'thumbnail metadata is atomically promoted'
	Assert-NotContains $thumbnailPath 'File.WriteAllText(path, PlainJsonReportWriter.Serialize(metadata), Encoding.UTF8);' 'thumbnail metadata is not overwritten in place'
	Assert-Contains $userSettingsPath 'Environment.SpecialFolder.LocalApplicationData' 'Admin/language settings use local user storage'
	Assert-Contains $userSettingsPath '"KKY", "FamilyBrowser", "Settings"' 'local user settings path is stable'
	Assert-NotContains $userSettingsPath 'ResolveManagedUserSettingsRoot' 'Admin/language settings do not block on the managed share'
	Assert-Contains $userSettingsPath 'public static bool HasAdminModePreference()' 'Admin mode distinguishes a real user selection from a missing first-run preference'
	Assert-Contains $userSettingsPath 'public static bool ResolveInitialAdminModeEnabled(bool canEnableAdminMode)' 'first-run Admin profiles default to Admin Mode ON without overwriting later choices'
	Assert-Contains $nativeGuardPath 'CentralPathCache' 'native guard caches workshared central identity'
	Assert-Contains $nativeGuardPath 'FamilyBrowserFileGuardPathMatcher.FindMatchingTarget' 'native guard delegates target selection to the shared path matcher'
	Assert-Contains $nativeGuardPath 'FamilyBrowserFileGuardPathMatcher.Resolve' 'native permission diagnostics use the shared path decision'
	Assert-NotContains $nativeGuardPath 'UNC-RELATIVE|' 'native guard contains no endpoint-neutral path bypass'
	Assert-NotContains $nativeGuardPath 'BuildEndpointNeutralUncKey' 'native guard contains no duplicate unsafe UNC matcher'
	Assert-Contains $nativeGuardPath 'ResolveCentralPath(document, allowResolve: true)' 'active document captures central identity once'
	Assert-Contains $nativeGuardPath 'RegisterProtectedChangeUpdater(application);' 'protected family/type updater is registered inside Revit OnStartup API context'
	Assert-WindowNotContains $nativeGuardPath 'private static void RefreshProtectedChangeUpdaterRegistration(Document document)' 'RegisterProtectedChangeUpdater(application);' 'modeless Admin/document changes never attempt illegal UpdaterRegistry registration'
	Assert-WindowNotContains $nativeGuardPath 'private static void RefreshProtectedChangeUpdaterRegistration(Document document)' 'UnregisterProtectedChangeUpdater();' 'modeless Admin/document changes keep the startup updater registered'
	Assert-WindowContains $nativeGuardPath 'public static void NotifyAdminModeChanged(bool enabled, FamilyBrowserStandardPolicy policy, bool refreshUiNow = true)' 'NativeGuardDecisionCache.Clear();' 'Admin ON/OFF invalidates native edit/type permission decisions immediately'
	Assert-Contains $nativeGuardPath 'public static string BuildRuntimeGuardDiagnostic(Document document)' 'runtime exposes target, permission, command-binding, and updater evidence'
	Assert-WindowContains $nativeGuardPath 'private static bool CanNativeGuardPermission' 'string.Equals(permission, "RenameFamilyOrType"' 'family/type rename uses an explicit combined permission route'
	Assert-WindowContains $nativeGuardPath 'private static bool CanNativeGuardPermission' '"EditFamilies"' 'family/type rename checks the family-edit guard'
	Assert-WindowContains $nativeGuardPath 'private static bool CanNativeGuardPermission' '"AddDeleteTypes"' 'family/type rename checks the type-change guard'
	Assert-Contains $nativeGuardPath '"ID_PRJBROWSER_RENAME"' 'actual Revit Project Browser F2 rename command is intercepted before every rename transaction'
	Assert-Contains $nativeGuardPath 'BuiltInParameter.ALL_MODEL_TYPE_NAME' 'system and loadable type-name changes have an explicit updater trigger'
	Assert-Contains $nativeGuardPath ';projectBrowserRenameBinding=' 'runtime diagnostics expose the real Project Browser rename binding state'
	Assert-Contains $nativeGuardPath 'private static bool ShouldRecordProtectedChange(string action, bool previousInfoAvailable, bool sameProtectedInfo)' 'protected-change classifier exposes the missing-baseline fail-closed seam'
	Assert-WindowContains $nativeGuardPath 'private static bool ShouldRecordProtectedChange' '!previousInfoAvailable' 'the first modification of an unseen Family or Type is blocked'
	Assert-WindowContains $nativeGuardPath 'public static void HandleDocumentChanged' 'UpdateProtectedElementIndexFromChanges(doc, e.GetAddedElementIds(), e.GetModifiedElementIds(), e.GetDeletedElementIds());' 'permitted Admin ON changes keep the guard index current' 64
	Assert-Contains $nativeGuardPath 'application.Idling += HandleIdling;' 'native ribbon state receives a Revit idle settle pass after Admin transitions'
	Assert-WindowContains $nativeGuardPath 'private static void HandleIdling' 'UpdateProtectedRibbonAvailability(force: true);' 'idle settle pass reapplies the native Load Family state after Revit refreshes its ribbon' 96
	Assert-WindowContains $nativeGuardPath 'private static void HandleIdling' 'EnsureProtectedElementIndexForGuard(LastActiveDocument);' 'Revit idle builds the protected Family/Type name baseline outside modeless UI callbacks' 96
	Assert-WindowContains $nativeGuardPath 'private static void EnsureProtectedElementIndexForGuard' 'EnsureProtectedElementIndexBaseline(doc);' 'Admin OFF protected documents receive a complete name baseline' 48
	Assert-WindowContains $nativeGuardPath 'private static void RefreshProtectedElementIndex' 'CompleteProtectedElementIndexDocumentTokens[documentKey]' 'complete baseline is distinguished from a partial change index' 64
	Assert-WindowContains $nativeGuardPath 'public static void NotifyAdminModeChanged(bool enabled, FamilyBrowserStandardPolicy policy, bool refreshUiNow = true)' 'ScheduleProtectedRibbonRefresh();' 'Admin ON/OFF schedules deferred ribbon state settlement' 64
	Assert-WindowContains $nativeGuardPath 'public static void NotifyAdminModeChanged(bool enabled, FamilyBrowserStandardPolicy policy, bool refreshUiNow = true)' 'ScheduleProtectedElementBaselineRefresh();' 'Admin OFF schedules the original-name baseline before user edits' 64
	Assert-Contains $nativeGuardPath ';pendingRibbonRefreshPasses=' 'runtime diagnostics expose deferred ribbon refresh progress'
	Assert-Contains $nativeGuardPath ';protectedElementBaselineComplete=' 'runtime diagnostics expose complete versus partial protected-name baselines'
	Assert-Contains $nativeGuardPath ';pendingPostRollbackUiRefreshPasses=' 'runtime diagnostics expose delayed rollback UI refresh work'
	Assert-WindowContains $nativeGuardPath 'Key = "native-load-family"' 'RequiredPermission = "EditFamilies"' 'native Load Family command uses the family-edit guard'
	Assert-WindowContains $nativeGuardPath 'Key = "native-rename-family-or-type"' 'RequiredPermission = "RenameFamilyOrType"' 'native rename command uses the combined guard'
	Assert-WindowContains $nativeGuardPath 'private static bool LoadAdminModeEnabledSetting()' 'FamilyBrowserUserSettingsStore.LoadAdminModeEnabled()' 'native guard restores the persisted Admin mode state'
	Assert-WindowNotContains $nativeGuardPath 'private static bool LoadAdminModeEnabledSetting()' 'SaveAdminModeEnabled(enabled: true)' 'native guard does not force an Admin profile back to Admin Mode ON'
	Assert-Contains $nativeGuardPath 'FamilyLoadingIntoDocument += HandleFamilyLoadingIntoDocument;' 'Revit family-load database event is attached'
	Assert-Contains $nativeGuardPath 'FailureProcessingResult.ProceedWithRollBack' 'unauthorized family/type mutations roll back before commit'
	Assert-WindowContains $nativeGuardPath 'private static void HandleFailuresProcessing' 'SchedulePostRollbackUiRefresh();' 'protected rollback schedules immediate Revit UI settlement after the dialog closes' 64
	Assert-WindowContains $nativeGuardPath 'private static void RefreshRevitUiAfterProtectedRollback' 'uiDocument.RefreshActiveView();' 'post-rollback idle refreshes the visible Revit UI immediately' 48
	Assert-Contains $nativeGuardPath 'new ElementClassFilter(typeof(FamilyInstance)), Element.GetChangeTypeElementAddition()' 'nested-only guard observes newly placed family instances'
	Assert-Contains $nativeGuardPath 'instance.SuperComponent != null' 'nested-only guard allows child instances created by their parent family'
	Assert-Contains $nativeGuardPath 'CollectNestedOnlyStandalonePlacements' 'nested-only direct placements are evaluated before commit'
	Assert-Contains $nativeGuardPath 'FamilyIsNestedOnlyAndFingerprintMatchesStandard' 'nested-only placement audit records exact standard fingerprint evidence'
	Assert-Contains $nativeGuardPath 'ShouldBlockNestedOnlyPlacementMatch' 'nested-only placement exposes one exact-match blocking decision'
	Assert-WindowContains $nativeGuardPath 'private static bool ShouldBlockNestedOnlyPlacementMatch' 'FamilyBrowserNestedOnlyPlacementMatchState.ExactMatch' 'nested-only direct placement blocks only an exact standard fingerprint match' 16
	Assert-NotContains $nativeGuardPath 'NestedOnlyFamilyPlacementVerificationPending' 'pending fingerprint verification never rolls back a user placement'
	Assert-Contains $nativeGuardPath 'FamilyBrowserNestedOnlyPlacementRuntimeService.EvaluatePlacement' 'nested-only placement evaluates project and standard fingerprint equality before commit'
	Assert-Contains $nativeGuardPath 'target.BlockNestedOnlyStandalonePlacement && !adminModeEnabled' 'nested-only placement guard follows the file flag and Admin mode'
	Assert-Contains $nestedOnlyCatalogPath 'SchemaVersion = 2' 'nested-only catalog uses the fingerprint-backed v2 schema'
	Assert-Contains $nestedOnlyCatalogPath 'ContentFingerprint = item.ContentFingerprint' 'nested-only catalog persists the precise standard family fingerprint'
	Assert-Contains $nestedOnlyCatalogPath 'public static bool IsExactMatch' 'nested-only placement requires exact fingerprint and shared-state equality'
	Assert-Contains $nativeGuardPath 'public static void NotifyStandardSnapshotChanged()' 'native guard exposes immediate standard snapshot cache invalidation'
	Assert-Contains $dashboardPath 'FamilyBrowserNativeCommandGuardService.NotifyStandardSnapshotChanged();' 'standard scan refresh invalidates nested-only catalog immediately'
	Assert-Contains $familySyncPath 'BeginTrustedOperation("Family Browser loadable family sync")' 'approved Family Browser loads use the trusted guard scope'
	Assert-Contains $systemApplyPath 'BeginTrustedOperation("Family Browser system type apply")' 'approved system-type applies use the trusted guard scope'
	Assert-Contains $nativeGuardPath 'NotifyPolicyChanged(FamilyBrowserStandardPolicy policy)' 'saved file guard policy is handed directly to runtime guard'
	Assert-Contains $nativeGuardPath 'ApplyLoadFamilyRibbonEnabledFast(allowed);' 'Admin toggle directly synchronizes the native Load Family ribbon state'
	Assert-Contains $nativeGuardPath 'FindBoundDefinition("native-load-family") ?? FamilyLoadingEventDefinition' 'Load Family ribbon guard stays fail-closed when a Revit release exposes no bindable command id'
	Assert-Contains $nativeGuardPath 'ResolveLoadFamilyRibbonControlsFast' 'Load Family ribbon controls are discovered once and cached'
	Assert-NotContains $nativeGuardPath 'ApplyLoadFamilyRibbonEnabledRecursive' 'Admin toggle avoids recursive traversal of the full Autodesk ribbon object graph'
	Assert-Contains $nativeGuardPath ';ribbonControls=' 'runtime diagnostic reports matched native Load Family ribbon controls'
	Assert-Contains $dashboardPath 'IsManagedDataRootAvailable(_workspaceRoot)' 'file guard settings preflight managed storage'
	Assert-Contains $dashboardPath 'NotifyAdminModeChanged(_adminModeEnabled, _standardPolicy)' 'Admin toggle seeds native guard from current policy'
	Assert-WindowContains $dashboardPath 'private void SetAdminMode(bool enabled)' '_dashboardPermissionCachedSnapshot = null;' 'Admin toggle invalidates the browser permission snapshot immediately'
	Assert-WindowContains $dashboardPath 'private void SetAdminMode(bool enabled)' 'SynchronizeAdminModeGuardState' 'Admin toggle immediately verifies browser and native guard state'
	Assert-WindowContains $dashboardPath 'private static bool IsDashboardUiOnlyAction(string action)' 'case "detail-window-open":' 'UI-only action classifier remains available'
	Assert-WindowNotContains $dashboardPath 'private static bool IsDashboardUiOnlyAction(string action)' 'case "admin-mode-off":' 'Admin OFF runs through the Revit ExternalEvent context'
	Assert-WindowNotContains $dashboardPath 'private static bool IsDashboardUiOnlyAction(string action)' 'case "admin-mode-on":' 'Admin ON runs through the Revit ExternalEvent context'
	Assert-Contains $dashboardPath 'adminModeSwitch' 'header uses an explicit Admin ON/OFF segmented control'
	Assert-Contains $dashboardPath 'kkyfb:admin-mode-on' 'explicit Admin ON segment is routed'
	Assert-Contains $dashboardPath 'kkyfb:admin-mode-off' 'explicit Admin OFF segment is routed'
	Assert-WindowNotContains $dashboardPath 'private void OpenStandardRegistrationSetup()' '_adminModeEnabled = true' 'standard RVT CTA cannot silently turn Admin Mode back on'
	Assert-WindowNotContains $dashboardPath 'private void OpenStandardListRegistrationSetup()' '_adminModeEnabled = true' 'standard list CTA cannot silently turn Admin Mode back on'
	Assert-WindowContains $dashboardPath 'private void CompleteInitialOpenRefresh' 'ApplyAdminModeAfterPolicyLoad(restorePersistedSelection: true, refreshUiNow: false)' 'startup restores the selected Admin mode and defers native Revit UI work to a valid idle context'
	Assert-WindowNotContains $dashboardPath 'private void CompleteInitialOpenRefresh' '_adminModeEnabled = CanEnableAdminMode(_standardPolicy)' 'startup does not force an Admin profile back to Admin Mode ON'
	Assert-WindowContains $dashboardPath 'private void ApplyAdminModeAfterPolicyLoad' 'FamilyBrowserUserSettingsStore.ResolveInitialAdminModeEnabled(canEnable)' 'first run defaults eligible Admin profiles to ON while persisted choices remain authoritative'
	Assert-WindowContains $dashboardPath 'private void ApplyAdminModeAfterPolicyLoad' '_adminModeEnabled = ResolveEffectiveAdminMode(requestedEnabled, canEnable);' 'effective Admin mode requires both an ON selection and Admin capability'
	Assert-Contains $dashboardPath 'ApplyAdminModeAfterPolicyLoad(restorePersistedSelection: false, refreshUiNow: false)' 'management-folder changes preserve Admin OFF and defer native Revit UI work to a valid idle context'
	Assert-WindowContains $dashboardPath 'private bool HasPermission(string permission)' 'CanNativeGuard(policy, currentUser, permission, context, _adminModeEnabled)' 'browser permissions follow effective Admin mode'
	Assert-WindowContains $dashboardPath 'private DashboardPermissionSnapshot BuildDashboardPermissionSnapshot()' '|admin-mode=' 'browser permission cache is partitioned by Admin mode'
	Assert-WindowContains $dashboardPath 'private DashboardPermissionSnapshot BuildDashboardPermissionSnapshot()' 'CanNativeGuard(policy, user, "LoadFamilies", context, _adminModeEnabled)' 'family load button follows file guard and Admin mode'
	Assert-Contains $dashboardPath 'NotifyPolicyChanged(_standardPolicy)' 'file guard save updates runtime without a network reread'
	Assert-WindowContains (Join-Path $root "$folder\FamilyBrowserSecurityPolicyService.cs") 'private static FamilyBrowserFileGuardPermissionDecision ResolveFileGuardPermission' '!string.Equals(permission, "LoadFamilies"' 'file guard covers Family Browser load actions'
	Assert-Contains (Join-Path $root "$folder\FamilyBrowserSecurityPolicyService.cs") 'IsProjectElementTrackingScopeEnabled' 'file guard provides a shared per-document element tracking decision'
	Assert-Contains (Join-Path $root "$folder\FamilyBrowserSecurityPolicyService.cs") 'matchingTarget.TrackElementChanges' 'only a matching checked RVT enters element tracking scope'
	Assert-WindowNotContains (Join-Path $root "$folder\FamilyBrowserSecurityPolicyService.cs") 'public static bool IsProjectElementTrackingScopeEnabled' 'return true;' 'unregistered RVTs cannot enter element tracking scope'
	Assert-NotContains $dashboardPath "ClearModelerAllSlotsCache();`r`n`t`t`tRefreshDocumentShellOnly();" 'Admin toggle avoids full browser data reload'
	Assert-Contains $errorHelpPath 'ResolveLocalDiagnosticLogFolder' 'error logging falls back to LocalAppData'
	Assert-Contains $errorHelpPath 'ApplyManagedFolderUnavailable' 'managed-folder failures have specific user guidance'
	Assert-Contains $errorHelpPath 'Guid.NewGuid().ToString("N").Substring(0, 8)' 'error logs use collision-free filenames across clients'
	Assert-Contains $errorHelpPath 'FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path)' 'error logs use a same-folder temporary file'
	Assert-Contains $errorHelpPath 'stream.Flush(true);' 'error logs are durably flushed before publication'
	Assert-Contains $errorHelpPath 'FamilyBrowserAtomicFileService.Promote(temporaryPath, path);' 'error logs are atomically promoted'
	Assert-NotContains $errorHelpPath 'File.WriteAllText(path, builder.ToString(), Encoding.UTF8);' 'error logs are not overwritten in place'
	if ($diagnosticsOutputPath) {
		Assert-Contains $diagnosticsOutputPath 'FamilyBrowserUniqueJsonReportStore.Save' 'diagnostic reports use collision-free atomic history storage'
		Assert-NotContains $diagnosticsOutputPath 'File.WriteAllText' 'diagnostic reports are not overwritten in place'
	} else {
		Add-Failure "Missing DiagnosticsOutputStore.cs under $folder"
	}
	Assert-Contains $bootstrapPath 'result.Message = "Managed folder was not reachable.";' 'unreachable homepage path stops bootstrap cleanly'
}

$performanceGatePath = Join-Path $root 'KKY_FamilyBrowser_Compile\Test-FamilyBrowserPerformance.ps1'
if (Test-Path -LiteralPath $performanceGatePath) {
    Assert-Contains $performanceGatePath 'syntheticFamilyCount' 'performance gate 1,000-family scenario'
    Assert-Contains $performanceGatePath 'syntheticSystemCount' 'performance gate 1,000-system scenario'
    Assert-Contains $performanceGatePath 'WarmUsableTargetMs = 1500' 'warm-list acceptance target'
    Assert-Contains $performanceGatePath 'ColdUsableTargetMs = 3000' 'cold-list acceptance target'
    Assert-Contains $performanceGatePath 'FilterTargetMs = 150' 'search/filter acceptance target'
	Assert-Contains $performanceGatePath 'DomRows = [int]$result.domRowCount' 'performance report records true DOM row count'
    Assert-Contains $performanceGatePath 'SKIP runtime-not-installed' 'missing Revit runtime skip record'
} else {
    Add-Failure "Missing performance gate: $performanceGatePath"
}

$managedDataAuditPath = Join-Path $root 'KKY_FamilyBrowser_Compile\Test-FamilyBrowserManagedData.ps1'
if (Test-Path -LiteralPath $managedDataAuditPath) {
	Assert-Contains $managedDataAuditPath 'family-browser-manifest-v2.json' 'managed-data audit validates V2 manifests'
	Assert-Contains $managedDataAuditPath 'project-scan-latest*.json' 'managed-data audit validates latest project scan records'
	Assert-Contains $managedDataAuditPath 'Get-FileHash -LiteralPath $artifactPath -Algorithm SHA256' 'managed-data audit validates artifact hashes'
	Assert-Contains $managedDataAuditPath 'TreatUnavailableAsFailure' 'managed-data audit distinguishes external path unavailability from integrity failure'
} else {
	Add-Failure "Missing managed-data audit: $managedDataAuditPath"
}

$harnessPath = Join-Path $root 'KKY_FamilyBrowser_Automation\KKY_FamilyBrowser_UiAuditHarness\Program.cs'
if (Test-Path -LiteralPath $harnessPath) {
	Assert-Contains $harnessPath 'CheckProjectCatalogState(browser, options, result);' 'UI harness validates contracted project catalog states'
	Assert-Contains $harnessPath 'CheckPendingTrackingQueue(browser, options, result);' 'UI harness validates offline tracking queue visibility'
	Assert-Contains $harnessPath 'project element change history action is missing' 'UI harness validates the element history action'
	Assert-Contains $harnessPath 'TrackingPendingCount' 'UI harness passes offline tracking queue count into audit rendering'
	Assert-Contains $harnessPath 'compareDetailedSystemTypeComponents' 'UI harness passes detailed System Type comparison scenario setting'
	Assert-Contains $harnessPath 'Railing detailed component tables remain visible while comparison is disabled' 'UI harness verifies disabled component details are hidden'
	Assert-Contains $harnessPath 'CheckNestedFamilyDifferenceState(browser, options, result);' 'UI harness nested-family difference validation call'
	Assert-Contains $harnessPath 'matching nested child should stay hidden' 'UI harness matching nested-child exclusion check'
	Assert-Contains $harnessPath 'differing nested child remains directly selectable' 'UI harness nested-child direct-load selection guard'
	Assert-Contains $harnessPath 'nested parent detail table lost propagated child difference' 'UI harness parent detail propagation check'
	Assert-Contains $harnessPath 'missing nested child was mislabeled as a fingerprint capture failure' 'UI harness distinguishes missing nested content from capture failure'
    Assert-Contains $harnessPath 'CheckBrowserDetailContent(browser, options, result);' 'UI harness detail content validation call'
    Assert-Contains $harnessPath 'CheckDebugDock(browser, result);' 'UI harness debug dock validation call'
    Assert-Contains $harnessPath 'CheckAutoDetachedDetailAction(options, result);' 'UI harness auto detached detail action validation call'
    Assert-Contains $harnessPath 'CheckSearchFilteringKeepsFocus(browser, options, result);' 'UI harness search focus validation call'
    Assert-Contains $harnessPath 'CheckDetailedFilterResetAcrossBrowserTabs(browser, options, result);' 'UI harness Family/System detailed-filter reset validation call'
	Assert-Contains $harnessPath 'Detailed filter was not reset across Family/System tab switch' 'UI harness reports cross-tab detailed-filter state leakage'
	Assert-Contains $harnessPath 'duplicate search summary status filter row is present' 'UI harness rejects the removed duplicate status row'
	Assert-Contains $harnessPath 'permission diagnostic path cell is too narrow' 'UI harness verifies the full-width Permission/Guard path layout'
    Assert-Contains $harnessPath 'Auto detached detail action missing' 'UI harness auto detached detail failure message'
	Assert-Contains $harnessPath 'search focus not retained' 'UI harness search focus retained failure message'
	Assert-Contains $harnessPath "currentFilter='All';" 'UI harness isolates search verification from prior status filters'
	Assert-Contains $harnessPath 'store&&store.rows&&store.rows.length' 'UI harness chooses its search term from the original virtual-row store'
	Assert-Contains $harnessPath 'search.readOnly=true;' 'UI harness shields the focused search test from external keyboard input'
	Assert-Contains $harnessPath 'Search filtering emitted focus-stealing host actions' 'UI harness search host-action suppression failure message'
	Assert-Contains $harnessPath 'window.KKYFB._stores[window.currentTab]' 'UI harness search term comes from the stable virtual-row store'
    Assert-Contains $harnessPath 'validation = EvalScript(browser, validationScript);' 'UI harness search validation waits for queued row rendering'
    Assert-Contains $harnessPath 'floating debug FAB still rendered' 'UI harness debug FAB absence validation'
    Assert-Contains $harnessPath 'debug dock is not attached to viewport bottom' 'UI harness debug bottom dock validation'
    Assert-Contains $harnessPath 'debug dock overlaps left menu' 'UI harness debug dock left-nav overlap validation'
    Assert-Contains $harnessPath 'ApplyInlinePreviewAction(browser, previewAction, result);' 'UI harness preview-inline host simulation'
    Assert-Contains $harnessPath 'CheckProjectSubtitle(browser, options, result);' 'UI harness project subtitle and language-transition validation call'
    Assert-Contains $harnessPath 'CheckLanguagePurity(browser, options, result);' 'UI harness language purity validation call'
    Assert-Contains $harnessPath 'CheckStandardSetupEmptyState(browser, options, result);' 'UI harness standard setup empty-state validation call'
	Assert-Contains $harnessPath 'CheckPendingCommitState(browser, options, result);' 'UI harness pending save/sync validation call'
	Assert-Contains $harnessPath 'Pending save/sync row is still selectable before commit.' 'UI harness pending row selection guard'
	Assert-Contains $harnessPath 'Pending save/sync detail card is missing its warning heading.' 'UI harness pending detail warning validation'
    Assert-Contains $harnessPath 'Registered RVT / missing list state incorrectly shows the standard-RVT registration CTA' 'UI harness registered-RVT missing-list branch isolation'
    Assert-Contains $harnessPath 'Missing RVT state incorrectly shows the standard-list registration CTA' 'UI harness missing-RVT branch isolation'
    Assert-Contains $harnessPath 'English mode contains Korean visible text' 'UI harness English-mode Korean leakage failure'
    Assert-Contains $harnessPath 'Korean mode contains untranslated English UI text' 'UI harness Korean-mode English leakage failure'
    Assert-Contains $harnessPath 'KoreanModeDisallowedEnglishPhrases' 'UI harness Korean-mode disallowed English phrase list'
    Assert-Contains $harnessPath 'central title path missing' 'UI harness central path title validation'
    Assert-Contains $harnessPath 'detailParameters' 'UI harness parameter detail validation'
    Assert-Contains $harnessPath 'parameter detail expected exactly one unified table' 'UI harness rejects split or duplicate parameter tables'
    Assert-Contains $harnessPath 'type switch did not update the single unified Width row' 'UI harness validates type dropdown updates the unified table'
	Assert-Contains $harnessPath 'Yes/No parameter expected Yes for the first type' 'UI harness validates Yes display value'
	Assert-Contains $harnessPath 'Yes/No parameter expected No after type switch' 'UI harness validates No display value after type switch'
    Assert-Contains $harnessPath 'type switch removed the common instance row' 'UI harness keeps instance parameters visible while switching type'
    Assert-Contains $harnessPath 'lookup CSV summary lost original table-name casing or size' 'UI harness lookup CSV detail casing validation'
    Assert-Contains $harnessPath 'hasLookupCsv' 'UI harness lookup CSV diff-modal validation'
    Assert-Contains $harnessPath 'CheckBrowserSystemDetailContent(browser, options, result);' 'UI harness system detail content validation call'
    Assert-Contains $harnessPath 'system routing preference table expected exactly one' 'UI harness unified routing table validation'
    Assert-Contains $harnessPath 'system routing preference rule row count expected 2' 'UI harness routing rule dedupe validation'
	Assert-Contains $harnessPath 'system routing/layer unit selectors default is not mm' 'UI harness routing and layer default mm validation'
	Assert-Contains $harnessPath 'system routing criteria row count expected 2' 'UI harness routing criteria table validation'
	Assert-Contains $harnessPath 'system routing minimum cell missing 100 mm' 'UI harness validates the bounded minimum cell independently'
	Assert-Contains $harnessPath 'system routing maximum cell missing 300 mm' 'UI harness validates the bounded maximum cell independently'
	Assert-Contains $harnessPath 'system routing unbounded minimum/maximum cells are incomplete' 'UI harness validates both unbounded criterion cells'
	Assert-Contains $harnessPath 'system routing minimum cell inch conversion missing' 'UI harness validates minimum unit conversion independently'
	Assert-Contains $harnessPath 'system routing mm conversion missing' 'UI harness feet-to-mm conversion validation'
	Assert-Contains $harnessPath 'system routing sentinel not rendered as no limit' 'UI harness unbounded sentinel presentation validation'
	Assert-Contains $harnessPath 'system routing scientific sentinel leaked into visible text' 'UI harness raw scientific notation rejection'
	Assert-Contains $harnessPath 'system routing inch conversion missing' 'UI harness feet-to-inch conversion validation'
	Assert-Contains $harnessPath 'system routing bilingual elbow label missing' 'UI harness Korean bilingual / English-only routing label validation'
    Assert-Contains $harnessPath 'system detail legacy split sections remain visible' 'UI harness legacy split-section rejection'
    Assert-Contains $harnessPath 'system detail bottom review block is still visible' 'UI harness legacy bottom review rejection'
    Assert-Contains $harnessPath 'system detail Segment/Material dash row remains visible' 'UI harness misleading identity-row rejection'
    Assert-Contains $harnessPath 'nested child duplicate with dash category was not removed' 'UI harness nested child dash-duplicate validation'
    Assert-Contains $harnessPath 'nested child table row count expected 2 after dedupe' 'UI harness nested child dedupe row-count validation'
    Assert-Contains $harnessPath 'load available detail should not show fingerprint diff detail button' 'UI harness load-available no-diff detail validation'
    Assert-Contains $harnessPath 'load available detail should not show fingerprint diff summary' 'UI harness load-available no concise diff validation'
    Assert-Contains $harnessPath 'fingerprint diff audit row missing' 'UI harness separate diff-row validation'
    Assert-Contains $harnessPath 'fingerprint diff detail button did not open modal' 'UI harness fingerprint diff modal validation'
    Assert-Contains $harnessPath 'fingerprint-diff-table' 'UI harness fingerprint diff table validation'
    Assert-Contains $harnessPath 'preview-fit-image' 'UI harness preview image validation'
    Assert-Contains $harnessPath 'filter still uses ellipsis' 'UI harness status/trade filter no-ellipsis layout check'
    Assert-Contains $harnessPath 'filter content is clipped' 'UI harness status/trade filter clipped-text layout check'
    Assert-Contains $harnessPath 'inline status filter still uses ellipsis' 'UI harness inline action status filter no-ellipsis layout check'
    Assert-Contains $harnessPath 'browser grid overlaps action status filters' 'UI harness action row dynamic height layout check'
    Assert-Contains $harnessPath "checkTwoColumnActionCards('admin-standard-action-grid','admin standard action grid',false)" 'UI harness admin standard two-column layout check'
    Assert-Contains $harnessPath "checkTwoColumnActionCards('audit-action-grid','audit action grid',true)" 'UI harness model-check 2x2 layout check'
    Assert-Contains $harnessPath 'missing audit target selector' 'UI harness model-check target selector presence check'
    Assert-Contains $harnessPath 'audit target selector active chip count' 'UI harness model-check target active chip check'
    Assert-Contains $harnessPath 'admin trade selector active target count' 'UI harness admin current target active-state check'
    Assert-Contains $harnessPath 'standard action row has unequal button widths' 'UI harness admin standard button width consistency check'
    Assert-Contains $harnessPath 'trade management is still inside baseline RVT actions' 'UI harness admin trade management placement check'
	Assert-Contains $harnessPath 'admin policy checkbox-to-copy gap is too small' 'UI harness administrator policy option spacing check'
	Assert-Contains $harnessPath 'admin policy title and description are not vertically separated' 'UI harness administrator policy copy hierarchy check'
	Assert-Contains $harnessPath 'removed global project tracking option is still visible' 'UI harness rejects the removed global tracking option'
	Assert-Contains $harnessPath 'history sidebar group title count is' 'UI harness history sidebar group check'
	Assert-Contains $harnessPath 'an RVT not registered in Permissions / Guard entered element tracking scope' 'UI harness unregistered RVT tracking rejection'
    Assert-Contains $harnessPath 'CheckAuditTargetScrollPersistence(browser, result);' 'UI harness model-check target scroll round-trip check'
    Assert-Contains $harnessPath 'Audit target scroll state did not round-trip' 'UI harness model-check target scroll failure diagnostic'
    Assert-Contains $harnessPath "label+' grid is missing'" 'UI harness action-grid missing-class regression check'
    Assert-Contains $harnessPath 'button escapes card right edge' 'UI harness action-card button overflow check'
	Assert-Contains $harnessPath 'RunSyntheticCacheAudit(options, result);' 'UI harness cold/warm/offline cache audit'
	Assert-Contains $harnessPath 'CheckPerformance(browser, options, result);' 'UI harness 2,000-row performance check'
	Assert-Contains $harnessPath 'RenderMessageDialogHtml(options, result.HostAssembly)' 'UI harness structured message HTML render seam'
	Assert-Contains $harnessPath 'RunMessageBodyAudit(html, options, result);' 'UI harness structured message browser audit'
	Assert-Contains $harnessPath 'Structured message is missing semantic element #' 'UI harness structured message semantic-region check'
	Assert-Contains $harnessPath 'English structured message contains Hangul text.' 'UI harness structured message English-language purity check'
	Assert-Contains $harnessPath 'Structured message body overflows horizontally' 'UI harness structured message overflow check'
	Assert-Contains $harnessPath 'Automatic result message did not expose data-message-auto-result=true.' 'UI harness automatic result semantic validation'
	Assert-Contains $harnessPath 'Automatic result message lost context, metric, or output-path content.' 'UI harness automatic result content validation'
	Assert-Contains $harnessPath 'Virtual row window exceeds 150 DOM rows' 'UI harness true DOM row-window acceptance check'
	Assert-Contains $harnessPath 'CheckVirtualRowPagingAndSelection' 'UI harness cross-page paging and selection check'
	Assert-Contains $harnessPath 'resizeColumnForAudit(tab,resizeIndex,resizeTarget)' 'UI harness synchronized header/body column resize check'
	Assert-Contains $harnessPath 'untouchedStable' 'UI harness untouched column pixel-width stability check'
	Assert-Contains $harnessPath 'Column resize changed an untouched column' 'UI harness isolated column resize failure diagnostic'
	Assert-Contains $harnessPath 'rowWindowPageSummary' 'UI harness numbered pager summary check'
	Assert-Contains $harnessPath 'All rows still exist in the DOM' 'UI harness rejects hide-only row windowing'
	Assert-Contains $harnessPath 'Search/filter response exceeded' 'UI harness filter response acceptance check'
	Assert-Contains $harnessPath 'RunFileGuardPolicyAudit(result.HostAssembly, result);' 'UI harness executes file guard policy regression audit'
	Assert-Contains $harnessPath 'cloning the managed policy dropped the per-file discipline assignment' 'UI harness verifies policy cloning preserves the assigned trade'
	Assert-Contains $harnessPath 'RunFileGuardExcelRoundTripAudit(assembly, policy, guard, result);' 'UI harness round-trips per-RVT trade assignments through Korean XLSX'
	Assert-Contains $harnessPath 'RunManagedFolderSetupAudit(result.HostAssembly, result);' 'UI harness executes management-folder setup regression audit'
	Assert-Contains $harnessPath 'CheckManagedFolderTransition(browser, options, result);' 'UI harness validates TEST-to-homepage transition actions'
	Assert-Contains $harnessPath 'a different homepage policy was not protected from overwrite' 'UI harness validates managed-folder conflict protection'
	Assert-Contains $harnessPath 'JSON path rebasing' 'UI harness validates migrated JSON path rebasing'
	Assert-Contains $harnessPath 'migratedProjectCatalogPath' 'UI harness verifies project catalog migration'
	Assert-Contains $harnessPath 'migratedRevisionManifestPath' 'UI harness verifies Standard RVT revision baseline migration'
	Assert-Contains $harnessPath 'RunFamilyEditDialogGuardAudit(result.HostAssembly, result);' 'UI harness executes family edit dialog button-topology audit'
	Assert-Contains $harnessPath 'RunMeasurementUnitPreferenceAudit(result.HostAssembly, result);' 'UI harness executes persisted measurement-unit audit'
	Assert-Contains $harnessPath 'system-legacy-preview-visible' 'UI harness rejects detached system legacy preview duplication'
	Assert-Contains $harnessPath 'system-family-composition-visible' 'UI harness rejects family composition in system detail'
	Assert-Contains $harnessPath 'Measurement unit preference did not default to mm.' 'UI harness verifies missing preference defaults to mm'
	Assert-Contains $harnessPath 'Measurement unit preference did not restore inch.' 'UI harness verifies inch preference survives reload'
	Assert-Contains $harnessPath 'system layer composition table expected exactly one' 'UI harness validates Revit-style layer table'
	Assert-Contains $harnessPath 'system routing/layer unit selectors expected exactly two' 'UI harness validates synchronized routing and layer selectors'
	Assert-Contains $harnessPath 'system layer inch conversion missing' 'UI harness validates layer inch conversion'
	Assert-Contains $harnessPath 'Railing detailed component/configuration difference tables expected 2' 'UI harness validates Railing component and difference tables'
	Assert-Contains $harnessPath 'Railing detailed component references are incomplete' 'UI harness validates Railing dependent component references'
	Assert-Contains $harnessPath 'Railing component inch conversion is incomplete' 'UI harness validates Railing component unit conversion'
	Assert-Contains $harnessPath 'curtain panel inch conversion is incomplete' 'UI harness validates curtain wall panel unit conversion'
	Assert-Contains $harnessPath 'direct PanelType inch conversion is incomplete' 'UI harness validates direct Curtain Panel unit conversion'
	Assert-Contains $harnessPath 'Opening not cutting anything.' 'UI harness covers Opening not cutting anything warning'
	Assert-Contains $harnessPath 'Delete Instance or Delete Type to continue.' 'UI harness covers delete-only Cancel choice'
	Assert-Contains $harnessPath 'idok:Delete Instance' 'UI harness rejects destructive button carrying IDOK control id'
	Assert-Contains $harnessPath 'ScanDialogs worksheet is missing.' 'UI harness validates scan-dialog XLSX worksheet'
	Assert-Contains $harnessPath 'Managed folder audit: internal UNC acceptance or local-path rejection failed.' 'UI harness validates network-only management-folder selection'
	Assert-Contains $harnessPath 'CheckOverflowTitleBehavior(browser, result);' 'UI harness executes clipped-text title behavior audit'
	Assert-Contains $harnessPath 'clipped text did not receive generated title' 'UI harness validates generated title for clipped text'
	Assert-Contains $harnessPath 'widened text kept stale generated title' 'UI harness validates title removal after widening'
	Assert-Contains $harnessPath 'authored title was overwritten' 'UI harness preserves authored tooltips'
	Assert-Contains $harnessPath 'family load/edit was allowed for a matching central RVT while Admin Mode was off' 'UI harness checks Admin OFF family guard'
	Assert-Contains $harnessPath 'type changes were allowed for a matching central RVT while Admin Mode was off' 'UI harness checks Admin OFF type guard'
	Assert-Contains $harnessPath 'a non-target RVT was blocked' 'UI harness checks file guard target isolation'
	Assert-Contains $harnessPath 'RunNestedOnlyPlacementCatalogAudit(assembly, result);' 'UI harness executes nested-only placement classification audit'
	Assert-Contains $harnessPath 'legacy snapshots without standalone-usage metadata were not kept fail-open' 'UI harness rejects unsafe legacy snapshot classification'
	Assert-Contains $harnessPath 'nested-only standalone placement did not follow Admin OFF/ON state' 'UI harness checks nested-only policy Admin bypass'
	Assert-Contains $harnessPath 'unresolved hostname/IP strings were treated as the same physical RVT without file identity evidence' 'UI harness rejects unproven hostname/IP equivalence'
	Assert-Contains $harnessPath 'two paths backed by the same physical file identity did not match' 'UI harness accepts aliases only with physical file evidence'
	Assert-Contains $harnessPath 'different UNC shares were treated as the same RVT path' 'UI harness keeps different shares isolated'
	Assert-Contains $harnessPath 'a matching workshared local was allowed before the deferred central-path refresh' 'UI harness checks File Guard before manual Refresh'
	Assert-Contains $harnessPath 'an unrelated standalone RVT inherited a guard from a same-name target' 'UI harness rejects unrelated same-name standalone files'
	Assert-Contains $harnessPath 'ambiguous workshared identity did not preserve the strictest combined guard while identity was unavailable' 'UI harness requires strict conservative guard during transient workshared identity uncertainty'
	Assert-Contains $harnessPath 'duplicate physical targets did not preserve the most restrictive guard' 'UI harness verifies conservative duplicate target handling'
	Assert-Contains $harnessPath 'nested-only placement was not limited to an exact standard fingerprint match' 'UI harness keeps pending fingerprint checks non-blocking'
} else {
    Add-Failure "Missing UI audit harness: $harnessPath"
}

$harnessRunnerPath = Join-Path $root 'KKY_FamilyBrowser_Compile\Invoke-FamilyBrowserUiAuditHarness.ps1'
if (Test-Path -LiteralPath $harnessRunnerPath) {
	Assert-Contains $harnessRunnerPath "'--projectCatalogBaselineMissing'" 'UI harness runner forwards missing project catalog baseline state'
	Assert-Contains $harnessRunnerPath "'--projectCatalogChanged'" 'UI harness runner forwards changed project catalog state'
	Assert-Contains $harnessRunnerPath "'--projectCatalogUntracked'" 'UI harness runner forwards untracked project catalog state'
	Assert-Contains $harnessRunnerPath "'--trackingPendingCount'" 'UI harness runner forwards offline tracking queue count'
} else {
	Add-Failure "Missing UI audit harness runner: $harnessRunnerPath"
}

$ribbonIconLoaderPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserRibbonIcon.cs'
$ribbonIconGeneratorPath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\New-FamilyBrowserRibbonIcons.ps1'
$ribbonIconSpecs = @(
	@{ Path = (Join-Path $root 'KKY_FamilyBrowser_SharedUi\family-browser-ribbon-16.png'); Size = 16 },
	@{ Path = (Join-Path $root 'KKY_FamilyBrowser_SharedUi\family-browser-ribbon-32.png'); Size = 32 }
)
$ribbonHostSpecs = @(
	@{
		Name = '2019-2023'
		App = (Join-Path $root 'KKY_FamilyBrowser_RevitHost_2019-2023\KKY_FamilyBrowser_RevitHost_2019_2023\App.cs')
		Project = (Join-Path $root 'KKY_FamilyBrowser_RevitHost_2019-2023\KKY_FamilyBrowser_RevitHost.csproj')
	},
	@{
		Name = '2025'
		App = (Join-Path $root 'KKY_FamilyBrowser_RevitHost_2025\KKY_FamilyBrowser_RevitHost_2025\App.cs')
		Project = (Join-Path $root 'KKY_FamilyBrowser_RevitHost_2025\KKY_FamilyBrowser_RevitHost_2025.csproj')
	},
	@{
		Name = '2027'
		App = (Join-Path $root 'KKY_FamilyBrowser_RevitHost_2027\KKY_FamilyBrowser_RevitHost_2027\App.cs')
		Project = (Join-Path $root 'KKY_FamilyBrowser_RevitHost_2027\KKY_FamilyBrowser_RevitHost_2027.csproj')
	}
)

if (Test-Path -LiteralPath $ribbonIconLoaderPath) {
	Assert-Contains $ribbonIconLoaderPath 'KKY.FamilyBrowser.RibbonAssets.family-browser-ribbon-16.png' 'Family Browser small ribbon icon resource loader'
	Assert-Contains $ribbonIconLoaderPath 'KKY.FamilyBrowser.RibbonAssets.family-browser-ribbon-32.png' 'Family Browser large ribbon icon resource loader'
	Assert-Contains $ribbonIconLoaderPath 'BitmapCacheOption.OnLoad' 'Family Browser ribbon icon stream-independent bitmap loading'
	Assert-Contains $ribbonIconLoaderPath 'image.Freeze();' 'Family Browser ribbon icon frozen ImageSource'
} else {
	Add-Failure "Missing Family Browser ribbon icon loader: $ribbonIconLoaderPath"
}

if (-not (Test-Path -LiteralPath $ribbonIconGeneratorPath)) {
	Add-Failure "Missing deterministic Family Browser ribbon icon generator: $ribbonIconGeneratorPath"
}

foreach ($iconSpec in $ribbonIconSpecs) {
	if (-not (Test-Path -LiteralPath $iconSpec.Path)) {
		Add-Failure "Missing Family Browser ribbon icon: $($iconSpec.Path)"
		continue
	}
	$bytes = [System.IO.File]::ReadAllBytes($iconSpec.Path)
	if ($bytes.Length -lt 24 -or $bytes[0] -ne 137 -or $bytes[1] -ne 80 -or $bytes[2] -ne 78 -or $bytes[3] -ne 71) {
		Add-Failure "Invalid Family Browser ribbon PNG: $($iconSpec.Path)"
		continue
	}
	$width = [System.Net.IPAddress]::NetworkToHostOrder([System.BitConverter]::ToInt32($bytes, 16))
	$height = [System.Net.IPAddress]::NetworkToHostOrder([System.BitConverter]::ToInt32($bytes, 20))
	if ($width -ne $iconSpec.Size -or $height -ne $iconSpec.Size) {
		Add-Failure "Family Browser ribbon icon expected $($iconSpec.Size)x$($iconSpec.Size), found ${width}x${height}: $($iconSpec.Path)"
	}
}

foreach ($hostSpec in $ribbonHostSpecs) {
	if (-not (Test-Path -LiteralPath $hostSpec.App)) {
		Add-Failure "Missing Family Browser ribbon App source: $($hostSpec.App)"
		continue
	}
	if (-not (Test-Path -LiteralPath $hostSpec.Project)) {
		Add-Failure "Missing Family Browser ribbon host project: $($hostSpec.Project)"
		continue
	}
	Assert-Contains $hostSpec.App 'private const string TabName = "KKY Tools";' "$($hostSpec.Name) shared KKY Tools ribbon tab"
	Assert-NotContains $hostSpec.App '"KKY Browser"' "$($hostSpec.Name) retired standalone KKY Browser tab"
	Assert-Contains $hostSpec.App 'application.CreateRibbonTab(TabName);' "$($hostSpec.Name) constant-based ribbon tab creation"
	Assert-Contains $hostSpec.App 'application.GetRibbonPanels(TabName)' "$($hostSpec.Name) shared-tab panel lookup"
	Assert-Contains $hostSpec.App 'application.CreateRibbonPanel(TabName, PanelName);' "$($hostSpec.Name) shared-tab Family Browser panel creation"
	Assert-Contains $hostSpec.App 'button.ToolTip = "KKY Family Browser 열기";' "$($hostSpec.Name) Family Browser ribbon tooltip"
	Assert-Contains $hostSpec.App 'FamilyBrowserRibbonIcon.LoadSmall();' "$($hostSpec.Name) small ribbon icon binding"
	Assert-Contains $hostSpec.App 'FamilyBrowserRibbonIcon.LoadLarge();' "$($hostSpec.Name) large ribbon icon binding"
	Assert-Contains $hostSpec.App 'button.LargeImage = largeIcon;' "$($hostSpec.Name) large ribbon image assignment"
	Assert-Contains $hostSpec.Project '<UseWPF>True</UseWPF>' "$($hostSpec.Name) WPF ImageSource project support"
	Assert-Contains $hostSpec.Project 'LogicalName="KKY.FamilyBrowser.RibbonAssets.family-browser-ribbon-16.png"' "$($hostSpec.Name) embedded 16px ribbon icon"
	Assert-Contains $hostSpec.Project 'LogicalName="KKY.FamilyBrowser.RibbonAssets.family-browser-ribbon-32.png"' "$($hostSpec.Name) embedded 32px ribbon icon"
}

$contractPath = Join-Path $root 'KKY_FamilyBrowser_Compile\FamilyBrowserUiAudit.contract.json'
if (Test-Path -LiteralPath $contractPath) {
    Assert-Contains $contractPath '"name": "admin-family-missing-standard-list"' 'Family missing-standard-list audit scenario'
	Assert-Contains $contractPath '"name": "admin-family-pending-save-sync"' 'Family pending save/sync audit scenario'
	Assert-Contains $contractPath '"name": "admin-system-pending-save-sync"' 'System pending save/sync audit scenario'
	Assert-Contains $contractPath '"name": "admin-system-components-disabled"' 'System detailed-component disabled audit scenario'
	Assert-Contains $contractPath '"name": "admin-home-standard-rvt-changed"' 'changed Standard RVT Home audit scenario'
	Assert-Contains $contractPath '"name": "modeler-family-standard-rvt-unavailable"' 'unavailable Standard RVT browser audit scenario'
	Assert-Contains $contractPath '"name": "modeler-manual-en"' 'modeler English manual audit scenario'
	Assert-Contains $contractPath '"name": "admin-permission-long-path-layout"' 'Permission/Guard long-path layout audit scenario'
	Assert-Contains $contractPath '"policyActiveDisciplineKey": "Mechanical"' 'browse trade differs from policy-active trade regression scenario'
	Assert-Contains $contractPath '"update-check"' 'update-check UI contract route'
	Assert-Contains $contractPath '"open-homepage"' 'homepage UI contract route'
	Assert-Contains $contractPath '"open-manual"' 'manual website UI contract route'
} else {
    Add-Failure "Missing UI audit contract: $contractPath"
}

$productUpdateServicePath = Join-Path $root 'KKY_FamilyBrowser_SharedUi\FamilyBrowserProductUpdateService.cs'
if (Test-Path -LiteralPath $productUpdateServicePath) {
	Assert-Contains $productUpdateServicePath 'public const string CurrentProductVersion = "1.0.1";' 'Family Browser product version source'
	Assert-Contains $productUpdateServicePath 'https://update.zerokky.com/Release/family-browser/latest.json' 'live Family Browser update feed URL'
	Assert-Contains $productUpdateServicePath 'https://update.zerokky.com/family-browser/index.html' 'live Family Browser manual URL'
	Assert-Contains $productUpdateServicePath 'Task.Factory.StartNew(' 'non-blocking Family Browser update check'
	Assert-Contains $productUpdateServicePath 'RequestCacheLevel.NoCacheNoStore' 'manual update check bypasses stale HTTP cache'
	Assert-Contains $productUpdateServicePath 'Math.Max(parsed.Build, 0)' 'update version comparison normalizes omitted build component'
	Assert-Contains $productUpdateServicePath 'Math.Max(parsed.Revision, 0)' 'update version comparison normalizes omitted revision component'
	Assert-Contains $productUpdateServicePath 'OpenFamilyBrowserSupportUri' 'safe homepage launcher'
	Assert-Contains $productUpdateServicePath '[DataMember(Name = "sha256")]' 'update manifest SHA-256 contract'
	Assert-Contains $productUpdateServicePath 'DownloadUpdateInstaller' 'verified update installer downloader'
	Assert-Contains $productUpdateServicePath 'TryGetTrustedInstallerUri' 'same-host HTTPS installer URL gate'
	Assert-Contains $productUpdateServicePath 'response.ResponseUri' 'installer redirect destination validation'
	Assert-Contains $productUpdateServicePath 'MaximumInstallerBytes' 'installer download size limit'
	Assert-Contains $productUpdateServicePath 'ValidateInstallerFile' 'downloaded installer integrity gate'
	Assert-Contains $productUpdateServicePath 'ComputeFileSha256' 'downloaded installer SHA-256 calculation'
	Assert-Contains $productUpdateServicePath 'FixedTimeEquals' 'constant-time installer hash comparison'
	Assert-Contains $productUpdateServicePath 'ValidateDownloadedInstallerBeforeLaunch' 'installer revalidation immediately before execution'
	Assert-Contains $productUpdateServicePath 'Verb = "runas"' 'verified installer elevated launch'
	Assert-Contains $productUpdateServicePath '_productUpdateDownloadPending' 'non-blocking update download state gate'
	Assert-Contains $productUpdateServicePath 'Save or synchronize every open project before starting the installer.' 'Revit save and close update guidance'
} else {
	Add-Failure "Missing Family Browser product update service: $productUpdateServicePath"
}

$harnessWrapperPath = Join-Path $root 'KKY_FamilyBrowser_Compile\Invoke-FamilyBrowserUiAuditHarness.ps1'
if (Test-Path -LiteralPath $harnessWrapperPath) {
	Assert-Contains $harnessWrapperPath "'--compareDetailedSystemTypeComponents'" 'UI harness wrapper forwards detailed System Type comparison setting'
    Assert-Contains $harnessWrapperPath 'Get-ScenarioLanguageVariants' 'UI harness wrapper runs scenario language variants'
    Assert-Contains $harnessWrapperPath 'Get-ScenarioThemeVariants' 'UI harness wrapper runs light and dark theme variants'
    Assert-Contains $harnessWrapperPath '$scenarioName = "$baseScenarioName-$languageCode-$themeCode"' 'UI harness wrapper suffixes scenario names by language and theme'
    Assert-Contains $harnessWrapperPath "'ko', 'en'" 'UI harness wrapper includes Korean and English scenarios'
    Assert-Contains $harnessWrapperPath '"structured-message-$languageCode-$themeCode"' 'UI harness wrapper runs structured message scenarios'
    Assert-Contains $harnessWrapperPath '"auto-result-message-$languageCode-$themeCode"' 'UI harness wrapper runs automatic result scenarios'
    Assert-Contains $harnessWrapperPath "'--messageFixture'," 'UI harness wrapper selects automatic result fixture'
    Assert-Contains $harnessWrapperPath "'--renderMode', 'message'" 'UI harness wrapper selects structured message render mode'
    Assert-Contains $harnessWrapperPath "'--themeCode'," 'UI harness wrapper passes the requested theme to the renderer'
	Assert-Contains $harnessWrapperPath "'--includePendingRows'," 'UI harness wrapper passes pending-state scenarios'
	Assert-Contains $harnessWrapperPath "'--policyActiveDisciplineKey'," 'UI harness wrapper passes the policy-active trade regression value'
	Assert-Contains $harnessWrapperPath "'admin-family-detail-with-preview', 'admin-system-with-data'" 'UI harness wrapper captures both family and system detail visuals'
	Assert-Contains $harnessWrapperPath "@('--revisionAudit', 'true')" 'UI harness wrapper runs Standard RVT revision primitives once per host'
} else {
    Add-Failure "Missing UI audit harness wrapper: $harnessWrapperPath"
}

$buildStagePath = Join-Path $root 'KKY_FamilyBrowser_Compile\Build-FamilyBrowserRecovered.ps1'
$buildInstallerPath = Join-Path $root 'KKY_FamilyBrowser_Compile\Build-FamilyBrowserInstaller.ps1'
$distributionAuditPath = Join-Path $root 'KKY_FamilyBrowser_Compile\Test-FamilyBrowserDistribution.ps1'
$installRecoveredPath = Join-Path $root 'KKY_FamilyBrowser_Compile\Install-FamilyBrowserRecovered.ps1'
$verifyRecoveredPath = Join-Path $root 'KKY_FamilyBrowser_Compile\Verify-FamilyBrowserRecovered.ps1'
$installerDefinitionPaths = @(
    (Join-Path $root 'KKY_FamilyBrowser_Compile\KKY_FamilyBrowser_Compiler.iss'),
    (Join-Path $root 'KKY_FamilyBrowser_Compile\KKY_FamilyBrowser_Compiler_NoSetupLdr.iss')
)

if (Test-Path -LiteralPath $buildStagePath) {
    Assert-Contains $buildStagePath 'payloadFileCount = $payloadFiles.Count' 'Stage manifest payload count'
    Assert-Contains $buildStagePath "sha256 = (Get-FileHash -LiteralPath `$_.FullName -Algorithm SHA256).Hash" 'Stage manifest payload hash'
} else {
    Add-Failure "Missing Family Browser stage builder: $buildStagePath"
}

if (Test-Path -LiteralPath $verifyRecoveredPath) {
    Assert-Contains $verifyRecoveredPath 'Assert-StageManifestIntegrity -Root $StageRoot' 'Stage payload-manifest integrity gate'
    Assert-Contains $verifyRecoveredPath '$StageRoot = [System.IO.Path]::GetFullPath($StageRoot)' 'Stage verifier normalizes relative Stage roots'
    Assert-Contains $verifyRecoveredPath '$manifestPathFull' 'Stage verifier compares the manifest by normalized full path'
    Assert-Contains $verifyRecoveredPath 'installed payload hash mismatch' 'ProgramData full payload hash gate'
    Assert-Contains $verifyRecoveredPath 'obsolete installed payload remains' 'ProgramData obsolete payload gate'
    Assert-Contains $verifyRecoveredPath 'obsolete addin manifest remains' 'ProgramData obsolete manifest gate'
} else {
    Add-Failure "Missing Family Browser payload verifier: $verifyRecoveredPath"
}

if (Test-Path -LiteralPath $installRecoveredPath) {
    Assert-Contains $installRecoveredPath 'Close every Revit process before installing Family Browser.' 'ProgramData install blocks running Revit'
    Assert-Contains $installRecoveredPath 'Family Browser ProgramData installation requires administrator rights or explicit write access' 'ProgramData install fails before mutation when neither elevation nor exact-root write access is available'
    Assert-Contains $installRecoveredPath '$probeFolders = @($installRoot)' 'ProgramData non-elevated replacement probes the exact addin root and existing payload before mutation'
    Assert-Contains $installRecoveredPath '.kky-family-browser-install-probe-' 'ProgramData non-elevated install verifies explicit create and delete access before mutation'
    Assert-Contains $installRecoveredPath '[System.IO.FileAccess]::ReadWrite' 'ProgramData non-elevated replacement verifies existing manifest write access before mutation'
    Assert-Contains $installRecoveredPath 'Assert-CopiedPayload -SourceFolder $stagedFolder -CopiedFolder $temporaryFolder' 'ProgramData temporary payload hash verification'
    Assert-Contains $installRecoveredPath "Remove-Item -LiteralPath `$destinationFolder -Recurse -Force" 'ProgramData obsolete folder cleanup'
    Assert-Contains $installRecoveredPath "'KKY_FamilyBrowser_RevitHost_2027.addin'" 'ProgramData obsolete manifest cleanup'
} else {
    Add-Failure "Missing Family Browser ProgramData installer: $installRecoveredPath"
}

if (Test-Path -LiteralPath $buildInstallerPath) {
    Assert-Contains $buildInstallerPath "[string]`$Version = '1.0.1'" 'Installer default version 1.0.1'
    Assert-Contains $buildInstallerPath '[double]$MailPackageMinimumMB = 15.9' 'Mail package default size threshold'
    Assert-Contains $buildInstallerPath 'Assert-PortableExecutable -Path $outputExe' 'Installer PE validation'
    Assert-Contains $buildInstallerPath 'Get-MailPackageValidation -ZipPath $mailPackage' 'Mail package embedded Setup hash validation'
    Assert-Contains $buildInstallerPath "Join-Path `$outputDir 'latest-build.json'" 'Installer latest-build metadata'
    Assert-Contains $buildInstallerPath "Join-Path `$mailPackageRoot 'latest-mail-package.json'" 'Mail package latest metadata'
    Assert-Contains $buildInstallerPath 'Remove-Item -LiteralPath $workDir -Recurse -Force' 'Mail package temporary folder cleanup'
	Assert-Contains $buildInstallerPath "Join-Path `$scriptRoot 'Test-FamilyBrowserDistribution.ps1'" 'Installer invokes independent distribution audit'
	Assert-Contains $buildInstallerPath 'Family Browser source/package version mismatch.' 'Installer build blocks source/package version mismatch'
	Assert-Contains $buildInstallerPath 'CurrentProductVersion' 'Installer build reads the runtime product version'
	Assert-NotContains $buildInstallerPath 'Distribution audit failed with exit code $LASTEXITCODE' 'PowerShell distribution audit must use exception-based failure handling'
} else {
    Add-Failure "Missing Family Browser installer builder: $buildInstallerPath"
}

if (Test-Path -LiteralPath $distributionAuditPath) {
    Assert-Contains $distributionAuditPath 'Installer is older than the current Stage manifest.' 'Distribution audit rejects stale installer'
    Assert-Contains $distributionAuditPath 'Mail ZIP Setup.exe does not match the latest standalone installer.' 'Distribution audit verifies embedded Setup hash'
    Assert-Contains $distributionAuditPath 'No abandoned mail packaging work folders' 'Distribution audit rejects abandoned package work folders'
    Assert-Contains $distributionAuditPath "-Installed -StageRoot `$StageRoot" 'Distribution audit can verify ProgramData payload'
    Assert-Contains $distributionAuditPath '$context = [ordered]@{' 'Distribution audit shares loaded metadata through an explicit context object'
    Assert-NotContains $distributionAuditPath 'Stage verification exited with code $LASTEXITCODE' 'PowerShell Stage verifier must use exception-based failure handling'
    Assert-NotContains $distributionAuditPath 'Installed verification exited with code $LASTEXITCODE' 'PowerShell installed verifier must use exception-based failure handling'
} else {
    Add-Failure "Missing Family Browser distribution audit: $distributionAuditPath"
}

foreach ($installerDefinitionPath in $installerDefinitionPaths) {
    if (-not (Test-Path -LiteralPath $installerDefinitionPath)) {
        Add-Failure "Missing Family Browser installer definition: $installerDefinitionPath"
        continue
    }
    Assert-Contains $installerDefinitionPath '#define MyAppVersion "1.0.1"' 'Installer definition default version 1.0.1'
    Assert-Contains $installerDefinitionPath '[InstallDelete]' 'Installer clean-up section'
	Assert-Contains $installerDefinitionPath 'CloseApplications=yes' 'Installer detects and closes locked Revit payloads through Restart Manager'
	Assert-Contains $installerDefinitionPath 'RestartApplications=no' 'Installer never restarts Revit automatically'
    Assert-MinimumOccurrences $installerDefinitionPath 'Type: filesandordirs; Name: "{commonappdata}\Autodesk\Revit\Addins\' 5 'Installer per-version obsolete payload cleanup'
}

$uiHarnessProgramPath = Join-Path $root 'KKY_FamilyBrowser_Automation\KKY_FamilyBrowser_UiAuditHarness\Program.cs'
if (Test-Path -LiteralPath $uiHarnessProgramPath) {
	Assert-Contains $uiHarnessProgramPath 'RunProductUpdatePrimitiveAudit(result.HostAssembly, result);' 'UI harness invokes offline updater security audit'
	Assert-Contains $uiHarnessProgramPath 'CheckBrowseRowDisciplineSynchronization(browser, result);' 'UI harness checks the visible row trade column against the selected trade'
	Assert-Contains $uiHarnessProgramPath 'Product update audit accepted an untrusted installer URL.' 'UI harness rejects untrusted updater URLs'
	Assert-Contains $uiHarnessProgramPath 'Product update audit accepted an installer with the wrong SHA-256.' 'UI harness rejects modified installer bytes'
	Assert-Contains $uiHarnessProgramPath 'Product update audit accepted a non-MZ file.' 'UI harness rejects non-installer content'
} else {
	Add-Failure "Missing UI audit harness program: $uiHarnessProgramPath"
}

if ($warnings.Count -gt 0) {
    Write-Host 'UI static warnings:'
    foreach ($warning in $warnings) {
        Write-Host "WARN $warning"
    }
}

if ($failures.Count -gt 0) {
    Write-Host 'UI static failures:'
    foreach ($failure in $failures) {
        Write-Host "FAIL $failure"
    }
    exit 1
}

Write-Host 'UI static checks passed.'

$contractScript = Join-Path $PSScriptRoot 'Test-FamilyBrowserUiContract.ps1'
if (Test-Path -LiteralPath $contractScript) {
    & $contractScript
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
