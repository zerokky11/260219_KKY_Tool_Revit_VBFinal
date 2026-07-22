param(
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $repoRoot ('artifacts\family-browser-workflow-audit\' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
}
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$checks = New-Object System.Collections.Generic.List[object]
$failures = New-Object System.Collections.Generic.List[string]

function Add-Check([string]$name, [bool]$passed, [string]$details = '', [string]$workflowId = '') {
    $checks.Add([pscustomobject]@{ Name = $name; Passed = $passed; Details = $details; WorkflowId = $workflowId }) | Out-Null
    if (-not $passed) {
        $failures.Add("${name}: $details") | Out-Null
    }
}

function Add-WorkflowCheck([string]$workflowId, [string]$name, [bool]$passed, [string]$details = '') {
    Add-Check $name $passed $details $workflowId
}

function Read-Text([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) {
        return ''
    }
    return Get-Content -Raw -Encoding UTF8 -LiteralPath $path
}

function Assert-Token([string]$path, [string]$token, [string]$name) {
    $text = Read-Text $path
    Add-Check $name ($text.IndexOf($token, [StringComparison]::Ordinal) -ge 0) "Missing '$token' in $path"
}

function Assert-WorkflowToken([string]$workflowId, [string]$path, [string]$token, [string]$name) {
    $text = Read-Text $path
    Add-WorkflowCheck $workflowId $name ($text.IndexOf($token, [StringComparison]::Ordinal) -ge 0) "Missing '$token' in $path"
}

function Assert-WorkflowMinimumOccurrences([string]$workflowId, [string]$path, [string]$token, [int]$minimum, [string]$name) {
    $text = Read-Text $path
    $count = if ([string]::IsNullOrEmpty($text)) { 0 } else { ([regex]::Matches($text, [regex]::Escape($token))).Count }
    Add-WorkflowCheck $workflowId $name ($count -ge $minimum) "Expected at least $minimum occurrence(s) of '$token' in $path; found $count."
}

function Assert-WorkflowMethodExactOccurrences([string]$workflowId, [string]$path, [string]$methodName, [string]$token, [int]$expected, [string]$name) {
    $body = Get-MethodText $path $methodName
    $count = if ([string]::IsNullOrEmpty($body)) { 0 } else { ([regex]::Matches($body, [regex]::Escape($token))).Count }
    Add-WorkflowCheck $workflowId $name ($count -eq $expected) "Expected exactly $expected occurrence(s) of '$token' in $methodName; found $count."
}

function Get-MethodText([string]$path, [string]$methodName, [string]$exactSignature = '') {
    $text = Read-Text $path
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }
    $pattern = if ([string]::IsNullOrWhiteSpace($exactSignature)) {
        '(?ms)^\s*(?:public|private|internal|protected)\s+(?:static\s+)?[^\r\n\{;=]*?\b' + [regex]::Escape($methodName) + '\s*\([^\{\};]*?\)\s*\{'
    } else {
        '(?m)^\s*' + [regex]::Escape($exactSignature) + '\s*\{'
    }
    $match = [regex]::Match($text, $pattern)
    if (-not $match.Success) {
        return ''
    }
    $open = $text.IndexOf('{', $match.Index)
    $depth = 0
    $inString = $false
    $escaped = $false
    for ($index = $open; $index -lt $text.Length; $index++) {
        $ch = $text[$index]
        if ($inString) {
            if ($escaped) {
                $escaped = $false
            } elseif ($ch -eq '\') {
                $escaped = $true
            } elseif ($ch -eq '"') {
                $inString = $false
            }
            continue
        }
        if ($ch -eq '"') {
            $inString = $true
            continue
        }
        if ($ch -eq '{') { $depth++ }
        if ($ch -eq '}') {
            $depth--
            if ($depth -eq 0) {
                return $text.Substring($match.Index, $index - $match.Index + 1)
            }
        }
    }
    return ''
}

function Assert-MethodToken([string]$path, [string]$methodName, [string]$token, [string]$name) {
    $body = Get-MethodText $path $methodName
    Add-Check $name (-not [string]::IsNullOrWhiteSpace($body) -and $body.IndexOf($token, [StringComparison]::Ordinal) -ge 0) "Method $methodName must contain '$token' in $path"
}

function Assert-WorkflowMethodToken([string]$workflowId, [string]$path, [string]$methodName, [string]$token, [string]$name, [string]$exactSignature = '') {
    $body = Get-MethodText $path $methodName $exactSignature
    Add-WorkflowCheck $workflowId $name (-not [string]::IsNullOrWhiteSpace($body) -and $body.IndexOf($token, [StringComparison]::Ordinal) -ge 0) "Method $methodName must contain '$token' in $path"
}

function Assert-WorkflowMethodExcludesToken([string]$workflowId, [string]$path, [string]$methodName, [string]$token, [string]$name) {
    $body = Get-MethodText $path $methodName
    Add-WorkflowCheck $workflowId $name (-not [string]::IsNullOrWhiteSpace($body) -and $body.IndexOf($token, [StringComparison]::Ordinal) -lt 0) "Method $methodName must not contain '$token' in $path"
}

function Assert-MethodOrder([string]$path, [string]$methodName, [string]$first, [string]$second, [string]$name) {
    $body = Get-MethodText $path $methodName
    $firstIndex = $body.IndexOf($first, [StringComparison]::Ordinal)
    $secondIndex = $body.IndexOf($second, [StringComparison]::Ordinal)
    Add-Check $name (-not [string]::IsNullOrWhiteSpace($body) -and $firstIndex -ge 0 -and $secondIndex -gt $firstIndex) "Expected '$first' before '$second' in $methodName"
}

function Assert-WorkflowMethodOrder([string]$workflowId, [string]$path, [string]$methodName, [string]$first, [string]$second, [string]$name) {
    $body = Get-MethodText $path $methodName
    $firstIndex = $body.IndexOf($first, [StringComparison]::Ordinal)
    $secondIndex = $body.IndexOf($second, [StringComparison]::Ordinal)
    Add-WorkflowCheck $workflowId $name (-not [string]::IsNullOrWhiteSpace($body) -and $firstIndex -ge 0 -and $secondIndex -gt $firstIndex) "Expected '$first' before '$second' in $methodName"
}

function Assert-WorkflowMethodTokenAfter([string]$workflowId, [string]$path, [string]$methodName, [string]$anchor, [string]$token, [string]$name) {
    $body = Get-MethodText $path $methodName
    $anchorIndex = $body.IndexOf($anchor, [StringComparison]::Ordinal)
    $tokenIndex = if ($anchorIndex -ge 0) { $body.IndexOf($token, $anchorIndex + $anchor.Length, [StringComparison]::Ordinal) } else { -1 }
    Add-WorkflowCheck $workflowId $name (-not [string]::IsNullOrWhiteSpace($body) -and $anchorIndex -ge 0 -and $tokenIndex -gt $anchorIndex) "Expected '$token' after '$anchor' in $methodName"
}

$contractPath = Join-Path $scriptRoot 'FamilyBrowserWorkflowAudit.contract.json'
$contract = Get-Content -Raw -Encoding UTF8 -LiteralPath $contractPath | ConvertFrom-Json
$workflowIds = @($contract.workflows | ForEach-Object { [string]$_.id })
$requiredWorkflowIds = @(
    'managed-folder-first-run',
    'managed-folder-homepage-return',
    'startup-cache-offline',
    'standard-setup-readiness',
    'standard-trade-switch',
    'standard-source-revision-block',
    'standard-edit-commit-lifecycle',
    'admin-live-state-refresh',
    'family-list-interaction',
    'family-detail-integrity',
    'nested-family-propagation',
    'family-load-save-lifecycle',
    'family-type-attribution',
    'system-type-apply-lifecycle',
    'system-detail-integrity',
    'current-model-check-baseline',
    'external-project-change-detection',
    'project-element-change-ledger',
    'element-tracking-integrity-recovery',
    'element-tracking-session-recovery',
    'element-tracking-schema-compatibility',
    'element-tracking-event-ambiguity',
    'tracking-policy-concurrency',
    'tracking-policy-session-isolation',
    'offline-tracking-recovery',
    'multi-client-tracking',
    'guard-admin-transition',
    'request-lifecycle',
    'request-concurrent-edit',
	'request-attachment-and-delete-audit',
    'scan-dialog-recovery',
    'language-result-export',
    'large-library-performance',
    'diagnostics-debug-log',
    'close-cancel-and-failure'
)
Add-Check 'workflow contract has unique IDs' (($workflowIds | Select-Object -Unique).Count -eq $workflowIds.Count) 'Duplicate workflow IDs found.'
foreach ($id in $requiredWorkflowIds) {
    Add-Check "workflow contract includes $id" ($workflowIds -contains $id) 'Required workflow is missing.'
}

$sharedRoot = Join-Path $repoRoot 'KKY_FamilyBrowser_SharedUi'
$trackingPath = Join-Path $sharedRoot 'FamilyBrowserTrackingPersistenceService.cs'
$catalogPath = Join-Path $sharedRoot 'FamilyBrowserProjectCatalogService.cs'
$revisionPath = Join-Path $sharedRoot 'FamilyBrowserStandardRevisionService.cs'
$revisionDashboardPath = Join-Path $sharedRoot 'FamilyBrowserDashboardStandardRevision.cs'
$projectCatalogPath = Join-Path $sharedRoot 'FamilyBrowserProjectCatalogService.cs'
$projectCatalogDashboardPath = Join-Path $sharedRoot 'FamilyBrowserDashboardProjectCatalog.cs'
$elementChangeTrackingPath = Join-Path $sharedRoot 'FamilyBrowserElementChangeTrackingService.cs'
$elementTrackingScopePolicyPath = Join-Path $sharedRoot 'FamilyBrowserElementTrackingScopePolicy.cs'
$elementTrackingTransitionPolicyPath = Join-Path $sharedRoot 'FamilyBrowserElementTrackingTransitionPolicy.cs'
$elementHistoryProjectionPolicyPath = Join-Path $sharedRoot 'FamilyBrowserElementHistoryProjectionPolicy.cs'
$trackingCommitOptimizationPath = Join-Path $sharedRoot 'FamilyBrowserTrackingCommitOptimizationPolicy.cs'
$elementActivityMatcherPath = Join-Path $sharedRoot 'FamilyBrowserElementActivityMatcher.cs'
$elementTrackingPolicyDecisionPath = Join-Path $sharedRoot 'FamilyBrowserElementTrackingPolicyDecision.cs'
$managedFolderSetupPath = Join-Path $sharedRoot 'FamilyBrowserManagedFolderSetupService.cs'
$managedFolderTransitionPath = Join-Path $sharedRoot 'FamilyBrowserDashboardManagedFolderTransition.cs'
$managementContextLockPath = Join-Path $sharedRoot 'FamilyBrowserManagementContextLock.cs'
$dataLoaderPath = Join-Path $sharedRoot 'FamilyBrowserDataLoader.cs'
$atomicFileServicePath = Join-Path $sharedRoot 'FamilyBrowserAtomicFileService.cs'
$uniqueJsonReportStorePath = Join-Path $sharedRoot 'FamilyBrowserUniqueJsonReportStore.cs'
$requestConcurrencyPath = Join-Path $sharedRoot 'FamilyBrowserRequestConcurrencyService.cs'
$requestFileTransactionPath = Join-Path $sharedRoot 'FamilyBrowserRequestFileTransactionService.cs'
$systemRoutingUnitPath = Join-Path $sharedRoot 'FamilyBrowserSystemRoutingUnitUi.cs'
$systemComponentUnitPath = Join-Path $sharedRoot 'FamilyBrowserSystemDetailedComponentUnitUi.cs'
$measurementPreferencePath = Join-Path $sharedRoot 'FamilyBrowserMeasurementUnitPreferenceService.cs'
$automaticModelCheckPath = Join-Path $sharedRoot 'FamilyBrowserAutomaticModelCheckService.cs'
$dashboardPath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\FamilyBrowserDashboardHtmlForm.cs'
$projectSnapshotStorePath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\ProjectSnapshotStore.cs'
$projectTrackingStorePath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\ProjectTrackingStoreService.cs'
$standardRegistryPath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\StandardLibraryRegistryStore.cs'
$standardRegistrationPath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\StandardLibraryRegistrationService.cs'
$loadableFamilySyncStorePath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\LoadableFamilySyncStore.cs'
$systemTypeApplyStorePath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\SystemTypeApplyStore.cs'
$systemTypePreflightStorePath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\SystemTypePreflightStore.cs'
$projectTrackingOutputStorePath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\ProjectTrackingOutputStore.cs'
$bootstrapPath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\FamilyBrowserDeploymentBootstrapService.cs'
$guardPath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\FamilyBrowserNativeCommandGuardService.cs'
$dialogGuardPath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\FamilyThumbnailConstraintDialogGuard.cs'
$thumbnailPreviewPath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\FamilyThumbnailPreviewService.cs'
$systemApplyPath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\SystemTypeApplyExecutionService.cs'
$requestStorePath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\FamilyBrowserRequestStore.cs'
$requestRecordPath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\FamilyBrowserRequestRecord.cs'
$requestAttachmentRecordPath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\FamilyBrowserRequestAttachmentFile.cs'
$userSettingsPath = Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025\FamilyBrowserUserSettingsStore.cs'
$nestedPropagationTestPath = Join-Path $scriptRoot 'Test-NestedFamilyDifferencePropagation.ps1'
$performanceTestPath = Join-Path $scriptRoot 'Test-FamilyBrowserPerformance.ps1'
$shellJsPath = Join-Path $sharedRoot 'family-browser-shell.js'
$rowWindowJsPath = Join-Path $sharedRoot 'family-browser-row-window.js'
$uiHarnessPath = Join-Path $repoRoot 'KKY_FamilyBrowser_Automation\KKY_FamilyBrowser_UiAuditHarness\Program.cs'
$workflowHarnessPath = Join-Path $repoRoot 'KKY_FamilyBrowser_Automation\KKY_FamilyBrowser_WorkflowAuditHarness\Program.cs'
$uiHarnessScriptPath = Join-Path $scriptRoot 'Invoke-FamilyBrowserUiAuditHarness.ps1'
$uiContractPath = Join-Path $scriptRoot 'FamilyBrowserUiAudit.contract.json'
$uiContract = Get-Content -Raw -Encoding UTF8 -LiteralPath $uiContractPath | ConvertFrom-Json
$uiScenarioNames = @($uiContract.scenarios | ForEach-Object { [string]$_.name })

function Assert-WorkflowScenario([string]$workflowId, [string]$scenarioName) {
    Add-WorkflowCheck $workflowId "UI contract includes $scenarioName" ($uiScenarioNames -contains $scenarioName) 'Required UI scenario is missing.'
}

Assert-Token $trackingPath 'PendingTracking' 'local write-ahead tracking spool exists'
Assert-Token $trackingPath 'PersistOperationEntries' 'durable operation persistence exists'
Assert-Token $trackingPath 'PersistStandardCandidateEntries' 'durable standard candidate persistence exists'
Assert-Token $trackingPath 'PersistElementChangeCommits' 'durable project element-change persistence exists'
Assert-Token $trackingPath 'LoadImmutableOperationEntries' 'immutable operation history reader exists'
Assert-Token $trackingPath 'LoadImmutableElementChangeCommits' 'immutable project element-change history reader exists'
Assert-Token $trackingPath 'File.Move(temporary, path)' 'tracking entries use atomic same-folder promotion'
Assert-Token $catalogPath 'FamilyBrowserTrackingPersistenceService.FlushPending(workspaceRoot)' 'project observation flushes offline tracking'
Assert-Token $catalogPath 'LoadImmutableOperationEntries(workspaceRoot, 10000)' 'project attribution reads immutable operation history'
Assert-MethodToken $catalogPath 'OperationMatches' '!string.IsNullOrWhiteSpace(operation.TypeName)' 'Family Type attribution requires an exact recorded type'
Assert-MethodToken $catalogPath 'OperationMatches' 'Normalize(operation.TypeName), Normalize(change.TypeName)' 'Family Type attribution compares exact names'

if (-not ('FamilyBrowserElementTrackingScopePolicy' -as [type])) {
    Add-Type -Path $elementTrackingScopePolicyPath
}
$auxiliaryScopeCases = @(
    @{ Name = 'PipingSystem class'; Class = 'PipingSystem'; CategoryId = ''; Category = ''; Expected = $true },
    @{ Name = 'MechanicalSystem class'; Class = 'MechanicalSystem'; CategoryId = ''; Category = ''; Expected = $true },
    @{ Name = 'ElectricalSystem class'; Class = 'ElectricalSystem'; CategoryId = ''; Category = ''; Expected = $true },
    @{ Name = 'CableTrayRun class'; Class = 'CableTrayRun'; CategoryId = ''; Category = ''; Expected = $true },
    @{ Name = 'ConduitRun class'; Class = 'ConduitRun'; CategoryId = ''; Category = ''; Expected = $true },
    @{ Name = 'pipe system category'; Class = 'Element'; CategoryId = '-2008043'; Category = ''; Expected = $true },
    @{ Name = 'duct system category'; Class = 'Element'; CategoryId = '-2008015'; Category = ''; Expected = $true },
    @{ Name = 'pipe centerline category'; Class = 'Element'; CategoryId = '-2008045'; Category = ''; Expected = $true },
    @{ Name = 'cable tray run category'; Class = 'Element'; CategoryId = '-2008150'; Category = ''; Expected = $true },
    @{ Name = 'localized centerline fallback'; Class = 'Element'; CategoryId = ''; Category = '중심선'; Expected = $true },
    @{ Name = 'localized piping-system fallback'; Class = 'Element'; CategoryId = ''; Category = '배관 시스템'; Expected = $true },
    @{ Name = 'ordinary pipe instance'; Class = 'Pipe'; CategoryId = '-2008044'; Category = 'Pipes'; Expected = $false },
    @{ Name = 'ordinary cable tray instance'; Class = 'CableTray'; CategoryId = '-2008132'; Category = 'Cable Trays'; Expected = $false },
    @{ Name = 'wall instance'; Class = 'Wall'; CategoryId = '-2000011'; Category = 'Walls'; Expected = $false },
    @{ Name = 'MEP system type definition'; Class = 'PipingSystemType'; CategoryId = '-2008043'; Category = 'Piping Systems'; IsElementType = $true; Expected = $false },
    @{ Name = 'material definition'; Class = 'Material'; CategoryId = '-2000032'; Category = 'Materials'; Expected = $false }
)
foreach ($scopeCase in $auxiliaryScopeCases) {
    $actual = [FamilyBrowserElementTrackingScopePolicy]::IsAuxiliarySupportRecord($scopeCase.Class, $scopeCase.CategoryId, $scopeCase.Category, [bool]$scopeCase.IsElementType)
    Add-WorkflowCheck 'project-element-change-ledger' ("top-level history scope classifies " + $scopeCase.Name) ($actual -eq $scopeCase.Expected) ("Expected {0}; actual {1}." -f $scopeCase.Expected, $actual)
}
foreach ($auxiliaryCategoryId in @(
    '-2000288', '-2008045', '-2008072', '-2008051', '-2008001', '-2008066', '-2008021',
    '-2008139', '-2008141', '-2008136', '-2008140', '-2008214', '-2008196', '-2008210',
    '-2008150', '-2008149', '-2008043', '-2008015', '-2008158', '-2008159', '-2008156', '-2008157',
    '-2008037', '-2008152'
)) {
    $actual = [FamilyBrowserElementTrackingScopePolicy]::IsAuxiliarySupportRecord('Element', $auxiliaryCategoryId, '', $false)
    Add-WorkflowCheck 'project-element-change-ledger' ("top-level history scope excludes auxiliary category " + $auxiliaryCategoryId) $actual 'Auxiliary category ID was not excluded.'
}

Assert-WorkflowMethodToken 'managed-folder-first-run' $dashboardPath 'CompleteInitialOpenRefresh' 'RefreshManagedFolderAvailabilityState(queueOnboarding: true);' 'startup checks managed-folder reachability'
Assert-WorkflowMethodToken 'managed-folder-first-run' $dashboardPath 'StartPostStartupServices' 'QueueManagedFolderOnboardingIfNeeded();' 'startup queues first-run guidance after shell render'
Assert-WorkflowMethodToken 'managed-folder-first-run' $managedFolderSetupPath 'Configure' 'CreateAndProbeManagedFolders(root);' 'TEST folder setup verifies write access before activation'
Assert-WorkflowMethodToken 'managed-folder-first-run' $managedFolderSetupPath 'Configure' 'if (!IsInternalNetworkShare(root))' 'TEST folder setup rejects local paths'
Assert-WorkflowScenario 'managed-folder-first-run' 'admin-home-managed-folder-unavailable'

Assert-WorkflowToken 'managed-folder-homepage-return' $managedFolderTransitionPath 'AnalyzeMigration(sourceRoot, probe.ManagedPolicyPath)' 'homepage return performs migration preflight'
Assert-WorkflowToken 'managed-folder-homepage-return' $managedFolderTransitionPath 'TryClearOverrideRoot' 'homepage return removes TEST pointer only after activation'
Assert-WorkflowToken 'managed-folder-homepage-return' $managedFolderTransitionPath 'FamilyBrowserTrackingPersistenceService.FlushPending(_workspaceRoot)' 'homepage return immediately flushes locally protected tracking records'
Assert-WorkflowMethodOrder 'managed-folder-homepage-return' $managedFolderTransitionPath 'SwitchToHomepageManagedFolder' 'GetPendingElementSessionCheckpointCount(_workspaceRoot)' 'TryClearOverrideRoot' 'Switch Only checks local-save tracking evidence before removing the TEST pointer'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $managedFolderTransitionPath 'SwitchToHomepageManagedFolder' 'GetInvalidElementSessionCheckpointCount()' 'Switch Only fails closed while a local tracking checkpoint requires recovery'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $managedFolderTransitionPath 'SwitchToHomepageManagedFolder' 'Switch Only is blocked because this PC still has workshared local-save tracking checkpoints' 'blocked Switch Only gives a recoverable synchronization or explicit migration path'
Assert-WorkflowMethodOrder 'managed-folder-homepage-return' $managedFolderTransitionPath 'SwitchToHomepageManagedFolder' 'FamilyBrowserTrackingPersistenceService.FlushPending(_workspaceRoot)' 'AnalyzeMigration(sourceRoot, probe.ManagedPolicyPath)' 'pending tracking is settled in the TEST source before migration analysis copies it'
Assert-WorkflowMethodOrder 'managed-folder-homepage-return' $managedFolderTransitionPath 'SwitchToHomepageManagedFolder' 'GetInvalidElementSessionCheckpointCount()' 'TryClearOverrideRoot' 'locked or corrupt local checkpoints stop the transition before the TEST pointer is removed'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $managedFolderTransitionPath 'SwitchToHomepageManagedFolder' 'GetActiveUncommittedSessionCount()' 'management-folder changes stop while this Revit process owns uncommitted tracking evidence'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $managedFolderTransitionPath 'SwitchToHomepageManagedFolder' 'GetOtherRevitProcessCount()' 'management-folder changes stop while another Revit process can still write to the old destination'
Assert-WorkflowMethodOrder 'managed-folder-homepage-return' $managedFolderTransitionPath 'SwitchToHomepageManagedFolder' 'FlushPendingForManagedFolderTransition(_workspaceRoot, sourceRoot)' 'TryClearOverrideRoot' 'checkpoint destination migration succeeds before the TEST pointer is removed'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $managedFolderTransitionPath 'SwitchToHomepageManagedFolder' 'ElementSessionCheckpointRebindFailedCount' 'checkpoint migration failures are visible to the transition controller'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $managedFolderTransitionPath 'SwitchToHomepageManagedFolder' 'TryApplyPersistedOverride' 'failed tracking migration rolls the active management path back to TEST'
Assert-WorkflowMethodOrder 'managed-folder-homepage-return' $managedFolderTransitionPath 'SwitchToHomepageManagedFolder' 'FamilyBrowserManagementContextLock.Acquire' 'GetOtherRevitProcessCount()' 'the cross-process management lock closes the process-start race before Revit sessions are inspected'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $managementContextLockPath 'Acquire' 'AbandonedMutexException' 'an abandoned management transition is recoverable instead of permanently blocking startup'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $managementContextLockPath 'Acquire' 'Timed out waiting for another Revit process' 'management-context lock contention fails explicitly'
Assert-WorkflowToken 'managed-folder-homepage-return' $workflowHarnessPath 'management-context changes are serialized across Revit processes and retry after release' 'workflow harness includes an actual cross-thread management-context lease contention fixture'
Assert-WorkflowMethodOrder 'managed-folder-homepage-return' $managedFolderSetupPath 'MigrateToHomepage' 'string sourceFingerprintBefore = ComputeManagedSourceFingerprint' 'string sourceFingerprintAfter = ComputeManagedSourceFingerprint' 'managed-folder migration compares a full source manifest before and after copying'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $managedFolderSetupPath 'MigrateToHomepage' 'result.SourceChangedDuringMigration = true;' 'source changes during migration are an explicit failed state'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $managedFolderSetupPath 'MigrateToHomepage' 'result.RolledBackFileCount++;' 'new destination files are rolled back after a partial migration failure'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $managedFolderSetupPath 'MigrateToHomepage' 'result.RollbackFailedFileCount++;' 'changed or locked rollback files remain an explicit administrator-recovery state'
Assert-WorkflowToken 'managed-folder-homepage-return' $uiHarnessPath 'a partial destination copy survived a deterministic mid-copy failure' 'IE harness injects a deterministic mid-copy failure and verifies rollback'
Assert-WorkflowMethodOrder 'managed-folder-first-run' $managedFolderSetupPath 'TryApplyPersistedOverride' 'FamilyBrowserMachineConfig previousMachineConfig = FamilyBrowserMachineConfigStore.Load();' 'RestoreMachineConfiguration(previousMachineConfig, currentUser);' 'persisted TEST activation restores the previous machine configuration when a later step fails'
Assert-WorkflowMethodOrder 'managed-folder-first-run' $managedFolderSetupPath 'Configure' 'FamilyBrowserMachineConfig previousMachineConfig = FamilyBrowserMachineConfigStore.Load();' 'FamilyBrowserStandardPolicyStore.SetRequestStore' 'first-run management setup captures rollback state before changing the active policy path'
Assert-WorkflowMethodOrder 'managed-folder-first-run' $managedFolderSetupPath 'Configure' 'FamilyBrowserStandardPolicyStore.SetRequestStore' 'SaveOverrideRoot(root);' 'the TEST pointer is committed only after the shared request store is ready'
Assert-WorkflowMethodToken 'managed-folder-first-run' $managedFolderSetupPath 'Configure' 'RestoreMachineConfiguration(previousMachineConfig, currentUser);' 'failed first-run setup restores the previous active management context'
Assert-WorkflowMethodToken 'managed-folder-first-run' $managedFolderSetupPath 'Configure' 'RestoreOverrideRoot(previousOverrideRoot);' 'failed first-run setup restores the previous TEST pointer state'
Assert-WorkflowMethodToken 'managed-folder-first-run' $managedFolderSetupPath 'SaveOverrideRoot' 'FamilyBrowserAtomicFileService.Promote(temporaryPath, pointerPath);' 'the per-user TEST pointer is promoted atomically instead of deleting the committed pointer first'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'GetDeferredPolicyDisableSessionCount' 'session.PolicyDisableDeferred' 'remote tracking disable deferral is exposed as live session state'
Assert-WorkflowToken 'project-element-change-ledger' $projectCatalogDashboardPath 'deferredTrackingDisablePill' 'the header exposes a remote tracking-disable session that still needs Save or Synchronize'
Assert-WorkflowToken 'project-element-change-ledger' $projectCatalogDashboardPath 'deferredTrackingDisableQueue' 'the home dashboard explains the remote tracking-disable commit boundary and mixed-window limitation'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $trackingPath 'RebindElementSessionCheckpointsNoLock' 'ElementSessionCheckpointLockUnavailable = true;' 'checkpoint migration lock contention is an explicit failure state'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $trackingPath 'RebindElementSessionCheckpointsNoLock' 'result.FailedCount++;' 'checkpoint rebind failures fail the overall management-folder transition'
Assert-WorkflowToken 'managed-folder-homepage-return' $managedFolderSetupPath '"ProjectCatalogs"' 'project catalog baselines migrate with TEST data'
Assert-WorkflowToken 'managed-folder-homepage-return' $managedFolderSetupPath '"StandardRevisionManifests"' 'Standard RVT revision baselines migrate with TEST data'
Assert-WorkflowToken 'managed-folder-homepage-return' $managedFolderSetupPath '"ElementChangeHistory"' 'element change history migrates with TEST data'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $managedFolderSetupPath 'ReadMigratedManagedFileBytes' 'IsElementChangeHistoryRelativePath(relativePath)' 'immutable element history is copied without checksum-breaking JSON rebasing'
Assert-WorkflowToken 'managed-folder-homepage-return' $managedFolderSetupPath 'name.EndsWith(".kky-lock"' 'shared policy lock files are excluded from managed-folder migration'
Assert-WorkflowScenario 'managed-folder-homepage-return' 'admin-home-test-folder-homepage-available'

Assert-WorkflowMethodToken 'startup-cache-offline' $dataLoaderPath 'LoadSnapshotProjection' 'OfflineCacheUsed' 'startup loader records validated offline-cache use'
Assert-WorkflowMethodToken 'startup-cache-offline' $dataLoaderPath 'LoadSnapshotProjection' 'TryResolveArtifact' 'startup loader resolves managed artifacts through the local cache'
Assert-WorkflowMethodToken 'startup-cache-offline' $dataLoaderPath 'TryLoadRowCache' 'row-cache-hit' 'startup row cache reports a warm hit'
Assert-WorkflowMethodToken 'startup-cache-offline' $dataLoaderPath 'ReplaceFileAtomic' 'FamilyBrowserAtomicFileService.Promote' 'V2 artifacts use the shared recoverable promotion path'
Assert-WorkflowMethodToken 'startup-cache-offline' $atomicFileServicePath 'Promote' 'File.Move(backupPath, destinationPath)' 'failed V2 artifact promotion restores the committed cache copy'
Assert-WorkflowToken 'startup-cache-offline' $atomicFileServicePath 'CreateSiblingTemporaryPath' 'legacy .NET hosts use short temporary file names on deep managed paths'
Assert-WorkflowToken 'startup-cache-offline' $performanceTestPath 'CacheOfflineMs' 'performance gate measures offline cache loading'

foreach ($scenarioName in @('admin-family-missing-standard-rvt', 'admin-family-missing-standard-list', 'admin-family-detail-with-preview', 'admin-system-missing-standard-list', 'admin-system-with-data')) {
    Assert-WorkflowScenario 'standard-setup-readiness' $scenarioName
}

Assert-WorkflowMethodToken 'standard-trade-switch' $dashboardPath 'SetBrowseDiscipline' 'ClearModelerAllSlotsCache();' 'trade switch invalidates combined rows'
Assert-WorkflowMethodToken 'standard-trade-switch' $dashboardPath 'SetBrowseDiscipline' 'allowStartupPreload: false, allowPreparedSlotData: true' 'trade switch reloads the selected prepared slot without a full startup scan'
Assert-WorkflowToken 'standard-trade-switch' $dashboardPath 'syncBrowseTreeTradeLabel(filter)' 'trade switch updates the left tree label immediately'
Assert-WorkflowToken 'standard-trade-switch' $dashboardPath 'resetBrowseTreeFilters();' 'trade switch clears stale category-tree selection'

Assert-WorkflowMethodToken 'standard-source-revision-block' $revisionPath 'Probe' 'ComputeRevisionHash(path)' 'Standard RVT probe compares sampled content revision'
Assert-WorkflowMethodToken 'standard-source-revision-block' $revisionDashboardPath 'ApplyStandardRevisionBlockIfNeeded' '_loadableRows = new List<BrowserRow>();' 'stale Standard RVT clears Family browser rows'
Assert-WorkflowMethodToken 'standard-source-revision-block' $revisionDashboardPath 'ApplyStandardRevisionBlockIfNeeded' '_systemRows = new List<SystemRow>();' 'stale Standard RVT clears System browser rows'
Assert-WorkflowScenario 'standard-source-revision-block' 'admin-home-standard-rvt-changed'
Assert-WorkflowScenario 'standard-source-revision-block' 'modeler-family-standard-rvt-unavailable'

Assert-WorkflowToken 'admin-live-state-refresh' $dashboardPath 'LoadStandardPolicy()' 'post-action refresh reads the live policy store'
Assert-WorkflowToken 'admin-live-state-refresh' $dashboardPath '_allowStartupPreloadDuringShellRefresh = false' 'post-action refresh always releases startup-preload state'
Assert-WorkflowToken 'admin-live-state-refresh' $dashboardPath 'shell-policy-source:' 'live versus startup policy source is diagnosable'
Assert-WorkflowScenario 'admin-live-state-refresh' 'admin-standard-settings-layout'

Assert-WorkflowScenario 'family-list-interaction' 'modeler-family-with-data'
Assert-WorkflowScenario 'family-list-interaction' 'viewport-1280-family'
Assert-WorkflowToken 'family-list-interaction' $dashboardPath 'function queueFilterRows()' 'search/filter uses a debounced browser-only path'
Assert-WorkflowToken 'family-list-interaction' $rowWindowJsPath 'w.resizeColumnForAudit = function' 'resizable columns have an automated seam'

Assert-WorkflowMethodToken 'family-detail-integrity' $dashboardPath 'HydrateSelectedRowDetailFromV2' 'FamilyBrowserDataLoader.Default.LoadDetail' 'Family detail is hydrated by stable detail key on selection'
Assert-WorkflowMethodToken 'family-detail-integrity' $dataLoaderPath 'LoadDetail' 'TryResolveArtifact' 'Family detail is loaded through the validated managed/local artifact resolver'
Assert-WorkflowToken 'family-detail-integrity' $dashboardPath 'AppendLookupCsvPreviewText' 'lookup CSV metadata is part of the unified Family detail'
Assert-WorkflowToken 'family-detail-integrity' $dashboardPath 'renderDetachedPreview' 'detached Family detail renders the captured preview'
Assert-WorkflowScenario 'family-detail-integrity' 'admin-family-detail-with-preview'

Assert-WorkflowToken 'nested-family-propagation' $nestedPropagationTestPath "Assert-True (`$parent.Status -eq 'DifferentFromStandard')" 'nested child difference marks its parent different'
Assert-WorkflowToken 'nested-family-propagation' $nestedPropagationTestPath "Assert-True (`$missingChild.Status -eq 'NestedMissingFromParent')" 'missing nested child has an explicit parent-composition state'
Assert-WorkflowToken 'nested-family-propagation' $nestedPropagationTestPath "Assert-True (`$grandParent.Status -eq 'DifferentFromStandard')" 'nested differences propagate transitively to top-level parents'

Assert-WorkflowToken 'system-type-apply-lifecycle' $systemApplyPath 'ApplyAuthoritativeCompoundStructure' 'System Type apply includes authoritative compound layers'
Assert-WorkflowToken 'system-type-apply-lifecycle' $systemApplyPath 'ApplyAuthoritativeDetailedSystemTypeDefinition' 'System Type apply includes Railing and Stair components'
Assert-WorkflowToken 'system-type-apply-lifecycle' $systemApplyPath 'GuardDependencyLoadDidNotCreateDuplicateFamilies' 'System Type apply guards dependency duplicates'
Assert-WorkflowToken 'system-type-apply-lifecycle' $systemTypeApplyStorePath 'FamilyBrowserUniqueJsonReportStore.Save' 'System Type apply history is collision-free and atomic'
Assert-WorkflowToken 'system-type-apply-lifecycle' $systemTypePreflightStorePath 'FamilyBrowserUniqueJsonReportStore.Save' 'System Type preflight history is collision-free and atomic'
Assert-WorkflowScenario 'system-type-apply-lifecycle' 'admin-system-with-data'

Assert-WorkflowToken 'system-detail-integrity' $dashboardPath 'data-system-component-table=\"true\"' 'System detail renders optional component rows as a table'
Assert-WorkflowToken 'system-detail-integrity' $dashboardPath 'data-system-curtain-panel-table=\\\"true\\\"' 'System detail always renders mandatory Curtain Panel dependencies'
Assert-WorkflowToken 'system-detail-integrity' $systemRoutingUnitPath 'kkyfb:measurement-unit/' 'routing/layer unit changes are routed to the host preference service'
Assert-WorkflowToken 'system-detail-integrity' $systemComponentUnitPath 'systemRoutingDisplayUnit' 'dependent-component tables share the routing/layer display unit'
Assert-WorkflowToken 'system-detail-integrity' $measurementPreferencePath 'measurement-unit.txt' 'System detail display units persist across browser sessions'
Assert-WorkflowToken 'system-detail-integrity' $userSettingsPath 'system-type-detail-components.txt' 'optional System component comparison preference persists'
Assert-WorkflowScenario 'system-detail-integrity' 'admin-system-with-data'
Assert-WorkflowScenario 'system-detail-integrity' 'admin-system-components-disabled'

Assert-WorkflowMethodToken 'current-model-check-baseline' $projectCatalogPath 'AcceptFromProjectSnapshot' 'PersistSnapshot(workspaceRoot, doc, snapshot, true, acceptedBy)' 'successful detailed check accepts the project catalog baseline'
Assert-WorkflowMethodToken 'current-model-check-baseline' $projectCatalogDashboardPath 'AcceptProjectCatalogAfterCurrentModelCheck' 'AcceptFromProjectSnapshot' 'dashboard accepts catalog only from completed check data'
Assert-WorkflowMethodToken 'current-model-check-baseline' $projectSnapshotStorePath 'CanPublishSharedProjectState' 'AllLocalChangesSavedToCentral' 'shared model-check publication rejects locally saved but unsynchronized workshared data'
Assert-WorkflowMethodToken 'current-model-check-baseline' $projectSnapshotStorePath 'CanPublishSharedProjectState' 'centralInfo.LatestCentralVersion > localInfo.LatestCentralVersion' 'shared model-check publication rejects a local file behind Central'
Assert-WorkflowMethodToken 'current-model-check-baseline' $projectSnapshotStorePath 'SaveLatestProjectScan' 'CanPublishSharedProjectState' 'latest project cache enforces publication readiness at its persistence boundary'
Assert-WorkflowMethodToken 'current-model-check-baseline' $projectSnapshotStorePath 'SaveLatestProjectScan' 'MatchesSnapshotGeneration' 'latest project cache rejects a Standard snapshot generation replaced after comparison construction'
Assert-WorkflowMethodToken 'current-model-check-baseline' $revisionPath 'BuildCurrentRevisionToken' 'state.SnapshotAtUtc' 'Standard revision identity includes the registered snapshot generation'
Assert-WorkflowMethodToken 'standard-source-revision-block' $revisionDashboardPath 'StandardRevisionStatesEquivalent' 'left.SnapshotPath' 'open dashboard detects a remotely replaced Standard snapshot generation'
Assert-WorkflowMethodToken 'standard-edit-commit-lifecycle' $standardRegistryPath 'PublishSnapshotAndActiveRegistration' 'AcquirePublicationLock' 'Standard snapshot, registration, revision, and browser artifacts share one publication lock'
Assert-WorkflowMethodOrder 'standard-edit-commit-lifecycle' $standardRegistryPath 'PublishSnapshotAndActiveRegistration' 'FamilyBrowserDataLoader.PublishStandardArtifacts' 'SaveActiveRegistrationCore' 'active Standard pointer changes only after V2 artifacts are ready'
Assert-WorkflowMethodToken 'standard-edit-commit-lifecycle' $standardRegistryPath 'WriteTextAtomic' 'FamilyBrowserAtomicFileService.Promote' 'Standard registration files are atomically promoted'
Assert-WorkflowMethodToken 'standard-edit-commit-lifecycle' $standardRegistrationPath 'Register' 'PublishSnapshotAndActiveRegistration' 'full Standard scan uses coordinated publication'
Assert-WorkflowMethodToken 'current-model-check-baseline' $automaticModelCheckPath 'Execute' 'standardRevisionBeforePublication' 'automatic model check revalidates the Standard RVT after report construction'
Assert-WorkflowMethodToken 'current-model-check-baseline' $automaticModelCheckPath 'Execute' 'finalPublicationReason' 'automatic model check revalidates project publication readiness immediately before publishing'
Assert-WorkflowMethodToken 'current-model-check-baseline' $automaticModelCheckPath 'Execute' 'dirtyMarkerAtCheckStart == null && cached != null && cached.Success' 'protected-content marker forces a fresh automatic check instead of cached reuse'
Assert-WorkflowMethodToken 'current-model-check-baseline' $projectTrackingStorePath 'MarkCurrentModelCheckRequired' 'MergeDirtyMarkerItems' 'multi-client protected-content markers merge their findings'
Assert-WorkflowMethodToken 'current-model-check-baseline' $projectTrackingStorePath 'ClearCurrentModelCheckRequired' 'DirtyMarkerIdentityMatches' 'a completed check cannot delete a newer marker generation'
Assert-WorkflowToken 'current-model-check-baseline' $projectTrackingOutputStorePath 'FamilyBrowserUniqueJsonReportStore.Save' 'tracking stamp history is collision-free and atomic'
Assert-WorkflowMethodToken 'current-model-check-baseline' $automaticModelCheckPath 'WriteStatus' 'stream.Flush(true);' 'automatic model-check latest status is durably flushed before publication'
Assert-WorkflowMethodToken 'current-model-check-baseline' $automaticModelCheckPath 'WriteStatus' 'temporaryPath = string.Empty;' 'automatic model-check status cleanup preserves the promoted file'
Assert-WorkflowMethodToken 'current-model-check-baseline' $automaticModelCheckPath 'WriteStatus' 'File.Delete(temporaryPath);' 'failed automatic model-check status writes clean their temporary file'
Assert-WorkflowToken 'family-load-save-lifecycle' $loadableFamilySyncStorePath 'FamilyBrowserUniqueJsonReportStore.Save' 'Family load history is collision-free and atomic'
Assert-WorkflowToken 'multi-client-tracking' $uniqueJsonReportStorePath 'stream.Flush(true)' 'multi-client report publication flushes bytes before atomic promotion'
Assert-WorkflowMethodOrder 'current-model-check-baseline' $automaticModelCheckPath 'Execute' 'FamilyBrowserFileGuardPathMatcher.FindMatchingTarget' 'ProjectSnapshotStore.CanPublishSharedProjectState' 'automatic model check tests publication readiness only after confirming the project is managed'
Assert-WorkflowMethodToken 'current-model-check-baseline' $automaticModelCheckPath 'ProcessPending' 'Last reason: ' 'automatic model check deferred status preserves the actual retry cause'
Assert-WorkflowMethodToken 'current-model-check-baseline' $dashboardPath 'RunCurrentModelCheckForTarget' 'RefreshDashboard(deepProjectScan: true)' 'manual model check routes through the guarded deep-scan workflow'
Assert-WorkflowMethodToken 'current-model-check-baseline' $dashboardPath 'RefreshDashboard' 'finalPublicationReason' 'manual model check revalidates project publication readiness immediately before publishing'
Assert-WorkflowScenario 'current-model-check-baseline' 'admin-home-project-catalog-baseline-missing'
Assert-WorkflowToken 'current-model-check-baseline' $uiHarnessScriptPath "'--projectCatalogBaselineMissing'" 'UI runner forwards the baseline-missing state into the IE harness'
Assert-WorkflowToken 'current-model-check-baseline' $uiHarnessPath 'CheckProjectCatalogState(browser, options, result);' 'IE harness validates the rendered baseline state'

Assert-WorkflowMethodToken 'external-project-change-detection' $projectCatalogPath 'Observe' 'PersistSnapshot(workspaceRoot, doc, snapshot, false, string.Empty)' 'automatic catalog observation cannot accept the baseline'
Assert-WorkflowMethodToken 'external-project-change-detection' $projectCatalogPath 'Observe' 'ProjectSnapshotStore.CanPublishSharedProjectState' 'automatic catalog observation cannot publish unsynchronized or stale workshared state'
Assert-WorkflowMethodToken 'external-project-change-detection' $projectCatalogPath 'PersistSnapshot' 'ProjectSnapshotStore.CanPublishSharedProjectState' 'project catalog publication revalidates worksharing state while holding its publication lock'
Assert-WorkflowMethodToken 'external-project-change-detection' $projectCatalogPath 'IsPublishedObservationState' 'BaselineMissing' 'a newly published catalog without an accepted baseline still counts as a durable observation'
Assert-WorkflowToken 'external-project-change-detection' $projectCatalogPath 'ExternalUntracked' 'unmatched catalog differences remain external or untracked'
Assert-WorkflowScenario 'external-project-change-detection' 'modeler-home-project-catalog-untracked-change'
Assert-WorkflowToken 'external-project-change-detection' $uiHarnessScriptPath "'--projectCatalogChanged'" 'UI runner forwards changed project catalog state'
Assert-WorkflowToken 'external-project-change-detection' $uiHarnessScriptPath "'--projectCatalogUntracked'" 'UI runner forwards external/untracked attribution state'
Assert-WorkflowMethodOrder 'external-project-change-detection' $projectCatalogDashboardPath 'QueueProjectCatalogObservation' '_modelessActionDispatcher("project-catalog-observe-auto")' '_lastProjectCatalogObservationUtc = DateTime.UtcNow;' 'failed project catalog dispatch does not consume the retry throttle window'
Assert-WorkflowMethodToken 'external-project-change-detection' $projectCatalogPath 'GetProjectFolder' 'GetStablePathIdentity(identityPath)' 'project catalog storage survives central-file identity replacement by using the stable project path'
Assert-WorkflowMethodToken 'external-project-change-detection' $projectCatalogPath 'ResolveProjectFolder' 'matchingLegacyFolders.Count > 1' 'ambiguous legacy catalog folders fail closed instead of selecting an arbitrary baseline'
Assert-WorkflowMethodToken 'external-project-change-detection' $projectCatalogPath 'SaveSnapshot' 'return Path.Combine("Snapshots", fileName);' 'catalog manifests store management-root-relative snapshot references across mapped-drive and UNC clients'
Assert-WorkflowMethodToken 'external-project-change-detection' $projectCatalogPath 'ResolveSnapshotPath' 'Path.GetFileName' 'legacy absolute snapshot references are rebound only to the current project catalog folder'
Assert-WorkflowMethodToken 'external-project-change-detection' $projectCatalogPath 'PersistSnapshot' 'File.Exists(manifestPath) && manifest == null' 'an unreadable catalog manifest is preserved and cannot be silently replaced'
Assert-WorkflowMethodToken 'external-project-change-detection' $projectCatalogPath 'ValidateSnapshot' 'ComputeCatalogHash(snapshot.Entries)' 'catalog snapshot entry keys, counters and content hash are verified before use'
Assert-WorkflowMethodToken 'external-project-change-detection' $projectCatalogPath 'LoadLatestState' 'state does not match its manifest' 'partial manifest/state publication is surfaced instead of shown as a healthy catalog'
Assert-WorkflowMethodToken 'external-project-change-detection' $projectCatalogPath 'LoadLatestState' 'AcquireCatalogLock(folder)' 'project catalog readers cannot observe a writer between manifest and state promotion'
Assert-WorkflowToken 'external-project-change-detection' $projectCatalogDashboardPath 'PublicationDeferred' 'deferred project catalog publication is visible to the user'
Assert-WorkflowToken 'external-project-change-detection' $projectCatalogDashboardPath 'Browser operation matched / actor unproven' 'name-only Browser operation matches are not presented as proof of the actor'

Assert-WorkflowToken 'project-element-change-ledger' $elementChangeTrackingPath 'GetAddedElementIds()' 'Revit-created element IDs are collected'
Assert-WorkflowToken 'project-element-change-ledger' $elementChangeTrackingPath 'GetModifiedElementIds()' 'Revit-modified element IDs are collected'
Assert-WorkflowToken 'project-element-change-ledger' $elementChangeTrackingPath 'GetDeletedElementIds()' 'Revit-deleted element IDs are collected'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'RequestDocumentSessionBaselineRefresh' 'System.Threading.Interlocked.Exchange' 'modeless policy loads can safely request a Revit-context baseline'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HasDocumentSession' 'Sessions.ContainsKey(BuildRuntimeKey(doc))' 'startup retries verify a real baseline session instead of assuming one exists'
Assert-WorkflowToken 'project-element-change-ledger' $elementChangeTrackingPath 'GetTransactionNames()' 'transaction names are retained for diagnostics'
Assert-WorkflowToken 'project-element-change-ledger' $elementChangeTrackingPath 'TransactionUndone' 'Undo removes reverted activity from the committed ledger'
Assert-WorkflowToken 'project-element-change-ledger' $elementChangeTrackingPath 'TransactionRedone' 'Redo restores reapplied activity to the committed ledger'
Assert-WorkflowMethodOrder 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentCommitted' 'if (!succeeded' 'PersistElementChangeCommits' 'failed or cancelled commits cannot append element history'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentCommitted' 'workshared && !string.Equals(kind, "SynchronizeWithCentral"' 'workshared element history commits only at synchronization'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentSynchronizingWithCentral' 'session.SynchronizingWithCentral = true;' 'incoming synchronization changes are excluded from local-user attribution'
Assert-WorkflowMethodOrder 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentSynchronizingWithCentral' 'ResolveDocumentPolicyEnabled(workspaceRoot, doc, false)' 'SynchronizingDocumentKeys.Add(runtimeKey);' 'disabled or out-of-scope tracking never leaves a synchronization suppression marker'
Assert-WorkflowMethodOrder 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentSynchronizingWithCentral' 'SynchronizingDocumentKeys.Add(runtimeKey);' 'BeginDocumentSession(workspaceRoot, doc);' 'synchronization suppression begins even when the tracking baseline cannot be created'
Assert-WorkflowToken 'project-element-change-ledger' $elementChangeTrackingPath 'StartReloadLatestBridge' 'Reload Latest event support is attached without a hard Revit-version dependency'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'StartReloadLatestBridge' 'Element tracking Reload Latest bridge failed' 'Reload Latest delegate binding failures are diagnosed instead of silently disabling suppression'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleReloadingLatestBridge' 'HandleDocumentReloadLatestStartFailure' 'Reload Latest start callback cannot throw silently or lose conservative suppression'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleReloadedLatestBridge' 'CloseExternalUpdateWindowAfterUnknownCompletion' 'Reload Latest completion callback cannot fail silently or leave suppression active'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentReloadingLatest' 'session.ReloadingLatest = true;' 'incoming Reload Latest changes are excluded from local-user attribution'
Assert-WorkflowMethodOrder 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentReloadingLatest' 'ResolveDocumentPolicyEnabled(workspaceRoot, doc, false)' 'ReloadingDocumentKeys.Add(runtimeKey);' 'disabled or out-of-scope tracking never leaves a Reload Latest suppression marker'
Assert-WorkflowMethodOrder 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentReloadingLatest' 'ReloadingDocumentKeys.Add(runtimeKey);' 'BeginDocumentSession(workspaceRoot, doc);' 'Reload Latest suppression begins even when the tracking baseline cannot be created'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentReloadedLatest' 'RebaseSessionAfterExternalUpdate(doc, session);' 'successful Reload Latest rebases remote state while preserving local pending activity'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentReloadedLatest' 'ReloadingDocumentKeys.Remove(runtimeKey);' 'Reload Latest completion always closes the document-level suppression window'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentChanged' 'ShouldUseReloadLatestTransactionFallback(activity)' 'transaction-name fallback protects Revit versions without Reload Latest events'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'ShouldUseReloadLatestTransactionFallback' '_reloadLatestEventSource != null' 'transaction-name Reload Latest inference is disabled when the explicit Revit event bridge is active'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'ShouldUseReloadLatestTransactionFallback' 'return IsReloadLatestActivity(activity);' '2019 or a failed explicit event bridge retains the conservative transaction-name fallback'
Assert-WorkflowMethodOrder 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentChanged' 'externalUpdateInProgress = SynchronizingDocumentKeys.Contains(runtimeKey)' 'if (session == null)' 'incoming activity is rejected before late-baseline session creation'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentChanged' 'IsUnknownExternalUpdateStartNoLock(workspaceRoot)' 'an unreadable Sync or Reload start still suppresses incoming changes conservatively'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentSynchronizationStartFailure' 'MarkUnknownExternalUpdateStart' 'an unreadable synchronization start opens a workspace-scoped suppression window'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentReloadLatestStartFailure' 'MarkUnknownExternalUpdateStart' 'an unreadable Reload Latest start opens a workspace-scoped suppression window'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentCommitted' 'SynchronizingDocumentKeys.Remove(runtimeKey);' 'synchronization completion always closes the document-level suppression window'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentCommitted' 'UnknownSynchronizationStartRoots.Remove' 'synchronization completion closes workspace-scoped unknown-start suppression'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentReloadedLatest' 'UnknownReloadLatestStartRoots.Remove' 'Reload Latest completion closes workspace-scoped unknown-start suppression'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'CaptureBaseline' 'states.Clear();' 'partial baseline failures cannot become mass-created element history'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'CaptureBaseline' 'BuildStateCaptureContext(doc)' 'the baseline captures project/shared parameter definitions with one binding-map index'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'IsTrackableElement' 'element is ParameterElement' 'categoryless project and shared parameter definition elements are explicitly tracked'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'CaptureBaseline' 'ignoredAuxiliaryElementIds.Add' 'baseline indexes auxiliary IDs once so later deletion filtering remains O(1)'
Assert-WorkflowMethodOrder 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentChanged' 'UpdateCurrentStates(session, doc, added, modified, deleted)' 'ApplyActivity(session, activity)' 'auxiliary support elements are removed before activity and undo/redo bookkeeping'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'UpdateCurrentStates' 'FamilyBrowserElementTrackingTransitionPolicy.ShouldIgnoreChangedElement' 'changed-element filtering uses the executable transition policy'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'UpdateCurrentStates' 'session.IgnoredAuxiliaryElementIds.Contains(key)' 'null or deleted auxiliary support elements use the session hash index without rescanning the document'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'UpdateCurrentStates' 'FamilyBrowserElementTrackingTransitionPolicy.RestoreVisibleElementId' 'an ID recaptured as visible is removed from both same-event and session auxiliary indexes'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'RemoveIgnoredElementIdsFromSession' 'session.AppliedActivities.Concat(session.UndoneActivities)' 'late auxiliary classification also removes prior undo/redo noise'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'IsAuxiliaryTrackedState(display)' 'commit construction has a final auxiliary-record defense'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'FamilyBrowserElementTrackingTransitionPolicy.ResolveChangeKind' 'commit change-kind resolution uses the executable transition policy'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'FamilyBrowserElementHistoryProjectionPolicy.UnresolvedTransientTrackingKind' 'same-boundary create/delete evidence without metadata is marked unresolved instead of becoming a blank user row'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'IsAuxiliaryRevitSupportElement' 'familyInstance.SuperComponent' 'dependent nested Family instances are omitted while their top-level parent remains tracked'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'visibleChangesByCommit' 'history and explicit Excel export share one filtered row projection'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'IsUserFacingProjectElementChange' 'FamilyBrowserElementHistoryProjectionPolicy.IsUserFacingChange' 'detail and Excel use the shared immutable-history projection'
Assert-WorkflowMethodToken 'project-element-change-ledger' $trackingPath 'LoadTrackedProjectHistorySummaries' 'FamilyBrowserElementHistoryProjectionPolicy.CountUserFacingChanges' 'all-project summary counts use the same immutable-history projection as detail and Excel'
Assert-WorkflowToken 'project-element-change-ledger' $elementHistoryProjectionPolicyPath 'HiddenUnresolvedTransientCount' 'the projection reports unresolved transient evidence separately from known auxiliary rows'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'hiddenAuxiliaryRows > 0 || hiddenUnresolvedTransientRows > 0' 'a history containing only hidden auxiliary or unresolved transient evidence is not mislabeled as never tracked'
Assert-WorkflowToken 'project-element-change-ledger' $elementTrackingTransitionPolicyPath 'ignoredInSession' 'the transition policy covers a null live element already known by the session auxiliary index'
Assert-WorkflowToken 'project-element-change-ledger' $workflowHarnessPath 'element transition decisions execute add, modify, delete, null-element, and same-boundary transient cases' 'the executable harness covers element transition boundaries'
Assert-WorkflowToken 'project-element-change-ledger' $workflowHarnessPath 'all-project history summaries use the same projection as detail and Excel without mutating immutable evidence' 'the executable harness covers legacy-summary projection without checksum mutation'
Assert-WorkflowToken 'project-element-change-ledger' $elementChangeTrackingPath 'SharedParameterGuid = string.IsNullOrWhiteSpace(parameterGuid)' 'shared-parameter GUID and binding metadata enter the tracked state'
Assert-WorkflowToken 'project-element-change-ledger' $elementChangeTrackingPath 'GridCurveSignature = gridCurve' 'grid curve, extent, and pin state enter the tracked state'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'RefreshProjectParameterStatesAtCommit' 'Project/shared parameter state verification at commit' 'Save and Sync recompare parameter definitions when DocumentChanged omits binding-map IDs'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'RefreshProjectParameterStatesAtCommit' 'ParameterBindingsReadSucceeded' 'an unreadable parameter binding map becomes an explicit commit-boundary coverage gap'
Assert-WorkflowMethodOrder 'project-element-change-ledger' $elementChangeTrackingPath 'StageWorksharedLocalSaveCheckpoint' 'RefreshProjectParameterStatesAtCommit(session, doc, commitCaptureContext);' 'BuildCommit(doc, session' 'workshared local Save verifies parameter metadata before protecting its checkpoint'
Assert-WorkflowMethodOrder 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentCommitted' 'RefreshProjectParameterStatesAtCommit(session, doc, commitCaptureContext);' 'BuildCommit(doc, session' 'standalone Save and central Sync verify parameter metadata before commit construction'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'ProjectElementChangeSummaryLabel' 'ParameterBoundCategories' 'history and Excel render readable shared-parameter binding differences'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'ProjectElementChangeSummaryLabel' 'GridCurveSignature' 'history and Excel render readable grid geometry differences'
Assert-WorkflowMethodOrder 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentCommitted' 'TryGetIsWorkshared(doc, out workshared)' 'if (workshared && !string.Equals(kind, "SynchronizeWithCentral"' 'commit classification requires a known worksharing state before deciding whether local Save can publish'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentCommitted' 'no standalone or central commit was inferred' 'unknown worksharing state retains the session instead of publishing a guessed commit'
Assert-WorkflowMethodOrder 'project-element-change-ledger' $elementChangeTrackingPath 'CreateSession' 'TryGetIsWorkshared(doc, out workshared)' 'CaptureBaseline(doc' 'session startup rejects an unknown worksharing state before building an untrusted baseline'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'HasPendingRecoveryCheckpoint' 'recovery remains fail-closed' 'unknown worksharing state cannot suppress a possible protected checkpoint'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'IsWorkshared = workshared' 'committed metadata uses the already-validated worksharing state without a second fallible lookup'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementTrackingTransitionPolicyPath 'ResolveChangeKind' 'baselineCapturedLate && hasActiveActivity && wasDeleted' 'late-baseline unknown events are classified as deletion only when Revit reported deletion'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'First observed' 'element history exposes the client observation window instead of only the commit boundary'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'Windows clocks are synchronized' 'cross-PC chronology discloses its clock-synchronization dependency'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'other categoryless Revit internals are intentionally excluded' 'history UI distinguishes tracked parameter definitions from excluded Revit-internal metadata'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $projectCatalogDashboardPath 'ResolvePendingElementSessionCheckpointStatus' 'GetPendingElementSessionCheckpointStatus(_workspaceRoot);' 'home and header expose protected local-save checkpoints from every project on this PC'
Assert-WorkflowMethodExcludesToken 'element-tracking-session-recovery' $projectCatalogDashboardPath 'ResolvePendingElementSessionCheckpointStatus' 'ResolveProjectIdentityPath' 'home checkpoint warning is not hidden by the currently active project identity'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'Exact attribution requires the add-in on every editing workstation.' 'the attribution boundary is explicit in every commit'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'ClientObservedWithIdentityGap' 'missing Revit-user identity cannot be presented as fully attributed client observation'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'exact Revit-user attribution requires review' 'commit evidence discloses the Windows-user-only attribution boundary'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'identityGapCommits' 'history summary visibly counts records without a Revit username'
Assert-WorkflowMethodToken 'tracking-policy-concurrency' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'Disabled remotely / protected through commit' 'deferred remote-disable commits are not mislabeled as live-enabled policy checks'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentCommitted' 'finalStateRefreshIds.UnionWith(incomingIds);' 'successful commit refreshes only locally pending and incoming changed elements'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentCommitted' 'current.PromoteCurrentToBaseline(stateRefreshElapsedMilliseconds);' 'a clean successful commit promotes the maintained state without a full-model recapture'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentCommitted' 'FamilyBrowserPostCommitBaselineMode.FullCapture' 'uncertain commit evidence retains a conservative full-model fallback'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentCommitted' 'SetProjectCatalogObservationDecisionNoLock(runtimeKey, projectCatalogObservationRequired);' 'commit analysis records whether a project catalog scan is actually needed'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'WriteTrackingPerformance' 'element-tracking-performance.log' 'element tracking exposes commit stage timing for field diagnosis'
Assert-WorkflowToken 'project-element-change-ledger' $trackingCommitOptimizationPath 'ShouldObserveProjectCatalog' 'catalog scan policy is shared with executable workflow tests'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentCommitted' 'session.ExternalOverlapIds.Add(id);' 'same-element incoming update overlap is retained instead of falsely claiming sole authorship'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'RebaseSessionAfterExternalUpdate' 'GetLocallyPendingElementIds(session)' 'external-update rebase preserves active local activity rather than every historically touched element'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'RebaseSessionAfterExternalUpdate' 'GetRecoveredCheckpointElementIds(session)' 'incoming updates overlapping restart-recovered local saves remain marked as mixed authorship'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'RebaseSessionAfterExternalUpdate' 'current.ExternalRebaseFailed = true;' 'a failed incoming-update rebase marks the session as uncertain'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentSynchronizationCompletionFailure' 'CloseExternalUpdateWindowAfterUnknownCompletion(' 'an unreadable synchronization completion closes its external-update window conservatively'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'CloseExternalUpdateWindowAfterUnknownCompletion' 'activeKeys.Remove(runtimeKey);' 'unknown completion cannot leave a permanent external-update suppression key'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'CloseExternalUpdateWindowAfterUnknownCompletion' 'session.ExternalRebaseFailed = true;' 'unknown completion records an attribution/rebase confidence gap'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'CloseExternalUpdateWindowAfterUnknownCompletion' 'session.CommitBoundaryProtectionFailed = HasUncommittedSessionEvidence(session);' 'an unreadable synchronization completion is immediately exposed as unsafe to close'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleReloadedLatestBridge' 'CloseExternalUpdateWindowAfterUnknownCompletion(' 'Reload Latest completion callback failure closes its suppression window'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentChanged' 'TryGetElementIds' 'DocumentChanged element-ID getter failures are not silently converted to an empty change set'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentChanged' 'active session(s) were conservatively marked for coverage review' 'an unreadable DocumentChanged document cannot disappear as a trusted no-op'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentChanged' 'TryCanTrackDocument' 'a document-kind API failure is distinct from a confirmed Family document'
Assert-WorkflowMethodExcludesToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentChanged' 'if (!CanTrackDocument(doc))' 'a transient document-kind failure cannot remove an active tracking session'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentSynchronizingWithCentral' 'session.ExternalRebaseFailed = true;' 'unknown document kind at synchronization start retains suppression and lowers attribution confidence'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentReloadingLatest' 'session.ExternalRebaseFailed = true;' 'unknown document kind at Reload Latest start retains suppression and lowers attribution confidence'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentChanged' 'RecoverActivityFromCurrentSnapshot' 'DocumentChanged ID gaps attempt a full current-state comparison'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'CreateActivity' 'UnknownOperation' 'an unreadable DocumentChanged operation is not guessed as a committed transaction'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'CreateActivity' 'metadataError = new AggregateException' 'transaction-name and operation read failures lower event coverage confidence'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'ClientObservedWithEventReadGap' 'a recovered DocumentChanged ID gap lowers attribution confidence'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'CoverageGapOnly = coverageGapOnly' 'an unidentified observed event becomes an explicit durable coverage-gap record'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'without inventing an element identity' 'coverage-gap history discloses that no synthetic element identity was created'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'BuildCommit' 'session.BaselineCapturedLate && !HasProtectedRecoveryEvidence(session)' 'checkpoint recovery does not fabricate a second coverage-only commit when no new gap occurred'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HasUncommittedSessionEvidence' 'session.EventReadFailureCount > 0' 'an unreadable event remains uncommitted evidence even when no element ID was recovered'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentSaveCompletionFailure' 'session.CommitBoundaryReadFailureCount++' 'an unreadable Save completion boundary remains visible in the next durable commit'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentSaveCompletionFailure' 'session.CommitBoundaryProtectionFailed = HasUncommittedSessionEvidence(session);' 'an unreadable Save boundary is immediately exposed as unsafe to close'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'ClientObservedWithCommitBoundaryGap' 'a commit spanning an unreadable Save boundary cannot claim exact boundary attribution'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'ApplyActivity' 'matchedActivity.LastObservedAtUtc' 'an exact Redo advances final observation time without discarding the original first observation time'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'BuildCommit' 'activity.LastObservedAtUtc' 'history rows use the Redo-aware final observation time'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'BuildCommit' 'OrderByDescending' 'Redo reordering cannot make an older callback look like the final observation'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'Commits with incomplete DocumentChanged event reads:' 'history surfaces DocumentChanged ID and operation-metadata observation gaps'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'Commits spanning an unreadable Save boundary:' 'history surfaces incomplete Save and Save As completion boundaries'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'Coverage-gap records without a trustworthy element ID:' 'history visibly counts commits whose element identity could not be recovered'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'ProjectElementAttributionLabel' 'Element ID unavailable' 'coverage-only commits produce a readable result and Excel row'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'AppendCommitCanonicalV4' 'commit.EventReadFailureCount' 'the frozen integrity-v4 canonical form still protects legacy event-gap metadata'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'AppendCommitCanonicalV5' 'AppendChangeExtensionCanonicalV5' 'integrity-v5 extends the protected record with parameter and grid metadata'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'ValidateElementChangeCommit' 'Schema-v6 element history requires integrity-v5 evidence.' 'schema-v6 parameter and grid records cannot be accepted as unsigned legacy history'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'SaveElementSessionCheckpoint' 'HasElementChangesOrCoverageGap' 'checkpoint writes retain coverage-only evidence and reject malformed empty commits'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'NotifyManagementContextChanged' '_cachedPolicyWorkspaceRoot = string.Empty;' 'managed-path changes invalidate the policy cache'
Assert-WorkflowMethodExcludesToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'NotifyManagementContextChanged' 'Sessions.Clear();' 'managed-path changes cannot erase live or checkpoint-recovered tracking sessions'
Assert-WorkflowMethodToken 'managed-folder-homepage-return' $managedFolderTransitionPath 'SwitchToHomepageManagedFolder' 'AuthorizeManagedFolderTransition()' 'only the verified homepage migration controller authorizes protected-checkpoint rebinding'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'localSyncPendingStatus.LockUnavailable' 'history distinguishes checkpoint lock contention from checksum corruption'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $projectCatalogDashboardPath 'AppendProjectCatalogPill' 'localSyncPendingStatus.LockUnavailable ? 0 : ResolveInvalidElementSessionCheckpointCount()' 'header does not mislabel a busy checkpoint as corrupt'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $projectCatalogDashboardPath 'AppendHomePendingTrackingQueueBoard' 'localSyncPendingStatus.LockUnavailable ? 0 : ResolveInvalidElementSessionCheckpointCount()' 'Home does not duplicate a busy checkpoint as a corruption alert'
Assert-WorkflowMethodExcludesToken 'project-element-change-ledger' $elementChangeTrackingPath 'RebaseSessionAfterExternalUpdate' 'EndDocumentSession(doc);' 'a transient incoming-update rebase failure cannot erase pending local activity'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'ClientObservedWithExternalRebaseGap' 'a retained session exposes failed external rebase coverage in attribution confidence'
Assert-WorkflowToken 'project-element-change-ledger' $elementChangeTrackingPath 'usedCachedFallback = true;' 'temporary policy read failure preserves the last confirmed tracking state'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'ResolvePolicyEnabledCore' '_cachedPolicyKnown' 'policy-read fallback distinguishes a deliberately stale-but-confirmed state from an uninitialized cache'
Assert-WorkflowMethodTokenAfter 'project-element-change-ledger' $elementChangeTrackingPath 'ResolvePolicyEnabledCore' 'if (string.Equals(root, _cachedPolicyWorkspaceRoot, StringComparison.OrdinalIgnoreCase) && _cachedPolicyKnown)' 'return _cachedPolicyEnabled;' 'Save or Sync policy-read failure retains the last confirmed deferred-disable state instead of discarding evidence'
Assert-WorkflowMethodToken 'tracking-policy-concurrency' $elementChangeTrackingPath 'NotifyPolicyChanged' 'EnableLiveTrackingNoLock(workspaceRoot);' 'locally re-enabling tracking promotes an already-open recovery-only checkpoint session'
Assert-WorkflowMethodToken 'tracking-policy-concurrency' $elementChangeTrackingPath 'ResolvePolicyEnabledCore' 'EnableLiveTrackingNoLock(root);' 'remotely re-enabling tracking promotes an already-open recovery-only checkpoint session'
Assert-WorkflowMethodToken 'tracking-policy-concurrency' $elementChangeTrackingPath 'EnableLiveTrackingNoLock' 'session.RecoveryOnly = false;' 're-enabled recovery sessions resume normal observation and commit behavior'
Assert-WorkflowMethodToken 'tracking-policy-concurrency' $elementChangeTrackingPath 'EnableLiveTrackingNoLock' 'session.PolicyDisableDeferred = false;' 'a later re-enable clears stale deferred-disable attribution state'
Assert-WorkflowToken 'tracking-policy-session-isolation' $elementTrackingPolicyDecisionPath 'policyEnabled && !policyStateIsFallbackOrDeferred' 'only an authoritative enabled policy permits live collection'
Assert-WorkflowMethodToken 'tracking-policy-session-isolation' $elementChangeTrackingPath 'BeginDocumentSession' 'FamilyBrowserElementTrackingPolicyDecision.Resolve' 'new document sessions use the fail-closed policy decision'
Assert-WorkflowMethodToken 'tracking-policy-session-isolation' $elementChangeTrackingPath 'HandleDocumentChanged' 'FamilyBrowserElementTrackingPolicyDecision.Resolve' 'DocumentChanged rechecks session eligibility before collecting evidence'
Assert-WorkflowMethodToken 'tracking-policy-session-isolation' $elementChangeTrackingPath 'DisableLiveTrackingNoLock' 'preserveUncommittedEvidence && HasUncommittedSessionEvidence(session)' 'remote disable preserves only sessions that already own evidence'
Assert-WorkflowMethodToken 'tracking-policy-session-isolation' $elementChangeTrackingPath 'DisableLiveTrackingNoLock' 'string.Equals(item.WorkspaceRoot ?? string.Empty, root' 'policy disable cleanup is scoped to the affected management context'
Assert-WorkflowToken 'tracking-policy-session-isolation' $workflowHarnessPath 'policy disable and read-fallback preserve existing evidence without starting collection in another document' 'workflow harness covers cross-document policy isolation'
Assert-WorkflowMethodToken 'project-element-change-ledger' $trackingPath 'PersistElementChangeCommits' 'BuildElementChangeSpoolPath' 'unavailable managed storage falls back to a local element-change spool'
Assert-WorkflowMethodToken 'project-element-change-ledger' $trackingPath 'PersistElementChangeCommitsDeferred' 'TryWriteJsonAtomic(BuildElementChangeSpoolPath(commit.EntryId), record)' 'Save and Sync completion first persists element history to the local durable spool'
Assert-WorkflowMethodExcludesToken 'project-element-change-ledger' $trackingPath 'PersistElementChangeCommitsDeferred' 'TryWriteElementHistoryAtomic' 'Save and Sync completion does not synchronously write element history to the managed network folder'
Assert-WorkflowMethodExcludesToken 'project-element-change-ledger' $trackingPath 'PersistElementChangeCommitsDeferred' 'BuildManagedDestinationIdentity(workspaceRoot)' 'Save and Sync completion does not probe the managed path while creating the local spool'
Assert-WorkflowMethodToken 'project-element-change-ledger' $trackingPath 'PersistElementChangeCommitsDeferred' 'BuildDeferredManagedDestinationIdentity(deferredDestinationPath)' 'the commit callback records a non-probing destination identity from local configuration text'
Assert-WorkflowMethodToken 'project-element-change-ledger' $trackingPath 'QueuePendingFlush' 'ThreadPool.QueueUserWorkItem' 'managed history publication is queued after the Revit commit callback'
Assert-WorkflowMethodToken 'project-element-change-ledger' $trackingPath 'QueuePendingFlush' 'PromoteDeferredElementSpoolDestinations(safeDeferredDestinationPath)' 'managed path canonicalization happens inside the deferred worker'
Assert-WorkflowMethodToken 'project-element-change-ledger' $trackingPath 'PromoteDeferredElementSpoolDestinations' 'BuildStableManagedDestinationIdentity(deferredDestinationPath)' 'the worker canonicalizes the captured destination before managed publication'
Assert-WorkflowMethodToken 'project-element-change-ledger' $trackingPath 'PromoteDeferredElementSpoolDestinations' 'TryWriteMutableJsonAtomic(spoolPath, record)' 'destination promotion atomically replaces the protected mutable spool envelope'
Assert-WorkflowMethodToken 'project-element-change-ledger' $trackingPath 'FlushPending' 'PromoteDeferredElementSpoolDestinations(ResolveDeferredManagedDestinationPath(workspaceRoot))' 'startup and later retry recover a spool written before the deferred worker ran'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentCommitted' 'PersistElementChangeCommitsDeferred' 'element commit completion uses deferred managed publication'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'BuildChangeActivityIndex(session.AppliedActivities)' 'element commit builds one activity index before classifying changed IDs'
Assert-WorkflowMethodExcludesToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'session.AppliedActivities.Any' 'element commit no longer scans every activity for every changed ID'
Assert-WorkflowMethodExcludesToken 'project-element-change-ledger' $elementChangeTrackingPath 'BuildCommit' 'session.AppliedActivities.Where' 'element commit reuses indexed per-element activity lists'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'HandleDocumentCommitted' 'StateCaptureContext commitCaptureContext = BuildStateCaptureContext(doc);' 'Save and Sync completion builds the project-parameter capture context once'
Assert-WorkflowMethodToken 'project-element-change-ledger' $elementChangeTrackingPath 'RefreshProjectParameterStatesAtCommit' 'captureContext.ParameterElements' 'parameter verification reuses the parameter elements collected by the shared context'
Assert-WorkflowToken 'project-element-change-ledger' $trackingPath 'ElementChangeHistory' 'managed element changes are partitioned into immutable history'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'LoadProjectElementHistoryBundle' 'LoadImmutableElementChangeCommitResult' 'administrator can review validated immutable project element history'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'ShowSelectedProjectElementChangeHistory' 'Task.Run' 'selected-project history file reads run outside the browser UI thread'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'ShowAllProjectElementChangeHistory' 'Task.Run' 'all-project history discovery runs outside the browser UI thread'
Assert-WorkflowMethodExcludesToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'FamilyBrowserTrackingPersistenceService.' 'history rendering does not return to synchronous disk scans on the browser UI thread'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'Recent changes' 'recent element rows are visible before Excel export'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'externalRebaseGapCommits' 'history summary and warning state expose incomplete external-update rebases'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'ProjectElementAttributionLabel' 'External update coverage gap' 'history rows never present an external rebase gap as ordinary client attribution'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'FamilyBrowserResultExcelExportUi.SaveRows' 'element history workbook is created only by explicit export'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'TrackedProjectElementHistoryHtmlForm' 'element history opens in the dedicated searchable table dialog'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'T("Time", "시각")' 'element history puts time first for rapid incident review'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'T("Element ID", "요소 ID")' 'element history exposes the deleted object identifier'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'change.FamilyName' 'element history exposes Family names for deleted objects'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'change.TypeName' 'element history exposes Type names for deleted objects'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'ShowAllProjectElementChangeHistory' 'LoadTrackedProjectHistorySummaries' 'administrator can select any tracked project from one history browser'
Assert-WorkflowMethodToken 'project-element-change-ledger' $projectCatalogDashboardPath 'ShowAllProjectElementChangeHistory' 'activeProjectTrackingEnabled' 'all-project history does not synthesize an unregistered active RVT'
Assert-WorkflowMinimumOccurrences 'project-element-change-ledger' $projectCatalogDashboardPath 'Only RVT files registered in Permissions / Guard with Element Change Tracking enabled are recorded.' 2 'current and all-project history windows disclose the per-RVT recording scope'
Assert-WorkflowMethodToken 'project-element-change-ledger' $trackingPath 'LoadTrackedProjectHistorySummaries' 'LoadAllImmutableElementChangeCommits' 'all-project history includes confirmed managed records'
Assert-WorkflowMethodToken 'project-element-change-ledger' $trackingPath 'LoadTrackedProjectHistorySummaries' 'LoadPendingElementChangeCommitResult' 'all-project history includes protected local upload records'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $trackingPath 'LoadElementSessionCheckpointHistory' 'SynchronizationSucceededEntryIds' 'local Save checkpoints retain their synchronization promotion status for history review'

Assert-WorkflowToken 'element-tracking-integrity-recovery' $trackingPath 'EnvelopeIntegritySha256' 'pending element destination metadata carries integrity evidence'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'ValidatePendingOperationEnvelope' 'ComputePendingOperationEnvelopeIntegrity(record)' 'pending operation destination and payload metadata are checksum validated'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'ValidatePendingStandardCandidateEnvelope' 'ComputePendingStandardCandidateEnvelopeIntegrity(record)' 'pending standard-candidate destination and payload metadata are checksum validated'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'FlushPendingNoLock' 'TryDeleteChecked(path)' 'pending flush reports local cleanup failure instead of claiming a settled transition'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'FlushPendingNoLock' 'EnumerateSpoolFilesOrMarkFailure' 'pending spool enumeration failures are surfaced instead of looking like an empty queue'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $trackingPath 'GetPendingElementSessionCheckpointStatus' 'result.LockUnavailable = true;' 'checkpoint enumeration failures remain a blocking unknown state'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $trackingPath 'HasBlockingElementSessionCheckpointForManagedPolicyPath' 'TryEnumerateSpoolFiles' 'managed-folder replacement fails closed when checkpoint enumeration is unavailable'
Assert-WorkflowMethodToken 'project-element-change-ledger' $trackingPath 'LoadImmutableElementChangeCommitResult' 'TryEnumerateDirectories(historyRoot' 'immutable-history root enumeration failures are explicitly counted'
Assert-WorkflowMethodToken 'project-element-change-ledger' $trackingPath 'LoadImmutableElementChangeCommitResult' 'result.InvalidRecordCount++;' 'immutable-history enumeration failure cannot appear as clean empty history'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'BuildStableManagedDestinationIdentity' 'GetCanonicalPath(path)' 'managed destination binding uses a stable path rather than a replaceable policy file identity'
Assert-WorkflowMethodOrder 'project-element-change-ledger' $revisionPath 'GetCanonicalPath' 'CanonicalPathCache.TryGetValue(normalized, out cached)' 'File.Exists(normalized)' 'repeated Save and Sync path identity checks reuse an already-verified canonical path before probing storage again'
Assert-WorkflowMethodToken 'project-element-change-ledger' $revisionPath 'GetCanonicalPath' 'Directory.Exists(normalized)' 'managed data-folder identities can be canonicalized and cached as directories'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'FlushPendingForManagedFolderTransition' 'BuildManagedDestinationIdentityForRoot(sourceManagedRoot)' 'managed-folder transition is restricted to the selected source root'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'RebindElementSessionCheckpointsNoLock' 'ValidateElementSessionCheckpointCommits(checkpoint)' 'managed-folder migration never rebinds a checkpoint with invalid inner commits'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'DeleteElementSessionCheckpointsForDestination' 'ValidateElementSessionCheckpointCommits(checkpoint)' 'checkpoint cleanup never deletes evidence whose inner records fail validation'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'DeleteElementSessionCheckpoint' 'ValidateElementSessionCheckpointCommits(checkpoint)' 'single-checkpoint cleanup never deletes corrupt inner evidence'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'SaveElementSessionCheckpoint' 'ValidateElementSessionCheckpointCommits(existing)' 'a new local Save cannot overwrite an invalid or foreign checkpoint'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'GetPendingElementSessionCheckpointStatus' 'ValidateElementSessionCheckpointCommits(checkpoint)' 'invalid checkpoints are not counted as trusted pending work'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'GetMismatchedElementSessionCheckpointCount' 'ValidateElementSessionCheckpointCommits(checkpoint)' 'management-folder mismatch status counts only valid protected checkpoints'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'GetMismatchedElementSessionCheckpointCount' '!CanFlushToDestination(checkpoint.DestinationIdentity' 'valid checkpoints bound to another management destination remain explicitly visible'
Assert-WorkflowMethodOrder 'element-tracking-integrity-recovery' $trackingPath 'LoadElementSessionCheckpoint' 'ValidateElementSessionCheckpointCommits(checkpoint)' 'CanFlushToDestination(checkpoint.DestinationIdentity' 'inner corruption is reported before a management-destination mismatch'
Assert-WorkflowMethodOrder 'element-tracking-integrity-recovery' $trackingPath 'SaveElementSessionCheckpoint' 'EnsureEntryId(commit);' '.GroupBy(delegate(FamilyBrowserElementChangeCommit commit)' 'checkpoint commit identities are created before deduplication'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'SaveElementSessionCheckpoint' '!FixedTimeEquals(candidate.IntegritySha256, first.IntegritySha256)' 'conflicting records with one checkpoint entry ID fail closed'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'ValidateElementSessionCheckpointCommits' 'string.IsNullOrWhiteSpace(commit.IntegritySha256)' 'checkpoint inner records must be signed even when the outer envelope is valid'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'ValidateElementSessionCheckpointCommits' 'commits.Count == 0' 'empty checkpoint envelopes are invalid rather than trusted pending work'
Assert-WorkflowToken 'element-tracking-integrity-recovery' $trackingPath 'FileMatchesPayload(path, payload)' 'an existing destination must exactly match before local evidence is deleted'
Assert-WorkflowToken 'element-tracking-integrity-recovery' $trackingPath 'Element history checksum mismatch.' 'tampered element history is rejected'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'LoadImmutableElementChangeCommitResult' 'ParseUtc(commit.CommittedAtUtc)' 'recent history is ordered by committed time'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'PendingCorruptRecordCount' 'history viewer exposes corrupt pending records'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $trackingPath 'LoadImmutableElementChangeCommitResult' 'result.TotalValidRecordCount = loaded.Count;' 'history load reports the full valid-record count before applying the display limit'
Assert-WorkflowMethodToken 'element-tracking-integrity-recovery' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'commitHistoryTruncated || rowHistoryTruncated' 'history display warns when its commit or row limit hides older evidence'

Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'StageWorksharedLocalSaveCheckpoint' 'SaveElementSessionCheckpoint' 'successful workshared local Save creates a restart-safe checkpoint'
Assert-WorkflowMethodTokenAfter 'element-tracking-session-recovery' $elementChangeTrackingPath 'StageWorksharedLocalSaveCheckpoint' 'if (!protectedLocally)' 'session.PromoteCurrentToBaseline(refreshElapsedMilliseconds);' 'a local-save boundary promotes live activity only after its checkpoint is durable'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'StageWorksharedLocalSaveCheckpoint' 'session.LocalSaveCheckpointFailed = currentCommit != null && HasUncommittedSessionEvidence(session);' 'a failed local-save checkpoint remains visible as an unsafe open-session state'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'StageWorksharedLocalSaveCheckpoint' 'catch (Exception checkpointError)' 'an unexpected checkpoint writer exception cannot bypass the unsafe open-session state'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'ResetBaseline' 'LocalSaveCheckpointFailed = false;' 'the unsafe local-save warning clears only after a durable boundary resets the baseline'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'GetUnprotectedLocalSaveSessionCount' 'session.LocalSaveCheckpointFailed && HasUncommittedSessionEvidence(session)' 'the dashboard counts only open sessions whose live evidence was not checkpointed'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'GetUnprotectedCommitBoundarySessionCount' 'session.CommitBoundaryProtectionFailed && HasUncommittedSessionEvidence(session)' 'the dashboard exposes observed evidence that did not reach a restart-safe Save or Sync boundary'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'StageWorksharedLocalSaveCheckpoint' 'session.RecoveredLocalSaveCommits.AddRange(staged.Where' 'same-session synchronization publishes the exact locally protected save intervals'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'CreateSession' 'RecoveredLocalSaveCommits.AddRange' 'reopening the same local file restores its pending local-save commits'
Assert-WorkflowMethodOrder 'element-tracking-session-recovery' $elementChangeTrackingPath 'HandleDocumentCommitted' 'SaveElementSessionCheckpoint' 'PersistElementChangeCommits' 'successful synchronization finalizes the write-ahead checkpoint before immutable persistence'
Assert-WorkflowMethodTokenAfter 'element-tracking-session-recovery' $elementChangeTrackingPath 'HandleDocumentCommitted' 'synchronization checkpoint finalization failed' 'return false;' 'failed checkpoint finalization stops immutable publication'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'HandleDocumentCommitted' 'DeleteElementSessionCheckpoint' 'successful synchronization removes the checkpoint only after durable persistence'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'HandleDocumentCommitted' 'recoveredCommit.PublishedAtUtc = synchronizationPublishedAtUtc;' 'recovered local-save work is published at the first successful central synchronization'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'HandleDocumentCommitted' 'recoveredCommit.LocalSaveProtectedAtUtc' 'local checkpoint protection time remains separate from central publication time'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'HandleDocumentCommitted' 'CommitMatchesProjectStableIdentity' 'a synchronization cannot publish a recovered checkpoint from another project identity'
Assert-WorkflowMethodExcludesToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'HandleDocumentCommitted' '.GroupBy(delegate(FamilyBrowserElementChangeCommit item)' 'synchronization does not discard conflicting recovered entry IDs before checkpoint validation'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'StageWorksharedLocalSaveCheckpoint' 'CommitMatchesProjectStableIdentity' 'Save As cannot copy a previous-project checkpoint into the new project checkpoint'
Assert-WorkflowMethodExcludesToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'StageWorksharedLocalSaveCheckpoint' '.GroupBy(delegate(FamilyBrowserElementChangeCommit item)' 'local Save cannot discard conflicting recovered entry IDs before checkpoint validation'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'StageWorksharedLocalSaveCheckpoint' 'SameCheckpointIdentity' 'Save As detects a changed local-file checkpoint identity'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'StageWorksharedLocalSaveCheckpoint' 'previous-path local-save checkpoint cleanup failed' 'Save As retains a diagnostic if its obsolete checkpoint cannot be removed'
Assert-WorkflowMethodExcludesToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'NotifyPolicyChanged' 'DeleteElementSessionCheckpointsForDestination' 'disabling tracking never deletes already-protected local-save evidence'
Assert-WorkflowMethodExcludesToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'ResolvePolicyEnabled' 'DeleteElementSessionCheckpointsForDestination' 'an observed disabled policy never performs implicit checkpoint deletion'
Assert-WorkflowMinimumOccurrences 'element-tracking-session-recovery' $elementChangeTrackingPath 'session.RecoveredLocalSaveCommits.AddRange' 2 'failed finalized-checkpoint replay remains attached to the session for retry'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'HandleDocumentCommitted' '.Select(CloneElementChangeCommit)' 'a failed checkpoint finalization cannot mutate recovered publication timestamps in memory'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'StageWorksharedLocalSaveCheckpoint' 'expectedCheckpointRevisionToken' 'local Save replaces only the exact checkpoint evidence revision observed by this Revit session'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'HandleDocumentCommitted' 'finalizedCheckpointRevisionToken' 'successful synchronization deletes only the checkpoint evidence revision it finalized'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'BuildRuntimeKey' 'RuntimeDocumentIdentities.GetValue(doc, CreateRuntimeDocumentIdentity)' 'same managed document wrapper retains its runtime session identity across Save As boundaries'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'CreateRuntimeDocumentIdentity' 'FamilyBrowserPathIdentityService.GetStablePathIdentity(SafeLocalDocumentPath(doc))' 'different Revit Document wrappers for the same saved local RVT converge on one tracking session key'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'CreateRuntimeDocumentIdentity' 'FamilyBrowserPathIdentityService.GetStablePathIdentity(SafeProjectIdentityPath(doc))' 'project or central identity remains a stable fallback when the local document path is unavailable'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'CreateRuntimeDocumentIdentity' 'document-unsaved:' 'unsaved documents retain a deterministic callback-stable session identity'
Assert-WorkflowMethodExcludesToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'CreateRuntimeDocumentIdentity' 'new RuntimeDocumentIdentity()' 'callback wrappers never receive unrelated random document identities while stable identity data is available'
Assert-WorkflowToken 'element-tracking-session-recovery' $workflowHarnessPath 'Save As keeps old and new local checkpoint identities isolated during conditional cleanup' 'workflow harness proves Save As cleanup cannot delete the new-path checkpoint'
Assert-WorkflowToken 'element-tracking-session-recovery' $projectCatalogDashboardPath 'pendingElementSessionQueue' 'Home explains local-save activity that still requires synchronization'
Assert-WorkflowToken 'element-tracking-session-recovery' $projectCatalogDashboardPath 'finalizedElementSessionPill' 'the header distinguishes a successful synchronization awaiting immutable-history promotion'
Assert-WorkflowToken 'element-tracking-session-recovery' $projectCatalogDashboardPath 'finalizedElementSessionQueue' 'Home explains that finalized checkpoint replay does not require another synchronization'
Assert-WorkflowToken 'element-tracking-session-recovery' $projectCatalogDashboardPath 'unprotectedLocalSavePill' 'the header warns immediately when a successful local Save did not protect its tracking evidence'
Assert-WorkflowToken 'element-tracking-session-recovery' $projectCatalogDashboardPath 'unprotectedLocalSaveQueue' 'Home tells the user not to close Revit until unprotected local-save evidence reaches a safe boundary'
Assert-WorkflowToken 'element-tracking-session-recovery' $projectCatalogDashboardPath 'unprotectedCommitBoundaryPill' 'the header warns when a Save or Sync boundary could not protect observed evidence'
Assert-WorkflowToken 'element-tracking-session-recovery' $projectCatalogDashboardPath 'unprotectedCommitBoundaryQueue' 'Home tells the user not to close Revit after an unreadable or failed protection boundary'

Assert-WorkflowToken 'element-tracking-schema-compatibility' $trackingPath 'commit.IntegrityVersion = commit.SchemaVersion >= 6 ? 5 : (commit.SchemaVersion >= 5 ? 4 : 3);' 'new schema-v6 history uses integrity-v5 while schema-v5 and older records retain their frozen verifiers'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'ComputeElementChangeIntegrity' 'ComputeElementChangeIntegrityV1(commit)' 'integrity-v1 records retain a frozen compatibility verifier'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'ComputeElementChangeIntegrity' 'AppendCommitCanonicalV2(canonical, commit);' 'integrity-v2 records retain their frozen canonical verifier'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'AppendCommitCanonicalV3' 'commit.PublishedAtUtc' 'integrity-v3 protects publication timing'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'EnsurePendingElementEnvelopeIntegrity' 'record.EnvelopeIntegrityVersion = 2;' 'new pending envelopes do not serialize the evolving full commit model'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'ComputePendingElementEnvelopeIntegrity' 'CreatePendingElementEnvelopeIntegrityV1EntryPayload' 'pending-envelope-v1 records retain their exact frozen compatibility schema'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'EnsureElementChangeIntegrity' 'return FixedTimeEquals(commit.IntegritySha256, ComputeElementChangeIntegrity(commit));' 'valid old checksums are preserved and invalid signed records are not silently re-signed'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'ValidateElementChangeCommit' 'MaximumSupportedElementCommitSchemaVersion' 'future element-history schemas are rejected instead of partially interpreted'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'TryWriteElementHistoryAtomic' 'ElementCommitMatchesSameProject' 'immutable replay checks entry identity and project binding before accepting an existing file'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'ValidateElementSessionCheckpointCommits' 'ElementCommitMatchesCheckpointProject' 'local checkpoints cannot contain commits from another project'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'LoadImmutableElementChangeCommitResult' 'allHistoryRoots' 'history lookup can recover records stored under a replaced legacy file identity'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'LoadElementSessionCheckpoint' 'result.Invalid = File.Exists(checkpointPath);' 'a corrupt checkpoint is not silently reported as absent'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'GetInvalidElementSessionCheckpointCount' 'ValidateElementSessionCheckpointCommits' 'local checkpoint health validates both envelope and commit checksums'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'SaveElementSessionCheckpoint' 'TryAcquireElementSessionFileLock' 'checkpoint writers are serialized across Revit processes before compare-and-swap'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'SaveElementSessionCheckpoint' 'FixedTimeEquals(ComputeElementSessionCheckpointRevisionToken(existing), expectedCheckpointRevisionToken)' 'a stale checkpoint evidence revision cannot overwrite newer local evidence'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'DeleteElementSessionCheckpoint' 'FixedTimeEquals(ComputeElementSessionCheckpointRevisionToken(checkpoint), expectedCheckpointRevisionToken)' 'stale cleanup cannot delete a newer checkpoint evidence revision'
Assert-WorkflowMethodExcludesToken 'element-tracking-schema-compatibility' $trackingPath 'ComputeElementSessionCheckpointRevisionToken' 'DestinationIdentity' 'explicit management-folder migration does not invalidate an unchanged checkpoint evidence revision'
Assert-WorkflowMethodExcludesToken 'element-tracking-schema-compatibility' $trackingPath 'ComputeElementSessionCheckpointRevisionToken' 'UpdatedAtUtc' 'checkpoint timestamp refresh does not masquerade as an evidence conflict'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'LoadElementSessionCheckpoint' 'result.LockUnavailable = true;' 'checkpoint lock contention is explicit instead of looking like no pending work'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $trackingPath 'GetPendingElementSessionCheckpointStatus' 'result.LockUnavailable = true;' 'checkpoint count exposes lock contention separately from a trusted zero count'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $trackingPath 'GetPendingElementSessionCheckpointStatus' 'result.SynchronizationSucceededCount++;' 'checkpoint status distinguishes finalized synchronization evidence from local saves still awaiting sync'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $trackingPath 'FlushPending' 'FlushFinalizedElementSessionCheckpointsNoLock' 'ordinary refresh retries finalized checkpoint promotion without another Revit synchronization'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $trackingPath 'DeleteElementSessionCheckpointById' 'expectedCheckpointRevisionToken' 'finalized checkpoint promotion deletes only the exact revision that was persisted'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'HasPendingRecoveryCheckpoint' 'return true;' 'checkpoint lookup failures keep recovery fail-closed instead of suppressing it'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'ValidateElementSessionCheckpointCommits' '!entryIds.Add(commit.EntryId)' 'duplicate fully signed checkpoint entry IDs fail closed during recovery'
Assert-WorkflowMethodToken 'element-tracking-schema-compatibility' $trackingPath 'LoadImmutableElementChangeCommitResult' 'conflictingEntryIds' 'conflicting immutable entry IDs across identity roots are quarantined instead of first-writer-wins display'
Assert-WorkflowToken 'element-tracking-schema-compatibility' $workflowHarnessPath 'stale checkpoint writers and cleanup operations fail closed across Revit processes' 'workflow harness proves checkpoint compare-and-swap and conditional cleanup'
Assert-WorkflowToken 'element-tracking-schema-compatibility' $workflowHarnessPath 'checkpoint load and count lock contention are explicit and never treated as trusted zero pending work' 'workflow harness proves checkpoint lock contention is fail-closed'
Assert-WorkflowToken 'element-tracking-schema-compatibility' $workflowHarnessPath 'management-folder checkpoint migration lock failures remain explicit, unchanged, and retryable' 'workflow harness proves a locked migration retains evidence and can be retried after lock release'
Assert-WorkflowToken 'element-tracking-schema-compatibility' $workflowHarnessPath 'live sessions continue with the same evidence revision' 'workflow harness proves explicit managed-folder migration does not strand a live Revit tracking session'
Assert-WorkflowToken 'element-tracking-integrity-recovery' $workflowHarnessPath 'management-folder checkpoint mismatches remain visible until explicit migration' 'workflow harness proves a valid destination mismatch is visible before migration and clears after migration'
Assert-WorkflowToken 'element-tracking-schema-compatibility' $workflowHarnessPath 'conflicting immutable entry IDs across legacy and current identity roots are quarantined together' 'workflow harness proves cross-root immutable collision quarantine'
Assert-WorkflowToken 'element-tracking-schema-compatibility' $workflowHarnessPath 'same-name project files remain isolated by stable path identity' 'workflow harness proves two same-name project files cannot mix immutable histories'
Assert-WorkflowToken 'element-tracking-schema-compatibility' $workflowHarnessPath "central identity changes retain the previous project's unsynchronized checkpoint independently" 'workflow harness proves a new central identity cannot consume or delete the old project checkpoint'
Assert-WorkflowToken 'element-tracking-session-recovery' $workflowHarnessPath 'global local-save status includes protected checkpoints from every project identity' 'workflow harness proves the home warning includes retained checkpoints outside the active project'
Assert-WorkflowToken 'element-tracking-schema-compatibility' $projectCatalogDashboardPath 'invalidElementSessionPill' 'corrupt local checkpoint health is visible in the header'
Assert-WorkflowToken 'element-tracking-schema-compatibility' $projectCatalogDashboardPath 'invalidElementSessionQueue' 'Home explains how to preserve and recover a corrupt local checkpoint'
Assert-WorkflowToken 'element-tracking-integrity-recovery' $projectCatalogDashboardPath 'mismatchedElementSessionPill' 'header exposes valid checkpoints still bound to another management folder'
Assert-WorkflowToken 'element-tracking-integrity-recovery' $projectCatalogDashboardPath 'mismatchedElementSessionQueue' 'Home explains explicit migration for destination-bound local evidence'

Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'BuildCommit' 'CreatedThenDeleted' 'create-then-delete activity remains visible at the successful boundary'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'CreateSession' 'capturedLate || SafeIsModified(doc)' 'enabling tracking on an already dirty document marks the baseline as incomplete'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'ApplyActivity' 'FindMatchingActivities' 'Undo and Redo can move more than one matched Revit activity'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementActivityMatcherPath 'Match' 'suffixIndexes.Add(i);' 'grouped Undo and Redo consume a contiguous LIFO activity suffix'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementActivityMatcherPath 'Match' 'equivalentShorterSuffixExists' 'indistinguishable repeated Undo or Redo evidence cannot select one guessed activity'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'ApplyActivity' 'AddAmbiguousActivityIds(session, activity)' 'partial or unmatched Undo and Redo ambiguity is scoped to the observed element IDs'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'ApplyActivity' 'AddAmbiguousActivityIds(session, matched)' 'partial Undo and Redo also quarantine every element from the guessed candidate activity'
Assert-WorkflowMethodOrder 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'ApplyActivity' 'if (exact && matched.Count > 0)' 'session.AppliedActivities.Remove(matchedActivity)' 'Undo and Redo stacks move only after an exact match'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementTrackingTransitionPolicyPath 'ResolveChangeKind' '!elementSequenceAmbiguous && hasActiveActivity' 'an ambiguous stale activity cannot fabricate a net modification while unrelated local changes remain observable'
Assert-WorkflowToken 'element-tracking-event-ambiguity' $workflowHarnessPath 'an ambiguous stale activity fabricated a modification without a state-signature change' 'the executable harness rejects ambiguous stale activity without a real state change'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'BuildCommit' 'ClientObservedWithEventAmbiguity' 'unmatched Undo or Redo is exposed in attribution confidence'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'NotifyPolicyChanged' 'DisableLiveTrackingNoLock(root, false);' 'tracking OFF removes inactive live sessions while retaining durable recovery sessions in the affected context'
Assert-WorkflowMethodOrder 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'NotifyPolicyChanged' 'bool hasUncommittedEvidence = Sessions.Values.Any' 'DisableLiveTrackingNoLock(root, true);' 'the service-level tracking OFF path checks uncommitted evidence before it can clear live sessions'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'DisableLiveTrackingNoLock' 'session.PolicyDisableDeferred = true;' 'a stale UI or direct local policy notification cannot discard already-observed uncommitted evidence'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'NotifyPolicyChanged' 'Element tracking local disable deferred until commit boundary' 'service-level deferred disable is diagnosed for administrator review'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'DisableLiveTrackingNoLock' 'SynchronizingDocumentKeys.ExceptWith(removedKeys);' 'tracking OFF clears synchronization suppression state only for sessions actually removed'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'DisableLiveTrackingNoLock' 'ReloadingDocumentKeys.ExceptWith(removedKeys);' 'tracking OFF clears Reload Latest suppression state only for sessions actually removed'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'EndDocumentSession' 'SynchronizingDocumentKeys.Remove(runtimeKey);' 'ending one document session also clears its stale synchronization suppression marker'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'EndDocumentSession' 'ReloadingDocumentKeys.Remove(runtimeKey);' 'ending one document session also clears its stale Reload Latest suppression marker'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'HandleDocumentSynchronizingWithCentral' 'HasRetainedSessionEvidence(doc)' 'synchronization start preserves already-observed evidence when policy lookup is disabled or temporarily unavailable'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'HandleDocumentReloadingLatest' 'HasRetainedSessionEvidence(doc)' 'Reload Latest start preserves already-observed evidence when policy lookup is disabled or temporarily unavailable'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'HandleDocumentReloadedLatest' 'FamilyBrowserElementTrackingPolicyDecision.Resolve' 'Reload Latest completion uses the explicit policy and retained-evidence decision table'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'HandleDocumentCommitted' 'FamilyBrowserElementTrackingPolicyDecision.Resolve' 'Save and Sync completion use the explicit policy and retained-evidence decision table'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'HandleDocumentCommitted' 'FamilyBrowserElementTrackingSessionMode.DeferredCommit' 'uncertain policy state cannot discard already-observed evidence at a successful boundary'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'BuildCommit' 'session.ExternalRebaseFailed' 'an external-update rebase failure qualifies as an explicit zero-row coverage gap'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'HasUncommittedSessionEvidence' 'session.ExternalRebaseFailed' 'external-update rebase failure is retained as uncommitted evidence until a successful boundary'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $trackingPath 'ValidateElementChangeCommit' 'externalRebaseGap' 'schema validation accepts only explicitly identified external-rebase zero-row coverage evidence'
Assert-WorkflowToken 'element-tracking-event-ambiguity' $workflowHarnessPath 'grouped Undo and Redo matching consumes only an exact contiguous LIFO activity suffix' 'workflow harness proves grouped Undo and Redo matching order'
Assert-WorkflowToken 'element-tracking-event-ambiguity' $workflowHarnessPath 'partial Undo and Redo matching remains explicitly ambiguous' 'workflow harness proves uncertain grouped events do not become exact attribution'
Assert-WorkflowToken 'element-tracking-event-ambiguity' $workflowHarnessPath 'indistinguishable repeated Undo and Redo evidence never mutates one guessed activity' 'workflow harness proves repeated same-element evidence remains conservative'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'BuildCommit' 'Tracking includes project/shared parameter definitions and binding metadata.' 'every committed ledger record discloses the newly tracked parameter scope'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'BuildCommit' 'It excludes View, DataStorage, ProjectInfo' 'every committed ledger record discloses the intentionally excluded Revit-internal element scope'
Assert-WorkflowToken 'tracking-policy-concurrency' $elementChangeTrackingPath 'session.PolicyDisableDeferred = true;' 'a remote disable preserves already-observed uncommitted evidence until a successful boundary'
Assert-WorkflowMethodToken 'tracking-policy-concurrency' $elementChangeTrackingPath 'BuildCommit' 'DisablePendingCommit' 'history marks records committed after a remote tracking disable request'
Assert-WorkflowMethodToken 'tracking-policy-concurrency' $elementChangeTrackingPath 'BeginDocumentSession' 'HasPendingRecoveryCheckpoint(resolvedWorkspaceRoot, doc)' 'disabled tracking can still reopen a protected workshared checkpoint for synchronization recovery'
Assert-WorkflowMethodToken 'tracking-policy-concurrency' $elementChangeTrackingPath 'HandleDocumentChanged' 'recoverySession.RecoveryOnly' 'a recovery-only disabled-policy session does not observe new local edits'
Assert-WorkflowMethodToken 'tracking-policy-concurrency' $elementChangeTrackingPath 'HandleDocumentCommitted' 'recoveryOnlyCommit' 'a successful synchronization can publish an existing checkpoint after tracking is disabled'
Assert-WorkflowMethodToken 'tracking-policy-concurrency' $elementChangeTrackingPath 'DisableLiveTrackingNoLock' 'session.RecoveryOnly = true;' 'disabling live tracking retains protected checkpoints as recovery-only sessions'
Assert-WorkflowMethodToken 'tracking-policy-concurrency' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'policyDisableDeferredCommits' 'history visibly counts evidence protected after a remote disable'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $projectCatalogDashboardPath 'RenderSelectedProjectElementChangeHistory' 'commits.Count(ProjectElementHasExternalRebaseGap)' 'history counts external-rebase gaps even when another attribution warning has higher display priority'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $projectCatalogDashboardPath 'ProjectElementAttributionLabel' 'Incoming update overlap' 'history rows expose incoming-update overlap as an independent warning'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $projectCatalogDashboardPath 'ProjectElementAttributionLabel' 'External update coverage gap' 'history rows expose external-rebase gaps alongside overlap and other warnings'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $projectCatalogDashboardPath 'ProjectElementAttributionLabel' 'DocumentChanged coverage gap' 'history rows expose event-read coverage gaps instead of hiding them behind one attribution label'
Assert-WorkflowMethodExcludesToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'RebaseAfterExternalUpdate' 'ExternalRebaseFailed = false;' 'a later successful rebase cannot erase an earlier external-coverage gap'
Assert-WorkflowMethodToken 'element-tracking-event-ambiguity' $elementChangeTrackingPath 'HandleDocumentCommitted' 'if (!succeeded)' 'failed or cancelled commit events leave no immutable element history'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'CreateSession' 'CheckpointProjectIdentityPath = SafeProjectIdentityPath(doc)' 'local checkpoints retain the project or central identity that created them'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'SameCheckpointIdentity' 'SameProjectIdentity(leftProject, rightProject)' 'checkpoint reuse requires both project identity and local-file identity to match'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'HandleDocumentCommitted' 'if (!sameCheckpointIdentity && sameCheckpointProject && !string.IsNullOrWhiteSpace(previousCheckpointRevisionToken)' 'successful new-project synchronization cleans a previous-path checkpoint only inside the same project identity'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'HandleDocumentCommitted' 'previousCheckpointProjectIdentity,' 'previous-path cleanup targets the identity that originally created the checkpoint'
Assert-WorkflowMethodToken 'element-tracking-session-recovery' $elementChangeTrackingPath 'HandleDocumentCommitted' 'previous-project checkpoint retained' 'central identity changes are diagnosed instead of silently consuming prior evidence'
Assert-WorkflowMethodToken 'tracking-policy-concurrency' $projectCatalogDashboardPath 'RecordProjectElementChangeTrackingPolicyChange' 'ProjectElementChangeTrackingPolicy' 'tracking policy changes create immutable operation evidence'
Assert-WorkflowMethodToken 'tracking-policy-concurrency' $projectCatalogDashboardPath 'RecordProjectElementChangeTrackingPolicyChange' ';source=' 'tracking policy audit distinguishes its mutation route from failure detail'

Assert-WorkflowMethodToken 'guard-admin-transition' $dashboardPath 'SetAdminMode' 'NotifyAdminModeChanged(_adminModeEnabled, _standardPolicy)' 'Admin ON/OFF pushes the already-loaded policy to the native guard'
Assert-WorkflowMethodToken 'guard-admin-transition' $dashboardPath 'SetAdminMode' '_dashboardPermissionCachedSnapshot = null;' 'Admin ON/OFF invalidates the browser permission snapshot'
Assert-WorkflowMethodToken 'guard-admin-transition' $dashboardPath 'SetAdminMode' 'SynchronizeAdminModeGuardState' 'Admin ON/OFF verifies the effective browser/native guard immediately'
Assert-WorkflowMethodToken 'guard-admin-transition' $dashboardPath 'CompleteInitialOpenRefresh' 'ApplyAdminModeAfterPolicyLoad(restorePersistedSelection: true, refreshUiNow: false)' 'startup restores the saved Admin selection and schedules native UI work through Revit idle'
Assert-WorkflowMethodExcludesToken 'guard-admin-transition' $dashboardPath 'CompleteInitialOpenRefresh' '_adminModeEnabled = CanEnableAdminMode(_standardPolicy)' 'startup cannot reinterpret Admin capability as Admin Mode ON'
Assert-WorkflowMethodToken 'guard-admin-transition' $dashboardPath 'ApplyAdminModeAfterPolicyLoad' 'ResolveEffectiveAdminMode(requestedEnabled, canEnable)' 'policy reload keeps Admin selection and capability as separate inputs'
Assert-WorkflowMethodToken 'guard-admin-transition' $dashboardPath 'ApplyAdminModeAfterPolicyLoad' 'ResolveInitialAdminModeEnabled(canEnable)' 'first-run Admin profile starts ON while explicit ON/OFF choices remain persisted'
Assert-WorkflowToken 'guard-admin-transition' $dashboardPath 'ApplyAdminModeAfterPolicyLoad(restorePersistedSelection: false, refreshUiNow: false)' 'management-folder setup preserves Admin OFF and schedules native UI work through Revit idle'
Assert-WorkflowMethodToken 'guard-admin-transition' $dashboardPath 'HasPermission' 'CanNativeGuard(policy, currentUser, permission, context, _adminModeEnabled)' 'browser action permissions follow effective Admin mode'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'LoadAdminModeEnabledSetting' 'FamilyBrowserUserSettingsStore.LoadAdminModeEnabled()' 'native guard reloads persisted Admin OFF state'
Assert-WorkflowMethodExcludesToken 'guard-admin-transition' $guardPath 'LoadAdminModeEnabledSetting' 'SaveAdminModeEnabled(enabled: true)' 'native guard does not silently force Admin Mode ON'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'Start' 'RegisterProtectedChangeUpdater(application)' 'family/type rollback updater is registered only in the Revit startup API context'
Assert-WorkflowMethodExcludesToken 'guard-admin-transition' $guardPath 'RefreshProtectedChangeUpdaterRegistration' 'RegisterProtectedChangeUpdater(application)' 'modeless Admin transition cannot fail by registering an updater outside API context'
Assert-WorkflowToken 'guard-admin-transition' $guardPath 'NativeGuardDecisionCache.Clear();' 'Admin transition invalidates native permission cache immediately'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'UpdateProtectedRibbonAvailability' 'ApplyLoadFamilyRibbonEnabledFast(allowed);' 'Admin OFF immediately synchronizes the native Load Family ribbon control'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'UpdateProtectedRibbonAvailability' 'FindBoundDefinition("native-load-family") ?? FamilyLoadingEventDefinition' 'Load Family ribbon guard remains active even when Revit exposes no bindable Load Family command id'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'ResolveLoadFamilyRibbonControlsFast' 'CachedLoadFamilyRibbonControls.AddRange(matches);' 'native Load Family ribbon controls are discovered once and cached'
Assert-WorkflowMethodExcludesToken 'guard-admin-transition' $guardPath 'UpdateProtectedRibbonAvailability' 'ApplyLoadFamilyRibbonEnabledRecursive' 'Admin transition does not traverse the full Autodesk ribbon object graph'
Assert-WorkflowToken 'guard-admin-transition' $guardPath 'FamilyLoadingIntoDocument += HandleFamilyLoadingIntoDocument' 'native Family load interception is attached'
Assert-WorkflowToken 'guard-admin-transition' $guardPath 'RequiredPermission = "RenameFamilyOrType"' 'native Family and Type rename permission is registered'
Assert-WorkflowToken 'guard-admin-transition' $guardPath '"ID_PRJBROWSER_RENAME"' 'the Revit journal-confirmed Project Browser F2 rename command is bound'
Assert-WorkflowToken 'guard-admin-transition' $guardPath 'BuiltInParameter.ALL_MODEL_TYPE_NAME' 'type-name updater fallback covers system and loadable ElementType names'
Assert-WorkflowToken 'guard-admin-transition' $guardPath ';projectBrowserRenameBinding=' 'runtime guard diagnostics report the concrete Project Browser rename binding'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'ShouldRecordProtectedChange' '!previousInfoAvailable' 'the first rename of an ElementType missing from a partial guard index fails closed'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'HandleDocumentChanged' 'UpdateProtectedElementIndexFromChanges(doc, e.GetAddedElementIds(), e.GetModifiedElementIds(), e.GetDeletedElementIds());' 'Admin ON changes keep partial guard baselines synchronized'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'Start' 'AttachUiEvents(application);' 'native guard attaches the deferred ribbon settlement event during Revit startup'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'HandleIdling' 'UpdateProtectedRibbonAvailability(force: true);' 'the next Revit idle cycles reassert Load Family availability after Admin transitions'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'HandleIdling' 'EnsureProtectedElementIndexForGuard(LastActiveDocument);' 'Revit idle prepares original Family/Type names before protected edits'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'EnsureProtectedElementIndexForGuard' 'EnsureProtectedElementIndexBaseline(doc);' 'Admin OFF creates a complete protected-name baseline'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'RefreshProtectedElementIndex' 'CompleteProtectedElementIndexDocumentTokens[documentKey]' 'a complete baseline cannot be mistaken for a partial changed-element index'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'NotifyAdminModeChanged' 'ScheduleProtectedRibbonRefresh();' 'Admin ON/OFF does not depend on a manual browser refresh to settle ribbon state' 'public static void NotifyAdminModeChanged(bool enabled, FamilyBrowserStandardPolicy policy, bool refreshUiNow = true)'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'NotifyAdminModeChanged' 'ScheduleProtectedElementBaselineRefresh();' 'Admin OFF schedules its baseline before the next user edit' 'public static void NotifyAdminModeChanged(bool enabled, FamilyBrowserStandardPolicy policy, bool refreshUiNow = true)'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'HandleFailuresProcessing' 'SchedulePostRollbackUiRefresh();' 'blocked transactions request visible UI settlement after rollback'
Assert-WorkflowMethodToken 'guard-admin-transition' $guardPath 'RefreshRevitUiAfterProtectedRollback' 'uiDocument.RefreshActiveView();' 'post-rollback idle refreshes Revit without another rename command'
Assert-WorkflowToken 'guard-admin-transition' $guardPath ';pendingRibbonRefreshPasses=' 'runtime diagnostics report pending ribbon settlement passes'
Assert-WorkflowToken 'guard-admin-transition' $guardPath ';protectedElementBaselineComplete=' 'runtime diagnostics report whether original-name restoration is ready'
Assert-WorkflowToken 'guard-admin-transition' $guardPath 'ShouldBlockNestedOnlyStandalonePlacement' 'nested-only standalone placement policy is enforced'
Assert-WorkflowScenario 'guard-admin-transition' 'admin-requests-and-permissions'

Assert-WorkflowToken 'request-lifecycle' $dashboardPath 'UpdateRequestStatusFromAction(action);' 'request status route is connected'
Assert-WorkflowToken 'request-lifecycle' $dashboardPath 'DeleteRequestFromAction(action);' 'request delete route is connected'
Assert-WorkflowToken 'request-lifecycle' $dashboardPath 'OpenRequestAttachmentFolderFromAction(action);' 'request attachment route is connected'
Assert-WorkflowMethodToken 'request-lifecycle' $requestStorePath 'WriteAllTextAtomic' 'FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path)' 'request records use a short same-folder temporary path'
Assert-WorkflowMethodToken 'request-lifecycle' $requestStorePath 'WriteAllTextAtomic' 'FamilyBrowserAtomicFileService.Promote(tempPath, path)' 'request records use recoverable atomic promotion'
Assert-WorkflowMethodToken 'request-lifecycle' $atomicFileServicePath 'Promote' 'File.Move(backupPath, destinationPath)' 'failed request promotion restores the previous committed record'
Assert-WorkflowScenario 'request-lifecycle' 'admin-requests-and-permissions'

Assert-WorkflowToken 'request-concurrent-edit' $requestConcurrencyPath 'FileShare.None' 'request-scoped lock excludes a second writer across processes and PCs'
Assert-WorkflowToken 'request-concurrent-edit' $requestConcurrencyPath 'EnsureExpectedRevision' 'request mutation validates the rendered revision and token'
Assert-WorkflowToken 'request-concurrent-edit' $requestConcurrencyPath 'ComputeFileToken' 'legacy and old-client request writes have a content compatibility token'
Assert-WorkflowToken 'request-concurrent-edit' $requestRecordPath 'public long Revision' 'request records persist an optimistic revision'
Assert-WorkflowToken 'request-concurrent-edit' $requestRecordPath 'public string RevisionToken' 'request records persist a revision token'
Assert-WorkflowMinimumOccurrences 'request-concurrent-edit' $requestStorePath 'FamilyBrowserRequestConcurrencyService.Acquire' 3 'create/save, status update, and delete are serialized by request ID'
Assert-WorkflowMinimumOccurrences 'request-concurrent-edit' $requestStorePath 'FamilyBrowserRequestConcurrencyService.EnsureExpectedRevision' 3 'every existing-record mutation rejects stale input'
Assert-WorkflowToken 'request-concurrent-edit' $dashboardPath 'expectedRevisionToken' 'request action routes carry the rendered content token'
Assert-WorkflowToken 'request-concurrent-edit' $dashboardPath 'HandleRequestConflict' 'stale request actions reload the latest list without overwriting data'
Assert-WorkflowToken 'request-concurrent-edit' $managedFolderSetupPath 'name.StartsWith(".kky-r-"' 'request lock files are excluded from managed-folder migration'
Assert-WorkflowToken 'request-concurrent-edit' $workflowHarnessPath 'request-scoped lock serializes writers and releases cleanly' 'workflow harness proves lock exclusivity and release'
Assert-WorkflowToken 'request-concurrent-edit' $workflowHarnessPath 'legacy and old-client edits are detected by the file-content token' 'workflow harness proves old-client edit detection'

Assert-WorkflowToken 'request-attachment-and-delete-audit' $requestFileTransactionPath 'CopyContentAddressed' 'request attachments use deterministic content-addressed storage'
Assert-WorkflowToken 'request-attachment-and-delete-audit' $requestFileTransactionPath 'CopyAndHash' 'attachment hashing and temporary copying occur in one source-file pass'
Assert-WorkflowToken 'request-attachment-and-delete-audit' $requestFileTransactionPath 'WriteImmutableText' 'request deletion audit files cannot overwrite prior events'
Assert-WorkflowToken 'request-attachment-and-delete-audit' $requestAttachmentRecordPath 'public string ContentSha256' 'request attachment metadata preserves the content identity'
Assert-WorkflowToken 'request-attachment-and-delete-audit' $requestStorePath 'RollbackAttachmentMutation' 'pre-commit request failures restore attachment metadata and files'
Assert-WorkflowToken 'request-attachment-and-delete-audit' $requestStorePath 'requestCommitted = true;' 'authoritative request JSON defines the attachment commit point'
Assert-WorkflowToken 'request-attachment-and-delete-audit' $requestStorePath '"RequestAudit", "Deleted"' 'deleted request snapshots live outside the active request list'
Assert-WorkflowToken 'request-attachment-and-delete-audit' $requestStorePath '"DeletePrepared"' 'a full deletion snapshot is durable before active files are removed'
Assert-WorkflowToken 'request-attachment-and-delete-audit' $requestStorePath '"DeleteCompleted"' 'successful deletion receives an immutable completion event'
Assert-WorkflowMinimumOccurrences 'request-attachment-and-delete-audit' $dashboardPath 'ShowRequestAuxiliaryWarning(' 3 'create and status actions visibly report committed auxiliary metadata failures'
Assert-WorkflowToken 'request-attachment-and-delete-audit' $dashboardPath 'Request auxiliary metadata warning' 'auxiliary failure details are written to diagnostics'
Assert-WorkflowToken 'request-attachment-and-delete-audit' $workflowHarnessPath 'request attachment retries are content-addressed and idempotent' 'workflow harness proves retry deduplication'
Assert-WorkflowToken 'request-attachment-and-delete-audit' $workflowHarnessPath 'request deletion audit entries are immutable and preserve the prepared snapshot' 'workflow harness proves deletion audit immutability'

Assert-WorkflowToken 'scan-dialog-recovery' $dialogGuardPath 'ResolveFamilyEditDialogActionForAudit' 'family-edit warning button topology has an audit seam'
Assert-WorkflowToken 'scan-dialog-recovery' $dialogGuardPath 'OpeningNotCuttingAnything' 'Opening not cutting anything is classified'
Assert-WorkflowToken 'scan-dialog-recovery' $dialogGuardPath 'HasOnlyDeleteOrCancelButtons' 'destructive-only warnings choose Cancel'
Assert-WorkflowToken 'scan-dialog-recovery' $dashboardPath 'AppendScanDialogRows' 'auto-handled scan warnings are included in exportable result rows'
Assert-WorkflowToken 'scan-dialog-recovery' $thumbnailPreviewPath 'FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path)' 'thumbnail metadata is written to a sibling temporary file'
Assert-WorkflowToken 'scan-dialog-recovery' $thumbnailPreviewPath 'stream.Flush(true);' 'thumbnail metadata is fully flushed before publication'
Assert-WorkflowToken 'scan-dialog-recovery' $thumbnailPreviewPath 'FamilyBrowserAtomicFileService.Promote(temporaryPath, path);' 'thumbnail metadata is atomically promoted'

Assert-WorkflowToken 'language-result-export' $uiHarnessScriptPath "foreach (`$languageCode in @('ko', 'en'))" 'quality harness renders Korean and English variants'
Assert-WorkflowToken 'language-result-export' $uiHarnessPath 'CheckLanguagePurity(browser, options, result);' 'rendered dashboard language purity is checked'
Assert-WorkflowToken 'language-result-export' $dashboardPath 'ShowDashboardResultWithExcelExport' 'rich HTML results expose on-demand Excel export'
Assert-WorkflowMethodToken 'language-result-export' $dashboardPath 'ShowDashboardResultWithExcelExport' 'AuxiliaryActionRequested' 'Excel is created only after the user requests export'

Assert-WorkflowToken 'large-library-performance' $performanceTestPath "'--syntheticFamilyCount', '1000'" 'performance gate renders 1000 synthetic Family rows'
Assert-WorkflowToken 'large-library-performance' $performanceTestPath "'--syntheticSystemCount', '1000'" 'performance gate renders 1000 synthetic System rows'
Assert-WorkflowToken 'large-library-performance' $performanceTestPath 'FilterTargetMs = 150' 'performance gate enforces the filter response target'
Assert-WorkflowToken 'large-library-performance' $rowWindowJsPath 'var windowSize = 150;' 'browser row window limits one rendered page to 150 rows'

Assert-WorkflowToken 'diagnostics-debug-log' $dashboardPath 'id=\"fbDebug\"' 'Debug Log is hosted in the bottom browser dock'
Assert-WorkflowToken 'diagnostics-debug-log' $dashboardPath 'case "debug-log":' 'Debug Log menu route is connected'
Assert-WorkflowToken 'diagnostics-debug-log' $uiHarnessPath 'CheckDebugDock(browser, result);' 'IE harness validates the Debug Log dock geometry and visibility'
Assert-WorkflowToken 'diagnostics-debug-log' $dashboardPath 'WriteDashboardRuntimeDiagnostic' 'browser workflow stages retain runtime diagnostics'

$scanRoots = @(
    $sharedRoot,
    (Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2019-2023'),
    (Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2025'),
    (Join-Path $repoRoot 'KKY_FamilyBrowser_RevitHost_2027')
)
$invalidDateStyle = 'DateTimeStyles\.RoundtripKind\s*\|\s*DateTimeStyles\.(?:AssumeUniversal|AdjustToUniversal)|DateTimeStyles\.(?:AssumeUniversal|AdjustToUniversal)\s*\|\s*DateTimeStyles\.RoundtripKind'
$invalidMatches = New-Object System.Collections.Generic.List[string]
foreach ($scanRoot in $scanRoots) {
    foreach ($file in Get-ChildItem -LiteralPath $scanRoot -Recurse -Filter '*.cs' -File) {
        $match = Select-String -LiteralPath $file.FullName -Pattern $invalidDateStyle
        foreach ($item in $match) {
            $invalidMatches.Add("$($file.FullName):$($item.LineNumber)") | Out-Null
        }
    }
}
Add-Check 'DateTimeStyles combinations are valid' ($invalidMatches.Count -eq 0) ($invalidMatches -join '; ')

$hostFolders = @(
    'KKY_FamilyBrowser_RevitHost_2019-2023',
    'KKY_FamilyBrowser_RevitHost_2025',
    'KKY_FamilyBrowser_RevitHost_2027'
)
foreach ($folder in $hostFolders) {
    $servicePath = Join-Path $repoRoot "$folder\StandardRvtChangeCandidateService.cs"
    $appPath = if ($folder -eq 'KKY_FamilyBrowser_RevitHost_2019-2023') {
        Join-Path $repoRoot "$folder\KKY_FamilyBrowser_RevitHost_2019_2023\App.cs"
    } else {
        Join-Path $repoRoot "$folder\$folder\App.cs"
    }
    $dashboardHostPath = Join-Path $repoRoot "$folder\FamilyBrowserDashboardHtmlForm.cs"
    $policyPath = Join-Path $repoRoot "$folder\FamilyBrowserStandardPolicy.cs"
    $policyStorePath = Join-Path $repoRoot "$folder\FamilyBrowserStandardPolicyStore.cs"
    $machineConfigPath = Join-Path $repoRoot "$folder\FamilyBrowserMachineConfigStore.cs"
    $fileGuardTargetPath = Join-Path $repoRoot "$folder\FamilyBrowserFileGuardTarget.cs"
    $fileGuardUiPath = Join-Path $repoRoot "$folder\FileGuardHtmlConfigurationForm.cs"
    $securityPolicyPath = Join-Path $repoRoot "$folder\FamilyBrowserSecurityPolicyService.cs"
	$standardDocumentResolverPath = Join-Path $repoRoot "$folder\StandardLibraryDocumentResolver.cs"
	Assert-WorkflowMethodOrder 'family-load-save-lifecycle' $standardDocumentResolverPath 'OpenRegisteredDocument' 'FindOpenDocument(application, resolvedPath)' 'File.Exists(resolvedPath)' "$folder reuses an already-open Standard RVT before probing the network path"
	Assert-WorkflowMethodToken 'family-load-save-lifecycle' $standardDocumentResolverPath 'FindOpenDocument' 'FamilyBrowserPathIdentityService.GetStablePathIdentity(doc.PathName)' "$folder reuses an already-open Standard RVT even when mapped-drive and UNC spellings differ"
	Assert-WorkflowMethodOrder 'family-load-save-lifecycle' $dashboardHostPath 'ApplyStandardFamilies' 'BeginLongOperationProgress(commandTitle' 'LoadRegistrationOrWarn(T("Apply Standard Families"' "$folder shows progress before standard revision and snapshot validation can touch network storage"
	Assert-WorkflowMethodOrder 'family-load-save-lifecycle' $dashboardHostPath 'ApplyStandardFamilies' 'CloseOwnedStandardDocumentBeforeResult' 'FamilyBrowserOperationHtmlDialog.ShowFamilyLoadResult' "$folder closes its owned Standard RVT before showing the terminal Family load result"
	Assert-WorkflowMethodOrder 'family-load-save-lifecycle' $dashboardHostPath 'ApplyStandardFamilies' 'FamilyLoad.PostApply.RefreshCompleted' 'FamilyBrowserOperationHtmlDialog.ShowFamilyLoadResult' "$folder finishes Family list refresh before showing the result dialog"
	Assert-WorkflowMethodOrder 'system-type-apply-lifecycle' $dashboardHostPath 'ApplyStandardSystemTypes' 'CloseOwnedStandardDocumentBeforeResult' 'FamilyBrowserOperationHtmlDialog.ShowSystemTypeApplyResult' "$folder closes its owned Standard RVT before showing the terminal System Type result"
	Assert-WorkflowMethodOrder 'system-type-apply-lifecycle' $dashboardHostPath 'ApplyStandardSystemTypes' 'BeginLongOperationProgress(commandTitle' 'LoadRegistrationOrWarn(T("Apply Standard System Types"' "$folder shows progress before System Type standard revision validation can touch network storage"
	Assert-WorkflowMethodOrder 'system-type-apply-lifecycle' $dashboardHostPath 'ApplyStandardSystemTypes' 'SystemApply.PostApply.RefreshCompleted' 'FamilyBrowserOperationHtmlDialog.ShowSystemTypeApplyResult' "$folder finishes System Type list refresh before showing the result dialog"
	Assert-WorkflowMethodExactOccurrences 'system-type-apply-lifecycle' $dashboardHostPath 'ApplyStandardSystemTypes' 'SystemTypeApplyStore.Save(_workspaceRoot, execution)' 1 "$folder persists the final System Type execution report only once"
	Assert-WorkflowMethodToken 'system-type-apply-lifecycle' $dashboardHostPath 'ApplyStandardSystemTypes' 'if (!isSelectedApply)' "$folder skips the intermediate full dashboard render for selected System Type apply"
	Assert-WorkflowMethodToken 'managed-folder-first-run' $appPath 'HandleDocumentOpened' 'RunEventHandlerSafely("Managed policy document-open preparation failed"' "$folder isolates managed-folder preparation failure at project open"
	Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'HandleDocumentOpened' 'RunEventHandlerSafely("Element tracking document-open baseline failed"' "$folder still starts tracking when another document-open subsystem fails"
	Assert-WorkflowMethodToken 'guard-admin-transition' $appPath 'HandleDocumentOpened' 'RunEventHandlerSafely("Native guard document-open policy preload failed"' "$folder still starts the native guard when another document-open subsystem fails"
	Assert-WorkflowMethodToken 'current-model-check-baseline' $appPath 'HandleDocumentOpened' 'RunEventHandlerSafely("Automatic Current Model Check document-open scheduling failed"' "$folder still schedules automatic checking when another document-open subsystem fails"
	Assert-WorkflowMethodToken 'guard-admin-transition' $appPath 'HandleViewActivated' 'RunEventHandlerSafely("Native guard view-activation policy preload failed"' "$folder isolates native guard preload failure during view activation"
    Assert-MethodToken $servicePath 'AppendImmutableCandidates' 'PersistStandardCandidateEntries' "$folder standard history uses write-ahead persistence"
    Assert-MethodToken $servicePath 'AppendOperationEntries' 'PersistOperationEntries' "$folder operation history uses write-ahead persistence"
    Assert-MethodToken $servicePath 'CommitPendingCandidateEntries' 'RestorePendingCandidateBatch' "$folder candidate batch is restored when persistence fails"
    Assert-MethodToken $servicePath 'CommitPendingOperationEntries' 'RestorePendingOperationBatch' "$folder operation batch is restored when persistence fails"
    Assert-MethodToken $servicePath 'LoadRecent' 'FlushPending(workspaceRoot)' "$folder standard history flushes local spool before reading"
    Assert-MethodToken $servicePath 'BuildLoadableFamilyOperationEntries' 'ResolveLoadedFamilyTypeNames' "$folder Family load captures exact type names"
    Assert-MethodToken $servicePath 'BuildLoadableFamilyOperationEntries' 'typeEntry.TypeName = typeName' "$folder Family Type operation entries are explicit"
    Assert-MethodToken $servicePath 'IsPendingOperationPresentAtSync' 'entry.TypeName' "$folder save verification checks the exact Family Type"
    Assert-MethodOrder $servicePath 'HandleDocumentSaved' 'IsSuccessfulRevitApiEventStatus(status)' 'CommitPendingCandidateEntries' "$folder failed save cannot commit tracking"
    Assert-MethodOrder $servicePath 'HandleDocumentSynchronizedWithCentral' 'IsSuccessfulRevitApiEventStatus(status)' 'CommitPendingOperationEntries' "$folder failed sync cannot commit tracking"
    Assert-MethodToken $servicePath 'HandleDocumentClosing' 'PendingOperationCloseKeysByDocumentId' "$folder close-start only remembers pending state"
    Assert-MethodToken $servicePath 'HandleDocumentClosed' 'PendingOperationEntriesByDocument.Remove' "$folder actual close discards uncommitted operation memory"

    Assert-WorkflowMethodOrder 'standard-edit-commit-lifecycle' $servicePath 'HandleDocumentSaved' 'IsSuccessfulRevitApiEventStatus(status)' 'CommitPendingCandidateEntries' "$folder successful Save gate precedes Standard RVT history commit"
    Assert-WorkflowMethodOrder 'standard-edit-commit-lifecycle' $servicePath 'HandleDocumentSynchronizedWithCentral' 'IsSuccessfulRevitApiEventStatus(status)' 'CommitPendingCandidateEntries' "$folder successful Sync gate precedes Standard RVT history commit"
    Assert-WorkflowMethodToken 'standard-edit-commit-lifecycle' $servicePath 'CommitPendingCandidateEntries' 'string.Equals(commitKind, "SaveAs"' "$folder Save As commits only when the final path is the registered Standard RVT"

    Assert-WorkflowMethodToken 'family-load-save-lifecycle' $servicePath 'AppendOperationEntries' 'PersistOperationEntries' "$folder Family/System operations use durable persistence"
    Assert-WorkflowMethodToken 'family-load-save-lifecycle' $servicePath 'CommitPendingOperationEntries' 'IsPendingOperationPresentAtSync' "$folder commit verifies that the loaded/applied item still exists"
    Assert-WorkflowMethodToken 'family-load-save-lifecycle' $servicePath 'CommitPendingOperationEntries' 'RestorePendingOperationBatch' "$folder failed persistence restores the live pending batch"

    Assert-WorkflowMethodToken 'family-type-attribution' $servicePath 'BuildLoadableFamilyOperationEntries' 'ResolveLoadedFamilyTypeNames' "$folder Family load enumerates the actual loaded type names"
    Assert-WorkflowMethodToken 'family-type-attribution' $servicePath 'BuildLoadableFamilyOperationEntries' 'typeEntry.TypeName = typeName' "$folder emits one explicit operation record per Family Type"
    Assert-WorkflowMethodToken 'family-type-attribution' $servicePath 'IsPendingOperationPresentAtSync' 'entry.TypeName' "$folder validates the exact Family Type at commit"

    Assert-WorkflowMethodToken 'close-cancel-and-failure' $servicePath 'HandleDocumentSaved' 'if (!IsSuccessfulRevitApiEventStatus(status))' "$folder failed or cancelled Save cannot commit"
    Assert-WorkflowMethodToken 'close-cancel-and-failure' $servicePath 'HandleDocumentClosing' 'PendingOperationCloseKeysByDocumentId' "$folder close-start preserves pending state until closure is final"
    Assert-WorkflowMethodToken 'close-cancel-and-failure' $servicePath 'HandleDocumentClosed' 'PendingOperationEntriesByDocument.Remove' "$folder completed close discards uncommitted in-memory operations"

	Assert-WorkflowToken 'project-element-change-ledger' $policyPath 'TrackProjectElementChanges' "$folder retains the legacy tracking mirror for policy compatibility"
	Assert-WorkflowToken 'project-element-change-ledger' $fileGuardTargetPath 'TrackElementChanges' "$folder File Guard target persists per-file element tracking scope"
	Assert-WorkflowToken 'project-element-change-ledger' $fileGuardTargetPath 'TrackElementChangesConfigured' "$folder legacy File Guard target migration is explicit"
	Assert-WorkflowMethodToken 'project-element-change-ledger' $fileGuardUiPath 'BuildPolicyFromRows' 'TrackElementChanges = row.TrackElements' "$folder File Guard UI persists the per-file tracking checkbox"
	Assert-WorkflowMethodToken 'project-element-change-ledger' $policyStorePath 'NormalizeFileGuardTarget' 'target.TrackElementChanges = true;' "$folder existing registered RVTs migrate to tracking enabled"
	Assert-WorkflowMethodToken 'project-element-change-ledger' $policyStorePath 'IsProjectElementChangeTrackingEnabled' 'target.Enabled && target.TrackElementChanges' "$folder registered checked RVTs are the only tracking master state"
	Assert-WorkflowMethodExcludesToken 'project-element-change-ledger' $policyStorePath 'IsProjectElementChangeTrackingEnabled' 'TrackProjectElementChanges == true' "$folder the legacy global field cannot enable an unregistered RVT"
	Assert-WorkflowMethodToken 'project-element-change-ledger' $securityPolicyPath 'IsProjectElementTrackingScopeEnabled' 'matchingTarget.TrackElementChanges' "$folder only the matching checked RVT enters element tracking scope"
	Assert-WorkflowMethodExcludesToken 'project-element-change-ledger' $securityPolicyPath 'IsProjectElementTrackingScopeEnabled' 'return true;' "$folder an absent or empty File Guard cannot enable element tracking"
	Assert-WorkflowToken 'tracking-policy-concurrency' $policyStorePath 'AcquirePolicyFileMutationLock(GetPolicyPath(workspaceRoot)' "$folder policy mutations acquire a shared-file lease"
    Assert-WorkflowMethodToken 'tracking-policy-concurrency' $policyStorePath 'AcquirePolicyFileMutationLock' 'FileShare.None' "$folder policy lock excludes another PC writer on the same SMB object"
    Assert-WorkflowMethodToken 'tracking-policy-concurrency' $policyStorePath 'SaveUnlocked' 'FamilyBrowserAtomicFileService.Promote' "$folder shared policy uses recoverable atomic promotion"
    Assert-WorkflowToken 'tracking-policy-concurrency' $policyStorePath 'FamilyBrowserAtomicFileService.Promote(temporaryPath, text)' "$folder standard registration uses recoverable atomic promotion"
    Assert-WorkflowMethodToken 'tracking-policy-concurrency' $policyStorePath 'LoadOrCreate' 'throw new InvalidDataException' "$folder corrupt shared policy is not silently replaced"
	Assert-WorkflowMethodExcludesToken 'project-element-change-ledger' $dashboardHostPath 'RunDashboardAction' 'project-element-change-tracking/' "$folder has no global tracking checkbox route"
	Assert-WorkflowToken 'project-element-change-ledger' $dashboardHostPath 'fbCurrentProjectHistoryTool' "$folder exposes current-project history in the sidebar"
	Assert-WorkflowToken 'project-element-change-ledger' $dashboardHostPath 'fbAllProjectHistoryTool' "$folder exposes all-project history in the sidebar"
	Assert-WorkflowToken 'project-element-change-ledger' $dashboardHostPath 'T("History", "이력 관리")' "$folder groups both history tools under a dedicated sidebar heading"
	Assert-WorkflowMethodToken 'element-tracking-session-recovery' $dashboardHostPath 'ResetAllFamilyBrowserSettings' 'FamilyBrowserElementChangeTrackingService.NotifyPolicyChanged(_workspaceRoot, false);' "$folder full settings reset immediately disables in-memory element tracking"
	Assert-WorkflowMethodToken 'element-tracking-session-recovery' $dashboardHostPath 'ResetAllFamilyBrowserSettings' 'Immutable element-change history and protected workshared local-save checkpoints are retained.' "$folder full settings reset discloses immutable history and checkpoint retention"
    Assert-WorkflowMethodOrder 'tracking-policy-concurrency' $dashboardHostPath 'ResetAllFamilyBrowserSettings' 'RecordProjectElementChangeTrackingPolicyChange(true, false, resetTrackingPolicyChangeId, "Prepared"' 'FamilyBrowserStandardPolicyStore.ResetToDefault' "$folder full reset protects the tracking-policy intent before disabling tracking"
    Assert-WorkflowMethodOrder 'tracking-policy-concurrency' $dashboardHostPath 'ResetAllFamilyBrowserSettings' 'FamilyBrowserStandardPolicyStore.ResetToDefault' 'RecordProjectElementChangeTrackingPolicyChange(true, false, resetTrackingPolicyChangeId, "Completed"' "$folder full reset records tracking-policy completion only after the policy reset succeeds"
    Assert-WorkflowMethodToken 'tracking-policy-concurrency' $dashboardHostPath 'ResetAllFamilyBrowserSettings' 'RecordProjectElementChangeTrackingPolicyChange(true, false, resetTrackingPolicyChangeId, "Failed"' "$folder failed full reset leaves an explicit failed tracking-policy audit record"
    Assert-WorkflowMethodToken 'tracking-policy-concurrency' $dashboardHostPath 'ResetAllFamilyBrowserSettings' 'trackingWasEnabled && !trackingPolicyResetApplied' "$folder auxiliary cleanup failures cannot falsely report that an already-applied tracking-policy reset failed"
    Assert-WorkflowMethodToken 'tracking-policy-concurrency' $dashboardHostPath 'ResetAllFamilyBrowserSettings' 'trackingWasEnabled && !trackingPolicyResetCompleted' "$folder missing completion audit is visibly reported after tracking is disabled"
    Assert-WorkflowMethodOrder 'tracking-policy-concurrency' $dashboardHostPath 'ResetAllFamilyBrowserSettings' 'GetActiveUncommittedSessionCount()' 'ResetToDefault(_workspaceRoot' "$folder full settings reset checks uncommitted tracking evidence before resetting the policy or deleting caches"
    Assert-WorkflowMethodExcludesToken 'tracking-policy-concurrency' $dashboardHostPath 'ResetAllFamilyBrowserSettings' 'trackingWasEnabled ? FamilyBrowserElementChangeTrackingService.GetActiveUncommittedSessionCount() : 0' "$folder full settings reset also blocks deferred remote-disable sessions after the shared policy is already off"
    Assert-WorkflowMethodToken 'tracking-policy-concurrency' $dashboardHostPath 'ResetAllFamilyBrowserSettings' 'Reset Blocked by Pending Tracking' "$folder full settings reset visibly blocks while uncommitted tracking sessions exist"
    Assert-WorkflowMethodExcludesToken 'element-tracking-session-recovery' $dashboardHostPath 'ResetAllFamilyBrowserSettings' 'ElementChangeHistory' "$folder full settings reset cannot delete immutable element history"
    Assert-WorkflowMethodExcludesToken 'element-tracking-session-recovery' $dashboardHostPath 'ResetAllFamilyBrowserSettings' 'OperationLogs' "$folder full settings reset cannot delete immutable tracking-policy audit history"
    Assert-WorkflowMethodExcludesToken 'element-tracking-session-recovery' $dashboardHostPath 'ResetAllFamilyBrowserSettings' 'DeleteElementSessionCheckpoint' "$folder full settings reset cannot delete protected local-save checkpoints"
    Assert-WorkflowMethodToken 'element-tracking-session-recovery' $machineConfigPath 'SetManagedPolicyPath' 'FamilyBrowserManagementContextLock.Acquire' "$folder serializes managed-path replacement across Revit processes"
    Assert-WorkflowMethodToken 'element-tracking-session-recovery' $machineConfigPath 'SetManagedPolicyPathNoLock' 'GetActiveUncommittedSessionCount()' "$folder refuses a managed-path replacement while this process owns uncommitted tracking evidence"
    Assert-WorkflowMethodOrder 'element-tracking-session-recovery' $machineConfigPath 'SetManagedPolicyPathNoLock' 'GetActiveUncommittedSessionCount()' 'Save(familyBrowserMachineConfig, currentUser)' "$folder checks active tracking evidence before saving a managed-path replacement"
    Assert-WorkflowMethodToken 'element-tracking-session-recovery' $machineConfigPath 'SetManagedPolicyPathNoLock' 'GetProtectedRecoverySessionCount()' "$folder refuses an automatic managed-path replacement while a live protected checkpoint awaits synchronization"
    Assert-WorkflowMethodToken 'element-tracking-session-recovery' $machineConfigPath 'SetManagedPolicyPathNoLock' 'HasBlockingElementSessionCheckpointForManagedPolicyPath(normalizedPolicyPath)' "$folder inspects on-disk checkpoints before automatic managed-path replacement"
    Assert-WorkflowMethodToken 'element-tracking-session-recovery' $machineConfigPath 'SetManagedPolicyPathNoLock' 'IsManagedFolderTransitionAuthorized()' "$folder permits checkpoint rebinding only through the verified migration controller"
    Assert-WorkflowMethodToken 'element-tracking-session-recovery' $machineConfigPath 'SetManagedPolicyPathNoLock' 'NotifyManagementContextChanged()' "$folder invalidates the tracking policy cache after a managed-path replacement"
    Assert-WorkflowToken 'project-element-change-ledger' $machineConfigPath 'last-known-managed-policy-path.txt' "$folder persists the last verified managed path as a cold-start hint"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $machineConfigPath 'TryRestoreLastKnownManagedPolicyPath' 'SetManagedPolicyPath(cachedPath, currentUser);' "$folder restores a verified managed path before model editing can begin"
    Assert-WorkflowMethodOrder 'project-element-change-ledger' $machineConfigPath 'SetManagedPolicyPathNoLock' 'Save(familyBrowserMachineConfig, currentUser)' 'RememberLastKnownManagedPolicyPath(normalizedPolicyPath);' "$folder remembers only a managed path that was accepted by the normal policy setter"
    Assert-WorkflowMethodToken 'element-tracking-session-recovery' $machineConfigPath 'ClearManagedPolicyPath' 'FamilyBrowserManagementContextLock.Acquire' "$folder serializes managed-path clearing across Revit processes"
    Assert-WorkflowMethodToken 'element-tracking-session-recovery' $machineConfigPath 'ClearManagedPolicyPathNoLock' 'GetActiveUncommittedSessionCount()' "$folder refuses to clear the managed path while this process owns uncommitted tracking evidence"
    Assert-WorkflowMethodToken 'element-tracking-session-recovery' $machineConfigPath 'ClearManagedPolicyPathNoLock' 'HasBlockingElementSessionCheckpointForManagedPolicyPath(string.Empty)' "$folder refuses to clear a destination while any protected local-save checkpoint remains"
    Assert-WorkflowMethodToken 'element-tracking-session-recovery' $machineConfigPath 'ClearManagedPolicyPathNoLock' 'NotifyManagementContextChanged()' "$folder invalidates the tracking policy cache after clearing the managed path"
	Assert-WorkflowMethodToken 'guard-admin-transition' $appPath 'HandleViewActivated' 'RunEventHandlerSafely("Native guard view-activation policy preload failed"' "$folder queues native guard policy before the Browser window opens"
    Assert-WorkflowMethodToken 'guard-admin-transition' $appPath 'QueueNativeGuardPolicyPreload' 'NativeGuardPolicyPreloadReady' "$folder skips cold policy preload only after a managed policy was actually resolved"
    Assert-WorkflowMethodToken 'guard-admin-transition' $appPath 'QueueNativeGuardPolicyPreload' 'FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot)' "$folder does not mark an unresolved management path as a completed preload"
    Assert-WorkflowMethodToken 'guard-admin-transition' $appPath 'QueueNativeGuardPolicyPreload' 'System.Threading.Interlocked.Exchange(ref NativeGuardPolicyPreloadStarted, 0);' "$folder releases the preload gate so an early unresolved attempt can be retried"
    Assert-WorkflowMethodToken 'guard-admin-transition' $appPath 'QueueNativeGuardPolicyPreload' 'FamilyBrowserDeploymentBootstrapService.TryApplyManagedPathOnly' "$folder resolves the homepage managed path without blocking Revit UI"
    Assert-WorkflowMethodToken 'guard-admin-transition' $appPath 'QueueNativeGuardPolicyPreload' 'Task.Run(delegate' "$folder keeps homepage/path probing off the Revit UI thread"
    Assert-WorkflowMethodToken 'guard-admin-transition' $appPath 'ApplyStartupPolicy' 'NotifyAdminModeChanged(enabled, policy, refreshUiNow: false)' "$folder schedules persisted Admin state through Revit Idling"
    Assert-WorkflowToken 'project-element-change-ledger' $appPath 'DocumentOpened += HandleDocumentOpened' "$folder captures a document baseline after open"
	Assert-WorkflowMethodOrder 'project-element-change-ledger' $appPath 'HandleDocumentOpened' 'RunEventHandlerSafely("Managed policy document-open preparation failed"' 'RunEventHandlerSafely("Element tracking document-open baseline failed"' "$folder resolves the managed policy before the first document baseline"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'EnsureManagedPolicyBeforeDocumentEditing' 'TryRestoreLastKnownManagedPolicyPath' "$folder restores the last verified management path during DocumentOpened"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'EnsureManagedPolicyBeforeDocumentEditing' 'FamilyBrowserDeploymentBootstrapService.TryApplyManagedPathOnly' "$folder performs a blocking first-run path resolution before returning an editable document"
    Assert-WorkflowMethodOrder 'project-element-change-ledger' $appPath 'EnsureManagedPolicyBeforeDocumentEditing' 'FamilyBrowserDeploymentBootstrapService.TryApplyManagedPathOnly' 'ApplyStartupPolicy(workspaceRoot, currentUser);' "$folder applies the resolved shared policy before baseline creation"
	Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'HandleDocumentOpened' 'RunEventHandlerSafely("Native guard document-open policy preload failed"' "$folder refreshes the homepage-managed path after the protected DocumentOpened baseline"
	Assert-WorkflowMethodToken 'current-model-check-baseline' $appPath 'HandleDocumentClosing' 'RunEventHandlerSafely("Automatic Current Model Check document-closing cleanup failed"' "$folder isolates automatic-check cleanup during document close"
	Assert-WorkflowMethodOrder 'external-project-change-detection' $appPath 'ObserveProjectCatalogAfterCommit' 'FamilyBrowserProjectCatalogService.IsPublishedObservationState' 'FamilyBrowserElementChangeTrackingService.RestoreProjectCatalogObservationRequired' "$folder restores a consumed catalog decision when publication is deferred or fails"
    Assert-WorkflowMethodOrder 'project-element-change-ledger' $dashboardHostPath 'TryGetCentralPath' 'if (!doc.IsWorkshared)' 'doc.GetWorksharingCentralModelPath()' "$folder never asks a standalone RVT for a worksharing central path"
    Assert-WorkflowMethodOrder 'standard-edit-commit-lifecycle' $servicePath 'BuildDocumentComparePaths' 'if (doc.IsWorkshared)' 'doc.GetWorksharingCentralModelPath()' "$folder resolves a Standard RVT central path only for a workshared document"
    Assert-WorkflowToken 'project-element-change-ledger' $appPath 'Idling += HandleIdling' "$folder can establish a tracking baseline after the managed path finishes loading"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'ApplyStartupPolicy' 'FamilyBrowserElementChangeTrackingService.NotifyPolicyChanged(workspaceRoot, trackingEnabled)' "$folder managed-path preload immediately informs element tracking"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $dashboardHostPath 'ApplyAdminModeAfterPolicyLoad' 'SynchronizeElementTrackingPolicy("policy-load:" + source);' "$folder every completed dashboard policy load queues element tracking before the user resumes model work"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $dashboardHostPath 'SynchronizeElementTrackingPolicy' 'RequestDocumentSessionBaselineRefresh' "$folder modeless policy changes defer baseline capture to a valid Revit Idling context"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $dashboardHostPath 'ConfigureFileGuardPolicy' 'SynchronizeElementTrackingPolicy("file-guard-save");' "$folder saving per-file tracking scope immediately queues a fresh baseline"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'HandleIdling' 'ConsumeDocumentSessionBaselineRefreshRequest' "$folder consumes dashboard baseline requests on the Revit UI event"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'HandleIdling' 'FamilyBrowserElementChangeTrackingService.BeginDocumentSession(workspaceRoot, document)' "$folder baseline capture is returned to the Revit UI thread"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'HandleIdling' 'FamilyBrowserElementChangeTrackingService.HasDocumentSession(document)' "$folder verifies that baseline creation really succeeded before ending retries"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'HandleDocumentSaving' 'PrepareDocumentCommit' "$folder Save prepares or recovers the element tracking session before the commit boundary"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'HandleDocumentSavingAs' 'PrepareDocumentCommit' "$folder Save As prepares or recovers the element tracking session before the commit boundary"
    Assert-WorkflowToken 'project-element-change-ledger' $appPath 'StartReloadLatestBridge(application.ControlledApplication' "$folder attaches Reload Latest tracking protection"
    Assert-WorkflowToken 'project-element-change-ledger' $appPath 'DocumentSynchronizingWithCentral += HandleDocumentSynchronizingWithCentral' "$folder marks the local synchronization boundary"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'HandleDocumentChanged' 'FamilyBrowserElementChangeTrackingService.HandleDocumentChanged' "$folder forwards Revit element changes to the ledger"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'HandleDocumentClosing' 'FamilyBrowserElementChangeTrackingService.HandleDocumentClosing' "$folder close-start preserves state until closure is final"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'HandleDocumentClosed' 'FamilyBrowserElementChangeTrackingService.HandleDocumentClosed' "$folder actual close releases only the closed document session"
    Assert-WorkflowToken 'project-element-change-ledger' $appPath 'RunEventHandlerSafely' "$folder isolates lifecycle services so one failure cannot block element history commit"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'HandleDocumentSynchronizedWithCentral' 'HandleDocumentSynchronizationCompletionFailure' "$folder closes synchronization suppression when completion document or status access fails"
    Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath 'HandleDocumentSynchronizingWithCentral' 'HandleDocumentSynchronizationStartFailure' "$folder opens conservative synchronization suppression when the start document is unreadable"
    foreach ($saveEventMethod in @('HandleDocumentSaved', 'HandleDocumentSavedAs')) {
        Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath $saveEventMethod 'HandleDocumentSaveCompletionFailure' "$folder preserves uncertainty when $saveEventMethod cannot read its document or status"
    }
    foreach ($eventMethod in @('HandleDocumentSaved', 'HandleDocumentSavedAs', 'HandleDocumentSynchronizedWithCentral')) {
        Assert-WorkflowMethodToken 'project-element-change-ledger' $appPath $eventMethod 'FamilyBrowserElementChangeTrackingService.HandleDocumentCommitted' "$folder $eventMethod evaluates the element-change commit boundary"
    }

    foreach ($eventMethod in @('HandleDocumentSaved', 'HandleDocumentSavedAs', 'HandleDocumentSynchronizedWithCentral')) {
        $operationCommitToken = if ($eventMethod -eq 'HandleDocumentSynchronizedWithCentral') { 'StandardRvtChangeCandidateService.HandleDocumentSynchronizedWithCentral' } else { 'StandardRvtChangeCandidateService.HandleDocumentSaved' }
        Assert-WorkflowMethodOrder 'external-project-change-detection' $appPath $eventMethod $operationCommitToken 'ObserveProjectCatalogAfterCommit' "$folder $eventMethod commits Browser operation attribution before conditional catalog comparison"
        Assert-WorkflowMethodOrder 'external-project-change-detection' $appPath $eventMethod 'FamilyBrowserProjectCatalogService.IsSuccessfulRevitEventStatus(status)' 'ObserveProjectCatalogAfterCommit' "$folder $eventMethod considers names only after a successful commit"
        Assert-WorkflowMethodOrder 'external-project-change-detection' $appPath $eventMethod 'ObserveProjectCatalogAfterCommit' 'FamilyBrowserDashboardModelessRuntime.NotifyDocumentCommitFinalized' "$folder $eventMethod refreshes the Browser after the conditional catalog decision"
    }
    Assert-WorkflowMethodOrder 'external-project-change-detection' $appPath 'ObserveProjectCatalogAfterCommit' 'ConsumeProjectCatalogObservationRequired' 'FamilyBrowserProjectCatalogService.Observe' "$folder evaluates the tracking decision before scanning the project catalog"
    Assert-WorkflowMethodToken 'external-project-change-detection' $appPath 'ObserveProjectCatalogAfterCommit' 'if (performed)' "$folder skips the full family/type catalog scan for ordinary instance-only commits"
    Assert-WorkflowMethodToken 'external-project-change-detection' $appPath 'ObserveProjectCatalogAfterCommit' 'RecordProjectCatalogObservationPerformance' "$folder records whether the catalog scan ran and how long it took"
}

$harnessProject = Join-Path $repoRoot 'KKY_FamilyBrowser_Automation\KKY_FamilyBrowser_WorkflowAuditHarness\KKY_FamilyBrowser_WorkflowAuditHarness.csproj'
$harnessOutput = Join-Path $OutputDir 'tracking-harness'
& dotnet run --project $harnessProject -c Release -- $harnessOutput
$harnessPassed = $LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath (Join-Path $harnessOutput 'tracking-persistence-summary.json'))
Add-Check 'actual-source tracking persistence harness' $harnessPassed "Harness exit code: $LASTEXITCODE"

$trackingHarnessSummary = $null
if ($harnessPassed) {
    $trackingHarnessSummary = Get-Content -Raw -Encoding UTF8 -LiteralPath (Join-Path $harnessOutput 'tracking-persistence-summary.json') | ConvertFrom-Json
}
$trackingHarnessChecks = if ($trackingHarnessSummary) { @($trackingHarnessSummary.checks) } else { @() }
Add-WorkflowCheck 'managed-folder-homepage-return' 'management-context transition lock excludes another Revit process' ($trackingHarnessChecks -contains 'management-context changes are serialized across Revit processes and retry after release') 'Tracking harness did not prove management-context lease exclusion and retry.'
Add-WorkflowCheck 'offline-tracking-recovery' 'offline tracking writes a durable local spool' ($trackingHarnessChecks -contains 'offline operations are write-ahead spooled') 'Tracking harness did not prove write-ahead spooling.'
Add-WorkflowCheck 'offline-tracking-recovery' 'reconnect flushes immutable records and clears the spool' ($trackingHarnessChecks -contains 'reconnected managed folder flushes and clears local spool') 'Tracking harness did not prove reconnect recovery.'
Add-WorkflowCheck 'offline-tracking-recovery' 'replayed pending records are idempotent' ($trackingHarnessChecks -contains 'stable entry IDs make replay idempotent') 'Tracking harness did not prove idempotent replay.'
Add-WorkflowCheck 'project-element-change-ledger' 'offline element changes use the write-ahead spool' ($trackingHarnessChecks -contains 'offline element changes are write-ahead spooled') 'Tracking harness did not prove durable offline element-change spooling.'
Add-WorkflowCheck 'project-element-change-ledger' 'element changes flush into per-project immutable history' ($trackingHarnessChecks -contains 'element changes flush into per-project immutable history') 'Tracking harness did not prove per-project immutable element history.'
Add-WorkflowCheck 'project-element-change-ledger' 'parallel element commits remain complete' ($trackingHarnessChecks -contains 'parallel element-change commits remain complete') 'Tracking harness did not prove concurrent element-change retention.'
Add-WorkflowCheck 'project-element-change-ledger' 'unidentified events persist as explicit coverage gaps' ($trackingHarnessChecks -contains 'unidentified DocumentChanged coverage gaps persist without fabricated element IDs and are integrity protected') 'Tracking harness did not prove durable coverage-gap evidence.'
Add-WorkflowCheck 'project-element-change-ledger' 'external-update rebase failures persist as explicit zero-row coverage gaps' ($trackingHarnessChecks -contains 'external-update rebase failures persist as zero-row coverage gaps without fabricated element IDs') 'Tracking harness did not prove durable external-update rebase-gap evidence.'
Add-WorkflowCheck 'element-tracking-integrity-recovery' 'empty non-coverage commits fail closed' ($trackingHarnessChecks -contains 'empty non-coverage commits fail closed instead of deleting or hiding evidence') 'Tracking harness did not reject malformed empty element commits.'
Add-WorkflowCheck 'element-tracking-session-recovery' 'coverage gaps survive local-save checkpoint recovery' ($trackingHarnessChecks -contains 'coverage-gap-only evidence survives workshared local-save checkpoint recovery') 'Tracking harness did not preserve coverage gaps through checkpoint recovery.'
Add-WorkflowCheck 'element-tracking-session-recovery' 'automatic path replacement respects protected checkpoints' ($trackingHarnessChecks -contains 'automatic management-folder replacement is blocked by protected local-save evidence') 'Tracking harness did not prove automatic managed-path replacement is blocked.'
Add-WorkflowCheck 'element-tracking-session-recovery' 'path replacement fails closed on checkpoint lock contention' ($trackingHarnessChecks -contains 'management-folder replacement fails closed while checkpoint state is locked') 'Tracking harness did not prove locked checkpoint state blocks a path replacement.'
Add-WorkflowCheck 'element-tracking-session-recovery' 'invalid empty checkpoint input preserves existing evidence' ($trackingHarnessChecks -contains 'invalid empty checkpoint commits cannot delete existing protected evidence') 'Tracking harness did not prove malformed checkpoint input is non-destructive.'
Add-WorkflowCheck 'element-tracking-integrity-recovery' 'pending records do not leak across management roots' ($trackingHarnessChecks -contains 'pending tracking records cannot leak into another management root without explicit migration') 'Tracking harness did not prove destination binding.'
Add-WorkflowCheck 'element-tracking-integrity-recovery' 'migration rebind is source scoped' ($trackingHarnessChecks -contains 'managed-folder migration rebinds only records from the selected source root') 'Tracking harness did not prove source-scoped migration.'
Add-WorkflowCheck 'element-tracking-integrity-recovery' 'pending destination envelope detects tampering' ($trackingHarnessChecks -contains 'pending element destination metadata is checksum protected') 'Tracking harness did not prove pending-envelope integrity.'
Add-WorkflowCheck 'element-tracking-integrity-recovery' 'operation and candidate envelopes detect tampering' ($trackingHarnessChecks -contains 'pending operation and standard-candidate destinations are checksum protected') 'Tracking harness did not prove operation/candidate pending-envelope integrity.'
Add-WorkflowCheck 'element-tracking-integrity-recovery' 'local spool cleanup failures remain retryable and block settlement' ($trackingHarnessChecks -contains 'managed-folder transitions can detect and retry local spool cleanup failures') 'Tracking harness did not prove cleanup failure visibility.'
Add-WorkflowCheck 'element-tracking-session-recovery' 'checkpoint enumeration failure remains blocking' ($trackingHarnessChecks -contains 'checkpoint enumeration failures remain blocking and cannot masquerade as trusted zero pending work') 'Tracking harness did not prove fail-closed checkpoint enumeration.'
Add-WorkflowCheck 'project-element-change-ledger' 'history enumeration failure remains visible' ($trackingHarnessChecks -contains 'immutable-history enumeration failures are explicit and recover after storage access returns') 'Tracking harness did not prove immutable-history enumeration failure visibility and recovery.'
Add-WorkflowCheck 'element-tracking-integrity-recovery' 'tampered immutable history is rejected' ($trackingHarnessChecks -contains 'element history checksum rejects tampered records') 'Tracking harness did not prove immutable-history checksum validation.'
Add-WorkflowCheck 'element-tracking-integrity-recovery' 'corrupt collisions preserve local evidence' ($trackingHarnessChecks -contains 'corrupt destination collisions preserve the valid local write-ahead copy') 'Tracking harness did not prove collision recovery.'
Add-WorkflowCheck 'element-tracking-integrity-recovery' 'history ordering uses committed time' ($trackingHarnessChecks -contains 'recent element history is ordered by committed time rather than file copy time') 'Tracking harness did not prove committed-time ordering.'
Add-WorkflowCheck 'element-tracking-integrity-recovery' 'operation and candidate ordering uses record time' ($trackingHarnessChecks -contains 'recent operation and standard-candidate history is ordered by record time after migration or recovery') 'Tracking harness did not prove operation/candidate record-time ordering.'
Add-WorkflowCheck 'tracking-policy-session-isolation' 'disabled or fallback policy cannot start another document session' ($trackingHarnessChecks -contains 'policy disable and read-fallback preserve existing evidence without starting collection in another document') 'Tracking harness did not prove cross-document policy isolation.'
Add-WorkflowCheck 'element-tracking-integrity-recovery' 'checkpoint identities precede deduplication' ($trackingHarnessChecks -contains 'checkpoint commits receive identities before deduplication') 'Tracking harness did not prove that blank entry IDs remain separate.'
Add-WorkflowCheck 'element-tracking-integrity-recovery' 'checkpoint entry-ID collision fails closed' ($trackingHarnessChecks -contains 'conflicting checkpoint entry-ID collisions fail closed') 'Tracking harness did not prove conflicting checkpoint collision rejection.'
Add-WorkflowCheck 'element-tracking-integrity-recovery' 'checkpoint inner records require integrity' ($trackingHarnessChecks -contains 'checkpoint inner commits must carry valid integrity evidence') 'Tracking harness did not prove unsigned inner checkpoint rejection.'
Add-WorkflowCheck 'element-tracking-session-recovery' 'workshared local Save survives restart until synchronization' ($trackingHarnessChecks -contains 'workshared local saves survive restart and publish only after successful synchronization') 'Tracking harness did not prove workshared session recovery.'
Add-WorkflowCheck 'element-tracking-session-recovery' 'checkpoint status separates sync and history-promotion states' ($trackingHarnessChecks -contains 'checkpoint status distinguishes synchronization pending from synchronized history-promotion pending') 'Tracking harness did not prove finalized checkpoint state classification.'
Add-WorkflowCheck 'element-tracking-session-recovery' 'ordinary refresh promotes finalized checkpoints' ($trackingHarnessChecks -contains 'ordinary refresh promotes finalized checkpoints without another synchronization') 'Tracking harness did not prove finalized checkpoint replay from the normal pending flush.'
Add-WorkflowCheck 'element-tracking-session-recovery' 'checkpoint rejects another project commit' ($trackingHarnessChecks -contains 'local checkpoints reject commits bound to another project') 'Tracking harness did not prove checkpoint-to-project binding.'
Add-WorkflowCheck 'element-tracking-schema-compatibility' 'integrity-v1 history survives schema evolution' ($trackingHarnessChecks -contains 'integrity-v1 element history remains verifiable after schema evolution') 'Tracking harness did not prove integrity-v1 compatibility.'
Add-WorkflowCheck 'element-tracking-schema-compatibility' 'integrity-v1 history replay remains idempotent' ($trackingHarnessChecks -contains 'already-persisted integrity-v1 history replays idempotently after schema evolution') 'Tracking harness did not prove old-history replay idempotency.'
Add-WorkflowCheck 'element-tracking-schema-compatibility' 'pending-envelope-v1 survives schema evolution' ($trackingHarnessChecks -contains 'integrity-v1 pending element envelopes survive commit schema evolution') 'Tracking harness did not prove pending-envelope-v1 compatibility.'
Add-WorkflowCheck 'element-tracking-schema-compatibility' 'invalid signed commit is not re-signed' ($trackingHarnessChecks -contains 'checksum-invalid signed commits are rejected instead of silently re-signed') 'Tracking harness did not prove invalid signed commit rejection.'
Add-WorkflowCheck 'element-tracking-schema-compatibility' 'future schema fails closed' ($trackingHarnessChecks -contains 'future element-history schemas fail closed instead of being partially interpreted') 'Tracking harness did not prove future element-history schema rejection.'
Add-WorkflowCheck 'element-tracking-schema-compatibility' 'central replacement cannot hide path-matching old history' ($trackingHarnessChecks -contains 'stable path fallback recovers history after the central file identity changes') 'Tracking harness did not prove legacy-root discovery.'
Add-WorkflowCheck 'element-tracking-schema-compatibility' 'corrupt local checkpoint is surfaced' ($trackingHarnessChecks -contains 'corrupt local checkpoints are surfaced instead of silently discarded') 'Tracking harness did not prove corrupt-checkpoint reporting.'
Assert-WorkflowScenario 'offline-tracking-recovery' 'admin-home-offline-tracking-pending'
Assert-WorkflowToken 'offline-tracking-recovery' $projectCatalogDashboardPath 'pendingTrackingPill' 'offline queue is visible in the header'
Assert-WorkflowToken 'offline-tracking-recovery' $projectCatalogDashboardPath 'pendingTrackingQueue' 'offline queue has an explanatory Home board'
Assert-WorkflowToken 'offline-tracking-recovery' $uiHarnessPath 'CheckPendingTrackingQueue(browser, options, result);' 'IE harness validates pending queue visibility and retry action'
Add-WorkflowCheck 'multi-client-tracking' 'parallel managed-folder writers retain every operation' ($trackingHarnessChecks -contains 'parallel writers retain every uniquely identified operation') 'Tracking harness did not prove parallel-writer retention.'
Add-WorkflowCheck 'multi-client-tracking' 'immutable history remains complete after concurrent writes' ($trackingHarnessChecks -contains 'immutable history is readable and complete') 'Tracking harness did not prove immutable-history completeness.'

foreach ($workflowId in $requiredWorkflowIds) {
    $evidenceCount = @($checks | Where-Object { $_.WorkflowId -eq $workflowId }).Count
    Add-Check "workflow $workflowId has scenario-specific evidence" ($evidenceCount -gt 0) 'Every workflow must have at least one dedicated source, fixture, contract or IE check.'
}

$passed = $failures.Count -eq 0
$workflowResults = @($contract.workflows | ForEach-Object {
    $workflow = $_
    $workflowChecks = @($checks | Where-Object { $_.WorkflowId -eq [string]$workflow.id })
    $workflowPassed = $workflowChecks.Count -gt 0 -and @($workflowChecks | Where-Object { -not $_.Passed }).Count -eq 0
    [pscustomobject]@{
        Id = $workflow.id
        Area = $workflow.area
        Automation = $workflow.automation
        AutomatedStatus = $(if ($workflowChecks.Count -eq 0) { 'NO_EVIDENCE' } elseif ($workflowPassed) { 'PASS' } else { 'FAIL' })
        EvidenceCount = $workflowChecks.Count
        RuntimeStatus = $(if ([bool]$workflow.needsRevit) { 'NEEDS_REVIT_CHECK' } else { 'AUTOMATED' })
        Sequence = $workflow.sequence
        Expected = $workflow.expected
    }
})

$checkArray = @($checks.ToArray())
$failureArray = @($failures.ToArray())
$summary = [ordered]@{
    generatedAt = (Get-Date).ToString('o')
    status = $(if ($passed) { 'PASS' } else { 'FAIL' })
    contract = $contractPath
    workflowCount = $workflowResults.Count
    checkCount = $checks.Count
    passedCheckCount = @($checkArray | Where-Object Passed).Count
    failures = $failureArray
    workflows = $workflowResults
    checks = $checkArray
}
$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDir 'workflow-audit-summary.json') -Encoding UTF8

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('# Family Browser Workflow Audit') | Out-Null
$lines.Add('') | Out-Null
$lines.Add("- Generated: $($summary.generatedAt)") | Out-Null
$lines.Add("- Status: $($summary.status)") | Out-Null
$lines.Add("- Workflows: $($summary.workflowCount)") | Out-Null
$lines.Add("- Checks: $($summary.passedCheckCount)/$($summary.checkCount)") | Out-Null
$lines.Add('') | Out-Null
$lines.Add('| Workflow | Area | Automated | Evidence | Runtime |') | Out-Null
$lines.Add('|---|---|---:|---:|---:|') | Out-Null
foreach ($workflow in $workflowResults) {
    $lines.Add("| $($workflow.Id) | $($workflow.Area) | $($workflow.AutomatedStatus) | $($workflow.EvidenceCount) | $($workflow.RuntimeStatus) |") | Out-Null
}
if ($failures.Count -gt 0) {
    $lines.Add('') | Out-Null
    $lines.Add('## Failures') | Out-Null
    foreach ($failure in $failures) {
        $lines.Add("- $failure") | Out-Null
    }
}
$lines | Set-Content -LiteralPath (Join-Path $OutputDir 'workflow-audit-summary.md') -Encoding UTF8

if (-not $passed) {
    Write-Error "Family Browser workflow audit failed. See $OutputDir"
    exit 1
}

Write-Host "Family Browser workflow audit passed: $OutputDir" -ForegroundColor Green
