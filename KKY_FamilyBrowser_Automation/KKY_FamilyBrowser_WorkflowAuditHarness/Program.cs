using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

public static class Program
{
    private static int Main(string[] args)
    {
        string output = args.Length > 0 ? Path.GetFullPath(args[0]) : Path.Combine(Path.GetTempPath(), "kky-family-browser-workflow-audit");
        string fixture = Path.Combine(output, "tracking-fixture");
        if (Directory.Exists(fixture))
        {
            Directory.Delete(fixture, true);
        }
        Directory.CreateDirectory(fixture);

        string managed = Path.Combine(fixture, "managed");
        string spool = Path.Combine(fixture, "spool");
        FamilyBrowserStandardPolicyStore.ManagedRoot = managed;
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = spool;

        List<string> checks = new List<string>();
        System.Reflection.MethodInfo elementIntegrityMethod = typeof(FamilyBrowserTrackingPersistenceService).GetMethod("ComputeElementChangeIntegrity", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        System.Reflection.MethodInfo pendingEnvelopeIntegrityMethod = typeof(FamilyBrowserTrackingPersistenceService).GetMethod("ComputePendingElementEnvelopeIntegrity", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        IDisposable managementContextBlocker = FamilyBrowserManagementContextLock.Acquire(TimeSpan.FromSeconds(1.0));
        bool managementContextTimedOut = Task.Run(delegate
        {
            try
            {
                using (FamilyBrowserManagementContextLock.Acquire(TimeSpan.FromMilliseconds(150.0)))
                {
                    return false;
                }
            }
            catch (IOException)
            {
                return true;
            }
        }).GetAwaiter().GetResult();
        Assert(managementContextTimedOut, "a second Revit process entered the management-context transition while the first process still held the lease");
        managementContextBlocker.Dispose();
        bool managementContextRetried = Task.Run(delegate
        {
            using (FamilyBrowserManagementContextLock.Acquire(TimeSpan.FromSeconds(1.0)))
            {
                return true;
            }
        }).GetAwaiter().GetResult();
        Assert(managementContextRetried, "management-context transition did not recover after the first process released its lease");
        checks.Add("management-context changes are serialized across Revit processes and retry after release");

        List<FamilyBrowserElementActivityMatchInput> groupedUndoCandidates = new List<FamilyBrowserElementActivityMatchInput>
        {
            ActivityMatch(new[] { "2001" }, new[] { "Place wall" }),
            ActivityMatch(new[] { "2002" }, new[] { "Place door" })
        };
        FamilyBrowserElementActivityMatchResult groupedUndoMatch = FamilyBrowserElementActivityMatcher.Match(
            groupedUndoCandidates,
            ActivityMatch(new[] { "2001", "2002" }, new[] { "Place wall", "Place door" }));
        Assert(groupedUndoMatch.Exact && groupedUndoMatch.CandidateIndexes.SequenceEqual(new[] { 1, 0 }), "a grouped Undo did not consume the complete contiguous LIFO activity suffix");
        FamilyBrowserElementActivityMatchResult sameElementUndoMatch = FamilyBrowserElementActivityMatcher.Match(
            new[]
            {
                ActivityMatch(new[] { "2010" }, new[] { "Move first" }),
                ActivityMatch(new[] { "2010" }, new[] { "Move second" })
            },
            ActivityMatch(new[] { "2010" }, new[] { "Move second" }));
        Assert(sameElementUndoMatch.Exact && sameElementUndoMatch.CandidateIndexes.SequenceEqual(new[] { 1 }), "same-element Undo did not select the newest matching activity");
        FamilyBrowserElementActivityMatchResult ambiguousUndoMatch = FamilyBrowserElementActivityMatcher.Match(
            groupedUndoCandidates,
            ActivityMatch(new[] { "2001", "2999" }, new[] { "Place wall", "Unknown external transaction" }));
        Assert(!ambiguousUndoMatch.Exact && ambiguousUndoMatch.CandidateIndexes.Count == 1, "a partial Undo match was incorrectly treated as exact");
        FamilyBrowserElementActivityMatchResult repeatedEvidenceUndoMatch = FamilyBrowserElementActivityMatcher.Match(
            new[]
            {
                ActivityMatch(new[] { "2020" }, new[] { "Move" }),
                ActivityMatch(new[] { "2020" }, new[] { "Move" })
            },
            ActivityMatch(new[] { "2020" }, new[] { "Move" }));
        Assert(!repeatedEvidenceUndoMatch.Exact && repeatedEvidenceUndoMatch.CandidateIndexes.SequenceEqual(new[] { 1, 0 }), "indistinguishable repeated same-element transactions were incorrectly attributed to one Undo activity");
        FamilyBrowserElementActivityMatchResult unnamedRepeatedUndoMatch = FamilyBrowserElementActivityMatcher.Match(
            new[]
            {
                ActivityMatch(new[] { "2030" }, new[] { "First action" }),
                ActivityMatch(new[] { "2030" }, new[] { "Second action" })
            },
            ActivityMatch(new[] { "2030" }, Array.Empty<string>()));
        Assert(!unnamedRepeatedUndoMatch.Exact && unnamedRepeatedUndoMatch.CandidateIndexes.SequenceEqual(new[] { 1, 0 }), "an Undo callback without transaction names incorrectly selected one of several same-element activities");
        checks.Add("grouped Undo and Redo matching consumes only an exact contiguous LIFO activity suffix");
        checks.Add("partial Undo and Redo matching remains explicitly ambiguous");
        checks.Add("indistinguishable repeated Undo and Redo evidence never mutates one guessed activity");

        Assert(FamilyBrowserElementTrackingTransitionPolicy.ShouldIgnoreChangedElement(true, false, false, false),
            "a resolved auxiliary element was allowed into activity bookkeeping");
        Assert(FamilyBrowserElementTrackingTransitionPolicy.ShouldIgnoreChangedElement(false, true, false, true),
            "a null live element already known by the session auxiliary index was allowed into activity bookkeeping");
        Assert(!FamilyBrowserElementTrackingTransitionPolicy.ShouldIgnoreChangedElement(false, false, false, true),
            "a live element was hidden only because its ID had previously belonged to an auxiliary element");
        Assert(FamilyBrowserElementTrackingTransitionPolicy.ResolveChangeKind(false, true, false, true, false, true, false, false) == "Created",
            "the transition policy did not classify a newly visible element as Created");
        Assert(FamilyBrowserElementTrackingTransitionPolicy.ResolveChangeKind(true, true, false, true, false, false, false, false) == "Modified",
            "the transition policy did not classify an active existing element as Modified");
        Assert(FamilyBrowserElementTrackingTransitionPolicy.ResolveChangeKind(true, false, false, true, false, false, true, false) == "Deleted",
            "the transition policy did not classify a removed baseline element as Deleted");
        Assert(FamilyBrowserElementTrackingTransitionPolicy.ResolveChangeKind(false, false, false, true, false, true, true, false) == "CreatedThenDeleted",
            "the transition policy lost a same-boundary create/delete sequence");
        Assert(string.IsNullOrEmpty(FamilyBrowserElementTrackingTransitionPolicy.ResolveChangeKind(false, false, true, true, false, true, true, false)),
            "an ambiguous same-boundary create/delete sequence was reported as exact");
        Assert(string.IsNullOrEmpty(FamilyBrowserElementTrackingTransitionPolicy.ResolveChangeKind(true, true, true, true, false, false, false, false)),
            "an ambiguous stale activity fabricated a modification without a state-signature change");
        Assert(FamilyBrowserElementTrackingTransitionPolicy.ResolveChangeKind(true, true, true, true, true, false, false, false) == "Modified",
            "an actual state-signature change was lost only because the event sequence was ambiguous");
        Assert(FamilyBrowserElementTrackingTransitionPolicy.IsUnresolvedTransient("CreatedThenDeleted", false, false, false),
            "a same-boundary create/delete sequence without any captured metadata was not marked unresolved");
        HashSet<string> currentEventIgnoredIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "reused-visible-id" };
        HashSet<string> sessionIgnoredIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "reused-visible-id" };
        FamilyBrowserElementTrackingTransitionPolicy.RestoreVisibleElementId(
            "reused-visible-id",
            currentEventIgnoredIds,
            sessionIgnoredIds);
        Assert(!currentEventIgnoredIds.Contains("reused-visible-id") && !sessionIgnoredIds.Contains("reused-visible-id"),
            "an ID recaptured as visible remained hidden by a stale auxiliary-element index");
        checks.Add("element transition decisions execute add, modify, delete, null-element, and same-boundary transient cases");

        FamilyBrowserTrackedElementState visibleState = CreateTrackedState("scope-visible", "after");
        FamilyBrowserTrackedElementState auxiliaryState = CreateTrackedState("scope-auxiliary", "after");
        auxiliaryState.ElementClass = "PipingSystem";
        auxiliaryState.CategoryName = "Piping Systems";
        auxiliaryState.CategoryId = "-2008043";
        FamilyBrowserTrackedElementState electricalSystemState = CreateTrackedState("scope-electrical-system", "after");
        electricalSystemState.ElementClass = "ElectricalSystem";
        electricalSystemState.CategoryName = "Electrical Circuits";
        electricalSystemState.CategoryId = "-2008037";
        FamilyBrowserTrackedElementState systemTypeState = CreateTrackedState("scope-system-type", "after");
        systemTypeState.ElementClass = "PipingSystemType";
        systemTypeState.CategoryName = "Piping Systems";
        systemTypeState.CategoryId = "-2008043";
        systemTypeState.IsElementType = true;
        List<FamilyBrowserElementChangeItem> projectionFixture = new List<FamilyBrowserElementChangeItem>
        {
            new FamilyBrowserElementChangeItem { ChangeKind = "Created", ElementId = visibleState.ElementId, After = visibleState },
            new FamilyBrowserElementChangeItem { ChangeKind = "Created", ElementId = auxiliaryState.ElementId, After = auxiliaryState },
            new FamilyBrowserElementChangeItem { ChangeKind = "Created", ElementId = electricalSystemState.ElementId, After = electricalSystemState },
            new FamilyBrowserElementChangeItem { ChangeKind = "Created", ElementId = systemTypeState.ElementId, After = systemTypeState },
            new FamilyBrowserElementChangeItem
            {
                ChangeKind = "CreatedThenDeleted",
                ElementId = "scope-unresolved",
                ElementClass = FamilyBrowserElementHistoryProjectionPolicy.UnresolvedTransientElementClass,
                TrackingKind = FamilyBrowserElementHistoryProjectionPolicy.UnresolvedTransientTrackingKind
            },
            new FamilyBrowserElementChangeItem { ChangeKind = "CreatedThenDeleted", ElementId = "scope-legacy-blank" }
        };
        FamilyBrowserElementHistoryProjectionCounts projectionCounts = FamilyBrowserElementHistoryProjectionPolicy.CountUserFacingChanges(projectionFixture);
        Assert(projectionCounts.VisibleChangeCount == 2 && projectionCounts.CreatedCount == 2 && projectionCounts.DeletedCount == 0,
            "the user-facing projection did not retain the ordinary instance and System Type definition only");
        Assert(projectionCounts.HiddenAuxiliaryCount == 2 && projectionCounts.HiddenUnresolvedTransientCount == 2,
            "the user-facing projection did not separate auxiliary and unresolved transient evidence");
        checks.Add("history projection hides known support rows and unresolved transients while preserving ordinary objects and System Type definitions");

        Assert(FamilyBrowserElementTrackingPolicyDecision.Resolve(true, false, false, false, false) == FamilyBrowserElementTrackingSessionMode.Live,
            "an authoritative enabled policy did not permit a new live session");
        Assert(FamilyBrowserElementTrackingPolicyDecision.Resolve(true, true, false, false, false) == FamilyBrowserElementTrackingSessionMode.Ignore,
            "a cached or deferred enabled state permitted a new document to start collecting evidence");
        Assert(FamilyBrowserElementTrackingPolicyDecision.Resolve(true, true, true, true, false) == FamilyBrowserElementTrackingSessionMode.DeferredCommit,
            "existing uncommitted evidence was not retained through a cached or deferred policy boundary");
        Assert(FamilyBrowserElementTrackingPolicyDecision.Resolve(false, false, true, true, false) == FamilyBrowserElementTrackingSessionMode.DeferredCommit,
            "an authoritative disable discarded evidence already observed before the policy boundary");
        Assert(FamilyBrowserElementTrackingPolicyDecision.Resolve(false, false, true, false, true) == FamilyBrowserElementTrackingSessionMode.RecoveryOnly,
            "a disabled policy did not preserve a protected workshared recovery checkpoint");
        Assert(FamilyBrowserElementTrackingPolicyDecision.Resolve(false, false, true, false, false) == FamilyBrowserElementTrackingSessionMode.Ignore,
            "a disabled policy kept an inactive live session eligible for future collection");
        checks.Add("policy disable and read-fallback preserve existing evidence without starting collection in another document");

        Assert(FamilyBrowserTrackingCommitOptimizationPolicy.ResolveBaselineMode(true, false, 0, 0) == FamilyBrowserPostCommitBaselineMode.Incremental,
            "a complete post-commit delta unnecessarily required a full model recapture");
        Assert(FamilyBrowserTrackingCommitOptimizationPolicy.ResolveBaselineMode(false, false, 0, 0) == FamilyBrowserPostCommitBaselineMode.FullCapture,
            "a failed incremental refresh did not fall back to a full model recapture");
        Assert(FamilyBrowserTrackingCommitOptimizationPolicy.ResolveBaselineMode(true, true, 0, 0) == FamilyBrowserPostCommitBaselineMode.FullCapture,
            "an external-update rebase gap incorrectly used the incremental baseline");
        Assert(FamilyBrowserTrackingCommitOptimizationPolicy.ResolveBaselineMode(true, false, 1, 0) == FamilyBrowserPostCommitBaselineMode.FullCapture,
            "an element-ID event read gap incorrectly used the incremental baseline");
        Assert(FamilyBrowserTrackingCommitOptimizationPolicy.ResolveBaselineMode(true, false, 0, 1) == FamilyBrowserPostCommitBaselineMode.FullCapture,
            "a commit-boundary metadata gap incorrectly used the incremental baseline");
        Assert(!FamilyBrowserTrackingCommitOptimizationPolicy.ShouldObserveProjectCatalog(true, false, false),
            "an ordinary instance-only commit unnecessarily requested a family/type catalog scan");
        Assert(FamilyBrowserTrackingCommitOptimizationPolicy.ShouldObserveProjectCatalog(true, true, false),
            "a family/type mutation did not request a project catalog scan");
        Assert(FamilyBrowserTrackingCommitOptimizationPolicy.ShouldObserveProjectCatalog(true, false, true),
            "an uncertain commit incorrectly skipped the project catalog safety scan");
        Assert(FamilyBrowserTrackingCommitOptimizationPolicy.ShouldObserveProjectCatalog(false, false, false),
            "a commit without a tracking-session decision incorrectly skipped the project catalog scan");
        checks.Add("successful small commits promote their maintained state without a full model recapture");
        checks.Add("tracking gaps retain the conservative full-capture fallback");
        checks.Add("project catalog scans run only for relevant or uncertain commits");

        FamilyBrowserStandardPolicyStore.ManagedAvailable = false;
        FamilyBrowserOperationLogEntry family = Operation("op-family", "LoadableFamily", "Door", "Single-Flush", "900 x 2100", string.Empty);
        FamilyBrowserOperationLogEntry system = Operation("op-system", "SystemType", string.Empty, string.Empty, "Domestic Cold Water", "PipingSystemType");
        Assert(FamilyBrowserTrackingPersistenceService.PersistOperationEntries("audit", new[] { family }), "offline family operation was not made durable");
        Assert(FamilyBrowserTrackingPersistenceService.PersistOperationEntries("audit", new[] { system }), "offline system operation was not made durable");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingCount() == 2, "offline operation spool count is not 2");
        checks.Add("offline operations are write-ahead spooled");

        StandardRvtChangeCandidateEntry candidate = new StandardRvtChangeCandidateEntry
        {
            EntryId = "candidate-a",
            RecordedAtUtc = DateTime.UtcNow.ToString("O"),
            CommittedAtUtc = DateTime.UtcNow.ToString("O"),
            CommitState = "Committed",
            CandidateKind = "LoadableFamily",
            FamilyName = "Single-Flush",
            TypeName = "900 x 2100"
        };
        Assert(FamilyBrowserTrackingPersistenceService.PersistStandardCandidateEntries("audit", "source-a", new[] { candidate }), "offline standard candidate was not made durable");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingCount() == 3, "combined spool count is not 3");
        checks.Add("offline standard changes are write-ahead spooled");

        string projectIdentity = Path.Combine(fixture, "central", "Project-A.rvt");
        Directory.CreateDirectory(Path.GetDirectoryName(projectIdentity)!);
        File.WriteAllText(projectIdentity, "project identity fixture");
        FamilyBrowserElementChangeCommit elementCommit = ElementCommit("element-commit-a", projectIdentity, "1001", "Created");
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { elementCommit }), "offline element-change commit was not made durable");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingCount() == 4, "combined spool count is not 4");
        checks.Add("offline element changes are write-ahead spooled");

        FamilyBrowserStandardPolicyStore.ManagedAvailable = true;
        FamilyBrowserTrackingFlushResult flush = FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        Assert(flush.FailedCount == 0, "flush reported a failure");
        Assert(flush.OperationFlushedCount == 2, "two operation records were not flushed");
        Assert(flush.StandardCandidateFlushedCount == 1, "one standard candidate was not flushed");
        Assert(flush.ElementChangeFlushedCount == 1, "one element-change commit was not flushed");

        string mixedScopeProjectIdentity = Path.Combine(fixture, "central", "Project-Scope-Summary.rvt");
        File.WriteAllText(mixedScopeProjectIdentity, "project scope summary fixture");
        FamilyBrowserElementChangeCommit mixedScopeCommit = ElementCommit("element-scope-summary", mixedScopeProjectIdentity, "scope-summary-visible", "Created");
        FamilyBrowserTrackedElementState mixedAuxiliaryState = CreateTrackedState("scope-summary-auxiliary", "after");
        mixedAuxiliaryState.ElementClass = "CableTrayRun";
        mixedAuxiliaryState.CategoryName = "Cable Tray Runs";
        mixedAuxiliaryState.CategoryId = "-2008150";
        mixedScopeCommit.Changes.Add(new FamilyBrowserElementChangeItem
        {
            ChangeKind = "Created",
            ElementId = mixedAuxiliaryState.ElementId,
            UniqueId = mixedAuxiliaryState.UniqueId,
            ElementName = mixedAuxiliaryState.ElementName,
            TrackingKind = mixedAuxiliaryState.TrackingKind,
            After = mixedAuxiliaryState
        });
        FamilyBrowserTrackedElementState mixedElectricalSystemState = CreateTrackedState("scope-summary-electrical-system", "after");
        mixedElectricalSystemState.ElementClass = "ElectricalSystem";
        mixedElectricalSystemState.CategoryName = "Electrical Circuits";
        mixedElectricalSystemState.CategoryId = "-2008037";
        mixedScopeCommit.Changes.Add(new FamilyBrowserElementChangeItem
        {
            ChangeKind = "Created",
            ElementId = mixedElectricalSystemState.ElementId,
            UniqueId = mixedElectricalSystemState.UniqueId,
            ElementName = mixedElectricalSystemState.ElementName,
            TrackingKind = mixedElectricalSystemState.TrackingKind,
            After = mixedElectricalSystemState
        });
        mixedScopeCommit.Changes.Add(new FamilyBrowserElementChangeItem
        {
            ChangeKind = "CreatedThenDeleted",
            ElementId = "scope-summary-unresolved",
            ElementClass = FamilyBrowserElementHistoryProjectionPolicy.UnresolvedTransientElementClass,
            TrackingKind = FamilyBrowserElementHistoryProjectionPolicy.UnresolvedTransientTrackingKind
        });
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { mixedScopeCommit }),
            "the mixed-scope immutable history fixture was not persisted");
        Assert(mixedScopeCommit.CreatedCount == 4 && mixedScopeCommit.DeletedCount == 1,
            "persistence unexpectedly mutated immutable evidence counts to match the display projection");
        FamilyBrowserTrackedProjectHistorySummary mixedScopeSummary = FamilyBrowserTrackingPersistenceService
            .LoadTrackedProjectHistorySummaries("audit")
            .Single(delegate(FamilyBrowserTrackedProjectHistorySummary summary)
            {
                return string.Equals(summary.ProjectComparableIdentity, mixedScopeCommit.ProjectComparableIdentity, StringComparison.OrdinalIgnoreCase);
            });
        Assert(mixedScopeSummary.CreatedCount == 1 && mixedScopeSummary.ModifiedCount == 0 && mixedScopeSummary.DeletedCount == 0,
            "the all-project selector used raw immutable counts instead of the shared user-facing projection");
        checks.Add("all-project history summaries use the same projection as detail and Excel without mutating immutable evidence");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingCount() == 0, "spool was not cleared after a successful flush");
        checks.Add("reconnected managed folder flushes and clears local spool");

        List<FamilyBrowserOperationLogEntry> operations = FamilyBrowserTrackingPersistenceService.LoadImmutableOperationEntries("audit", 100);
        Assert(operations.Select(delegate(FamilyBrowserOperationLogEntry entry) { return entry.EntryId; }).Distinct(StringComparer.OrdinalIgnoreCase).Count() == 2, "immutable operation history did not retain both records");
        List<StandardRvtChangeCandidateEntry> candidates = FamilyBrowserTrackingPersistenceService.LoadImmutableStandardCandidateEntries("audit", "source-a", 100);
        Assert(candidates.Count(delegate(StandardRvtChangeCandidateEntry entry) { return entry.EntryId == "candidate-a"; }) == 1, "immutable standard history did not retain the candidate");
        List<FamilyBrowserElementChangeCommit> elementCommits = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommits("audit", projectIdentity, 100);
        Assert(elementCommits.Count(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId == "element-commit-a"; }) == 1, "immutable element history did not retain the commit");
        checks.Add("immutable history is readable and complete");
        checks.Add("element changes flush into per-project immutable history");

        string deferredManagedRoot = Path.Combine(fixture, "deferred-element-managed");
        string deferredSpoolRoot = Path.Combine(fixture, "deferred-element-spool");
        FamilyBrowserStandardPolicyStore.ManagedRoot = deferredManagedRoot;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = false;
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = deferredSpoolRoot;
        FamilyBrowserElementChangeCommit deferredCommit = ElementCommit("element-deferred-publish", projectIdentity, "1002", "Modified");
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommitsDeferred("audit", new[] { deferredCommit }), "deferred element commit was not made locally durable");
        string deferredEnvelopePath = Directory.EnumerateFiles(Path.Combine(deferredSpoolRoot, "ElementChanges"), "*.json").Single();
        FamilyBrowserPendingElementChangeRecord deferredEnvelope = JsonSerializer.Deserialize<FamilyBrowserPendingElementChangeRecord>(File.ReadAllText(deferredEnvelopePath))!;
        Assert(deferredEnvelope.DestinationIdentity.StartsWith("MANAGED-DEFERRED-PATH:", StringComparison.Ordinal), "deferred commit probed or bound the managed destination on the caller thread");
        FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        deferredEnvelope = JsonSerializer.Deserialize<FamilyBrowserPendingElementChangeRecord>(File.ReadAllText(deferredEnvelopePath))!;
        Assert(deferredEnvelope.DestinationIdentity.StartsWith("MANAGED-PATH:", StringComparison.Ordinal), "a later flush did not canonicalize the deferred destination after caller-thread durability");
        FamilyBrowserStandardPolicyStore.ManagedAvailable = true;
        FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingCount() == 0, "deferred element spool was not cleared after managed storage reconnected");
        List<FamilyBrowserElementChangeCommit> deferredHistory = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommits("audit", projectIdentity, 100);
        Assert(deferredHistory.Count(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId == "element-deferred-publish"; }) == 1, "deferred element commit did not reach immutable managed history exactly once");
        checks.Add("Save and Sync element history is locally durable before managed-path canonicalization");
        checks.Add("deferred element history survives a later retry and publishes exactly once after reconnect");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = spool;

        string coverageGapRoot = Path.Combine(fixture, "coverage-gap-managed");
        FamilyBrowserStandardPolicyStore.ManagedRoot = coverageGapRoot;
        FamilyBrowserElementChangeCommit coverageGapCommit = ElementCommit("element-coverage-gap", projectIdentity, "unused", "Modified");
        coverageGapCommit.Changes.Clear();
        coverageGapCommit.CoverageGapOnly = true;
        coverageGapCommit.BaselineCapturedLate = true;
        coverageGapCommit.EventReadFailureCount = 1;
        coverageGapCommit.AttributionConfidence = "ClientObservedWithEventReadGap";
        coverageGapCommit.CoverageNote = "DocumentChanged ID or operation-metadata was unavailable; no element identity was invented.";
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { coverageGapCommit }), "coverage-gap-only evidence was not persisted");
        FamilyBrowserElementChangeHistoryLoadResult coverageGapLoad = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 100);
        FamilyBrowserElementChangeCommit retainedCoverageGap = coverageGapLoad.Commits.Single(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId == "element-coverage-gap"; });
        Assert(retainedCoverageGap.CoverageGapOnly && retainedCoverageGap.Changes.Count == 0 && retainedCoverageGap.EventReadFailureCount == 1 && retainedCoverageGap.IntegrityVersion == 5,
            "coverage-gap-only history lost its explicit gap evidence or integrity-v5 protection");
        FamilyBrowserElementChangeCommit externalRebaseGapCommit = ElementCommit("element-external-rebase-gap", projectIdentity, "unused", "Modified");
        externalRebaseGapCommit.Changes.Clear();
        externalRebaseGapCommit.ModifiedCount = 0;
        externalRebaseGapCommit.CoverageGapOnly = true;
        externalRebaseGapCommit.AttributionConfidence = "ClientObservedWithExternalRebaseGap";
        externalRebaseGapCommit.CoverageNote = "An incoming central/reload update could not be fully rebased. No element identity was invented.";
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { externalRebaseGapCommit }), "external-update rebase coverage gap was not persisted");
        coverageGapLoad = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 100);
        FamilyBrowserElementChangeCommit retainedExternalRebaseGap = coverageGapLoad.Commits.Single(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId == "element-external-rebase-gap"; });
        Assert(retainedExternalRebaseGap.CoverageGapOnly && retainedExternalRebaseGap.Changes.Count == 0 && retainedExternalRebaseGap.IntegrityVersion == 5,
            "external-update rebase coverage gap lost its zero-row evidence or integrity protection");
        FamilyBrowserElementChangeCommit invalidEmptyCommit = ElementCommit("element-invalid-empty", projectIdentity, "unused", "Modified");
        invalidEmptyCommit.Changes.Clear();
        Assert(!FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { invalidEmptyCommit }), "an empty non-coverage commit was silently accepted");
        FamilyBrowserElementChangeCommit invalidUnsupportedGap = ElementCommit("element-invalid-gap", projectIdentity, "unused", "Modified");
        invalidUnsupportedGap.Changes.Clear();
        invalidUnsupportedGap.CoverageGapOnly = true;
        Assert(!FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { invalidUnsupportedGap }), "a coverage-only commit without gap evidence was silently accepted");
        string coverageGapPath = Directory.EnumerateFiles(Path.Combine(coverageGapRoot, "ElementChangeHistory"), "element-coverage-gap.json", SearchOption.AllDirectories).Single();
        string coverageGapJson = File.ReadAllText(coverageGapPath);
        string tamperedCoverageGapJson = coverageGapJson.Replace("\"EventReadFailureCount\":1", "\"EventReadFailureCount\":2", StringComparison.Ordinal);
        Assert(!string.Equals(coverageGapJson, tamperedCoverageGapJson, StringComparison.Ordinal), "coverage-gap tamper fixture did not change the protected field");
        File.WriteAllText(coverageGapPath, tamperedCoverageGapJson);
        coverageGapLoad = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 100);
        Assert(coverageGapLoad.InvalidRecordCount == 1 && coverageGapLoad.Commits.All(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId != "element-coverage-gap"; }),
            "tampering with coverage-gap metadata passed integrity-v5 validation");
        checks.Add("unidentified DocumentChanged coverage gaps persist without fabricated element IDs and are integrity protected");
        checks.Add("external-update rebase failures persist as zero-row coverage gaps without fabricated element IDs");
        checks.Add("empty non-coverage commits fail closed instead of deleting or hiding evidence");

        FamilyBrowserElementChangeCommit inconsistentCountCommit = ElementCommit("element-inconsistent-count", projectIdentity, "1007", "Created");
        inconsistentCountCommit.CreatedCount = 2;
        inconsistentCountCommit.IntegrityVersion = 5;
        inconsistentCountCommit.IntegritySha256 = (string)elementIntegrityMethod.Invoke(null, new object[] { inconsistentCountCommit })!;
        Assert(!FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { inconsistentCountCommit }), "a fully checksummed schema-v6 record with contradictory counters was accepted");
        FamilyBrowserElementChangeCommit inconsistentStateCommit = ElementCommit("element-inconsistent-state", projectIdentity, "1008", "Modified");
        inconsistentStateCommit.Changes[0].After!.ElementId = "9998";
        inconsistentStateCommit.ModifiedCount = 1;
        inconsistentStateCommit.IntegrityVersion = 5;
        inconsistentStateCommit.IntegritySha256 = (string)elementIntegrityMethod.Invoke(null, new object[] { inconsistentStateCommit })!;
        Assert(!FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { inconsistentStateCommit }), "a fully checksummed schema-v6 row whose state belongs to another element was accepted");
        FamilyBrowserElementChangeCommit inconsistentTimeCommit = ElementCommit("element-inconsistent-time", projectIdentity, "1009", "Deleted");
        inconsistentTimeCommit.DeletedCount = 1;
        inconsistentTimeCommit.LocalSaveProtectedAtUtc = DateTime.UtcNow.AddMinutes(5).ToString("O");
        inconsistentTimeCommit.PublishedAtUtc = DateTime.UtcNow.ToString("O");
        inconsistentTimeCommit.IntegrityVersion = 5;
        inconsistentTimeCommit.IntegritySha256 = (string)elementIntegrityMethod.Invoke(null, new object[] { inconsistentTimeCommit })!;
        Assert(!FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { inconsistentTimeCommit }), "a fully checksummed schema-v6 record published before its protection time was accepted");
        FamilyBrowserElementChangeCommit inconsistentProjectCommit = ElementCommit("element-inconsistent-project", projectIdentity, "1011", "Modified");
        inconsistentProjectCommit.ProjectComparableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(Path.Combine(fixture, "central", "Different-Project.rvt"));
        inconsistentProjectCommit.IntegrityVersion = 5;
        inconsistentProjectCommit.IntegritySha256 = (string)elementIntegrityMethod.Invoke(null, new object[] { inconsistentProjectCommit })!;
        Assert(!FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { inconsistentProjectCommit }), "a recomputed schema-v6 record whose project identity contradicts its paths was accepted");
        FamilyBrowserElementChangeCommit observationAfterCommit = ElementCommit("element-observation-after-commit", projectIdentity, "1012", "Modified");
        DateTime observationCommitTime = DateTime.UtcNow;
        observationAfterCommit.TrackingStartedAtUtc = observationCommitTime.AddMinutes(-2).ToString("O");
        observationAfterCommit.BaselineCapturedAtUtc = observationCommitTime.AddMinutes(-1).ToString("O");
        observationAfterCommit.CommittedAtUtc = observationCommitTime.ToString("O");
        observationAfterCommit.Changes[0].FirstObservedAtUtc = observationCommitTime.AddSeconds(-30).ToString("O");
        observationAfterCommit.Changes[0].LastObservedAtUtc = observationCommitTime.AddSeconds(30).ToString("O");
        observationAfterCommit.IntegrityVersion = 5;
        observationAfterCommit.IntegritySha256 = (string)elementIntegrityMethod.Invoke(null, new object[] { observationAfterCommit })!;
        Assert(!FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { observationAfterCommit }), "a recomputed schema-v6 record observed after its commit boundary was accepted");
        checks.Add("schema-v6 checksums cannot legitimize contradictory projects, row states, counters, observations, or protection times");

        string parameterMetadataRoot = Path.Combine(fixture, "parameter-metadata-managed");
        FamilyBrowserStandardPolicyStore.ManagedRoot = parameterMetadataRoot;
        FamilyBrowserElementChangeCommit sharedParameterCommit = ElementCommit("element-shared-parameter", projectIdentity, "1701", "Modified");
        FamilyBrowserElementChangeItem sharedParameterChange = sharedParameterCommit.Changes.Single();
        sharedParameterChange.ElementClass = "SharedParameterElement";
        sharedParameterChange.CategoryName = "Shared Parameter";
        sharedParameterChange.ElementName = "KKY_Fire_Rating";
        sharedParameterChange.FamilyName = string.Empty;
        sharedParameterChange.TypeName = "KKY_Fire_Rating";
        sharedParameterChange.TrackingKind = "SharedParameter";
        sharedParameterChange.Before!.ElementClass = "SharedParameterElement";
        sharedParameterChange.Before.CategoryName = "Shared Parameter";
        sharedParameterChange.Before.ElementName = "KKY_Fire_Rating";
        sharedParameterChange.Before.FamilyName = string.Empty;
        sharedParameterChange.Before.TypeName = "KKY_Fire_Rating";
        sharedParameterChange.Before.TrackingKind = "SharedParameter";
        sharedParameterChange.Before.SharedParameterGuid = "11111111-2222-3333-4444-555555555555";
        sharedParameterChange.Before.ParameterBindingKind = "Instance";
        sharedParameterChange.Before.ParameterBoundCategories = "Walls";
        sharedParameterChange.Before.ParameterBoundCategoryIds = "-2000011";
        sharedParameterChange.Before.ParameterGroup = "PG_DATA";
        sharedParameterChange.Before.ParameterDataType = "Text";
        sharedParameterChange.Before.StateSignature = "shared-before";
        sharedParameterChange.After!.ElementClass = "SharedParameterElement";
        sharedParameterChange.After.CategoryName = "Shared Parameter";
        sharedParameterChange.After.ElementName = "KKY_Fire_Rating";
        sharedParameterChange.After.FamilyName = string.Empty;
        sharedParameterChange.After.TypeName = "KKY_Fire_Rating";
        sharedParameterChange.After.TrackingKind = "SharedParameter";
        sharedParameterChange.After.SharedParameterGuid = "11111111-2222-3333-4444-555555555555";
        sharedParameterChange.After.ParameterBindingKind = "Instance";
        sharedParameterChange.After.ParameterBoundCategories = "Floors, Walls";
        sharedParameterChange.After.ParameterBoundCategoryIds = "-2000032,-2000011";
        sharedParameterChange.After.ParameterGroup = "PG_DATA";
        sharedParameterChange.After.ParameterDataType = "Text";
        sharedParameterChange.After.StateSignature = "shared-after";
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { sharedParameterCommit }), "schema-v6 shared-parameter metadata history was not persisted");
        FamilyBrowserElementChangeHistoryLoadResult sharedParameterLoad = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 10);
        FamilyBrowserElementChangeCommit retainedSharedParameter = sharedParameterLoad.Commits.Single(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId == "element-shared-parameter"; });
        Assert(retainedSharedParameter.SchemaVersion == 6 && retainedSharedParameter.IntegrityVersion == 5 && retainedSharedParameter.Changes[0].After!.ParameterBoundCategories == "Floors, Walls",
            "schema-v6 shared-parameter binding metadata lost its integrity-v5 state");
        string sharedParameterPath = Directory.EnumerateFiles(Path.Combine(parameterMetadataRoot, "ElementChangeHistory"), "element-shared-parameter.json", SearchOption.AllDirectories).Single();
        string sharedParameterJson = File.ReadAllText(sharedParameterPath);
        string tamperedSharedParameterJson = sharedParameterJson.Replace("Floors, Walls", "Roofs, Walls", StringComparison.Ordinal);
        Assert(!string.Equals(sharedParameterJson, tamperedSharedParameterJson, StringComparison.Ordinal), "shared-parameter metadata tamper fixture did not change the protected field");
        File.WriteAllText(sharedParameterPath, tamperedSharedParameterJson);
        sharedParameterLoad = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 10);
        Assert(sharedParameterLoad.InvalidRecordCount == 1 && sharedParameterLoad.Commits.All(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId != "element-shared-parameter"; }),
            "tampering with shared-parameter binding metadata passed integrity-v5 validation");
        checks.Add("schema-v6 integrity-v5 protects shared-parameter GUID, binding, category, group, and data-type metadata");

        string gridMetadataRoot = Path.Combine(fixture, "grid-metadata-managed");
        FamilyBrowserStandardPolicyStore.ManagedRoot = gridMetadataRoot;
        FamilyBrowserElementChangeCommit gridCommit = ElementCommit("element-grid-metadata", projectIdentity, "1702", "Modified");
        FamilyBrowserElementChangeItem gridChange = gridCommit.Changes.Single();
        gridChange.ElementClass = "Grid";
        gridChange.CategoryName = "Grids";
        gridChange.ElementName = "A";
        gridChange.FamilyName = string.Empty;
        gridChange.TypeName = "Grid";
        gridChange.TrackingKind = "Grid";
        gridChange.Before!.ElementClass = "Grid";
        gridChange.Before.CategoryName = "Grids";
        gridChange.Before.ElementName = "A";
        gridChange.Before.FamilyName = string.Empty;
        gridChange.Before.TypeName = "Grid";
        gridChange.Before.TrackingKind = "Grid";
        gridChange.Before.GridCurveSignature = "Line|0,0,0|10000,0,0";
        gridChange.Before.GridExtentsSignature = "Min=-1000,-1000,0|Max=1000,1000,3000";
        gridChange.Before.GridPinnedState = "False";
        gridChange.Before.StateSignature = "grid-before";
        gridChange.After!.ElementClass = "Grid";
        gridChange.After.CategoryName = "Grids";
        gridChange.After.ElementName = "A1";
        gridChange.After.FamilyName = string.Empty;
        gridChange.After.TypeName = "Grid";
        gridChange.After.TrackingKind = "Grid";
        gridChange.After.GridCurveSignature = "Line|500,0,0|10500,0,0";
        gridChange.After.GridExtentsSignature = "Min=-500,-1000,0|Max=1500,1000,3000";
        gridChange.After.GridPinnedState = "True";
        gridChange.After.StateSignature = "grid-after";
        gridChange.ElementName = gridChange.After.ElementName;
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { gridCommit }), "schema-v6 grid metadata history was not persisted");
        FamilyBrowserElementChangeHistoryLoadResult gridLoad = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 10);
        FamilyBrowserElementChangeCommit retainedGrid = gridLoad.Commits.Single(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId == "element-grid-metadata"; });
        Assert(retainedGrid.SchemaVersion == 6 && retainedGrid.IntegrityVersion == 5 && retainedGrid.Changes[0].After!.GridPinnedState == "True" && retainedGrid.Changes[0].After.GridCurveSignature.StartsWith("Line|500", StringComparison.Ordinal),
            "schema-v6 grid curve, extent, and pin metadata lost its integrity-v5 state");
        string gridPath = Directory.EnumerateFiles(Path.Combine(gridMetadataRoot, "ElementChangeHistory"), "element-grid-metadata.json", SearchOption.AllDirectories).Single();
        string gridJson = File.ReadAllText(gridPath);
        string tamperedGridJson = gridJson.Replace("Line|500,0,0|10500,0,0", "Arc|500,0,0|10500,0,0", StringComparison.Ordinal);
        Assert(!string.Equals(gridJson, tamperedGridJson, StringComparison.Ordinal), "grid metadata tamper fixture did not change the protected field");
        File.WriteAllText(gridPath, tamperedGridJson);
        gridLoad = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 10);
        Assert(gridLoad.InvalidRecordCount == 1 && gridLoad.Commits.All(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId != "element-grid-metadata"; }),
            "tampering with grid curve metadata passed integrity-v5 validation");
        checks.Add("schema-v6 integrity-v5 protects grid name, curve, extent, and pin-state metadata");

        string schemaV5CompatibilityRoot = Path.Combine(fixture, "schema-v5-compatibility-managed");
        FamilyBrowserStandardPolicyStore.ManagedRoot = schemaV5CompatibilityRoot;
        FamilyBrowserElementChangeCommit schemaV5Commit = ElementCommit("element-schema-v5-compatibility", projectIdentity, "1703", "Modified");
        schemaV5Commit.SchemaVersion = 5;
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { schemaV5Commit }), "schema-v5 compatibility record was rejected after schema-v6 introduction");
        FamilyBrowserElementChangeCommit retainedSchemaV5 = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommits("audit", projectIdentity, 10)
            .Single(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId == "element-schema-v5-compatibility"; });
        Assert(retainedSchemaV5.IntegrityVersion == 4, "schema-v5 history did not retain its frozen integrity-v4 verifier");
        checks.Add("schema-v5 integrity-v4 history remains readable after schema-v6 metadata tracking");

        string unboundSpool = Path.Combine(fixture, "unbound-destination-spool");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = unboundSpool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = string.Empty;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = false;
        FamilyBrowserOperationLogEntry unboundOperation = Operation("operation-unbound-destination", "LoadableFamily", "Door", "Unbound Door", "900 x 2100", string.Empty);
        StandardRvtChangeCandidateEntry unboundCandidate = new StandardRvtChangeCandidateEntry
        {
            EntryId = "candidate-unbound-destination",
            RecordedAtUtc = DateTime.UtcNow.ToString("O"),
            CommittedAtUtc = DateTime.UtcNow.ToString("O"),
            CommitState = "Committed",
            CandidateKind = "LoadableFamily",
            FamilyName = "Unbound Door",
            TypeName = "900 x 2100"
        };
        Assert(!FamilyBrowserTrackingPersistenceService.PersistOperationEntries("audit", new[] { unboundOperation }), "a new operation audit record without a management destination was accepted");
        Assert(!FamilyBrowserTrackingPersistenceService.PersistStandardCandidateEntries("audit", "source-unbound", new[] { unboundCandidate }), "a new standard candidate without a management destination was accepted");
        FamilyBrowserElementChangeCommit unboundCommit = ElementCommit("element-unbound-destination", projectIdentity, "1010", "Modified");
        Assert(!FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { unboundCommit }), "a new pending element record without a management destination was accepted");
        string unboundCheckpointToken;
        Assert(!FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, Path.Combine(fixture, "unbound-local.rvt"), "AuditRevitUser", new[] { unboundCommit }, false, string.Empty, out unboundCheckpointToken), "a new local checkpoint without a management destination was accepted");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingCount() == 0 &&
            (!Directory.Exists(Path.Combine(unboundSpool, "Operations")) || !Directory.EnumerateFiles(Path.Combine(unboundSpool, "Operations"), "*.json").Any()) &&
            (!Directory.Exists(Path.Combine(unboundSpool, "StandardCandidates")) || !Directory.EnumerateFiles(Path.Combine(unboundSpool, "StandardCandidates"), "*.json").Any()) &&
            (!Directory.Exists(Path.Combine(unboundSpool, "ElementSessions")) || !Directory.EnumerateFiles(Path.Combine(unboundSpool, "ElementSessions"), "*.json").Any()),
            "destination-less tracking evidence entered a trusted local queue");

        string reboundSpool = Path.Combine(fixture, "recomputed-unbound-envelope-spool");
        string reboundManagedA = Path.Combine(fixture, "recomputed-unbound-managed-a");
        string reboundManagedB = Path.Combine(fixture, "recomputed-unbound-managed-b");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = reboundSpool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = reboundManagedA;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = false;
        FamilyBrowserElementChangeCommit reboundCommit = ElementCommit("element-recomputed-unbound", projectIdentity, "1013", "Modified");
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { reboundCommit }), "destination-envelope fixture was not spooled");
        string reboundEnvelopePath = Directory.EnumerateFiles(Path.Combine(reboundSpool, "ElementChanges"), "*.json").Single();
        FamilyBrowserPendingElementChangeRecord reboundEnvelope = JsonSerializer.Deserialize<FamilyBrowserPendingElementChangeRecord>(File.ReadAllText(reboundEnvelopePath))!;
        reboundEnvelope.DestinationIdentity = string.Empty;
        reboundEnvelope.EnvelopeIntegrityVersion = 2;
        reboundEnvelope.EnvelopeIntegritySha256 = (string)pendingEnvelopeIntegrityMethod.Invoke(null, new object[] { reboundEnvelope })!;
        File.WriteAllText(reboundEnvelopePath, JsonSerializer.Serialize(reboundEnvelope));
        FamilyBrowserStandardPolicyStore.ManagedRoot = reboundManagedB;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = true;
        flush = FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        Assert(flush.CorruptRecordCount == 1 && FamilyBrowserTrackingPersistenceService.GetPendingCount() == 1 &&
            !Directory.Exists(Path.Combine(reboundManagedB, "ElementChangeHistory")),
            "a recomputed but destination-less envelope was rebound to an unrelated management folder");
        File.Delete(reboundEnvelopePath);
        checks.Add("new operation, standard candidate, checkpoint, and element records require a non-empty management destination");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = spool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = managed;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = true;

        FamilyBrowserElementChangeCommit legacyIdentityCommit = ElementCommit("element-legacy-identity", projectIdentity, "1002", "Modified");
        legacyIdentityCommit.SchemaVersion = 4;
        legacyIdentityCommit.ProjectComparableIdentity = legacyIdentityCommit.ProjectLegacyComparableIdentity;
        legacyIdentityCommit.ProjectLegacyComparableIdentity = string.Empty;
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { legacyIdentityCommit }), "legacy file-identity commit was not written");
        elementCommits = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommits("audit", projectIdentity, 100);
        Assert(elementCommits.Any(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId == "element-legacy-identity"; }), "stable project history lookup lost a legacy file-identity record");
        checks.Add("stable project history lookup remains compatible with legacy file identities");

        FamilyBrowserElementChangeCommit futureSchemaCommit = ElementCommit("element-future-schema", projectIdentity, "1006", "Modified");
        futureSchemaCommit.SchemaVersion = 99;
        Assert(!FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { futureSchemaCommit }), "an unsupported future element-history schema was persisted as if this client understood it");
        elementCommits = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommits("audit", projectIdentity, 100);
        Assert(!elementCommits.Any(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId == "element-future-schema"; }), "an unsupported future element-history schema appeared in trusted history");
        checks.Add("future element-history schemas fail closed instead of being partially interpreted");

        FamilyBrowserElementChangeCommit replacedFileIdentityCommit = ElementCommit("element-replaced-file-identity", projectIdentity, "1004", "Modified");
        replacedFileIdentityCommit.SchemaVersion = 4;
        replacedFileIdentityCommit.ProjectComparableIdentity = "FILE:OBSOLETE-PRE-REPLACEMENT-IDENTITY";
        replacedFileIdentityCommit.ProjectLegacyComparableIdentity = string.Empty;
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { replacedFileIdentityCommit }), "obsolete file-identity commit was not written");
        elementCommits = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommits("audit", projectIdentity, 100);
        Assert(elementCommits.Any(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId == "element-replaced-file-identity"; }), "a central-file replacement hid history stored under the previous file identity");
        checks.Add("stable path fallback recovers history after the central file identity changes");

        string sameNameProjectA = Path.Combine(fixture, "project-isolation-a", "SharedName.rvt");
        string sameNameProjectB = Path.Combine(fixture, "project-isolation-b", "SharedName.rvt");
        Directory.CreateDirectory(Path.GetDirectoryName(sameNameProjectA)!);
        Directory.CreateDirectory(Path.GetDirectoryName(sameNameProjectB)!);
        File.WriteAllText(sameNameProjectA, "project isolation A");
        File.WriteAllText(sameNameProjectB, "project isolation B");
        FamilyBrowserElementChangeCommit isolatedCommitA = ElementCommit("isolated-project-a", sameNameProjectA, "1011", "Created");
        FamilyBrowserElementChangeCommit isolatedCommitB = ElementCommit("isolated-project-b", sameNameProjectB, "1012", "Deleted");
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { isolatedCommitA, isolatedCommitB }), "same-name project isolation fixtures were not persisted");
        FamilyBrowserElementChangeHistoryLoadResult isolatedLoadA = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", sameNameProjectA, 100);
        FamilyBrowserElementChangeHistoryLoadResult isolatedLoadB = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", sameNameProjectB, 100);
        Assert(isolatedLoadA.Commits.Any(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId == "isolated-project-a"; }) &&
            !isolatedLoadA.Commits.Any(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId == "isolated-project-b"; }), "project A history included a same-name project B record");
        Assert(isolatedLoadB.Commits.Any(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId == "isolated-project-b"; }) &&
            !isolatedLoadB.Commits.Any(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId == "isolated-project-a"; }), "project B history included a same-name project A record");
        checks.Add("same-name project files remain isolated by stable path identity");

        string legacyV1Root = Path.Combine(fixture, "legacy-v1-managed");
        FamilyBrowserStandardPolicyStore.ManagedRoot = legacyV1Root;
        FamilyBrowserElementChangeCommit legacyV1Commit = ElementCommit("element-integrity-v1", projectIdentity, "1003", "Modified");
        legacyV1Commit.SchemaVersion = 1;
        WriteLegacyV1Commit(legacyV1Root, legacyV1Commit);
        FamilyBrowserElementChangeHistoryLoadResult legacyV1Load = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 100);
        Assert(legacyV1Load.Commits.Any(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId == "element-integrity-v1"; }), "a valid integrity-v1 record was rejected after the model schema changed");
        Assert(legacyV1Load.InvalidRecordCount == 0, "a valid integrity-v1 record was reported as corrupt");
        checks.Add("integrity-v1 element history remains verifiable after schema evolution");

        string legacyV1ReplaySpool = Path.Combine(fixture, "legacy-v1-replay-spool");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = legacyV1ReplaySpool;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = false;
        WriteLegacyPendingV1Envelope(legacyV1ReplaySpool, legacyV1Root, legacyV1Commit);
        FamilyBrowserStandardPolicyStore.ManagedAvailable = true;
        flush = FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        Assert(flush.ElementChangeFlushedCount == 1 && flush.FailedCount == 0 && flush.CorruptRecordCount == 0, "an already-persisted integrity-v1 record was not replay-idempotent");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingCount() == 0, "idempotent integrity-v1 replay left a local spool record");
        legacyV1Load = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 100);
        Assert(legacyV1Load.Commits.Count(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId == "element-integrity-v1"; }) == 1, "integrity-v1 replay duplicated immutable history");
        checks.Add("already-persisted integrity-v1 history replays idempotently after schema evolution");

        string legacyPendingV1Root = Path.Combine(fixture, "legacy-pending-v1-managed");
        string legacyPendingV1Spool = Path.Combine(fixture, "legacy-pending-v1-spool");
        FamilyBrowserStandardPolicyStore.ManagedRoot = legacyPendingV1Root;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = false;
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = legacyPendingV1Spool;
        FamilyBrowserElementChangeCommit legacyPendingV1Commit = ElementCommit("element-pending-envelope-v1", projectIdentity, "1005", "Modified");
        legacyPendingV1Commit.SchemaVersion = 2;
        legacyPendingV1Commit.UnmatchedUndoCount = 2;
        legacyPendingV1Commit.UnmatchedRedoCount = 1;
        WriteLegacyPendingV1Envelope(legacyPendingV1Spool, legacyPendingV1Root, legacyPendingV1Commit);
        FamilyBrowserStandardPolicyStore.ManagedAvailable = true;
        flush = FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        Assert(flush.ElementChangeFlushedCount == 1 && flush.CorruptRecordCount == 0, "a valid pending-envelope-v1 record did not flush after schema evolution");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingCount() == 0, "a successfully replayed pending-envelope-v1 record remained in the spool");
        FamilyBrowserElementChangeHistoryLoadResult legacyPendingV1Load = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 100);
        Assert(legacyPendingV1Load.Commits.Any(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId == "element-pending-envelope-v1"; }), "pending-envelope-v1 replay did not reach immutable project history");
        checks.Add("integrity-v1 pending element envelopes survive commit schema evolution");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = spool;

        string checkpointRoot = Path.Combine(fixture, "checkpoint-managed");
        string checkpointLocal = Path.Combine(fixture, "locals", "Project-A_user.rvt");
        Directory.CreateDirectory(Path.GetDirectoryName(checkpointLocal)!);
        File.WriteAllText(checkpointLocal, "local project checkpoint fixture");
        FamilyBrowserStandardPolicyStore.ManagedRoot = checkpointRoot;
        string coverageGapCheckpointLocal = Path.Combine(fixture, "locals", "Project-A_coverage-gap.rvt");
        File.WriteAllText(coverageGapCheckpointLocal, "coverage gap checkpoint fixture");
        string coverageGapCheckpointToken;
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, coverageGapCheckpointLocal, "AuditRevitUser", new[] { coverageGapCommit }, false, string.Empty, out coverageGapCheckpointToken), "coverage-gap-only evidence was not protected by the workshared local-save checkpoint");
        FamilyBrowserElementSessionCheckpointLoadResult coverageGapCheckpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, coverageGapCheckpointLocal, "AuditRevitUser");
        Assert(coverageGapCheckpointLoad.Checkpoint != null && coverageGapCheckpointLoad.Checkpoint.Commits.Count == 1 && coverageGapCheckpointLoad.Checkpoint.Commits[0].CoverageGapOnly,
            "coverage-gap-only evidence was lost while reloading a local-save checkpoint");
        Assert(FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(projectIdentity, coverageGapCheckpointLocal, "AuditRevitUser", coverageGapCheckpointToken), "coverage-gap checkpoint fixture could not be cleaned up");
        checks.Add("coverage-gap-only evidence survives workshared local-save checkpoint recovery");
        FamilyBrowserElementChangeCommit localSaveCommit = ElementCommit("local-save-checkpoint", projectIdentity, "1101", "Modified");
        localSaveCommit.CommitKind = "WorksharedLocalSavePendingSync";
        string checkpointToken;
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { localSaveCommit }, false, string.Empty, out checkpointToken), "workshared local-save checkpoint was not protected");
        FamilyBrowserElementSessionCheckpointLoadResult checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Checkpoint != null && !checkpointLoad.Checkpoint.SynchronizationSucceeded, "unfinalized workshared checkpoint was not restored");
        Assert(checkpointLoad.Checkpoint.Commits.Count == 1 && checkpointLoad.Checkpoint.Commits[0].EntryId == "local-save-checkpoint", "restored workshared checkpoint lost its element commit");
        FamilyBrowserElementSessionCheckpointCountResult checkpointStatus = FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointStatus("audit", projectIdentity);
        Assert(checkpointStatus.Count == 1 && checkpointStatus.SynchronizationSucceededCount == 0, "unfinalized workshared checkpoint status was not project scoped");
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", checkpointLoad.Checkpoint.Commits, true, FamilyBrowserTrackingPersistenceService.GetElementSessionCheckpointRevisionToken(checkpointLoad.Checkpoint), out checkpointToken), "successful synchronization did not finalize the local checkpoint");
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Checkpoint != null && checkpointLoad.Checkpoint.SynchronizationSucceeded, "successful synchronization marker was not durable");
        checkpointStatus = FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointStatus("audit", projectIdentity);
        Assert(checkpointStatus.Count == 1 && checkpointStatus.SynchronizationSucceededCount == 1, "a synchronized checkpoint awaiting history promotion was mislabeled as waiting for synchronization");
        flush = FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        Assert(flush.FinalizedElementSessionPromotedCount == 1 && flush.FailedCount == 0, "ordinary pending flush did not promote a finalized synchronized checkpoint");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointCount("audit", projectIdentity) == 0, "finalized checkpoint remained pending after cleanup");
        FamilyBrowserElementChangeHistoryLoadResult finalizedHistory = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 100);
        Assert(finalizedHistory.Commits.Any(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId == "local-save-checkpoint"; }), "finalized checkpoint was deleted before immutable history could read it");
        flush = FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        Assert(flush.FinalizedElementSessionPromotedCount == 0 && flush.FailedCount == 0, "finalized checkpoint replay was not idempotent after cleanup");
        checks.Add("workshared local saves survive restart and publish only after successful synchronization");
        checks.Add("checkpoint status distinguishes synchronization pending from synchronized history-promotion pending");
        checks.Add("ordinary refresh promotes finalized checkpoints without another synchronization");

        string checkpointCasSpool = Path.Combine(fixture, "checkpoint-cas-spool");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = checkpointCasSpool;
        FamilyBrowserElementChangeCommit casBase = ElementCommit("checkpoint-cas-base", projectIdentity, "1121", "Modified");
        FamilyBrowserElementChangeCommit casWinner = ElementCommit("checkpoint-cas-winner", projectIdentity, "1122", "Created");
        FamilyBrowserElementChangeCommit casLoser = ElementCommit("checkpoint-cas-loser", projectIdentity, "1123", "Deleted");
        string casOriginalToken;
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { casBase }, false, string.Empty, out casOriginalToken), "checkpoint compare-and-swap fixture was not written");
        FamilyBrowserElementSessionCheckpointLoadResult casActorA = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        FamilyBrowserElementSessionCheckpointLoadResult casActorB = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        string casActorARevision = FamilyBrowserTrackingPersistenceService.GetElementSessionCheckpointRevisionToken(casActorA.Checkpoint);
        string casActorBRevision = FamilyBrowserTrackingPersistenceService.GetElementSessionCheckpointRevisionToken(casActorB.Checkpoint);
        Assert(casActorA.Checkpoint != null && casActorB.Checkpoint != null && !string.IsNullOrWhiteSpace(casActorARevision) && string.Equals(casActorARevision, casActorBRevision, StringComparison.OrdinalIgnoreCase), "two simulated Revit processes did not observe the same checkpoint revision");
        string casWinnerToken;
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { casBase, casWinner }, false, casActorARevision, out casWinnerToken), "the first checkpoint writer could not advance the expected revision");
        string casRejectedToken;
        Assert(!FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { casBase, casLoser }, false, casActorBRevision, out casRejectedToken), "a stale checkpoint writer overwrote a newer revision");
        Assert(!FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(projectIdentity, checkpointLocal, "AuditRevitUser", casOriginalToken), "a stale cleanup token deleted a newer checkpoint revision");
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Checkpoint != null && checkpointLoad.Checkpoint.Commits.Any(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId == "checkpoint-cas-winner"; }) &&
            !checkpointLoad.Checkpoint.Commits.Any(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId == "checkpoint-cas-loser"; }),
            "checkpoint compare-and-swap did not preserve the winning revision exactly");
        Assert(FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(projectIdentity, checkpointLocal, "AuditRevitUser", casWinnerToken), "the current checkpoint revision could not be cleaned up with its exact token");
        checks.Add("stale checkpoint writers and cleanup operations fail closed across Revit processes");
        string checkpointLockPath = Path.Combine(checkpointCasSpool, "ElementSessions", ".checkpoint-write.lock");
        using (FileStream checkpointBlocker = new FileStream(checkpointLockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
        {
            FamilyBrowserElementSessionCheckpointLoadResult blockedLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
            Assert(blockedLoad.LockUnavailable && blockedLoad.Checkpoint == null && !blockedLoad.Invalid, "checkpoint lock contention was silently reported as an empty trusted state");
            FamilyBrowserElementSessionCheckpointCountResult blockedCount = FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointStatus("audit", projectIdentity);
            Assert(blockedCount.LockUnavailable && blockedCount.Count == 0, "checkpoint count lock contention was silently reported as a trusted zero-pending state");
        }
        checks.Add("checkpoint load and count lock contention are explicit and never treated as trusted zero pending work");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = spool;

        string destinationCheckpointSpool = Path.Combine(fixture, "destination-checkpoint-spool");
        string destinationCheckpointManagedA = Path.Combine(fixture, "destination-checkpoint-managed-a");
        string destinationCheckpointManagedB = Path.Combine(fixture, "destination-checkpoint-managed-b");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = destinationCheckpointSpool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = destinationCheckpointManagedA;
        FamilyBrowserElementChangeCommit destinationCheckpointCommit = ElementCommit("destination-bound-checkpoint", projectIdentity, "1103", "Modified");
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { destinationCheckpointCommit }, false, string.Empty, out checkpointToken), "destination-bound checkpoint fixture was not written");
        string destinationCheckpointPath = Directory.EnumerateFiles(Path.Combine(destinationCheckpointSpool, "ElementSessions"), "*.json", SearchOption.TopDirectoryOnly).Single();
        string destinationCheckpointOriginal = File.ReadAllText(destinationCheckpointPath);
        string destinationCheckpointPolicyA = Path.Combine(destinationCheckpointManagedA, "Config", "standard-policy.json");
        string destinationCheckpointPolicyB = Path.Combine(destinationCheckpointManagedB, "Config", "standard-policy.json");
        Assert(!FamilyBrowserTrackingPersistenceService.HasBlockingElementSessionCheckpointForManagedPolicyPath(destinationCheckpointPolicyA), "a checkpoint already bound to the active management folder blocked an unchanged path");
        Assert(FamilyBrowserTrackingPersistenceService.HasBlockingElementSessionCheckpointForManagedPolicyPath(destinationCheckpointPolicyB), "automatic management-path replacement ignored a protected checkpoint bound to the previous folder");
        string destinationGuardLockPath = Path.Combine(destinationCheckpointSpool, "ElementSessions", ".checkpoint-write.lock");
        using (FileStream destinationGuardBlocker = new FileStream(destinationGuardLockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
        {
            Assert(FamilyBrowserTrackingPersistenceService.HasBlockingElementSessionCheckpointForManagedPolicyPath(destinationCheckpointPolicyA), "checkpoint lock contention was treated as safe for automatic management-path replacement");
        }
        FamilyBrowserElementChangeCommit invalidCheckpointCommit = ElementCommit("destination-invalid-empty", projectIdentity, "unused", "Modified");
        invalidCheckpointCommit.Changes.Clear();
        string invalidCheckpointToken;
        Assert(!FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { invalidCheckpointCommit }, false, checkpointToken, out invalidCheckpointToken), "an invalid empty checkpoint commit was reported as saved");
        Assert(string.Equals(destinationCheckpointOriginal, File.ReadAllText(destinationCheckpointPath), StringComparison.Ordinal), "an invalid empty checkpoint commit deleted or rewrote existing protected evidence");
        FamilyBrowserStandardPolicyStore.ManagedRoot = string.Empty;
        flush = FamilyBrowserTrackingPersistenceService.FlushPendingForManagedFolderTransition("audit", destinationCheckpointManagedA);
        Assert(flush.ElementSessionCheckpointReboundCount == 0 && flush.ElementSessionCheckpointRebindFailedCount == 1 && flush.FailedCount >= 1,
            "management-folder migration without a verified target destination was reported as successful");
        Assert(string.Equals(destinationCheckpointOriginal, File.ReadAllText(destinationCheckpointPath), StringComparison.Ordinal),
            "management-folder migration with an empty target rewrote protected checkpoint evidence");
        FamilyBrowserStandardPolicyStore.ManagedRoot = destinationCheckpointManagedB;
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.DestinationMismatch && checkpointLoad.Checkpoint == null, "a checkpoint bound to another management folder was loaded as current state");
        Assert(FamilyBrowserTrackingPersistenceService.GetMismatchedElementSessionCheckpointCount("audit") == 1, "a valid checkpoint bound to another management folder was hidden from local recovery status");
        string rejectedDestinationCheckpointToken;
        Assert(!FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { destinationCheckpointCommit }, false, checkpointToken, out rejectedDestinationCheckpointToken), "a local Save silently rebound a checkpoint from another management folder");
        Assert(string.Equals(destinationCheckpointOriginal, File.ReadAllText(destinationCheckpointPath), StringComparison.Ordinal), "rejected destination-bound checkpoint overwrite changed the evidence file");
        flush = FamilyBrowserTrackingPersistenceService.FlushPendingForManagedFolderTransition("audit", destinationCheckpointManagedA);
        Assert(flush.ElementSessionCheckpointReboundCount == 1, "explicit management-folder migration did not rebind a valid local checkpoint");
        Assert(FamilyBrowserTrackingPersistenceService.GetMismatchedElementSessionCheckpointCount("audit") == 0, "an explicitly migrated checkpoint remained marked as a management-folder mismatch");
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Checkpoint != null && !checkpointLoad.Invalid && !checkpointLoad.DestinationMismatch, "a migrated valid checkpoint could not be loaded from the new management folder");
        Assert(string.Equals(checkpointToken, FamilyBrowserTrackingPersistenceService.GetElementSessionCheckpointRevisionToken(checkpointLoad.Checkpoint), StringComparison.OrdinalIgnoreCase), "management-folder migration changed the checkpoint evidence revision");
        Assert(!FamilyBrowserTrackingPersistenceService.HasBlockingElementSessionCheckpointForManagedPolicyPath(destinationCheckpointPolicyB), "an explicitly rebound checkpoint still blocked its verified destination");
        FamilyBrowserElementChangeCommit destinationCheckpointNextCommit = ElementCommit("destination-bound-checkpoint-next", projectIdentity, "1104", "Created");
        string migratedCheckpointToken;
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { destinationCheckpointCommit, destinationCheckpointNextCommit }, false, checkpointToken, out migratedCheckpointToken), "a live Revit session could not continue its checkpoint after explicit management-folder migration");
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Checkpoint != null && checkpointLoad.Checkpoint.Commits.Count == 2, "continued checkpoint after management-folder migration lost evidence");
        Assert(FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(projectIdentity, checkpointLocal, "AuditRevitUser", migratedCheckpointToken), "migrated valid checkpoint could not be cleaned up");
        checks.Add("local checkpoints require explicit management-folder migration and live sessions continue with the same evidence revision");
        checks.Add("management-folder checkpoint mismatches remain visible until explicit migration");
        checks.Add("automatic management-folder replacement is blocked by protected local-save evidence");
        checks.Add("management-folder replacement fails closed while checkpoint state is locked");
        checks.Add("management-folder migration fails closed without a verified target and preserves checkpoint bytes");
        checks.Add("invalid empty checkpoint commits cannot delete existing protected evidence");

        FamilyBrowserStandardPolicyStore.ManagedRoot = destinationCheckpointManagedA;
        FamilyBrowserElementChangeCommit blockedMigrationCommit = ElementCommit("checkpoint-migration-lock", projectIdentity, "1107", "Modified");
        string blockedMigrationToken;
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { blockedMigrationCommit }, false, string.Empty, out blockedMigrationToken), "checkpoint migration lock fixture was not written");
        string blockedMigrationPath = Directory.EnumerateFiles(Path.Combine(destinationCheckpointSpool, "ElementSessions"), "*.json", SearchOption.TopDirectoryOnly).Single();
        string blockedMigrationOriginal = File.ReadAllText(blockedMigrationPath);
        FamilyBrowserStandardPolicyStore.ManagedRoot = destinationCheckpointManagedB;
        string destinationCheckpointLockPath = Path.Combine(destinationCheckpointSpool, "ElementSessions", ".checkpoint-write.lock");
        using (FileStream checkpointMigrationBlocker = new FileStream(destinationCheckpointLockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
        {
            flush = FamilyBrowserTrackingPersistenceService.FlushPendingForManagedFolderTransition("audit", destinationCheckpointManagedA);
            Assert(flush.ElementSessionCheckpointLockUnavailable, "management-folder checkpoint migration did not expose lock contention");
            Assert(flush.ElementSessionCheckpointRebindFailedCount == 1 && flush.FailedCount >= 1, "management-folder checkpoint migration lock contention was reported as successful zero work");
            Assert(flush.ElementSessionCheckpointReboundCount == 0, "locked management-folder checkpoint migration reported a rebound checkpoint");
            Assert(string.Equals(blockedMigrationOriginal, File.ReadAllText(blockedMigrationPath), StringComparison.Ordinal), "locked management-folder checkpoint migration changed protected evidence");
        }
        Assert(FamilyBrowserTrackingPersistenceService.GetMismatchedElementSessionCheckpointCount("audit") == 1, "failed checkpoint migration no longer exposed the source-bound recovery record");
        flush = FamilyBrowserTrackingPersistenceService.FlushPendingForManagedFolderTransition("audit", destinationCheckpointManagedA);
        Assert(flush.ElementSessionCheckpointReboundCount == 1 && flush.ElementSessionCheckpointRebindFailedCount == 0 && flush.FailedCount == 0, "checkpoint migration could not be retried after lock release");
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Checkpoint != null && !checkpointLoad.Invalid && !checkpointLoad.DestinationMismatch, "retried checkpoint migration did not restore a trusted current checkpoint");
        Assert(FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(projectIdentity, checkpointLocal, "AuditRevitUser", blockedMigrationToken), "retried checkpoint migration fixture could not be cleaned up");
        checks.Add("management-folder checkpoint migration lock failures remain explicit, unchanged, and retryable");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = spool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = checkpointRoot;

        string saveAsCheckpointSpool = Path.Combine(fixture, "save-as-checkpoint-spool");
        string saveAsLocalPath = Path.Combine(fixture, "locals", "Project-A_user_SaveAs.rvt");
        File.WriteAllText(saveAsLocalPath, "save-as local checkpoint fixture");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = saveAsCheckpointSpool;
        FamilyBrowserElementChangeCommit saveAsBeforeCommit = ElementCommit("save-as-before", projectIdentity, "1105", "Modified");
        FamilyBrowserElementChangeCommit saveAsAfterCommit = ElementCommit("save-as-after", projectIdentity, "1106", "Created");
        string oldPathCheckpointToken;
        string newPathCheckpointToken;
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { saveAsBeforeCommit }, false, string.Empty, out oldPathCheckpointToken), "pre-Save-As local checkpoint was not written");
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, saveAsLocalPath, "AuditRevitUser", new[] { saveAsBeforeCommit, saveAsAfterCommit }, false, string.Empty, out newPathCheckpointToken), "post-Save-As local checkpoint was not written independently");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointCount("audit", projectIdentity) == 2, "Save As checkpoint identities collapsed into one local record");
        Assert(FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(projectIdentity, checkpointLocal, "AuditRevitUser", oldPathCheckpointToken), "pre-Save-As checkpoint could not be conditionally cleaned up");
        FamilyBrowserElementSessionCheckpointLoadResult saveAsCheckpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, saveAsLocalPath, "AuditRevitUser");
        Assert(saveAsCheckpointLoad.Checkpoint != null && saveAsCheckpointLoad.Checkpoint.Commits.Count == 2, "pre-Save-As cleanup deleted or damaged the new-path checkpoint");
        Assert(FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(projectIdentity, saveAsLocalPath, "AuditRevitUser", newPathCheckpointToken), "post-Save-As checkpoint could not be cleaned up");
        checks.Add("Save As keeps old and new local checkpoint identities isolated during conditional cleanup");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = spool;

        string previousCentralIdentity = Path.Combine(fixture, "central-before-save-as", "Project-Central.rvt");
        string replacementCentralIdentity = Path.Combine(fixture, "central-after-save-as", "Project-Central.rvt");
        Directory.CreateDirectory(Path.GetDirectoryName(previousCentralIdentity)!);
        Directory.CreateDirectory(Path.GetDirectoryName(replacementCentralIdentity)!);
        File.WriteAllText(previousCentralIdentity, "previous central identity");
        File.WriteAllText(replacementCentralIdentity, "replacement central identity");
        string centralChangeLocalPath = Path.Combine(fixture, "locals", "central-change-user.rvt");
        FamilyBrowserElementChangeCommit previousCentralCommit = ElementCommit("checkpoint-previous-central", previousCentralIdentity, "1201", "Modified");
        FamilyBrowserElementChangeCommit replacementCentralCommit = ElementCommit("checkpoint-replacement-central", replacementCentralIdentity, "1202", "Modified");
        string previousCentralToken;
        string replacementCentralToken;
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", previousCentralIdentity, centralChangeLocalPath, "AuditRevitUser", new[] { previousCentralCommit }, false, string.Empty, out previousCentralToken), "previous-central checkpoint fixture was not protected");
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", replacementCentralIdentity, centralChangeLocalPath, "AuditRevitUser", new[] { replacementCentralCommit }, false, string.Empty, out replacementCentralToken), "replacement-central checkpoint collided with the previous project identity");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointCount("audit", previousCentralIdentity) == 1 && FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointCount("audit", replacementCentralIdentity) == 1, "central-identity change collapsed two project checkpoints into one");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointCount("audit") == 2, "global local-save checkpoint status hid a pending project identity");
        Assert(FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(replacementCentralIdentity, centralChangeLocalPath, "AuditRevitUser", replacementCentralToken), "replacement-central checkpoint could not be cleaned independently");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointCount("audit") == 1, "cleaning the replacement central checkpoint also hid or removed the previous central checkpoint");
        FamilyBrowserElementSessionCheckpointLoadResult retainedPreviousCentral = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", previousCentralIdentity, centralChangeLocalPath, "AuditRevitUser");
        Assert(retainedPreviousCentral.Checkpoint != null && string.Equals(retainedPreviousCentral.Checkpoint.Commits.Single().EntryId, previousCentralCommit.EntryId, StringComparison.OrdinalIgnoreCase), "new-project cleanup deleted the previous central project's unsynchronized evidence");
        Assert(FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(previousCentralIdentity, centralChangeLocalPath, "AuditRevitUser", previousCentralToken), "previous-central checkpoint could not be cleaned with its own identity and revision");
        checks.Add("central identity changes retain the previous project's unsynchronized checkpoint independently");
        checks.Add("global local-save status includes protected checkpoints from every project identity");

        string otherCheckpointProject = Path.Combine(fixture, "central", "Project-Other.rvt");
        FamilyBrowserElementChangeCommit crossProjectCheckpointCommit = ElementCommit("cross-project-checkpoint", otherCheckpointProject, "1102", "Modified");
        Assert(!FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { crossProjectCheckpointCommit }, false, string.Empty, out checkpointToken), "a checkpoint accepted an element commit from another project");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointCount("audit", projectIdentity) == 0, "rejected cross-project checkpoint created local state");
        checks.Add("local checkpoints reject commits bound to another project");

        FamilyBrowserElementChangeCommit missingIdCommitA = ElementCommit(string.Empty, projectIdentity, "1111", "Created");
        FamilyBrowserElementChangeCommit missingIdCommitB = ElementCommit(string.Empty, projectIdentity, "1112", "Modified");
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { missingIdCommitA, missingIdCommitB }, false, string.Empty, out checkpointToken), "checkpoint commits without entry IDs were not assigned independent identities");
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Checkpoint != null && checkpointLoad.Checkpoint.Commits.Count == 2, "checkpoint entry-ID assignment collapsed two independent commits");
        Assert(checkpointLoad.Checkpoint.Commits.All(delegate(FamilyBrowserElementChangeCommit commit) { return !string.IsNullOrWhiteSpace(commit.EntryId); }) &&
            checkpointLoad.Checkpoint.Commits.Select(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId; }).Distinct(StringComparer.OrdinalIgnoreCase).Count() == 2,
            "checkpoint entry-ID assignment did not produce two unique durable identities");
        Assert(FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(projectIdentity, checkpointLocal, "AuditRevitUser", FamilyBrowserTrackingPersistenceService.GetElementSessionCheckpointRevisionToken(checkpointLoad.Checkpoint)), "entry-ID assignment checkpoint could not be cleaned up");
        checks.Add("checkpoint commits receive identities before deduplication");

        FamilyBrowserElementChangeCommit duplicateIdCommitA = ElementCommit("checkpoint-collision", projectIdentity, "1113", "Created");
        FamilyBrowserElementChangeCommit duplicateIdCommitB = ElementCommit("checkpoint-collision", projectIdentity, "1114", "Deleted");
        Assert(!FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { duplicateIdCommitA, duplicateIdCommitB }, false, string.Empty, out checkpointToken), "conflicting checkpoint commits with the same entry ID were silently collapsed");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointCount("audit", projectIdentity) == 0, "rejected checkpoint entry-ID collision created trusted pending state");
        checks.Add("conflicting checkpoint entry-ID collisions fail closed");

        string duplicateInnerCheckpointSpool = Path.Combine(fixture, "duplicate-inner-checkpoint-spool");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = duplicateInnerCheckpointSpool;
        FamilyBrowserElementChangeCommit duplicateInnerA = ElementCommit("checkpoint-inner-a", projectIdentity, "1141", "Created");
        FamilyBrowserElementChangeCommit duplicateInnerB = ElementCommit("checkpoint-inner-b", projectIdentity, "1142", "Modified");
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { duplicateInnerA, duplicateInnerB }, false, string.Empty, out checkpointToken), "duplicate-inner checkpoint fixture was not written");
        string duplicateInnerCheckpointPath = Directory.EnumerateFiles(Path.Combine(duplicateInnerCheckpointSpool, "ElementSessions"), "*.json", SearchOption.TopDirectoryOnly).Single();
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Checkpoint != null && checkpointLoad.Checkpoint.Commits.Count == 2, "duplicate-inner checkpoint fixture could not be loaded before tampering");
        checkpointLoad.Checkpoint.Commits[1].EntryId = checkpointLoad.Checkpoint.Commits[0].EntryId;
        checkpointLoad.Checkpoint.Commits[1].IntegritySha256 = (string)elementIntegrityMethod.Invoke(null, new object[] { checkpointLoad.Checkpoint.Commits[1] })!;
        System.Reflection.MethodInfo duplicateCheckpointIntegrityMethod = typeof(FamilyBrowserTrackingPersistenceService).GetMethod("ComputeElementSessionCheckpointIntegrity", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        checkpointLoad.Checkpoint.EnvelopeIntegritySha256 = (string)duplicateCheckpointIntegrityMethod.Invoke(null, new object[] { checkpointLoad.Checkpoint })!;
        File.WriteAllText(duplicateInnerCheckpointPath, JsonSerializer.Serialize(checkpointLoad.Checkpoint));
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Invalid && checkpointLoad.Checkpoint == null, "a checkpoint containing duplicate fully signed entry IDs was trusted");
        File.Delete(duplicateInnerCheckpointPath);
        checks.Add("checkpoint recovery rejects duplicate fully signed entry IDs");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = spool;

        string unsignedCheckpointSpool = Path.Combine(fixture, "unsigned-checkpoint-spool");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = unsignedCheckpointSpool;
        FamilyBrowserElementChangeCommit unsignedCheckpointCommit = ElementCommit("unsigned-checkpoint", projectIdentity, "1115", "Modified");
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { unsignedCheckpointCommit }, false, string.Empty, out checkpointToken), "unsigned-checkpoint fixture was not written");
        string unsignedCheckpointPath = Directory.EnumerateFiles(Path.Combine(unsignedCheckpointSpool, "ElementSessions"), "*.json", SearchOption.TopDirectoryOnly).Single();
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Checkpoint != null && checkpointLoad.Checkpoint.Commits.Count == 1, "unsigned-checkpoint fixture could not be loaded before tampering");
        checkpointLoad.Checkpoint.Commits[0].IntegrityVersion = 0;
        checkpointLoad.Checkpoint.Commits[0].IntegritySha256 = string.Empty;
        System.Reflection.MethodInfo checkpointIntegrityMethod = typeof(FamilyBrowserTrackingPersistenceService).GetMethod("ComputeElementSessionCheckpointIntegrity", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        checkpointLoad.Checkpoint.EnvelopeIntegritySha256 = (string)checkpointIntegrityMethod.Invoke(null, new object[] { checkpointLoad.Checkpoint })!;
        File.WriteAllText(unsignedCheckpointPath, JsonSerializer.Serialize(checkpointLoad.Checkpoint));
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Invalid && checkpointLoad.Checkpoint == null, "an unsigned inner checkpoint commit was accepted behind a valid outer envelope");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointCount("audit", projectIdentity) == 0 && FamilyBrowserTrackingPersistenceService.GetInvalidElementSessionCheckpointCount() == 1, "unsigned inner checkpoint evidence was counted as trusted pending work");
        Assert(!FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(projectIdentity, checkpointLocal, "AuditRevitUser", checkpointToken) && File.Exists(unsignedCheckpointPath), "direct cleanup deleted unsigned checkpoint evidence");
        File.Delete(unsignedCheckpointPath);
        checks.Add("checkpoint inner commits must carry valid integrity evidence");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = spool;

        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { localSaveCommit }, false, string.Empty, out checkpointToken), "checkpoint corruption fixture was not written");
        string corruptCheckpointPath = Directory.EnumerateFiles(Path.Combine(spool, "ElementSessions"), "*.json", SearchOption.TopDirectoryOnly).Single();
        File.WriteAllText(corruptCheckpointPath, "{not-valid-json");
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Invalid && checkpointLoad.Checkpoint == null, "a corrupt local checkpoint was silently treated as no pending work");
        Assert(FamilyBrowserTrackingPersistenceService.GetInvalidElementSessionCheckpointCount() == 1, "corrupt local checkpoint health was not visible to the dashboard");
        File.Delete(corruptCheckpointPath);
        Assert(FamilyBrowserTrackingPersistenceService.GetInvalidElementSessionCheckpointCount() == 0, "removed corrupt checkpoint remained in local health status");
        checks.Add("corrupt local checkpoints are surfaced instead of silently discarded");

        string innerCorruptCheckpointSpool = Path.Combine(fixture, "inner-corrupt-checkpoint-spool");
        string innerCorruptManagedA = Path.Combine(fixture, "inner-corrupt-managed-a");
        string innerCorruptManagedB = Path.Combine(fixture, "inner-corrupt-managed-b");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = innerCorruptCheckpointSpool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = innerCorruptManagedA;
        FamilyBrowserElementChangeCommit innerCorruptCommit = ElementCommit("inner-corrupt-checkpoint", projectIdentity, "1151", "Modified");
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { innerCorruptCommit }, false, string.Empty, out checkpointToken), "inner-corrupt checkpoint fixture was not written");
        string innerCorruptCheckpointPath = Directory.EnumerateFiles(Path.Combine(innerCorruptCheckpointSpool, "ElementSessions"), "*.json", SearchOption.TopDirectoryOnly).Single();
        string innerCorruptCheckpointJson = File.ReadAllText(innerCorruptCheckpointPath);
        string innerCorruptTamperedJson = innerCorruptCheckpointJson.Replace("1151", "1199");
        Assert(!string.Equals(innerCorruptCheckpointJson, innerCorruptTamperedJson, StringComparison.Ordinal), "inner checkpoint fixture was not changed");
        File.WriteAllText(innerCorruptCheckpointPath, innerCorruptTamperedJson);
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Invalid && checkpointLoad.Checkpoint == null, "a checkpoint with a corrupt signed inner commit was accepted");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointCount("audit", projectIdentity) == 0 && FamilyBrowserTrackingPersistenceService.GetInvalidElementSessionCheckpointCount() == 1, "a corrupt inner commit was double-counted as both trusted pending work and invalid state");
        Assert(!FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(projectIdentity, checkpointLocal, "AuditRevitUser", checkpointToken) && File.Exists(innerCorruptCheckpointPath), "direct checkpoint cleanup deleted corrupt evidence");
        string beforeRejectedOverwrite = File.ReadAllText(innerCorruptCheckpointPath);
        Assert(!FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { innerCorruptCommit }, false, checkpointToken, out checkpointToken), "a new local Save overwrote a corrupt checkpoint");
        Assert(string.Equals(beforeRejectedOverwrite, File.ReadAllText(innerCorruptCheckpointPath), StringComparison.Ordinal), "rejected checkpoint overwrite changed the corrupt evidence file");
        FamilyBrowserStandardPolicyStore.ManagedRoot = innerCorruptManagedB;
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Invalid && !checkpointLoad.DestinationMismatch, "checkpoint corruption was hidden behind a management-destination mismatch");
        flush = FamilyBrowserTrackingPersistenceService.FlushPendingForManagedFolderTransition("audit", innerCorruptManagedA);
        Assert(flush.ElementSessionCheckpointReboundCount == 0 && File.Exists(innerCorruptCheckpointPath), "managed-folder migration rebound or discarded a checkpoint with a corrupt inner commit");
        FamilyBrowserStandardPolicyStore.ManagedRoot = innerCorruptManagedA;
        Assert(FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpointsForDestination("audit") == 0 && File.Exists(innerCorruptCheckpointPath), "destination cleanup deleted a checkpoint with a corrupt inner commit");
        File.Delete(innerCorruptCheckpointPath);
        checks.Add("corrupt signed checkpoint commits cannot be rebound or deleted as trusted state");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = spool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = checkpointRoot;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = true;

        string enumerationCheckpointSpool = Path.Combine(fixture, "enumeration-checkpoint-spool");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = enumerationCheckpointSpool;
        FamilyBrowserElementChangeCommit enumerationCheckpointCommit = ElementCommit("checkpoint-enumeration-failure", projectIdentity, "1171", "Modified");
        Assert(FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser", new[] { enumerationCheckpointCommit }, false, string.Empty, out checkpointToken), "checkpoint enumeration failure fixture was not protected");
        string enumerationCheckpointFolder = Path.Combine(enumerationCheckpointSpool, "ElementSessions");
        FamilyBrowserTrackingPersistenceService.EnumerationFailureFolderOverrideForAudit = enumerationCheckpointFolder;
        FamilyBrowserElementSessionCheckpointCountResult enumerationCheckpointStatus = FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointStatus("audit", projectIdentity);
        Assert(enumerationCheckpointStatus.LockUnavailable && enumerationCheckpointStatus.Count == 0, "an unreadable checkpoint folder was reported as a trustworthy zero-pending state");
        Assert(FamilyBrowserTrackingPersistenceService.GetInvalidElementSessionCheckpointCount() == 1, "checkpoint enumeration failure was hidden from local health status");
        Assert(FamilyBrowserTrackingPersistenceService.GetMismatchedElementSessionCheckpointCount("audit") == 1, "checkpoint enumeration failure was hidden from management-destination health status");
        Assert(FamilyBrowserTrackingPersistenceService.HasBlockingElementSessionCheckpointForManagedPolicyPath(FamilyBrowserStandardPolicyStore.GetConfiguredManagedPolicyPath()), "management-folder replacement was allowed while checkpoint enumeration was unavailable");
        Assert(FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpointsForDestination("audit") == 0 && Directory.EnumerateFiles(enumerationCheckpointFolder, "*.json", SearchOption.TopDirectoryOnly).Any(), "unreadable checkpoint evidence was reported as deleted");
        FamilyBrowserTrackingPersistenceService.EnumerationFailureFolderOverrideForAudit = string.Empty;
        checkpointLoad = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint("audit", projectIdentity, checkpointLocal, "AuditRevitUser");
        Assert(checkpointLoad.Checkpoint != null && FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(projectIdentity, checkpointLocal, "AuditRevitUser", FamilyBrowserTrackingPersistenceService.GetElementSessionCheckpointRevisionToken(checkpointLoad.Checkpoint)), "checkpoint enumeration fixture could not be recovered after access returned");
        checks.Add("checkpoint enumeration failures remain blocking and cannot masquerade as trusted zero pending work");

        string enumerationHistoryRoot = Path.Combine(fixture, "enumeration-history-managed");
        string enumerationHistorySpool = Path.Combine(fixture, "enumeration-history-spool");
        FamilyBrowserStandardPolicyStore.ManagedRoot = enumerationHistoryRoot;
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = enumerationHistorySpool;
        FamilyBrowserElementChangeCommit enumerationHistoryCommit = ElementCommit("history-enumeration-failure", projectIdentity, "1172", "Modified");
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { enumerationHistoryCommit }), "history enumeration failure fixture was not persisted");
        string enumerationHistoryFolder = Path.Combine(enumerationHistoryRoot, "ElementChangeHistory");
        FamilyBrowserTrackingPersistenceService.EnumerationFailureFolderOverrideForAudit = enumerationHistoryFolder;
        FamilyBrowserElementChangeHistoryLoadResult enumerationHistoryLoad = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 100);
        Assert(enumerationHistoryLoad.InvalidRecordCount >= 1 && enumerationHistoryLoad.Commits.Count == 0, "an unreadable immutable-history root was presented as clean readable history");
        FamilyBrowserTrackingPersistenceService.EnumerationFailureFolderOverrideForAudit = string.Empty;
        enumerationHistoryLoad = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 100);
        Assert(enumerationHistoryLoad.Commits.Any(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId == enumerationHistoryCommit.EntryId; }), "immutable history did not recover after enumeration access returned");
        checks.Add("immutable-history enumeration failures are explicit and recover after storage access returns");

        string historyCollisionRoot = Path.Combine(fixture, "history-collision-managed");
        string historyCollisionSpool = Path.Combine(fixture, "history-collision-spool");
        FamilyBrowserStandardPolicyStore.ManagedRoot = historyCollisionRoot;
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = historyCollisionSpool;
        FamilyBrowserElementChangeCommit historyCollisionA = ElementCommit("history-cross-root-collision", projectIdentity, "1181", "Created");
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { historyCollisionA }), "history collision baseline record was not persisted");
        FamilyBrowserElementChangeCommit historyCollisionB = ElementCommit("history-cross-root-collision", projectIdentity, "1182", "Deleted");
        historyCollisionB.IntegrityVersion = 5;
        historyCollisionB.IntegritySha256 = (string)elementIntegrityMethod.Invoke(null, new object[] { historyCollisionB })!;
        string manualCollisionFolder = Path.Combine(historyCollisionRoot, "ElementChangeHistory", "manual-conflicting-root", "20260719");
        Directory.CreateDirectory(manualCollisionFolder);
        File.WriteAllText(Path.Combine(manualCollisionFolder, "history-cross-root-collision.json"), JsonSerializer.Serialize(historyCollisionB));
        FamilyBrowserElementChangeHistoryLoadResult historyCollisionLoad = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 100);
        Assert(historyCollisionLoad.InvalidRecordCount >= 2 && !historyCollisionLoad.Commits.Any(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId == "history-cross-root-collision"; }), "conflicting immutable records with one entry ID silently selected an arbitrary winner");
        checks.Add("conflicting immutable entry IDs across legacy and current identity roots are quarantined together");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = spool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = checkpointRoot;

        Assert(FamilyBrowserTrackingPersistenceService.PersistOperationEntries("audit", new[] { family }), "idempotent operation replay failed");
        operations = FamilyBrowserTrackingPersistenceService.LoadImmutableOperationEntries("audit", 100);
        Assert(operations.Count(delegate(FamilyBrowserOperationLogEntry entry) { return entry.EntryId == "op-family"; }) == 1, "replay created a duplicate immutable entry");
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { elementCommit }), "idempotent element-change replay failed");
        elementCommits = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommits("audit", projectIdentity, 100);
        Assert(elementCommits.Count(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId == "element-commit-a"; }) == 1, "element replay created a duplicate immutable entry");
        checks.Add("stable entry IDs make replay idempotent");

        Parallel.For(0, 40, delegate(int index)
        {
            FamilyBrowserTrackingPersistenceService.PersistOperationEntries("audit", new[]
            {
                Operation("parallel-" + index.ToString("00"), "LoadableFamily", "Generic Models", "Parallel " + index, "Type " + index, string.Empty)
            });
        });
        operations = FamilyBrowserTrackingPersistenceService.LoadImmutableOperationEntries("audit", 1000);
        Assert(operations.Count(delegate(FamilyBrowserOperationLogEntry entry) { return entry.EntryId.StartsWith("parallel-", StringComparison.Ordinal); }) == 40, "parallel immutable writes lost records");
        checks.Add("parallel writers retain every uniquely identified operation");

        Parallel.For(0, 20, delegate(int index)
        {
            FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[]
            {
                ElementCommit("parallel-element-" + index.ToString("00"), projectIdentity, (2000 + index).ToString(), "Modified")
            });
        });
        elementCommits = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommits("audit", projectIdentity, 1000);
        Assert(elementCommits.Count(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId.StartsWith("parallel-element-", StringComparison.Ordinal); }) == 20, "parallel element-change writes lost commits");
        checks.Add("parallel element-change commits remain complete");

        string destinationSpool = Path.Combine(fixture, "destination-bound-spool");
        string managedA = Path.Combine(fixture, "managed-a");
        string managedB = Path.Combine(fixture, "managed-b");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = destinationSpool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = managedA;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = false;
        FamilyBrowserElementChangeCommit destinationBound = ElementCommit("destination-bound", projectIdentity, "3001", "Created");
        FamilyBrowserElementChangeCommit otherProjectPending = ElementCommit("destination-other-project", Path.Combine(fixture, "central", "Project-B.rvt"), "3002", "Created");
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { destinationBound, otherProjectPending }), "destination-bound commits were not spooled");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingElementChangeCount("audit", projectIdentity) == 1, "project-specific pending count included another project");
        string managedC = Path.Combine(fixture, "managed-c");
        FamilyBrowserStandardPolicyStore.ManagedRoot = managedC;
        FamilyBrowserElementChangeCommit unrelatedDestination = ElementCommit("destination-unrelated", projectIdentity, "3003", "Created");
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { unrelatedDestination }), "unrelated-destination commit was not spooled");
        FamilyBrowserStandardPolicyStore.ManagedRoot = managedB;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = true;
        flush = FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        Assert(flush.DestinationMismatchCount == 3, "pending records were not bound to their original management destination");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingCount() == 3, "destination mismatch deleted protected pending records");
        Assert(FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommits("audit", projectIdentity, 10).Count == 0, "destination mismatch leaked history into another management root");
        flush = FamilyBrowserTrackingPersistenceService.FlushPendingForManagedFolderTransition("audit", string.Empty);
        Assert(flush.FailedCount == 1 && flush.DestinationMismatchCount == 1 && FamilyBrowserTrackingPersistenceService.GetPendingCount() == 3,
            "a management-folder transition without a verified source rebound or deleted protected pending records");
        flush = FamilyBrowserTrackingPersistenceService.FlushPendingForManagedFolderTransition("audit", managedA);
        Assert(flush.ElementChangeFlushedCount == 2 && flush.DestinationMismatchCount == 1 && FamilyBrowserTrackingPersistenceService.GetPendingCount() == 1, "explicit managed-folder transition did not restrict rebind to its source root");
        Assert(FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommits("audit", projectIdentity, 10).Any(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId == "destination-bound"; }), "explicit transition lost the destination-bound project commit");
        Assert(!FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommits("audit", projectIdentity, 10).Any(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId == "destination-unrelated"; }), "explicit transition leaked an unrelated management-root commit");
        FamilyBrowserStandardPolicyStore.ManagedRoot = managedC;
        flush = FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        Assert(flush.ElementChangeFlushedCount == 1 && FamilyBrowserTrackingPersistenceService.GetPendingCount() == 0, "unrelated destination record could not later flush to its own root");
        checks.Add("pending tracking records cannot leak into another management root without explicit migration");
        checks.Add("managed-folder migration rebinds only records from the selected source root");
        checks.Add("managed-folder migration rejects an empty or unverifiable source root");
        checks.Add("pending element counts are scoped to the current project and destination");

        string envelopeSpool = Path.Combine(fixture, "envelope-integrity-spool");
        string envelopeManaged = Path.Combine(fixture, "envelope-managed-a");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = envelopeSpool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = envelopeManaged;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = false;
        FamilyBrowserElementChangeCommit envelopeCommit = ElementCommit("envelope-protected", projectIdentity, "3501", "Modified");
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { envelopeCommit }), "envelope-protected commit was not spooled");
        string envelopePath = Directory.EnumerateFiles(Path.Combine(envelopeSpool, "ElementChanges"), "envelope-protected.json", SearchOption.TopDirectoryOnly).Single();
        string originalEnvelopeJson = File.ReadAllText(envelopePath);
        string envelopeJson = originalEnvelopeJson.Replace("ENVELOPE-MANAGED-A", "ENVELOPE-MANAGED-Z");
        Assert(!string.Equals(originalEnvelopeJson, envelopeJson, StringComparison.Ordinal), "envelope destination fixture was not changed");
        File.WriteAllText(envelopePath, envelopeJson);
        FamilyBrowserStandardPolicyStore.ManagedAvailable = true;
        flush = FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        Assert(flush.CorruptRecordCount == 1 && FamilyBrowserTrackingPersistenceService.GetPendingCount() == 1, "tampered pending destination envelope was accepted or deleted");
        File.Delete(envelopePath);
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingCount() == 0, "tampered envelope fixture was not cleaned up");
        checks.Add("pending element destination metadata is checksum protected");

        string metadataEnvelopeSpool = Path.Combine(fixture, "metadata-envelope-integrity-spool");
        string metadataEnvelopeManaged = Path.Combine(fixture, "envelope-operation-managed-a");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = metadataEnvelopeSpool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = metadataEnvelopeManaged;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = false;
        FamilyBrowserOperationLogEntry envelopeOperation = Operation("operation-envelope-protected", "ProjectElementChangeTrackingPolicy", string.Empty, string.Empty, string.Empty, string.Empty);
        StandardRvtChangeCandidateEntry envelopeCandidate = new StandardRvtChangeCandidateEntry
        {
            EntryId = "candidate-envelope-protected",
            RecordedAtUtc = DateTime.UtcNow.ToString("O"),
            CommittedAtUtc = DateTime.UtcNow.ToString("O"),
            CommitState = "Committed",
            CandidateKind = "LoadableFamily",
            FamilyName = "Envelope Candidate",
            TypeName = "Type A"
        };
        Assert(FamilyBrowserTrackingPersistenceService.PersistOperationEntries("audit", new[] { envelopeOperation }), "operation envelope fixture was not spooled");
        Assert(FamilyBrowserTrackingPersistenceService.PersistStandardCandidateEntries("audit", "source-envelope", new[] { envelopeCandidate }), "candidate envelope fixture was not spooled");
        string operationEnvelopePath = Directory.EnumerateFiles(Path.Combine(metadataEnvelopeSpool, "Operations"), "operation-envelope-protected.json", SearchOption.TopDirectoryOnly).Single();
        string candidateEnvelopePath = Directory.EnumerateFiles(Path.Combine(metadataEnvelopeSpool, "StandardCandidates"), "candidate-envelope-protected.json", SearchOption.TopDirectoryOnly).Single();
        string originalOperationEnvelope = File.ReadAllText(operationEnvelopePath);
        string originalCandidateEnvelope = File.ReadAllText(candidateEnvelopePath);
        string tamperedOperationEnvelope = originalOperationEnvelope.Replace("ENVELOPE-OPERATION-MANAGED-A", "ENVELOPE-OPERATION-MANAGED-Z", StringComparison.Ordinal);
        string tamperedCandidateEnvelope = originalCandidateEnvelope.Replace("ENVELOPE-OPERATION-MANAGED-A", "ENVELOPE-OPERATION-MANAGED-Z", StringComparison.Ordinal);
        Assert(!string.Equals(originalOperationEnvelope, tamperedOperationEnvelope, StringComparison.Ordinal) && !string.Equals(originalCandidateEnvelope, tamperedCandidateEnvelope, StringComparison.Ordinal), "operation/candidate destination fixtures were not changed");
        File.WriteAllText(operationEnvelopePath, tamperedOperationEnvelope);
        File.WriteAllText(candidateEnvelopePath, tamperedCandidateEnvelope);
        FamilyBrowserStandardPolicyStore.ManagedAvailable = true;
        flush = FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        Assert(flush.CorruptRecordCount == 2 && flush.FailedCount == 2 && FamilyBrowserTrackingPersistenceService.GetPendingCount() == 2,
            "tampered operation or standard-candidate destination metadata was accepted or deleted");
        File.Delete(operationEnvelopePath);
        File.Delete(candidateEnvelopePath);
        checks.Add("pending operation and standard-candidate destinations are checksum protected");

        string cleanupSpool = Path.Combine(fixture, "cleanup-confirmation-spool");
        string cleanupManaged = Path.Combine(fixture, "cleanup-confirmation-managed");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = cleanupSpool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = cleanupManaged;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = false;
        FamilyBrowserOperationLogEntry cleanupOperation = Operation("operation-cleanup-locked", "LoadableFamily", "Door", "Cleanup Door", "Type A", string.Empty);
        Assert(FamilyBrowserTrackingPersistenceService.PersistOperationEntries("audit", new[] { cleanupOperation }), "cleanup confirmation fixture was not spooled");
        string cleanupSpoolPath = Directory.EnumerateFiles(Path.Combine(cleanupSpool, "Operations"), "operation-cleanup-locked.json", SearchOption.TopDirectoryOnly).Single();
        FamilyBrowserStandardPolicyStore.ManagedAvailable = true;
        using (FileStream cleanupLease = new FileStream(cleanupSpoolPath, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            flush = FamilyBrowserTrackingPersistenceService.FlushPending("audit");
            Assert(flush.FailedCount == 1 && flush.OperationFlushedCount == 0 && FamilyBrowserTrackingPersistenceService.GetPendingCount() == 1,
                "a published record whose local spool cleanup failed was reported as fully settled");
        }
        flush = FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        Assert(flush.FailedCount == 0 && flush.OperationFlushedCount == 1 && FamilyBrowserTrackingPersistenceService.GetPendingCount() == 0,
            "a cleanup-only retry did not settle the already durable operation record");
        checks.Add("managed-folder transitions can detect and retry local spool cleanup failures");

        string corruptSpool = Path.Combine(fixture, "corrupt-destination-spool");
        string corruptManaged = Path.Combine(fixture, "corrupt-managed");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = corruptSpool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = corruptManaged;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = true;
        FamilyBrowserElementChangeCommit protectedCommit = ElementCommit("protected-commit", projectIdentity, "4001", "Modified");
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { protectedCommit }), "protected commit was not written");
        string protectedPath = Directory.EnumerateFiles(Path.Combine(corruptManaged, "ElementChangeHistory"), "protected-commit.json", SearchOption.AllDirectories).Single();
        string tamperedJson = File.ReadAllText(protectedPath).Replace("4001", "4999");
        File.WriteAllText(protectedPath, tamperedJson);
        FamilyBrowserElementChangeHistoryLoadResult tamperedLoad = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 10);
        Assert(tamperedLoad.InvalidRecordCount == 1 && tamperedLoad.Commits.All(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId != "protected-commit"; }), "tampered element history passed checksum validation");
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { protectedCommit }), "valid replay was not retained locally after a corrupt destination collision");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingCount() == 1, "corrupt destination collision discarded the valid local copy");
        flush = FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        Assert(flush.FailedCount == 1 && FamilyBrowserTrackingPersistenceService.GetPendingCount() == 1, "corrupt destination was incorrectly accepted as an idempotent write");
        File.Delete(protectedPath);
        flush = FamilyBrowserTrackingPersistenceService.FlushPending("audit");
        Assert(flush.ElementChangeFlushedCount == 1 && FamilyBrowserTrackingPersistenceService.GetPendingCount() == 0, "valid local copy did not recover after removing the corrupt destination");
        checks.Add("element history checksum rejects tampered records");
        checks.Add("corrupt destination collisions preserve the valid local write-ahead copy");

        string resignSpool = Path.Combine(fixture, "tampered-resign-spool");
        string resignManaged = Path.Combine(fixture, "tampered-resign-managed");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = resignSpool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = resignManaged;
        FamilyBrowserStandardPolicyStore.ManagedAvailable = true;
        FamilyBrowserElementChangeCommit resignCommit = ElementCommit("tampered-resign", projectIdentity, "4501", "Modified");
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { resignCommit }), "signed tamper fixture was not written");
        resignCommit.Changes[0].ElementId = "4599";
        Assert(!FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { resignCommit }), "a checksum-invalid in-memory commit was silently re-signed");
        Assert(FamilyBrowserTrackingPersistenceService.GetPendingCount() == 0, "rejected checksum-invalid commit entered the local spool");
        FamilyBrowserElementChangeHistoryLoadResult resignLoad = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 20);
        FamilyBrowserElementChangeCommit retainedResignCommit = resignLoad.Commits.Single(delegate(FamilyBrowserElementChangeCommit entry) { return entry.EntryId == "tampered-resign"; });
        Assert(retainedResignCommit.Changes[0].ElementId == "4501", "rejected checksum-invalid commit altered immutable history");
        checks.Add("checksum-invalid signed commits are rejected instead of silently re-signed");

        string orderingSpool = Path.Combine(fixture, "ordering-spool");
        string orderingManaged = Path.Combine(fixture, "ordering-managed");
        FamilyBrowserTrackingPersistenceService.LocalSpoolRootOverrideForAudit = orderingSpool;
        FamilyBrowserStandardPolicyStore.ManagedRoot = orderingManaged;
        FamilyBrowserElementChangeCommit olderCommit = ElementCommit("ordering-old", projectIdentity, "5001", "Modified");
        FamilyBrowserElementChangeCommit newerCommit = ElementCommit("ordering-new", projectIdentity, "5002", "Modified");
        DateTime sameDay = DateTime.UtcNow.Date;
        olderCommit.CommittedAtUtc = sameDay.AddHours(1).ToString("O");
        newerCommit.CommittedAtUtc = sameDay.AddHours(22).ToString("O");
        Assert(FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits("audit", new[] { olderCommit, newerCommit }), "ordering fixture commits were not written");
        string oldPath = Directory.EnumerateFiles(Path.Combine(orderingManaged, "ElementChangeHistory"), "ordering-old.json", SearchOption.AllDirectories).Single();
        string newPath = Directory.EnumerateFiles(Path.Combine(orderingManaged, "ElementChangeHistory"), "ordering-new.json", SearchOption.AllDirectories).Single();
        File.SetLastWriteTimeUtc(oldPath, DateTime.UtcNow.AddMinutes(5));
        File.SetLastWriteTimeUtc(newPath, DateTime.UtcNow.AddMinutes(-5));
        FamilyBrowserElementChangeHistoryLoadResult orderedLoad = FamilyBrowserTrackingPersistenceService.LoadImmutableElementChangeCommitResult("audit", projectIdentity, 1);
        Assert(orderedLoad.Commits.Count == 1 && orderedLoad.Commits[0].EntryId == "ordering-new", "latest history used file copy time instead of committed time");
        checks.Add("recent element history is ordered by committed time rather than file copy time");

        FamilyBrowserOperationLogEntry olderOperation = Operation("operation-ordering-old", "ProjectElementChangeTrackingPolicy", string.Empty, string.Empty, string.Empty, string.Empty);
        FamilyBrowserOperationLogEntry newerOperation = Operation("operation-ordering-new", "ProjectElementChangeTrackingPolicy", string.Empty, string.Empty, string.Empty, string.Empty);
        olderOperation.RecordedAtUtc = olderOperation.CommittedAtUtc = sameDay.AddHours(2).ToString("O");
        newerOperation.RecordedAtUtc = newerOperation.CommittedAtUtc = sameDay.AddHours(21).ToString("O");
        Assert(FamilyBrowserTrackingPersistenceService.PersistOperationEntries("audit", new[] { olderOperation, newerOperation }), "operation ordering fixtures were not written");
        string oldOperationPath = Directory.EnumerateFiles(Path.Combine(orderingManaged, "OperationLogs"), "operation-ordering-old.json", SearchOption.AllDirectories).Single();
        string newOperationPath = Directory.EnumerateFiles(Path.Combine(orderingManaged, "OperationLogs"), "operation-ordering-new.json", SearchOption.AllDirectories).Single();
        File.SetLastWriteTimeUtc(oldOperationPath, DateTime.UtcNow.AddMinutes(5));
        File.SetLastWriteTimeUtc(newOperationPath, DateTime.UtcNow.AddMinutes(-5));
        List<FamilyBrowserOperationLogEntry> orderedOperations = FamilyBrowserTrackingPersistenceService.LoadImmutableOperationEntries("audit", 1);
        Assert(orderedOperations.Count == 1 && orderedOperations[0].EntryId == "operation-ordering-new", "latest operation history used file copy time instead of committed time");

        StandardRvtChangeCandidateEntry olderCandidate = new StandardRvtChangeCandidateEntry
        {
            EntryId = "candidate-ordering-old",
            RecordedAtUtc = sameDay.AddHours(3).ToString("O"),
            CommittedAtUtc = sameDay.AddHours(3).ToString("O"),
            CommitState = "Committed",
            CandidateKind = "LoadableFamily"
        };
        StandardRvtChangeCandidateEntry newerCandidate = new StandardRvtChangeCandidateEntry
        {
            EntryId = "candidate-ordering-new",
            RecordedAtUtc = sameDay.AddHours(20).ToString("O"),
            CommittedAtUtc = sameDay.AddHours(20).ToString("O"),
            CommitState = "Committed",
            CandidateKind = "LoadableFamily"
        };
        Assert(FamilyBrowserTrackingPersistenceService.PersistStandardCandidateEntries("audit", "source-ordering", new[] { olderCandidate, newerCandidate }), "candidate ordering fixtures were not written");
        string oldCandidatePath = Directory.EnumerateFiles(Path.Combine(orderingManaged, "StandardChangeCandidates"), "candidate-ordering-old.json", SearchOption.AllDirectories).Single();
        string newCandidatePath = Directory.EnumerateFiles(Path.Combine(orderingManaged, "StandardChangeCandidates"), "candidate-ordering-new.json", SearchOption.AllDirectories).Single();
        File.SetLastWriteTimeUtc(oldCandidatePath, DateTime.UtcNow.AddMinutes(5));
        File.SetLastWriteTimeUtc(newCandidatePath, DateTime.UtcNow.AddMinutes(-5));
        List<StandardRvtChangeCandidateEntry> orderedCandidates = FamilyBrowserTrackingPersistenceService.LoadImmutableStandardCandidateEntries("audit", "source-ordering", 1);
        Assert(orderedCandidates.Count == 1 && orderedCandidates[0].EntryId == "candidate-ordering-new", "latest standard candidate history used file copy time instead of committed time");
        checks.Add("recent operation and standard-candidate history is ordered by record time after migration or recovery");

		string requestStore = Path.Combine(fixture, "requests");
		Directory.CreateDirectory(requestStore);
		using (FamilyBrowserRequestMutationLease firstLease = FamilyBrowserRequestConcurrencyService.Acquire(requestStore, "FBR-CONCURRENCY", 500))
		{
			bool lockConflict = Task.Run(delegate
			{
				try
				{
					using FamilyBrowserRequestMutationLease ignored = FamilyBrowserRequestConcurrencyService.Acquire(requestStore, "FBR-CONCURRENCY", 250);
					return false;
				}
				catch (FamilyBrowserRequestConflictException ex)
				{
					return ex.LockTimedOut;
				}
			}).GetAwaiter().GetResult();
			Assert(lockConflict, "request-scoped lock allowed a second writer");
		}
		using (FamilyBrowserRequestMutationLease releasedLease = FamilyBrowserRequestConcurrencyService.Acquire(requestStore, "FBR-CONCURRENCY", 500))
		{
			Assert(releasedLease != null, "request-scoped lock could not be reacquired after release");
		}
		checks.Add("request-scoped lock serializes writers and releases cleanly");

		string revisionToken = FamilyBrowserRequestConcurrencyService.CreateRevisionToken();
		FamilyBrowserRequestConcurrencyService.EnsureExpectedRevision("FBR-CONCURRENCY", 1L, revisionToken, 1L, revisionToken);
		bool staleRevisionBlocked = false;
		try
		{
			FamilyBrowserRequestConcurrencyService.EnsureExpectedRevision("FBR-CONCURRENCY", 1L, revisionToken, 2L, FamilyBrowserRequestConcurrencyService.CreateRevisionToken());
		}
		catch (FamilyBrowserRequestConflictException)
		{
			staleRevisionBlocked = true;
		}
		Assert(staleRevisionBlocked, "stale request revision was not blocked");
		checks.Add("stale request revisions are rejected before mutation");

		string legacyRequest = Path.Combine(requestStore, "legacy-request.json");
		File.WriteAllText(legacyRequest, "{\"Status\":\"Submitted\"}");
		string firstLegacyToken = FamilyBrowserRequestConcurrencyService.ComputeFileToken(legacyRequest);
		File.WriteAllText(legacyRequest, "{\"Status\":\"Approved\"}");
		string secondLegacyToken = FamilyBrowserRequestConcurrencyService.ComputeFileToken(legacyRequest);
		Assert(!string.Equals(firstLegacyToken, secondLegacyToken, StringComparison.OrdinalIgnoreCase), "legacy request content change did not alter its compatibility token");
		bool legacyChangeBlocked = false;
		try
		{
			FamilyBrowserRequestConcurrencyService.EnsureExpectedRevision("FBR-LEGACY", 1L, firstLegacyToken, 1L, secondLegacyToken);
		}
		catch (FamilyBrowserRequestConflictException)
		{
			legacyChangeBlocked = true;
		}
		Assert(legacyChangeBlocked, "an external or old-client request edit was not detected");
		checks.Add("legacy and old-client edits are detected by the file-content token");

		string attachmentSource = Path.Combine(fixture, "request-attachment.txt");
		string attachmentFolder = Path.Combine(requestStore, "Attachments", "FBR-TRANSACTION");
		File.WriteAllText(attachmentSource, "attachment revision one");
		FamilyBrowserRequestAttachmentCopyResult firstAttachment = FamilyBrowserRequestFileTransactionService.CopyContentAddressed(attachmentSource, attachmentFolder, "Evidence.txt");
		FamilyBrowserRequestAttachmentCopyResult repeatedAttachment = FamilyBrowserRequestFileTransactionService.CopyContentAddressed(attachmentSource, attachmentFolder, "Evidence.txt");
		Assert(firstAttachment.Created, "first content-addressed request attachment was not created");
		Assert(!repeatedAttachment.Created, "retry created a duplicate request attachment");
		Assert(string.Equals(firstAttachment.StoredPath, repeatedAttachment.StoredPath, StringComparison.OrdinalIgnoreCase), "retry did not reuse the same request attachment path");
		Assert(Directory.GetFiles(attachmentFolder).Count(delegate(string path) { return !Path.GetFileName(path).StartsWith(".kky-", StringComparison.OrdinalIgnoreCase); }) == 1, "retry left more than one attachment file");
		checks.Add("request attachment retries are content-addressed and idempotent");

		File.WriteAllText(attachmentSource, "attachment revision two");
		FamilyBrowserRequestAttachmentCopyResult changedAttachment = FamilyBrowserRequestFileTransactionService.CopyContentAddressed(attachmentSource, attachmentFolder, "Evidence.txt");
		Assert(changedAttachment.Created, "changed request attachment content did not create a new content address");
		Assert(!string.Equals(firstAttachment.StoredPath, changedAttachment.StoredPath, StringComparison.OrdinalIgnoreCase), "different request attachment content reused a stale path");
		FamilyBrowserRequestFileTransactionService.RollbackCreatedFile(changedAttachment.StoredPath);
		Assert(!File.Exists(changedAttachment.StoredPath), "request attachment rollback did not remove the newly created file");
		checks.Add("changed attachments receive a new identity and pre-commit rollback removes new files");

		string deletionAuditPath = Path.Combine(requestStore, "RequestAudit", "Deleted", "audit-delete-prepared.json");
		FamilyBrowserRequestFileTransactionService.WriteImmutableText(deletionAuditPath, "{\"EventType\":\"DeletePrepared\"}");
		bool immutableOverwriteBlocked = false;
		try
		{
			FamilyBrowserRequestFileTransactionService.WriteImmutableText(deletionAuditPath, "{\"EventType\":\"DeleteCompleted\"}");
		}
		catch (IOException)
		{
			immutableOverwriteBlocked = true;
		}
		Assert(immutableOverwriteBlocked, "request deletion audit allowed an existing event to be overwritten");
		Assert(File.ReadAllText(deletionAuditPath).Contains("DeletePrepared"), "request deletion audit content changed after rejected overwrite");
		checks.Add("request deletion audit entries are immutable and preserve the prepared snapshot");

        Directory.CreateDirectory(output);
        string summaryPath = Path.Combine(output, "tracking-persistence-summary.json");
        File.WriteAllText(summaryPath, JsonSerializer.Serialize(new
        {
            generatedAt = DateTime.UtcNow.ToString("O"),
            status = "PASS",
            checks,
            operationCount = operations.Count,
            pendingCount = FamilyBrowserTrackingPersistenceService.GetPendingCount(),
            fixture
        }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine(summaryPath);
        return 0;
    }

    private static FamilyBrowserElementChangeCommit ElementCommit(string entryId, string projectIdentity, string elementId, string changeKind)
    {
        FamilyBrowserTrackedElementState before = CreateTrackedState(elementId, "before");
        FamilyBrowserTrackedElementState after = CreateTrackedState(elementId, "after");
        FamilyBrowserElementChangeItem change = new FamilyBrowserElementChangeItem
        {
            ElementId = elementId,
            UniqueId = after.UniqueId,
            ChangeKind = changeKind,
            ElementName = after.ElementName,
            TrackingKind = after.TrackingKind
        };
        if (string.Equals(changeKind, "Created", StringComparison.OrdinalIgnoreCase))
        {
            change.Before = null;
            change.After = after;
        }
        else if (string.Equals(changeKind, "Deleted", StringComparison.OrdinalIgnoreCase))
        {
            change.UniqueId = before.UniqueId;
            change.Before = before;
            change.After = null;
        }
        else if (string.Equals(changeKind, "CreatedThenDeleted", StringComparison.OrdinalIgnoreCase))
        {
            change.UniqueId = string.Empty;
            change.Before = null;
            change.After = null;
        }
        else
        {
            change.Before = before;
            change.After = after;
        }
        return new FamilyBrowserElementChangeCommit
        {
            EntryId = entryId,
            CommittedAtUtc = DateTime.UtcNow.ToString("O"),
            ProjectIdentityPath = projectIdentity,
            ProjectCanonicalPath = FamilyBrowserPathIdentityService.GetCanonicalPath(projectIdentity),
            ProjectComparableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(projectIdentity),
            ProjectLegacyComparableIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(projectIdentity),
            CreatedCount = string.Equals(changeKind, "Created", StringComparison.OrdinalIgnoreCase) ? 1 : 0,
            ModifiedCount = string.Equals(changeKind, "Modified", StringComparison.OrdinalIgnoreCase) ? 1 : 0,
            DeletedCount = string.Equals(changeKind, "Deleted", StringComparison.OrdinalIgnoreCase) ? 1 : 0,
            Changes = new List<FamilyBrowserElementChangeItem>
            {
                change
            }
        };
    }

    private static FamilyBrowserTrackedElementState CreateTrackedState(string elementId, string state)
    {
        return new FamilyBrowserTrackedElementState
        {
            ElementId = elementId,
            UniqueId = "UID-" + (elementId ?? string.Empty),
            ElementClass = "FamilyInstance",
            CategoryName = "Generic Models",
            CategoryId = "-2000151",
            ElementName = "Audit Element",
            FamilyName = "Audit Family",
            TypeName = "Audit Type",
            TypeId = "9001",
            TrackingKind = "Element",
            StateSignature = state ?? string.Empty
        };
    }

    private static FamilyBrowserOperationLogEntry Operation(string id, string candidateKind, string category, string family, string type, string systemKind)
    {
        return new FamilyBrowserOperationLogEntry
        {
            EntryId = id,
            RecordedAtUtc = DateTime.UtcNow.ToString("O"),
            CommittedAtUtc = DateTime.UtcNow.ToString("O"),
            CommitState = "Committed",
            CommitKind = "Save",
            CandidateKind = candidateKind,
            CategoryName = category,
            FamilyName = family,
            TypeName = type,
            SystemFamilyKind = systemKind,
            DocumentPath = @"C:\Projects\Audit.rvt",
            Outcome = "Loaded"
        };
    }

    public sealed class LegacyElementChangeIntegrityV1Payload
    {
        public int SchemaVersion { get; set; }
        public string EntryId { get; set; } = string.Empty;
        public string ProjectTitle { get; set; } = string.Empty;
        public string ProjectIdentityPath { get; set; } = string.Empty;
        public string ProjectCanonicalPath { get; set; } = string.Empty;
        public string ProjectComparableIdentity { get; set; } = string.Empty;
        public string CommitKind { get; set; } = string.Empty;
        public string CommittedAtUtc { get; set; } = string.Empty;
        public string RevitVersion { get; set; } = string.Empty;
        public string RevitUserName { get; set; } = string.Empty;
        public string WindowsUserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
        public string AttributionConfidence { get; set; } = string.Empty;
        public string PolicyValidationState { get; set; } = string.Empty;
        public string CoverageNote { get; set; } = string.Empty;
        public bool IsWorkshared { get; set; }
        public bool BaselineCapturedLate { get; set; }
        public string TrackingStartedAtUtc { get; set; } = string.Empty;
        public string BaselineCapturedAtUtc { get; set; } = string.Empty;
        public long BaselineElapsedMilliseconds { get; set; }
        public int BaselineElementCount { get; set; }
        public int ActivityCount { get; set; }
        public int UndoCount { get; set; }
        public int RedoCount { get; set; }
        public int CreatedCount { get; set; }
        public int ModifiedCount { get; set; }
        public int DeletedCount { get; set; }
        public int ExternalUpdateOverlapCount { get; set; }
        public int IntegrityVersion { get; set; }
        public string IntegritySha256 { get; set; } = string.Empty;
        public List<string> TransactionNames { get; set; } = new List<string>();
        public List<LegacyElementChangeItemV1Payload> Changes { get; set; } = new List<LegacyElementChangeItemV1Payload>();
    }

    public sealed class LegacyElementChangeItemV1Payload
    {
        public string ChangeKind { get; set; } = string.Empty;
        public string ElementId { get; set; } = string.Empty;
        public string UniqueId { get; set; } = string.Empty;
        public string ElementClass { get; set; } = string.Empty;
        public string CategoryName { get; set; } = string.Empty;
        public string FamilyName { get; set; } = string.Empty;
        public string TypeName { get; set; } = string.Empty;
        public string FirstObservedAtUtc { get; set; } = string.Empty;
        public string LastObservedAtUtc { get; set; } = string.Empty;
        public string ChangeSummary { get; set; } = string.Empty;
        public bool PreviousStateUnavailable { get; set; }
        public bool ExternalUpdateOverlap { get; set; }
        public LegacyTrackedElementStateV1Payload Before { get; set; }
        public LegacyTrackedElementStateV1Payload After { get; set; }
        public List<string> TransactionNames { get; set; } = new List<string>();
    }

    public sealed class LegacyTrackedElementStateV1Payload
    {
        public string ElementId { get; set; } = string.Empty;
        public string UniqueId { get; set; } = string.Empty;
        public string ElementClass { get; set; } = string.Empty;
        public string CategoryName { get; set; } = string.Empty;
        public string CategoryId { get; set; } = string.Empty;
        public string ElementName { get; set; } = string.Empty;
        public string FamilyName { get; set; } = string.Empty;
        public string TypeName { get; set; } = string.Empty;
        public string TypeId { get; set; } = string.Empty;
        public string LevelId { get; set; } = string.Empty;
        public string WorksetId { get; set; } = string.Empty;
        public string LocationSignature { get; set; } = string.Empty;
        public string StateSignature { get; set; } = string.Empty;
        public bool IsElementType { get; set; }
        public bool IsViewSpecific { get; set; }
    }

    public sealed class LegacyPendingElementEnvelopeV1EntryPayload
    {
        public int SchemaVersion { get; set; }
        public string EntryId { get; set; } = string.Empty;
        public string ProjectTitle { get; set; } = string.Empty;
        public string ProjectIdentityPath { get; set; } = string.Empty;
        public string ProjectCanonicalPath { get; set; } = string.Empty;
        public string ProjectComparableIdentity { get; set; } = string.Empty;
        public string ProjectLegacyComparableIdentity { get; set; } = string.Empty;
        public string CommitKind { get; set; } = string.Empty;
        public string CommittedAtUtc { get; set; } = string.Empty;
        public string RevitVersion { get; set; } = string.Empty;
        public string RevitUserName { get; set; } = string.Empty;
        public string WindowsUserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
        public string AttributionConfidence { get; set; } = string.Empty;
        public string PolicyValidationState { get; set; } = string.Empty;
        public string CoverageNote { get; set; } = string.Empty;
        public bool IsWorkshared { get; set; }
        public bool BaselineCapturedLate { get; set; }
        public string TrackingStartedAtUtc { get; set; } = string.Empty;
        public string BaselineCapturedAtUtc { get; set; } = string.Empty;
        public long BaselineElapsedMilliseconds { get; set; }
        public int BaselineElementCount { get; set; }
        public int ActivityCount { get; set; }
        public int UndoCount { get; set; }
        public int RedoCount { get; set; }
        public int UnmatchedUndoCount { get; set; }
        public int UnmatchedRedoCount { get; set; }
        public int CreatedCount { get; set; }
        public int ModifiedCount { get; set; }
        public int DeletedCount { get; set; }
        public int ExternalUpdateOverlapCount { get; set; }
        public int IntegrityVersion { get; set; }
        public string IntegritySha256 { get; set; } = string.Empty;
        public List<string> TransactionNames { get; set; } = new List<string>();
        public List<LegacyElementChangeItemV1Payload> Changes { get; set; } = new List<LegacyElementChangeItemV1Payload>();
    }

    public sealed class LegacyPendingElementEnvelopeV1Payload
    {
        public string DestinationIdentity { get; set; } = string.Empty;
        public int EnvelopeIntegrityVersion { get; set; }
        public string EnvelopeIntegritySha256 { get; set; } = string.Empty;
        public LegacyPendingElementEnvelopeV1EntryPayload Entry { get; set; } = new LegacyPendingElementEnvelopeV1EntryPayload();
    }

    private static void WriteLegacyV1Commit(string managedRoot, FamilyBrowserElementChangeCommit commit)
    {
        LegacyElementChangeIntegrityV1Payload payload = new LegacyElementChangeIntegrityV1Payload
        {
            SchemaVersion = 1,
            EntryId = commit.EntryId,
            ProjectTitle = commit.ProjectTitle,
            ProjectIdentityPath = commit.ProjectIdentityPath,
            ProjectCanonicalPath = commit.ProjectCanonicalPath,
            ProjectComparableIdentity = commit.ProjectComparableIdentity,
            CommitKind = commit.CommitKind,
            CommittedAtUtc = commit.CommittedAtUtc,
            RevitVersion = commit.RevitVersion,
            RevitUserName = commit.RevitUserName,
            WindowsUserName = commit.WindowsUserName,
            MachineName = commit.MachineName,
            AttributionConfidence = commit.AttributionConfidence,
            PolicyValidationState = commit.PolicyValidationState,
            CoverageNote = commit.CoverageNote,
            IsWorkshared = commit.IsWorkshared,
            BaselineCapturedLate = commit.BaselineCapturedLate,
            TrackingStartedAtUtc = commit.TrackingStartedAtUtc,
            BaselineCapturedAtUtc = commit.BaselineCapturedAtUtc,
            BaselineElapsedMilliseconds = commit.BaselineElapsedMilliseconds,
            BaselineElementCount = commit.BaselineElementCount,
            ActivityCount = commit.ActivityCount,
            UndoCount = commit.UndoCount,
            RedoCount = commit.RedoCount,
            CreatedCount = commit.CreatedCount,
            ModifiedCount = commit.ModifiedCount,
            DeletedCount = commit.DeletedCount,
            ExternalUpdateOverlapCount = commit.ExternalUpdateOverlapCount,
            IntegrityVersion = 1,
            TransactionNames = commit.TransactionNames,
            Changes = CreateLegacyElementChanges(commit.Changes)
        };
        byte[] unsigned = SerializeLegacyV1(payload);
        using (SHA256 sha = SHA256.Create())
        {
            payload.IntegritySha256 = Convert.ToHexString(sha.ComputeHash(unsigned));
        }
        string identity = commit.ProjectComparableIdentity;
        string projectKey;
        using (SHA256 sha = SHA256.Create())
        {
            projectKey = Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(identity ?? string.Empty))).Substring(0, 32);
        }
        string day = DateTime.Parse(commit.CommittedAtUtc).ToUniversalTime().ToString("yyyyMMdd");
        string folder = Path.Combine(managedRoot, "ElementChangeHistory", projectKey, day);
        Directory.CreateDirectory(folder);
        File.WriteAllBytes(Path.Combine(folder, commit.EntryId + ".json"), SerializeLegacyV1(payload));
    }

    private static void WriteLegacyPendingV1Envelope(string spoolRoot, string managedRoot, FamilyBrowserElementChangeCommit commit)
    {
        LegacyElementChangeIntegrityV1Payload immutablePayload = CreateLegacyV1Payload(commit);
        using (SHA256 sha = SHA256.Create())
        {
            immutablePayload.IntegritySha256 = Convert.ToHexString(sha.ComputeHash(SerializeLegacyV1(immutablePayload)));
        }
        commit.IntegrityVersion = 1;
        commit.IntegritySha256 = immutablePayload.IntegritySha256;

        LegacyPendingElementEnvelopeV1Payload envelope = new LegacyPendingElementEnvelopeV1Payload
        {
            DestinationIdentity = "MANAGED-PATH:" + FamilyBrowserPathIdentityService.GetCanonicalPath(Path.Combine(managedRoot, "Config", "standard-policy.json")).ToUpperInvariant(),
            EnvelopeIntegrityVersion = 1,
            EnvelopeIntegritySha256 = string.Empty,
            Entry = CreateLegacyPendingV1Entry(commit)
        };
        using (SHA256 sha = SHA256.Create())
        {
            envelope.EnvelopeIntegritySha256 = Convert.ToHexString(sha.ComputeHash(SerializeLegacyPendingEnvelopeV1(envelope)));
        }
        string folder = Path.Combine(spoolRoot, "ElementChanges");
        Directory.CreateDirectory(folder);
        File.WriteAllBytes(Path.Combine(folder, commit.EntryId + ".json"), SerializeLegacyPendingEnvelopeV1(envelope));
    }

    private static LegacyElementChangeIntegrityV1Payload CreateLegacyV1Payload(FamilyBrowserElementChangeCommit commit)
    {
        return new LegacyElementChangeIntegrityV1Payload
        {
            SchemaVersion = commit.SchemaVersion,
            EntryId = commit.EntryId,
            ProjectTitle = commit.ProjectTitle,
            ProjectIdentityPath = commit.ProjectIdentityPath,
            ProjectCanonicalPath = commit.ProjectCanonicalPath,
            ProjectComparableIdentity = commit.ProjectComparableIdentity,
            CommitKind = commit.CommitKind,
            CommittedAtUtc = commit.CommittedAtUtc,
            RevitVersion = commit.RevitVersion,
            RevitUserName = commit.RevitUserName,
            WindowsUserName = commit.WindowsUserName,
            MachineName = commit.MachineName,
            AttributionConfidence = commit.AttributionConfidence,
            PolicyValidationState = commit.PolicyValidationState,
            CoverageNote = commit.CoverageNote,
            IsWorkshared = commit.IsWorkshared,
            BaselineCapturedLate = commit.BaselineCapturedLate,
            TrackingStartedAtUtc = commit.TrackingStartedAtUtc,
            BaselineCapturedAtUtc = commit.BaselineCapturedAtUtc,
            BaselineElapsedMilliseconds = commit.BaselineElapsedMilliseconds,
            BaselineElementCount = commit.BaselineElementCount,
            ActivityCount = commit.ActivityCount,
            UndoCount = commit.UndoCount,
            RedoCount = commit.RedoCount,
            CreatedCount = commit.CreatedCount,
            ModifiedCount = commit.ModifiedCount,
            DeletedCount = commit.DeletedCount,
            ExternalUpdateOverlapCount = commit.ExternalUpdateOverlapCount,
            IntegrityVersion = 1,
            IntegritySha256 = string.Empty,
            TransactionNames = commit.TransactionNames,
            Changes = CreateLegacyElementChanges(commit.Changes)
        };
    }

    private static LegacyPendingElementEnvelopeV1EntryPayload CreateLegacyPendingV1Entry(FamilyBrowserElementChangeCommit commit)
    {
        return new LegacyPendingElementEnvelopeV1EntryPayload
        {
            SchemaVersion = commit.SchemaVersion,
            EntryId = commit.EntryId,
            ProjectTitle = commit.ProjectTitle,
            ProjectIdentityPath = commit.ProjectIdentityPath,
            ProjectCanonicalPath = commit.ProjectCanonicalPath,
            ProjectComparableIdentity = commit.ProjectComparableIdentity,
            ProjectLegacyComparableIdentity = commit.ProjectLegacyComparableIdentity,
            CommitKind = commit.CommitKind,
            CommittedAtUtc = commit.CommittedAtUtc,
            RevitVersion = commit.RevitVersion,
            RevitUserName = commit.RevitUserName,
            WindowsUserName = commit.WindowsUserName,
            MachineName = commit.MachineName,
            AttributionConfidence = commit.AttributionConfidence,
            PolicyValidationState = commit.PolicyValidationState,
            CoverageNote = commit.CoverageNote,
            IsWorkshared = commit.IsWorkshared,
            BaselineCapturedLate = commit.BaselineCapturedLate,
            TrackingStartedAtUtc = commit.TrackingStartedAtUtc,
            BaselineCapturedAtUtc = commit.BaselineCapturedAtUtc,
            BaselineElapsedMilliseconds = commit.BaselineElapsedMilliseconds,
            BaselineElementCount = commit.BaselineElementCount,
            ActivityCount = commit.ActivityCount,
            UndoCount = commit.UndoCount,
            RedoCount = commit.RedoCount,
            UnmatchedUndoCount = commit.UnmatchedUndoCount,
            UnmatchedRedoCount = commit.UnmatchedRedoCount,
            CreatedCount = commit.CreatedCount,
            ModifiedCount = commit.ModifiedCount,
            DeletedCount = commit.DeletedCount,
            ExternalUpdateOverlapCount = commit.ExternalUpdateOverlapCount,
            IntegrityVersion = commit.IntegrityVersion,
            IntegritySha256 = commit.IntegritySha256,
            TransactionNames = commit.TransactionNames,
            Changes = CreateLegacyElementChanges(commit.Changes)
        };
    }

    private static List<LegacyElementChangeItemV1Payload> CreateLegacyElementChanges(IEnumerable<FamilyBrowserElementChangeItem> changes)
    {
        return (changes ?? Enumerable.Empty<FamilyBrowserElementChangeItem>())
            .Select(delegate(FamilyBrowserElementChangeItem item)
            {
                if (item == null) return null;
                return new LegacyElementChangeItemV1Payload
                {
                    ChangeKind = item.ChangeKind,
                    ElementId = item.ElementId,
                    UniqueId = item.UniqueId,
                    ElementClass = item.ElementClass,
                    CategoryName = item.CategoryName,
                    FamilyName = item.FamilyName,
                    TypeName = item.TypeName,
                    FirstObservedAtUtc = item.FirstObservedAtUtc,
                    LastObservedAtUtc = item.LastObservedAtUtc,
                    ChangeSummary = item.ChangeSummary,
                    PreviousStateUnavailable = item.PreviousStateUnavailable,
                    ExternalUpdateOverlap = item.ExternalUpdateOverlap,
                    Before = CreateLegacyTrackedState(item.Before),
                    After = CreateLegacyTrackedState(item.After),
                    TransactionNames = item.TransactionNames
                };
            })
            .ToList();
    }

    private static LegacyTrackedElementStateV1Payload CreateLegacyTrackedState(FamilyBrowserTrackedElementState state)
    {
        if (state == null) return null;
        return new LegacyTrackedElementStateV1Payload
        {
            ElementId = state.ElementId,
            UniqueId = state.UniqueId,
            ElementClass = state.ElementClass,
            CategoryName = state.CategoryName,
            CategoryId = state.CategoryId,
            ElementName = state.ElementName,
            FamilyName = state.FamilyName,
            TypeName = state.TypeName,
            TypeId = state.TypeId,
            LevelId = state.LevelId,
            WorksetId = state.WorksetId,
            LocationSignature = state.LocationSignature,
            StateSignature = state.StateSignature,
            IsElementType = state.IsElementType,
            IsViewSpecific = state.IsViewSpecific
        };
    }

    private static byte[] SerializeLegacyV1(LegacyElementChangeIntegrityV1Payload payload)
    {
        using MemoryStream stream = new MemoryStream();
        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(LegacyElementChangeIntegrityV1Payload));
        serializer.WriteObject(stream, payload);
        return stream.ToArray();
    }

    private static byte[] SerializeLegacyPendingEnvelopeV1(LegacyPendingElementEnvelopeV1Payload payload)
    {
        using MemoryStream stream = new MemoryStream();
        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(LegacyPendingElementEnvelopeV1Payload));
        serializer.WriteObject(stream, payload);
        return stream.ToArray();
    }

    private static FamilyBrowserElementActivityMatchInput ActivityMatch(IEnumerable<string> elementIds, IEnumerable<string> transactionNames)
    {
        return new FamilyBrowserElementActivityMatchInput
        {
            ElementIds = (elementIds ?? Enumerable.Empty<string>()).ToList(),
            TransactionNames = (transactionNames ?? Enumerable.Empty<string>()).ToList()
        };
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
        {
            throw new InvalidOperationException(message);
        }
    }
}
