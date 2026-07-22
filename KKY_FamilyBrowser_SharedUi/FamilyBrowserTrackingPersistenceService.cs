using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

public sealed class FamilyBrowserTrackingFlushResult
{
    public int OperationFlushedCount { get; set; }
    public int StandardCandidateFlushedCount { get; set; }
    public int ElementChangeFlushedCount { get; set; }
    public int FailedCount { get; set; }
    public int DestinationMismatchCount { get; set; }
    public int CorruptRecordCount { get; set; }
    public int ElementSessionCheckpointReboundCount { get; set; }
    public int ElementSessionCheckpointRebindFailedCount { get; set; }
    public int FinalizedElementSessionPromotedCount { get; set; }
    public bool ElementSessionCheckpointLockUnavailable { get; set; }
}

public sealed class FamilyBrowserPendingOperationRecord
{
    public string DestinationIdentity { get; set; }
    public int EnvelopeIntegrityVersion { get; set; }
    public string EnvelopeIntegritySha256 { get; set; }
    public FamilyBrowserOperationLogEntry Entry { get; set; }

    public FamilyBrowserPendingOperationRecord()
    {
        DestinationIdentity = string.Empty;
        EnvelopeIntegritySha256 = string.Empty;
    }
}

public sealed class FamilyBrowserPendingStandardCandidateRecord
{
    public string DestinationIdentity { get; set; }
    public string SourceId { get; set; }
    public int EnvelopeIntegrityVersion { get; set; }
    public string EnvelopeIntegritySha256 { get; set; }
    public StandardRvtChangeCandidateEntry Entry { get; set; }

    public FamilyBrowserPendingStandardCandidateRecord()
    {
        DestinationIdentity = string.Empty;
        SourceId = string.Empty;
        EnvelopeIntegritySha256 = string.Empty;
    }
}

public sealed class FamilyBrowserPendingElementChangeRecord
{
    public string DestinationIdentity { get; set; }
    public int EnvelopeIntegrityVersion { get; set; }
    public string EnvelopeIntegritySha256 { get; set; }
    public FamilyBrowserElementChangeCommit Entry { get; set; }

    public FamilyBrowserPendingElementChangeRecord()
    {
        DestinationIdentity = string.Empty;
        EnvelopeIntegritySha256 = string.Empty;
    }
}

public sealed class FamilyBrowserElementChangeHistoryLoadResult
{
    public List<FamilyBrowserElementChangeCommit> Commits { get; set; }
    public int ScannedFileCount { get; set; }
    public int TotalValidRecordCount { get; set; }
    public int InvalidRecordCount { get; set; }
    public int LegacyUnverifiedCount { get; set; }
    public int PendingDestinationMismatchCount { get; set; }
    public int PendingCorruptRecordCount { get; set; }
    public int PendingFailedCount { get; set; }

    public FamilyBrowserElementChangeHistoryLoadResult()
    {
        Commits = new List<FamilyBrowserElementChangeCommit>();
    }
}

public sealed class FamilyBrowserElementSessionCheckpoint
{
    public int SchemaVersion { get; set; }
    public string CheckpointId { get; set; }
    public string DestinationIdentity { get; set; }
    public string ProjectStableIdentity { get; set; }
    public string LocalDocumentStableIdentity { get; set; }
    public string RevitUserName { get; set; }
    public string UpdatedAtUtc { get; set; }
    public bool SynchronizationSucceeded { get; set; }
    public int EnvelopeIntegrityVersion { get; set; }
    public string EnvelopeIntegritySha256 { get; set; }
    public List<FamilyBrowserElementChangeCommit> Commits { get; set; }

    public FamilyBrowserElementSessionCheckpoint()
    {
        SchemaVersion = 1;
        CheckpointId = string.Empty;
        DestinationIdentity = string.Empty;
        ProjectStableIdentity = string.Empty;
        LocalDocumentStableIdentity = string.Empty;
        RevitUserName = string.Empty;
        UpdatedAtUtc = string.Empty;
        EnvelopeIntegritySha256 = string.Empty;
        Commits = new List<FamilyBrowserElementChangeCommit>();
    }
}

public sealed class FamilyBrowserElementSessionCheckpointLoadResult
{
    public FamilyBrowserElementSessionCheckpoint Checkpoint { get; set; }
    public bool Invalid { get; set; }
    public bool DestinationMismatch { get; set; }
    public bool LockUnavailable { get; set; }
}

public sealed class FamilyBrowserElementSessionCheckpointCountResult
{
    public int Count { get; set; }
    public int SynchronizationSucceededCount { get; set; }
    public bool LockUnavailable { get; set; }
}

public sealed class FamilyBrowserElementSessionCheckpointHistoryLoadResult
{
    public List<FamilyBrowserElementChangeCommit> Commits { get; set; }
    public HashSet<string> SynchronizationSucceededEntryIds { get; set; }
    public int TotalValidRecordCount { get; set; }
    public int InvalidRecordCount { get; set; }
    public int DestinationMismatchCount { get; set; }
    public bool LockUnavailable { get; set; }

    public FamilyBrowserElementSessionCheckpointHistoryLoadResult()
    {
        Commits = new List<FamilyBrowserElementChangeCommit>();
        SynchronizationSucceededEntryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    }
}

public sealed class FamilyBrowserTrackedProjectHistorySummary
{
    public string ProjectIdentityPath { get; set; }
    public string ProjectComparableIdentity { get; set; }
    public string ProjectTitle { get; set; }
    public string LastActivityAtUtc { get; set; }
    public int ConfirmedCommitCount { get; set; }
    public int UploadPendingCommitCount { get; set; }
    public int LocalSavePendingCommitCount { get; set; }
    public int CreatedCount { get; set; }
    public int ModifiedCount { get; set; }
    public int DeletedCount { get; set; }

    public FamilyBrowserTrackedProjectHistorySummary()
    {
        ProjectIdentityPath = string.Empty;
        ProjectComparableIdentity = string.Empty;
        ProjectTitle = string.Empty;
        LastActivityAtUtc = string.Empty;
    }
}

public static class FamilyBrowserTrackingPersistenceService
{
    // Integrity v1 used DataContractJsonSerializer over this exact schema. Keep the
    // projection frozen so future model fields do not invalidate valid old records.
    public sealed class ElementChangeIntegrityV1Payload
    {
        public int SchemaVersion { get; set; }
        public string EntryId { get; set; }
        public string ProjectTitle { get; set; }
        public string ProjectIdentityPath { get; set; }
        public string ProjectCanonicalPath { get; set; }
        public string ProjectComparableIdentity { get; set; }
        public string CommitKind { get; set; }
        public string CommittedAtUtc { get; set; }
        public string RevitVersion { get; set; }
        public string RevitUserName { get; set; }
        public string WindowsUserName { get; set; }
        public string MachineName { get; set; }
        public string AttributionConfidence { get; set; }
        public string PolicyValidationState { get; set; }
        public string CoverageNote { get; set; }
        public bool IsWorkshared { get; set; }
        public bool BaselineCapturedLate { get; set; }
        public string TrackingStartedAtUtc { get; set; }
        public string BaselineCapturedAtUtc { get; set; }
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
        public string IntegritySha256 { get; set; }
        public List<string> TransactionNames { get; set; }
        public List<ElementChangeIntegrityV1ItemPayload> Changes { get; set; }
    }

    public sealed class ElementChangeIntegrityV1ItemPayload
    {
        public string ChangeKind { get; set; }
        public string ElementId { get; set; }
        public string UniqueId { get; set; }
        public string ElementClass { get; set; }
        public string CategoryName { get; set; }
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        public string FirstObservedAtUtc { get; set; }
        public string LastObservedAtUtc { get; set; }
        public string ChangeSummary { get; set; }
        public bool PreviousStateUnavailable { get; set; }
        public bool ExternalUpdateOverlap { get; set; }
        public ElementChangeIntegrityV1StatePayload Before { get; set; }
        public ElementChangeIntegrityV1StatePayload After { get; set; }
        public List<string> TransactionNames { get; set; }
    }

    public sealed class ElementChangeIntegrityV1StatePayload
    {
        public string ElementId { get; set; }
        public string UniqueId { get; set; }
        public string ElementClass { get; set; }
        public string CategoryName { get; set; }
        public string CategoryId { get; set; }
        public string ElementName { get; set; }
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        public string TypeId { get; set; }
        public string LevelId { get; set; }
        public string WorksetId { get; set; }
        public string LocationSignature { get; set; }
        public string StateSignature { get; set; }
        public bool IsElementType { get; set; }
        public bool IsViewSpecific { get; set; }
    }

    // Pending-envelope integrity v1 was introduced after the immutable-history
    // v1 schema. It serialized the then-current full commit model, which already
    // contained the stable/legacy identity and unmatched Undo/Redo counters.
    // Keep this separate projection frozen so either schema can evolve safely.
    public sealed class PendingElementEnvelopeIntegrityV1EntryPayload
    {
        public int SchemaVersion { get; set; }
        public string EntryId { get; set; }
        public string ProjectTitle { get; set; }
        public string ProjectIdentityPath { get; set; }
        public string ProjectCanonicalPath { get; set; }
        public string ProjectComparableIdentity { get; set; }
        public string ProjectLegacyComparableIdentity { get; set; }
        public string CommitKind { get; set; }
        public string CommittedAtUtc { get; set; }
        public string RevitVersion { get; set; }
        public string RevitUserName { get; set; }
        public string WindowsUserName { get; set; }
        public string MachineName { get; set; }
        public string AttributionConfidence { get; set; }
        public string PolicyValidationState { get; set; }
        public string CoverageNote { get; set; }
        public bool IsWorkshared { get; set; }
        public bool BaselineCapturedLate { get; set; }
        public string TrackingStartedAtUtc { get; set; }
        public string BaselineCapturedAtUtc { get; set; }
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
        public string IntegritySha256 { get; set; }
        public List<string> TransactionNames { get; set; }
        public List<ElementChangeIntegrityV1ItemPayload> Changes { get; set; }
    }

    public sealed class PendingElementEnvelopeIntegrityV1Payload
    {
        public string DestinationIdentity { get; set; }
        public int EnvelopeIntegrityVersion { get; set; }
        public string EnvelopeIntegritySha256 { get; set; }
        public PendingElementEnvelopeIntegrityV1EntryPayload Entry { get; set; }
    }

    private static readonly object SyncRoot = new object();
    private static readonly object DeferredSpoolRoot = new object();
    private static readonly object DeferredFlushStateRoot = new object();
    private static readonly HashSet<string> DeferredFlushActiveRoots = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private static readonly HashSet<string> DeferredFlushRequestedRoots = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private const int MinimumSupportedElementCommitSchemaVersion = 1;
    private const int MaximumSupportedElementCommitSchemaVersion = 6;
    private const int ElementSessionLockTimeoutMilliseconds = 750;
    private const int ElementSessionLockRetryMilliseconds = 25;
    private static string _localSpoolRootOverrideForAudit = string.Empty;
    private static string _enumerationFailureFolderOverrideForAudit = string.Empty;

    public static string LocalSpoolRootOverrideForAudit
    {
        get
        {
            lock (SyncRoot)
            {
                return _localSpoolRootOverrideForAudit;
            }
        }
        set
        {
            lock (SyncRoot)
            {
                _localSpoolRootOverrideForAudit = value ?? string.Empty;
            }
        }
    }

    public static string EnumerationFailureFolderOverrideForAudit
    {
        get
        {
            lock (SyncRoot)
            {
                return _enumerationFailureFolderOverrideForAudit;
            }
        }
        set
        {
            lock (SyncRoot)
            {
                _enumerationFailureFolderOverrideForAudit = value ?? string.Empty;
            }
        }
    }

    public static bool PersistOperationEntries(string workspaceRoot, IEnumerable<FamilyBrowserOperationLogEntry> entries)
    {
        List<FamilyBrowserOperationLogEntry> list = (entries ?? Enumerable.Empty<FamilyBrowserOperationLogEntry>())
            .Where(delegate(FamilyBrowserOperationLogEntry entry) { return entry != null; })
            .ToList();
        if (list.Count == 0)
        {
            return true;
        }
        lock (SyncRoot)
        {
            bool allDurable = true;
            bool managedAvailable = FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot);
            string destinationIdentity = BuildManagedDestinationIdentity(workspaceRoot);
            if (string.IsNullOrWhiteSpace(destinationIdentity))
            {
                return false;
            }
            foreach (FamilyBrowserOperationLogEntry entry in list)
            {
                EnsureEntryId(entry);
                string spoolPath = BuildOperationSpoolPath(entry.EntryId);
                FamilyBrowserPendingOperationRecord record = new FamilyBrowserPendingOperationRecord
                {
                    DestinationIdentity = destinationIdentity,
                    Entry = entry
                };
                EnsurePendingOperationEnvelopeIntegrity(record);
                bool spooled = TryWriteJsonAtomic(spoolPath, record);
                bool managed = managedAvailable && TryWriteJsonAtomic(BuildOperationHistoryPath(workspaceRoot, entry), entry);
                if (managed && spooled)
                {
                    TryDelete(spoolPath);
                }
                if (!spooled && !managed)
                {
                    allDurable = false;
                }
            }
            FlushPendingNoLock(workspaceRoot);
            return allDurable;
        }
    }

    public static bool PersistStandardCandidateEntries(string workspaceRoot, string sourceId, IEnumerable<StandardRvtChangeCandidateEntry> entries)
    {
        List<StandardRvtChangeCandidateEntry> list = (entries ?? Enumerable.Empty<StandardRvtChangeCandidateEntry>())
            .Where(delegate(StandardRvtChangeCandidateEntry entry) { return entry != null; })
            .ToList();
        if (list.Count == 0)
        {
            return true;
        }
        if (string.IsNullOrWhiteSpace(sourceId))
        {
            return false;
        }
        lock (SyncRoot)
        {
            bool allDurable = true;
            bool managedAvailable = FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot);
            string destinationIdentity = BuildManagedDestinationIdentity(workspaceRoot);
            if (string.IsNullOrWhiteSpace(destinationIdentity))
            {
                return false;
            }
            foreach (StandardRvtChangeCandidateEntry entry in list)
            {
                EnsureEntryId(entry);
                FamilyBrowserPendingStandardCandidateRecord record = new FamilyBrowserPendingStandardCandidateRecord
                {
                    DestinationIdentity = destinationIdentity,
                    SourceId = sourceId,
                    Entry = entry
                };
                EnsurePendingStandardCandidateEnvelopeIntegrity(record);
                string spoolPath = BuildCandidateSpoolPath(entry.EntryId);
                bool spooled = TryWriteJsonAtomic(spoolPath, record);
                bool managed = managedAvailable && TryWriteJsonAtomic(BuildCandidateHistoryPath(workspaceRoot, sourceId, entry), entry);
                if (managed && spooled)
                {
                    TryDelete(spoolPath);
                }
                if (!spooled && !managed)
                {
                    allDurable = false;
                }
            }
            FlushPendingNoLock(workspaceRoot);
            return allDurable;
        }
    }

    public static bool PersistElementChangeCommits(string workspaceRoot, IEnumerable<FamilyBrowserElementChangeCommit> commits)
    {
        List<FamilyBrowserElementChangeCommit> supplied = (commits ?? Enumerable.Empty<FamilyBrowserElementChangeCommit>())
            .Where(delegate(FamilyBrowserElementChangeCommit commit) { return commit != null; })
            .ToList();
        if (supplied.Any(delegate(FamilyBrowserElementChangeCommit commit) { return !HasElementChangesOrCoverageGap(commit); }))
        {
            return false;
        }
        List<FamilyBrowserElementChangeCommit> list = supplied;
        if (list.Count == 0)
        {
            return true;
        }
        lock (SyncRoot)
        {
            bool allDurable = true;
            bool managedAvailable = FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot);
            string destinationIdentity = BuildManagedDestinationIdentity(workspaceRoot);
            if (string.IsNullOrWhiteSpace(destinationIdentity))
            {
                return false;
            }
            foreach (FamilyBrowserElementChangeCommit commit in list)
            {
                EnsureEntryId(commit);
                string validationIssue;
                if (!ValidateElementChangeCommit(commit, commit.ProjectComparableIdentity, true, out validationIssue))
                {
                    allDurable = false;
                    continue;
                }
                string spoolPath = BuildElementChangeSpoolPath(commit.EntryId);
                FamilyBrowserPendingElementChangeRecord record = new FamilyBrowserPendingElementChangeRecord
                {
                    DestinationIdentity = destinationIdentity,
                    Entry = commit
                };
                EnsurePendingElementEnvelopeIntegrity(record);
                bool spooled = TryWriteJsonAtomic(spoolPath, record);
                bool managed = managedAvailable && TryWriteElementHistoryAtomic(BuildElementChangeHistoryPath(workspaceRoot, commit), commit);
                if (managed && spooled)
                {
                    TryDelete(spoolPath);
                }
                if (!spooled && !managed)
                {
                    allDurable = false;
                }
            }
            FlushPendingNoLock(workspaceRoot);
            return allDurable;
        }
    }

    public static bool PersistElementChangeCommitsDeferred(string workspaceRoot, IEnumerable<FamilyBrowserElementChangeCommit> commits)
    {
        List<FamilyBrowserElementChangeCommit> supplied = (commits ?? Enumerable.Empty<FamilyBrowserElementChangeCommit>())
            .Where(delegate(FamilyBrowserElementChangeCommit commit) { return commit != null; })
            .ToList();
        if (supplied.Any(delegate(FamilyBrowserElementChangeCommit commit) { return !HasElementChangesOrCoverageGap(commit); }))
        {
            return false;
        }
        if (supplied.Count == 0)
        {
            return true;
        }

        bool allDurable = true;
        string deferredDestinationPath = ResolveDeferredManagedDestinationPath(workspaceRoot);
        lock (DeferredSpoolRoot)
        {
            string destinationIdentity = BuildDeferredManagedDestinationIdentity(deferredDestinationPath);
            if (string.IsNullOrWhiteSpace(destinationIdentity))
            {
                return false;
            }
            foreach (FamilyBrowserElementChangeCommit commit in supplied)
            {
                EnsureEntryId(commit);
                string validationIssue;
                if (!ValidateElementChangeCommit(commit, commit.ProjectComparableIdentity, true, out validationIssue))
                {
                    allDurable = false;
                    continue;
                }
                FamilyBrowserPendingElementChangeRecord record = new FamilyBrowserPendingElementChangeRecord
                {
                    DestinationIdentity = destinationIdentity,
                    Entry = commit
                };
                EnsurePendingElementEnvelopeIntegrity(record);
                if (!TryWriteJsonAtomic(BuildElementChangeSpoolPath(commit.EntryId), record))
                {
                    allDurable = false;
                }
            }
        }

        if (allDurable)
        {
            QueuePendingFlush(workspaceRoot, deferredDestinationPath);
        }
        return allDurable;
    }

    public static void QueuePendingFlush(string workspaceRoot, string deferredDestinationPath)
    {
        string safeWorkspaceRoot = workspaceRoot ?? string.Empty;
        string safeDeferredDestinationPath = deferredDestinationPath ?? string.Empty;
        string flushKey = BuildDeferredManagedDestinationIdentity(safeDeferredDestinationPath);
        if (string.IsNullOrWhiteSpace(flushKey))
        {
            flushKey = safeWorkspaceRoot.Trim();
        }
        if (string.IsNullOrWhiteSpace(flushKey))
        {
            return;
        }

        lock (DeferredFlushStateRoot)
        {
            if (DeferredFlushActiveRoots.Contains(flushKey))
            {
                DeferredFlushRequestedRoots.Add(flushKey);
                return;
            }
            DeferredFlushActiveRoots.Add(flushKey);
        }

        ThreadPool.QueueUserWorkItem(delegate
        {
            // Let the Revit Save/Synchronize completion callback release its UI-thread work first.
            Thread.Sleep(250);
            while (true)
            {
                try
                {
                    PromoteDeferredElementSpoolDestinations(safeDeferredDestinationPath);
                    FlushPending(safeWorkspaceRoot);
                }
                catch
                {
                    // The validated local spool remains authoritative and will be retried later.
                }

                lock (DeferredFlushStateRoot)
                {
                    if (DeferredFlushRequestedRoots.Remove(flushKey))
                    {
                        continue;
                    }
                    DeferredFlushActiveRoots.Remove(flushKey);
                    break;
                }
            }
        });
    }

    public static bool SaveElementSessionCheckpoint(
        string workspaceRoot,
        string projectIdentity,
        string localDocumentPath,
        string revitUserName,
        IEnumerable<FamilyBrowserElementChangeCommit> commits,
        bool synchronizationSucceeded,
        string expectedCheckpointRevisionToken,
        out string savedCheckpointRevisionToken)
    {
        savedCheckpointRevisionToken = string.Empty;
        List<FamilyBrowserElementChangeCommit> supplied = (commits ?? Enumerable.Empty<FamilyBrowserElementChangeCommit>())
            .Where(delegate(FamilyBrowserElementChangeCommit commit) { return commit != null; })
            .ToList();
        if (supplied.Any(delegate(FamilyBrowserElementChangeCommit commit) { return !HasElementChangesOrCoverageGap(commit); }))
        {
            return false;
        }
        List<FamilyBrowserElementChangeCommit> list = supplied;
        if (list.Count == 0)
        {
            return DeleteElementSessionCheckpoint(projectIdentity, localDocumentPath, revitUserName, expectedCheckpointRevisionToken);
        }
        lock (SyncRoot)
        {
            using (FileStream checkpointLock = TryAcquireElementSessionFileLock())
            {
                if (checkpointLock == null)
                {
                    return false;
                }
                string expectedProjectIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(projectIdentity);
                foreach (FamilyBrowserElementChangeCommit commit in list)
                {
                    if (!ElementCommitMatchesCheckpointProject(commit, expectedProjectIdentity))
                    {
                        return false;
                    }
                    EnsureEntryId(commit);
                    string validationIssue;
                    if (!ValidateElementChangeCommit(commit, expectedProjectIdentity, true, out validationIssue))
                    {
                        return false;
                    }
                }
                List<IGrouping<string, FamilyBrowserElementChangeCommit>> duplicateGroups = list
                    .GroupBy(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId ?? string.Empty; }, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                foreach (IGrouping<string, FamilyBrowserElementChangeCommit> group in duplicateGroups)
                {
                    FamilyBrowserElementChangeCommit first = group.First();
                    if (group.Any(delegate(FamilyBrowserElementChangeCommit candidate)
                    {
                        return candidate.IntegrityVersion != first.IntegrityVersion ||
                            !FixedTimeEquals(candidate.IntegritySha256, first.IntegritySha256);
                    }))
                    {
                        return false;
                    }
                }
                list = duplicateGroups
                    .Select(delegate(IGrouping<string, FamilyBrowserElementChangeCommit> group) { return group.First(); })
                    .ToList();
                FamilyBrowserElementSessionCheckpoint checkpoint = BuildElementSessionCheckpoint(
                    workspaceRoot,
                    projectIdentity,
                    localDocumentPath,
                    revitUserName,
                    list,
                    synchronizationSucceeded);
                if (string.IsNullOrWhiteSpace(checkpoint.CheckpointId) || string.IsNullOrWhiteSpace(checkpoint.DestinationIdentity))
                {
                    return false;
                }
                string checkpointPath = BuildElementSessionCheckpointPath(checkpoint.CheckpointId);
                if (File.Exists(checkpointPath))
                {
                    FamilyBrowserElementSessionCheckpoint existing;
                    if (!TryReadJson(checkpointPath, out existing) || existing == null ||
                        !ValidateElementSessionCheckpoint(existing, projectIdentity, localDocumentPath, revitUserName) ||
                        !ValidateElementSessionCheckpointCommits(existing) ||
                        !CanFlushToDestination(existing.DestinationIdentity, checkpoint.DestinationIdentity, false) ||
                        string.IsNullOrWhiteSpace(expectedCheckpointRevisionToken) ||
                        !FixedTimeEquals(ComputeElementSessionCheckpointRevisionToken(existing), expectedCheckpointRevisionToken))
                    {
                        return false;
                    }
                }
                else if (!string.IsNullOrWhiteSpace(expectedCheckpointRevisionToken))
                {
                    return false;
                }
                if (!TryWriteMutableJsonAtomic(checkpointPath, checkpoint))
                {
                    return false;
                }
                savedCheckpointRevisionToken = ComputeElementSessionCheckpointRevisionToken(checkpoint);
                return !string.IsNullOrWhiteSpace(savedCheckpointRevisionToken);
            }
        }
    }

    public static FamilyBrowserElementSessionCheckpointLoadResult LoadElementSessionCheckpoint(
        string workspaceRoot,
        string projectIdentity,
        string localDocumentPath,
        string revitUserName)
    {
        FamilyBrowserElementSessionCheckpointLoadResult result = new FamilyBrowserElementSessionCheckpointLoadResult();
        string checkpointId = BuildElementSessionCheckpointId(projectIdentity, localDocumentPath, revitUserName);
        if (string.IsNullOrWhiteSpace(checkpointId))
        {
            return result;
        }
        lock (SyncRoot)
        {
            using (FileStream checkpointLock = TryAcquireElementSessionFileLock())
            {
                if (checkpointLock == null)
                {
                    result.LockUnavailable = true;
                    return result;
                }
                string checkpointPath = BuildElementSessionCheckpointPath(checkpointId);
                FamilyBrowserElementSessionCheckpoint checkpoint;
                if (!TryReadJson(checkpointPath, out checkpoint) || checkpoint == null)
                {
                    result.Invalid = File.Exists(checkpointPath);
                    return result;
                }
                if (!ValidateElementSessionCheckpoint(checkpoint, projectIdentity, localDocumentPath, revitUserName))
                {
                    result.Invalid = true;
                    return result;
                }
                if (!ValidateElementSessionCheckpointCommits(checkpoint))
                {
                    result.Invalid = true;
                    return result;
                }
                string currentDestination = BuildManagedDestinationIdentity(workspaceRoot);
                if (!CanFlushToDestination(checkpoint.DestinationIdentity, currentDestination, false))
                {
                    result.DestinationMismatch = true;
                    return result;
                }
                result.Checkpoint = checkpoint;
                return result;
            }
        }
    }

    public static bool DeleteElementSessionCheckpoint(
        string projectIdentity,
        string localDocumentPath,
        string revitUserName,
        string expectedCheckpointRevisionToken)
    {
        string checkpointId = BuildElementSessionCheckpointId(projectIdentity, localDocumentPath, revitUserName);
        if (string.IsNullOrWhiteSpace(checkpointId))
        {
            return true;
        }
        lock (SyncRoot)
        {
            using (FileStream checkpointLock = TryAcquireElementSessionFileLock())
            {
                if (checkpointLock == null)
                {
                    return false;
                }
                string checkpointPath = BuildElementSessionCheckpointPath(checkpointId);
                if (!File.Exists(checkpointPath))
                {
                    return true;
                }
                FamilyBrowserElementSessionCheckpoint checkpoint;
                if (!TryReadJson(checkpointPath, out checkpoint) || checkpoint == null ||
                    !ValidateElementSessionCheckpoint(checkpoint, projectIdentity, localDocumentPath, revitUserName) ||
                    !ValidateElementSessionCheckpointCommits(checkpoint) ||
                    string.IsNullOrWhiteSpace(expectedCheckpointRevisionToken) ||
                    !FixedTimeEquals(ComputeElementSessionCheckpointRevisionToken(checkpoint), expectedCheckpointRevisionToken))
                {
                    return false;
                }
                return TryDeleteChecked(checkpointPath);
            }
        }
    }

    public static string GetElementSessionCheckpointRevisionToken(FamilyBrowserElementSessionCheckpoint checkpoint)
    {
        if (!ValidateElementSessionCheckpointEnvelope(checkpoint) || !ValidateElementSessionCheckpointCommits(checkpoint))
        {
            return string.Empty;
        }
        return ComputeElementSessionCheckpointRevisionToken(checkpoint);
    }

    public static int DeleteElementSessionCheckpointsForDestination(string workspaceRoot)
    {
        string destinationIdentity = BuildManagedDestinationIdentity(workspaceRoot);
        int deleted = 0;
        lock (SyncRoot)
        {
            using (FileStream checkpointLock = TryAcquireElementSessionFileLock())
            {
                if (checkpointLock == null)
                {
                    return 0;
                }
                List<string> checkpointPaths;
                if (!TryEnumerateSpoolFiles(BuildSpoolFolder("ElementSessions"), out checkpointPaths))
                {
                    return 0;
                }
                foreach (string path in checkpointPaths)
                {
                    FamilyBrowserElementSessionCheckpoint checkpoint;
                    if (TryReadJson(path, out checkpoint) && checkpoint != null &&
                        ValidateElementSessionCheckpointEnvelope(checkpoint) &&
                        ValidateElementSessionCheckpointCommits(checkpoint) &&
                        CanFlushToDestination(checkpoint.DestinationIdentity, destinationIdentity, false) &&
                        TryDeleteChecked(path))
                    {
                        deleted++;
                    }
                }
            }
        }
        return deleted;
    }

    public static int GetPendingElementSessionCheckpointCount(string workspaceRoot, string projectIdentity = "")
    {
        return GetPendingElementSessionCheckpointStatus(workspaceRoot, projectIdentity).Count;
    }

    public static FamilyBrowserElementSessionCheckpointCountResult GetPendingElementSessionCheckpointStatus(string workspaceRoot, string projectIdentity = "")
    {
        FamilyBrowserElementSessionCheckpointCountResult result = new FamilyBrowserElementSessionCheckpointCountResult();
        string destinationIdentity = BuildManagedDestinationIdentity(workspaceRoot);
        string stableProject = FamilyBrowserPathIdentityService.GetStablePathIdentity(projectIdentity);
        lock (SyncRoot)
        {
            using (FileStream checkpointLock = TryAcquireElementSessionFileLock())
            {
                if (checkpointLock == null)
                {
                    result.LockUnavailable = true;
                    return result;
                }
                List<string> checkpointPaths;
                if (!TryEnumerateSpoolFiles(BuildSpoolFolder("ElementSessions"), out checkpointPaths))
                {
                    result.LockUnavailable = true;
                    return result;
                }
                foreach (string path in checkpointPaths)
                {
                    FamilyBrowserElementSessionCheckpoint checkpoint;
                    if (!TryReadJson(path, out checkpoint) || checkpoint == null ||
                        !ValidateElementSessionCheckpointEnvelope(checkpoint) ||
                        !ValidateElementSessionCheckpointCommits(checkpoint) ||
                        !CanFlushToDestination(checkpoint.DestinationIdentity, destinationIdentity, false))
                    {
                        continue;
                    }
                    if (string.IsNullOrWhiteSpace(stableProject) ||
                        string.Equals(checkpoint.ProjectStableIdentity, stableProject, StringComparison.OrdinalIgnoreCase))
                    {
                        result.Count++;
                        if (checkpoint.SynchronizationSucceeded)
                        {
                            result.SynchronizationSucceededCount++;
                        }
                    }
                }
            }
        }
        return result;
    }

    public static FamilyBrowserElementSessionCheckpointHistoryLoadResult LoadElementSessionCheckpointHistory(
        string workspaceRoot,
        string projectIdentity,
        int limit)
    {
        FamilyBrowserElementSessionCheckpointHistoryLoadResult result = new FamilyBrowserElementSessionCheckpointHistoryLoadResult();
        string destinationIdentity = BuildManagedDestinationIdentity(workspaceRoot);
        string stableProject = FamilyBrowserPathIdentityService.GetStablePathIdentity(projectIdentity);
        lock (SyncRoot)
        {
            using (FileStream checkpointLock = TryAcquireElementSessionFileLock())
            {
                if (checkpointLock == null)
                {
                    result.LockUnavailable = true;
                    return result;
                }
                List<string> checkpointPaths;
                if (!TryEnumerateSpoolFiles(BuildSpoolFolder("ElementSessions"), out checkpointPaths))
                {
                    result.LockUnavailable = true;
                    return result;
                }
                Dictionary<string, FamilyBrowserElementChangeCommit> loaded = new Dictionary<string, FamilyBrowserElementChangeCommit>(StringComparer.OrdinalIgnoreCase);
                HashSet<string> conflictingEntryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (string path in checkpointPaths)
                {
                    FamilyBrowserElementSessionCheckpoint checkpoint;
                    if (!TryReadJson(path, out checkpoint) || checkpoint == null ||
                        !ValidateElementSessionCheckpointEnvelope(checkpoint) ||
                        !ValidateElementSessionCheckpointCommits(checkpoint))
                    {
                        result.InvalidRecordCount++;
                        continue;
                    }
                    if (!CanFlushToDestination(checkpoint.DestinationIdentity, destinationIdentity, false))
                    {
                        result.DestinationMismatchCount++;
                        continue;
                    }
                    if (!string.IsNullOrWhiteSpace(stableProject) &&
                        !string.Equals(checkpoint.ProjectStableIdentity, stableProject, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    foreach (FamilyBrowserElementChangeCommit commit in checkpoint.Commits ?? new List<FamilyBrowserElementChangeCommit>())
                    {
                        if (commit == null || string.IsNullOrWhiteSpace(commit.EntryId) || conflictingEntryIds.Contains(commit.EntryId))
                        {
                            continue;
                        }
                        FamilyBrowserElementChangeCommit existing;
                        if (loaded.TryGetValue(commit.EntryId, out existing))
                        {
                            if (!ElementCommitsHaveEquivalentPayload(existing, commit))
                            {
                                loaded.Remove(commit.EntryId);
                                conflictingEntryIds.Add(commit.EntryId);
                                result.InvalidRecordCount += 2;
                            }
                            continue;
                        }
                        loaded[commit.EntryId] = commit;
                        if (checkpoint.SynchronizationSucceeded)
                        {
                            result.SynchronizationSucceededEntryIds.Add(commit.EntryId);
                        }
                    }
                }
                result.TotalValidRecordCount = loaded.Count;
                result.Commits = loaded.Values
                    .OrderByDescending(delegate(FamilyBrowserElementChangeCommit commit)
                    {
                        return ParseUtc(string.IsNullOrWhiteSpace(commit.LocalSaveProtectedAtUtc) ? commit.CommittedAtUtc : commit.LocalSaveProtectedAtUtc);
                    })
                    .ThenByDescending(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId ?? string.Empty; }, StringComparer.OrdinalIgnoreCase)
                    .Take(Math.Max(1, limit))
                    .ToList();
            }
        }
        return result;
    }

    public static FamilyBrowserElementChangeHistoryLoadResult LoadPendingElementChangeCommitResult(
        string workspaceRoot,
        string projectIdentity,
        int limit)
    {
        FamilyBrowserElementChangeHistoryLoadResult result = new FamilyBrowserElementChangeHistoryLoadResult();
        List<string> comparableIdentities = string.IsNullOrWhiteSpace(projectIdentity)
            ? new List<string>()
            : BuildProjectHistoryIdentities(projectIdentity);
        string destinationIdentity = BuildManagedDestinationIdentity(workspaceRoot);
        lock (SyncRoot)
        {
            List<string> pendingPaths;
            if (!TryEnumerateSpoolFiles(BuildSpoolFolder("ElementChanges"), out pendingPaths))
            {
                result.PendingFailedCount++;
                return result;
            }
            Dictionary<string, FamilyBrowserElementChangeCommit> loaded = new Dictionary<string, FamilyBrowserElementChangeCommit>(StringComparer.OrdinalIgnoreCase);
            foreach (string path in pendingPaths)
            {
                FamilyBrowserPendingElementChangeRecord record;
                FamilyBrowserElementChangeCommit commit;
                string intendedDestination;
                string validationIssue;
                if (!TryReadPendingElementChange(path, out record, out commit, out intendedDestination) || commit == null ||
                    !ValidateElementChangeCommit(commit, string.Empty, false, out validationIssue))
                {
                    result.InvalidRecordCount++;
                    continue;
                }
                if (!CanFlushToDestination(intendedDestination, destinationIdentity, false))
                {
                    result.PendingDestinationMismatchCount++;
                    continue;
                }
                if (comparableIdentities.Count > 0 && !ElementCommitMatchesProject(commit, projectIdentity, comparableIdentities))
                {
                    continue;
                }
                FamilyBrowserElementChangeCommit existing;
                if (loaded.TryGetValue(commit.EntryId ?? string.Empty, out existing) && !ElementCommitsHaveEquivalentPayload(existing, commit))
                {
                    loaded.Remove(commit.EntryId ?? string.Empty);
                    result.InvalidRecordCount += 2;
                    continue;
                }
                loaded[commit.EntryId ?? string.Empty] = commit;
            }
            result.TotalValidRecordCount = loaded.Count;
            result.Commits = loaded.Values
                .OrderByDescending(delegate(FamilyBrowserElementChangeCommit commit) { return ParseUtc(commit.CommittedAtUtc); })
                .ThenByDescending(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId ?? string.Empty; }, StringComparer.OrdinalIgnoreCase)
                .Take(Math.Max(1, limit))
                .ToList();
        }
        return result;
    }

    public static List<FamilyBrowserTrackedProjectHistorySummary> LoadTrackedProjectHistorySummaries(string workspaceRoot)
    {
        List<FamilyBrowserElementChangeCommit> immutable = LoadAllImmutableElementChangeCommits(workspaceRoot);
        FamilyBrowserElementChangeHistoryLoadResult uploadPending = LoadPendingElementChangeCommitResult(workspaceRoot, string.Empty, int.MaxValue);
        FamilyBrowserElementSessionCheckpointHistoryLoadResult localPending = LoadElementSessionCheckpointHistory(workspaceRoot, string.Empty, int.MaxValue);
        HashSet<string> immutableIds = new HashSet<string>(immutable.Select(delegate(FamilyBrowserElementChangeCommit commit) { return commit == null ? string.Empty : commit.EntryId ?? string.Empty; }), StringComparer.OrdinalIgnoreCase);
        HashSet<string> uploadPendingIds = new HashSet<string>((uploadPending.Commits ?? new List<FamilyBrowserElementChangeCommit>()).Select(delegate(FamilyBrowserElementChangeCommit commit) { return commit == null ? string.Empty : commit.EntryId ?? string.Empty; }), StringComparer.OrdinalIgnoreCase);
        HashSet<string> localPendingIds = new HashSet<string>((localPending.Commits ?? new List<FamilyBrowserElementChangeCommit>()).Select(delegate(FamilyBrowserElementChangeCommit commit) { return commit == null ? string.Empty : commit.EntryId ?? string.Empty; }), StringComparer.OrdinalIgnoreCase);
        List<FamilyBrowserElementChangeCommit> all = immutable
            .Concat(uploadPending.Commits ?? new List<FamilyBrowserElementChangeCommit>())
            .Concat(localPending.Commits ?? new List<FamilyBrowserElementChangeCommit>())
            .Where(delegate(FamilyBrowserElementChangeCommit commit) { return commit != null; })
            .GroupBy(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId ?? string.Empty; }, StringComparer.OrdinalIgnoreCase)
            .Select(delegate(IGrouping<string, FamilyBrowserElementChangeCommit> group) { return group.First(); })
            .ToList();

        return all
            .GroupBy(delegate(FamilyBrowserElementChangeCommit commit)
            {
                string identity = commit.ProjectComparableIdentity ?? string.Empty;
                return string.IsNullOrWhiteSpace(identity)
                    ? FamilyBrowserPathIdentityService.GetStablePathIdentity(commit.ProjectIdentityPath)
                    : identity;
            }, StringComparer.OrdinalIgnoreCase)
            .Where(delegate(IGrouping<string, FamilyBrowserElementChangeCommit> group) { return !string.IsNullOrWhiteSpace(group.Key); })
            .Select(delegate(IGrouping<string, FamilyBrowserElementChangeCommit> group)
            {
                List<FamilyBrowserElementChangeCommit> projectCommits = group.ToList();
                FamilyBrowserElementChangeCommit latest = projectCommits
                    .OrderByDescending(delegate(FamilyBrowserElementChangeCommit commit) { return ParseUtc(ResolveElementCommitActivityAtUtc(commit)); })
                    .First();
                FamilyBrowserElementHistoryProjectionCounts visibleCounts = FamilyBrowserElementHistoryProjectionPolicy.CountUserFacingChanges(
                    projectCommits.SelectMany(delegate(FamilyBrowserElementChangeCommit commit)
                    {
                        return commit.Changes ?? new List<FamilyBrowserElementChangeItem>();
                    }));
                return new FamilyBrowserTrackedProjectHistorySummary
                {
                    ProjectIdentityPath = latest.ProjectIdentityPath ?? string.Empty,
                    ProjectComparableIdentity = group.Key,
                    ProjectTitle = string.IsNullOrWhiteSpace(latest.ProjectTitle) ? Path.GetFileNameWithoutExtension(latest.ProjectIdentityPath ?? string.Empty) : latest.ProjectTitle,
                    LastActivityAtUtc = ResolveElementCommitActivityAtUtc(latest),
                    ConfirmedCommitCount = projectCommits.Count(delegate(FamilyBrowserElementChangeCommit commit) { return immutableIds.Contains(commit.EntryId ?? string.Empty); }),
                    UploadPendingCommitCount = projectCommits.Count(delegate(FamilyBrowserElementChangeCommit commit) { return uploadPendingIds.Contains(commit.EntryId ?? string.Empty) && !immutableIds.Contains(commit.EntryId ?? string.Empty); }),
                    LocalSavePendingCommitCount = projectCommits.Count(delegate(FamilyBrowserElementChangeCommit commit) { return localPendingIds.Contains(commit.EntryId ?? string.Empty) && !immutableIds.Contains(commit.EntryId ?? string.Empty); }),
                    CreatedCount = visibleCounts.CreatedCount,
                    ModifiedCount = visibleCounts.ModifiedCount,
                    DeletedCount = visibleCounts.DeletedCount
                };
            })
            .OrderByDescending(delegate(FamilyBrowserTrackedProjectHistorySummary summary) { return ParseUtc(summary.LastActivityAtUtc); })
            .ThenBy(delegate(FamilyBrowserTrackedProjectHistorySummary summary) { return summary.ProjectTitle ?? string.Empty; }, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public static int GetInvalidElementSessionCheckpointCount()
    {
        int count = 0;
        lock (SyncRoot)
        {
            using (FileStream checkpointLock = TryAcquireElementSessionFileLock())
            {
                if (checkpointLock == null)
                {
                    return 1;
                }
                List<string> checkpointPaths;
                if (!TryEnumerateSpoolFiles(BuildSpoolFolder("ElementSessions"), out checkpointPaths))
                {
                    return 1;
                }
                foreach (string path in checkpointPaths)
                {
                    FamilyBrowserElementSessionCheckpoint checkpoint;
                    if (!TryReadJson(path, out checkpoint) || checkpoint == null ||
                        !ValidateElementSessionCheckpointEnvelope(checkpoint) ||
                        !ValidateElementSessionCheckpointCommits(checkpoint))
                    {
                        count++;
                    }
                }
            }
        }
        return count;
    }

    public static int GetMismatchedElementSessionCheckpointCount(string workspaceRoot)
    {
        string destinationIdentity = BuildManagedDestinationIdentity(workspaceRoot);
        int count = 0;
        lock (SyncRoot)
        {
            using (FileStream checkpointLock = TryAcquireElementSessionFileLock())
            {
                if (checkpointLock == null)
                {
                    return 1;
                }
                List<string> checkpointPaths;
                if (!TryEnumerateSpoolFiles(BuildSpoolFolder("ElementSessions"), out checkpointPaths))
                {
                    return 1;
                }
                foreach (string path in checkpointPaths)
                {
                    FamilyBrowserElementSessionCheckpoint checkpoint;
                    if (TryReadJson(path, out checkpoint) && checkpoint != null &&
                        ValidateElementSessionCheckpointEnvelope(checkpoint) &&
                        ValidateElementSessionCheckpointCommits(checkpoint) &&
                        !CanFlushToDestination(checkpoint.DestinationIdentity, destinationIdentity, false))
                    {
                        count++;
                    }
                }
            }
        }
        return count;
    }

    public static bool HasBlockingElementSessionCheckpointForManagedPolicyPath(string managedPolicyPath)
    {
        string targetDestination = string.IsNullOrWhiteSpace(managedPolicyPath)
            ? string.Empty
            : BuildStableManagedDestinationIdentity(managedPolicyPath);
        lock (SyncRoot)
        {
            using (FileStream checkpointLock = TryAcquireElementSessionFileLock())
            {
                if (checkpointLock == null)
                {
                    return true;
                }
                try
                {
                    List<string> checkpointPaths;
                    if (!TryEnumerateSpoolFiles(BuildSpoolFolder("ElementSessions"), out checkpointPaths))
                    {
                        return true;
                    }
                    foreach (string path in checkpointPaths)
                    {
                        FamilyBrowserElementSessionCheckpoint checkpoint;
                        if (!TryReadJson(path, out checkpoint) || checkpoint == null ||
                            !ValidateElementSessionCheckpointEnvelope(checkpoint) ||
                            !ValidateElementSessionCheckpointCommits(checkpoint) ||
                            string.IsNullOrWhiteSpace(checkpoint.DestinationIdentity) ||
                            string.IsNullOrWhiteSpace(targetDestination) ||
                            !string.Equals(checkpoint.DestinationIdentity, targetDestination, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                    return false;
                }
                catch
                {
                    return true;
                }
            }
        }
    }

    public static FamilyBrowserTrackingFlushResult FlushPending(string workspaceRoot)
    {
        PromoteDeferredElementSpoolDestinations(ResolveDeferredManagedDestinationPath(workspaceRoot));
        lock (SyncRoot)
        {
            FamilyBrowserTrackingFlushResult result = FlushPendingNoLock(workspaceRoot, false);
            result.FinalizedElementSessionPromotedCount = FlushFinalizedElementSessionCheckpointsNoLock(workspaceRoot, result);
            return result;
        }
    }

    public static FamilyBrowserTrackingFlushResult FlushPendingForManagedFolderTransition(string workspaceRoot, string sourceManagedRoot)
    {
        PromoteDeferredElementSpoolDestinations(Path.Combine(sourceManagedRoot ?? string.Empty, "Config", "standard-policy.json"));
        lock (SyncRoot)
        {
            string sourceDestination = BuildManagedDestinationIdentityForRoot(sourceManagedRoot);
            if (string.IsNullOrWhiteSpace(sourceDestination))
            {
                return new FamilyBrowserTrackingFlushResult
                {
                    FailedCount = 1,
                    DestinationMismatchCount = 1
                };
            }
            FamilyBrowserTrackingFlushResult result = FlushPendingNoLock(workspaceRoot, true, sourceDestination);
            result.ElementSessionCheckpointReboundCount = RebindElementSessionCheckpointsNoLock(workspaceRoot, sourceDestination, result);
            return result;
        }
    }

    public static List<FamilyBrowserOperationLogEntry> LoadImmutableOperationEntries(string workspaceRoot, int limit)
    {
        List<FamilyBrowserOperationLogEntry> loaded = new List<FamilyBrowserOperationLogEntry>();
        FlushPending(workspaceRoot);
        string dataFolder = FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "OperationLogs");
        if (string.IsNullOrWhiteSpace(dataFolder))
        {
            return loaded;
        }
        string root = Path.Combine(dataFolder, "History");
        if (!Directory.Exists(root))
        {
            return loaded;
        }
        try
        {
            IEnumerable<string> files = SafeEnumerateDirectories(root)
                .OrderByDescending(delegate(string path) { return Path.GetFileName(path); }, StringComparer.OrdinalIgnoreCase)
                .Take(90)
                .SelectMany(delegate(string path) { return SafeEnumerateFiles(path, "*.json", SearchOption.TopDirectoryOnly); });
            foreach (string path in files)
            {
                FamilyBrowserOperationLogEntry entry;
                if (TryReadJson(path, out entry) && entry != null)
                {
                    loaded.Add(entry);
                }
            }
        }
        catch
        {
        }
        return loaded
            .OrderByDescending(delegate(FamilyBrowserOperationLogEntry entry)
            {
                return ParseUtcOrMin(string.IsNullOrWhiteSpace(entry == null ? string.Empty : entry.CommittedAtUtc)
                    ? entry == null ? string.Empty : entry.RecordedAtUtc
                    : entry.CommittedAtUtc);
            })
            .ThenByDescending(delegate(FamilyBrowserOperationLogEntry entry) { return entry == null ? string.Empty : entry.EntryId ?? string.Empty; }, StringComparer.OrdinalIgnoreCase)
            .Take(Math.Max(1, limit))
            .ToList();
    }

    private static bool TryReadPendingOperation(
        string path,
        out FamilyBrowserPendingOperationRecord record,
        out FamilyBrowserOperationLogEntry entry,
        out string destinationIdentity)
    {
        record = null;
        entry = null;
        destinationIdentity = string.Empty;
        if (TryReadJson(path, out record) && record != null && record.Entry != null)
        {
            if (!ValidatePendingOperationEnvelope(record))
            {
                record = null;
                return false;
            }
            entry = record.Entry;
            destinationIdentity = record.DestinationIdentity ?? string.Empty;
            return true;
        }
        FamilyBrowserOperationLogEntry legacy;
        if (TryReadJson(path, out legacy) && legacy != null && !string.IsNullOrWhiteSpace(legacy.EntryId))
        {
            entry = legacy;
            return true;
        }
        return false;
    }

    private static bool TryReadPendingStandardCandidate(
        string path,
        out FamilyBrowserPendingStandardCandidateRecord record,
        out StandardRvtChangeCandidateEntry entry,
        out string sourceId,
        out string destinationIdentity)
    {
        record = null;
        entry = null;
        sourceId = string.Empty;
        destinationIdentity = string.Empty;
        if (TryReadJson(path, out record) && record != null && record.Entry != null && !string.IsNullOrWhiteSpace(record.SourceId))
        {
            if (!ValidatePendingStandardCandidateEnvelope(record))
            {
                record = null;
                return false;
            }
            entry = record.Entry;
            sourceId = record.SourceId;
            destinationIdentity = record.DestinationIdentity ?? string.Empty;
            return true;
        }
        FamilyBrowserPendingStandardCandidateRecord legacyRecord;
        if (TryReadJson(path, out legacyRecord) && legacyRecord != null && legacyRecord.Entry != null && !string.IsNullOrWhiteSpace(legacyRecord.SourceId))
        {
            record = legacyRecord;
            entry = legacyRecord.Entry;
            sourceId = legacyRecord.SourceId;
            return true;
        }
        return false;
    }

    private static bool TryReadPendingElementChange(
        string path,
        out FamilyBrowserPendingElementChangeRecord record,
        out FamilyBrowserElementChangeCommit commit,
        out string destinationIdentity)
    {
        record = null;
        commit = null;
        destinationIdentity = string.Empty;
        if (TryReadJson(path, out record) && record != null && record.Entry != null)
        {
            if (!ValidatePendingElementEnvelope(record))
            {
                record = null;
                return false;
            }
            commit = record.Entry;
            destinationIdentity = record.DestinationIdentity ?? string.Empty;
            return true;
        }
        FamilyBrowserElementChangeCommit legacy;
        if (TryReadJson(path, out legacy) && legacy != null && !string.IsNullOrWhiteSpace(legacy.EntryId))
        {
            commit = legacy;
            return true;
        }
        return false;
    }

    private static string BuildManagedDestinationIdentity(string workspaceRoot)
    {
        try
        {
            string policyPath = FamilyBrowserStandardPolicyStore.GetConfiguredManagedPolicyPath();
            if (!string.IsNullOrWhiteSpace(policyPath))
            {
                return BuildStableManagedDestinationIdentity(policyPath);
            }
        }
        catch
        {
        }
        try
        {
            return BuildStableManagedDestinationIdentity(FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, string.Empty));
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string ResolveDeferredManagedDestinationPath(string workspaceRoot)
    {
        try
        {
            string policyPath = FamilyBrowserStandardPolicyStore.GetConfiguredManagedPolicyPath();
            if (!string.IsNullOrWhiteSpace(policyPath))
            {
                return policyPath;
            }
        }
        catch
        {
        }
        try
        {
            return FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, string.Empty);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string BuildDeferredManagedDestinationIdentity(string path)
    {
        string normalized = FamilyBrowserPathIdentityService.NormalizePath(path ?? string.Empty);
        return string.IsNullOrWhiteSpace(normalized)
            ? string.Empty
            : "MANAGED-DEFERRED-PATH:" + normalized.ToUpperInvariant();
    }

    private static void PromoteDeferredElementSpoolDestinations(string deferredDestinationPath)
    {
        string deferredIdentity = BuildDeferredManagedDestinationIdentity(deferredDestinationPath);
        if (string.IsNullOrWhiteSpace(deferredIdentity))
        {
            return;
        }
        string stableIdentity = BuildStableManagedDestinationIdentity(deferredDestinationPath);
        if (string.IsNullOrWhiteSpace(stableIdentity) ||
            string.Equals(deferredIdentity, stableIdentity, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        lock (DeferredSpoolRoot)
        {
            List<string> spoolPaths;
            if (!TryEnumerateSpoolFiles(BuildSpoolFolder("ElementChanges"), out spoolPaths))
            {
                return;
            }
            foreach (string spoolPath in spoolPaths)
            {
                FamilyBrowserPendingElementChangeRecord record;
                if (!TryReadJson(spoolPath, out record) ||
                    record == null ||
                    !ValidatePendingElementEnvelope(record) ||
                    !string.Equals(record.DestinationIdentity, deferredIdentity, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                record.DestinationIdentity = stableIdentity;
                EnsurePendingElementEnvelopeIntegrity(record);
                TryWriteMutableJsonAtomic(spoolPath, record);
            }
        }
    }

    private static string BuildManagedDestinationIdentityForRoot(string managedRoot)
    {
        if (string.IsNullOrWhiteSpace(managedRoot))
        {
            return string.Empty;
        }
        return BuildStableManagedDestinationIdentity(Path.Combine(managedRoot, "Config", "standard-policy.json"));
    }

    private static string BuildStableManagedDestinationIdentity(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }
        string canonical = FamilyBrowserPathIdentityService.GetCanonicalPath(path);
        if (string.IsNullOrWhiteSpace(canonical))
        {
            canonical = FamilyBrowserPathIdentityService.NormalizePath(path);
        }
        return string.IsNullOrWhiteSpace(canonical) ? string.Empty : "MANAGED-PATH:" + canonical.ToUpperInvariant();
    }

    private static FamilyBrowserElementSessionCheckpoint BuildElementSessionCheckpoint(
        string workspaceRoot,
        string projectIdentity,
        string localDocumentPath,
        string revitUserName,
        List<FamilyBrowserElementChangeCommit> commits,
        bool synchronizationSucceeded)
    {
        FamilyBrowserElementSessionCheckpoint checkpoint = new FamilyBrowserElementSessionCheckpoint
        {
            CheckpointId = BuildElementSessionCheckpointId(projectIdentity, localDocumentPath, revitUserName),
            DestinationIdentity = BuildManagedDestinationIdentity(workspaceRoot),
            ProjectStableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(projectIdentity),
            LocalDocumentStableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(localDocumentPath),
            RevitUserName = (revitUserName ?? string.Empty).Trim(),
            UpdatedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
            SynchronizationSucceeded = synchronizationSucceeded,
            Commits = commits ?? new List<FamilyBrowserElementChangeCommit>()
        };
        checkpoint.EnvelopeIntegrityVersion = 1;
        checkpoint.EnvelopeIntegritySha256 = ComputeElementSessionCheckpointIntegrity(checkpoint);
        return checkpoint;
    }

    private static string BuildElementSessionCheckpointId(string projectIdentity, string localDocumentPath, string revitUserName)
    {
        string project = FamilyBrowserPathIdentityService.GetStablePathIdentity(projectIdentity);
        string local = FamilyBrowserPathIdentityService.GetStablePathIdentity(localDocumentPath);
        if (string.IsNullOrWhiteSpace(project))
        {
            return string.Empty;
        }
        if (string.IsNullOrWhiteSpace(local))
        {
            local = project;
        }
        return HashText(project + "|" + local + "|" + (revitUserName ?? string.Empty).Trim().ToUpperInvariant()).Substring(0, 40);
    }

    private static string BuildElementSessionCheckpointPath(string checkpointId)
    {
        return Path.Combine(BuildSpoolFolder("ElementSessions"), SafeFileName(checkpointId) + ".json");
    }

    private static bool ValidateElementSessionCheckpoint(
        FamilyBrowserElementSessionCheckpoint checkpoint,
        string projectIdentity,
        string localDocumentPath,
        string revitUserName)
    {
        if (!ValidateElementSessionCheckpointEnvelope(checkpoint))
        {
            return false;
        }
        string expectedId = BuildElementSessionCheckpointId(projectIdentity, localDocumentPath, revitUserName);
        return !string.IsNullOrWhiteSpace(expectedId) &&
            string.Equals(checkpoint.CheckpointId, expectedId, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(checkpoint.ProjectStableIdentity, FamilyBrowserPathIdentityService.GetStablePathIdentity(projectIdentity), StringComparison.OrdinalIgnoreCase) &&
            string.Equals(checkpoint.LocalDocumentStableIdentity, FamilyBrowserPathIdentityService.GetStablePathIdentity(localDocumentPath), StringComparison.OrdinalIgnoreCase) &&
            string.Equals(checkpoint.RevitUserName ?? string.Empty, (revitUserName ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private static bool ValidateElementSessionCheckpointEnvelope(FamilyBrowserElementSessionCheckpoint checkpoint)
    {
        return checkpoint != null &&
            checkpoint.SchemaVersion == 1 &&
            checkpoint.EnvelopeIntegrityVersion == 1 &&
            !string.IsNullOrWhiteSpace(checkpoint.CheckpointId) &&
            !string.IsNullOrWhiteSpace(checkpoint.DestinationIdentity) &&
            !string.IsNullOrWhiteSpace(checkpoint.ProjectStableIdentity) &&
            !string.IsNullOrWhiteSpace(checkpoint.EnvelopeIntegritySha256) &&
            FixedTimeEquals(checkpoint.EnvelopeIntegritySha256, ComputeElementSessionCheckpointIntegrity(checkpoint));
    }

    private static bool ValidateElementSessionCheckpointCommits(FamilyBrowserElementSessionCheckpoint checkpoint)
    {
        string expectedProjectIdentity = checkpoint == null ? string.Empty : checkpoint.ProjectStableIdentity;
        List<FamilyBrowserElementChangeCommit> commits = checkpoint == null || checkpoint.Commits == null
            ? new List<FamilyBrowserElementChangeCommit>()
            : checkpoint.Commits;
        if (commits.Count == 0)
        {
            return false;
        }
        HashSet<string> entryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (FamilyBrowserElementChangeCommit commit in commits)
        {
            string issue;
            if (commit == null || string.IsNullOrWhiteSpace(commit.EntryId) || !entryIds.Add(commit.EntryId) ||
                string.IsNullOrWhiteSpace(commit.IntegritySha256) ||
                !ElementCommitMatchesCheckpointProject(commit, expectedProjectIdentity) ||
                !ValidateElementChangeCommit(commit, string.Empty, false, out issue))
            {
                return false;
            }
        }
        return true;
    }

    private static bool ElementCommitMatchesCheckpointProject(FamilyBrowserElementChangeCommit commit, string expectedStableIdentity)
    {
        if (commit == null || string.IsNullOrWhiteSpace(expectedStableIdentity))
        {
            return false;
        }
        return string.Equals(commit.ProjectComparableIdentity ?? string.Empty, expectedStableIdentity, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(FamilyBrowserPathIdentityService.GetStablePathIdentity(commit.ProjectIdentityPath), expectedStableIdentity, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(FamilyBrowserPathIdentityService.GetStablePathIdentity(commit.ProjectCanonicalPath), expectedStableIdentity, StringComparison.OrdinalIgnoreCase);
    }

    private static string ComputeElementSessionCheckpointIntegrity(FamilyBrowserElementSessionCheckpoint checkpoint)
    {
        StringBuilder canonical = new StringBuilder(1024);
        AppendCanonical(canonical, checkpoint == null ? 0 : checkpoint.SchemaVersion);
        AppendCanonical(canonical, checkpoint == null ? string.Empty : checkpoint.CheckpointId);
        AppendCanonical(canonical, checkpoint == null ? string.Empty : checkpoint.DestinationIdentity);
        AppendCanonical(canonical, checkpoint == null ? string.Empty : checkpoint.ProjectStableIdentity);
        AppendCanonical(canonical, checkpoint == null ? string.Empty : checkpoint.LocalDocumentStableIdentity);
        AppendCanonical(canonical, checkpoint == null ? string.Empty : checkpoint.RevitUserName);
        AppendCanonical(canonical, checkpoint == null ? string.Empty : checkpoint.UpdatedAtUtc);
        AppendCanonical(canonical, checkpoint != null && checkpoint.SynchronizationSucceeded);
        List<FamilyBrowserElementChangeCommit> commits = checkpoint == null || checkpoint.Commits == null
            ? new List<FamilyBrowserElementChangeCommit>()
            : checkpoint.Commits;
        AppendCanonical(canonical, commits.Count);
        foreach (FamilyBrowserElementChangeCommit commit in commits)
        {
            AppendCanonical(canonical, commit == null ? string.Empty : commit.EntryId);
            AppendCanonical(canonical, commit == null ? 0 : commit.IntegrityVersion);
            AppendCanonical(canonical, commit == null ? string.Empty : commit.IntegritySha256);
        }
        using (SHA256 sha = SHA256.Create())
        {
            return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(canonical.ToString()))).Replace("-", string.Empty);
        }
    }

    private static string ComputeElementSessionCheckpointRevisionToken(FamilyBrowserElementSessionCheckpoint checkpoint)
    {
        StringBuilder canonical = new StringBuilder(1024);
        AppendCanonical(canonical, checkpoint == null ? 0 : checkpoint.SchemaVersion);
        AppendCanonical(canonical, checkpoint == null ? string.Empty : checkpoint.CheckpointId);
        AppendCanonical(canonical, checkpoint == null ? string.Empty : checkpoint.ProjectStableIdentity);
        AppendCanonical(canonical, checkpoint == null ? string.Empty : checkpoint.LocalDocumentStableIdentity);
        AppendCanonical(canonical, checkpoint == null ? string.Empty : checkpoint.RevitUserName);
        AppendCanonical(canonical, checkpoint != null && checkpoint.SynchronizationSucceeded);
        List<FamilyBrowserElementChangeCommit> commits = checkpoint == null || checkpoint.Commits == null
            ? new List<FamilyBrowserElementChangeCommit>()
            : checkpoint.Commits;
        AppendCanonical(canonical, commits.Count);
        foreach (FamilyBrowserElementChangeCommit commit in commits)
        {
            AppendCanonical(canonical, commit == null ? string.Empty : commit.EntryId);
            AppendCanonical(canonical, commit == null ? 0 : commit.IntegrityVersion);
            AppendCanonical(canonical, commit == null ? string.Empty : commit.IntegritySha256);
        }
        using (SHA256 sha = SHA256.Create())
        {
            return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(canonical.ToString()))).Replace("-", string.Empty);
        }
    }

    private static int RebindElementSessionCheckpointsNoLock(
        string workspaceRoot,
        string sourceDestination,
        FamilyBrowserTrackingFlushResult result)
    {
        if (string.IsNullOrWhiteSpace(sourceDestination))
        {
            return 0;
        }
        using (FileStream checkpointLock = TryAcquireElementSessionFileLock())
        {
            if (checkpointLock == null)
            {
                if (result != null)
                {
                    result.ElementSessionCheckpointLockUnavailable = true;
                    result.ElementSessionCheckpointRebindFailedCount++;
                    result.FailedCount++;
                }
                return 0;
            }
            string destination = BuildManagedDestinationIdentity(workspaceRoot);
            if (string.IsNullOrWhiteSpace(destination))
            {
                if (result != null)
                {
                    result.ElementSessionCheckpointRebindFailedCount++;
                    result.FailedCount++;
                }
                return 0;
            }
            int rebound = 0;
            List<string> checkpointPaths;
            if (!TryEnumerateSpoolFiles(BuildSpoolFolder("ElementSessions"), out checkpointPaths))
            {
                if (result != null)
                {
                    result.ElementSessionCheckpointLockUnavailable = true;
                    result.ElementSessionCheckpointRebindFailedCount++;
                    result.FailedCount++;
                }
                return 0;
            }
            foreach (string path in checkpointPaths)
            {
                FamilyBrowserElementSessionCheckpoint checkpoint;
                if (!TryReadJson(path, out checkpoint) || checkpoint == null ||
                    !ValidateElementSessionCheckpointEnvelope(checkpoint) ||
                    !ValidateElementSessionCheckpointCommits(checkpoint) ||
                    !string.Equals(checkpoint.DestinationIdentity, sourceDestination, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                checkpoint.DestinationIdentity = destination;
                checkpoint.UpdatedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
                checkpoint.EnvelopeIntegritySha256 = ComputeElementSessionCheckpointIntegrity(checkpoint);
                if (TryWriteMutableJsonAtomic(path, checkpoint))
                {
                    rebound++;
                }
                else if (result != null)
                {
                    result.ElementSessionCheckpointRebindFailedCount++;
                    result.FailedCount++;
                }
            }
            return rebound;
        }
    }

    private static bool CanFlushToDestination(string intendedDestination, string currentDestination, bool allowDestinationRebind, string rebindSourceDestination = "")
    {
        if (string.IsNullOrWhiteSpace(intendedDestination) ||
            string.Equals(intendedDestination, currentDestination ?? string.Empty, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
        return allowDestinationRebind &&
            !string.IsNullOrWhiteSpace(rebindSourceDestination) &&
            string.Equals(intendedDestination, rebindSourceDestination, StringComparison.OrdinalIgnoreCase);
    }

    private static void EnsurePendingOperationEnvelopeIntegrity(FamilyBrowserPendingOperationRecord record)
    {
        if (record == null)
        {
            return;
        }
        record.EnvelopeIntegrityVersion = 1;
        record.EnvelopeIntegritySha256 = ComputePendingOperationEnvelopeIntegrity(record);
    }

    private static bool ValidatePendingOperationEnvelope(FamilyBrowserPendingOperationRecord record)
    {
        if (record == null || record.Entry == null || string.IsNullOrWhiteSpace(record.DestinationIdentity))
        {
            return false;
        }
        if (record.EnvelopeIntegrityVersion == 0 && string.IsNullOrWhiteSpace(record.EnvelopeIntegritySha256))
        {
            return true;
        }
        return record.EnvelopeIntegrityVersion == 1 &&
            !string.IsNullOrWhiteSpace(record.EnvelopeIntegritySha256) &&
            FixedTimeEquals(record.EnvelopeIntegritySha256, ComputePendingOperationEnvelopeIntegrity(record));
    }

    private static string ComputePendingOperationEnvelopeIntegrity(FamilyBrowserPendingOperationRecord record)
    {
        FamilyBrowserPendingOperationRecord payload = new FamilyBrowserPendingOperationRecord
        {
            DestinationIdentity = record == null ? string.Empty : record.DestinationIdentity,
            EnvelopeIntegrityVersion = record == null ? 0 : record.EnvelopeIntegrityVersion,
            EnvelopeIntegritySha256 = string.Empty,
            Entry = record == null ? null : record.Entry
        };
        using (SHA256 sha = SHA256.Create())
        {
            return BitConverter.ToString(sha.ComputeHash(SerializeJson(payload))).Replace("-", string.Empty);
        }
    }

    private static void EnsurePendingStandardCandidateEnvelopeIntegrity(FamilyBrowserPendingStandardCandidateRecord record)
    {
        if (record == null)
        {
            return;
        }
        record.EnvelopeIntegrityVersion = 1;
        record.EnvelopeIntegritySha256 = ComputePendingStandardCandidateEnvelopeIntegrity(record);
    }

    private static bool ValidatePendingStandardCandidateEnvelope(FamilyBrowserPendingStandardCandidateRecord record)
    {
        if (record == null || record.Entry == null || string.IsNullOrWhiteSpace(record.DestinationIdentity) || string.IsNullOrWhiteSpace(record.SourceId))
        {
            return false;
        }
        if (record.EnvelopeIntegrityVersion == 0 && string.IsNullOrWhiteSpace(record.EnvelopeIntegritySha256))
        {
            return true;
        }
        return record.EnvelopeIntegrityVersion == 1 &&
            !string.IsNullOrWhiteSpace(record.EnvelopeIntegritySha256) &&
            FixedTimeEquals(record.EnvelopeIntegritySha256, ComputePendingStandardCandidateEnvelopeIntegrity(record));
    }

    private static string ComputePendingStandardCandidateEnvelopeIntegrity(FamilyBrowserPendingStandardCandidateRecord record)
    {
        FamilyBrowserPendingStandardCandidateRecord payload = new FamilyBrowserPendingStandardCandidateRecord
        {
            DestinationIdentity = record == null ? string.Empty : record.DestinationIdentity,
            SourceId = record == null ? string.Empty : record.SourceId,
            EnvelopeIntegrityVersion = record == null ? 0 : record.EnvelopeIntegrityVersion,
            EnvelopeIntegritySha256 = string.Empty,
            Entry = record == null ? null : record.Entry
        };
        using (SHA256 sha = SHA256.Create())
        {
            return BitConverter.ToString(sha.ComputeHash(SerializeJson(payload))).Replace("-", string.Empty);
        }
    }

    private static void EnsurePendingElementEnvelopeIntegrity(FamilyBrowserPendingElementChangeRecord record)
    {
        if (record == null)
        {
            return;
        }
        record.EnvelopeIntegrityVersion = 2;
        record.EnvelopeIntegritySha256 = ComputePendingElementEnvelopeIntegrity(record);
    }

    private static bool ValidatePendingElementEnvelope(FamilyBrowserPendingElementChangeRecord record)
    {
        if (record == null || record.Entry == null || string.IsNullOrWhiteSpace(record.DestinationIdentity))
        {
            return false;
        }
        if (string.IsNullOrWhiteSpace(record.EnvelopeIntegritySha256))
        {
            return false;
        }
        return (record.EnvelopeIntegrityVersion == 1 || record.EnvelopeIntegrityVersion == 2) &&
            FixedTimeEquals(record.EnvelopeIntegritySha256, ComputePendingElementEnvelopeIntegrity(record));
    }

    private static string ComputePendingElementEnvelopeIntegrity(FamilyBrowserPendingElementChangeRecord record)
    {
        if (record == null)
        {
            return string.Empty;
        }
        if (record.EnvelopeIntegrityVersion == 1)
        {
            PendingElementEnvelopeIntegrityV1Payload payload = new PendingElementEnvelopeIntegrityV1Payload
            {
                DestinationIdentity = record.DestinationIdentity,
                EnvelopeIntegrityVersion = 1,
                EnvelopeIntegritySha256 = string.Empty,
                Entry = CreatePendingElementEnvelopeIntegrityV1EntryPayload(record.Entry)
            };
            using (SHA256 sha = SHA256.Create())
            {
                return BitConverter.ToString(sha.ComputeHash(SerializeJson(payload))).Replace("-", string.Empty);
            }
        }
        StringBuilder canonical = new StringBuilder(512);
        AppendCanonical(canonical, 2);
        AppendCanonical(canonical, record.DestinationIdentity);
        AppendCanonical(canonical, record.Entry == null ? string.Empty : record.Entry.EntryId);
        AppendCanonical(canonical, record.Entry == null ? string.Empty : record.Entry.ProjectComparableIdentity);
        AppendCanonical(canonical, record.Entry == null ? 0 : record.Entry.IntegrityVersion);
        AppendCanonical(canonical, record.Entry == null ? string.Empty : record.Entry.IntegritySha256);
        using (SHA256 sha = SHA256.Create())
        {
            return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(canonical.ToString()))).Replace("-", string.Empty);
        }
    }

    private static bool HasElementChangesOrCoverageGap(FamilyBrowserElementChangeCommit commit)
    {
        return commit != null &&
            ((commit.Changes != null && commit.Changes.Count > 0) ||
             (commit.SchemaVersion >= 5 && commit.CoverageGapOnly));
    }

    private static bool ValidateElementChangeCommit(
        FamilyBrowserElementChangeCommit commit,
        string expectedComparableIdentity,
        bool upgradeLegacyIntegrity,
        out string issue)
    {
        issue = string.Empty;
        if (commit == null || string.IsNullOrWhiteSpace(commit.EntryId))
        {
            issue = "Missing entry identity.";
            return false;
        }
        if (commit.SchemaVersion < MinimumSupportedElementCommitSchemaVersion || commit.SchemaVersion > MaximumSupportedElementCommitSchemaVersion)
        {
            issue = "Unsupported element history schema version.";
            return false;
        }
        if (string.IsNullOrWhiteSpace(commit.ProjectComparableIdentity) ||
            (!string.IsNullOrWhiteSpace(expectedComparableIdentity) &&
             !string.Equals(commit.ProjectComparableIdentity, expectedComparableIdentity, StringComparison.OrdinalIgnoreCase)))
        {
            issue = "Project identity mismatch.";
            return false;
        }
        DateTime committed;
        if (!DateTime.TryParse(commit.CommittedAtUtc ?? string.Empty, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out committed))
        {
            issue = "Invalid commit time.";
            return false;
        }
        bool coverageGapOnly = commit.SchemaVersion >= 5 && commit.CoverageGapOnly;
        bool externalRebaseGap = string.Equals(commit.AttributionConfidence ?? string.Empty, "ClientObservedWithExternalRebaseGap", StringComparison.OrdinalIgnoreCase) ||
            (commit.CoverageNote ?? string.Empty).IndexOf("incoming central/reload update could not be fully rebased", StringComparison.OrdinalIgnoreCase) >= 0;
        bool hasCoverageGapEvidence = commit.BaselineCapturedLate ||
            commit.EventReadFailureCount > 0 ||
            commit.CommitBoundaryReadFailureCount > 0 ||
            commit.UnmatchedUndoCount > 0 ||
            commit.UnmatchedRedoCount > 0 ||
            externalRebaseGap;
        if (commit.Changes == null ||
            (commit.Changes.Count == 0 && (!coverageGapOnly || !hasCoverageGapEvidence)) ||
            (commit.Changes.Count > 0 && coverageGapOnly) ||
            commit.Changes.Any(delegate(FamilyBrowserElementChangeItem item)
        {
            return item == null || string.IsNullOrWhiteSpace(item.ElementId) ||
                !(string.Equals(item.ChangeKind, "Created", StringComparison.OrdinalIgnoreCase) ||
                  string.Equals(item.ChangeKind, "Modified", StringComparison.OrdinalIgnoreCase) ||
                  string.Equals(item.ChangeKind, "Deleted", StringComparison.OrdinalIgnoreCase) ||
                  string.Equals(item.ChangeKind, "CreatedThenDeleted", StringComparison.OrdinalIgnoreCase));
        }))
        {
            issue = "Invalid element change payload.";
            return false;
        }
        if (commit.SchemaVersion >= 5)
        {
            if (upgradeLegacyIntegrity && string.IsNullOrWhiteSpace(commit.IntegritySha256))
            {
                NormalizeDerivedElementChangeCounts(commit);
            }
            if (!ValidateElementChangeSemanticsV5(commit, out issue))
            {
                return false;
            }
            if (commit.SchemaVersion >= 6 && !ValidateElementChangeSemanticsV6(commit, out issue))
            {
                return false;
            }
        }
        if (string.IsNullOrWhiteSpace(commit.IntegritySha256))
        {
            if (upgradeLegacyIntegrity)
            {
                if (!EnsureElementChangeIntegrity(commit))
                {
                    issue = "Element history checksum could not be created.";
                    return false;
                }
                return true;
            }
            if (commit.SchemaVersion >= 6)
            {
                issue = "Schema-v6 element history requires integrity-v5 evidence.";
                return false;
            }
            if (commit.SchemaVersion >= 5)
            {
                issue = "Schema-v5 element history requires integrity-v4 evidence.";
                return false;
            }
            return true;
        }
        if (commit.IntegrityVersion != 1 && commit.IntegrityVersion != 2 && commit.IntegrityVersion != 3 && commit.IntegrityVersion != 4 && commit.IntegrityVersion != 5)
        {
            issue = "Unsupported element history integrity version.";
            return false;
        }
        if (commit.SchemaVersion >= 6 && commit.IntegrityVersion != 5)
        {
            issue = "Schema-v6 element history requires integrity-v5 evidence.";
            return false;
        }
        if (commit.SchemaVersion == 5 && commit.IntegrityVersion != 4)
        {
            issue = "Schema-v5 element history requires integrity-v4 evidence.";
            return false;
        }
        string expectedHash = ComputeElementChangeIntegrity(commit);
        if (!FixedTimeEquals(commit.IntegritySha256, expectedHash))
        {
            issue = "Element history checksum mismatch.";
            return false;
        }
        return true;
    }

    private static void NormalizeDerivedElementChangeCounts(FamilyBrowserElementChangeCommit commit)
    {
        if (commit == null)
        {
            return;
        }
        List<FamilyBrowserElementChangeItem> changes = commit.Changes ?? new List<FamilyBrowserElementChangeItem>();
        commit.CreatedCount = changes.Count(delegate(FamilyBrowserElementChangeItem item)
        {
            return item != null && (string.Equals(item.ChangeKind, "Created", StringComparison.OrdinalIgnoreCase) || string.Equals(item.ChangeKind, "CreatedThenDeleted", StringComparison.OrdinalIgnoreCase));
        });
        commit.ModifiedCount = changes.Count(delegate(FamilyBrowserElementChangeItem item)
        {
            return item != null && string.Equals(item.ChangeKind, "Modified", StringComparison.OrdinalIgnoreCase);
        });
        commit.DeletedCount = changes.Count(delegate(FamilyBrowserElementChangeItem item)
        {
            return item != null && (string.Equals(item.ChangeKind, "Deleted", StringComparison.OrdinalIgnoreCase) || string.Equals(item.ChangeKind, "CreatedThenDeleted", StringComparison.OrdinalIgnoreCase));
        });
        commit.TransientCreatedDeletedCount = changes.Count(delegate(FamilyBrowserElementChangeItem item)
        {
            return item != null && string.Equals(item.ChangeKind, "CreatedThenDeleted", StringComparison.OrdinalIgnoreCase);
        });
        commit.ExternalUpdateOverlapCount = changes.Count(delegate(FamilyBrowserElementChangeItem item)
        {
            return item != null && item.ExternalUpdateOverlap;
        });
    }

    private static bool ValidateElementChangeSemanticsV5(FamilyBrowserElementChangeCommit commit, out string issue)
    {
        issue = string.Empty;
        if (commit == null || commit.TransactionNames == null || commit.Changes == null)
        {
            issue = "Schema-v5 element history requires complete collection fields.";
            return false;
        }
        if (commit.BaselineElapsedMilliseconds < 0L || commit.BaselineElementCount < 0 || commit.ActivityCount < 0 ||
            commit.UndoCount < 0 || commit.RedoCount < 0 || commit.UnmatchedUndoCount < 0 || commit.UnmatchedRedoCount < 0 ||
            commit.CreatedCount < 0 || commit.ModifiedCount < 0 || commit.DeletedCount < 0 || commit.TransientCreatedDeletedCount < 0 ||
            commit.ExternalUpdateOverlapCount < 0 || commit.EventReadFailureCount < 0 || commit.CommitBoundaryReadFailureCount < 0 ||
            commit.UnmatchedUndoCount > commit.UndoCount || commit.UnmatchedRedoCount > commit.RedoCount)
        {
            issue = "Schema-v5 element history contains impossible negative or unmatched counters.";
            return false;
        }

        string identityPath = FamilyBrowserPathIdentityService.GetStablePathIdentity(commit.ProjectIdentityPath);
        string canonicalPath = FamilyBrowserPathIdentityService.GetStablePathIdentity(commit.ProjectCanonicalPath);
        if (string.IsNullOrWhiteSpace(commit.ProjectComparableIdentity) ||
            ((string.IsNullOrWhiteSpace(identityPath) || !string.Equals(commit.ProjectComparableIdentity, identityPath, StringComparison.OrdinalIgnoreCase)) &&
             (string.IsNullOrWhiteSpace(canonicalPath) || !string.Equals(commit.ProjectComparableIdentity, canonicalPath, StringComparison.OrdinalIgnoreCase))))
        {
            issue = "Schema-v5 element history project paths contradict its stable project identity.";
            return false;
        }

        DateTime committedAt;
        DateTime trackingStartedAt;
        DateTime baselineCapturedAt;
        bool hasCommittedAt = TryParseOptionalUtc(commit.CommittedAtUtc, out committedAt);
        bool hasTrackingStartedAt = TryParseOptionalUtc(commit.TrackingStartedAtUtc, out trackingStartedAt);
        bool hasBaselineCapturedAt = TryParseOptionalUtc(commit.BaselineCapturedAtUtc, out baselineCapturedAt);
        if (!hasCommittedAt ||
            (!string.IsNullOrWhiteSpace(commit.TrackingStartedAtUtc) && !hasTrackingStartedAt) ||
            (!string.IsNullOrWhiteSpace(commit.BaselineCapturedAtUtc) && !hasBaselineCapturedAt) ||
            (hasTrackingStartedAt && hasBaselineCapturedAt && baselineCapturedAt < trackingStartedAt) ||
            (hasBaselineCapturedAt && committedAt < baselineCapturedAt) ||
            (hasTrackingStartedAt && committedAt < trackingStartedAt))
        {
            issue = "Schema-v5 element history contains an invalid tracking, baseline, or commit time range.";
            return false;
        }

        int created = 0;
        int modified = 0;
        int deleted = 0;
        int transient = 0;
        int overlap = 0;
        foreach (FamilyBrowserElementChangeItem item in commit.Changes)
        {
            if (item == null || item.TransactionNames == null)
            {
                issue = "Schema-v5 element history contains an incomplete change row.";
                return false;
            }
            string kind = item.ChangeKind ?? string.Empty;
            bool isCreated = string.Equals(kind, "Created", StringComparison.OrdinalIgnoreCase);
            bool isModified = string.Equals(kind, "Modified", StringComparison.OrdinalIgnoreCase);
            bool isDeleted = string.Equals(kind, "Deleted", StringComparison.OrdinalIgnoreCase);
            bool isTransient = string.Equals(kind, "CreatedThenDeleted", StringComparison.OrdinalIgnoreCase);
            if ((isCreated && (item.Before != null || item.After == null || item.PreviousStateUnavailable)) ||
                (isModified && (item.After == null || (!item.PreviousStateUnavailable && item.Before == null))) ||
                (isDeleted && (item.After != null || (!item.PreviousStateUnavailable && item.Before == null))) ||
                (isTransient && (item.Before != null || item.After != null || item.PreviousStateUnavailable)))
            {
                issue = "Schema-v5 element history change kind contradicts its before/after states.";
                return false;
            }
            if (!StateMatchesChangeIdentity(item.Before, item.ElementId) || !StateMatchesChangeIdentity(item.After, item.ElementId))
            {
                issue = "Schema-v5 element history row and state element identities do not match.";
                return false;
            }
            FamilyBrowserTrackedElementState display = item.After ?? item.Before;
            if (display != null && !string.IsNullOrWhiteSpace(display.UniqueId) &&
                !string.Equals(display.UniqueId, item.UniqueId ?? string.Empty, StringComparison.Ordinal))
            {
                issue = "Schema-v5 element history row and state unique identities do not match.";
                return false;
            }
            if (item.Before != null && item.After != null &&
                !string.IsNullOrWhiteSpace(item.Before.UniqueId) &&
                !string.IsNullOrWhiteSpace(item.After.UniqueId) &&
                !string.Equals(item.Before.UniqueId, item.After.UniqueId, StringComparison.Ordinal))
            {
                issue = "Schema-v5 element history before/after states belong to different unique elements.";
                return false;
            }
            DateTime firstObserved;
            DateTime lastObserved;
            bool hasFirstObserved = TryParseOptionalUtc(item.FirstObservedAtUtc, out firstObserved);
            bool hasLastObserved = TryParseOptionalUtc(item.LastObservedAtUtc, out lastObserved);
            if ((!string.IsNullOrWhiteSpace(item.FirstObservedAtUtc) && !hasFirstObserved) ||
                (!string.IsNullOrWhiteSpace(item.LastObservedAtUtc) && !hasLastObserved) ||
                (hasFirstObserved && hasLastObserved && lastObserved < firstObserved))
            {
                issue = "Schema-v5 element history contains an invalid observation time range.";
                return false;
            }
            if ((hasFirstObserved && hasTrackingStartedAt && firstObserved < trackingStartedAt) ||
                (hasFirstObserved && firstObserved > committedAt) ||
                (hasLastObserved && lastObserved > committedAt))
            {
                issue = "Schema-v5 element history observation times fall outside its tracking and commit boundary.";
                return false;
            }
            if (isCreated || isTransient)
            {
                created++;
            }
            if (isModified)
            {
                modified++;
            }
            if (isDeleted || isTransient)
            {
                deleted++;
            }
            if (isTransient)
            {
                transient++;
            }
            if (item.ExternalUpdateOverlap)
            {
                overlap++;
            }
        }
        if (commit.CreatedCount != created || commit.ModifiedCount != modified || commit.DeletedCount != deleted ||
            commit.TransientCreatedDeletedCount != transient || commit.ExternalUpdateOverlapCount != overlap)
        {
            issue = "Schema-v5 element history counters do not match its change rows.";
            return false;
        }
        if (commit.CoverageGapOnly && (created != 0 || modified != 0 || deleted != 0 || overlap != 0))
        {
            issue = "Schema-v5 coverage-gap-only history contains element rows or derived counts.";
            return false;
        }
        DateTime protectedAt;
        DateTime publishedAt;
        bool hasProtectedAt = TryParseOptionalUtc(commit.LocalSaveProtectedAtUtc, out protectedAt);
        bool hasPublishedAt = TryParseOptionalUtc(commit.PublishedAtUtc, out publishedAt);
        if ((!string.IsNullOrWhiteSpace(commit.LocalSaveProtectedAtUtc) && !hasProtectedAt) ||
            (!string.IsNullOrWhiteSpace(commit.PublishedAtUtc) && !hasPublishedAt) ||
            (hasProtectedAt && hasPublishedAt && publishedAt < protectedAt) ||
            (hasProtectedAt && protectedAt > committedAt) ||
            (hasPublishedAt && publishedAt > committedAt))
        {
            issue = "Schema-v5 element history contains an invalid protection/publication time range.";
            return false;
        }
        return true;
    }

    private static bool ValidateElementChangeSemanticsV6(FamilyBrowserElementChangeCommit commit, out string issue)
    {
        issue = string.Empty;
        foreach (FamilyBrowserElementChangeItem item in commit == null || commit.Changes == null
            ? new List<FamilyBrowserElementChangeItem>()
            : commit.Changes)
        {
            if (item == null || !ValidateTrackedElementStateV6(item.Before) || !ValidateTrackedElementStateV6(item.After))
            {
                issue = "Schema-v6 element history contains an invalid extended element state.";
                return false;
            }
            FamilyBrowserTrackedElementState display = item.After ?? item.Before;
            if (display != null &&
                (!string.Equals(item.ElementName ?? string.Empty, display.ElementName ?? string.Empty, StringComparison.Ordinal) ||
                 !string.Equals(item.TrackingKind ?? string.Empty, display.TrackingKind ?? string.Empty, StringComparison.Ordinal)))
            {
                issue = "Schema-v6 element history row and extended state metadata do not match.";
                return false;
            }
        }
        return true;
    }

    private static bool ValidateTrackedElementStateV6(FamilyBrowserTrackedElementState state)
    {
        if (state == null)
        {
            return true;
        }
        string kind = state.TrackingKind ?? string.Empty;
        return !string.IsNullOrWhiteSpace(state.StateSignature) &&
            (string.Equals(kind, "Element", StringComparison.Ordinal) ||
             string.Equals(kind, "Grid", StringComparison.Ordinal) ||
             string.Equals(kind, "SharedParameter", StringComparison.Ordinal) ||
             string.Equals(kind, "ProjectParameter", StringComparison.Ordinal));
    }

    private static bool StateMatchesChangeIdentity(FamilyBrowserTrackedElementState state, string elementId)
    {
        return state == null || string.Equals(state.ElementId ?? string.Empty, elementId ?? string.Empty, StringComparison.Ordinal);
    }

    private static bool TryParseOptionalUtc(string value, out DateTime parsed)
    {
        parsed = DateTime.MinValue;
        return !string.IsNullOrWhiteSpace(value) && DateTime.TryParse(
            value,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
            out parsed);
    }

    private static bool EnsureElementChangeIntegrity(FamilyBrowserElementChangeCommit commit)
    {
        if (commit == null)
        {
            return false;
        }
        if (commit.SchemaVersion < MinimumSupportedElementCommitSchemaVersion || commit.SchemaVersion > MaximumSupportedElementCommitSchemaVersion)
        {
            return false;
        }
        if (!string.IsNullOrWhiteSpace(commit.IntegritySha256))
        {
            if (commit.IntegrityVersion != 1 && commit.IntegrityVersion != 2 && commit.IntegrityVersion != 3 && commit.IntegrityVersion != 4 && commit.IntegrityVersion != 5)
            {
                return false;
            }
            if (commit.SchemaVersion >= 6 && commit.IntegrityVersion != 5)
            {
                return false;
            }
            if (commit.SchemaVersion == 5 && commit.IntegrityVersion != 4)
            {
                return false;
            }
            return FixedTimeEquals(commit.IntegritySha256, ComputeElementChangeIntegrity(commit));
        }
        NormalizeDerivedElementChangeCounts(commit);
        commit.IntegrityVersion = commit.SchemaVersion >= 6 ? 5 : (commit.SchemaVersion >= 5 ? 4 : 3);
        commit.IntegritySha256 = ComputeElementChangeIntegrity(commit);
        return !string.IsNullOrWhiteSpace(commit.IntegritySha256);
    }

    private static string ComputeElementChangeIntegrity(FamilyBrowserElementChangeCommit commit)
    {
        if (commit == null)
        {
            return string.Empty;
        }
        if (commit.IntegrityVersion == 1)
        {
            return ComputeElementChangeIntegrityV1(commit);
        }
        StringBuilder canonical = new StringBuilder(4096);
        if (commit.IntegrityVersion == 2)
        {
            AppendCommitCanonicalV2(canonical, commit);
        }
        else if (commit.IntegrityVersion == 3)
        {
            AppendCommitCanonicalV3(canonical, commit);
        }
        else if (commit.IntegrityVersion == 4)
        {
            AppendCommitCanonicalV4(canonical, commit);
        }
        else
        {
            AppendCommitCanonicalV5(canonical, commit);
        }
        using (SHA256 sha = SHA256.Create())
        {
            return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(canonical.ToString()))).Replace("-", string.Empty);
        }
    }

    private static string ComputeElementChangeIntegrityV1(FamilyBrowserElementChangeCommit commit)
    {
        ElementChangeIntegrityV1Payload payload = CreateElementChangeIntegrityV1Payload(commit, false);
        using (SHA256 sha = SHA256.Create())
        {
            return BitConverter.ToString(sha.ComputeHash(SerializeJson(payload))).Replace("-", string.Empty);
        }
    }

    private static ElementChangeIntegrityV1Payload CreateElementChangeIntegrityV1Payload(FamilyBrowserElementChangeCommit commit, bool includeIntegrityHash)
    {
        if (commit == null)
        {
            return null;
        }
        return new ElementChangeIntegrityV1Payload
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
            IntegritySha256 = includeIntegrityHash ? commit.IntegritySha256 : string.Empty,
            TransactionNames = commit.TransactionNames,
            Changes = CreateElementChangeIntegrityV1Items(commit.Changes)
        };
    }

    private static PendingElementEnvelopeIntegrityV1EntryPayload CreatePendingElementEnvelopeIntegrityV1EntryPayload(FamilyBrowserElementChangeCommit commit)
    {
        if (commit == null)
        {
            return null;
        }
        return new PendingElementEnvelopeIntegrityV1EntryPayload
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
            Changes = CreateElementChangeIntegrityV1Items(commit.Changes)
        };
    }

    private static List<ElementChangeIntegrityV1ItemPayload> CreateElementChangeIntegrityV1Items(IEnumerable<FamilyBrowserElementChangeItem> changes)
    {
        return (changes ?? Enumerable.Empty<FamilyBrowserElementChangeItem>())
            .Select(delegate(FamilyBrowserElementChangeItem item)
            {
                if (item == null)
                {
                    return null;
                }
                return new ElementChangeIntegrityV1ItemPayload
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
                    Before = CreateElementChangeIntegrityV1State(item.Before),
                    After = CreateElementChangeIntegrityV1State(item.After),
                    TransactionNames = item.TransactionNames
                };
            })
            .ToList();
    }

    private static ElementChangeIntegrityV1StatePayload CreateElementChangeIntegrityV1State(FamilyBrowserTrackedElementState state)
    {
        if (state == null)
        {
            return null;
        }
        return new ElementChangeIntegrityV1StatePayload
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

    private static void AppendCommitCanonicalV2(StringBuilder target, FamilyBrowserElementChangeCommit commit)
    {
        AppendCanonical(target, commit.SchemaVersion);
        AppendCanonical(target, commit.EntryId);
        AppendCanonical(target, commit.ProjectTitle);
        AppendCanonical(target, commit.ProjectIdentityPath);
        AppendCanonical(target, commit.ProjectCanonicalPath);
        AppendCanonical(target, commit.ProjectComparableIdentity);
        AppendCanonical(target, commit.ProjectLegacyComparableIdentity);
        AppendCanonical(target, commit.CommitKind);
        AppendCanonical(target, commit.CommittedAtUtc);
        AppendCanonical(target, commit.RevitVersion);
        AppendCanonical(target, commit.RevitUserName);
        AppendCanonical(target, commit.WindowsUserName);
        AppendCanonical(target, commit.MachineName);
        AppendCanonical(target, commit.AttributionConfidence);
        AppendCanonical(target, commit.PolicyValidationState);
        AppendCanonical(target, commit.CoverageNote);
        AppendCanonical(target, commit.IsWorkshared);
        AppendCanonical(target, commit.BaselineCapturedLate);
        AppendCanonical(target, commit.TrackingStartedAtUtc);
        AppendCanonical(target, commit.BaselineCapturedAtUtc);
        AppendCanonical(target, commit.BaselineElapsedMilliseconds);
        AppendCanonical(target, commit.BaselineElementCount);
        AppendCanonical(target, commit.ActivityCount);
        AppendCanonical(target, commit.UndoCount);
        AppendCanonical(target, commit.RedoCount);
        AppendCanonical(target, commit.UnmatchedUndoCount);
        AppendCanonical(target, commit.UnmatchedRedoCount);
        AppendCanonical(target, commit.CreatedCount);
        AppendCanonical(target, commit.ModifiedCount);
        AppendCanonical(target, commit.DeletedCount);
        AppendCanonical(target, commit.TransientCreatedDeletedCount);
        AppendCanonical(target, commit.ExternalUpdateOverlapCount);
        AppendCanonicalList(target, commit.TransactionNames, delegate(StringBuilder builder, string value) { AppendCanonical(builder, value); });
        AppendCanonicalList(target, commit.Changes, AppendChangeCanonicalV2);
    }

    private static void AppendCommitCanonicalV3(StringBuilder target, FamilyBrowserElementChangeCommit commit)
    {
        AppendCommitCanonicalV2(target, commit);
        AppendCanonical(target, commit.LocalSaveProtectedAtUtc);
        AppendCanonical(target, commit.PublishedAtUtc);
    }

    private static void AppendCommitCanonicalV4(StringBuilder target, FamilyBrowserElementChangeCommit commit)
    {
        AppendCommitCanonicalV3(target, commit);
        AppendCanonical(target, commit.CoverageGapOnly);
        AppendCanonical(target, commit.EventReadFailureCount);
        AppendCanonical(target, commit.CommitBoundaryReadFailureCount);
    }

    private static void AppendCommitCanonicalV5(StringBuilder target, FamilyBrowserElementChangeCommit commit)
    {
        AppendCommitCanonicalV4(target, commit);
        AppendCanonicalList(target, commit.Changes, AppendChangeExtensionCanonicalV5);
    }

    private static void AppendChangeExtensionCanonicalV5(StringBuilder target, FamilyBrowserElementChangeItem item)
    {
        if (item == null)
        {
            AppendCanonical(target, "<null-change-extension>");
            return;
        }
        AppendCanonical(target, item.ElementName);
        AppendCanonical(target, item.TrackingKind);
        AppendStateExtensionCanonicalV5(target, item.Before);
        AppendStateExtensionCanonicalV5(target, item.After);
    }

    private static void AppendStateExtensionCanonicalV5(StringBuilder target, FamilyBrowserTrackedElementState state)
    {
        if (state == null)
        {
            AppendCanonical(target, "<null-state-extension>");
            return;
        }
        AppendCanonical(target, state.TrackingKind);
        AppendCanonical(target, state.SharedParameterGuid);
        AppendCanonical(target, state.ParameterBindingKind);
        AppendCanonical(target, state.ParameterBoundCategories);
        AppendCanonical(target, state.ParameterBoundCategoryIds);
        AppendCanonical(target, state.ParameterGroup);
        AppendCanonical(target, state.ParameterDataType);
        AppendCanonical(target, state.ParameterVariesAcrossGroups);
        AppendCanonical(target, state.GridCurveSignature);
        AppendCanonical(target, state.GridExtentsSignature);
        AppendCanonical(target, state.GridPinnedState);
    }

    private static void AppendChangeCanonicalV2(StringBuilder target, FamilyBrowserElementChangeItem item)
    {
        if (item == null)
        {
            AppendCanonical(target, "<null-change>");
            return;
        }
        AppendCanonical(target, item.ChangeKind);
        AppendCanonical(target, item.ElementId);
        AppendCanonical(target, item.UniqueId);
        AppendCanonical(target, item.ElementClass);
        AppendCanonical(target, item.CategoryName);
        AppendCanonical(target, item.FamilyName);
        AppendCanonical(target, item.TypeName);
        AppendCanonical(target, item.FirstObservedAtUtc);
        AppendCanonical(target, item.LastObservedAtUtc);
        AppendCanonical(target, item.ChangeSummary);
        AppendCanonical(target, item.PreviousStateUnavailable);
        AppendCanonical(target, item.ExternalUpdateOverlap);
        AppendStateCanonicalV2(target, item.Before);
        AppendStateCanonicalV2(target, item.After);
        AppendCanonicalList(target, item.TransactionNames, delegate(StringBuilder builder, string value) { AppendCanonical(builder, value); });
    }

    private static void AppendStateCanonicalV2(StringBuilder target, FamilyBrowserTrackedElementState state)
    {
        if (state == null)
        {
            AppendCanonical(target, "<null-state>");
            return;
        }
        AppendCanonical(target, state.ElementId);
        AppendCanonical(target, state.UniqueId);
        AppendCanonical(target, state.ElementClass);
        AppendCanonical(target, state.CategoryName);
        AppendCanonical(target, state.CategoryId);
        AppendCanonical(target, state.ElementName);
        AppendCanonical(target, state.FamilyName);
        AppendCanonical(target, state.TypeName);
        AppendCanonical(target, state.TypeId);
        AppendCanonical(target, state.LevelId);
        AppendCanonical(target, state.WorksetId);
        AppendCanonical(target, state.LocationSignature);
        AppendCanonical(target, state.StateSignature);
        AppendCanonical(target, state.IsElementType);
        AppendCanonical(target, state.IsViewSpecific);
    }

    private static void AppendCanonicalList<T>(StringBuilder target, IEnumerable<T> values, Action<StringBuilder, T> appendItem)
    {
        List<T> list = (values ?? Enumerable.Empty<T>()).ToList();
        AppendCanonical(target, list.Count);
        foreach (T value in list)
        {
            appendItem(target, value);
        }
    }

    private static void AppendCanonical(StringBuilder target, object value)
    {
        string text = value == null ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        target.Append(text.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(text).Append('|');
    }

    private static bool FixedTimeEquals(string left, string right)
    {
        byte[] leftBytes;
        byte[] rightBytes;
        try
        {
            leftBytes = Encoding.ASCII.GetBytes((left ?? string.Empty).Trim().ToUpperInvariant());
            rightBytes = Encoding.ASCII.GetBytes((right ?? string.Empty).Trim().ToUpperInvariant());
        }
        catch
        {
            return false;
        }
        if (leftBytes.Length != rightBytes.Length)
        {
            return false;
        }
        int difference = 0;
        for (int i = 0; i < leftBytes.Length; i++)
        {
            difference |= leftBytes[i] ^ rightBytes[i];
        }
        return difference == 0;
    }

    public static List<StandardRvtChangeCandidateEntry> LoadImmutableStandardCandidateEntries(string workspaceRoot, string sourceId, int limit)
    {
        List<StandardRvtChangeCandidateEntry> loaded = new List<StandardRvtChangeCandidateEntry>();
        FlushPending(workspaceRoot);
        string dataFolder = FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "StandardChangeCandidates");
        if (string.IsNullOrWhiteSpace(dataFolder))
        {
            return loaded;
        }
        string root = Path.Combine(dataFolder, "History", SafeFileName(sourceId));
        if (!Directory.Exists(root))
        {
            return loaded;
        }
        try
        {
            foreach (string path in SafeEnumerateFiles(root, "*.json", SearchOption.AllDirectories))
            {
                StandardRvtChangeCandidateEntry entry;
                if (TryReadJson(path, out entry) && entry != null)
                {
                    loaded.Add(entry);
                }
            }
        }
        catch
        {
        }
        return loaded
            .OrderByDescending(delegate(StandardRvtChangeCandidateEntry entry)
            {
                return ParseUtcOrMin(string.IsNullOrWhiteSpace(entry == null ? string.Empty : entry.CommittedAtUtc)
                    ? entry == null ? string.Empty : entry.RecordedAtUtc
                    : entry.CommittedAtUtc);
            })
            .ThenByDescending(delegate(StandardRvtChangeCandidateEntry entry) { return entry == null ? string.Empty : entry.EntryId ?? string.Empty; }, StringComparer.OrdinalIgnoreCase)
            .Take(Math.Max(1, limit))
            .ToList();
    }

    public static List<FamilyBrowserElementChangeCommit> LoadImmutableElementChangeCommits(string workspaceRoot, string projectIdentity, int limit)
    {
        return LoadImmutableElementChangeCommitResult(workspaceRoot, projectIdentity, limit).Commits;
    }

    private static List<FamilyBrowserElementChangeCommit> LoadAllImmutableElementChangeCommits(string workspaceRoot)
    {
        FlushPending(workspaceRoot);
        string historyRoot = FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "ElementChangeHistory");
        if (string.IsNullOrWhiteSpace(historyRoot) || !Directory.Exists(historyRoot))
        {
            return new List<FamilyBrowserElementChangeCommit>();
        }
        Dictionary<string, FamilyBrowserElementChangeCommit> loaded = new Dictionary<string, FamilyBrowserElementChangeCommit>(StringComparer.OrdinalIgnoreCase);
        HashSet<string> conflicts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        List<string> projectFolders;
        if (!TryEnumerateDirectories(historyRoot, out projectFolders))
        {
            return new List<FamilyBrowserElementChangeCommit>();
        }
        foreach (string projectFolder in projectFolders)
        {
            List<string> dayFolders;
            if (!TryEnumerateDirectories(projectFolder, out dayFolders))
            {
                continue;
            }
            foreach (string dayFolder in dayFolders)
            {
                List<string> historyFiles;
                if (!TryEnumerateFiles(dayFolder, "*.json", SearchOption.TopDirectoryOnly, out historyFiles))
                {
                    continue;
                }
                foreach (string path in historyFiles)
                {
                    FamilyBrowserElementChangeCommit commit;
                    string validationIssue;
                    if (!TryReadJson(path, out commit) || commit == null || string.IsNullOrWhiteSpace(commit.EntryId) ||
                        !ValidateElementChangeCommit(commit, string.Empty, false, out validationIssue) ||
                        conflicts.Contains(commit.EntryId))
                    {
                        continue;
                    }
                    FamilyBrowserElementChangeCommit existing;
                    if (loaded.TryGetValue(commit.EntryId, out existing))
                    {
                        if (!ElementCommitsHaveEquivalentPayload(existing, commit))
                        {
                            loaded.Remove(commit.EntryId);
                            conflicts.Add(commit.EntryId);
                        }
                        continue;
                    }
                    loaded[commit.EntryId] = commit;
                }
            }
        }
        return loaded.Values
            .OrderByDescending(delegate(FamilyBrowserElementChangeCommit commit) { return ParseUtc(ResolveElementCommitActivityAtUtc(commit)); })
            .ToList();
    }

    private static string ResolveElementCommitActivityAtUtc(FamilyBrowserElementChangeCommit commit)
    {
        if (commit == null)
        {
            return string.Empty;
        }
        if (!string.IsNullOrWhiteSpace(commit.PublishedAtUtc))
        {
            return commit.PublishedAtUtc;
        }
        if (!string.IsNullOrWhiteSpace(commit.LocalSaveProtectedAtUtc))
        {
            return commit.LocalSaveProtectedAtUtc;
        }
        return commit.CommittedAtUtc ?? string.Empty;
    }

    public static FamilyBrowserElementChangeHistoryLoadResult LoadImmutableElementChangeCommitResult(string workspaceRoot, string projectIdentity, int limit)
    {
        FamilyBrowserElementChangeHistoryLoadResult result = new FamilyBrowserElementChangeHistoryLoadResult();
        FamilyBrowserTrackingFlushResult flush = FlushPending(workspaceRoot);
        result.PendingDestinationMismatchCount = flush == null ? 0 : flush.DestinationMismatchCount;
        result.PendingCorruptRecordCount = flush == null ? 0 : flush.CorruptRecordCount;
        result.PendingFailedCount = flush == null ? 0 : flush.FailedCount;
        List<string> comparableIdentities = BuildProjectHistoryIdentities(projectIdentity);
        if (comparableIdentities.Count == 0)
        {
            return result;
        }
        string historyRoot = FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "ElementChangeHistory");
        if (string.IsNullOrWhiteSpace(historyRoot))
        {
            return result;
        }
        try
        {
            List<FamilyBrowserElementChangeCommit> loaded = new List<FamilyBrowserElementChangeCommit>();
            Dictionary<string, FamilyBrowserElementChangeCommit> seenEntries = new Dictionary<string, FamilyBrowserElementChangeCommit>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> conflictingEntryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<string> allHistoryRoots;
            if (!TryEnumerateDirectories(historyRoot, out allHistoryRoots))
            {
                result.InvalidRecordCount++;
                return result;
            }
            List<string> directRoots = comparableIdentities
                .Select(delegate(string identity) { return Path.Combine(historyRoot, HashText(identity).Substring(0, 32)); })
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(Directory.Exists)
                .ToList();
            List<KeyValuePair<string, bool>> roots = directRoots
                .Select(delegate(string root) { return new KeyValuePair<string, bool>(root, true); })
                .Concat(allHistoryRoots
                    .Where(delegate(string root) { return !directRoots.Contains(root, StringComparer.OrdinalIgnoreCase); })
                    .Select(delegate(string root) { return new KeyValuePair<string, bool>(root, false); }))
                .ToList();
            foreach (KeyValuePair<string, bool> rootCandidate in roots)
            {
                List<string> dayFolders;
                if (!TryEnumerateDirectories(rootCandidate.Key, out dayFolders))
                {
                    result.InvalidRecordCount++;
                    continue;
                }
                foreach (string dayFolder in dayFolders.OrderByDescending(delegate(string path) { return Path.GetFileName(path); }, StringComparer.OrdinalIgnoreCase))
                {
                    List<string> historyFiles;
                    if (!TryEnumerateFiles(dayFolder, "*.json", SearchOption.TopDirectoryOnly, out historyFiles))
                    {
                        result.InvalidRecordCount++;
                        continue;
                    }
                    foreach (string path in historyFiles)
                    {
                        FamilyBrowserElementChangeCommit commit;
                        if (!TryReadJson(path, out commit) || commit == null)
                        {
                            if (rootCandidate.Value)
                            {
                                result.ScannedFileCount++;
                                result.InvalidRecordCount++;
                            }
                            continue;
                        }
                        bool projectMatch = ElementCommitMatchesProject(commit, projectIdentity, comparableIdentities);
                        if (!projectMatch)
                        {
                            if (rootCandidate.Value)
                            {
                                result.ScannedFileCount++;
                                result.InvalidRecordCount++;
                            }
                            continue;
                        }
                        result.ScannedFileCount++;
                        string validationIssue;
                        if (!ValidateElementChangeCommit(commit, string.Empty, false, out validationIssue))
                        {
                            result.InvalidRecordCount++;
                            continue;
                        }
                        string entryId = commit.EntryId ?? string.Empty;
                        if (conflictingEntryIds.Contains(entryId))
                        {
                            result.InvalidRecordCount++;
                            continue;
                        }
                        FamilyBrowserElementChangeCommit existingEntry;
                        if (seenEntries.TryGetValue(entryId, out existingEntry))
                        {
                            if (ElementCommitsHaveEquivalentPayload(existingEntry, commit))
                            {
                                continue;
                            }
                            conflictingEntryIds.Add(entryId);
                            seenEntries.Remove(entryId);
                            loaded.Remove(existingEntry);
                            if (string.IsNullOrWhiteSpace(existingEntry.IntegritySha256) && result.LegacyUnverifiedCount > 0)
                            {
                                result.LegacyUnverifiedCount--;
                            }
                            result.InvalidRecordCount += 2;
                            continue;
                        }
                        if (string.IsNullOrWhiteSpace(commit.IntegritySha256))
                        {
                            result.LegacyUnverifiedCount++;
                        }
                        seenEntries[entryId] = commit;
                        loaded.Add(commit);
                    }
                }
            }
            result.TotalValidRecordCount = loaded.Count;
            result.Commits = loaded
                .OrderByDescending(delegate(FamilyBrowserElementChangeCommit commit) { return ParseUtc(commit.CommittedAtUtc); })
                .ThenByDescending(delegate(FamilyBrowserElementChangeCommit commit) { return commit.EntryId ?? string.Empty; }, StringComparer.OrdinalIgnoreCase)
                .Take(Math.Max(1, limit))
                .ToList();
        }
        catch
        {
            result.InvalidRecordCount++;
        }
        return result;
    }

    public static int GetPendingCount()
    {
        lock (SyncRoot)
        {
            return CountJsonFiles(BuildSpoolFolder("Operations")) + CountJsonFiles(BuildSpoolFolder("StandardCandidates")) + CountJsonFiles(BuildSpoolFolder("ElementChanges"));
        }
    }

    public static int GetPendingElementChangeCount(string workspaceRoot, string projectIdentity)
    {
        List<string> comparableIdentities = BuildProjectHistoryIdentities(projectIdentity);
        string destinationIdentity = BuildManagedDestinationIdentity(workspaceRoot);
        if (comparableIdentities.Count == 0)
        {
            return 0;
        }
        lock (SyncRoot)
        {
            int count = 0;
            List<string> pendingPaths;
            if (!TryEnumerateSpoolFiles(BuildSpoolFolder("ElementChanges"), out pendingPaths))
            {
                return 1;
            }
            foreach (string path in pendingPaths)
            {
                FamilyBrowserPendingElementChangeRecord record;
                FamilyBrowserElementChangeCommit commit;
                string intendedDestination;
                if (TryReadPendingElementChange(path, out record, out commit, out intendedDestination) &&
                    commit != null &&
                    ElementCommitMatchesProject(commit, projectIdentity, comparableIdentities) &&
                    CanFlushToDestination(intendedDestination, destinationIdentity, false))
                {
                    count++;
                }
            }
            return count;
        }
    }

    private static List<string> BuildProjectHistoryIdentities(string projectIdentity)
    {
        return new[]
        {
            FamilyBrowserPathIdentityService.GetStablePathIdentity(projectIdentity),
            FamilyBrowserPathIdentityService.GetComparableIdentity(projectIdentity)
        }
        .Where(delegate(string value) { return !string.IsNullOrWhiteSpace(value); })
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();
    }

    private static bool ElementCommitMatchesProject(FamilyBrowserElementChangeCommit commit, string projectIdentity, ICollection<string> comparableIdentities)
    {
        if (commit == null)
        {
            return false;
        }
        HashSet<string> expected = new HashSet<string>(comparableIdentities ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
        if (expected.Contains(commit.ProjectComparableIdentity ?? string.Empty) ||
            expected.Contains(commit.ProjectLegacyComparableIdentity ?? string.Empty))
        {
            return true;
        }
        string stableExpected = FamilyBrowserPathIdentityService.GetStablePathIdentity(projectIdentity);
        if (string.IsNullOrWhiteSpace(stableExpected))
        {
            return false;
        }
        return string.Equals(FamilyBrowserPathIdentityService.GetStablePathIdentity(commit.ProjectIdentityPath), stableExpected, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(FamilyBrowserPathIdentityService.GetStablePathIdentity(commit.ProjectCanonicalPath), stableExpected, StringComparison.OrdinalIgnoreCase);
    }

    private static FamilyBrowserTrackingFlushResult FlushPendingNoLock(string workspaceRoot, bool allowDestinationRebind = false, string rebindSourceDestination = "")
    {
        FamilyBrowserTrackingFlushResult result = new FamilyBrowserTrackingFlushResult();
        if (!FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
        {
            return result;
        }

        string destinationIdentity = BuildManagedDestinationIdentity(workspaceRoot);
        if (string.IsNullOrWhiteSpace(destinationIdentity))
        {
            result.FailedCount++;
            result.DestinationMismatchCount++;
            return result;
        }

        foreach (string path in EnumerateSpoolFilesOrMarkFailure(BuildSpoolFolder("Operations"), result))
        {
            FamilyBrowserPendingOperationRecord record;
            FamilyBrowserOperationLogEntry entry;
            string intendedDestination;
            if (!TryReadPendingOperation(path, out record, out entry, out intendedDestination) || entry == null)
            {
                result.FailedCount++;
                result.CorruptRecordCount++;
                continue;
            }
            if (!CanFlushToDestination(intendedDestination, destinationIdentity, allowDestinationRebind, rebindSourceDestination))
            {
                result.DestinationMismatchCount++;
                continue;
            }
            EnsureEntryId(entry);
            if (TryWriteJsonAtomic(BuildOperationHistoryPath(workspaceRoot, entry), entry))
            {
                if (TryDeleteChecked(path))
                {
                    result.OperationFlushedCount++;
                }
                else
                {
                    result.FailedCount++;
                }
            }
            else
            {
                result.FailedCount++;
            }
        }

        foreach (string path in EnumerateSpoolFilesOrMarkFailure(BuildSpoolFolder("StandardCandidates"), result))
        {
            FamilyBrowserPendingStandardCandidateRecord record;
            StandardRvtChangeCandidateEntry candidateEntry;
            string candidateSourceId;
            string intendedDestination;
            if (!TryReadPendingStandardCandidate(path, out record, out candidateEntry, out candidateSourceId, out intendedDestination) ||
                candidateEntry == null || string.IsNullOrWhiteSpace(candidateSourceId))
            {
                result.FailedCount++;
                result.CorruptRecordCount++;
                continue;
            }
            if (!CanFlushToDestination(intendedDestination, destinationIdentity, allowDestinationRebind, rebindSourceDestination))
            {
                result.DestinationMismatchCount++;
                continue;
            }
            EnsureEntryId(candidateEntry);
            if (TryWriteJsonAtomic(BuildCandidateHistoryPath(workspaceRoot, candidateSourceId, candidateEntry), candidateEntry))
            {
                if (TryDeleteChecked(path))
                {
                    result.StandardCandidateFlushedCount++;
                }
                else
                {
                    result.FailedCount++;
                }
            }
            else
            {
                result.FailedCount++;
            }
        }

        foreach (string path in EnumerateSpoolFilesOrMarkFailure(BuildSpoolFolder("ElementChanges"), result))
        {
            FamilyBrowserPendingElementChangeRecord record;
            FamilyBrowserElementChangeCommit commit;
            string intendedDestination;
            if (!TryReadPendingElementChange(path, out record, out commit, out intendedDestination) ||
                !ValidateElementChangeCommit(commit, commit == null ? string.Empty : commit.ProjectComparableIdentity, true, out _))
            {
                result.FailedCount++;
                result.CorruptRecordCount++;
                continue;
            }
            if (!CanFlushToDestination(intendedDestination, destinationIdentity, allowDestinationRebind, rebindSourceDestination))
            {
                result.DestinationMismatchCount++;
                continue;
            }
            EnsureEntryId(commit);
            if (TryWriteElementHistoryAtomic(BuildElementChangeHistoryPath(workspaceRoot, commit), commit))
            {
                if (TryDeleteChecked(path))
                {
                    result.ElementChangeFlushedCount++;
                }
                else
                {
                    result.FailedCount++;
                }
            }
            else
            {
                result.FailedCount++;
            }
        }
        return result;
    }

    private static int FlushFinalizedElementSessionCheckpointsNoLock(string workspaceRoot, FamilyBrowserTrackingFlushResult result)
    {
        if (!FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
        {
            return 0;
        }
        string destinationIdentity = BuildManagedDestinationIdentity(workspaceRoot);
        List<KeyValuePair<FamilyBrowserElementSessionCheckpoint, string>> finalized = new List<KeyValuePair<FamilyBrowserElementSessionCheckpoint, string>>();
        using (FileStream checkpointLock = TryAcquireElementSessionFileLock())
        {
            if (checkpointLock == null)
            {
                result.ElementSessionCheckpointLockUnavailable = true;
                return 0;
            }
            List<string> checkpointPaths;
            if (!TryEnumerateSpoolFiles(BuildSpoolFolder("ElementSessions"), out checkpointPaths))
            {
                result.ElementSessionCheckpointLockUnavailable = true;
                result.FailedCount++;
                return 0;
            }
            foreach (string path in checkpointPaths)
            {
                FamilyBrowserElementSessionCheckpoint checkpoint;
                if (!TryReadJson(path, out checkpoint) || checkpoint == null ||
                    !ValidateElementSessionCheckpointEnvelope(checkpoint) ||
                    !ValidateElementSessionCheckpointCommits(checkpoint) ||
                    !checkpoint.SynchronizationSucceeded ||
                    !CanFlushToDestination(checkpoint.DestinationIdentity, destinationIdentity, false))
                {
                    continue;
                }
                string revisionToken = ComputeElementSessionCheckpointRevisionToken(checkpoint);
                if (!string.IsNullOrWhiteSpace(revisionToken))
                {
                    finalized.Add(new KeyValuePair<FamilyBrowserElementSessionCheckpoint, string>(checkpoint, revisionToken));
                }
            }
        }

        int promoted = 0;
        foreach (KeyValuePair<FamilyBrowserElementSessionCheckpoint, string> item in finalized)
        {
            FamilyBrowserElementSessionCheckpoint checkpoint = item.Key;
            bool durable = PersistElementChangeCommits(workspaceRoot, checkpoint.Commits);
            if (!durable || !DeleteElementSessionCheckpointById(checkpoint.CheckpointId, item.Value))
            {
                result.FailedCount++;
                continue;
            }
            promoted++;
        }
        return promoted;
    }

    private static bool DeleteElementSessionCheckpointById(string checkpointId, string expectedCheckpointRevisionToken)
    {
        if (string.IsNullOrWhiteSpace(checkpointId))
        {
            return true;
        }
        using (FileStream checkpointLock = TryAcquireElementSessionFileLock())
        {
            if (checkpointLock == null)
            {
                return false;
            }
            string checkpointPath = BuildElementSessionCheckpointPath(checkpointId);
            if (!File.Exists(checkpointPath))
            {
                return true;
            }
            FamilyBrowserElementSessionCheckpoint checkpoint;
            if (!TryReadJson(checkpointPath, out checkpoint) || checkpoint == null ||
                !ValidateElementSessionCheckpointEnvelope(checkpoint) ||
                !ValidateElementSessionCheckpointCommits(checkpoint) ||
                !string.Equals(checkpoint.CheckpointId, checkpointId, StringComparison.OrdinalIgnoreCase) ||
                string.IsNullOrWhiteSpace(expectedCheckpointRevisionToken) ||
                !FixedTimeEquals(ComputeElementSessionCheckpointRevisionToken(checkpoint), expectedCheckpointRevisionToken))
            {
                return false;
            }
            return TryDeleteChecked(checkpointPath);
        }
    }

    private static string BuildOperationHistoryPath(string workspaceRoot, FamilyBrowserOperationLogEntry entry)
    {
        DateTime committed = ParseUtc(entry == null ? string.Empty : string.IsNullOrWhiteSpace(entry.CommittedAtUtc) ? entry.RecordedAtUtc : entry.CommittedAtUtc);
        string folder = Path.Combine(FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "OperationLogs"), "History", committed.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
        return Path.Combine(folder, SafeFileName(entry == null ? string.Empty : entry.EntryId) + ".json");
    }

    private static string BuildCandidateHistoryPath(string workspaceRoot, string sourceId, StandardRvtChangeCandidateEntry entry)
    {
        DateTime committed = ParseUtc(entry == null ? string.Empty : string.IsNullOrWhiteSpace(entry.CommittedAtUtc) ? entry.RecordedAtUtc : entry.CommittedAtUtc);
        string folder = Path.Combine(FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "StandardChangeCandidates"), "History", SafeFileName(sourceId), committed.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
        return Path.Combine(folder, SafeFileName(entry == null ? string.Empty : entry.EntryId) + ".json");
    }

    private static string BuildElementChangeHistoryPath(string workspaceRoot, FamilyBrowserElementChangeCommit commit)
    {
        string identity = commit == null ? string.Empty : commit.ProjectComparableIdentity;
        if (string.IsNullOrWhiteSpace(identity) && commit != null)
        {
            identity = FamilyBrowserPathIdentityService.GetComparableIdentity(commit.ProjectIdentityPath);
        }
        DateTime committed = ParseUtc(commit == null ? string.Empty : commit.CommittedAtUtc);
        string projectKey = HashText(identity).Substring(0, 32);
        string folder = Path.Combine(FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "ElementChangeHistory"), projectKey, committed.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
        return Path.Combine(folder, SafeFileName(commit == null ? string.Empty : commit.EntryId) + ".json");
    }

    private static string BuildOperationSpoolPath(string entryId)
    {
        return Path.Combine(BuildSpoolFolder("Operations"), SafeFileName(entryId) + ".json");
    }

    private static string BuildCandidateSpoolPath(string entryId)
    {
        return Path.Combine(BuildSpoolFolder("StandardCandidates"), SafeFileName(entryId) + ".json");
    }

    private static string BuildElementChangeSpoolPath(string entryId)
    {
        return Path.Combine(BuildSpoolFolder("ElementChanges"), SafeFileName(entryId) + ".json");
    }

    private static string BuildSpoolFolder(string name)
    {
        string root = string.IsNullOrWhiteSpace(_localSpoolRootOverrideForAudit)
            ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KKY", "FamilyBrowser", "PendingTracking")
            : _localSpoolRootOverrideForAudit;
        return Path.Combine(root, name ?? string.Empty);
    }

    private static FileStream TryAcquireElementSessionFileLock()
    {
        string folder = BuildSpoolFolder("ElementSessions");
        try
        {
            Directory.CreateDirectory(folder);
        }
        catch
        {
            return null;
        }
        string lockPath = Path.Combine(folder, ".checkpoint-write.lock");
        DateTime deadline = DateTime.UtcNow.AddMilliseconds(ElementSessionLockTimeoutMilliseconds);
        do
        {
            try
            {
                return new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
                return null;
            }
            Thread.Sleep(ElementSessionLockRetryMilliseconds);
        }
        while (DateTime.UtcNow < deadline);
        return null;
    }

    private static IEnumerable<string> EnumerateSpoolFilesOrMarkFailure(string folder, FamilyBrowserTrackingFlushResult result)
    {
        List<string> paths;
        if (TryEnumerateSpoolFiles(folder, out paths))
        {
            return paths;
        }
        if (result != null)
        {
            result.FailedCount++;
        }
        return Enumerable.Empty<string>();
    }

    private static bool TryEnumerateSpoolFiles(string folder, out List<string> paths)
    {
        return TryEnumerateFiles(folder, "*.json", SearchOption.TopDirectoryOnly, out paths);
    }

    private static bool TryEnumerateDirectories(string folder, out List<string> paths)
    {
        paths = new List<string>();
        if (string.IsNullOrWhiteSpace(folder))
        {
            return true;
        }
        if (IsEnumerationFailureForcedForAudit(folder))
        {
            return false;
        }
        try
        {
            if (!Directory.Exists(folder))
            {
                return true;
            }
            paths = Directory.EnumerateDirectories(folder, "*", SearchOption.TopDirectoryOnly).ToList();
            return true;
        }
        catch
        {
            paths.Clear();
            return false;
        }
    }

    private static bool TryEnumerateFiles(string folder, string pattern, SearchOption searchOption, out List<string> paths)
    {
        paths = new List<string>();
        if (string.IsNullOrWhiteSpace(folder))
        {
            return true;
        }
        if (IsEnumerationFailureForcedForAudit(folder))
        {
            return false;
        }
        try
        {
            if (!Directory.Exists(folder))
            {
                return true;
            }
            paths = Directory.EnumerateFiles(folder, string.IsNullOrWhiteSpace(pattern) ? "*" : pattern, searchOption).ToList();
            return true;
        }
        catch
        {
            paths.Clear();
            return false;
        }
    }

    private static bool IsEnumerationFailureForcedForAudit(string folder)
    {
        string forced;
        lock (SyncRoot)
        {
            forced = _enumerationFailureFolderOverrideForAudit;
        }
        if (string.IsNullOrWhiteSpace(forced))
        {
            return false;
        }
        return string.Equals(NormalizeDirectoryComparisonPath(folder), NormalizeDirectoryComparisonPath(forced), StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeDirectoryComparisonPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }
        try
        {
            return Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
        catch
        {
            return path.Trim().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
    }

    private static IEnumerable<string> SafeEnumerateDirectories(string folder)
    {
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            return Enumerable.Empty<string>();
        }
        try
        {
            return Directory.EnumerateDirectories(folder, "*", SearchOption.TopDirectoryOnly).ToList();
        }
        catch
        {
            return Enumerable.Empty<string>();
        }
    }

    private static IEnumerable<string> SafeEnumerateFiles(string folder, string pattern, SearchOption searchOption)
    {
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            return Enumerable.Empty<string>();
        }
        try
        {
            return Directory.EnumerateFiles(folder, string.IsNullOrWhiteSpace(pattern) ? "*" : pattern, searchOption).ToList();
        }
        catch
        {
            return Enumerable.Empty<string>();
        }
    }

    private static int CountJsonFiles(string folder)
    {
        List<string> files;
        if (!TryEnumerateFiles(folder, "*.json", SearchOption.TopDirectoryOnly, out files))
        {
            return 1;
        }
        return files.Count;
    }

    private static bool TryWriteJsonAtomic<T>(string path, T value)
    {
        if (string.IsNullOrWhiteSpace(path) || value == null)
        {
            return false;
        }
        string temporary = string.Empty;
        try
        {
            byte[] payload = SerializeJson(value);
            if (File.Exists(path))
            {
                return FileMatchesPayload(path, payload);
            }
            string folder = Path.GetDirectoryName(path);
            if (string.IsNullOrWhiteSpace(folder))
            {
                return false;
            }
            Directory.CreateDirectory(folder);
            temporary = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
            using (FileStream stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                stream.Write(payload, 0, payload.Length);
                stream.Flush(true);
            }
            try
            {
                File.Move(temporary, path);
            }
            catch (IOException)
            {
                if (!File.Exists(path))
                {
                    throw;
                }
            }
            return FileMatchesPayload(path, payload);
        }
        catch
        {
            return false;
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(temporary) && File.Exists(temporary))
            {
                TryDelete(temporary);
            }
        }
    }

    private static bool TryWriteElementHistoryAtomic(string path, FamilyBrowserElementChangeCommit commit)
    {
        if (!File.Exists(path))
        {
            return TryWriteJsonAtomic(path, commit);
        }
        FamilyBrowserElementChangeCommit existing;
        string issue;
        if (!TryReadJson(path, out existing) || existing == null ||
            !ValidateElementChangeCommit(existing, string.Empty, false, out issue) ||
            !string.Equals(existing.EntryId ?? string.Empty, commit == null ? string.Empty : commit.EntryId ?? string.Empty, StringComparison.OrdinalIgnoreCase) ||
            !ElementCommitMatchesSameProject(existing, commit))
        {
            return false;
        }
        if (!string.IsNullOrWhiteSpace(existing.IntegritySha256) && commit != null &&
            existing.IntegrityVersion == commit.IntegrityVersion &&
            FixedTimeEquals(existing.IntegritySha256, commit.IntegritySha256))
        {
            return true;
        }
        if (string.IsNullOrWhiteSpace(existing.IntegritySha256) && commit != null)
        {
            StringBuilder existingCanonical = new StringBuilder(4096);
            StringBuilder incomingCanonical = new StringBuilder(4096);
            AppendCommitCanonicalV2(existingCanonical, existing);
            AppendCommitCanonicalV2(incomingCanonical, commit);
            return string.Equals(existingCanonical.ToString(), incomingCanonical.ToString(), StringComparison.Ordinal);
        }
        return false;
    }

    private static bool ElementCommitMatchesSameProject(FamilyBrowserElementChangeCommit left, FamilyBrowserElementChangeCommit right)
    {
        if (left == null || right == null)
        {
            return false;
        }
        if (string.Equals(left.ProjectComparableIdentity ?? string.Empty, right.ProjectComparableIdentity ?? string.Empty, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
        string leftStable = FamilyBrowserPathIdentityService.GetStablePathIdentity(left.ProjectIdentityPath);
        string rightStable = FamilyBrowserPathIdentityService.GetStablePathIdentity(right.ProjectIdentityPath);
        return !string.IsNullOrWhiteSpace(leftStable) && string.Equals(leftStable, rightStable, StringComparison.OrdinalIgnoreCase);
    }

    private static bool ElementCommitsHaveEquivalentPayload(FamilyBrowserElementChangeCommit left, FamilyBrowserElementChangeCommit right)
    {
        if (left == null || right == null ||
            !string.Equals(left.EntryId ?? string.Empty, right.EntryId ?? string.Empty, StringComparison.OrdinalIgnoreCase) ||
            !ElementCommitMatchesSameProject(left, right))
        {
            return false;
        }
        bool leftSigned = !string.IsNullOrWhiteSpace(left.IntegritySha256);
        bool rightSigned = !string.IsNullOrWhiteSpace(right.IntegritySha256);
        if (leftSigned && rightSigned && left.IntegrityVersion == right.IntegrityVersion)
        {
            return FixedTimeEquals(left.IntegritySha256, right.IntegritySha256);
        }
        StringBuilder leftCanonical = new StringBuilder(4096);
        StringBuilder rightCanonical = new StringBuilder(4096);
        if (leftSigned && rightSigned && (left.IntegrityVersion >= 4 || right.IntegrityVersion >= 4))
        {
            AppendCommitCanonicalV4(leftCanonical, left);
            AppendCommitCanonicalV4(rightCanonical, right);
        }
        else if (leftSigned && rightSigned && (left.IntegrityVersion >= 3 || right.IntegrityVersion >= 3))
        {
            AppendCommitCanonicalV3(leftCanonical, left);
            AppendCommitCanonicalV3(rightCanonical, right);
        }
        else
        {
            AppendCommitCanonicalV2(leftCanonical, left);
            AppendCommitCanonicalV2(rightCanonical, right);
        }
        return string.Equals(leftCanonical.ToString(), rightCanonical.ToString(), StringComparison.Ordinal);
    }

    private static bool TryWriteMutableJsonAtomic<T>(string path, T value)
    {
        if (string.IsNullOrWhiteSpace(path) || value == null)
        {
            return false;
        }
        string temporary = string.Empty;
        try
        {
            byte[] payload = SerializeJson(value);
            string folder = Path.GetDirectoryName(path);
            if (string.IsNullOrWhiteSpace(folder))
            {
                return false;
            }
            Directory.CreateDirectory(folder);
            temporary = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
            using (FileStream stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                stream.Write(payload, 0, payload.Length);
                stream.Flush(true);
            }
            FamilyBrowserAtomicFileService.Promote(temporary, path);
            temporary = string.Empty;
            return FileMatchesPayload(path, payload);
        }
        catch
        {
            return false;
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(temporary) && File.Exists(temporary))
            {
                TryDelete(temporary);
            }
        }
    }

    private static byte[] SerializeJson<T>(T value)
    {
        using (MemoryStream stream = new MemoryStream())
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
            serializer.WriteObject(stream, value);
            return stream.ToArray();
        }
    }

    private static bool FileMatchesPayload(string path, byte[] expected)
    {
        if (string.IsNullOrWhiteSpace(path) || expected == null)
        {
            return false;
        }
        try
        {
            FileInfo info = new FileInfo(path);
            if (!info.Exists || info.Length != expected.LongLength)
            {
                return false;
            }
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                int offset = 0;
                byte[] buffer = new byte[8192];
                while (offset < expected.Length)
                {
                    int read = stream.Read(buffer, 0, Math.Min(buffer.Length, expected.Length - offset));
                    if (read <= 0)
                    {
                        return false;
                    }
                    for (int i = 0; i < read; i++)
                    {
                        if (buffer[i] != expected[offset + i])
                        {
                            return false;
                        }
                    }
                    offset += read;
                }
                return stream.ReadByte() < 0;
            }
        }
        catch
        {
            return false;
        }
    }

    private static bool TryReadJson<T>(string path, out T value)
    {
        value = default(T);
        try
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
                object loaded = serializer.ReadObject(stream);
                if (loaded is T)
                {
                    value = (T)loaded;
                    return true;
                }
            }
        }
        catch
        {
        }
        return false;
    }

    private static DateTime ParseUtc(string value)
    {
        DateTime parsed;
        if (DateTime.TryParse(value ?? string.Empty, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out parsed))
        {
            return parsed.ToUniversalTime();
        }
        return DateTime.UtcNow;
    }

    private static DateTime ParseUtcOrMin(string value)
    {
        DateTime parsed;
        if (DateTime.TryParse(value ?? string.Empty, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out parsed))
        {
            return parsed.ToUniversalTime();
        }
        return DateTime.MinValue;
    }

    private static void EnsureEntryId(FamilyBrowserOperationLogEntry entry)
    {
        if (entry != null && string.IsNullOrWhiteSpace(entry.EntryId))
        {
            entry.EntryId = Guid.NewGuid().ToString("N");
        }
    }

    private static void EnsureEntryId(StandardRvtChangeCandidateEntry entry)
    {
        if (entry != null && string.IsNullOrWhiteSpace(entry.EntryId))
        {
            entry.EntryId = Guid.NewGuid().ToString("N");
        }
    }

    private static void EnsureEntryId(FamilyBrowserElementChangeCommit commit)
    {
        if (commit != null && string.IsNullOrWhiteSpace(commit.EntryId))
        {
            commit.EntryId = Guid.NewGuid().ToString("N");
        }
    }

    private static string HashText(string value)
    {
        using (SHA256 sha = SHA256.Create())
        {
            return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty))).Replace("-", string.Empty);
        }
    }

    private static string SafeFileName(string value)
    {
        char[] invalid = Path.GetInvalidFileNameChars();
        string safe = new string((value ?? string.Empty).Select(delegate(char ch) { return invalid.Contains(ch) ? '_' : ch; }).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(safe) ? "entry" : safe;
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
        }
    }

    private static bool TryDeleteChecked(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            return !File.Exists(path);
        }
        catch
        {
            return false;
        }
    }
}
