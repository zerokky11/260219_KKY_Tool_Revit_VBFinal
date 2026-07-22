using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;

public sealed class FamilyBrowserTrackedElementState
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
    public string TrackingKind { get; set; }
    public string SharedParameterGuid { get; set; }
    public string ParameterBindingKind { get; set; }
    public string ParameterBoundCategories { get; set; }
    public string ParameterBoundCategoryIds { get; set; }
    public string ParameterGroup { get; set; }
    public string ParameterDataType { get; set; }
    public string ParameterVariesAcrossGroups { get; set; }
    public string GridCurveSignature { get; set; }
    public string GridExtentsSignature { get; set; }
    public string GridPinnedState { get; set; }
    public string StateSignature { get; set; }
    public bool IsElementType { get; set; }
    public bool IsViewSpecific { get; set; }

    public FamilyBrowserTrackedElementState()
    {
        ElementId = string.Empty;
        UniqueId = string.Empty;
        ElementClass = string.Empty;
        CategoryName = string.Empty;
        CategoryId = string.Empty;
        ElementName = string.Empty;
        FamilyName = string.Empty;
        TypeName = string.Empty;
        TypeId = string.Empty;
        LevelId = string.Empty;
        WorksetId = string.Empty;
        LocationSignature = string.Empty;
        TrackingKind = string.Empty;
        SharedParameterGuid = string.Empty;
        ParameterBindingKind = string.Empty;
        ParameterBoundCategories = string.Empty;
        ParameterBoundCategoryIds = string.Empty;
        ParameterGroup = string.Empty;
        ParameterDataType = string.Empty;
        ParameterVariesAcrossGroups = string.Empty;
        GridCurveSignature = string.Empty;
        GridExtentsSignature = string.Empty;
        GridPinnedState = string.Empty;
        StateSignature = string.Empty;
    }
}

public sealed class FamilyBrowserElementChangeItem
{
    public string ChangeKind { get; set; }
    public string ElementId { get; set; }
    public string UniqueId { get; set; }
    public string ElementClass { get; set; }
    public string CategoryName { get; set; }
    public string ElementName { get; set; }
    public string FamilyName { get; set; }
    public string TypeName { get; set; }
    public string TrackingKind { get; set; }
    public string FirstObservedAtUtc { get; set; }
    public string LastObservedAtUtc { get; set; }
    public string ChangeSummary { get; set; }
    public bool PreviousStateUnavailable { get; set; }
    public bool ExternalUpdateOverlap { get; set; }
    public FamilyBrowserTrackedElementState Before { get; set; }
    public FamilyBrowserTrackedElementState After { get; set; }
    public List<string> TransactionNames { get; set; }

    public FamilyBrowserElementChangeItem()
    {
        ChangeKind = string.Empty;
        ElementId = string.Empty;
        UniqueId = string.Empty;
        ElementClass = string.Empty;
        CategoryName = string.Empty;
        ElementName = string.Empty;
        FamilyName = string.Empty;
        TypeName = string.Empty;
        TrackingKind = string.Empty;
        FirstObservedAtUtc = string.Empty;
        LastObservedAtUtc = string.Empty;
        ChangeSummary = string.Empty;
        TransactionNames = new List<string>();
    }
}

public sealed class FamilyBrowserElementChangeCommit
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
    public string LocalSaveProtectedAtUtc { get; set; }
    public string PublishedAtUtc { get; set; }
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
    public int TransientCreatedDeletedCount { get; set; }
    public int ExternalUpdateOverlapCount { get; set; }
    public bool CoverageGapOnly { get; set; }
    public int EventReadFailureCount { get; set; }
    public int CommitBoundaryReadFailureCount { get; set; }
    public int IntegrityVersion { get; set; }
    public string IntegritySha256 { get; set; }
    public List<string> TransactionNames { get; set; }
    public List<FamilyBrowserElementChangeItem> Changes { get; set; }

    public FamilyBrowserElementChangeCommit()
    {
        SchemaVersion = 6;
        EntryId = Guid.NewGuid().ToString("N");
        ProjectTitle = string.Empty;
        ProjectIdentityPath = string.Empty;
        ProjectCanonicalPath = string.Empty;
        ProjectComparableIdentity = string.Empty;
        ProjectLegacyComparableIdentity = string.Empty;
        CommitKind = string.Empty;
        CommittedAtUtc = string.Empty;
        LocalSaveProtectedAtUtc = string.Empty;
        PublishedAtUtc = string.Empty;
        RevitVersion = string.Empty;
        RevitUserName = string.Empty;
        WindowsUserName = string.Empty;
        MachineName = string.Empty;
        AttributionConfidence = "ClientObserved";
        PolicyValidationState = "LiveEnabled";
        CoverageNote = string.Empty;
        TrackingStartedAtUtc = string.Empty;
        BaselineCapturedAtUtc = string.Empty;
        IntegritySha256 = string.Empty;
        TransactionNames = new List<string>();
        Changes = new List<FamilyBrowserElementChangeItem>();
    }
}

public static class FamilyBrowserElementChangeTrackingService
{
    private sealed class ParameterBindingState
    {
        public string ElementId { get; set; }
        public string DefinitionName { get; set; }
        public string SharedGuid { get; set; }
        public string BindingKind { get; set; }
        public string CategoryNames { get; set; }
        public string CategoryIds { get; set; }
        public string ParameterGroup { get; set; }
        public string DataType { get; set; }
        public string VariesAcrossGroups { get; set; }

        public ParameterBindingState()
        {
            ElementId = string.Empty;
            DefinitionName = string.Empty;
            SharedGuid = string.Empty;
            BindingKind = string.Empty;
            CategoryNames = string.Empty;
            CategoryIds = string.Empty;
            ParameterGroup = string.Empty;
            DataType = string.Empty;
            VariesAcrossGroups = string.Empty;
        }
    }

    private sealed class StateCaptureContext
    {
        public Dictionary<string, ParameterBindingState> ParameterByElementId { get; private set; }
        public Dictionary<string, ParameterBindingState> ParameterByGuid { get; private set; }
        public Dictionary<string, ParameterBindingState> ParameterByName { get; private set; }
        public List<ParameterElement> ParameterElements { get; private set; }
        public bool ParameterBindingsReadSucceeded { get; set; }

        public StateCaptureContext()
        {
            ParameterByElementId = new Dictionary<string, ParameterBindingState>(StringComparer.Ordinal);
            ParameterByGuid = new Dictionary<string, ParameterBindingState>(StringComparer.OrdinalIgnoreCase);
            ParameterByName = new Dictionary<string, ParameterBindingState>(StringComparer.OrdinalIgnoreCase);
            ParameterElements = new List<ParameterElement>();
        }
    }

    private sealed class ChangeActivity
    {
        public string Operation { get; set; }
        public string ObservedAtUtc { get; set; }
        public string LastObservedAtUtc { get; set; }
        public HashSet<string> AddedIds { get; private set; }
        public HashSet<string> ModifiedIds { get; private set; }
        public HashSet<string> DeletedIds { get; private set; }
        public List<string> TransactionNames { get; private set; }

        public ChangeActivity()
        {
            Operation = string.Empty;
            ObservedAtUtc = string.Empty;
            LastObservedAtUtc = string.Empty;
            AddedIds = new HashSet<string>(StringComparer.Ordinal);
            ModifiedIds = new HashSet<string>(StringComparer.Ordinal);
            DeletedIds = new HashSet<string>(StringComparer.Ordinal);
            TransactionNames = new List<string>();
        }

        public IEnumerable<string> AllElementIds()
        {
            return AddedIds.Concat(ModifiedIds).Concat(DeletedIds).Distinct(StringComparer.Ordinal);
        }
    }

    private sealed class ChangeActivityIndex
    {
        public HashSet<string> ActiveIds { get; private set; }
        public HashSet<string> AddedIds { get; private set; }
        public HashSet<string> DeletedIds { get; private set; }
        public Dictionary<string, List<ChangeActivity>> ActivitiesByElementId { get; private set; }
        public List<string> TransactionNames { get; private set; }

        public ChangeActivityIndex()
        {
            ActiveIds = new HashSet<string>(StringComparer.Ordinal);
            AddedIds = new HashSet<string>(StringComparer.Ordinal);
            DeletedIds = new HashSet<string>(StringComparer.Ordinal);
            ActivitiesByElementId = new Dictionary<string, List<ChangeActivity>>(StringComparer.Ordinal);
            TransactionNames = new List<string>();
        }
    }

    private sealed class DocumentSession
    {
        public string RuntimeKey { get; set; }
        public string WorkspaceRoot { get; set; }
        public string TrackingStartedAtUtc { get; set; }
        public string BaselineCapturedAtUtc { get; set; }
        public string CheckpointProjectIdentityPath { get; set; }
        public string CheckpointLocalDocumentPath { get; set; }
        public string CheckpointRevitUserName { get; set; }
        public string CheckpointRevisionToken { get; set; }
        public long BaselineElapsedMilliseconds { get; set; }
        public bool BaselineCapturedLate { get; set; }
        public bool SynchronizingWithCentral { get; set; }
        public bool ReloadingLatest { get; set; }
        public bool RecoveryOnly { get; set; }
        public bool PolicyDisableDeferred { get; set; }
        public int UndoCount { get; set; }
        public int RedoCount { get; set; }
        public int UnmatchedUndoCount { get; set; }
        public int UnmatchedRedoCount { get; set; }
        public int EventReadFailureCount { get; set; }
        public int CommitBoundaryReadFailureCount { get; set; }
        public bool LocalSaveCheckpointFailed { get; set; }
        public bool CommitBoundaryProtectionFailed { get; set; }
        public bool ExternalRebaseFailed { get; set; }
        public Dictionary<string, FamilyBrowserTrackedElementState> Baseline { get; private set; }
        public Dictionary<string, FamilyBrowserTrackedElementState> Current { get; private set; }
        public Dictionary<string, FamilyBrowserTrackedElementState> DeletedLastKnown { get; private set; }
        public HashSet<string> IgnoredAuxiliaryElementIds { get; private set; }
        public HashSet<string> TouchedIds { get; private set; }
        public HashSet<string> UnknownPreviousStateIds { get; private set; }
        public HashSet<string> SuppressedExternalIds { get; private set; }
        public HashSet<string> ExternalOverlapIds { get; private set; }
        public HashSet<string> AmbiguousActivityIds { get; private set; }
        public List<ChangeActivity> AppliedActivities { get; private set; }
        public List<ChangeActivity> UndoneActivities { get; private set; }
        public List<FamilyBrowserElementChangeCommit> RecoveredLocalSaveCommits { get; private set; }

        public DocumentSession()
        {
            RuntimeKey = string.Empty;
            WorkspaceRoot = string.Empty;
            TrackingStartedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
            BaselineCapturedAtUtc = string.Empty;
            CheckpointProjectIdentityPath = string.Empty;
            CheckpointLocalDocumentPath = string.Empty;
            CheckpointRevitUserName = string.Empty;
            CheckpointRevisionToken = string.Empty;
            Baseline = new Dictionary<string, FamilyBrowserTrackedElementState>(StringComparer.Ordinal);
            Current = new Dictionary<string, FamilyBrowserTrackedElementState>(StringComparer.Ordinal);
            DeletedLastKnown = new Dictionary<string, FamilyBrowserTrackedElementState>(StringComparer.Ordinal);
            IgnoredAuxiliaryElementIds = new HashSet<string>(StringComparer.Ordinal);
            TouchedIds = new HashSet<string>(StringComparer.Ordinal);
            UnknownPreviousStateIds = new HashSet<string>(StringComparer.Ordinal);
            SuppressedExternalIds = new HashSet<string>(StringComparer.Ordinal);
            ExternalOverlapIds = new HashSet<string>(StringComparer.Ordinal);
            AmbiguousActivityIds = new HashSet<string>(StringComparer.Ordinal);
            AppliedActivities = new List<ChangeActivity>();
            UndoneActivities = new List<ChangeActivity>();
            RecoveredLocalSaveCommits = new List<FamilyBrowserElementChangeCommit>();
        }

        public void ResetBaseline(
            Dictionary<string, FamilyBrowserTrackedElementState> states,
            HashSet<string> ignoredAuxiliaryElementIds,
            long elapsedMilliseconds,
            bool capturedLate)
        {
            Baseline = new Dictionary<string, FamilyBrowserTrackedElementState>(states ?? new Dictionary<string, FamilyBrowserTrackedElementState>(), StringComparer.Ordinal);
            Current = new Dictionary<string, FamilyBrowserTrackedElementState>(Baseline, StringComparer.Ordinal);
            DeletedLastKnown.Clear();
            IgnoredAuxiliaryElementIds = new HashSet<string>(ignoredAuxiliaryElementIds ?? new HashSet<string>(), StringComparer.Ordinal);
            TouchedIds.Clear();
            UnknownPreviousStateIds.Clear();
            SuppressedExternalIds.Clear();
            ExternalOverlapIds.Clear();
            AmbiguousActivityIds.Clear();
            AppliedActivities.Clear();
            UndoneActivities.Clear();
            UndoCount = 0;
            RedoCount = 0;
            UnmatchedUndoCount = 0;
            UnmatchedRedoCount = 0;
            EventReadFailureCount = 0;
            CommitBoundaryReadFailureCount = 0;
            LocalSaveCheckpointFailed = false;
            CommitBoundaryProtectionFailed = false;
            ExternalRebaseFailed = false;
            PolicyDisableDeferred = false;
            BaselineCapturedLate = capturedLate;
            BaselineElapsedMilliseconds = elapsedMilliseconds;
            BaselineCapturedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
            TrackingStartedAtUtc = BaselineCapturedAtUtc;
        }

        public void PromoteCurrentToBaseline(long elapsedMilliseconds)
        {
            Baseline = Current ?? new Dictionary<string, FamilyBrowserTrackedElementState>(StringComparer.Ordinal);
            Current = new Dictionary<string, FamilyBrowserTrackedElementState>(Baseline, StringComparer.Ordinal);
            DeletedLastKnown.Clear();
            TouchedIds.Clear();
            UnknownPreviousStateIds.Clear();
            SuppressedExternalIds.Clear();
            ExternalOverlapIds.Clear();
            AmbiguousActivityIds.Clear();
            AppliedActivities.Clear();
            UndoneActivities.Clear();
            UndoCount = 0;
            RedoCount = 0;
            UnmatchedUndoCount = 0;
            UnmatchedRedoCount = 0;
            EventReadFailureCount = 0;
            CommitBoundaryReadFailureCount = 0;
            LocalSaveCheckpointFailed = false;
            CommitBoundaryProtectionFailed = false;
            ExternalRebaseFailed = false;
            PolicyDisableDeferred = false;
            BaselineCapturedLate = false;
            BaselineElapsedMilliseconds = Math.Max(0L, elapsedMilliseconds);
            BaselineCapturedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
            TrackingStartedAtUtc = BaselineCapturedAtUtc;
        }

        public void RebaseAfterExternalUpdate(
            Dictionary<string, FamilyBrowserTrackedElementState> rebasedBaseline,
            Dictionary<string, FamilyBrowserTrackedElementState> refreshedCurrent,
            Dictionary<string, FamilyBrowserTrackedElementState> deletedLastKnown,
            HashSet<string> ignoredAuxiliaryElementIds,
            long elapsedMilliseconds)
        {
            Baseline = new Dictionary<string, FamilyBrowserTrackedElementState>(rebasedBaseline ?? new Dictionary<string, FamilyBrowserTrackedElementState>(), StringComparer.Ordinal);
            Current = new Dictionary<string, FamilyBrowserTrackedElementState>(refreshedCurrent ?? new Dictionary<string, FamilyBrowserTrackedElementState>(), StringComparer.Ordinal);
            DeletedLastKnown = new Dictionary<string, FamilyBrowserTrackedElementState>(deletedLastKnown ?? new Dictionary<string, FamilyBrowserTrackedElementState>(), StringComparer.Ordinal);
            IgnoredAuxiliaryElementIds = new HashSet<string>(ignoredAuxiliaryElementIds ?? new HashSet<string>(), StringComparer.Ordinal);
            BaselineElapsedMilliseconds += elapsedMilliseconds;
        }

        public void ReplaceIgnoredAuxiliaryElementIds(IEnumerable<string> elementIds)
        {
            IgnoredAuxiliaryElementIds = new HashSet<string>(elementIds ?? Enumerable.Empty<string>(), StringComparer.Ordinal);
        }
    }

    private sealed class RuntimeDocumentIdentity
    {
        public string Value { get; private set; }

        public RuntimeDocumentIdentity(string value)
        {
            Value = string.IsNullOrWhiteSpace(value)
                ? "document-runtime:" + Guid.NewGuid().ToString("N")
                : value;
        }
    }

    private const int PolicyRefreshSeconds = 5;
    private static readonly object SyncRoot = new object();
    private static readonly object PerformanceLogSyncRoot = new object();
    private static readonly Dictionary<string, DocumentSession> Sessions = new Dictionary<string, DocumentSession>(StringComparer.Ordinal);
    private static readonly ConditionalWeakTable<Document, RuntimeDocumentIdentity> RuntimeDocumentIdentities = new ConditionalWeakTable<Document, RuntimeDocumentIdentity>();
    private static readonly Dictionary<int, string> ClosingSessionKeysByDocumentId = new Dictionary<int, string>();
    private static readonly HashSet<string> SynchronizingDocumentKeys = new HashSet<string>(StringComparer.Ordinal);
    private static readonly HashSet<string> ReloadingDocumentKeys = new HashSet<string>(StringComparer.Ordinal);
    private static readonly HashSet<string> UnknownSynchronizationStartRoots = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private static readonly HashSet<string> UnknownReloadLatestStartRoots = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private static readonly Dictionary<string, DateTime> LastDiagnosticUtcByKey = new Dictionary<string, DateTime>(StringComparer.Ordinal);
    private static readonly Dictionary<string, bool> ProjectCatalogObservationRequiredByRuntimeKey = new Dictionary<string, bool>(StringComparer.Ordinal);
    private static string _cachedPolicyWorkspaceRoot = string.Empty;
    private static bool _cachedPolicyEnabled;
    private static bool _cachedPolicyKnown;
    private static DateTime _cachedPolicyAtUtc = DateTime.MinValue;
    private static FamilyBrowserStandardPolicy _cachedPolicy;
    private static object _reloadLatestEventSource;
    private static EventInfo _reloadingLatestEvent;
    private static EventInfo _reloadedLatestEvent;
    private static Delegate _reloadingLatestHandler;
    private static Delegate _reloadedLatestHandler;
    private static Func<string> _workspaceRootResolver;
    private static int _documentSessionBaselineRefreshRequested;
    [ThreadStatic]
    private static int _managedFolderTransitionAuthorizationDepth;

    private sealed class ManagedFolderTransitionAuthorization : IDisposable
    {
        private bool _disposed;

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            if (_managedFolderTransitionAuthorizationDepth > 0)
            {
                _managedFolderTransitionAuthorizationDepth--;
            }
        }
    }

    public static void NotifyPolicyChanged(string workspaceRoot, bool enabled)
    {
        bool disableDeferred = false;
        string root = workspaceRoot ?? string.Empty;
        lock (SyncRoot)
        {
            _cachedPolicyWorkspaceRoot = root;
            _cachedPolicyKnown = true;
            _cachedPolicy = null;
            bool hasUncommittedEvidence = Sessions.Values.Any(delegate(DocumentSession session)
            {
                return session != null &&
                    string.Equals(session.WorkspaceRoot ?? string.Empty, root, StringComparison.OrdinalIgnoreCase) &&
                    HasUncommittedSessionEvidence(session);
            });
            if (!enabled && hasUncommittedEvidence)
            {
                DisableLiveTrackingNoLock(root, true);
                _cachedPolicyEnabled = true;
                _cachedPolicyAtUtc = DateTime.MinValue;
                disableDeferred = true;
            }
            else
            {
                _cachedPolicyEnabled = enabled;
                _cachedPolicyAtUtc = DateTime.MinValue;
                if (enabled)
                {
                    EnableLiveTrackingNoLock(workspaceRoot);
                }
                else
                {
                    DisableLiveTrackingNoLock(root, false);
                }
            }
        }
        if (disableDeferred)
        {
            WriteTrackingDiagnostic(
                workspaceRoot,
                "Element tracking local disable deferred until commit boundary",
                new InvalidOperationException("A local policy notification attempted to disable tracking while observed activity had not reached a successful Save or Synchronize boundary. Existing evidence remains tracked until that boundary instead of being discarded."),
                null);
        }
    }

    public static bool IsEnabled(string workspaceRoot)
    {
        return ResolvePolicyEnabled(workspaceRoot, false);
    }

    public static void RequestDocumentSessionBaselineRefresh()
    {
        System.Threading.Interlocked.Exchange(ref _documentSessionBaselineRefreshRequested, 1);
    }

    public static bool ConsumeDocumentSessionBaselineRefreshRequest()
    {
        return System.Threading.Interlocked.Exchange(ref _documentSessionBaselineRefreshRequested, 0) == 1;
    }

    public static bool HasDocumentSession(Document doc)
    {
        if (doc == null)
        {
            return false;
        }
        lock (SyncRoot)
        {
            return Sessions.ContainsKey(BuildRuntimeKey(doc));
        }
    }

    public static bool ConsumeProjectCatalogObservationRequired(Document doc)
    {
        if (doc == null)
        {
            return true;
        }
        lock (SyncRoot)
        {
            string runtimeKey = BuildRuntimeKey(doc);
            bool required;
            bool available = ProjectCatalogObservationRequiredByRuntimeKey.TryGetValue(runtimeKey, out required);
            ProjectCatalogObservationRequiredByRuntimeKey.Remove(runtimeKey);
            return FamilyBrowserTrackingCommitOptimizationPolicy.ShouldObserveProjectCatalog(
                available,
                available && required,
                false);
        }
    }

    public static void RestoreProjectCatalogObservationRequired(Document doc)
    {
        if (doc == null)
        {
            return;
        }
        lock (SyncRoot)
        {
            SetProjectCatalogObservationDecisionNoLock(BuildRuntimeKey(doc), true);
        }
    }

    public static void RecordProjectCatalogObservationPerformance(Document doc, string commitKind, bool performed, long elapsedMilliseconds)
    {
        WriteTrackingPerformance(
            doc,
            string.IsNullOrWhiteSpace(commitKind) ? "Save" : commitKind,
            "project-catalog",
            Math.Max(0L, elapsedMilliseconds),
            "performed=" + (performed ? "yes" : "no"));
    }

    public static int GetActiveUncommittedSessionCount()
    {
        lock (SyncRoot)
        {
            return Sessions.Values.Count(HasUncommittedSessionEvidence);
        }
    }

    public static int GetProtectedRecoverySessionCount()
    {
        lock (SyncRoot)
        {
            return Sessions.Values.Count(HasProtectedRecoveryEvidence);
        }
    }

    public static bool IsManagedFolderTransitionAuthorized()
    {
        return _managedFolderTransitionAuthorizationDepth > 0;
    }

    public static IDisposable AuthorizeManagedFolderTransition()
    {
        _managedFolderTransitionAuthorizationDepth++;
        return new ManagedFolderTransitionAuthorization();
    }

    public static int GetDeferredPolicyDisableSessionCount()
    {
        lock (SyncRoot)
        {
            return Sessions.Values.Count(delegate(DocumentSession session)
            {
                return session != null && session.PolicyDisableDeferred && HasUncommittedSessionEvidence(session);
            });
        }
    }

    public static int GetUnprotectedLocalSaveSessionCount()
    {
        lock (SyncRoot)
        {
            return Sessions.Values.Count(delegate(DocumentSession session)
            {
                return session != null && session.LocalSaveCheckpointFailed && HasUncommittedSessionEvidence(session);
            });
        }
    }

    public static int GetUnprotectedCommitBoundarySessionCount()
    {
        lock (SyncRoot)
        {
            return Sessions.Values.Count(delegate(DocumentSession session)
            {
                return session != null && session.CommitBoundaryProtectionFailed && HasUncommittedSessionEvidence(session);
            });
        }
    }

    public static void NotifyManagementContextChanged()
    {
        lock (SyncRoot)
        {
            _cachedPolicyWorkspaceRoot = string.Empty;
            _cachedPolicyEnabled = false;
            _cachedPolicyKnown = false;
            _cachedPolicyAtUtc = DateTime.MinValue;
            _cachedPolicy = null;
        }
    }

    public static void StartReloadLatestBridge(object controlledApplication, Func<string> workspaceRootResolver)
    {
        lock (SyncRoot)
        {
            StopReloadLatestBridgeNoLock();
            _workspaceRootResolver = workspaceRootResolver;
            if (controlledApplication == null)
            {
                return;
            }
            try
            {
                Type sourceType = controlledApplication.GetType();
                _reloadingLatestEvent = sourceType.GetEvent("DocumentReloadingLatest", BindingFlags.Public | BindingFlags.Instance);
                _reloadedLatestEvent = sourceType.GetEvent("DocumentReloadedLatest", BindingFlags.Public | BindingFlags.Instance);
                if (_reloadingLatestEvent == null || _reloadedLatestEvent == null)
                {
                    _reloadingLatestEvent = null;
                    _reloadedLatestEvent = null;
                    return;
                }
                MethodInfo reloadingMethod = typeof(FamilyBrowserElementChangeTrackingService).GetMethod("HandleReloadingLatestBridge", BindingFlags.NonPublic | BindingFlags.Static);
                MethodInfo reloadedMethod = typeof(FamilyBrowserElementChangeTrackingService).GetMethod("HandleReloadedLatestBridge", BindingFlags.NonPublic | BindingFlags.Static);
                _reloadingLatestHandler = Delegate.CreateDelegate(_reloadingLatestEvent.EventHandlerType, reloadingMethod, true);
                _reloadedLatestHandler = Delegate.CreateDelegate(_reloadedLatestEvent.EventHandlerType, reloadedMethod, true);
                _reloadLatestEventSource = controlledApplication;
                _reloadingLatestEvent.AddEventHandler(controlledApplication, _reloadingLatestHandler);
                _reloadedLatestEvent.AddEventHandler(controlledApplication, _reloadedLatestHandler);
            }
            catch (Exception ex)
            {
                string diagnosticRoot = ResolveBridgeWorkspaceRoot();
                StopReloadLatestBridgeNoLock();
                WriteTrackingDiagnostic(diagnosticRoot, "Element tracking Reload Latest bridge failed", ex, null);
            }
        }
    }

    public static void BeginDocumentSession(string workspaceRoot, Document doc)
    {
        bool canTrack;
        Exception documentKindError;
        if (!TryCanTrackDocument(doc, out canTrack, out documentKindError))
        {
            WriteTrackingDiagnostic(workspaceRoot, "Element tracking document kind is unavailable at session start", documentKindError, doc);
            return;
        }
        if (!canTrack)
        {
            return;
        }
        string runtimeKey = BuildRuntimeKey(doc);
        string resolvedWorkspaceRoot = workspaceRoot ?? string.Empty;
        DocumentSession existingSession;
        lock (SyncRoot)
        {
            Sessions.TryGetValue(runtimeKey, out existingSession);
        }
        bool policyStateIsFallbackOrDeferred;
        bool policyEnabled = ResolveDocumentPolicyEnabledCore(resolvedWorkspaceRoot, doc, false, out policyStateIsFallbackOrDeferred);
        if (existingSession != null)
        {
            bool keepExistingSession = true;
            lock (SyncRoot)
            {
                DocumentSession current;
                if (!Sessions.TryGetValue(runtimeKey, out current) || !object.ReferenceEquals(current, existingSession))
                {
                    keepExistingSession = false;
                }
                else if (!string.Equals(current.WorkspaceRoot ?? string.Empty, resolvedWorkspaceRoot, StringComparison.OrdinalIgnoreCase))
                {
                    bool hasProtectedCheckpoint = HasProtectedRecoveryEvidence(current) || !string.IsNullOrWhiteSpace(current.CheckpointRevisionToken);
                    bool canPromoteUnresolvedRoot = policyEnabled &&
                        !hasProtectedCheckpoint &&
                        string.IsNullOrWhiteSpace(current.WorkspaceRoot) &&
                        !string.IsNullOrWhiteSpace(resolvedWorkspaceRoot);
                    if (canPromoteUnresolvedRoot)
                    {
                        current.WorkspaceRoot = resolvedWorkspaceRoot;
                        current.RecoveryOnly = false;
                    }
                    else if (!HasUncommittedSessionEvidence(current) && !hasProtectedCheckpoint)
                    {
                        Sessions.Remove(runtimeKey);
                        keepExistingSession = false;
                    }
                }
            }
            if (keepExistingSession)
            {
                FamilyBrowserTrackingPersistenceService.FlushPending(resolvedWorkspaceRoot);
                return;
            }
        }
        bool hasPendingRecoveryCheckpoint = HasPendingRecoveryCheckpoint(resolvedWorkspaceRoot, doc);
        FamilyBrowserElementTrackingSessionMode sessionMode = FamilyBrowserElementTrackingPolicyDecision.Resolve(
            policyEnabled,
            policyStateIsFallbackOrDeferred,
            false,
            false,
            hasPendingRecoveryCheckpoint);
        if (sessionMode == FamilyBrowserElementTrackingSessionMode.Ignore)
        {
            return;
        }
        DocumentSession session = CreateSession(resolvedWorkspaceRoot, doc, false);
        if (session == null)
        {
            return;
        }
        session.RecoveryOnly = sessionMode == FamilyBrowserElementTrackingSessionMode.RecoveryOnly;
        if (session.RecoveryOnly && !HasProtectedRecoveryEvidence(session))
        {
            return;
        }
        lock (SyncRoot)
        {
            if (!Sessions.ContainsKey(runtimeKey))
            {
                Sessions[runtimeKey] = session;
            }
        }
        FamilyBrowserTrackingPersistenceService.FlushPending(resolvedWorkspaceRoot);
    }

    public static void PrepareDocumentCommit(string workspaceRoot, Document doc, string commitKind)
    {
        bool canTrack;
        Exception documentKindError;
        if (!TryCanTrackDocument(doc, out canTrack, out documentKindError))
        {
            WriteTrackingDiagnostic(workspaceRoot, "Element tracking document kind is unavailable before commit", documentKindError, doc);
            return;
        }
        if (!canTrack)
        {
            return;
        }

        string resolvedWorkspaceRoot = workspaceRoot ?? string.Empty;
        bool policyStateIsFallbackOrDeferred;
        bool policyEnabled = ResolveDocumentPolicyEnabledCore(resolvedWorkspaceRoot, doc, true, out policyStateIsFallbackOrDeferred);
        if (!policyEnabled)
        {
            return;
        }

        string runtimeKey = BuildRuntimeKey(doc);
        lock (SyncRoot)
        {
            DocumentSession existing;
            if (Sessions.TryGetValue(runtimeKey, out existing) && existing != null)
            {
                if (!string.IsNullOrWhiteSpace(resolvedWorkspaceRoot) &&
                    string.IsNullOrWhiteSpace(existing.WorkspaceRoot) &&
                    !HasProtectedRecoveryEvidence(existing) &&
                    string.IsNullOrWhiteSpace(existing.CheckpointRevisionToken))
                {
                    existing.WorkspaceRoot = resolvedWorkspaceRoot;
                }
                return;
            }
        }

        DocumentSession lateSession = CreateSession(resolvedWorkspaceRoot, doc, true);
        if (lateSession == null)
        {
            return;
        }
        lock (SyncRoot)
        {
            if (!Sessions.ContainsKey(runtimeKey))
            {
                Sessions[runtimeKey] = lateSession;
            }
        }
        WriteTrackingDiagnostic(
            resolvedWorkspaceRoot,
            "Element tracking session recovered at Save boundary",
            new InvalidOperationException((string.IsNullOrWhiteSpace(commitKind) ? "Save" : commitKind) + " began while tracking was enabled but no document baseline session existed. A truthful coverage-gap record will be protected instead of silently dropping the boundary."),
            doc);
    }

    public static void HandleDocumentChanged(string workspaceRoot, DocumentChangedEventArgs e)
    {
        if (e == null)
        {
            return;
        }
        Document doc;
        try
        {
            doc = e.GetDocument();
        }
        catch (Exception ex)
        {
            int affectedSessionCount = 0;
            lock (SyncRoot)
            {
                foreach (DocumentSession affectedSession in Sessions.Values.Where(delegate(DocumentSession item)
                {
                    return item != null && string.Equals(item.WorkspaceRoot ?? string.Empty, workspaceRoot ?? string.Empty, StringComparison.OrdinalIgnoreCase);
                }))
                {
                    affectedSession.EventReadFailureCount++;
                    affectedSessionCount++;
                }
            }
            WriteTrackingDiagnostic(
                workspaceRoot,
                "Element tracking DocumentChanged document was unavailable",
                new InvalidOperationException("The DocumentChanged event could not identify its document. " + affectedSessionCount.ToString(CultureInfo.InvariantCulture) + " active session(s) were conservatively marked for coverage review.", ex),
                null);
            return;
        }
        bool canTrack;
        Exception documentKindError;
        if (!TryCanTrackDocument(doc, out canTrack, out documentKindError))
        {
            lock (SyncRoot)
            {
                DocumentSession uncertainSession;
                if (Sessions.TryGetValue(BuildRuntimeKey(doc), out uncertainSession) && uncertainSession != null)
                {
                    uncertainSession.EventReadFailureCount++;
                }
            }
            WriteTrackingDiagnostic(workspaceRoot, "Element tracking DocumentChanged document kind was unavailable", documentKindError, doc);
            return;
        }
        if (!canTrack)
        {
            EndDocumentSession(doc);
            return;
        }

        string runtimeKey = BuildRuntimeKey(doc);
        ICollection<ElementId> added;
        ICollection<ElementId> modified;
        ICollection<ElementId> deleted;
        Exception addedReadError;
        Exception modifiedReadError;
        Exception deletedReadError;
        bool addedRead = TryGetElementIds(delegate { return e.GetAddedElementIds(); }, out added, out addedReadError);
        bool modifiedRead = TryGetElementIds(delegate { return e.GetModifiedElementIds(); }, out modified, out modifiedReadError);
        bool deletedRead = TryGetElementIds(delegate { return e.GetDeletedElementIds(); }, out deleted, out deletedReadError);
        bool eventReadFailed = !addedRead || !modifiedRead || !deletedRead;
        Exception eventReadError = eventReadFailed
            ? new AggregateException("One or more DocumentChanged element-ID collections could not be read.", new[] { addedReadError, modifiedReadError, deletedReadError }.Where(delegate(Exception error) { return error != null; }))
            : null;
        bool policyStateIsFallbackOrDeferred;
        bool policyEnabled = ResolveDocumentPolicyEnabledCore(workspaceRoot, doc, false, out policyStateIsFallbackOrDeferred);
        DocumentSession session;
        lock (SyncRoot)
        {
            Sessions.TryGetValue(runtimeKey, out session);
        }
        FamilyBrowserElementTrackingSessionMode sessionMode = FamilyBrowserElementTrackingPolicyDecision.Resolve(
            policyEnabled,
            policyStateIsFallbackOrDeferred,
            session != null,
            HasUncommittedSessionEvidence(session),
            HasProtectedRecoveryEvidence(session));
        if (sessionMode == FamilyBrowserElementTrackingSessionMode.RecoveryOnly)
        {
            bool recoveryHandled = false;
            lock (SyncRoot)
            {
                DocumentSession recoverySession;
                bool unknownExternalRecoveryUpdate = IsUnknownExternalUpdateStartNoLock(workspaceRoot);
                bool externalRecoveryUpdate = SynchronizingDocumentKeys.Contains(runtimeKey) || ReloadingDocumentKeys.Contains(runtimeKey) || unknownExternalRecoveryUpdate;
                if (Sessions.TryGetValue(runtimeKey, out recoverySession) && recoverySession.RecoveryOnly && HasProtectedRecoveryEvidence(recoverySession))
                {
                    if (externalRecoveryUpdate)
                    {
                        AddIds(recoverySession.SuppressedExternalIds, added);
                        AddIds(recoverySession.SuppressedExternalIds, modified);
                        AddIds(recoverySession.SuppressedExternalIds, deleted);
                        if (eventReadFailed)
                        {
                            recoverySession.ExternalRebaseFailed = true;
                        }
                        if (unknownExternalRecoveryUpdate)
                        {
                            recoverySession.ExternalRebaseFailed = true;
                        }
                    }
                    recoveryHandled = true;
                }
            }
            if (recoveryHandled)
            {
                if (eventReadFailed)
                {
                    WriteTrackingDiagnostic(workspaceRoot, "Element tracking recovery DocumentChanged IDs were incomplete", eventReadError, doc);
                }
                return;
            }
            EndDocumentSession(doc);
            return;
        }
        if (sessionMode == FamilyBrowserElementTrackingSessionMode.Ignore)
        {
            EndDocumentSession(doc);
            return;
        }
        if (sessionMode == FamilyBrowserElementTrackingSessionMode.DeferredCommit)
        {
            lock (SyncRoot)
            {
                if (session != null)
                {
                    session.PolicyDisableDeferred = true;
                }
            }
        }
        bool capturedLate = false;
        bool externalUpdateInProgress;
        bool unknownExternalUpdateStart;
        lock (SyncRoot)
        {
            Sessions.TryGetValue(runtimeKey, out session);
            unknownExternalUpdateStart = IsUnknownExternalUpdateStartNoLock(workspaceRoot);
            externalUpdateInProgress = SynchronizingDocumentKeys.Contains(runtimeKey) || ReloadingDocumentKeys.Contains(runtimeKey) || unknownExternalUpdateStart;
            if (externalUpdateInProgress && session != null)
            {
                AddIds(session.SuppressedExternalIds, added);
                AddIds(session.SuppressedExternalIds, modified);
                AddIds(session.SuppressedExternalIds, deleted);
                if (unknownExternalUpdateStart)
                {
                    session.ExternalRebaseFailed = true;
                }
            }
        }
        if (externalUpdateInProgress)
        {
            if (eventReadFailed)
            {
                lock (SyncRoot)
                {
                    if (session != null)
                    {
                        session.ExternalRebaseFailed = true;
                    }
                }
                WriteTrackingDiagnostic(workspaceRoot, "Element tracking external DocumentChanged IDs were incomplete", eventReadError, doc);
            }
            return;
        }
        if (session == null)
        {
            session = CreateSession(workspaceRoot, doc, true);
            if (session == null)
            {
                return;
            }
            capturedLate = true;
            lock (SyncRoot)
            {
                Sessions[runtimeKey] = session;
            }
        }
        if (session.SynchronizingWithCentral || session.ReloadingLatest)
        {
            lock (SyncRoot)
            {
                AddIds(session.SuppressedExternalIds, added);
                AddIds(session.SuppressedExternalIds, modified);
                AddIds(session.SuppressedExternalIds, deleted);
                if (eventReadFailed)
                {
                    session.ExternalRebaseFailed = true;
                }
            }
            if (eventReadFailed)
            {
                WriteTrackingDiagnostic(workspaceRoot, "Element tracking external DocumentChanged IDs were incomplete", eventReadError, doc);
            }
            return;
        }

        Exception activityMetadataError;
        ChangeActivity activity = CreateActivity(e, added, modified, deleted, out activityMetadataError);
        if (eventReadFailed)
        {
            Exception recoveryError;
            bool recovered = RecoverActivityFromCurrentSnapshot(doc, session, activity, out recoveryError);
            WriteTrackingDiagnostic(
                workspaceRoot,
                recovered ? "Element tracking DocumentChanged ID gap recovered by full-state comparison" : "Element tracking DocumentChanged ID gap recovery failed",
                recovered ? eventReadError : new AggregateException(eventReadError, recoveryError ?? new InvalidOperationException("Full-state comparison did not complete.")),
                doc);
        }
        if (eventReadFailed || activityMetadataError != null)
        {
            lock (SyncRoot)
            {
                session.EventReadFailureCount++;
            }
            if (activityMetadataError != null)
            {
                WriteTrackingDiagnostic(workspaceRoot, "Element tracking DocumentChanged operation metadata was incomplete", activityMetadataError, doc);
            }
        }
        if (ShouldUseReloadLatestTransactionFallback(activity))
        {
            lock (SyncRoot)
            {
                foreach (string id in activity.AllElementIds())
                {
                    session.SuppressedExternalIds.Add(id);
                }
            }
            RebaseSessionAfterExternalUpdate(doc, session);
            return;
        }
        if (!activity.AllElementIds().Any())
        {
            return;
        }

        lock (SyncRoot)
        {
            HashSet<string> ignoredAuxiliaryElementIds = UpdateCurrentStates(session, doc, added, modified, deleted);
            FilterActivityElementIds(activity, ignoredAuxiliaryElementIds);
            RemoveIgnoredElementIdsFromSession(session, ignoredAuxiliaryElementIds);
            if (!activity.AllElementIds().Any())
            {
                return;
            }
            if (capturedLate)
            {
                foreach (string id in activity.AddedIds)
                {
                    session.Baseline.Remove(id);
                }
                foreach (string key in activity.ModifiedIds.Concat(activity.DeletedIds))
                {
                    if (!string.IsNullOrWhiteSpace(key))
                    {
                        session.UnknownPreviousStateIds.Add(key);
                    }
                }
            }
            ApplyActivity(session, activity);
        }
    }

    public static void HandleDocumentSynchronizingWithCentral(string workspaceRoot, Document doc)
    {
        if (doc == null)
        {
            HandleDocumentSynchronizationStartFailure(workspaceRoot, new InvalidOperationException("Synchronize with Central started without a readable document."));
            return;
        }
        bool canTrack;
        Exception documentKindError;
        bool documentKindKnown = TryCanTrackDocument(doc, out canTrack, out documentKindError);
        if (documentKindKnown && !canTrack)
        {
            return;
        }
        if (!documentKindKnown)
        {
            WriteTrackingDiagnostic(workspaceRoot, "Element tracking synchronization document kind was unavailable", documentKindError, doc);
        }
        bool policyEnabled = ResolveDocumentPolicyEnabled(workspaceRoot, doc, false);
        if (!policyEnabled && !HasRetainedSessionEvidence(doc) && !HasPendingRecoveryCheckpoint(workspaceRoot, doc))
        {
            return;
        }
        string runtimeKey = BuildRuntimeKey(doc);
        lock (SyncRoot)
        {
            SynchronizingDocumentKeys.Add(runtimeKey);
        }
        BeginDocumentSession(workspaceRoot, doc);
        lock (SyncRoot)
        {
            DocumentSession session;
            if (Sessions.TryGetValue(runtimeKey, out session))
            {
                session.SynchronizingWithCentral = true;
                if (!documentKindKnown)
                {
                    session.ExternalRebaseFailed = true;
                }
            }
        }
    }

    public static void HandleDocumentReloadingLatest(string workspaceRoot, Document doc)
    {
        if (doc == null)
        {
            HandleDocumentReloadLatestStartFailure(workspaceRoot, new InvalidOperationException("Reload Latest started without a readable document."));
            return;
        }
        bool canTrack;
        Exception documentKindError;
        bool documentKindKnown = TryCanTrackDocument(doc, out canTrack, out documentKindError);
        if (documentKindKnown && !canTrack)
        {
            return;
        }
        if (!documentKindKnown)
        {
            WriteTrackingDiagnostic(workspaceRoot, "Element tracking Reload Latest document kind was unavailable", documentKindError, doc);
        }
        bool policyEnabled = ResolveDocumentPolicyEnabled(workspaceRoot, doc, false);
        if (!policyEnabled && !HasRetainedSessionEvidence(doc) && !HasPendingRecoveryCheckpoint(workspaceRoot, doc))
        {
            return;
        }
        string runtimeKey = BuildRuntimeKey(doc);
        lock (SyncRoot)
        {
            ReloadingDocumentKeys.Add(runtimeKey);
        }
        BeginDocumentSession(workspaceRoot, doc);
        lock (SyncRoot)
        {
            DocumentSession session;
            if (Sessions.TryGetValue(runtimeKey, out session))
            {
                session.ReloadingLatest = true;
                if (!documentKindKnown)
                {
                    session.ExternalRebaseFailed = true;
                }
            }
        }
    }

    public static void HandleDocumentReloadedLatest(string workspaceRoot, Document doc, object status)
    {
        if (doc == null)
        {
            return;
        }
        string runtimeKey = BuildRuntimeKey(doc);
        DocumentSession session;
        lock (SyncRoot)
        {
            ReloadingDocumentKeys.Remove(runtimeKey);
            UnknownReloadLatestStartRoots.Remove(NormalizeWorkspaceRootKey(workspaceRoot));
            Sessions.TryGetValue(runtimeKey, out session);
            if (session != null)
            {
                session.ReloadingLatest = false;
            }
        }
        bool succeeded = status != null && string.Equals(status.ToString(), "Succeeded", StringComparison.OrdinalIgnoreCase);
        if (!succeeded)
        {
            bool hasSuppressedChanges;
            lock (SyncRoot)
            {
                hasSuppressedChanges = session != null && session.SuppressedExternalIds.Count > 0;
            }
            if (hasSuppressedChanges)
            {
                RebaseSessionAfterExternalUpdate(doc, session);
            }
            else
            {
                lock (SyncRoot)
                {
                    if (session != null)
                    {
                        session.SuppressedExternalIds.Clear();
                    }
                }
            }
            return;
        }
        bool policyStateIsFallbackOrDeferred;
        bool policyEnabled = ResolveDocumentPolicyEnabledCore(workspaceRoot, doc, true, out policyStateIsFallbackOrDeferred);
        bool pendingRecoveryCheckpoint = session == null && HasPendingRecoveryCheckpoint(workspaceRoot, doc);
        FamilyBrowserElementTrackingSessionMode sessionMode = FamilyBrowserElementTrackingPolicyDecision.Resolve(
            policyEnabled,
            policyStateIsFallbackOrDeferred,
            session != null,
            HasUncommittedSessionEvidence(session),
            HasProtectedRecoveryEvidence(session) || pendingRecoveryCheckpoint);
        if (sessionMode == FamilyBrowserElementTrackingSessionMode.Ignore)
        {
            EndDocumentSession(doc);
            return;
        }
        if (session == null)
        {
            BeginDocumentSession(workspaceRoot, doc);
            return;
        }
        RebaseSessionAfterExternalUpdate(doc, session);
    }

    public static bool HandleDocumentCommitted(string workspaceRoot, Document doc, object status, string commitKind)
    {
        if (doc == null)
        {
            return false;
        }
        Stopwatch completionStopwatch = Stopwatch.StartNew();
        long stateRefreshElapsedMilliseconds = 0L;
        long persistenceElapsedMilliseconds = 0L;
        long postCommitBaselineElapsedMilliseconds = 0L;
        int locallyPendingElementCount = 0;
        int incomingElementCount = 0;
        bool projectCatalogObservationRequired = true;
        string postCommitBaselineMode = "none";
        string kind = string.IsNullOrWhiteSpace(commitKind) ? "Save" : commitKind;
        bool synchronization = string.Equals(kind, "SynchronizeWithCentral", StringComparison.OrdinalIgnoreCase);
        string runtimeKey = BuildRuntimeKey(doc);
        DocumentSession session;
        lock (SyncRoot)
        {
            if (synchronization)
            {
                SynchronizingDocumentKeys.Remove(runtimeKey);
                UnknownSynchronizationStartRoots.Remove(NormalizeWorkspaceRootKey(workspaceRoot));
            }
            Sessions.TryGetValue(runtimeKey, out session);
            if (session != null && synchronization)
            {
                session.SynchronizingWithCentral = false;
            }
        }
        bool succeeded;
        try
        {
            succeeded = status != null && string.Equals(status.ToString(), "Succeeded", StringComparison.OrdinalIgnoreCase);
        }
        catch (Exception ex)
        {
            MarkCommitBoundaryProtectionFailed(session);
            if (synchronization)
            {
                CloseExternalUpdateWindowAfterUnknownCompletion(
                    workspaceRoot,
                    doc,
                    true,
                    "Element tracking synchronization completion status failed",
                    ex);
            }
            else
            {
                MarkCommitBoundaryProtectionFailed(session);
                WriteTrackingDiagnostic(workspaceRoot, "Element tracking save completion status failed", ex, doc);
            }
            return false;
        }
        if (!succeeded)
        {
            bool hasSuppressedChanges;
            lock (SyncRoot)
            {
                hasSuppressedChanges = session != null && session.SuppressedExternalIds.Count > 0;
            }
            if (hasSuppressedChanges)
            {
                RebaseSessionAfterExternalUpdate(doc, session);
            }
            else
            {
                lock (SyncRoot)
                {
                    if (session != null)
                    {
                        session.SuppressedExternalIds.Clear();
                    }
                }
            }
            return false;
        }
        bool policyUsedCachedFallback = false;
        if (session == null)
        {
            if (!synchronization)
            {
                PrepareDocumentCommit(workspaceRoot, doc, kind);
                lock (SyncRoot)
                {
                    Sessions.TryGetValue(runtimeKey, out session);
                }
            }
            if (session == null)
            {
                return false;
            }
        }
        bool policyEnabledAtCommit = ResolveDocumentPolicyEnabledCore(workspaceRoot, doc, true, out policyUsedCachedFallback);
        FamilyBrowserElementTrackingSessionMode commitSessionMode = FamilyBrowserElementTrackingPolicyDecision.Resolve(
            policyEnabledAtCommit,
            policyUsedCachedFallback,
            true,
            HasUncommittedSessionEvidence(session),
            HasProtectedRecoveryEvidence(session));
        bool recoveryOnlyCommit = commitSessionMode == FamilyBrowserElementTrackingSessionMode.RecoveryOnly &&
            string.Equals(kind, "SynchronizeWithCentral", StringComparison.OrdinalIgnoreCase);
        if (commitSessionMode == FamilyBrowserElementTrackingSessionMode.Ignore)
        {
            EndDocumentSession(doc);
            return false;
        }
        if (commitSessionMode == FamilyBrowserElementTrackingSessionMode.RecoveryOnly && !recoveryOnlyCommit)
        {
            return false;
        }
        if (commitSessionMode == FamilyBrowserElementTrackingSessionMode.DeferredCommit && !session.PolicyDisableDeferred)
        {
            policyUsedCachedFallback = true;
        }
        bool workshared;
        if (!TryGetIsWorkshared(doc, out workshared))
        {
            MarkCommitBoundaryProtectionFailed(session);
            WriteTrackingDiagnostic(workspaceRoot, "Element tracking worksharing state is unavailable at commit", new InvalidOperationException("A successful Save or Synchronize callback did not expose a trustworthy worksharing state. The session was retained and no standalone or central commit was inferred."), doc);
            return false;
        }
        if (workshared && !string.Equals(kind, "SynchronizeWithCentral", StringComparison.OrdinalIgnoreCase))
        {
            StageWorksharedLocalSaveCheckpoint(workspaceRoot, doc, session, policyUsedCachedFallback);
            return false;
        }

        string projectIdentity = SafeProjectIdentityPath(doc);
        string currentProjectStableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(projectIdentity);
        if (string.IsNullOrWhiteSpace(currentProjectStableIdentity))
        {
            MarkCommitBoundaryProtectionFailed(session);
            WriteTrackingDiagnostic(workspaceRoot, "Element tracking commit project identity is unavailable", new InvalidOperationException("A successful Save or Synchronize callback did not expose a stable project identity. The session was retained and no history was fabricated."), doc);
            return false;
        }
        string localDocumentPath = SafeLocalDocumentPath(doc);
        string revitUserName = SafeRevitUserName(doc);
        FamilyBrowserElementChangeCommit commit;
        string persistenceWorkspaceRoot;
        List<FamilyBrowserElementChangeCommit> commitsToPersist;
        int foreignRecoveredCheckpointCount;
        bool incrementalRefreshSucceeded = true;
        Exception incrementalRefreshError = null;
        FamilyBrowserPostCommitBaselineMode baselineMode = FamilyBrowserPostCommitBaselineMode.FullCapture;
        string synchronizationPublishedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
        StateCaptureContext commitCaptureContext = BuildStateCaptureContext(doc);
        lock (SyncRoot)
        {
            List<FamilyBrowserElementChangeCommit> recoveredForCurrentProject = session.RecoveredLocalSaveCommits
                .Where(delegate(FamilyBrowserElementChangeCommit recovered) { return CommitMatchesProjectStableIdentity(recovered, currentProjectStableIdentity); })
                .Select(CloneElementChangeCommit)
                .ToList();
            foreignRecoveredCheckpointCount = session.RecoveredLocalSaveCommits.Count - recoveredForCurrentProject.Count;
            RefreshProjectParameterStatesAtCommit(session, doc, commitCaptureContext);
            HashSet<string> locallyPendingIds = GetLocallyPendingElementIds(session);
            HashSet<string> incomingIds = new HashSet<string>(session.SuppressedExternalIds, StringComparer.Ordinal);
            locallyPendingElementCount = locallyPendingIds.Count;
            incomingElementCount = incomingIds.Count;
            HashSet<string> locallyAttributedIds = new HashSet<string>(locallyPendingIds, StringComparer.Ordinal);
            foreach (FamilyBrowserElementChangeCommit recoveredCommit in recoveredForCurrentProject)
            {
                foreach (FamilyBrowserElementChangeItem recoveredChange in recoveredCommit == null || recoveredCommit.Changes == null
                    ? new List<FamilyBrowserElementChangeItem>()
                    : recoveredCommit.Changes)
                {
                    if (recoveredChange != null && !string.IsNullOrWhiteSpace(recoveredChange.ElementId))
                    {
                        locallyAttributedIds.Add(recoveredChange.ElementId);
                    }
                }
            }
            session.ExternalOverlapIds.IntersectWith(locallyAttributedIds);
            foreach (string id in incomingIds)
            {
                if (locallyAttributedIds.Contains(id))
                {
                    session.ExternalOverlapIds.Add(id);
                }
            }
            HashSet<string> finalStateRefreshIds = new HashSet<string>(locallyPendingIds, StringComparer.Ordinal);
            finalStateRefreshIds.UnionWith(incomingIds);
            incrementalRefreshSucceeded = RefreshTouchedCurrentStates(
                session,
                doc,
                finalStateRefreshIds,
                commitCaptureContext,
                out stateRefreshElapsedMilliseconds,
                out incrementalRefreshError);
            if (!incrementalRefreshSucceeded)
            {
                session.CommitBoundaryReadFailureCount++;
            }
            bool evidenceGap = !incrementalRefreshSucceeded ||
                session.ExternalRebaseFailed ||
                session.EventReadFailureCount > 0 ||
                session.CommitBoundaryReadFailureCount > 0;
            bool catalogRelevantChange = HasProjectCatalogRelevantChanges(session, finalStateRefreshIds, recoveredForCurrentProject);
            projectCatalogObservationRequired = FamilyBrowserTrackingCommitOptimizationPolicy.ShouldObserveProjectCatalog(
                true,
                catalogRelevantChange,
                evidenceGap);
            SetProjectCatalogObservationDecisionNoLock(runtimeKey, projectCatalogObservationRequired);
            baselineMode = FamilyBrowserTrackingCommitOptimizationPolicy.ResolveBaselineMode(
                incrementalRefreshSucceeded,
                session.ExternalRebaseFailed,
                session.EventReadFailureCount,
                session.CommitBoundaryReadFailureCount);
            session.SuppressedExternalIds.Clear();
            commit = session.RecoveryOnly ? null : BuildCommit(doc, session, kind, policyUsedCachedFallback, workshared);
            persistenceWorkspaceRoot = string.IsNullOrWhiteSpace(session.WorkspaceRoot) ? workspaceRoot : session.WorkspaceRoot;
            foreach (FamilyBrowserElementChangeCommit recoveredCommit in recoveredForCurrentProject)
            {
                if (recoveredCommit == null)
                {
                    continue;
                }
                if (string.IsNullOrWhiteSpace(recoveredCommit.PublishedAtUtc))
                {
                    string localSaveAtUtc = string.IsNullOrWhiteSpace(recoveredCommit.LocalSaveProtectedAtUtc)
                        ? recoveredCommit.CommittedAtUtc ?? string.Empty
                        : recoveredCommit.LocalSaveProtectedAtUtc;
                    recoveredCommit.LocalSaveProtectedAtUtc = localSaveAtUtc;
                    recoveredCommit.PublishedAtUtc = synchronizationPublishedAtUtc;
                    recoveredCommit.CommittedAtUtc = synchronizationPublishedAtUtc;
                    recoveredCommit.CoverageNote = (recoveredCommit.CoverageNote ?? string.Empty)
                        + " Local Save protected at " + localSaveAtUtc
                        + "; published after successful Synchronize with Central at " + synchronizationPublishedAtUtc + ".";
                    recoveredCommit.IntegritySha256 = string.Empty;
                }
                bool recoveredOverlap = false;
                foreach (FamilyBrowserElementChangeItem recoveredChange in recoveredCommit.Changes ?? new List<FamilyBrowserElementChangeItem>())
                {
                    if (recoveredChange != null && session.ExternalOverlapIds.Contains(recoveredChange.ElementId ?? string.Empty))
                    {
                        recoveredChange.ExternalUpdateOverlap = true;
                        recoveredOverlap = true;
                    }
                }
                if (recoveredOverlap)
                {
                    recoveredCommit.ExternalUpdateOverlapCount = (recoveredCommit.Changes ?? new List<FamilyBrowserElementChangeItem>()).Count(delegate(FamilyBrowserElementChangeItem item) { return item != null && item.ExternalUpdateOverlap; });
                    recoveredCommit.AttributionConfidence = "ClientObservedWithExternalOverlap";
                    recoveredCommit.CoverageNote = (recoveredCommit.CoverageNote ?? string.Empty) + " A recovered local-save change also appeared in an incoming central/reload update; authorship is mixed or uncertain.";
                    recoveredCommit.IntegritySha256 = string.Empty;
                }
                if (session.ExternalRebaseFailed)
                {
                    recoveredCommit.AttributionConfidence = recoveredOverlap ? "ClientObservedWithExternalOverlap" : "ClientObservedWithExternalRebaseGap";
                    recoveredCommit.CoverageNote = (recoveredCommit.CoverageNote ?? string.Empty) + " An incoming central/reload update could not be fully rebased before publication; exact intervening authorship requires review.";
                    recoveredCommit.IntegritySha256 = string.Empty;
                }
            }
            commitsToPersist = recoveredForCurrentProject
                .Where(delegate(FamilyBrowserElementChangeCommit recovered) { return recovered != null; })
                .Concat(commit == null ? Enumerable.Empty<FamilyBrowserElementChangeCommit>() : new[] { commit })
                .ToList();
        }
        if (!incrementalRefreshSucceeded)
        {
            WriteTrackingDiagnostic(
                persistenceWorkspaceRoot,
                "Element tracking incremental post-commit refresh failed",
                incrementalRefreshError ?? new InvalidOperationException("The changed-element state refresh did not complete; the conservative full-model baseline fallback will be used."),
                doc);
        }
        if (foreignRecoveredCheckpointCount > 0)
        {
            WriteTrackingDiagnostic(persistenceWorkspaceRoot, "Element tracking recovered checkpoint project mismatch", new InvalidDataException(foreignRecoveredCheckpointCount.ToString(CultureInfo.InvariantCulture) + " recovered local-save checkpoint commit(s) belong to a different project identity. They were not published by this project synchronization and remain protected under their original checkpoint identity."), doc);
        }
        string previousCheckpointLocalPath;
        string previousCheckpointProjectIdentity;
        string previousCheckpointRevitUserName;
        string previousCheckpointRevisionToken;
        lock (SyncRoot)
        {
            previousCheckpointLocalPath = session.CheckpointLocalDocumentPath;
            previousCheckpointProjectIdentity = session.CheckpointProjectIdentityPath;
            previousCheckpointRevitUserName = session.CheckpointRevitUserName;
            previousCheckpointRevisionToken = session.CheckpointRevisionToken;
        }
        bool sameCheckpointProject = SameProjectIdentity(projectIdentity, previousCheckpointProjectIdentity);
        bool sameCheckpointIdentity = SameCheckpointIdentity(projectIdentity, localDocumentPath, revitUserName, previousCheckpointProjectIdentity, previousCheckpointLocalPath, previousCheckpointRevitUserName);
        string expectedCheckpointRevisionToken = sameCheckpointIdentity ? previousCheckpointRevisionToken : string.Empty;
        string finalizedCheckpointRevisionToken = expectedCheckpointRevisionToken;
        bool checkpointFinalized;
        Stopwatch persistenceStopwatch = Stopwatch.StartNew();
        try
        {
            checkpointFinalized = commitsToPersist.Count == 0 || FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint(
                persistenceWorkspaceRoot,
                projectIdentity,
                localDocumentPath,
                revitUserName,
                commitsToPersist,
                true,
                expectedCheckpointRevisionToken,
                out finalizedCheckpointRevisionToken);
        }
        catch (Exception checkpointFinalizationError)
        {
            MarkCommitBoundaryProtectionFailed(session);
            WriteTrackingDiagnostic(persistenceWorkspaceRoot, "Element tracking synchronization checkpoint finalization threw an exception", checkpointFinalizationError, doc);
            return false;
        }
        if (!checkpointFinalized)
        {
            MarkCommitBoundaryProtectionFailed(session);
            WriteTrackingDiagnostic(persistenceWorkspaceRoot, "Element tracking synchronization checkpoint finalization failed", new IOException("The successful synchronization could not update its local write-ahead checkpoint before immutable history persistence."), doc);
            return false;
        }
        if (workshared)
        {
            lock (SyncRoot)
            {
                session.LocalSaveCheckpointFailed = false;
                session.CommitBoundaryProtectionFailed = false;
            }
        }
        bool durable;
        try
        {
            durable = commitsToPersist.Count == 0 || FamilyBrowserTrackingPersistenceService.PersistElementChangeCommitsDeferred(persistenceWorkspaceRoot, commitsToPersist);
        }
        catch (Exception persistenceError)
        {
            if (!workshared)
            {
                MarkCommitBoundaryProtectionFailed(session);
            }
            WriteTrackingDiagnostic(persistenceWorkspaceRoot, "Element change history persistence threw an exception", persistenceError, doc);
            return false;
        }
        if (!durable)
        {
            if (!workshared)
            {
                MarkCommitBoundaryProtectionFailed(session);
            }
            WriteTrackingDiagnostic(persistenceWorkspaceRoot, "Element change history persistence failed", new IOException("Neither the managed history nor the local write-ahead spool could preserve the element change commit."), doc);
            return false;
        }
        bool currentCheckpointCleaned = FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(
            projectIdentity,
            localDocumentPath,
            revitUserName,
            finalizedCheckpointRevisionToken);
        if (!currentCheckpointCleaned)
        {
            WriteTrackingDiagnostic(persistenceWorkspaceRoot, "Element tracking synchronized checkpoint cleanup failed", new IOException("Immutable history is durable, but the finalized local checkpoint could not be removed. A later replay is idempotent."), doc);
            EndDocumentSession(doc);
            return commitsToPersist.Count > 0;
        }
        if (!sameCheckpointIdentity && sameCheckpointProject && !string.IsNullOrWhiteSpace(previousCheckpointRevisionToken) &&
            !FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(
                previousCheckpointProjectIdentity,
                previousCheckpointLocalPath,
                previousCheckpointRevitUserName,
                previousCheckpointRevisionToken))
        {
            WriteTrackingDiagnostic(persistenceWorkspaceRoot, "Element tracking previous-path checkpoint cleanup failed", new IOException("Immutable history is durable, but a checkpoint from the previous local-file path could not be removed. Its stable entry IDs prevent duplicate history on replay."), doc);
        }
        if (!sameCheckpointProject && !string.IsNullOrWhiteSpace(previousCheckpointRevisionToken))
        {
            WriteTrackingDiagnostic(persistenceWorkspaceRoot, "Element tracking previous-project checkpoint retained", new InvalidOperationException("The project or central-model identity changed. The previous project's unsynchronized checkpoint was intentionally retained and was not published or deleted by the new project."), doc);
        }
        persistenceStopwatch.Stop();
        persistenceElapsedMilliseconds = persistenceStopwatch.ElapsedMilliseconds;
        if (recoveryOnlyCommit)
        {
            EndDocumentSession(doc);
            completionStopwatch.Stop();
            WriteTrackingPerformance(
                doc,
                kind,
                "commit",
                completionStopwatch.ElapsedMilliseconds,
                BuildCommitPerformanceDetail(
                    postCommitBaselineMode,
                    stateRefreshElapsedMilliseconds,
                    persistenceElapsedMilliseconds,
                    postCommitBaselineElapsedMilliseconds,
                    locallyPendingElementCount,
                    incomingElementCount,
                    projectCatalogObservationRequired));
            return commitsToPersist.Count > 0;
        }

        if (baselineMode == FamilyBrowserPostCommitBaselineMode.Incremental)
        {
            postCommitBaselineMode = "incremental";
            Stopwatch promotionStopwatch = Stopwatch.StartNew();
            lock (SyncRoot)
            {
                DocumentSession current;
                if (Sessions.TryGetValue(runtimeKey, out current) && object.ReferenceEquals(current, session))
                {
                    current.WorkspaceRoot = workspaceRoot ?? string.Empty;
                    current.PromoteCurrentToBaseline(stateRefreshElapsedMilliseconds);
                    current.RecoveredLocalSaveCommits.Clear();
                    current.CheckpointProjectIdentityPath = projectIdentity;
                    current.CheckpointLocalDocumentPath = localDocumentPath;
                    current.CheckpointRevitUserName = revitUserName;
                    current.CheckpointRevisionToken = string.Empty;
                }
            }
            promotionStopwatch.Stop();
            postCommitBaselineElapsedMilliseconds = promotionStopwatch.ElapsedMilliseconds;
        }
        else
        {
            postCommitBaselineMode = "full-fallback";
            Dictionary<string, FamilyBrowserTrackedElementState> refreshed;
            HashSet<string> ignoredAuxiliaryElementIds;
            long elapsed;
            Exception captureError;
            if (!CaptureBaseline(doc, out refreshed, out ignoredAuxiliaryElementIds, out elapsed, out captureError))
            {
                WriteTrackingDiagnostic(workspaceRoot, "Element tracking post-commit baseline failed", captureError ?? new InvalidOperationException("The post-commit element baseline could not be captured."), doc);
                EndDocumentSession(doc);
                completionStopwatch.Stop();
                WriteTrackingPerformance(
                    doc,
                    kind,
                    "commit",
                    completionStopwatch.ElapsedMilliseconds,
                    BuildCommitPerformanceDetail(
                        postCommitBaselineMode + "-failed",
                        stateRefreshElapsedMilliseconds,
                        persistenceElapsedMilliseconds,
                        elapsed,
                        locallyPendingElementCount,
                        incomingElementCount,
                        projectCatalogObservationRequired));
                return commit != null;
            }
            postCommitBaselineElapsedMilliseconds = elapsed;
            lock (SyncRoot)
            {
                DocumentSession current;
                if (Sessions.TryGetValue(runtimeKey, out current) && object.ReferenceEquals(current, session))
                {
                    current.WorkspaceRoot = workspaceRoot ?? string.Empty;
                    current.ResetBaseline(refreshed, ignoredAuxiliaryElementIds, elapsed, false);
                    current.RecoveredLocalSaveCommits.Clear();
                    current.CheckpointProjectIdentityPath = projectIdentity;
                    current.CheckpointLocalDocumentPath = localDocumentPath;
                    current.CheckpointRevitUserName = revitUserName;
                    current.CheckpointRevisionToken = string.Empty;
                }
            }
        }
        completionStopwatch.Stop();
        WriteTrackingPerformance(
            doc,
            kind,
            "commit",
            completionStopwatch.ElapsedMilliseconds,
            BuildCommitPerformanceDetail(
                postCommitBaselineMode,
                stateRefreshElapsedMilliseconds,
                persistenceElapsedMilliseconds,
                postCommitBaselineElapsedMilliseconds,
                locallyPendingElementCount,
                incomingElementCount,
                projectCatalogObservationRequired));
        return commitsToPersist.Count > 0;
    }

    public static void HandleDocumentSynchronizationCompletionFailure(string workspaceRoot, Document doc, Exception error)
    {
        CloseExternalUpdateWindowAfterUnknownCompletion(
            workspaceRoot,
            doc,
            true,
            "Element tracking synchronization completion status failed",
            error ?? new InvalidOperationException("Synchronize with Central completed without a readable document or status."));
    }

    public static void HandleDocumentSynchronizationStartFailure(string workspaceRoot, Exception error)
    {
        MarkUnknownExternalUpdateStart(
            workspaceRoot,
            true,
            "Element tracking synchronization start document failed",
            error ?? new InvalidOperationException("Synchronize with Central started without a readable document. Incoming changes are suppressed conservatively until completion."));
    }

    public static void HandleDocumentReloadLatestStartFailure(string workspaceRoot, Exception error)
    {
        MarkUnknownExternalUpdateStart(
            workspaceRoot,
            false,
            "Element tracking Reload Latest start document failed",
            error ?? new InvalidOperationException("Reload Latest started without a readable document. Incoming changes are suppressed conservatively until completion."));
    }

    public static void HandleDocumentSaveCompletionFailure(string workspaceRoot, Document doc, string commitKind, Exception error)
    {
        string root = workspaceRoot ?? string.Empty;
        int affectedSessionCount = 0;
        lock (SyncRoot)
        {
            if (doc != null)
            {
                DocumentSession session;
                if (Sessions.TryGetValue(BuildRuntimeKey(doc), out session) && session != null)
                {
                    session.CommitBoundaryReadFailureCount++;
                    session.CommitBoundaryProtectionFailed = HasUncommittedSessionEvidence(session);
                    affectedSessionCount = 1;
                }
            }
            else
            {
                foreach (DocumentSession session in Sessions.Values.Where(delegate(DocumentSession item)
                {
                    return item != null && string.Equals(item.WorkspaceRoot ?? string.Empty, root, StringComparison.OrdinalIgnoreCase);
                }))
                {
                    session.CommitBoundaryReadFailureCount++;
                    session.CommitBoundaryProtectionFailed = HasUncommittedSessionEvidence(session);
                    affectedSessionCount++;
                }
            }
        }
        string kind = string.IsNullOrWhiteSpace(commitKind) ? "Save" : commitKind;
        WriteTrackingDiagnostic(
            root,
            "Element tracking " + kind + " completion boundary was unreadable",
            error ?? new InvalidOperationException(kind + " completed without a readable document or status. " + affectedSessionCount.ToString(CultureInfo.InvariantCulture) + " active session(s) were marked for boundary review."),
            doc);
    }

    private static void CloseExternalUpdateWindowAfterUnknownCompletion(
        string workspaceRoot,
        Document doc,
        bool synchronization,
        string caption,
        Exception error)
    {
        DocumentSession sessionToRebase = null;
        bool shouldRebase = false;
        lock (SyncRoot)
        {
            HashSet<string> activeKeys = synchronization ? SynchronizingDocumentKeys : ReloadingDocumentKeys;
            if (synchronization)
            {
                UnknownSynchronizationStartRoots.Remove(NormalizeWorkspaceRootKey(workspaceRoot));
            }
            else
            {
                UnknownReloadLatestStartRoots.Remove(NormalizeWorkspaceRootKey(workspaceRoot));
            }
            List<string> affectedKeys;
            if (doc == null)
            {
                affectedKeys = activeKeys.ToList();
            }
            else
            {
                affectedKeys = new List<string> { BuildRuntimeKey(doc) };
            }
            foreach (string runtimeKey in affectedKeys)
            {
                activeKeys.Remove(runtimeKey);
                DocumentSession session;
                if (!Sessions.TryGetValue(runtimeKey, out session) || session == null)
                {
                    continue;
                }
                if (synchronization)
                {
                    session.SynchronizingWithCentral = false;
                    session.CommitBoundaryProtectionFailed = HasUncommittedSessionEvidence(session);
                }
                else
                {
                    session.ReloadingLatest = false;
                }
                session.ExternalRebaseFailed = true;
                if (doc != null && session.SuppressedExternalIds.Count > 0)
                {
                    sessionToRebase = session;
                    shouldRebase = true;
                }
            }
        }
        WriteTrackingDiagnostic(workspaceRoot, caption, error, doc);
        if (shouldRebase && sessionToRebase != null)
        {
            RebaseSessionAfterExternalUpdate(doc, sessionToRebase);
        }
    }

    private static void MarkUnknownExternalUpdateStart(string workspaceRoot, bool synchronization, string caption, Exception error)
    {
        string rootKey = NormalizeWorkspaceRootKey(workspaceRoot);
        lock (SyncRoot)
        {
            HashSet<string> roots = synchronization ? UnknownSynchronizationStartRoots : UnknownReloadLatestStartRoots;
            roots.Add(rootKey);
        }
        WriteTrackingDiagnostic(workspaceRoot, caption, error, null);
    }

    private static bool IsUnknownExternalUpdateStartNoLock(string workspaceRoot)
    {
        string rootKey = NormalizeWorkspaceRootKey(workspaceRoot);
        return UnknownSynchronizationStartRoots.Contains(rootKey) || UnknownReloadLatestStartRoots.Contains(rootKey);
    }

    private static string NormalizeWorkspaceRootKey(string workspaceRoot)
    {
        string value = (workspaceRoot ?? string.Empty).Trim();
        try
        {
            value = Path.GetFullPath(value);
        }
        catch
        {
        }
        return value.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }

    private static void StageWorksharedLocalSaveCheckpoint(
        string workspaceRoot,
        Document doc,
        DocumentSession session,
        bool policyUsedCachedFallback)
    {
        Stopwatch saveStopwatch = Stopwatch.StartNew();
        long refreshElapsedMilliseconds = 0L;
        int pendingElementCount = 0;
        bool catalogObservationRequired = true;
        string projectIdentity = SafeProjectIdentityPath(doc);
        string currentProjectStableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(projectIdentity);
        if (string.IsNullOrWhiteSpace(currentProjectStableIdentity))
        {
            lock (SyncRoot)
            {
                session.LocalSaveCheckpointFailed = HasUncommittedSessionEvidence(session);
            }
            WriteTrackingDiagnostic(workspaceRoot, "Element tracking local-save project identity is unavailable", new InvalidOperationException("The workshared local Save did not expose a stable central project identity. The in-memory session was retained and no checkpoint was written under an ambiguous project."), doc);
            return;
        }
        string localDocumentPath = SafeLocalDocumentPath(doc);
        string revitUserName = SafeRevitUserName(doc);
        FamilyBrowserElementChangeCommit currentCommit;
        List<FamilyBrowserElementChangeCommit> staged;
        string persistenceWorkspaceRoot;
        string previousCheckpointProjectIdentity;
        string previousCheckpointLocalPath;
        string previousCheckpointRevitUserName;
        string previousCheckpointRevisionToken;
        int foreignRecoveredCheckpointCount;
        StateCaptureContext commitCaptureContext = BuildStateCaptureContext(doc);
        lock (SyncRoot)
        {
            RefreshProjectParameterStatesAtCommit(session, doc, commitCaptureContext);
            HashSet<string> locallyPendingIds = GetLocallyPendingElementIds(session);
            pendingElementCount = locallyPendingIds.Count;
            Exception refreshError;
            bool refreshSucceeded = RefreshTouchedCurrentStates(
                session,
                doc,
                locallyPendingIds,
                commitCaptureContext,
                out refreshElapsedMilliseconds,
                out refreshError);
            if (!refreshSucceeded)
            {
                session.CommitBoundaryReadFailureCount++;
                WriteTrackingDiagnostic(
                    workspaceRoot,
                    "Element tracking local-save changed-state refresh failed",
                    refreshError ?? new InvalidOperationException("The local-save changed-element refresh did not complete."),
                    doc);
            }
            bool evidenceGap = !refreshSucceeded || session.EventReadFailureCount > 0 || session.CommitBoundaryReadFailureCount > 0;
            bool catalogRelevantChange = HasProjectCatalogRelevantChanges(
                session,
                locallyPendingIds,
                session.RecoveredLocalSaveCommits);
            catalogObservationRequired = FamilyBrowserTrackingCommitOptimizationPolicy.ShouldObserveProjectCatalog(
                true,
                catalogRelevantChange,
                evidenceGap);
            SetProjectCatalogObservationDecisionNoLock(BuildRuntimeKey(doc), catalogObservationRequired);
            currentCommit = BuildCommit(doc, session, "WorksharedLocalSavePendingSync", policyUsedCachedFallback, true);
            if (currentCommit != null)
            {
                currentCommit.CoverageNote += " This activity was protected after a successful local Save and is published to managed immutable history only after a later successful Synchronize with Central.";
            }
            List<FamilyBrowserElementChangeCommit> recoveredForCurrentProject = session.RecoveredLocalSaveCommits
                .Where(delegate(FamilyBrowserElementChangeCommit recovered) { return CommitMatchesProjectStableIdentity(recovered, currentProjectStableIdentity); })
                .ToList();
            foreignRecoveredCheckpointCount = session.RecoveredLocalSaveCommits.Count - recoveredForCurrentProject.Count;
            staged = recoveredForCurrentProject
                .Where(delegate(FamilyBrowserElementChangeCommit recovered) { return recovered != null; })
                .Concat(currentCommit == null ? Enumerable.Empty<FamilyBrowserElementChangeCommit>() : new[] { currentCommit })
                .ToList();
            persistenceWorkspaceRoot = string.IsNullOrWhiteSpace(session.WorkspaceRoot) ? workspaceRoot : session.WorkspaceRoot;
            previousCheckpointProjectIdentity = session.CheckpointProjectIdentityPath;
            previousCheckpointLocalPath = session.CheckpointLocalDocumentPath;
            previousCheckpointRevitUserName = session.CheckpointRevitUserName;
            previousCheckpointRevisionToken = session.CheckpointRevisionToken;
        }
        if (foreignRecoveredCheckpointCount > 0)
        {
            WriteTrackingDiagnostic(persistenceWorkspaceRoot, "Element tracking local-save checkpoint project mismatch", new InvalidDataException(foreignRecoveredCheckpointCount.ToString(CultureInfo.InvariantCulture) + " recovered checkpoint commit(s) belong to a different project identity and were not copied into this project's local-save checkpoint."), doc);
        }
        bool sameCheckpointProject = SameProjectIdentity(projectIdentity, previousCheckpointProjectIdentity);
        bool sameCheckpointIdentity = SameCheckpointIdentity(projectIdentity, localDocumentPath, revitUserName, previousCheckpointProjectIdentity, previousCheckpointLocalPath, previousCheckpointRevitUserName);
        string expectedCheckpointRevisionToken = sameCheckpointIdentity ? previousCheckpointRevisionToken : string.Empty;
        string savedCheckpointRevisionToken;
        bool protectedLocally;
        try
        {
            protectedLocally = FamilyBrowserTrackingPersistenceService.SaveElementSessionCheckpoint(
                persistenceWorkspaceRoot,
                projectIdentity,
                localDocumentPath,
                revitUserName,
                staged,
                false,
                expectedCheckpointRevisionToken,
                out savedCheckpointRevisionToken);
        }
        catch (Exception checkpointError)
        {
            lock (SyncRoot)
            {
                session.LocalSaveCheckpointFailed = currentCommit != null && HasUncommittedSessionEvidence(session);
            }
            WriteTrackingDiagnostic(persistenceWorkspaceRoot, "Element tracking workshared local-save checkpoint threw an exception", checkpointError, doc);
            return;
        }
        if (!protectedLocally)
        {
            lock (SyncRoot)
            {
                session.LocalSaveCheckpointFailed = currentCommit != null && HasUncommittedSessionEvidence(session);
            }
            WriteTrackingDiagnostic(persistenceWorkspaceRoot, "Element tracking workshared local-save checkpoint failed", new IOException("Pending workshared element activity could not be protected across a Revit restart. Keep the current Revit session open and synchronize before closing."), doc);
            return;
        }
        lock (SyncRoot)
        {
            session.PromoteCurrentToBaseline(refreshElapsedMilliseconds);
            session.RecoveredLocalSaveCommits.Clear();
            session.RecoveredLocalSaveCommits.AddRange(staged.Where(delegate(FamilyBrowserElementChangeCommit stagedCommit) { return stagedCommit != null; }));
            session.CheckpointProjectIdentityPath = projectIdentity;
            session.CheckpointLocalDocumentPath = localDocumentPath;
            session.CheckpointRevitUserName = revitUserName;
            session.CheckpointRevisionToken = savedCheckpointRevisionToken;
        }
        if (!sameCheckpointIdentity && sameCheckpointProject && !string.IsNullOrWhiteSpace(previousCheckpointRevisionToken) &&
            !FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(
                previousCheckpointProjectIdentity,
                previousCheckpointLocalPath,
                previousCheckpointRevitUserName,
                previousCheckpointRevisionToken))
        {
            WriteTrackingDiagnostic(persistenceWorkspaceRoot, "Element tracking previous-path local-save checkpoint cleanup failed", new IOException("The new local-file checkpoint is protected, but the checkpoint under the previous local-file path could not be removed."), doc);
        }
        if (!sameCheckpointProject && !string.IsNullOrWhiteSpace(previousCheckpointRevisionToken))
        {
            WriteTrackingDiagnostic(persistenceWorkspaceRoot, "Element tracking previous-project local-save checkpoint retained", new InvalidOperationException("The project or central-model identity changed. The previous project's unsynchronized checkpoint remains protected under its original identity."), doc);
        }
        saveStopwatch.Stop();
        WriteTrackingPerformance(
            doc,
            "WorksharedLocalSave",
            "commit",
            saveStopwatch.ElapsedMilliseconds,
            BuildCommitPerformanceDetail(
                "incremental",
                refreshElapsedMilliseconds,
                Math.Max(0L, saveStopwatch.ElapsedMilliseconds - refreshElapsedMilliseconds),
                0L,
                pendingElementCount,
                0,
                catalogObservationRequired));
    }

    public static void EndDocumentSession(Document doc)
    {
        if (doc == null)
        {
            return;
        }
        lock (SyncRoot)
        {
            string runtimeKey = BuildRuntimeKey(doc);
            Sessions.Remove(runtimeKey);
            SynchronizingDocumentKeys.Remove(runtimeKey);
            ReloadingDocumentKeys.Remove(runtimeKey);
            ProjectCatalogObservationRequiredByRuntimeKey.Remove(runtimeKey);
            foreach (int documentId in ClosingSessionKeysByDocumentId
                .Where(pair => string.Equals(pair.Value, runtimeKey, StringComparison.Ordinal))
                .Select(pair => pair.Key)
                .ToList())
            {
                ClosingSessionKeysByDocumentId.Remove(documentId);
            }
        }
    }

    public static void HandleDocumentClosing(Document doc, int documentId)
    {
        if (doc == null || documentId < 0)
        {
            return;
        }
        string runtimeKey = BuildRuntimeKey(doc);
        lock (SyncRoot)
        {
            ClosingSessionKeysByDocumentId.Remove(documentId);
            if (Sessions.ContainsKey(runtimeKey) || SynchronizingDocumentKeys.Contains(runtimeKey) || ReloadingDocumentKeys.Contains(runtimeKey))
            {
                ClosingSessionKeysByDocumentId[documentId] = runtimeKey;
            }
        }
    }

    public static void HandleDocumentClosed(int documentId)
    {
        if (documentId < 0)
        {
            return;
        }
        lock (SyncRoot)
        {
            string runtimeKey;
            if (ClosingSessionKeysByDocumentId.TryGetValue(documentId, out runtimeKey))
            {
                ClosingSessionKeysByDocumentId.Remove(documentId);
                Sessions.Remove(runtimeKey);
                SynchronizingDocumentKeys.Remove(runtimeKey);
                ReloadingDocumentKeys.Remove(runtimeKey);
                ProjectCatalogObservationRequiredByRuntimeKey.Remove(runtimeKey);
            }
        }
    }

    public static void Stop()
    {
        lock (SyncRoot)
        {
            StopReloadLatestBridgeNoLock();
            Sessions.Clear();
            ClosingSessionKeysByDocumentId.Clear();
            SynchronizingDocumentKeys.Clear();
            ReloadingDocumentKeys.Clear();
            UnknownSynchronizationStartRoots.Clear();
            UnknownReloadLatestStartRoots.Clear();
            ProjectCatalogObservationRequiredByRuntimeKey.Clear();
            _cachedPolicyWorkspaceRoot = string.Empty;
            _cachedPolicyEnabled = false;
            _cachedPolicyKnown = false;
            _cachedPolicyAtUtc = DateTime.MinValue;
            _cachedPolicy = null;
            LastDiagnosticUtcByKey.Clear();
            System.Threading.Interlocked.Exchange(ref _documentSessionBaselineRefreshRequested, 0);
        }
    }

    private static DocumentSession CreateSession(string workspaceRoot, Document doc, bool capturedLate)
    {
        bool workshared;
        if (!TryGetIsWorkshared(doc, out workshared))
        {
            WriteTrackingDiagnostic(workspaceRoot, "Element tracking worksharing state is unavailable at session start", new InvalidOperationException("The document worksharing state could not be read. Tracking did not start from an inferred standalone state."), doc);
            return null;
        }
        Dictionary<string, FamilyBrowserTrackedElementState> baseline;
        HashSet<string> ignoredAuxiliaryElementIds;
        long elapsed;
        Exception captureError;
        if (!CaptureBaseline(doc, out baseline, out ignoredAuxiliaryElementIds, out elapsed, out captureError))
        {
            WriteTrackingDiagnostic(workspaceRoot, "Element tracking initial baseline failed", captureError ?? new InvalidOperationException("The initial element baseline could not be captured."), doc);
            return null;
        }
        DocumentSession session = new DocumentSession
        {
            RuntimeKey = BuildRuntimeKey(doc),
            WorkspaceRoot = workspaceRoot ?? string.Empty,
            CheckpointProjectIdentityPath = SafeProjectIdentityPath(doc),
            CheckpointLocalDocumentPath = SafeLocalDocumentPath(doc),
            CheckpointRevitUserName = SafeRevitUserName(doc)
        };
        session.ResetBaseline(baseline, ignoredAuxiliaryElementIds, elapsed, capturedLate || SafeIsModified(doc));
        if (workshared)
        {
            string projectIdentity = SafeProjectIdentityPath(doc);
            string localDocumentPath = SafeLocalDocumentPath(doc);
            string revitUserName = SafeRevitUserName(doc);
            FamilyBrowserElementSessionCheckpointLoadResult checkpointResult = FamilyBrowserTrackingPersistenceService.LoadElementSessionCheckpoint(
                workspaceRoot,
                projectIdentity,
                localDocumentPath,
                revitUserName);
            if (checkpointResult.LockUnavailable)
            {
                WriteTrackingDiagnostic(
                    workspaceRoot,
                    "Element tracking local checkpoint is busy",
                    new IOException("Another Revit process is updating the local worksharing checkpoint. Tracking did not start from an unverified checkpoint state."),
                    doc);
                return null;
            }
            if (checkpointResult.Invalid || checkpointResult.DestinationMismatch)
            {
                WriteTrackingDiagnostic(
                    workspaceRoot,
                    checkpointResult.Invalid ? "Element tracking local checkpoint is invalid" : "Element tracking local checkpoint belongs to another management folder",
                    new InvalidDataException(checkpointResult.Invalid
                        ? "The saved local worksharing checkpoint failed identity or checksum validation."
                        : "The saved local worksharing checkpoint is bound to another management destination and was not loaded."),
                    doc);
            }
            else if (checkpointResult.Checkpoint != null)
            {
                session.CheckpointRevisionToken = FamilyBrowserTrackingPersistenceService.GetElementSessionCheckpointRevisionToken(checkpointResult.Checkpoint);
                List<FamilyBrowserElementChangeCommit> recovered = checkpointResult.Checkpoint.Commits ?? new List<FamilyBrowserElementChangeCommit>();
                if (checkpointResult.Checkpoint.SynchronizationSucceeded)
                {
                    bool replayed = FamilyBrowserTrackingPersistenceService.PersistElementChangeCommits(workspaceRoot, recovered);
                    if (replayed)
                    {
                        if (FamilyBrowserTrackingPersistenceService.DeleteElementSessionCheckpoint(
                            projectIdentity,
                            localDocumentPath,
                            revitUserName,
                            session.CheckpointRevisionToken))
                        {
                            session.CheckpointRevisionToken = string.Empty;
                        }
                    }
                    else
                    {
                        WriteTrackingDiagnostic(workspaceRoot, "Element tracking finalized local checkpoint replay failed", new IOException("A checkpoint from a successful synchronization remains protected locally because it could not be promoted to immutable history."), doc);
                        session.RecoveredLocalSaveCommits.AddRange(recovered.Where(delegate(FamilyBrowserElementChangeCommit commit) { return commit != null; }));
                        if (session.RecoveredLocalSaveCommits.Count > 0)
                        {
                            session.BaselineCapturedLate = true;
                        }
                    }
                }
                else
                {
                    session.RecoveredLocalSaveCommits.AddRange(recovered.Where(delegate(FamilyBrowserElementChangeCommit commit) { return commit != null; }));
                    if (session.RecoveredLocalSaveCommits.Count > 0)
                    {
                        session.BaselineCapturedLate = true;
                    }
                }
            }
        }
        return session;
    }

    private static bool CaptureBaseline(
        Document doc,
        out Dictionary<string, FamilyBrowserTrackedElementState> states,
        out HashSet<string> ignoredAuxiliaryElementIds,
        out long elapsedMilliseconds,
        out Exception captureError)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        states = new Dictionary<string, FamilyBrowserTrackedElementState>(StringComparer.Ordinal);
        ignoredAuxiliaryElementIds = new HashSet<string>(StringComparer.Ordinal);
        captureError = null;
        if (!CanTrackDocument(doc))
        {
            elapsedMilliseconds = 0L;
            return false;
        }
        try
        {
            StateCaptureContext captureContext = BuildStateCaptureContext(doc);
            ElementFilter allElementKinds = new LogicalOrFilter(new ElementIsElementTypeFilter(false), new ElementIsElementTypeFilter(true));
            IEnumerable<Element> elements = new FilteredElementCollector(doc).WherePasses(allElementKinds).ToElements();
            foreach (Element element in elements)
            {
                bool ignoredAuxiliary;
                FamilyBrowserTrackedElementState state = CaptureState(doc, element, captureContext, out ignoredAuxiliary);
                if (state != null && !string.IsNullOrWhiteSpace(state.ElementId))
                {
                    states[state.ElementId] = state;
                }
                else if (ignoredAuxiliary)
                {
                    string ignoredId = SafeElementId(element == null ? null : element.Id);
                    if (!string.IsNullOrWhiteSpace(ignoredId))
                    {
                        ignoredAuxiliaryElementIds.Add(ignoredId);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            captureError = ex;
            states.Clear();
            ignoredAuxiliaryElementIds.Clear();
            stopwatch.Stop();
            elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
            return false;
        }
        stopwatch.Stop();
        elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
        return true;
    }

    private static void RebaseSessionAfterExternalUpdate(Document doc, DocumentSession expectedSession)
    {
        Dictionary<string, FamilyBrowserTrackedElementState> refreshed;
        HashSet<string> ignoredAuxiliaryElementIds;
        long elapsed;
        Exception captureError;
        if (!CaptureBaseline(doc, out refreshed, out ignoredAuxiliaryElementIds, out elapsed, out captureError))
        {
            WriteTrackingDiagnostic(expectedSession == null ? string.Empty : expectedSession.WorkspaceRoot, "Element tracking external-update rebase failed", captureError ?? new InvalidOperationException("The element baseline could not be recaptured after an incoming update."), doc);
            lock (SyncRoot)
            {
                DocumentSession current;
                if (Sessions.TryGetValue(BuildRuntimeKey(doc), out current) && object.ReferenceEquals(current, expectedSession))
                {
                    current.ExternalRebaseFailed = true;
                }
            }
            return;
        }
        lock (SyncRoot)
        {
            DocumentSession session;
            if (!Sessions.TryGetValue(BuildRuntimeKey(doc), out session) || !object.ReferenceEquals(session, expectedSession))
            {
                return;
            }
            HashSet<string> pendingIds = GetLocallyPendingElementIds(session);
            HashSet<string> locallyAttributedIds = new HashSet<string>(pendingIds, StringComparer.Ordinal);
            locallyAttributedIds.UnionWith(GetRecoveredCheckpointElementIds(session));
            session.ExternalOverlapIds.IntersectWith(locallyAttributedIds);
            foreach (string id in session.SuppressedExternalIds)
            {
                if (locallyAttributedIds.Contains(id))
                {
                    session.ExternalOverlapIds.Add(id);
                }
            }
            session.SuppressedExternalIds.Clear();
            Dictionary<string, FamilyBrowserTrackedElementState> rebasedBaseline = new Dictionary<string, FamilyBrowserTrackedElementState>(refreshed, StringComparer.Ordinal);
            foreach (string id in pendingIds)
            {
                FamilyBrowserTrackedElementState original;
                if (session.Baseline.TryGetValue(id, out original))
                {
                    rebasedBaseline[id] = original;
                }
                else
                {
                    rebasedBaseline.Remove(id);
                }
            }
            Dictionary<string, FamilyBrowserTrackedElementState> deletedLastKnown = new Dictionary<string, FamilyBrowserTrackedElementState>(StringComparer.Ordinal);
            foreach (string id in pendingIds)
            {
                FamilyBrowserTrackedElementState previous;
                if (!refreshed.ContainsKey(id) && (session.DeletedLastKnown.TryGetValue(id, out previous) || session.Baseline.TryGetValue(id, out previous)))
                {
                    deletedLastKnown[id] = previous;
                }
            }
            session.RebaseAfterExternalUpdate(rebasedBaseline, refreshed, deletedLastKnown, ignoredAuxiliaryElementIds, elapsed);
        }
    }

    private static HashSet<string> GetLocallyPendingElementIds(DocumentSession session)
    {
        HashSet<string> pending = new HashSet<string>(StringComparer.Ordinal);
        if (session == null)
        {
            return pending;
        }
        foreach (ChangeActivity activity in session.AppliedActivities ?? new List<ChangeActivity>())
        {
            pending.UnionWith(activity == null ? Enumerable.Empty<string>() : activity.AllElementIds());
        }
        if (session.UnmatchedUndoCount > 0 || session.UnmatchedRedoCount > 0)
        {
            pending.UnionWith(session.TouchedIds);
        }
        return pending;
    }

    private static HashSet<string> GetRecoveredCheckpointElementIds(DocumentSession session)
    {
        HashSet<string> recoveredIds = new HashSet<string>(StringComparer.Ordinal);
        if (session == null)
        {
            return recoveredIds;
        }
        foreach (FamilyBrowserElementChangeCommit commit in session.RecoveredLocalSaveCommits ?? new List<FamilyBrowserElementChangeCommit>())
        {
            foreach (FamilyBrowserElementChangeItem change in commit == null || commit.Changes == null
                ? new List<FamilyBrowserElementChangeItem>()
                : commit.Changes)
            {
                if (change != null && !string.IsNullOrWhiteSpace(change.ElementId))
                {
                    recoveredIds.Add(change.ElementId);
                }
            }
        }
        return recoveredIds;
    }

    private static bool CommitMatchesProjectStableIdentity(FamilyBrowserElementChangeCommit commit, string expectedStableIdentity)
    {
        if (commit == null || string.IsNullOrWhiteSpace(expectedStableIdentity))
        {
            return false;
        }
        return string.Equals(commit.ProjectComparableIdentity ?? string.Empty, expectedStableIdentity, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(FamilyBrowserPathIdentityService.GetStablePathIdentity(commit.ProjectIdentityPath), expectedStableIdentity, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(FamilyBrowserPathIdentityService.GetStablePathIdentity(commit.ProjectCanonicalPath), expectedStableIdentity, StringComparison.OrdinalIgnoreCase);
    }

    private static FamilyBrowserElementChangeCommit CloneElementChangeCommit(FamilyBrowserElementChangeCommit commit)
    {
        if (commit == null)
        {
            throw new ArgumentNullException("commit");
        }
        using (MemoryStream stream = new MemoryStream())
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(FamilyBrowserElementChangeCommit));
            serializer.WriteObject(stream, commit);
            stream.Position = 0L;
            FamilyBrowserElementChangeCommit clone = serializer.ReadObject(stream) as FamilyBrowserElementChangeCommit;
            if (clone == null)
            {
                throw new InvalidDataException("The recovered element checkpoint commit could not be cloned safely.");
            }
            return clone;
        }
    }

    private static bool IsReloadLatestActivity(ChangeActivity activity)
    {
        if (activity == null)
        {
            return false;
        }
        string compact = CompactActivityToken((activity.Operation ?? string.Empty) + " " + string.Join(" ", activity.TransactionNames ?? new List<string>()));
        return compact.Contains("reloadlatest") || compact.Contains("latestreload") || compact.Contains("최신항목다시로드") || compact.Contains("최신버전다시로드") || compact.Contains("최신다시로드");
    }

    private static bool ShouldUseReloadLatestTransactionFallback(ChangeActivity activity)
    {
        lock (SyncRoot)
        {
            if (_reloadLatestEventSource != null &&
                _reloadingLatestEvent != null &&
                _reloadedLatestEvent != null &&
                _reloadingLatestHandler != null &&
                _reloadedLatestHandler != null)
            {
                return false;
            }
        }
        return IsReloadLatestActivity(activity);
    }

    private static string CompactActivityToken(string value)
    {
        return new string((value ?? string.Empty).Where(delegate(char ch) { return char.IsLetterOrDigit(ch); }).ToArray()).ToLowerInvariant();
    }

    private static void HandleReloadingLatestBridge(object sender, EventArgs e)
    {
        string workspaceRoot = ResolveBridgeWorkspaceRoot();
        Document doc = SafeEventDocument(e);
        try
        {
            HandleDocumentReloadingLatest(workspaceRoot, doc);
        }
        catch (Exception ex)
        {
            HandleDocumentReloadLatestStartFailure(workspaceRoot, ex);
        }
    }

    private static void HandleReloadedLatestBridge(object sender, EventArgs e)
    {
        string workspaceRoot = ResolveBridgeWorkspaceRoot();
        Document doc;
        Exception eventAccessError;
        if (!TryGetEventDocument(e, out doc, out eventAccessError))
        {
            CloseExternalUpdateWindowAfterUnknownCompletion(
                workspaceRoot,
                null,
                false,
                "Element tracking Reload Latest completion document failed",
                eventAccessError ?? new InvalidOperationException("Reload Latest completed without a readable document. Active reload suppression windows were closed conservatively."));
            return;
        }
        object status;
        if (!TryGetEventStatus(e, out status, out eventAccessError))
        {
            CloseExternalUpdateWindowAfterUnknownCompletion(
                workspaceRoot,
                doc,
                false,
                "Element tracking Reload Latest completion status failed",
                eventAccessError ?? new InvalidOperationException("Reload Latest completed without a readable status. The reload suppression window was closed conservatively."));
            return;
        }
        try
        {
            HandleDocumentReloadedLatest(workspaceRoot, doc, status);
        }
        catch (Exception ex)
        {
            CloseExternalUpdateWindowAfterUnknownCompletion(
                workspaceRoot,
                doc,
                false,
                "Element tracking Reload Latest completion callback failed",
                ex);
        }
    }

    private static Document SafeEventDocument(EventArgs e)
    {
        Document doc;
        Exception ignored;
        return TryGetEventDocument(e, out doc, out ignored) ? doc : null;
    }

    private static bool TryGetEventDocument(EventArgs e, out Document doc, out Exception error)
    {
        doc = null;
        error = null;
        try
        {
            PropertyInfo property = e == null ? null : e.GetType().GetProperty("Document", BindingFlags.Public | BindingFlags.Instance);
            if (property == null)
            {
                error = new MissingMemberException("The Revit completion event did not expose a Document property.");
                return false;
            }
            doc = property.GetValue(e, null) as Document;
            if (doc == null)
            {
                error = new InvalidOperationException("The Revit completion event exposed a null Document.");
                return false;
            }
            return true;
        }
        catch (Exception ex)
        {
            error = ex;
            return false;
        }
    }

    private static bool TryGetEventStatus(EventArgs e, out object status, out Exception error)
    {
        status = null;
        error = null;
        try
        {
            PropertyInfo property = e == null ? null : e.GetType().GetProperty("Status", BindingFlags.Public | BindingFlags.Instance);
            if (property == null)
            {
                error = new MissingMemberException("The Revit completion event did not expose a Status property.");
                return false;
            }
            status = property.GetValue(e, null);
            if (status == null)
            {
                error = new InvalidOperationException("The Revit completion event exposed a null Status.");
                return false;
            }
            return true;
        }
        catch (Exception ex)
        {
            error = ex;
            return false;
        }
    }

    private static string ResolveBridgeWorkspaceRoot()
    {
        Func<string> resolver;
        lock (SyncRoot)
        {
            resolver = _workspaceRootResolver;
        }
        try
        {
            return resolver == null ? string.Empty : resolver() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static void StopReloadLatestBridgeNoLock()
    {
        try
        {
            if (_reloadLatestEventSource != null && _reloadingLatestEvent != null && _reloadingLatestHandler != null)
            {
                _reloadingLatestEvent.RemoveEventHandler(_reloadLatestEventSource, _reloadingLatestHandler);
            }
        }
        catch
        {
        }
        try
        {
            if (_reloadLatestEventSource != null && _reloadedLatestEvent != null && _reloadedLatestHandler != null)
            {
                _reloadedLatestEvent.RemoveEventHandler(_reloadLatestEventSource, _reloadedLatestHandler);
            }
        }
        catch
        {
        }
        _reloadLatestEventSource = null;
        _reloadingLatestEvent = null;
        _reloadedLatestEvent = null;
        _reloadingLatestHandler = null;
        _reloadedLatestHandler = null;
        _workspaceRootResolver = null;
    }

    private static ChangeActivity CreateActivity(DocumentChangedEventArgs e, IEnumerable<ElementId> added, IEnumerable<ElementId> modified, IEnumerable<ElementId> deleted, out Exception metadataError)
    {
        metadataError = null;
        List<Exception> metadataErrors = new List<Exception>();
        string operation;
        Exception operationError;
        if (!TryGetOperationName(e, out operation, out operationError))
        {
            operation = "UnknownOperation";
            metadataErrors.Add(operationError ?? new InvalidOperationException("DocumentChanged operation was unavailable."));
        }
        string observedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
        ChangeActivity activity = new ChangeActivity
        {
            Operation = operation,
            ObservedAtUtc = observedAtUtc,
            LastObservedAtUtc = observedAtUtc
        };
        AddIds(activity.AddedIds, added);
        AddIds(activity.ModifiedIds, modified);
        AddIds(activity.DeletedIds, deleted);
        try
        {
            foreach (string name in e.GetTransactionNames() ?? new List<string>())
            {
                string trimmed = (name ?? string.Empty).Trim();
                if (trimmed.Length > 0 && !activity.TransactionNames.Contains(trimmed, StringComparer.OrdinalIgnoreCase))
                {
                    activity.TransactionNames.Add(trimmed);
                }
            }
        }
        catch (Exception ex)
        {
            metadataErrors.Add(ex);
        }
        if (metadataErrors.Count > 0)
        {
            metadataError = new AggregateException("DocumentChanged operation or transaction metadata could not be read completely. No operation kind was guessed.", metadataErrors);
        }
        return activity;
    }

    private static bool RecoverActivityFromCurrentSnapshot(Document doc, DocumentSession expectedSession, ChangeActivity activity, out Exception error)
    {
        error = null;
        Dictionary<string, FamilyBrowserTrackedElementState> refreshed;
        HashSet<string> refreshedIgnoredAuxiliaryElementIds;
        long elapsed;
        if (!CaptureBaseline(doc, out refreshed, out refreshedIgnoredAuxiliaryElementIds, out elapsed, out error))
        {
            return false;
        }
        lock (SyncRoot)
        {
            DocumentSession session;
            if (!Sessions.TryGetValue(BuildRuntimeKey(doc), out session) || !object.ReferenceEquals(session, expectedSession))
            {
                error = new InvalidOperationException("The document tracking session changed before the fallback snapshot could be applied.");
                return false;
            }
            HashSet<string> ignoredAuxiliaryElementIds = new HashSet<string>(session.IgnoredAuxiliaryElementIds, StringComparer.Ordinal);
            ignoredAuxiliaryElementIds.UnionWith(refreshedIgnoredAuxiliaryElementIds);
            FilterActivityElementIds(activity, ignoredAuxiliaryElementIds);
            RemoveIgnoredElementIdsFromSession(session, ignoredAuxiliaryElementIds);
            foreach (KeyValuePair<string, FamilyBrowserTrackedElementState> pair in refreshed)
            {
                FamilyBrowserTrackedElementState previous;
                if (!session.Current.TryGetValue(pair.Key, out previous))
                {
                    activity.AddedIds.Add(pair.Key);
                }
                else if (!string.Equals(previous == null ? string.Empty : previous.StateSignature, pair.Value == null ? string.Empty : pair.Value.StateSignature, StringComparison.Ordinal))
                {
                    activity.ModifiedIds.Add(pair.Key);
                }
            }
            foreach (string id in session.Current.Keys.Where(delegate(string key) { return !refreshed.ContainsKey(key); }).ToList())
            {
                activity.DeletedIds.Add(id);
                FamilyBrowserTrackedElementState previous;
                if (session.Current.TryGetValue(id, out previous) && previous != null)
                {
                    session.DeletedLastKnown[id] = previous;
                }
            }
            foreach (string id in activity.AddedIds.Concat(activity.ModifiedIds).Distinct(StringComparer.Ordinal))
            {
                FamilyBrowserTrackedElementState current;
                if (refreshed.TryGetValue(id, out current) && current != null)
                {
                    session.Current[id] = current;
                    session.DeletedLastKnown.Remove(id);
                }
            }
            foreach (string id in activity.DeletedIds)
            {
                session.Current.Remove(id);
            }
            session.ReplaceIgnoredAuxiliaryElementIds(refreshedIgnoredAuxiliaryElementIds);
            session.BaselineElapsedMilliseconds += elapsed;
        }
        return true;
    }

    private static void ApplyActivity(DocumentSession session, ChangeActivity activity)
    {
        foreach (string id in activity.AllElementIds())
        {
            session.TouchedIds.Add(id);
        }
        if (string.Equals(activity.Operation, "TransactionUndone", StringComparison.OrdinalIgnoreCase))
        {
            session.UndoCount++;
            bool exact;
            List<ChangeActivity> matched = FindMatchingActivities(session.AppliedActivities, activity, out exact);
            if (exact && matched.Count > 0)
            {
                foreach (ChangeActivity matchedActivity in matched)
                {
                    session.AppliedActivities.Remove(matchedActivity);
                    session.UndoneActivities.Add(matchedActivity);
                }
            }
            else
            {
                session.UnmatchedUndoCount++;
                AddAmbiguousActivityIds(session, activity);
                AddAmbiguousActivityIds(session, matched);
            }
            return;
        }
        if (string.Equals(activity.Operation, "TransactionRedone", StringComparison.OrdinalIgnoreCase))
        {
            session.RedoCount++;
            bool exact;
            List<ChangeActivity> matched = FindMatchingActivities(session.UndoneActivities, activity, out exact);
            if (exact && matched.Count > 0)
            {
                foreach (ChangeActivity matchedActivity in matched)
                {
                    session.UndoneActivities.Remove(matchedActivity);
                    matchedActivity.LastObservedAtUtc = string.IsNullOrWhiteSpace(activity.LastObservedAtUtc) ? activity.ObservedAtUtc : activity.LastObservedAtUtc;
                    session.AppliedActivities.Add(matchedActivity);
                }
            }
            else
            {
                session.UnmatchedRedoCount++;
                AddAmbiguousActivityIds(session, activity);
                AddAmbiguousActivityIds(session, matched);
            }
            return;
        }
        if (activity.Operation.IndexOf("RolledBack", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return;
        }
        session.AppliedActivities.Add(activity);
        session.UndoneActivities.Clear();
    }

    private static List<ChangeActivity> FindMatchingActivities(IList<ChangeActivity> candidates, ChangeActivity observed, out bool exact)
    {
        List<ChangeActivity> source = (candidates ?? new List<ChangeActivity>()).Where(delegate(ChangeActivity candidate) { return candidate != null; }).ToList();
        List<FamilyBrowserElementActivityMatchInput> inputs = source.Select(ToActivityMatchInput).ToList();
        FamilyBrowserElementActivityMatchResult result = FamilyBrowserElementActivityMatcher.Match(inputs, ToActivityMatchInput(observed));
        exact = result.Exact;
        return result.CandidateIndexes
            .Where(delegate(int index) { return index >= 0 && index < source.Count; })
            .Select(delegate(int index) { return source[index]; })
            .ToList();
    }

    private static FamilyBrowserElementActivityMatchInput ToActivityMatchInput(ChangeActivity activity)
    {
        return new FamilyBrowserElementActivityMatchInput
        {
            ElementIds = (activity == null ? Enumerable.Empty<string>() : activity.AllElementIds()).Distinct(StringComparer.Ordinal).ToList(),
            TransactionNames = activity == null || activity.TransactionNames == null
                ? new List<string>()
                : activity.TransactionNames.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
        };
    }

    private static void AddAmbiguousActivityIds(DocumentSession session, ChangeActivity activity)
    {
        if (session == null)
        {
            return;
        }
        foreach (string id in activity == null ? Enumerable.Empty<string>() : activity.AllElementIds())
        {
            session.AmbiguousActivityIds.Add(id);
        }
    }

    private static void AddAmbiguousActivityIds(DocumentSession session, IEnumerable<ChangeActivity> activities)
    {
        foreach (ChangeActivity activity in activities ?? Enumerable.Empty<ChangeActivity>())
        {
            AddAmbiguousActivityIds(session, activity);
        }
    }

    private static void RefreshProjectParameterStatesAtCommit(DocumentSession session, Document doc, StateCaptureContext captureContext)
    {
        if (session == null || doc == null || session.RecoveryOnly)
        {
            return;
        }

        captureContext = captureContext ?? BuildStateCaptureContext(doc);
        if (!captureContext.ParameterBindingsReadSucceeded)
        {
            session.CommitBoundaryReadFailureCount++;
            WriteTrackingDiagnostic(
                session.WorkspaceRoot,
                "Project parameter metadata verification failed at commit",
                new InvalidOperationException("The project/shared parameter binding map could not be read at the Save or Synchronize boundary. Existing DocumentChanged evidence remains protected, but parameter binding additions, removals, or category changes require review."),
                doc);
            return;
        }

        Dictionary<string, FamilyBrowserTrackedElementState> refreshed = new Dictionary<string, FamilyBrowserTrackedElementState>(StringComparer.Ordinal);
        try
        {
            foreach (ParameterElement parameterElement in captureContext.ParameterElements)
            {
                FamilyBrowserTrackedElementState state = CaptureState(doc, parameterElement, captureContext);
                if (state != null && !string.IsNullOrWhiteSpace(state.ElementId))
                {
                    refreshed[state.ElementId] = state;
                }
            }
        }
        catch (Exception ex)
        {
            session.CommitBoundaryReadFailureCount++;
            WriteTrackingDiagnostic(session.WorkspaceRoot, "Project parameter definition verification failed at commit", ex, doc);
            return;
        }

        HashSet<string> previousIds = new HashSet<string>(session.Current
            .Where(delegate(KeyValuePair<string, FamilyBrowserTrackedElementState> pair)
            {
                return IsProjectParameterTrackingKind(pair.Value == null ? string.Empty : pair.Value.TrackingKind);
            })
            .Select(delegate(KeyValuePair<string, FamilyBrowserTrackedElementState> pair) { return pair.Key; }), StringComparer.Ordinal);

        ChangeActivity verification = new ChangeActivity
        {
            Operation = "ProjectParameterStateVerification",
            ObservedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
            LastObservedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
        };
        verification.TransactionNames.Add("Project/shared parameter state verification at commit");

        foreach (KeyValuePair<string, FamilyBrowserTrackedElementState> pair in refreshed)
        {
            FamilyBrowserTrackedElementState previous;
            if (!session.Current.TryGetValue(pair.Key, out previous))
            {
                verification.AddedIds.Add(pair.Key);
            }
            else if (!string.Equals(previous == null ? string.Empty : previous.StateSignature, pair.Value == null ? string.Empty : pair.Value.StateSignature, StringComparison.Ordinal))
            {
                verification.ModifiedIds.Add(pair.Key);
            }
        }
        foreach (string id in previousIds.Where(delegate(string value) { return !refreshed.ContainsKey(value); }))
        {
            verification.DeletedIds.Add(id);
        }

        if (!verification.AllElementIds().Any())
        {
            return;
        }

        foreach (string id in verification.DeletedIds)
        {
            FamilyBrowserTrackedElementState previous;
            if (session.Current.TryGetValue(id, out previous) && previous != null)
            {
                session.DeletedLastKnown[id] = previous;
            }
            session.Current.Remove(id);
        }
        foreach (string id in verification.AddedIds.Concat(verification.ModifiedIds).Distinct(StringComparer.Ordinal))
        {
            FamilyBrowserTrackedElementState current;
            if (refreshed.TryGetValue(id, out current) && current != null)
            {
                session.Current[id] = current;
                session.DeletedLastKnown.Remove(id);
            }
        }
        ApplyActivity(session, verification);
    }

    private static HashSet<string> UpdateCurrentStates(DocumentSession session, Document doc, IEnumerable<ElementId> added, IEnumerable<ElementId> modified, IEnumerable<ElementId> deleted)
    {
        HashSet<string> ignoredAuxiliaryElementIds = new HashSet<string>(StringComparer.Ordinal);
        HashSet<string> deletedElementIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (ElementId id in deleted ?? Enumerable.Empty<ElementId>())
        {
            string key = SafeElementId(id);
            if (string.IsNullOrWhiteSpace(key))
            {
                continue;
            }
            deletedElementIds.Add(key);
            if (session.IgnoredAuxiliaryElementIds.Contains(key))
            {
                ignoredAuxiliaryElementIds.Add(key);
                session.Current.Remove(key);
                session.DeletedLastKnown.Remove(key);
                continue;
            }
            FamilyBrowserTrackedElementState previous;
            if (session.Current.TryGetValue(key, out previous) || session.Baseline.TryGetValue(key, out previous))
            {
                session.DeletedLastKnown[key] = previous;
            }
            session.Current.Remove(key);
        }
        foreach (ElementId id in (added ?? Enumerable.Empty<ElementId>()).Concat(modified ?? Enumerable.Empty<ElementId>()).GroupBy(SafeElementId).Select(delegate(IGrouping<string, ElementId> group) { return group.First(); }))
        {
            string key = SafeElementId(id);
            if (string.IsNullOrWhiteSpace(key))
            {
                continue;
            }
            Element element = SafeGetElement(doc, id);
            bool ignoredAuxiliary;
            FamilyBrowserTrackedElementState state = CaptureState(doc, element, out ignoredAuxiliary);
            if (FamilyBrowserElementTrackingTransitionPolicy.ShouldIgnoreChangedElement(
                ignoredAuxiliary,
                element == null,
                ignoredAuxiliaryElementIds.Contains(key),
                session.IgnoredAuxiliaryElementIds.Contains(key)))
            {
                ignoredAuxiliaryElementIds.Add(key);
                session.IgnoredAuxiliaryElementIds.Add(key);
                session.Current.Remove(key);
                session.DeletedLastKnown.Remove(key);
                continue;
            }
            if (state == null)
            {
                if (element == null)
                {
                    session.Current.Remove(key);
                }
            }
            else
            {
                FamilyBrowserElementTrackingTransitionPolicy.RestoreVisibleElementId(
                    key,
                    ignoredAuxiliaryElementIds,
                    session.IgnoredAuxiliaryElementIds);
                session.Current[key] = state;
                session.DeletedLastKnown.Remove(key);
            }
        }
        session.IgnoredAuxiliaryElementIds.ExceptWith(deletedElementIds);
        return ignoredAuxiliaryElementIds;
    }

    private static void FilterActivityElementIds(ChangeActivity activity, HashSet<string> ignoredElementIds)
    {
        if (activity == null || ignoredElementIds == null || ignoredElementIds.Count == 0)
        {
            return;
        }
        activity.AddedIds.ExceptWith(ignoredElementIds);
        activity.ModifiedIds.ExceptWith(ignoredElementIds);
        activity.DeletedIds.ExceptWith(ignoredElementIds);
    }

    private static void RemoveIgnoredElementIdsFromSession(DocumentSession session, HashSet<string> ignoredElementIds)
    {
        if (session == null || ignoredElementIds == null || ignoredElementIds.Count == 0)
        {
            return;
        }
        HashSet<string> ignored = ignoredElementIds;
        bool storedInPriorActivity = ignored.Overlaps(session.TouchedIds) || ignored.Overlaps(session.AmbiguousActivityIds);
        foreach (string id in ignored)
        {
            session.Baseline.Remove(id);
            session.Current.Remove(id);
            session.DeletedLastKnown.Remove(id);
        }
        session.TouchedIds.ExceptWith(ignored);
        session.UnknownPreviousStateIds.ExceptWith(ignored);
        session.SuppressedExternalIds.ExceptWith(ignored);
        session.ExternalOverlapIds.ExceptWith(ignored);
        session.AmbiguousActivityIds.ExceptWith(ignored);
        if (!storedInPriorActivity)
        {
            return;
        }
        foreach (ChangeActivity previousActivity in session.AppliedActivities.Concat(session.UndoneActivities))
        {
            FilterActivityElementIds(previousActivity, ignored);
        }
    }

    private static bool RefreshTouchedCurrentStates(
        DocumentSession session,
        Document doc,
        IEnumerable<string> elementIds,
        StateCaptureContext captureContext,
        out long elapsedMilliseconds,
        out Exception refreshError)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        elapsedMilliseconds = 0L;
        refreshError = null;
        if (session == null || doc == null)
        {
            stopwatch.Stop();
            elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
            refreshError = new ArgumentNullException(session == null ? "session" : "doc");
            return false;
        }
        try
        {
            captureContext = captureContext ?? BuildStateCaptureContext(doc);
            foreach (string id in (elementIds ?? Enumerable.Empty<string>()).Distinct(StringComparer.Ordinal).ToList())
            {
                FamilyBrowserTrackedElementState reference;
                if (!session.Current.TryGetValue(id, out reference) &&
                    !session.Baseline.TryGetValue(id, out reference) &&
                    !session.DeletedLastKnown.TryGetValue(id, out reference))
                {
                    reference = null;
                }
                Element element = SafeGetElementByTrackingIdentity(doc, id, reference == null ? string.Empty : reference.UniqueId);
                bool ignoredAuxiliary;
                FamilyBrowserTrackedElementState refreshed = CaptureState(doc, element, captureContext, out ignoredAuxiliary);
                if (ignoredAuxiliary)
                {
                    session.IgnoredAuxiliaryElementIds.Add(id);
                    session.Current.Remove(id);
                    session.DeletedLastKnown.Remove(id);
                    continue;
                }
                if (refreshed == null)
                {
                    if (element == null)
                    {
                        FamilyBrowserTrackedElementState previous;
                        if (session.Current.TryGetValue(id, out previous) || session.Baseline.TryGetValue(id, out previous))
                        {
                            session.DeletedLastKnown[id] = previous;
                        }
                        session.Current.Remove(id);
                    }
                }
                else
                {
                    session.Current[id] = refreshed;
                    session.DeletedLastKnown.Remove(id);
                }
            }
        }
        catch (Exception ex)
        {
            refreshError = ex;
            stopwatch.Stop();
            elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
            return false;
        }
        stopwatch.Stop();
        elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
        return true;
    }

    private static bool HasProjectCatalogRelevantChanges(
        DocumentSession session,
        IEnumerable<string> elementIds,
        IEnumerable<FamilyBrowserElementChangeCommit> recoveredCommits)
    {
        if (session != null)
        {
            foreach (string id in (elementIds ?? Enumerable.Empty<string>()).Distinct(StringComparer.Ordinal))
            {
                FamilyBrowserTrackedElementState state;
                if ((session.Baseline.TryGetValue(id, out state) && IsProjectCatalogRelevantState(state)) ||
                    (session.Current.TryGetValue(id, out state) && IsProjectCatalogRelevantState(state)) ||
                    (session.DeletedLastKnown.TryGetValue(id, out state) && IsProjectCatalogRelevantState(state)))
                {
                    return true;
                }
            }
        }
        foreach (FamilyBrowserElementChangeItem change in (recoveredCommits ?? Enumerable.Empty<FamilyBrowserElementChangeCommit>())
            .Where(delegate(FamilyBrowserElementChangeCommit commit) { return commit != null; })
            .SelectMany(delegate(FamilyBrowserElementChangeCommit commit)
            {
                return commit.Changes ?? new List<FamilyBrowserElementChangeItem>();
            }))
        {
            if (change != null &&
                (IsProjectCatalogRelevantState(change.Before) || IsProjectCatalogRelevantState(change.After)))
            {
                return true;
            }
        }
        return false;
    }

    private static bool IsProjectCatalogRelevantState(FamilyBrowserTrackedElementState state)
    {
        return state != null &&
            (state.IsElementType || string.Equals(state.ElementClass ?? string.Empty, "Family", StringComparison.OrdinalIgnoreCase));
    }

    private static void SetProjectCatalogObservationDecisionNoLock(string runtimeKey, bool required)
    {
        if (string.IsNullOrWhiteSpace(runtimeKey))
        {
            return;
        }
        bool existing;
        ProjectCatalogObservationRequiredByRuntimeKey[runtimeKey] =
            ProjectCatalogObservationRequiredByRuntimeKey.TryGetValue(runtimeKey, out existing)
                ? existing || required
                : required;
    }

    private static ChangeActivityIndex BuildChangeActivityIndex(IEnumerable<ChangeActivity> activities)
    {
        ChangeActivityIndex index = new ChangeActivityIndex();
        HashSet<string> transactionNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (ChangeActivity activity in activities ?? Enumerable.Empty<ChangeActivity>())
        {
            if (activity == null)
            {
                continue;
            }
            foreach (string transactionName in activity.TransactionNames ?? new List<string>())
            {
                if (!string.IsNullOrWhiteSpace(transactionName) && transactionNames.Add(transactionName))
                {
                    index.TransactionNames.Add(transactionName);
                }
            }
            foreach (string id in activity.AllElementIds())
            {
                if (string.IsNullOrWhiteSpace(id))
                {
                    continue;
                }
                index.ActiveIds.Add(id);
                if (activity.AddedIds.Contains(id))
                {
                    index.AddedIds.Add(id);
                }
                if (activity.DeletedIds.Contains(id))
                {
                    index.DeletedIds.Add(id);
                }
                List<ChangeActivity> elementActivities;
                if (!index.ActivitiesByElementId.TryGetValue(id, out elementActivities))
                {
                    elementActivities = new List<ChangeActivity>();
                    index.ActivitiesByElementId[id] = elementActivities;
                }
                elementActivities.Add(activity);
            }
        }
        return index;
    }

    private static FamilyBrowserElementChangeCommit BuildCommit(Document doc, DocumentSession session, string commitKind, bool policyUsedCachedFallback, bool workshared)
    {
        bool eventSequenceAmbiguous = session.UnmatchedUndoCount > 0 || session.UnmatchedRedoCount > 0;
        bool eventReadGap = session.EventReadFailureCount > 0;
        bool commitBoundaryReadGap = session.CommitBoundaryReadFailureCount > 0;
        ChangeActivityIndex activityIndex = BuildChangeActivityIndex(session.AppliedActivities);
        HashSet<string> activeIds = activityIndex.ActiveIds;
        HashSet<string> candidateIds = new HashSet<string>(session.TouchedIds, StringComparer.Ordinal);
        candidateIds.UnionWith(activeIds);
        List<FamilyBrowserElementChangeItem> changes = new List<FamilyBrowserElementChangeItem>();

        foreach (string id in candidateIds.OrderBy(delegate(string value) { return value; }, StringComparer.Ordinal))
        {
            FamilyBrowserTrackedElementState before;
            FamilyBrowserTrackedElementState after;
            session.Baseline.TryGetValue(id, out before);
            session.Current.TryGetValue(id, out after);
            bool wasAdded = activityIndex.AddedIds.Contains(id);
            bool wasDeleted = activityIndex.DeletedIds.Contains(id);
            bool elementSequenceAmbiguous = session.AmbiguousActivityIds.Contains(id);
            string kind = FamilyBrowserElementTrackingTransitionPolicy.ResolveChangeKind(
                before != null,
                after != null,
                elementSequenceAmbiguous,
                activeIds.Contains(id),
                before != null && after != null && !string.Equals(before.StateSignature, after.StateSignature, StringComparison.Ordinal),
                wasAdded,
                wasDeleted,
                session.BaselineCapturedLate);
            if (string.IsNullOrWhiteSpace(kind))
            {
                continue;
            }

            FamilyBrowserTrackedElementState lastKnown;
            session.DeletedLastKnown.TryGetValue(id, out lastKnown);
            FamilyBrowserTrackedElementState display = after ?? before ?? lastKnown;
            bool unresolvedTransient = FamilyBrowserElementTrackingTransitionPolicy.IsUnresolvedTransient(
                kind,
                before != null,
                after != null,
                lastKnown != null);
            if (session.IgnoredAuxiliaryElementIds.Contains(id) || IsAuxiliaryTrackedState(display))
            {
                continue;
            }
            bool previousUnavailable = !string.Equals(kind, "CreatedThenDeleted", StringComparison.Ordinal) &&
                (session.UnknownPreviousStateIds.Contains(id) ||
                 (string.Equals(kind, "Deleted", StringComparison.Ordinal) && before == null && lastKnown == null));
            List<ChangeActivity> activities;
            if (!activityIndex.ActivitiesByElementId.TryGetValue(id, out activities))
            {
                activities = new List<ChangeActivity>();
            }
            FamilyBrowserElementChangeItem item = new FamilyBrowserElementChangeItem
            {
                ChangeKind = kind,
                ElementId = id,
                UniqueId = display == null ? string.Empty : display.UniqueId,
                ElementClass = unresolvedTransient ? FamilyBrowserElementHistoryProjectionPolicy.UnresolvedTransientElementClass : (display == null ? string.Empty : display.ElementClass),
                CategoryName = display == null ? string.Empty : display.CategoryName,
                ElementName = display == null ? string.Empty : display.ElementName,
                FamilyName = display == null ? string.Empty : display.FamilyName,
                TypeName = display == null ? string.Empty : display.TypeName,
                TrackingKind = unresolvedTransient ? FamilyBrowserElementHistoryProjectionPolicy.UnresolvedTransientTrackingKind : (display == null ? string.Empty : display.TrackingKind),
                FirstObservedAtUtc = activities.Select(delegate(ChangeActivity activity) { return activity.ObservedAtUtc; }).Where(delegate(string value) { return !string.IsNullOrWhiteSpace(value); }).OrderBy(delegate(string value) { return value; }, StringComparer.Ordinal).FirstOrDefault() ?? string.Empty,
                LastObservedAtUtc = activities.Select(delegate(ChangeActivity activity) { return string.IsNullOrWhiteSpace(activity.LastObservedAtUtc) ? activity.ObservedAtUtc : activity.LastObservedAtUtc; }).Where(delegate(string value) { return !string.IsNullOrWhiteSpace(value); }).OrderByDescending(delegate(string value) { return value; }, StringComparer.Ordinal).FirstOrDefault() ?? string.Empty,
                Before = previousUnavailable ? null : before ?? (string.Equals(kind, "Deleted", StringComparison.Ordinal) ? lastKnown : null),
                After = after,
                PreviousStateUnavailable = previousUnavailable,
                ExternalUpdateOverlap = session.ExternalOverlapIds.Contains(id),
                TransactionNames = activities.SelectMany(delegate(ChangeActivity activity) { return activity.TransactionNames; }).Distinct(StringComparer.OrdinalIgnoreCase).ToList()
            };
            item.ChangeSummary = unresolvedTransient
                ? "Element metadata was unavailable after a same-boundary create/delete sequence."
                : BuildChangeSummary(kind, item.Before, item.After, item.PreviousStateUnavailable);
            changes.Add(item);
        }
        bool unprotectedLateBaselineGap = session.BaselineCapturedLate && !HasProtectedRecoveryEvidence(session);
        bool coverageGapOnly = changes.Count == 0 &&
            (unprotectedLateBaselineGap || eventReadGap || commitBoundaryReadGap || eventSequenceAmbiguous || session.ExternalRebaseFailed);
        if (changes.Count == 0 && !coverageGapOnly)
        {
            return null;
        }

        string identityPath = SafeProjectIdentityPath(doc);
        string revitUserName = SafeRevitUserName(doc);
        string committedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
        bool localSavePendingSynchronization = string.Equals(commitKind, "WorksharedLocalSavePendingSync", StringComparison.OrdinalIgnoreCase);
        FamilyBrowserElementChangeCommit commit = new FamilyBrowserElementChangeCommit
        {
            ProjectTitle = SafeDocumentTitle(doc),
            ProjectIdentityPath = identityPath,
            ProjectCanonicalPath = FamilyBrowserPathIdentityService.GetCanonicalPath(identityPath),
            ProjectComparableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(identityPath),
            ProjectLegacyComparableIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(identityPath),
            CommitKind = commitKind ?? string.Empty,
            CommittedAtUtc = committedAtUtc,
            LocalSaveProtectedAtUtc = localSavePendingSynchronization ? committedAtUtc : string.Empty,
            PublishedAtUtc = localSavePendingSynchronization ? string.Empty : committedAtUtc,
            RevitVersion = SafeRevitVersion(doc),
            RevitUserName = revitUserName,
            WindowsUserName = Environment.UserName,
            MachineName = Environment.MachineName,
            PolicyValidationState = session.PolicyDisableDeferred ? "DisablePendingCommit" : (policyUsedCachedFallback ? "LastKnownEnabled" : "LiveEnabled"),
            IsWorkshared = workshared,
            BaselineCapturedLate = session.BaselineCapturedLate,
            TrackingStartedAtUtc = session.TrackingStartedAtUtc,
            BaselineCapturedAtUtc = session.BaselineCapturedAtUtc,
            BaselineElapsedMilliseconds = session.BaselineElapsedMilliseconds,
            BaselineElementCount = session.Baseline.Count,
            ActivityCount = session.AppliedActivities.Count,
            UndoCount = session.UndoCount,
            RedoCount = session.RedoCount,
            UnmatchedUndoCount = session.UnmatchedUndoCount,
            UnmatchedRedoCount = session.UnmatchedRedoCount,
            CreatedCount = changes.Count(delegate(FamilyBrowserElementChangeItem item) { return string.Equals(item.ChangeKind, "Created", StringComparison.Ordinal) || string.Equals(item.ChangeKind, "CreatedThenDeleted", StringComparison.Ordinal); }),
            ModifiedCount = changes.Count(delegate(FamilyBrowserElementChangeItem item) { return string.Equals(item.ChangeKind, "Modified", StringComparison.Ordinal); }),
            DeletedCount = changes.Count(delegate(FamilyBrowserElementChangeItem item) { return string.Equals(item.ChangeKind, "Deleted", StringComparison.Ordinal) || string.Equals(item.ChangeKind, "CreatedThenDeleted", StringComparison.Ordinal); }),
            TransientCreatedDeletedCount = changes.Count(delegate(FamilyBrowserElementChangeItem item) { return string.Equals(item.ChangeKind, "CreatedThenDeleted", StringComparison.Ordinal); }),
            ExternalUpdateOverlapCount = changes.Count(delegate(FamilyBrowserElementChangeItem item) { return item.ExternalUpdateOverlap; }),
            CoverageGapOnly = coverageGapOnly,
            EventReadFailureCount = session.EventReadFailureCount,
            CommitBoundaryReadFailureCount = session.CommitBoundaryReadFailureCount,
            TransactionNames = new List<string>(activityIndex.TransactionNames),
            Changes = changes
        };
        commit.AttributionConfidence = commit.ExternalUpdateOverlapCount > 0
            ? "ClientObservedWithExternalOverlap"
            : (session.ExternalRebaseFailed
                ? "ClientObservedWithExternalRebaseGap"
                : (commitBoundaryReadGap
                    ? "ClientObservedWithCommitBoundaryGap"
                    : (eventReadGap
                        ? "ClientObservedWithEventReadGap"
                        : (eventSequenceAmbiguous
                            ? "ClientObservedWithEventAmbiguity"
                            : (string.IsNullOrWhiteSpace(revitUserName) ? "ClientObservedWithIdentityGap" : "ClientObserved")))));
        string baselineNote = session.BaselineCapturedLate
            ? "Client observed the change after a late baseline; previous metadata can be incomplete."
            : "Client observed from a pre-change baseline. Exact attribution requires the add-in on every editing workstation.";
        commit.CoverageNote = commit.ExternalUpdateOverlapCount > 0
            ? baselineNote + " One or more locally touched elements also appeared in an incoming central/reload update; final state was recaptured but authorship is mixed or uncertain."
            : baselineNote;
        if (policyUsedCachedFallback)
        {
            commit.CoverageNote += session.PolicyDisableDeferred
                ? " The shared policy was disabled while this client still held uncommitted observed activity. Tracking continued only until this successful Save or Synchronize boundary so the already-observed evidence was not discarded."
                : " The shared policy could not be re-read at commit time; tracking continued from the last confirmed enabled state and the record was protected locally if managed storage was unavailable.";
        }
        if (eventSequenceAmbiguous)
        {
            commit.CoverageNote += " One or more Undo/Redo callbacks could not be matched to an observed transaction. Ambiguity was limited to the affected element IDs; their final document state was retained, but transaction-level attribution requires review.";
        }
        if (eventReadGap)
        {
            commit.CoverageNote += " One or more DocumentChanged ID or operation-metadata reads were incomplete. A full current-state comparison was attempted when element IDs were unavailable, but event-level coverage and transaction attribution require review.";
        }
        if (commitBoundaryReadGap)
        {
            commit.CoverageNote += " One or more earlier Save, Save As, synchronization, or external-update completion boundaries could not expose a trustworthy document or status. Final element state is retained, but exact grouping across those boundaries requires review.";
        }
        if (string.IsNullOrWhiteSpace(revitUserName))
        {
            commit.CoverageNote += " The Revit username was unavailable; Windows user and machine identify the observing client, but exact Revit-user attribution requires review.";
        }
        commit.CoverageNote += " Tracking includes project/shared parameter definitions and binding metadata. It excludes View, DataStorage, ProjectInfo, temporary negative-ID, and other categoryless internal Revit elements.";
        if (session.ExternalRebaseFailed)
        {
            commit.CoverageNote += " An incoming central/reload update could not be fully rebased. The final local state was retained, but exact intervening authorship requires review.";
        }
        if (coverageGapOnly)
        {
            commit.CoverageNote += " No trustworthy element ID could be retained for this boundary. This coverage-gap record is intentionally preserved without inventing an element identity.";
        }
        return commit;
    }

    private static string BuildChangeSummary(string kind, FamilyBrowserTrackedElementState before, FamilyBrowserTrackedElementState after, bool previousUnavailable)
    {
        FamilyBrowserTrackedElementState display = after ?? before;
        string trackingKind = display == null ? string.Empty : display.TrackingKind ?? string.Empty;
        if (string.Equals(kind, "Created", StringComparison.Ordinal))
        {
            if (IsProjectParameterTrackingKind(trackingKind))
            {
                return ParameterTrackingSubject(trackingKind) + " created: " + DescribeParameterState(after) + ".";
            }
            if (string.Equals(trackingKind, "Grid", StringComparison.OrdinalIgnoreCase))
            {
                return "Grid created: " + EmptyDash(after == null ? string.Empty : after.ElementName) + ".";
            }
            return "Element created.";
        }
        if (string.Equals(kind, "Deleted", StringComparison.Ordinal))
        {
            if (!previousUnavailable && IsProjectParameterTrackingKind(trackingKind))
            {
                return ParameterTrackingSubject(trackingKind) + " deleted: " + DescribeParameterState(before) + ".";
            }
            if (!previousUnavailable && string.Equals(trackingKind, "Grid", StringComparison.OrdinalIgnoreCase))
            {
                return "Grid deleted: " + EmptyDash(before == null ? string.Empty : before.ElementName) + ".";
            }
            return previousUnavailable ? "Element deleted; previous metadata was unavailable." : "Element deleted.";
        }
        if (string.Equals(kind, "CreatedThenDeleted", StringComparison.Ordinal))
        {
            return "Element was created and deleted before the successful save or synchronization boundary.";
        }
        if (previousUnavailable)
        {
            return "Element modified; previous metadata was unavailable because tracking started after the change event.";
        }
        List<string> parts = new List<string>();
        if (IsProjectParameterTrackingKind(trackingKind))
        {
            AddDifference(parts, "parameter name", before == null ? string.Empty : before.ElementName, after == null ? string.Empty : after.ElementName);
            AddDifference(parts, "GUID", before == null ? string.Empty : before.SharedParameterGuid, after == null ? string.Empty : after.SharedParameterGuid);
            AddDifference(parts, "binding", before == null ? string.Empty : before.ParameterBindingKind, after == null ? string.Empty : after.ParameterBindingKind);
            AddDifference(parts, "categories", before == null ? string.Empty : before.ParameterBoundCategories, after == null ? string.Empty : after.ParameterBoundCategories);
            AddDifference(parts, "category ids", before == null ? string.Empty : before.ParameterBoundCategoryIds, after == null ? string.Empty : after.ParameterBoundCategoryIds);
            AddDifference(parts, "group", before == null ? string.Empty : before.ParameterGroup, after == null ? string.Empty : after.ParameterGroup);
            AddDifference(parts, "data type", before == null ? string.Empty : before.ParameterDataType, after == null ? string.Empty : after.ParameterDataType);
            AddDifference(parts, "varies across groups", before == null ? string.Empty : before.ParameterVariesAcrossGroups, after == null ? string.Empty : after.ParameterVariesAcrossGroups);
            return parts.Count == 0
                ? ParameterTrackingSubject(trackingKind) + " definition or binding changed."
                : ParameterTrackingSubject(trackingKind) + " modified; " + string.Join("; ", parts) + ".";
        }
        if (string.Equals(trackingKind, "Grid", StringComparison.OrdinalIgnoreCase))
        {
            AddDifference(parts, "grid name", before == null ? string.Empty : before.ElementName, after == null ? string.Empty : after.ElementName);
            AddDifference(parts, "grid type", before == null ? string.Empty : before.TypeName, after == null ? string.Empty : after.TypeName);
            AddDifference(parts, "curve", before == null ? string.Empty : before.GridCurveSignature, after == null ? string.Empty : after.GridCurveSignature);
            AddDifference(parts, "extents", before == null ? string.Empty : before.GridExtentsSignature, after == null ? string.Empty : after.GridExtentsSignature);
            AddDifference(parts, "pinned", before == null ? string.Empty : before.GridPinnedState, after == null ? string.Empty : after.GridPinnedState);
            AddDifference(parts, "workset", before == null ? string.Empty : before.WorksetId, after == null ? string.Empty : after.WorksetId);
            return parts.Count == 0 ? "Grid parameters, geometry, or datum state changed." : "Grid modified; " + string.Join("; ", parts) + ".";
        }
        AddDifference(parts, "name", before == null ? string.Empty : before.ElementName, after == null ? string.Empty : after.ElementName);
        AddDifference(parts, "family", before == null ? string.Empty : before.FamilyName, after == null ? string.Empty : after.FamilyName);
        AddDifference(parts, "type", before == null ? string.Empty : before.TypeName, after == null ? string.Empty : after.TypeName);
        AddDifference(parts, "type id", before == null ? string.Empty : before.TypeId, after == null ? string.Empty : after.TypeId);
        AddDifference(parts, "level", before == null ? string.Empty : before.LevelId, after == null ? string.Empty : after.LevelId);
        AddDifference(parts, "workset", before == null ? string.Empty : before.WorksetId, after == null ? string.Empty : after.WorksetId);
        AddDifference(parts, "location", before == null ? string.Empty : before.LocationSignature, after == null ? string.Empty : after.LocationSignature);
        return parts.Count == 0 ? "Element parameters, geometry, or internal state changed." : string.Join("; ", parts) + ".";
    }

    private static string ParameterTrackingSubject(string trackingKind)
    {
        return string.Equals(trackingKind, "SharedParameter", StringComparison.OrdinalIgnoreCase) ? "Shared parameter" : "Project parameter";
    }

    private static string DescribeParameterState(FamilyBrowserTrackedElementState state)
    {
        if (state == null)
        {
            return "-";
        }
        List<string> parts = new List<string> { EmptyDash(state.ElementName) };
        if (!string.IsNullOrWhiteSpace(state.SharedParameterGuid)) parts.Add("GUID " + state.SharedParameterGuid);
        if (!string.IsNullOrWhiteSpace(state.ParameterBindingKind)) parts.Add(state.ParameterBindingKind + " binding");
        if (!string.IsNullOrWhiteSpace(state.ParameterBoundCategories)) parts.Add("categories " + state.ParameterBoundCategories);
        if (!string.IsNullOrWhiteSpace(state.ParameterGroup)) parts.Add("group " + state.ParameterGroup);
        if (!string.IsNullOrWhiteSpace(state.ParameterDataType)) parts.Add("data type " + state.ParameterDataType);
        return string.Join("; ", parts);
    }

    private static void AddDifference(ICollection<string> parts, string label, string before, string after)
    {
        if (!string.Equals(before ?? string.Empty, after ?? string.Empty, StringComparison.Ordinal))
        {
            parts.Add(label + ": " + EmptyDash(before) + " -> " + EmptyDash(after));
        }
    }

    private static string EmptyDash(string value)
    {
        return string.IsNullOrWhiteSpace(value) ? "-" : value;
    }

    private static StateCaptureContext BuildStateCaptureContext(Document doc)
    {
        StateCaptureContext context = new StateCaptureContext();
        if (doc == null)
        {
            return context;
        }

        try
        {
            foreach (ParameterElement parameterElement in new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterElement))
                .ToElements()
                .OfType<ParameterElement>())
            {
                context.ParameterElements.Add(parameterElement);
                Definition definition = SafeParameterDefinition(parameterElement);
                ParameterBindingState state = BuildParameterBindingState(doc, parameterElement, definition);
                IndexParameterBindingState(context, state);
            }
        }
        catch
        {
        }

        try
        {
            BindingMap map = doc.ParameterBindings;
            if (map == null)
            {
                context.ParameterBindingsReadSucceeded = true;
                return context;
            }
            DefinitionBindingMapIterator iterator = map.ForwardIterator();
            iterator.Reset();
            while (iterator.MoveNext())
            {
                Definition definition = iterator.Key as Definition;
                Binding binding = iterator.Current as Binding;
                ElementBinding elementBinding = binding as ElementBinding;
                if (definition == null || elementBinding == null)
                {
                    continue;
                }

                string definitionName = SafeDefinitionName(definition);
                string sharedGuid = SafeDefinitionGuid(definition);
                string elementId = SafeDefinitionElementId(doc, definition);
                ParameterBindingState state = ResolveParameterBindingState(context, elementId, sharedGuid, definitionName);
                if (state == null)
                {
                    state = new ParameterBindingState
                    {
                        ElementId = elementId,
                        DefinitionName = definitionName,
                        SharedGuid = sharedGuid,
                        ParameterGroup = SafeDefinitionGroup(definition),
                        DataType = SafeDefinitionDataType(definition),
                        VariesAcrossGroups = SafeDefinitionVariesAcrossGroups(definition)
                    };
                }
                state.BindingKind = binding is InstanceBinding ? "Instance" : (binding is TypeBinding ? "Type" : binding.GetType().Name);
                ApplyBoundCategories(state, elementBinding);
                IndexParameterBindingState(context, state);
            }
            context.ParameterBindingsReadSucceeded = true;
        }
        catch
        {
            context.ParameterBindingsReadSucceeded = false;
        }
        return context;
    }

    private static ParameterBindingState BuildParameterBindingState(Document doc, ParameterElement parameterElement, Definition definition)
    {
        string sharedGuid = SafeSharedParameterGuid(parameterElement);
        if (string.IsNullOrWhiteSpace(sharedGuid))
        {
            sharedGuid = SafeDefinitionGuid(definition);
        }
        return new ParameterBindingState
        {
            ElementId = SafeElementId(parameterElement == null ? null : parameterElement.Id),
            DefinitionName = SafeDefinitionName(definition),
            SharedGuid = sharedGuid,
            BindingKind = "Unbound",
            ParameterGroup = SafeDefinitionGroup(definition),
            DataType = SafeDefinitionDataType(definition),
            VariesAcrossGroups = SafeDefinitionVariesAcrossGroups(definition)
        };
    }

    private static void IndexParameterBindingState(StateCaptureContext context, ParameterBindingState state)
    {
        if (context == null || state == null)
        {
            return;
        }
        if (!string.IsNullOrWhiteSpace(state.ElementId))
        {
            context.ParameterByElementId[state.ElementId] = state;
        }
        if (!string.IsNullOrWhiteSpace(state.SharedGuid))
        {
            context.ParameterByGuid[state.SharedGuid] = state;
        }
        string nameKey = NormalizeParameterDefinitionKey(state.DefinitionName);
        if (!string.IsNullOrWhiteSpace(nameKey) && !context.ParameterByName.ContainsKey(nameKey))
        {
            context.ParameterByName[nameKey] = state;
        }
    }

    private static ParameterBindingState ResolveParameterBindingState(StateCaptureContext context, string elementId, string sharedGuid, string definitionName)
    {
        if (context == null)
        {
            return null;
        }
        ParameterBindingState state;
        if (!string.IsNullOrWhiteSpace(elementId) && context.ParameterByElementId.TryGetValue(elementId, out state))
        {
            return state;
        }
        if (!string.IsNullOrWhiteSpace(sharedGuid) && context.ParameterByGuid.TryGetValue(sharedGuid, out state))
        {
            return state;
        }
        string nameKey = NormalizeParameterDefinitionKey(definitionName);
        return !string.IsNullOrWhiteSpace(nameKey) && context.ParameterByName.TryGetValue(nameKey, out state) ? state : null;
    }

    private static void ApplyBoundCategories(ParameterBindingState state, ElementBinding binding)
    {
        SortedSet<string> names = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        SortedSet<string> ids = new SortedSet<string>(StringComparer.Ordinal);
        try
        {
            foreach (Category category in binding == null || binding.Categories == null ? Enumerable.Empty<Category>() : binding.Categories.Cast<Category>())
            {
                if (category == null)
                {
                    continue;
                }
                string name = (category.Name ?? string.Empty).Trim();
                string id = SafeElementId(category.Id);
                if (!string.IsNullOrWhiteSpace(name))
                {
                    names.Add(name);
                }
                if (!string.IsNullOrWhiteSpace(id))
                {
                    ids.Add(id);
                }
            }
        }
        catch
        {
        }
        state.CategoryNames = string.Join(", ", names);
        state.CategoryIds = string.Join(",", ids);
    }

    private static FamilyBrowserTrackedElementState CaptureState(Document doc, Element element)
    {
        StateCaptureContext context = element is ParameterElement ? BuildStateCaptureContext(doc) : null;
        return CaptureState(doc, element, context);
    }

    private static FamilyBrowserTrackedElementState CaptureState(Document doc, Element element, out bool ignoredAuxiliary)
    {
        StateCaptureContext context = element is ParameterElement ? BuildStateCaptureContext(doc) : null;
        return CaptureState(doc, element, context, out ignoredAuxiliary);
    }

    private static FamilyBrowserTrackedElementState CaptureState(Document doc, Element element, StateCaptureContext captureContext)
    {
        bool ignoredAuxiliary;
        return CaptureState(doc, element, captureContext, out ignoredAuxiliary);
    }

    private static FamilyBrowserTrackedElementState CaptureState(Document doc, Element element, StateCaptureContext captureContext, out bool ignoredAuxiliary)
    {
        if (!IsTrackableElement(element, out ignoredAuxiliary))
        {
            return null;
        }
        string elementId = SafeElementId(element.Id);
        ElementType elementType = element as ElementType;
        ElementType assignedType = elementType ?? SafeGetElement(doc, SafeGetTypeId(element)) as ElementType;
        ParameterElement parameterElement = element as ParameterElement;
        SharedParameterElement sharedParameterElement = element as SharedParameterElement;
        Grid grid = element as Grid;
        Definition parameterDefinition = SafeParameterDefinition(parameterElement);
        string parameterGuid = SafeSharedParameterGuid(parameterElement);
        ParameterBindingState parameterBinding = parameterElement == null
            ? null
            : ResolveParameterBindingState(captureContext, elementId, parameterGuid, SafeDefinitionName(parameterDefinition));
        string elementName = parameterElement == null ? SafeElementName(element) : SafeDefinitionName(parameterDefinition);
        string trackingKind = sharedParameterElement != null
            ? "SharedParameter"
            : (parameterElement != null ? "ProjectParameter" : (grid != null ? "Grid" : "Element"));
        string gridCurve = SafeGridCurveSignature(grid);
        string locationSignature = SafeLocationSignature(element);
        if (grid != null && !string.IsNullOrWhiteSpace(gridCurve))
        {
            locationSignature = gridCurve;
        }
        string familyName = parameterElement == null ? ResolveFamilyName(element, assignedType) : string.Empty;
        string typeName = parameterElement == null
            ? (elementType == null ? SafeElementName(assignedType) : SafeElementName(elementType))
            : elementName;
        FamilyBrowserTrackedElementState state = new FamilyBrowserTrackedElementState
        {
            ElementId = elementId,
            UniqueId = SafeUniqueId(element),
            ElementClass = element.GetType().Name,
            CategoryName = sharedParameterElement != null ? "Shared Parameter" : (parameterElement != null ? "Project Parameter" : SafeCategoryName(element)),
            CategoryId = SafeCategoryId(element),
            ElementName = elementName,
            FamilyName = familyName,
            TypeName = typeName,
            TypeId = SafeElementId(SafeGetTypeId(element)),
            LevelId = SafeLevelId(element),
            WorksetId = SafeWorksetId(element),
            LocationSignature = locationSignature,
            TrackingKind = trackingKind,
            SharedParameterGuid = string.IsNullOrWhiteSpace(parameterGuid) && parameterBinding != null ? parameterBinding.SharedGuid : parameterGuid,
            ParameterBindingKind = parameterBinding == null ? (parameterElement == null ? string.Empty : "Unbound") : parameterBinding.BindingKind,
            ParameterBoundCategories = parameterBinding == null ? string.Empty : parameterBinding.CategoryNames,
            ParameterBoundCategoryIds = parameterBinding == null ? string.Empty : parameterBinding.CategoryIds,
            ParameterGroup = parameterBinding == null ? SafeDefinitionGroup(parameterDefinition) : parameterBinding.ParameterGroup,
            ParameterDataType = parameterBinding == null ? SafeDefinitionDataType(parameterDefinition) : parameterBinding.DataType,
            ParameterVariesAcrossGroups = parameterBinding == null ? SafeDefinitionVariesAcrossGroups(parameterDefinition) : parameterBinding.VariesAcrossGroups,
            GridCurveSignature = gridCurve,
            GridExtentsSignature = SafeGridExtentsSignature(grid),
            GridPinnedState = SafeGridPinnedState(grid),
            IsElementType = elementType != null,
            IsViewSpecific = SafeViewSpecific(element)
        };
        state.StateSignature = HashText(string.Join("|", new[]
        {
            state.ElementClass, state.CategoryId, state.ElementName, state.FamilyName, state.TypeName,
            state.TypeId, state.LevelId, state.WorksetId, state.LocationSignature,
            state.TrackingKind, state.SharedParameterGuid, state.ParameterBindingKind,
            state.ParameterBoundCategories, state.ParameterBoundCategoryIds, state.ParameterGroup,
            state.ParameterDataType, state.ParameterVariesAcrossGroups,
            state.GridCurveSignature, state.GridExtentsSignature, state.GridPinnedState,
            state.IsElementType ? "T" : "I", state.IsViewSpecific ? "V" : "M"
        }));
        return state;
    }

    private static bool IsTrackableElement(Element element, out bool ignoredAuxiliary)
    {
        ignoredAuxiliary = false;
        string elementId = element == null ? string.Empty : SafeElementId(element.Id);
        if (element == null || string.IsNullOrWhiteSpace(elementId) || elementId.StartsWith("-", StringComparison.Ordinal))
        {
            return false;
        }
        string className = element.GetType().Name;
        if (element is View || string.Equals(className, "DataStorage", StringComparison.Ordinal) || string.Equals(className, "ProjectInfo", StringComparison.Ordinal))
        {
            return false;
        }
        if (element is ElementType || element is Family || element is Material || element is ParameterElement)
        {
            return true;
        }
        bool hasCategory;
        if (IsAuxiliaryRevitSupportElement(element, out hasCategory))
        {
            ignoredAuxiliary = true;
            return false;
        }
        return hasCategory;
    }

    private static bool IsTrackableElement(Element element)
    {
        bool ignoredAuxiliary;
        return IsTrackableElement(element, out ignoredAuxiliary);
    }

    private static bool IsAuxiliaryRevitSupportElement(Element element, out bool hasCategory)
    {
        hasCategory = false;
        if (element == null || element is ElementType || element is Family || element is Material || element is ParameterElement)
        {
            return false;
        }
        FamilyInstance familyInstance = element as FamilyInstance;
        if (familyInstance != null)
        {
            try
            {
                if (familyInstance.SuperComponent != null)
                {
                    return true;
                }
            }
            catch
            {
            }
        }
        string className = element.GetType().Name;
        if (FamilyBrowserElementTrackingScopePolicy.IsAuxiliarySupportRecord(className, string.Empty, string.Empty))
        {
            return true;
        }
        Category category;
        try
        {
            category = element.Category;
        }
        catch
        {
            category = null;
        }
        hasCategory = category != null;
        if (!hasCategory)
        {
            return false;
        }
        string categoryId = SafeElementId(category.Id);
        if (FamilyBrowserElementTrackingScopePolicy.IsAuxiliarySupportRecord(string.Empty, categoryId, string.Empty))
        {
            return true;
        }
        if (!string.IsNullOrWhiteSpace(categoryId))
        {
            return false;
        }
        string categoryName;
        try
        {
            categoryName = category.Name;
        }
        catch
        {
            categoryName = string.Empty;
        }
        return FamilyBrowserElementTrackingScopePolicy.IsAuxiliarySupportRecord(string.Empty, string.Empty, categoryName);
    }

    private static bool IsAuxiliaryTrackedState(FamilyBrowserTrackedElementState state)
    {
        return state != null && FamilyBrowserElementTrackingScopePolicy.IsAuxiliarySupportRecord(
            state.ElementClass,
            state.CategoryId,
            state.CategoryName,
            state.IsElementType);
    }

    private static bool IsProjectParameterTrackingKind(string trackingKind)
    {
        return string.Equals(trackingKind, "SharedParameter", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(trackingKind, "ProjectParameter", StringComparison.OrdinalIgnoreCase);
    }

    private static Definition SafeParameterDefinition(ParameterElement parameterElement)
    {
        try
        {
            return parameterElement == null ? null : parameterElement.GetDefinition();
        }
        catch
        {
            return null;
        }
    }

    private static string SafeDefinitionName(Definition definition)
    {
        try
        {
            return definition == null ? string.Empty : (definition.Name ?? string.Empty).Trim();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeSharedParameterGuid(ParameterElement parameterElement)
    {
        try
        {
            SharedParameterElement shared = parameterElement as SharedParameterElement;
            Guid guid = shared == null ? Guid.Empty : shared.GuidValue;
            return guid == Guid.Empty ? string.Empty : guid.ToString("D").ToUpperInvariant();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeDefinitionGuid(Definition definition)
    {
        try
        {
            ExternalDefinition external = definition as ExternalDefinition;
            return external == null || external.GUID == Guid.Empty ? string.Empty : external.GUID.ToString("D").ToUpperInvariant();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeDefinitionElementId(Document doc, Definition definition)
    {
        try
        {
            InternalDefinition internalDefinition = definition as InternalDefinition;
            if (internalDefinition != null)
            {
                PropertyInfo idProperty = internalDefinition.GetType().GetProperty("Id", BindingFlags.Public | BindingFlags.Instance);
                ElementId id = idProperty == null ? null : idProperty.GetValue(internalDefinition, null) as ElementId;
                string value = SafeElementId(id);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }
            ExternalDefinition external = definition as ExternalDefinition;
            if (external != null && doc != null)
            {
                SharedParameterElement shared = SharedParameterElement.Lookup(doc, external.GUID);
                return shared == null ? string.Empty : SafeElementId(shared.Id);
            }
        }
        catch
        {
        }
        return string.Empty;
    }

    private static string SafeDefinitionGroup(Definition definition)
    {
        return SafeDefinitionMemberText(definition, "GetGroupTypeId", "ParameterGroup");
    }

    private static string SafeDefinitionDataType(Definition definition)
    {
        return SafeDefinitionMemberText(definition, "GetDataType", "ParameterType");
    }

    private static string SafeDefinitionVariesAcrossGroups(Definition definition)
    {
        try
        {
            PropertyInfo property = definition == null ? null : definition.GetType().GetProperty("VariesAcrossGroups", BindingFlags.Public | BindingFlags.Instance);
            if (property == null)
            {
                return string.Empty;
            }
            object value = property.GetValue(definition, null);
            return value is bool ? ((bool)value ? "Yes" : "No") : (value == null ? string.Empty : value.ToString());
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeDefinitionMemberText(Definition definition, string methodName, string propertyName)
    {
        try
        {
            if (definition == null)
            {
                return string.Empty;
            }
            MethodInfo method = definition.GetType().GetMethod(methodName, BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
            object value = method == null ? null : method.Invoke(definition, null);
            if (value == null)
            {
                PropertyInfo property = definition.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
                value = property == null ? null : property.GetValue(definition, null);
            }
            return value == null ? string.Empty : value.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string NormalizeParameterDefinitionKey(string value)
    {
        return (value ?? string.Empty).Trim().ToUpperInvariant();
    }

    private static string SafeGridCurveSignature(Grid grid)
    {
        try
        {
            Curve curve = grid == null ? null : grid.Curve;
            if (curve == null)
            {
                return string.Empty;
            }
            XYZ start = curve.GetEndPoint(0);
            XYZ end = curve.GetEndPoint(1);
            XYZ middle = curve.Evaluate(0.5d, true);
            return curve.GetType().Name + ": " + FormatPointMillimeters(start) + " -> " + FormatPointMillimeters(end) + "; mid " + FormatPointMillimeters(middle);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeGridExtentsSignature(Grid grid)
    {
        try
        {
            MethodInfo method = grid == null ? null : grid.GetType().GetMethod("GetExtents", BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
            BoundingBoxXYZ extents = method == null ? null : method.Invoke(grid, null) as BoundingBoxXYZ;
            return extents == null ? string.Empty : FormatPointMillimeters(extents.Min) + " -> " + FormatPointMillimeters(extents.Max);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeGridPinnedState(Grid grid)
    {
        try
        {
            return grid == null ? string.Empty : (grid.Pinned ? "Pinned" : "Unpinned");
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string ResolveFamilyName(Element element, ElementType assignedType)
    {
        Family family = element as Family;
        if (family != null)
        {
            return SafeElementName(family);
        }
        FamilyInstance instance = element as FamilyInstance;
        if (instance != null)
        {
            try
            {
                return instance.Symbol == null || instance.Symbol.Family == null ? string.Empty : SafeElementName(instance.Symbol.Family);
            }
            catch
            {
                return string.Empty;
            }
        }
        FamilySymbol symbol = element as FamilySymbol ?? assignedType as FamilySymbol;
        if (symbol != null)
        {
            try
            {
                return symbol.Family == null ? string.Empty : SafeElementName(symbol.Family);
            }
            catch
            {
                return string.Empty;
            }
        }
        return SafeStringProperty(assignedType, "FamilyName");
    }

    private static string SafeLocationSignature(Element element)
    {
        try
        {
            LocationPoint point = element.Location as LocationPoint;
            if (point != null && point.Point != null)
            {
                return "P:" + FormatDouble(point.Point.X) + "," + FormatDouble(point.Point.Y) + "," + FormatDouble(point.Point.Z) + "|R:" + FormatDouble(point.Rotation);
            }
            LocationCurve curve = element.Location as LocationCurve;
            if (curve != null && curve.Curve != null)
            {
                XYZ start = curve.Curve.GetEndPoint(0);
                XYZ end = curve.Curve.GetEndPoint(1);
                return "C:" + FormatPoint(start) + "->" + FormatPoint(end);
            }
        }
        catch
        {
        }
        return string.Empty;
    }

    private static string FormatPoint(XYZ point)
    {
        return point == null ? string.Empty : FormatDouble(point.X) + "," + FormatDouble(point.Y) + "," + FormatDouble(point.Z);
    }

    private static string FormatPointMillimeters(XYZ point)
    {
        return point == null
            ? string.Empty
            : "(" + FormatMillimeters(point.X) + ", " + FormatMillimeters(point.Y) + ", " + FormatMillimeters(point.Z) + ") mm";
    }

    private static string FormatMillimeters(double internalFeet)
    {
        return (internalFeet * 304.8d).ToString("0.###", CultureInfo.InvariantCulture);
    }

    private static string FormatDouble(double value)
    {
        return value.ToString("0.########", CultureInfo.InvariantCulture);
    }

    private static string SafeLevelId(Element element)
    {
        try
        {
            PropertyInfo property = element.GetType().GetProperty("LevelId", BindingFlags.Public | BindingFlags.Instance);
            return property == null ? string.Empty : SafeElementId(property.GetValue(element, null) as ElementId);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeWorksetId(Element element)
    {
        try
        {
            return element.WorksetId == null ? string.Empty : element.WorksetId.IntegerValue.ToString(CultureInfo.InvariantCulture);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static bool SafeViewSpecific(Element element)
    {
        try
        {
            return element.ViewSpecific;
        }
        catch
        {
            return false;
        }
    }

    private static string SafeUniqueId(Element element)
    {
        try
        {
            return element.UniqueId ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeElementName(Element element)
    {
        try
        {
            return element == null ? string.Empty : element.Name ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeStringProperty(object value, string propertyName)
    {
        try
        {
            if (value == null)
            {
                return string.Empty;
            }
            PropertyInfo property = value.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
            object result = property == null ? null : property.GetValue(value, null);
            return result == null ? string.Empty : result.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeCategoryName(Element element)
    {
        try
        {
            return element.Category == null ? string.Empty : element.Category.Name ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeCategoryId(Element element)
    {
        try
        {
            return element.Category == null ? string.Empty : SafeElementId(element.Category.Id);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static ElementId SafeGetTypeId(Element element)
    {
        try
        {
            return element == null ? null : element.GetTypeId();
        }
        catch
        {
            return null;
        }
    }

    private static Element SafeGetElement(Document doc, ElementId id)
    {
        try
        {
            return doc == null || id == null ? null : doc.GetElement(id);
        }
        catch
        {
            return null;
        }
    }

    private static Element SafeGetElementByTrackingIdentity(Document doc, string elementId, string uniqueId)
    {
        if (doc == null)
        {
            return null;
        }
        if (!string.IsNullOrWhiteSpace(uniqueId))
        {
            try
            {
                Element byUniqueId = doc.GetElement(uniqueId);
                if (byUniqueId != null)
                {
                    return byUniqueId;
                }
            }
            catch
            {
            }
        }
        long numericId;
        if (!long.TryParse(elementId ?? string.Empty, NumberStyles.Integer, CultureInfo.InvariantCulture, out numericId))
        {
            return null;
        }
        try
        {
            ConstructorInfo longConstructor = typeof(ElementId).GetConstructor(new[] { typeof(long) });
            if (longConstructor != null)
            {
                return SafeGetElement(doc, longConstructor.Invoke(new object[] { numericId }) as ElementId);
            }
            if (numericId >= int.MinValue && numericId <= int.MaxValue)
            {
                ConstructorInfo integerConstructor = typeof(ElementId).GetConstructor(new[] { typeof(int) });
                if (integerConstructor != null)
                {
                    return SafeGetElement(doc, integerConstructor.Invoke(new object[] { (int)numericId }) as ElementId);
                }
            }
        }
        catch
        {
        }
        return null;
    }

    private static string SafeElementId(ElementId id)
    {
        if (id == null)
        {
            return string.Empty;
        }
        try
        {
            PropertyInfo valueProperty = id.GetType().GetProperty("Value", BindingFlags.Public | BindingFlags.Instance);
            if (valueProperty != null)
            {
                object value = valueProperty.GetValue(id, null);
                if (value != null)
                {
                    return Convert.ToInt64(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);
                }
            }
        }
        catch
        {
        }
        try
        {
            PropertyInfo integerValueProperty = id.GetType().GetProperty("IntegerValue", BindingFlags.Public | BindingFlags.Instance);
            object integerValue = integerValueProperty == null ? null : integerValueProperty.GetValue(id, null);
            return integerValue == null ? string.Empty : Convert.ToInt64(integerValue, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static bool TryGetOperationName(DocumentChangedEventArgs e, out string operation, out Exception error)
    {
        operation = "UnknownOperation";
        error = null;
        try
        {
            if (e == null)
            {
                error = new ArgumentNullException("e");
                return false;
            }
            object value = e.Operation;
            operation = value == null ? string.Empty : value.ToString();
            if (string.IsNullOrWhiteSpace(operation))
            {
                error = new InvalidOperationException("DocumentChanged operation was empty.");
                operation = "UnknownOperation";
                return false;
            }
            return true;
        }
        catch (Exception ex)
        {
            error = ex;
            operation = "UnknownOperation";
            return false;
        }
    }

    private static bool TryGetElementIds(Func<ICollection<ElementId>> getter, out ICollection<ElementId> ids, out Exception error)
    {
        ids = new List<ElementId>();
        error = null;
        try
        {
            ids = getter() ?? new List<ElementId>();
            return true;
        }
        catch (Exception ex)
        {
            error = ex;
            return false;
        }
    }

    private static void AddIds(ISet<string> target, IEnumerable<ElementId> ids)
    {
        foreach (ElementId id in ids ?? Enumerable.Empty<ElementId>())
        {
            string value = SafeElementId(id);
            if (!string.IsNullOrWhiteSpace(value))
            {
                target.Add(value);
            }
        }
    }

    private static bool ResolvePolicyEnabled(string workspaceRoot, bool force)
    {
        bool usedCachedFallback;
        return ResolvePolicyEnabledCore(workspaceRoot, force, out usedCachedFallback);
    }

    private static bool ResolveDocumentPolicyEnabled(string workspaceRoot, Document doc, bool force)
    {
        bool usedCachedFallback;
        return ResolveDocumentPolicyEnabledCore(workspaceRoot, doc, force, out usedCachedFallback);
    }

    private static bool ResolveDocumentPolicyEnabledCore(string workspaceRoot, Document doc, bool force, out bool usedCachedFallback)
    {
        bool globallyEnabled = ResolvePolicyEnabledCore(workspaceRoot, force, out usedCachedFallback);
        if (!globallyEnabled)
        {
            return false;
        }
        FamilyBrowserStandardPolicy policy;
        lock (SyncRoot)
        {
            policy = _cachedPolicy;
        }
        if (policy == null)
        {
            usedCachedFallback = true;
            return false;
        }
        try
        {
            return FamilyBrowserSecurityPolicyService.IsProjectElementTrackingScopeEnabled(policy, BuildTrackingProjectContext(doc));
        }
        catch (Exception ex)
        {
            usedCachedFallback = true;
            WriteTrackingDiagnostic(workspaceRoot, "Element tracking file scope evaluation failed", ex, doc);
            return false;
        }
    }

    private static FamilyBrowserProjectPolicyContext BuildTrackingProjectContext(Document doc)
    {
        bool workshared;
        bool worksharingKnown = TryGetIsWorkshared(doc, out workshared);
        return new FamilyBrowserProjectPolicyContext
        {
            ProjectTitle = SafeDocumentTitle(doc),
            ModelPath = SafeLocalDocumentPath(doc),
            CentralPath = SafeProjectIdentityPath(doc),
            IsWorkshared = worksharingKnown && workshared
        };
    }

    private static bool ResolvePolicyEnabledCore(string workspaceRoot, bool force, out bool usedCachedFallback)
    {
        usedCachedFallback = false;
        string root = workspaceRoot ?? string.Empty;
        lock (SyncRoot)
        {
            if (!force && _cachedPolicy != null && string.Equals(root, _cachedPolicyWorkspaceRoot, StringComparison.OrdinalIgnoreCase) && DateTime.UtcNow - _cachedPolicyAtUtc < TimeSpan.FromSeconds(PolicyRefreshSeconds))
            {
                return _cachedPolicyEnabled;
            }
        }
        bool enabled;
        FamilyBrowserStandardPolicy loadedPolicy;
        try
        {
            loadedPolicy = FamilyBrowserStandardPolicyStore.LoadOrCreate(root, Environment.UserName);
            enabled = FamilyBrowserStandardPolicyStore.IsProjectElementChangeTrackingEnabled(loadedPolicy);
        }
        catch (Exception ex)
        {
            WriteTrackingDiagnostic(root, "Element tracking policy read failed", ex, null);
            lock (SyncRoot)
            {
                if (string.Equals(root, _cachedPolicyWorkspaceRoot, StringComparison.OrdinalIgnoreCase) && _cachedPolicyKnown)
                {
                    usedCachedFallback = true;
                    return _cachedPolicyEnabled;
                }
            }
            return false;
        }
        bool disableDeferred = false;
        lock (SyncRoot)
        {
            if (!enabled && Sessions.Values.Any(delegate(DocumentSession session)
            {
                return session != null &&
                    string.Equals(session.WorkspaceRoot ?? string.Empty, root, StringComparison.OrdinalIgnoreCase) &&
                    HasUncommittedSessionEvidence(session);
            }))
            {
                foreach (DocumentSession session in Sessions.Values.Where(delegate(DocumentSession item)
                {
                    return item != null &&
                        string.Equals(item.WorkspaceRoot ?? string.Empty, root, StringComparison.OrdinalIgnoreCase) &&
                        HasUncommittedSessionEvidence(item);
                }))
                {
                    session.PolicyDisableDeferred = true;
                }
                DisableLiveTrackingNoLock(root, true);
                _cachedPolicyWorkspaceRoot = root;
                _cachedPolicyEnabled = true;
                _cachedPolicyKnown = true;
                _cachedPolicyAtUtc = DateTime.MinValue;
                _cachedPolicy = loadedPolicy;
                usedCachedFallback = true;
                disableDeferred = true;
            }
            else
            {
                _cachedPolicyWorkspaceRoot = root;
                _cachedPolicyEnabled = enabled;
                _cachedPolicyKnown = true;
                _cachedPolicyAtUtc = DateTime.UtcNow;
                _cachedPolicy = loadedPolicy;
                if (enabled)
                {
                    EnableLiveTrackingNoLock(root);
                }
                else
                {
                    DisableLiveTrackingNoLock(root, false);
                }
            }
        }
        if (disableDeferred)
        {
            WriteTrackingDiagnostic(
                root,
                "Element tracking disable deferred until commit boundary",
                new InvalidOperationException("The shared policy was disabled while this Revit client still held observed activity that had not reached a successful Save or Synchronize boundary. Existing evidence remains tracked until that boundary; no new session will start afterward."),
                null);
        }
        return enabled || disableDeferred;
    }

    private static void WriteTrackingDiagnostic(string workspaceRoot, string caption, Exception error, Document doc)
    {
        string project = doc == null ? string.Empty : SafeProjectIdentityPath(doc);
        string key = (caption ?? string.Empty) + "|" + project;
        lock (SyncRoot)
        {
            DateTime last;
            if (LastDiagnosticUtcByKey.TryGetValue(key, out last) && DateTime.UtcNow - last < TimeSpan.FromMinutes(1.0))
            {
                return;
            }
            LastDiagnosticUtcByKey[key] = DateTime.UtcNow;
        }
        try
        {
            FamilyBrowserErrorHelp.WriteLog(
                workspaceRoot ?? string.Empty,
                caption ?? "Element tracking failure",
                error ?? new InvalidOperationException("Unknown element tracking failure."),
                "Project=" + project + Environment.NewLine + "Tracking data was not fabricated; review the next successful Save/Sync and the local pending queue.");
        }
        catch
        {
        }
    }

    private static string BuildCommitPerformanceDetail(
        string baselineMode,
        long changedStateRefreshMilliseconds,
        long persistenceMilliseconds,
        long baselineMilliseconds,
        int locallyPendingElementCount,
        int incomingElementCount,
        bool projectCatalogObservationRequired)
    {
        return "baselineMode=" + (baselineMode ?? string.Empty) +
            ";changedStateRefreshMs=" + Math.Max(0L, changedStateRefreshMilliseconds).ToString(CultureInfo.InvariantCulture) +
            ";persistenceMs=" + Math.Max(0L, persistenceMilliseconds).ToString(CultureInfo.InvariantCulture) +
            ";managedPublish=deferred-after-local-durability" +
            ";baselineMs=" + Math.Max(0L, baselineMilliseconds).ToString(CultureInfo.InvariantCulture) +
            ";localIds=" + Math.Max(0, locallyPendingElementCount).ToString(CultureInfo.InvariantCulture) +
            ";incomingIds=" + Math.Max(0, incomingElementCount).ToString(CultureInfo.InvariantCulture) +
            ";catalogScan=" + (projectCatalogObservationRequired ? "required" : "skipped");
    }

    private static void WriteTrackingPerformance(
        Document doc,
        string commitKind,
        string stage,
        long elapsedMilliseconds,
        string detail)
    {
        try
        {
            string folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "KKY",
                "FamilyBrowser",
                "Diagnostics");
            Directory.CreateDirectory(folder);
            string path = Path.Combine(folder, "element-tracking-performance.log");
            string project = SafeProjectIdentityPath(doc);
            string line = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture) +
                "|stage=" + SanitizePerformanceValue(stage) +
                "|commit=" + SanitizePerformanceValue(commitKind) +
                "|elapsedMs=" + Math.Max(0L, elapsedMilliseconds).ToString(CultureInfo.InvariantCulture) +
                "|project=" + SanitizePerformanceValue(project) +
                "|" + SanitizePerformanceValue(detail) +
                Environment.NewLine;
            lock (PerformanceLogSyncRoot)
            {
                if (File.Exists(path) && new FileInfo(path).Length > 2L * 1024L * 1024L)
                {
                    File.WriteAllText(path, string.Empty, new UTF8Encoding(false));
                }
                File.AppendAllText(path, line, new UTF8Encoding(false));
            }
        }
        catch
        {
        }
    }

    private static string SanitizePerformanceValue(string value)
    {
        return (value ?? string.Empty)
            .Replace("\r", " ")
            .Replace("\n", " ")
            .Replace("|", "/");
    }

    private static bool CanTrackDocument(Document doc)
    {
        bool canTrack;
        Exception ignored;
        return TryCanTrackDocument(doc, out canTrack, out ignored) && canTrack;
    }

    private static bool TryCanTrackDocument(Document doc, out bool canTrack, out Exception error)
    {
        canTrack = false;
        error = null;
        if (doc == null)
        {
            return true;
        }
        try
        {
            canTrack = !doc.IsFamilyDocument;
            return true;
        }
        catch (Exception ex)
        {
            error = ex;
            return false;
        }
    }

    private static bool HasUncommittedSessionEvidence(DocumentSession session)
    {
        if (session == null)
        {
            return false;
        }
        bool lateBaselineWithoutProtectedCheckpoint = session.BaselineCapturedLate &&
            (session.RecoveredLocalSaveCommits == null || session.RecoveredLocalSaveCommits.Count == 0);
        return lateBaselineWithoutProtectedCheckpoint ||
            (session.TouchedIds != null && session.TouchedIds.Count > 0) ||
            (session.AppliedActivities != null && session.AppliedActivities.Count > 0) ||
            (session.AmbiguousActivityIds != null && session.AmbiguousActivityIds.Count > 0) ||
            session.EventReadFailureCount > 0 ||
            session.CommitBoundaryReadFailureCount > 0 ||
            session.UnmatchedUndoCount > 0 ||
            session.UnmatchedRedoCount > 0 ||
            session.ExternalRebaseFailed;
    }

    private static bool HasRetainedSessionEvidence(Document doc)
    {
        if (doc == null)
        {
            return false;
        }
        lock (SyncRoot)
        {
            DocumentSession session;
            return Sessions.TryGetValue(BuildRuntimeKey(doc), out session) &&
                (HasUncommittedSessionEvidence(session) || HasProtectedRecoveryEvidence(session));
        }
    }

    private static void MarkCommitBoundaryProtectionFailed(DocumentSession session)
    {
        lock (SyncRoot)
        {
            if (session != null && HasUncommittedSessionEvidence(session))
            {
                session.CommitBoundaryProtectionFailed = true;
            }
        }
    }

    private static bool HasProtectedRecoveryEvidence(DocumentSession session)
    {
        return session != null &&
            session.RecoveredLocalSaveCommits != null &&
            session.RecoveredLocalSaveCommits.Count > 0;
    }

    private static bool HasPendingRecoveryCheckpoint(string workspaceRoot, Document doc)
    {
        bool workshared;
        if (!TryGetIsWorkshared(doc, out workshared))
        {
            WriteTrackingDiagnostic(workspaceRoot, "Element tracking recovery worksharing state lookup failed", new InvalidOperationException("The document worksharing state is unknown, so recovery remains fail-closed."), doc);
            return true;
        }
        if (!workshared)
        {
            return false;
        }
        try
        {
            string projectIdentity = SafeProjectIdentityPath(doc);
            if (string.IsNullOrWhiteSpace(projectIdentity))
            {
                return false;
            }
            FamilyBrowserElementSessionCheckpointCountResult status =
                FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointStatus(workspaceRoot, projectIdentity);
            if (status.LockUnavailable)
            {
                WriteTrackingDiagnostic(workspaceRoot, "Element tracking recovery checkpoint is busy; recovery remains fail-closed", null, doc);
                return true;
            }
            return status.Count > 0;
        }
        catch (Exception ex)
        {
            WriteTrackingDiagnostic(workspaceRoot, "Element tracking recovery checkpoint lookup failed", ex, doc);
            return true;
        }
    }

    private static void EnableLiveTrackingNoLock(string workspaceRoot)
    {
        string root = workspaceRoot ?? string.Empty;
        foreach (DocumentSession session in Sessions.Values.Where(delegate(DocumentSession item)
        {
            return item != null && string.Equals(item.WorkspaceRoot ?? string.Empty, root, StringComparison.OrdinalIgnoreCase);
        }))
        {
            session.RecoveryOnly = false;
            session.PolicyDisableDeferred = false;
        }
    }

    private static void DisableLiveTrackingNoLock(string workspaceRoot, bool preserveUncommittedEvidence)
    {
        string root = workspaceRoot ?? string.Empty;
        HashSet<string> removedKeys = new HashSet<string>(
            Sessions.Where(delegate(KeyValuePair<string, DocumentSession> pair)
            {
                DocumentSession session = pair.Value;
                if (session == null || !string.Equals(session.WorkspaceRoot ?? string.Empty, root, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
                return !HasProtectedRecoveryEvidence(session) &&
                    !(preserveUncommittedEvidence && HasUncommittedSessionEvidence(session));
            }).Select(delegate(KeyValuePair<string, DocumentSession> pair) { return pair.Key; }),
            StringComparer.Ordinal);
        foreach (string key in removedKeys)
        {
            Sessions.Remove(key);
        }
        foreach (DocumentSession session in Sessions.Values.Where(delegate(DocumentSession item)
        {
            return item != null && string.Equals(item.WorkspaceRoot ?? string.Empty, root, StringComparison.OrdinalIgnoreCase);
        }))
        {
            if (preserveUncommittedEvidence && HasUncommittedSessionEvidence(session))
            {
                session.RecoveryOnly = false;
                session.PolicyDisableDeferred = true;
            }
            else
            {
                session.RecoveryOnly = true;
                session.PolicyDisableDeferred = false;
            }
        }
        foreach (int documentId in ClosingSessionKeysByDocumentId
            .Where(delegate(KeyValuePair<int, string> pair) { return removedKeys.Contains(pair.Value); })
            .Select(delegate(KeyValuePair<int, string> pair) { return pair.Key; })
            .ToList())
        {
            ClosingSessionKeysByDocumentId.Remove(documentId);
        }
        SynchronizingDocumentKeys.ExceptWith(removedKeys);
        ReloadingDocumentKeys.ExceptWith(removedKeys);
        if (!Sessions.Values.Any(delegate(DocumentSession session)
        {
            return session != null && string.Equals(session.WorkspaceRoot ?? string.Empty, root, StringComparison.OrdinalIgnoreCase);
        }))
        {
            UnknownSynchronizationStartRoots.Remove(NormalizeWorkspaceRootKey(root));
            UnknownReloadLatestStartRoots.Remove(NormalizeWorkspaceRootKey(root));
        }
    }

    private static bool TryGetIsWorkshared(Document doc, out bool workshared)
    {
        workshared = false;
        try
        {
            if (doc == null)
            {
                return false;
            }
            workshared = doc.IsWorkshared;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool SafeIsModified(Document doc)
    {
        try
        {
            return doc != null && doc.IsModified;
        }
        catch
        {
            return true;
        }
    }

    private static string SafeProjectIdentityPath(Document doc)
    {
        try
        {
            string identity = ProjectSnapshotStore.ResolveProjectIdentityPath(doc);
            if (!string.IsNullOrWhiteSpace(identity))
            {
                return identity;
            }
        }
        catch
        {
        }
        try
        {
            return doc == null ? string.Empty : doc.PathName ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeLocalDocumentPath(Document doc)
    {
        try
        {
            return doc == null ? string.Empty : doc.PathName ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeDocumentTitle(Document doc)
    {
        try
        {
            return doc == null ? string.Empty : doc.Title ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeRevitVersion(Document doc)
    {
        try
        {
            return doc == null || doc.Application == null ? string.Empty : doc.Application.VersionNumber ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeRevitUserName(Document doc)
    {
        try
        {
            return doc == null || doc.Application == null ? string.Empty : doc.Application.Username ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static bool SameCheckpointIdentity(string leftProject, string leftPath, string leftUser, string rightProject, string rightPath, string rightUser)
    {
        return SameProjectIdentity(leftProject, rightProject) && string.Equals(
                FamilyBrowserPathIdentityService.GetStablePathIdentity(leftPath),
                FamilyBrowserPathIdentityService.GetStablePathIdentity(rightPath),
                StringComparison.OrdinalIgnoreCase) &&
            string.Equals((leftUser ?? string.Empty).Trim(), (rightUser ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private static bool SameProjectIdentity(string leftProject, string rightProject)
    {
        string left = FamilyBrowserPathIdentityService.GetStablePathIdentity(leftProject);
        string right = FamilyBrowserPathIdentityService.GetStablePathIdentity(rightProject);
        return !string.IsNullOrWhiteSpace(left) && string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildRuntimeKey(Document doc)
    {
        if (doc == null)
        {
            return string.Empty;
        }
        return RuntimeDocumentIdentities.GetValue(doc, CreateRuntimeDocumentIdentity).Value;
    }

    private static RuntimeDocumentIdentity CreateRuntimeDocumentIdentity(Document doc)
    {
        // Revit can expose a different managed Document wrapper for each callback.
        string stableLocalPath = FamilyBrowserPathIdentityService.GetStablePathIdentity(SafeLocalDocumentPath(doc));
        if (!string.IsNullOrWhiteSpace(stableLocalPath))
        {
            return new RuntimeDocumentIdentity("document-local:" + HashText(stableLocalPath));
        }

        string stableProjectPath = FamilyBrowserPathIdentityService.GetStablePathIdentity(SafeProjectIdentityPath(doc));
        if (!string.IsNullOrWhiteSpace(stableProjectPath))
        {
            return new RuntimeDocumentIdentity("document-project:" + HashText(stableProjectPath));
        }

        string title = SafeDocumentTitle(doc).Trim();
        if (!string.IsNullOrWhiteSpace(title))
        {
            string unsavedIdentity = SafeRevitVersion(doc) + "|" + SafeRevitUserName(doc) + "|" + title;
            return new RuntimeDocumentIdentity("document-unsaved:" + HashText(unsavedIdentity));
        }

        return new RuntimeDocumentIdentity(string.Empty);
    }

    private static string HashText(string value)
    {
        using (SHA256 sha = SHA256.Create())
        {
            return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty))).Replace("-", string.Empty);
        }
    }
}
