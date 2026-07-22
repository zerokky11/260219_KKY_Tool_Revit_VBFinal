using System;
using System.Collections.Generic;
using System.IO;

public class FamilyBrowserOperationLogEntry
{
    public string EntryId { get; set; } = string.Empty;
    public string RecordedAtUtc { get; set; } = string.Empty;
    public string UserName { get; set; } = string.Empty;
    public string OperationKind { get; set; } = string.Empty;
    public string DocumentTitle { get; set; } = string.Empty;
    public string DocumentPath { get; set; } = string.Empty;
    public string SourceId { get; set; } = string.Empty;
    public string StandardSourceKey { get; set; } = string.Empty;
    public string StandardDisplayName { get; set; } = string.Empty;
    public string CandidateKind { get; set; } = string.Empty;
    public string CategoryName { get; set; } = string.Empty;
    public string FamilyName { get; set; } = string.Empty;
    public string TypeName { get; set; } = string.Empty;
    public string SystemFamilyKind { get; set; } = string.Empty;
    public string PlannedAction { get; set; } = string.Empty;
    public string Outcome { get; set; } = string.Empty;
    public string Details { get; set; } = string.Empty;
    public string CommitState { get; set; } = string.Empty;
    public string CommitKind { get; set; } = string.Empty;
    public string CommittedAtUtc { get; set; } = string.Empty;
}

public class StandardRvtChangeCandidateEntry
{
    public string EntryId { get; set; } = string.Empty;
    public string RecordedAtUtc { get; set; } = string.Empty;
    public string UserName { get; set; } = string.Empty;
    public string RevitUserName { get; set; } = string.Empty;
    public string MachineName { get; set; } = string.Empty;
    public string DocumentTitle { get; set; } = string.Empty;
    public string DocumentPath { get; set; } = string.Empty;
    public string CanonicalDocumentIdentity { get; set; } = string.Empty;
    public string SourceId { get; set; } = string.Empty;
    public string SlotKey { get; set; } = string.Empty;
    public string DisciplineKey { get; set; } = string.Empty;
    public string DisciplineLabel { get; set; } = string.Empty;
    public string CandidateKind { get; set; } = string.Empty;
    public string CategoryName { get; set; } = string.Empty;
    public string CategoryId { get; set; } = string.Empty;
    public string FamilyName { get; set; } = string.Empty;
    public string TypeName { get; set; } = string.Empty;
    public string SystemFamilyKind { get; set; } = string.Empty;
    public string ChangeKind { get; set; } = string.Empty;
    public string Reason { get; set; } = string.Empty;
    public string Details { get; set; } = string.Empty;
    public string BeforeFingerprint { get; set; } = string.Empty;
    public string AfterFingerprint { get; set; } = string.Empty;
    public string CommitState { get; set; } = string.Empty;
    public string CommitKind { get; set; } = string.Empty;
    public string CommittedAtUtc { get; set; } = string.Empty;
}

public class FamilyBrowserElementChangeItem
{
    public string ChangeKind { get; set; } = string.Empty;
    public string ElementId { get; set; } = string.Empty;
    public string UniqueId { get; set; } = string.Empty;
    public string ElementClass { get; set; } = string.Empty;
    public string CategoryName { get; set; } = string.Empty;
    public string ElementName { get; set; } = string.Empty;
    public string FamilyName { get; set; } = string.Empty;
    public string TypeName { get; set; } = string.Empty;
    public string TrackingKind { get; set; } = string.Empty;
    public string FirstObservedAtUtc { get; set; } = string.Empty;
    public string LastObservedAtUtc { get; set; } = string.Empty;
    public string ChangeSummary { get; set; } = string.Empty;
    public bool PreviousStateUnavailable { get; set; }
    public bool ExternalUpdateOverlap { get; set; }
    public FamilyBrowserTrackedElementState Before { get; set; }
    public FamilyBrowserTrackedElementState After { get; set; }
    public List<string> TransactionNames { get; set; } = new List<string>();
}

public class FamilyBrowserTrackedElementState
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
    public string TrackingKind { get; set; } = string.Empty;
    public string SharedParameterGuid { get; set; } = string.Empty;
    public string ParameterBindingKind { get; set; } = string.Empty;
    public string ParameterBoundCategories { get; set; } = string.Empty;
    public string ParameterBoundCategoryIds { get; set; } = string.Empty;
    public string ParameterGroup { get; set; } = string.Empty;
    public string ParameterDataType { get; set; } = string.Empty;
    public string ParameterVariesAcrossGroups { get; set; } = string.Empty;
    public string GridCurveSignature { get; set; } = string.Empty;
    public string GridExtentsSignature { get; set; } = string.Empty;
    public string GridPinnedState { get; set; } = string.Empty;
    public string StateSignature { get; set; } = string.Empty;
    public bool IsElementType { get; set; }
    public bool IsViewSpecific { get; set; }
}

public class FamilyBrowserElementChangeCommit
{
    public int SchemaVersion { get; set; } = 6;
    public string EntryId { get; set; } = string.Empty;
    public string ProjectTitle { get; set; } = string.Empty;
    public string ProjectIdentityPath { get; set; } = string.Empty;
    public string ProjectCanonicalPath { get; set; } = string.Empty;
    public string ProjectComparableIdentity { get; set; } = string.Empty;
    public string ProjectLegacyComparableIdentity { get; set; } = string.Empty;
    public string CommitKind { get; set; } = string.Empty;
    public string CommittedAtUtc { get; set; } = string.Empty;
    public string LocalSaveProtectedAtUtc { get; set; } = string.Empty;
    public string PublishedAtUtc { get; set; } = string.Empty;
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
    public int TransientCreatedDeletedCount { get; set; }
    public int ExternalUpdateOverlapCount { get; set; }
    public bool CoverageGapOnly { get; set; }
    public int EventReadFailureCount { get; set; }
    public int CommitBoundaryReadFailureCount { get; set; }
    public int IntegrityVersion { get; set; }
    public string IntegritySha256 { get; set; } = string.Empty;
    public List<string> TransactionNames { get; set; } = new List<string>();
    public List<FamilyBrowserElementChangeItem> Changes { get; set; } = new List<FamilyBrowserElementChangeItem>();
}

public static class FamilyBrowserPathIdentityService
{
    public static string GetComparableIdentity(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }
        if (File.Exists(path))
        {
            return "FILE-STUB:" + Path.GetFullPath(path).Replace('/', '\\').TrimEnd('\\').ToUpperInvariant();
        }
        try
        {
            return Path.GetFullPath(path).Replace('/', '\\').TrimEnd('\\').ToUpperInvariant();
        }
        catch
        {
            return path.Replace('/', '\\').Trim().TrimEnd('\\').ToUpperInvariant();
        }
    }

    public static string GetCanonicalPath(string path)
    {
        return NormalizePath(path);
    }

    public static string GetStablePathIdentity(string path)
    {
        string canonical = GetCanonicalPath(path);
        return string.IsNullOrWhiteSpace(canonical) ? string.Empty : "PATH:" + canonical.ToUpperInvariant();
    }

    public static string NormalizePath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }
        try
        {
            return Path.GetFullPath(path).Replace('/', '\\').TrimEnd('\\');
        }
        catch
        {
            return path.Replace('/', '\\').Trim().TrimEnd('\\');
        }
    }
}

public static class FamilyBrowserStandardPolicyStore
{
    public static bool ManagedAvailable { get; set; }
    public static string ManagedRoot { get; set; } = string.Empty;

    public static string GetConfiguredManagedPolicyPath()
    {
        return string.IsNullOrWhiteSpace(ManagedRoot) ? string.Empty : Path.Combine(ManagedRoot, "Config", "standard-policy.json");
    }

    public static bool IsManagedDataRootAvailable(string workspaceRoot = "")
    {
        return ManagedAvailable && !string.IsNullOrWhiteSpace(ManagedRoot);
    }

    public static string GetDataFolder(string workspaceRoot, string folderName)
    {
        return string.IsNullOrWhiteSpace(folderName) ? ManagedRoot : Path.Combine(ManagedRoot, folderName);
    }
}
