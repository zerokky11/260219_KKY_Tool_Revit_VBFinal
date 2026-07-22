using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using Autodesk.Revit.DB;

public sealed class FamilyBrowserProjectCatalogEntry
{
    public string EntryKind { get; set; }
    public string Key { get; set; }
    public string CategoryName { get; set; }
    public string CategoryId { get; set; }
    public string FamilyName { get; set; }
    public string TypeName { get; set; }
    public string TypeClassName { get; set; }

    public FamilyBrowserProjectCatalogEntry()
    {
        EntryKind = string.Empty;
        Key = string.Empty;
        CategoryName = string.Empty;
        CategoryId = string.Empty;
        FamilyName = string.Empty;
        TypeName = string.Empty;
        TypeClassName = string.Empty;
    }
}

public sealed class FamilyBrowserProjectCatalogSnapshot
{
    public int SchemaVersion { get; set; }
    public string ProjectTitle { get; set; }
    public string ProjectIdentityPath { get; set; }
    public string ProjectCanonicalPath { get; set; }
    public string ProjectComparableIdentity { get; set; }
    public string CapturedAtUtc { get; set; }
    public string Trigger { get; set; }
    public string RevitVersion { get; set; }
    public string RevitUserName { get; set; }
    public string WindowsUserName { get; set; }
    public string MachineName { get; set; }
    public long ElapsedMilliseconds { get; set; }
    public int FamilyCount { get; set; }
    public int FamilyTypeCount { get; set; }
    public int SystemTypeCount { get; set; }
    public string CatalogHash { get; set; }
    public List<FamilyBrowserProjectCatalogEntry> Entries { get; set; }

    public FamilyBrowserProjectCatalogSnapshot()
    {
        SchemaVersion = 1;
        ProjectTitle = string.Empty;
        ProjectIdentityPath = string.Empty;
        ProjectCanonicalPath = string.Empty;
        ProjectComparableIdentity = string.Empty;
        CapturedAtUtc = string.Empty;
        Trigger = string.Empty;
        RevitVersion = string.Empty;
        RevitUserName = string.Empty;
        WindowsUserName = string.Empty;
        MachineName = string.Empty;
        CatalogHash = string.Empty;
        Entries = new List<FamilyBrowserProjectCatalogEntry>();
    }
}

public sealed class FamilyBrowserProjectCatalogManifest
{
    public int SchemaVersion { get; set; }
    public string ProjectTitle { get; set; }
    public string ProjectIdentityPath { get; set; }
    public string ProjectCanonicalPath { get; set; }
    public string ProjectComparableIdentity { get; set; }
    public string AcceptedSnapshotPath { get; set; }
    public string AcceptedCatalogHash { get; set; }
    public string AcceptedAtUtc { get; set; }
    public string AcceptedBy { get; set; }
    public string LastObservedSnapshotPath { get; set; }
    public string LastObservedCatalogHash { get; set; }
    public string LastObservedAtUtc { get; set; }
    public string LastObservedTrigger { get; set; }

    public FamilyBrowserProjectCatalogManifest()
    {
        SchemaVersion = 1;
        ProjectTitle = string.Empty;
        ProjectIdentityPath = string.Empty;
        ProjectCanonicalPath = string.Empty;
        ProjectComparableIdentity = string.Empty;
        AcceptedSnapshotPath = string.Empty;
        AcceptedCatalogHash = string.Empty;
        AcceptedAtUtc = string.Empty;
        AcceptedBy = string.Empty;
        LastObservedSnapshotPath = string.Empty;
        LastObservedCatalogHash = string.Empty;
        LastObservedAtUtc = string.Empty;
        LastObservedTrigger = string.Empty;
    }
}

public sealed class FamilyBrowserProjectCatalogChange
{
    public string ChangeKind { get; set; }
    public string EntryKind { get; set; }
    public string Key { get; set; }
    public string CategoryName { get; set; }
    public string FamilyName { get; set; }
    public string TypeName { get; set; }
    public string TypeClassName { get; set; }
    public string Attribution { get; set; }
    public string OperationKind { get; set; }
    public string OperationUser { get; set; }
    public string OperationAtUtc { get; set; }

    public FamilyBrowserProjectCatalogChange()
    {
        ChangeKind = string.Empty;
        EntryKind = string.Empty;
        Key = string.Empty;
        CategoryName = string.Empty;
        FamilyName = string.Empty;
        TypeName = string.Empty;
        TypeClassName = string.Empty;
        Attribution = "ExternalUntracked";
        OperationKind = string.Empty;
        OperationUser = string.Empty;
        OperationAtUtc = string.Empty;
    }
}

public sealed class FamilyBrowserProjectCatalogState
{
    public int SchemaVersion { get; set; }
    public string StateCode { get; set; }
    public string ProjectTitle { get; set; }
    public string ProjectIdentityPath { get; set; }
    public string ProjectComparableIdentity { get; set; }
    public string CheckedAtUtc { get; set; }
    public string AcceptedAtUtc { get; set; }
    public string AcceptedBy { get; set; }
    public string AcceptedCatalogHash { get; set; }
    public string CurrentCatalogHash { get; set; }
    public string CurrentSnapshotPath { get; set; }
    public string Trigger { get; set; }
    public long ElapsedMilliseconds { get; set; }
    public int FamilyCount { get; set; }
    public int FamilyTypeCount { get; set; }
    public int SystemTypeCount { get; set; }
    public int AddedCount { get; set; }
    public int RemovedCount { get; set; }
    public int BrowserTrackedChangeCount { get; set; }
    public int ExternalUntrackedChangeCount { get; set; }
    public int RecentAddedCount { get; set; }
    public int RecentRemovedCount { get; set; }
    public int RecentUntrackedChangeCount { get; set; }
    public string Reason { get; set; }
    public string ErrorMessage { get; set; }
    public List<FamilyBrowserProjectCatalogChange> Changes { get; set; }
    public List<FamilyBrowserProjectCatalogChange> RecentChanges { get; set; }

    public bool BaselineMissing
    {
        get { return string.Equals(StateCode, "BaselineMissing", StringComparison.OrdinalIgnoreCase); }
    }

    public bool Changed
    {
        get { return string.Equals(StateCode, "Changed", StringComparison.OrdinalIgnoreCase); }
    }

    public FamilyBrowserProjectCatalogState()
    {
        SchemaVersion = 1;
        StateCode = "NotChecked";
        ProjectTitle = string.Empty;
        ProjectIdentityPath = string.Empty;
        ProjectComparableIdentity = string.Empty;
        CheckedAtUtc = string.Empty;
        AcceptedAtUtc = string.Empty;
        AcceptedBy = string.Empty;
        AcceptedCatalogHash = string.Empty;
        CurrentCatalogHash = string.Empty;
        CurrentSnapshotPath = string.Empty;
        Trigger = string.Empty;
        Reason = string.Empty;
        ErrorMessage = string.Empty;
        Changes = new List<FamilyBrowserProjectCatalogChange>();
        RecentChanges = new List<FamilyBrowserProjectCatalogChange>();
    }
}

public static class FamilyBrowserProjectCatalogService
{
    private const int SchemaVersion = 1;
    private static readonly object SyncRoot = new object();
    private static readonly HashSet<string> AllowedSystemTypeNames = new HashSet<string>(new[]
    {
        "WallType", "FloorType", "RoofType", "CeilingType", "StairsType", "RailingType", "CurtainSystemType", "PanelType",
        "DuctType", "PipeType", "FlexDuctType", "FlexPipeType", "DuctSystemType", "PipingSystemType", "MechanicalSystemType",
        "ElectricalSystemType", "CableTrayType", "ConduitType", "WireType", "DuctInsulationType", "PipeInsulationType",
        "DuctLiningType", "MullionType"
    }, StringComparer.OrdinalIgnoreCase);

    public static bool IsSuccessfulRevitEventStatus(object status)
    {
        return status != null && string.Equals(status.ToString(), "Succeeded", StringComparison.OrdinalIgnoreCase);
    }

    public static FamilyBrowserProjectCatalogSnapshot Capture(Document doc, string trigger)
    {
        if (doc == null)
        {
            throw new ArgumentNullException("doc");
        }
        Stopwatch stopwatch = Stopwatch.StartNew();
        FamilyBrowserProjectCatalogSnapshot snapshot = CreateSnapshotShell(doc, trigger);
        Dictionary<string, FamilyBrowserProjectCatalogEntry> entries = new Dictionary<string, FamilyBrowserProjectCatalogEntry>(StringComparer.Ordinal);

        foreach (Family family in new FilteredElementCollector(doc).OfClass(typeof(Family)).Cast<Family>())
        {
            if (family == null || IsInPlaceFamily(family))
            {
                continue;
            }
            string familyName = SafeElementName(family);
            string categoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family);
            string categoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family);
            AddEntry(entries, CreateEntry("Family", categoryName, categoryId, familyName, string.Empty, string.Empty));
            ICollection<ElementId> symbolIds = null;
            try
            {
                symbolIds = family.GetFamilySymbolIds();
            }
            catch
            {
                symbolIds = null;
            }
            if (symbolIds == null)
            {
                continue;
            }
            foreach (ElementId symbolId in symbolIds)
            {
                ElementType symbol = null;
                try
                {
                    symbol = doc.GetElement(symbolId) as ElementType;
                }
                catch
                {
                    symbol = null;
                }
                if (symbol != null)
                {
                    AddEntry(entries, CreateEntry("FamilyType", categoryName, categoryId, familyName, SafeElementName(symbol), symbol.GetType().Name));
                }
            }
        }

        foreach (ElementType elementType in new FilteredElementCollector(doc).WhereElementIsElementType().Cast<ElementType>())
        {
            if (elementType == null)
            {
                continue;
            }
            string className = elementType.GetType().Name;
            if (!AllowedSystemTypeNames.Contains(className))
            {
                continue;
            }
            if (elementType is FamilySymbol && !SystemTypeDetailedComponentSnapshotService.SupportsRequiredCurtainPanelComponents(className))
            {
                continue;
            }
            string categoryName = elementType.Category == null ? string.Empty : elementType.Category.Name ?? string.Empty;
            string categoryId = elementType.Category == null ? string.Empty : elementType.Category.Id.CompatIntegerValue().ToString(CultureInfo.InvariantCulture);
            AddEntry(entries, CreateEntry("SystemType", categoryName, categoryId, string.Empty, SafeElementName(elementType), className));
        }

        snapshot.Entries = entries.Values.OrderBy(delegate(FamilyBrowserProjectCatalogEntry x) { return x.Key; }, StringComparer.Ordinal).ToList();
        FinalizeSnapshot(snapshot, stopwatch.ElapsedMilliseconds);
        return snapshot;
    }

    public static FamilyBrowserProjectCatalogState Observe(string workspaceRoot, Document doc, string trigger)
    {
        if (doc == null)
        {
            return CreateUnavailableState("NoProject", "No active project is available.");
        }
        string publicationReason;
        if (!ProjectSnapshotStore.CanPublishSharedProjectState(doc, out publicationReason))
        {
            return CreatePublicationDeferredState(doc, null, trigger, publicationReason);
        }
        try
        {
            FamilyBrowserProjectCatalogSnapshot snapshot = Capture(doc, trigger);
            return PersistSnapshot(workspaceRoot, doc, snapshot, false, string.Empty);
        }
        catch (Exception ex)
        {
            return CreateErrorState(doc, trigger, ex);
        }
    }

    public static FamilyBrowserProjectCatalogState AcceptCurrent(string workspaceRoot, Document doc, string trigger, string acceptedBy)
    {
        if (doc == null)
        {
            return CreateUnavailableState("NoProject", "No active project is available.");
        }
        string publicationReason;
        if (!ProjectSnapshotStore.CanPublishSharedProjectState(doc, out publicationReason))
        {
            return CreatePublicationDeferredState(doc, null, trigger, publicationReason);
        }
        try
        {
            FamilyBrowserProjectCatalogSnapshot snapshot = Capture(doc, trigger);
            return PersistSnapshot(workspaceRoot, doc, snapshot, true, acceptedBy);
        }
        catch (Exception ex)
        {
            return CreateErrorState(doc, trigger, ex);
        }
    }

    public static FamilyBrowserProjectCatalogState AcceptFromProjectSnapshot(string workspaceRoot, Document doc, ProjectContentSnapshot source, string trigger, string acceptedBy)
    {
        if (doc == null || source == null)
        {
            return CreateUnavailableState("NoProject", "No completed Current Model Check snapshot is available.");
        }
        string publicationReason;
        if (!ProjectSnapshotStore.CanPublishSharedProjectState(doc, out publicationReason))
        {
            return CreatePublicationDeferredState(doc, null, trigger, publicationReason);
        }
        try
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            FamilyBrowserProjectCatalogSnapshot snapshot = CreateSnapshotShell(doc, trigger);
            Dictionary<string, FamilyBrowserProjectCatalogEntry> entries = new Dictionary<string, FamilyBrowserProjectCatalogEntry>(StringComparer.Ordinal);
            foreach (ProjectLoadableFamilySnapshotItem family in source.LoadableFamilies ?? new List<ProjectLoadableFamilySnapshotItem>())
            {
                if (family == null)
                {
                    continue;
                }
                AddEntry(entries, CreateEntry("Family", family.CategoryName, family.CategoryId, family.FamilyName, string.Empty, string.Empty));
                foreach (string typeName in family.TypeNames ?? new List<string>())
                {
                    AddEntry(entries, CreateEntry("FamilyType", family.CategoryName, family.CategoryId, family.FamilyName, typeName, "FamilySymbol"));
                }
            }
            foreach (ProjectSystemTypeSnapshotItem systemType in source.SystemTypes ?? new List<ProjectSystemTypeSnapshotItem>())
            {
                if (systemType != null)
                {
                    AddEntry(entries, CreateEntry("SystemType", systemType.CategoryName, systemType.CategoryId, string.Empty, systemType.TypeName, systemType.TypeClassName));
                }
            }
            snapshot.Entries = entries.Values.OrderBy(delegate(FamilyBrowserProjectCatalogEntry x) { return x.Key; }, StringComparer.Ordinal).ToList();
            FinalizeSnapshot(snapshot, stopwatch.ElapsedMilliseconds);
            return PersistSnapshot(workspaceRoot, doc, snapshot, true, acceptedBy);
        }
        catch (Exception ex)
        {
            return CreateErrorState(doc, trigger, ex);
        }
    }

    public static FamilyBrowserProjectCatalogState LoadLatestState(string workspaceRoot, Document doc)
    {
        if (doc == null)
        {
            return CreateUnavailableState("NoProject", "No active project is available.");
        }
        string identityPath = ProjectSnapshotStore.ResolveProjectIdentityPath(doc);
        if (string.IsNullOrWhiteSpace(identityPath))
        {
            return CreateUnavailableState("NoIdentity", "Save the project before enabling project catalog tracking.");
        }
        try
        {
            string folder = ResolveProjectFolder(workspaceRoot, identityPath);
            if (!Directory.Exists(folder))
            {
                FamilyBrowserProjectCatalogState notChecked = CreateUnavailableState("NotChecked", "The project catalog has not been checked yet.");
                notChecked.ProjectTitle = SafeDocumentTitle(doc);
                notChecked.ProjectIdentityPath = identityPath;
                notChecked.ProjectComparableIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(identityPath);
                return notChecked;
            }
            lock (SyncRoot)
            {
                using (FileStream catalogLock = AcquireCatalogLock(folder))
                {
                    string manifestPath = Path.Combine(folder, "manifest.json");
                    string statePath = Path.Combine(folder, "state.json");
                    FamilyBrowserProjectCatalogManifest manifest = LoadJson<FamilyBrowserProjectCatalogManifest>(manifestPath);
                    FamilyBrowserProjectCatalogState state = LoadJson<FamilyBrowserProjectCatalogState>(statePath);
                    if ((File.Exists(manifestPath) && manifest == null) || (File.Exists(statePath) && state == null))
                    {
                        throw new InvalidDataException("The project catalog manifest or state file is unreadable. Existing tracking data was preserved and was not replaced.");
                    }
                    if (manifest != null)
                    {
                        ValidateManifest(manifest, identityPath);
                        if (state == null ||
                            !string.Equals(NormalizeSnapshotReference(state.CurrentSnapshotPath), NormalizeSnapshotReference(manifest.LastObservedSnapshotPath), StringComparison.OrdinalIgnoreCase) ||
                            !string.Equals(state.CurrentCatalogHash ?? string.Empty, manifest.LastObservedCatalogHash ?? string.Empty, StringComparison.Ordinal) ||
                            !string.Equals(state.AcceptedCatalogHash ?? string.Empty, manifest.AcceptedCatalogHash ?? string.Empty, StringComparison.Ordinal))
                        {
                            throw new InvalidDataException("The project catalog state does not match its manifest. Refresh or restore the managed tracking data before accepting a new baseline.");
                        }
                    }
                    if (state == null)
                    {
                        state = CreateUnavailableState("NotChecked", "The project catalog has not been checked yet.");
                        state.ProjectTitle = SafeDocumentTitle(doc);
                        state.ProjectIdentityPath = identityPath;
                        state.ProjectComparableIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(identityPath);
                    }
                    return state;
                }
            }
        }
        catch (Exception ex)
        {
            return CreateErrorState(doc, "LoadLatestState", ex);
        }
    }

    public static bool IsPublishedObservationState(FamilyBrowserProjectCatalogState state)
    {
        if (state == null || !string.IsNullOrWhiteSpace(state.ErrorMessage))
        {
            return false;
        }
        return string.Equals(state.StateCode, "Current", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(state.StateCode, "Changed", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(state.StateCode, "BaselineMissing", StringComparison.OrdinalIgnoreCase);
    }

    public static string GetProjectFolder(string workspaceRoot, string identityPath)
    {
        string stableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(identityPath);
        if (string.IsNullOrWhiteSpace(stableIdentity))
        {
            stableIdentity = "PATH:" + FamilyBrowserPathIdentityService.NormalizePath(identityPath).ToUpperInvariant();
        }
        return Path.Combine(FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "ProjectCatalogs"), HashText(stableIdentity).Substring(0, 32));
    }

    private static FamilyBrowserProjectCatalogState PersistSnapshot(string workspaceRoot, Document doc, FamilyBrowserProjectCatalogSnapshot snapshot, bool accept, string acceptedBy)
    {
        if (snapshot == null || string.IsNullOrWhiteSpace(snapshot.ProjectIdentityPath))
        {
            FamilyBrowserProjectCatalogState noIdentity = CreateUnavailableState("NoIdentity", "Save the project before enabling project catalog tracking.");
            if (snapshot != null)
            {
                CopySnapshotSummary(snapshot, noIdentity);
            }
            return noIdentity;
        }
        if (!FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
        {
            FamilyBrowserProjectCatalogState unavailable = CreateUnavailableState("StorageUnavailable", "The managed data folder is not available.");
            CopySnapshotSummary(snapshot, unavailable);
            return unavailable;
        }

        string folder = ResolveProjectFolder(workspaceRoot, snapshot.ProjectIdentityPath);
        Directory.CreateDirectory(folder);
        lock (SyncRoot)
        {
            using (FileStream catalogLock = AcquireCatalogLock(folder))
            {
                string publicationReason;
                if (!ProjectSnapshotStore.CanPublishSharedProjectState(doc, out publicationReason))
                {
                    return CreatePublicationDeferredState(doc, snapshot, snapshot.Trigger, publicationReason);
                }
                string manifestPath = Path.Combine(folder, "manifest.json");
                FamilyBrowserProjectCatalogManifest manifest = LoadJson<FamilyBrowserProjectCatalogManifest>(manifestPath);
                if (File.Exists(manifestPath) && manifest == null)
                {
                    throw new InvalidDataException("The project catalog manifest is unreadable. Existing tracking data was preserved and was not replaced.");
                }
                manifest = manifest ?? new FamilyBrowserProjectCatalogManifest();
                ValidateManifest(manifest, snapshot.ProjectIdentityPath);
                FamilyBrowserProjectCatalogSnapshot previousObserved = LoadSnapshot(
                    folder,
                    manifest.LastObservedSnapshotPath,
                    manifest.LastObservedCatalogHash,
                    snapshot.ProjectIdentityPath,
                    false);
                string snapshotPath = SaveSnapshot(folder, snapshot);
                if (accept)
                {
                    manifest.AcceptedSnapshotPath = snapshotPath;
                    manifest.AcceptedCatalogHash = snapshot.CatalogHash;
                    manifest.AcceptedAtUtc = snapshot.CapturedAtUtc;
                    manifest.AcceptedBy = string.IsNullOrWhiteSpace(acceptedBy) ? Environment.UserName : acceptedBy.Trim();
                }
                FamilyBrowserProjectCatalogSnapshot accepted = LoadSnapshot(
                    folder,
                    manifest.AcceptedSnapshotPath,
                    manifest.AcceptedCatalogHash,
                    snapshot.ProjectIdentityPath,
                    !string.IsNullOrWhiteSpace(manifest.AcceptedSnapshotPath));
                manifest.SchemaVersion = SchemaVersion;
                manifest.ProjectTitle = snapshot.ProjectTitle;
                manifest.ProjectIdentityPath = snapshot.ProjectIdentityPath;
                manifest.ProjectCanonicalPath = snapshot.ProjectCanonicalPath;
                manifest.ProjectComparableIdentity = snapshot.ProjectComparableIdentity;
                manifest.LastObservedSnapshotPath = snapshotPath;
                manifest.LastObservedCatalogHash = snapshot.CatalogHash;
                manifest.LastObservedAtUtc = snapshot.CapturedAtUtc;
                manifest.LastObservedTrigger = snapshot.Trigger;

                List<FamilyBrowserOperationLogEntry> operations = LoadCommittedOperations(workspaceRoot, snapshot.ProjectComparableIdentity, manifest.AcceptedAtUtc);
                FamilyBrowserProjectCatalogState state = BuildState(snapshot, accepted, previousObserved, manifest, operations);
                if (!ProjectSnapshotStore.CanPublishSharedProjectState(doc, out publicationReason))
                {
                    return CreatePublicationDeferredState(doc, snapshot, snapshot.Trigger, publicationReason);
                }
                SaveJson(manifestPath, manifest);
                SaveJson(Path.Combine(folder, "state.json"), state);
                return state;
            }
        }
    }

    private static FamilyBrowserProjectCatalogState BuildState(
        FamilyBrowserProjectCatalogSnapshot current,
        FamilyBrowserProjectCatalogSnapshot accepted,
        FamilyBrowserProjectCatalogSnapshot previousObserved,
        FamilyBrowserProjectCatalogManifest manifest,
        IList<FamilyBrowserOperationLogEntry> operations)
    {
        FamilyBrowserProjectCatalogState state = new FamilyBrowserProjectCatalogState();
        CopySnapshotSummary(current, state);
        state.AcceptedAtUtc = manifest.AcceptedAtUtc ?? string.Empty;
        state.AcceptedBy = manifest.AcceptedBy ?? string.Empty;
        state.AcceptedCatalogHash = manifest.AcceptedCatalogHash ?? string.Empty;
        state.CurrentSnapshotPath = manifest.LastObservedSnapshotPath ?? string.Empty;
        if (accepted == null || string.IsNullOrWhiteSpace(manifest.AcceptedCatalogHash))
        {
            state.StateCode = "BaselineMissing";
            state.Reason = "No accepted project catalog baseline is available. Run Current Model Check or accept the current name catalog as the baseline.";
        }
        else
        {
            state.Changes = BuildChanges(accepted, current, operations);
            state.AddedCount = state.Changes.Count(delegate(FamilyBrowserProjectCatalogChange x) { return string.Equals(x.ChangeKind, "Added", StringComparison.OrdinalIgnoreCase); });
            state.RemovedCount = state.Changes.Count(delegate(FamilyBrowserProjectCatalogChange x) { return string.Equals(x.ChangeKind, "Removed", StringComparison.OrdinalIgnoreCase); });
            state.BrowserTrackedChangeCount = state.Changes.Count(delegate(FamilyBrowserProjectCatalogChange x) { return string.Equals(x.Attribution, "KnownBrowser", StringComparison.OrdinalIgnoreCase); });
            state.ExternalUntrackedChangeCount = state.Changes.Count - state.BrowserTrackedChangeCount;
            state.StateCode = state.Changes.Count == 0 ? "Current" : "Changed";
            state.Reason = state.Changes.Count == 0
                ? "The current family and system type name catalog matches the accepted baseline."
                : "Family or type names differ from the accepted project catalog baseline.";
        }
        if (previousObserved != null)
        {
            state.RecentChanges = BuildChanges(previousObserved, current, operations);
            state.RecentAddedCount = state.RecentChanges.Count(delegate(FamilyBrowserProjectCatalogChange x) { return string.Equals(x.ChangeKind, "Added", StringComparison.OrdinalIgnoreCase); });
            state.RecentRemovedCount = state.RecentChanges.Count(delegate(FamilyBrowserProjectCatalogChange x) { return string.Equals(x.ChangeKind, "Removed", StringComparison.OrdinalIgnoreCase); });
            state.RecentUntrackedChangeCount = state.RecentChanges.Count(delegate(FamilyBrowserProjectCatalogChange x) { return !string.Equals(x.Attribution, "KnownBrowser", StringComparison.OrdinalIgnoreCase); });
        }
        return state;
    }

    private static List<FamilyBrowserProjectCatalogChange> BuildChanges(
        FamilyBrowserProjectCatalogSnapshot before,
        FamilyBrowserProjectCatalogSnapshot after,
        IList<FamilyBrowserOperationLogEntry> operations)
    {
        Dictionary<string, FamilyBrowserProjectCatalogEntry> beforeEntries = (before == null ? new List<FamilyBrowserProjectCatalogEntry>() : before.Entries ?? new List<FamilyBrowserProjectCatalogEntry>())
            .Where(delegate(FamilyBrowserProjectCatalogEntry x) { return x != null && !string.IsNullOrWhiteSpace(x.Key); })
            .GroupBy(delegate(FamilyBrowserProjectCatalogEntry x) { return x.Key; }, StringComparer.Ordinal)
            .ToDictionary(delegate(IGrouping<string, FamilyBrowserProjectCatalogEntry> x) { return x.Key; }, delegate(IGrouping<string, FamilyBrowserProjectCatalogEntry> x) { return x.First(); }, StringComparer.Ordinal);
        Dictionary<string, FamilyBrowserProjectCatalogEntry> afterEntries = (after == null ? new List<FamilyBrowserProjectCatalogEntry>() : after.Entries ?? new List<FamilyBrowserProjectCatalogEntry>())
            .Where(delegate(FamilyBrowserProjectCatalogEntry x) { return x != null && !string.IsNullOrWhiteSpace(x.Key); })
            .GroupBy(delegate(FamilyBrowserProjectCatalogEntry x) { return x.Key; }, StringComparer.Ordinal)
            .ToDictionary(delegate(IGrouping<string, FamilyBrowserProjectCatalogEntry> x) { return x.Key; }, delegate(IGrouping<string, FamilyBrowserProjectCatalogEntry> x) { return x.First(); }, StringComparer.Ordinal);
        List<FamilyBrowserProjectCatalogChange> result = new List<FamilyBrowserProjectCatalogChange>();
        foreach (FamilyBrowserProjectCatalogEntry entry in afterEntries.Values.Where(delegate(FamilyBrowserProjectCatalogEntry x) { return !beforeEntries.ContainsKey(x.Key); }))
        {
            FamilyBrowserProjectCatalogChange change = CreateChange("Added", entry);
            ApplyOperationAttribution(change, operations);
            result.Add(change);
        }
        foreach (FamilyBrowserProjectCatalogEntry entry in beforeEntries.Values.Where(delegate(FamilyBrowserProjectCatalogEntry x) { return !afterEntries.ContainsKey(x.Key); }))
        {
            result.Add(CreateChange("Removed", entry));
        }
        return result.OrderBy(delegate(FamilyBrowserProjectCatalogChange x) { return x.ChangeKind; }, StringComparer.Ordinal)
            .ThenBy(delegate(FamilyBrowserProjectCatalogChange x) { return x.Key; }, StringComparer.Ordinal)
            .ToList();
    }

    private static void ApplyOperationAttribution(FamilyBrowserProjectCatalogChange change, IList<FamilyBrowserOperationLogEntry> operations)
    {
        if (change == null || operations == null || !string.Equals(change.ChangeKind, "Added", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }
        FamilyBrowserOperationLogEntry matched = operations
            .Where(delegate(FamilyBrowserOperationLogEntry operation) { return OperationMatches(change, operation); })
            .OrderByDescending(delegate(FamilyBrowserOperationLogEntry operation) { return operation.CommittedAtUtc ?? operation.RecordedAtUtc ?? string.Empty; }, StringComparer.Ordinal)
            .FirstOrDefault();
        if (matched == null)
        {
            return;
        }
        change.Attribution = "KnownBrowser";
        change.OperationKind = matched.OperationKind ?? matched.PlannedAction ?? string.Empty;
        change.OperationUser = matched.UserName ?? string.Empty;
        change.OperationAtUtc = string.IsNullOrWhiteSpace(matched.CommittedAtUtc) ? matched.RecordedAtUtc ?? string.Empty : matched.CommittedAtUtc;
    }

    private static bool OperationMatches(FamilyBrowserProjectCatalogChange change, FamilyBrowserOperationLogEntry operation)
    {
        if (operation == null || !string.Equals(operation.CommitState, "Committed", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }
        if (string.Equals(change.EntryKind, "Family", StringComparison.OrdinalIgnoreCase))
        {
            return string.Equals(operation.CandidateKind, "LoadableFamily", StringComparison.OrdinalIgnoreCase)
                && string.Equals(Normalize(operation.FamilyName), Normalize(change.FamilyName), StringComparison.Ordinal)
                && (string.IsNullOrWhiteSpace(operation.CategoryName) || string.IsNullOrWhiteSpace(change.CategoryName)
                    || string.Equals(Normalize(operation.CategoryName), Normalize(change.CategoryName), StringComparison.Ordinal));
        }
        if (string.Equals(change.EntryKind, "FamilyType", StringComparison.OrdinalIgnoreCase))
        {
            return string.Equals(operation.CandidateKind, "LoadableFamily", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(operation.TypeName)
                && string.Equals(Normalize(operation.FamilyName), Normalize(change.FamilyName), StringComparison.Ordinal)
                && string.Equals(Normalize(operation.TypeName), Normalize(change.TypeName), StringComparison.Ordinal)
                && (string.IsNullOrWhiteSpace(operation.CategoryName) || string.IsNullOrWhiteSpace(change.CategoryName)
                    || string.Equals(Normalize(operation.CategoryName), Normalize(change.CategoryName), StringComparison.Ordinal));
        }
        return string.Equals(change.EntryKind, "SystemType", StringComparison.OrdinalIgnoreCase)
            && string.Equals(operation.CandidateKind, "SystemType", StringComparison.OrdinalIgnoreCase)
            && string.Equals(Normalize(operation.TypeName), Normalize(change.TypeName), StringComparison.Ordinal)
            && (string.IsNullOrWhiteSpace(operation.SystemFamilyKind) || string.IsNullOrWhiteSpace(change.TypeClassName)
                || string.Equals(Normalize(operation.SystemFamilyKind), Normalize(change.TypeClassName), StringComparison.Ordinal));
    }

    private static List<FamilyBrowserOperationLogEntry> LoadCommittedOperations(string workspaceRoot, string projectComparableIdentity, string sinceUtc)
    {
        List<FamilyBrowserOperationLogEntry> result = new List<FamilyBrowserOperationLogEntry>();
        FamilyBrowserTrackingPersistenceService.FlushPending(workspaceRoot);
        string folder = FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "OperationLogs");
        if (!Directory.Exists(folder))
        {
            return result;
        }
        DateTime since;
        bool hasSince = DateTime.TryParse(sinceUtc ?? string.Empty, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out since);
        List<FamilyBrowserOperationLogEntry> candidates = FamilyBrowserTrackingPersistenceService.LoadImmutableOperationEntries(workspaceRoot, 10000);
        foreach (string path in Directory.GetFiles(folder, "family-browser-operations-*.json").OrderByDescending(delegate(string x) { return x; }, StringComparer.OrdinalIgnoreCase).Take(90))
        {
            FamilyBrowserOperationLog log = LoadJson<FamilyBrowserOperationLog>(path);
            candidates.AddRange(log == null ? new List<FamilyBrowserOperationLogEntry>() : log.Entries ?? new List<FamilyBrowserOperationLogEntry>());
        }
        foreach (FamilyBrowserOperationLogEntry entry in candidates
            .Where(delegate(FamilyBrowserOperationLogEntry value) { return value != null; })
            .GroupBy(delegate(FamilyBrowserOperationLogEntry value) { return string.IsNullOrWhiteSpace(value.EntryId) ? BuildLegacyOperationIdentity(value) : value.EntryId; }, StringComparer.OrdinalIgnoreCase)
            .Select(delegate(IGrouping<string, FamilyBrowserOperationLogEntry> group) { return group.OrderByDescending(delegate(FamilyBrowserOperationLogEntry value) { return value.CommittedAtUtc ?? value.RecordedAtUtc ?? string.Empty; }, StringComparer.Ordinal).First(); }))
        {
            if (!string.Equals(entry.CommitState, "Committed", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }
            DateTime committed;
            string committedText = string.IsNullOrWhiteSpace(entry.CommittedAtUtc) ? entry.RecordedAtUtc : entry.CommittedAtUtc;
            if (hasSince && DateTime.TryParse(committedText ?? string.Empty, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out committed) && committed < since)
            {
                continue;
            }
            string operationIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(entry.DocumentPath ?? string.Empty);
            if (string.Equals(operationIdentity, projectComparableIdentity, StringComparison.OrdinalIgnoreCase))
            {
                result.Add(entry);
            }
        }
        return result;
    }

    private static string BuildLegacyOperationIdentity(FamilyBrowserOperationLogEntry entry)
    {
        if (entry == null)
        {
            return string.Empty;
        }
        return (entry.CommittedAtUtc ?? entry.RecordedAtUtc ?? string.Empty) + "|" + (entry.CandidateKind ?? string.Empty) + "|" + (entry.CategoryName ?? string.Empty) + "|" + (entry.FamilyName ?? string.Empty) + "|" + (entry.TypeName ?? string.Empty) + "|" + (entry.SystemFamilyKind ?? string.Empty);
    }

    private static FamilyBrowserProjectCatalogSnapshot CreateSnapshotShell(Document doc, string trigger)
    {
        string identityPath = ProjectSnapshotStore.ResolveProjectIdentityPath(doc);
        return new FamilyBrowserProjectCatalogSnapshot
        {
            ProjectTitle = SafeDocumentTitle(doc),
            ProjectIdentityPath = identityPath ?? string.Empty,
            ProjectCanonicalPath = FamilyBrowserPathIdentityService.GetCanonicalPath(identityPath),
            ProjectComparableIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(identityPath),
            CapturedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
            Trigger = trigger ?? string.Empty,
            RevitVersion = SafeRevitVersion(doc),
            RevitUserName = SafeRevitUserName(doc),
            WindowsUserName = Environment.UserName,
            MachineName = Environment.MachineName
        };
    }

    private static FamilyBrowserProjectCatalogEntry CreateEntry(string entryKind, string categoryName, string categoryId, string familyName, string typeName, string typeClassName)
    {
        FamilyBrowserProjectCatalogEntry entry = new FamilyBrowserProjectCatalogEntry
        {
            EntryKind = entryKind ?? string.Empty,
            CategoryName = categoryName ?? string.Empty,
            CategoryId = categoryId ?? string.Empty,
            FamilyName = familyName ?? string.Empty,
            TypeName = typeName ?? string.Empty,
            TypeClassName = typeClassName ?? string.Empty
        };
        entry.Key = BuildEntryKey(entry);
        return entry;
    }

    private static string BuildEntryKey(FamilyBrowserProjectCatalogEntry entry)
    {
        return (entry.EntryKind ?? string.Empty).Trim() + "|" + (entry.TypeClassName ?? string.Empty).Trim() + "|" + (entry.CategoryId ?? string.Empty).Trim() + "|" + (entry.CategoryName ?? string.Empty).Trim() + "|" + (entry.FamilyName ?? string.Empty).Trim() + "|" + (entry.TypeName ?? string.Empty).Trim();
    }

    private static void AddEntry(IDictionary<string, FamilyBrowserProjectCatalogEntry> entries, FamilyBrowserProjectCatalogEntry entry)
    {
        if (entry != null && !string.IsNullOrWhiteSpace(entry.Key) && !entries.ContainsKey(entry.Key))
        {
            entries.Add(entry.Key, entry);
        }
    }

    private static void FinalizeSnapshot(FamilyBrowserProjectCatalogSnapshot snapshot, long elapsedMilliseconds)
    {
        snapshot.FamilyCount = snapshot.Entries.Count(delegate(FamilyBrowserProjectCatalogEntry x) { return string.Equals(x.EntryKind, "Family", StringComparison.OrdinalIgnoreCase); });
        snapshot.FamilyTypeCount = snapshot.Entries.Count(delegate(FamilyBrowserProjectCatalogEntry x) { return string.Equals(x.EntryKind, "FamilyType", StringComparison.OrdinalIgnoreCase); });
        snapshot.SystemTypeCount = snapshot.Entries.Count(delegate(FamilyBrowserProjectCatalogEntry x) { return string.Equals(x.EntryKind, "SystemType", StringComparison.OrdinalIgnoreCase); });
        snapshot.CatalogHash = ComputeCatalogHash(snapshot.Entries);
        snapshot.ElapsedMilliseconds = elapsedMilliseconds;
    }

    private static string ComputeCatalogHash(IEnumerable<FamilyBrowserProjectCatalogEntry> entries)
    {
        string payload = string.Join("\n", (entries ?? new List<FamilyBrowserProjectCatalogEntry>()).Where(delegate(FamilyBrowserProjectCatalogEntry x) { return x != null; }).Select(delegate(FamilyBrowserProjectCatalogEntry x) { return x.Key ?? string.Empty; }).OrderBy(delegate(string x) { return x; }, StringComparer.Ordinal));
        return HashText(payload);
    }

    private static string SaveSnapshot(string folder, FamilyBrowserProjectCatalogSnapshot snapshot)
    {
        string snapshotFolder = Path.Combine(folder, "Snapshots");
        Directory.CreateDirectory(snapshotFolder);
        string fileName = "project-catalog-" + snapshot.CatalogHash + ".json";
        string path = Path.Combine(snapshotFolder, fileName);
        if (!File.Exists(path))
        {
            SaveJson(path, snapshot);
        }
        else
        {
            LoadSnapshot(folder, Path.Combine("Snapshots", fileName), snapshot.CatalogHash, snapshot.ProjectIdentityPath, true);
        }
        return Path.Combine("Snapshots", fileName);
    }

    private static FamilyBrowserProjectCatalogSnapshot LoadSnapshot(
        string folder,
        string storedPath,
        string expectedCatalogHash,
        string expectedProjectIdentityPath,
        bool required)
    {
        if (string.IsNullOrWhiteSpace(storedPath))
        {
            if (required)
            {
                throw new InvalidDataException("The project catalog snapshot reference is missing.");
            }
            return null;
        }
        string path = ResolveSnapshotPath(folder, storedPath);
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("The referenced project catalog snapshot is missing.", path);
        }
        FamilyBrowserProjectCatalogSnapshot snapshot = LoadJson<FamilyBrowserProjectCatalogSnapshot>(path);
        if (snapshot == null)
        {
            throw new InvalidDataException("The referenced project catalog snapshot is unreadable: " + path);
        }
        ValidateSnapshot(snapshot, expectedCatalogHash, expectedProjectIdentityPath);
        return snapshot;
    }

    private static string ResolveSnapshotPath(string folder, string storedPath)
    {
        string fileName = Path.GetFileName((storedPath ?? string.Empty).Trim());
        if (string.IsNullOrWhiteSpace(fileName) || fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
        {
            throw new InvalidDataException("The project catalog snapshot reference is invalid.");
        }
        return Path.Combine(folder, "Snapshots", fileName);
    }

    private static string NormalizeSnapshotReference(string storedPath)
    {
        string fileName = Path.GetFileName((storedPath ?? string.Empty).Trim());
        return string.IsNullOrWhiteSpace(fileName) ? string.Empty : fileName;
    }

    private static void ValidateSnapshot(FamilyBrowserProjectCatalogSnapshot snapshot, string expectedCatalogHash, string expectedProjectIdentityPath)
    {
        if (snapshot == null || snapshot.SchemaVersion != SchemaVersion || snapshot.Entries == null)
        {
            throw new InvalidDataException("The project catalog snapshot schema is invalid.");
        }
        string expectedStableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(expectedProjectIdentityPath);
        string snapshotStableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(snapshot.ProjectIdentityPath);
        if (!string.IsNullOrWhiteSpace(expectedStableIdentity) &&
            !string.Equals(expectedStableIdentity, snapshotStableIdentity, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidDataException("The project catalog snapshot belongs to a different project path.");
        }
        HashSet<string> keys = new HashSet<string>(StringComparer.Ordinal);
        foreach (FamilyBrowserProjectCatalogEntry entry in snapshot.Entries)
        {
            if (entry == null || string.IsNullOrWhiteSpace(entry.Key) ||
                !string.Equals(entry.Key, BuildEntryKey(entry), StringComparison.Ordinal) ||
                !keys.Add(entry.Key))
            {
                throw new InvalidDataException("The project catalog snapshot contains an invalid or duplicate entry key.");
            }
        }
        string computedHash = ComputeCatalogHash(snapshot.Entries);
        if (!string.Equals(snapshot.CatalogHash ?? string.Empty, computedHash, StringComparison.Ordinal) ||
            (!string.IsNullOrWhiteSpace(expectedCatalogHash) && !string.Equals(expectedCatalogHash, computedHash, StringComparison.Ordinal)) ||
            snapshot.FamilyCount != snapshot.Entries.Count(delegate(FamilyBrowserProjectCatalogEntry x) { return string.Equals(x.EntryKind, "Family", StringComparison.OrdinalIgnoreCase); }) ||
            snapshot.FamilyTypeCount != snapshot.Entries.Count(delegate(FamilyBrowserProjectCatalogEntry x) { return string.Equals(x.EntryKind, "FamilyType", StringComparison.OrdinalIgnoreCase); }) ||
            snapshot.SystemTypeCount != snapshot.Entries.Count(delegate(FamilyBrowserProjectCatalogEntry x) { return string.Equals(x.EntryKind, "SystemType", StringComparison.OrdinalIgnoreCase); }))
        {
            throw new InvalidDataException("The project catalog snapshot hash or summary counters are inconsistent.");
        }
    }

    private static void ValidateManifest(FamilyBrowserProjectCatalogManifest manifest, string expectedProjectIdentityPath)
    {
        if (manifest == null)
        {
            return;
        }
        if (manifest.SchemaVersion != 0 && manifest.SchemaVersion != SchemaVersion)
        {
            throw new InvalidDataException("The project catalog manifest schema is unsupported.");
        }
        string expectedStableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(expectedProjectIdentityPath);
        string manifestStableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(manifest.ProjectIdentityPath);
        if (!string.IsNullOrWhiteSpace(manifest.ProjectIdentityPath) &&
            !string.IsNullOrWhiteSpace(expectedStableIdentity) &&
            !string.Equals(expectedStableIdentity, manifestStableIdentity, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidDataException("The project catalog manifest belongs to a different project path.");
        }
        bool acceptedReferencePresent = !string.IsNullOrWhiteSpace(manifest.AcceptedSnapshotPath);
        bool acceptedHashPresent = !string.IsNullOrWhiteSpace(manifest.AcceptedCatalogHash);
        bool observedReferencePresent = !string.IsNullOrWhiteSpace(manifest.LastObservedSnapshotPath);
        bool observedHashPresent = !string.IsNullOrWhiteSpace(manifest.LastObservedCatalogHash);
        if (acceptedReferencePresent != acceptedHashPresent || observedReferencePresent != observedHashPresent)
        {
            throw new InvalidDataException("The project catalog manifest contains an incomplete snapshot binding.");
        }
    }

    private static string ResolveProjectFolder(string workspaceRoot, string identityPath)
    {
        string preferred = GetProjectFolder(workspaceRoot, identityPath);
        if (Directory.Exists(preferred))
        {
            return preferred;
        }
        string root = FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "ProjectCatalogs");
        string legacyIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(identityPath);
        if (!string.IsNullOrWhiteSpace(legacyIdentity))
        {
            string legacy = Path.Combine(root, HashText(legacyIdentity).Substring(0, 32));
            if (Directory.Exists(legacy))
            {
                return legacy;
            }
        }
        if (!Directory.Exists(root))
        {
            return preferred;
        }
        string expectedStableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(identityPath);
        List<string> matchingLegacyFolders = new List<string>();
        foreach (string candidate in Directory.EnumerateDirectories(root, "*", SearchOption.TopDirectoryOnly))
        {
            FamilyBrowserProjectCatalogManifest manifest = LoadJson<FamilyBrowserProjectCatalogManifest>(Path.Combine(candidate, "manifest.json"));
            if (manifest == null)
            {
                continue;
            }
            string candidateStableIdentity = FamilyBrowserPathIdentityService.GetStablePathIdentity(manifest.ProjectIdentityPath);
            if (!string.IsNullOrWhiteSpace(expectedStableIdentity) &&
                string.Equals(expectedStableIdentity, candidateStableIdentity, StringComparison.OrdinalIgnoreCase))
            {
                matchingLegacyFolders.Add(candidate);
            }
        }
        if (matchingLegacyFolders.Count > 1)
        {
            throw new InvalidDataException("Multiple legacy project catalog folders match the same stable project path. Preserve them and complete an administrator-reviewed merge before accepting a new baseline.");
        }
        return matchingLegacyFolders.Count == 1 ? matchingLegacyFolders[0] : preferred;
    }

    private static void SaveJson<T>(string path, T value) where T : class
    {
        if (value == null || string.IsNullOrWhiteSpace(path))
        {
            return;
        }
        Directory.CreateDirectory(Path.GetDirectoryName(path));
        string temporary = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
        try
        {
            using (FileStream stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
                serializer.WriteObject(stream, value);
                stream.Flush(true);
            }
            FamilyBrowserAtomicFileService.Promote(temporary, path);
        }
        finally
        {
            if (File.Exists(temporary))
            {
                try { File.Delete(temporary); } catch { }
            }
        }
    }

    private static T LoadJson<T>(string path) where T : class
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return null;
        }
        try
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
                return serializer.ReadObject(stream) as T;
            }
        }
        catch
        {
            return null;
        }
    }

    private static FileStream AcquireCatalogLock(string folder)
    {
        string path = Path.Combine(folder, ".project-catalog.lock");
        IOException last = null;
        for (int attempt = 0; attempt < 20; attempt++)
        {
            try
            {
                return new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException ex)
            {
                last = ex;
                Thread.Sleep(50);
            }
        }
        throw new IOException("The project catalog is currently being updated by another process.", last);
    }

    private static FamilyBrowserProjectCatalogChange CreateChange(string changeKind, FamilyBrowserProjectCatalogEntry entry)
    {
        return new FamilyBrowserProjectCatalogChange
        {
            ChangeKind = changeKind ?? string.Empty,
            EntryKind = entry == null ? string.Empty : entry.EntryKind ?? string.Empty,
            Key = entry == null ? string.Empty : entry.Key ?? string.Empty,
            CategoryName = entry == null ? string.Empty : entry.CategoryName ?? string.Empty,
            FamilyName = entry == null ? string.Empty : entry.FamilyName ?? string.Empty,
            TypeName = entry == null ? string.Empty : entry.TypeName ?? string.Empty,
            TypeClassName = entry == null ? string.Empty : entry.TypeClassName ?? string.Empty,
            Attribution = "ExternalUntracked"
        };
    }

    private static FamilyBrowserProjectCatalogState CreateUnavailableState(string stateCode, string reason)
    {
        return new FamilyBrowserProjectCatalogState
        {
            StateCode = stateCode ?? "NotChecked",
            CheckedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
            Reason = reason ?? string.Empty
        };
    }

    private static FamilyBrowserProjectCatalogState CreatePublicationDeferredState(Document doc, FamilyBrowserProjectCatalogSnapshot snapshot, string trigger, string reason)
    {
        FamilyBrowserProjectCatalogState state = CreateUnavailableState("PublicationDeferred", reason);
        if (snapshot != null)
        {
            CopySnapshotSummary(snapshot, state);
        }
        else
        {
            state.ProjectTitle = SafeDocumentTitle(doc);
            state.ProjectIdentityPath = doc == null ? string.Empty : ProjectSnapshotStore.ResolveProjectIdentityPath(doc);
            state.ProjectComparableIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(state.ProjectIdentityPath);
            state.Trigger = trigger ?? string.Empty;
        }
        return state;
    }

    private static FamilyBrowserProjectCatalogState CreateErrorState(Document doc, string trigger, Exception ex)
    {
        FamilyBrowserProjectCatalogState state = CreateUnavailableState("Error", "The project catalog check failed.");
        state.ProjectTitle = SafeDocumentTitle(doc);
        state.ProjectIdentityPath = doc == null ? string.Empty : ProjectSnapshotStore.ResolveProjectIdentityPath(doc);
        state.ProjectComparableIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(state.ProjectIdentityPath);
        state.Trigger = trigger ?? string.Empty;
        state.ErrorMessage = ex == null ? string.Empty : ex.Message;
        return state;
    }

    private static void CopySnapshotSummary(FamilyBrowserProjectCatalogSnapshot snapshot, FamilyBrowserProjectCatalogState state)
    {
        if (snapshot == null || state == null)
        {
            return;
        }
        state.ProjectTitle = snapshot.ProjectTitle;
        state.ProjectIdentityPath = snapshot.ProjectIdentityPath;
        state.ProjectComparableIdentity = snapshot.ProjectComparableIdentity;
        state.CheckedAtUtc = snapshot.CapturedAtUtc;
        state.CurrentCatalogHash = snapshot.CatalogHash;
        state.Trigger = snapshot.Trigger;
        state.ElapsedMilliseconds = snapshot.ElapsedMilliseconds;
        state.FamilyCount = snapshot.FamilyCount;
        state.FamilyTypeCount = snapshot.FamilyTypeCount;
        state.SystemTypeCount = snapshot.SystemTypeCount;
    }

    private static bool IsInPlaceFamily(Family family)
    {
        try { return family != null && family.IsInPlace; } catch { return false; }
    }

    private static string SafeElementName(Element element)
    {
        try { return element == null ? string.Empty : element.Name ?? string.Empty; } catch { return string.Empty; }
    }

    private static string SafeDocumentTitle(Document doc)
    {
        try { return doc == null ? string.Empty : doc.Title ?? string.Empty; } catch { return string.Empty; }
    }

    private static string SafeRevitVersion(Document doc)
    {
        try { return doc == null || doc.Application == null ? string.Empty : doc.Application.VersionNumber ?? string.Empty; } catch { return string.Empty; }
    }

    private static string SafeRevitUserName(Document doc)
    {
        try { return doc == null || doc.Application == null ? string.Empty : doc.Application.Username ?? string.Empty; } catch { return string.Empty; }
    }

    private static string Normalize(string value)
    {
        return (value ?? string.Empty).Trim().ToUpperInvariant();
    }

    private static string HashText(string value)
    {
        using (SHA256 sha = SHA256.Create())
        {
            byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty));
            StringBuilder builder = new StringBuilder(hash.Length * 2);
            foreach (byte item in hash)
            {
                builder.Append(item.ToString("x2", CultureInfo.InvariantCulture));
            }
            return builder.ToString();
        }
    }
}
