using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Microsoft.VisualBasic.CompilerServices;

public sealed class StandardRvtChangeCandidateService
{
	private sealed class StandardRvtTarget
	{
		public string SourceId { get; set; }

		public string StandardRvtPath { get; set; }

		public string SnapshotPath { get; set; }

		public string SlotKey { get; set; }

		public string DisciplineKey { get; set; }

		public string DisciplineLabel { get; set; }

		public StandardRvtTarget()
		{
			SourceId = string.Empty;
			StandardRvtPath = string.Empty;
			SnapshotPath = string.Empty;
			SlotKey = string.Empty;
			DisciplineKey = string.Empty;
			DisciplineLabel = string.Empty;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__25_002D0
	{
		public StandardRvtChangeCandidateEntry _0024VB_0024Local_entry;

		public _Closure_0024__25_002D0(_Closure_0024__25_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_entry = arg0._0024VB_0024Local_entry;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__1(StandardRvtChangeCandidateEntry x)
		{
			return IsSameCandidate(x, _0024VB_0024Local_entry);
		}
	}

	private const int CacheRefreshSeconds = 30;

	private const int MaxLogEntriesPerSource = 5000;

	private static readonly object SyncRoot = RuntimeHelpers.GetObjectValue(new object());

	private static string CachedWorkspaceRoot = string.Empty;

	private static DateTime CachedAtUtc = DateTime.MinValue;

	private static List<StandardRvtTarget> CachedTargets = new List<StandardRvtTarget>();

	private static readonly Dictionary<string, Dictionary<string, StandardRvtChangeCandidateEntry>> PendingCandidatesByDocument = new Dictionary<string, Dictionary<string, StandardRvtChangeCandidateEntry>>(StringComparer.OrdinalIgnoreCase);

	private static readonly Dictionary<string, PendingCandidateSaveBatch> PendingCandidateSaveBatchesByDocument = new Dictionary<string, PendingCandidateSaveBatch>(StringComparer.OrdinalIgnoreCase);

	private static readonly Dictionary<int, string> PendingCandidateCloseKeysByDocumentId = new Dictionary<int, string>();

	private static readonly Dictionary<string, PendingOperationBatch> PendingOperationEntriesByDocument = new Dictionary<string, PendingOperationBatch>(StringComparer.OrdinalIgnoreCase);

	private static readonly Dictionary<int, string> PendingOperationCloseKeysByDocumentId = new Dictionary<int, string>();

	private static readonly HashSet<string> AllowedSystemTypeNames = new HashSet<string>(new string[23]
	{
		"WallType", "FloorType", "RoofType", "CeilingType", "StairsType", "RailingType", "CurtainSystemType", "PanelType", "DuctType", "PipeType", "FlexDuctType", "FlexPipeType",
		"DuctSystemType", "PipingSystemType", "MechanicalSystemType", "ElectricalSystemType", "CableTrayType", "ConduitType", "WireType", "DuctInsulationType", "PipeInsulationType", "DuctLiningType",
		"MullionType"
	}, StringComparer.OrdinalIgnoreCase);

	private StandardRvtChangeCandidateService()
	{
	}

	public static void NotifyPolicyChanged()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			CachedAtUtc = DateTime.MinValue;
			CachedTargets = new List<StandardRvtTarget>();
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static void HandleDocumentChanged(DocumentChangedEventArgs e)
	{
		if (e == null)
		{
			return;
		}
		Document doc = null;
		try
		{
			doc = e.GetDocument();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
			return;
		}
		if (doc == null || string.IsNullOrWhiteSpace(doc.PathName))
		{
			return;
		}
		string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
		StandardRvtTarget target = ResolveTargetForDocument(workspaceRoot, doc);
		if (target != null)
		{
			Dictionary<string, StandardRvtChangeCandidateEntry> entries = new Dictionary<string, StandardRvtChangeCandidateEntry>(StringComparer.Ordinal);
			CollectCandidates(entries, doc, e.GetAddedElementIds(), target, "Added");
			CollectCandidates(entries, doc, e.GetModifiedElementIds(), target, "Modified");
			AddPendingCandidates(doc, target, entries.Values);
		}
	}

	private sealed class PendingOperationBatch
	{
		public string WorkspaceRoot { get; set; }

		public string DocumentKey { get; set; }

		public List<FamilyBrowserOperationLogEntry> Entries { get; set; }

		public PendingOperationBatch()
		{
			WorkspaceRoot = string.Empty;
			DocumentKey = string.Empty;
			Entries = new List<FamilyBrowserOperationLogEntry>();
		}
	}

	private sealed class PendingCandidateSaveBatch
	{
		public string WorkspaceRoot { get; set; }

		public string DocumentKey { get; set; }

		public StandardRvtTarget Target { get; set; }

		public List<StandardRvtChangeCandidateEntry> Entries { get; set; }

		public PendingCandidateSaveBatch()
		{
			WorkspaceRoot = string.Empty;
			DocumentKey = string.Empty;
			Entries = new List<StandardRvtChangeCandidateEntry>();
		}
	}

	public static void HandleDocumentSaving(Document doc)
	{
		if (doc == null || string.IsNullOrWhiteSpace(doc.PathName))
		{
			return;
		}
		string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
		StandardRvtTarget target = ResolveTargetForDocument(workspaceRoot, doc);
		if (target == null)
		{
			return;
		}
		string documentKey = BuildPendingOperationKey(workspaceRoot, doc);
		Dictionary<string, StandardRvtChangeCandidateEntry> entries = new Dictionary<string, StandardRvtChangeCandidateEntry>(StringComparer.Ordinal);
		PendingCandidateSaveBatch existingBatch = null;
		object existingSync = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(existingSync);
		bool existingLockTaken = false;
		try
		{
			Monitor.Enter(existingSync, ref existingLockTaken);
			PendingCandidateSaveBatchesByDocument.TryGetValue(documentKey, out existingBatch);
		}
		finally
		{
			if (existingLockTaken)
			{
				Monitor.Exit(existingSync);
			}
		}
		foreach (StandardRvtChangeCandidateEntry entry in existingBatch == null ? Enumerable.Empty<StandardRvtChangeCandidateEntry>() : existingBatch.Entries)
		{
			AddOrMerge(entries, entry);
		}
		foreach (StandardRvtChangeCandidateEntry entry in DrainPendingCandidates(doc, target))
		{
			AddOrMerge(entries, entry);
		}
		CollectSnapshotDeltaCandidates(entries, doc, target);
		if (entries.Count != 0)
		{
			PendingCandidateSaveBatch batch = new PendingCandidateSaveBatch
			{
				WorkspaceRoot = workspaceRoot,
				DocumentKey = documentKey,
				Target = target,
				Entries = entries.Values.ToList()
			};
			object syncRoot = SyncRoot;
			ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
			bool lockTaken = false;
			try
			{
				Monitor.Enter(syncRoot, ref lockTaken);
				PendingCandidateSaveBatchesByDocument[documentKey] = batch;
			}
			finally
			{
				if (lockTaken)
				{
					Monitor.Exit(syncRoot);
				}
			}
		}
	}

	public static bool HandleDocumentSaved(Document doc, object status, string commitKind)
	{
		if (!IsSuccessfulRevitApiEventStatus(status))
		{
			return false;
		}
		string finalKind = string.IsNullOrWhiteSpace(commitKind) ? "Save" : commitKind;
		bool candidateCommitted = CommitPendingCandidateEntries(doc, finalKind);
		bool operationCommitted = CommitPendingOperationEntries(doc, finalKind);
		return candidateCommitted || operationCommitted;
	}

	public static bool HandleDocumentSynchronizedWithCentral(Document doc, object status)
	{
		if (!IsSuccessfulRevitApiEventStatus(status))
		{
			return false;
		}
		HandleDocumentSaving(doc);
		bool candidateCommitted = CommitPendingCandidateEntries(doc, "SynchronizeWithCentral");
		bool operationCommitted = CommitPendingOperationEntries(doc, "SynchronizeWithCentral");
		return candidateCommitted || operationCommitted;
	}

	private static bool CommitPendingCandidateEntries(Document doc, string commitKind)
	{
		if (doc == null)
		{
			return false;
		}
		string documentKey = BuildPendingOperationKey(HostWorkspacePathResolver.ResolveRoot(), doc);
		PendingCandidateSaveBatch batch = null;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (!PendingCandidateSaveBatchesByDocument.TryGetValue(documentKey, out batch) || batch == null)
			{
				return false;
			}
			PendingCandidateSaveBatchesByDocument.Remove(documentKey);
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		List<StandardRvtChangeCandidateEntry> entries = (batch.Entries ?? new List<StandardRvtChangeCandidateEntry>()).Where([SpecialName] (StandardRvtChangeCandidateEntry x) => x != null).ToList();
		if (entries.Count == 0 || batch.Target == null)
		{
			return false;
		}
		string committedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		string documentPath = doc.PathName ?? string.Empty;
		string canonicalIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(documentPath);
		string targetIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(batch.Target.StandardRvtPath);
		if (string.Equals(commitKind, "SaveAs", StringComparison.OrdinalIgnoreCase) && !string.Equals(canonicalIdentity, targetIdentity, StringComparison.OrdinalIgnoreCase))
		{
			return false;
		}
		foreach (StandardRvtChangeCandidateEntry entry in entries)
		{
			entry.CommitState = "Committed";
			entry.CommitKind = commitKind ?? string.Empty;
			entry.CommittedAtUtc = committedAtUtc;
			entry.DocumentPath = documentPath;
			entry.CanonicalDocumentIdentity = canonicalIdentity;
			if (string.IsNullOrWhiteSpace(entry.BeforeFingerprint))
			{
				entry.BeforeFingerprint = string.Equals(entry.ChangeKind, "Added", StringComparison.OrdinalIgnoreCase) ? "missing" : "last-scan fingerprint unavailable";
			}
			if (string.IsNullOrWhiteSpace(entry.AfterFingerprint))
			{
				entry.AfterFingerprint = string.Equals(entry.ChangeKind, "Deleted", StringComparison.OrdinalIgnoreCase) ? "missing" : "changed; precise rescan required";
			}
		}
		if (AppendCandidates(batch.WorkspaceRoot, batch.Target, entries))
		{
			return true;
		}
		RestorePendingCandidateBatch(documentKey, batch);
		return false;
	}

	private static void RestorePendingCandidateBatch(string documentKey, PendingCandidateSaveBatch batch)
	{
		if (string.IsNullOrWhiteSpace(documentKey) || batch == null)
		{
			return;
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			PendingCandidateSaveBatch current = null;
			if (!PendingCandidateSaveBatchesByDocument.TryGetValue(documentKey, out current) || current == null)
			{
				PendingCandidateSaveBatchesByDocument[documentKey] = batch;
				return;
			}
			current.Entries.AddRange(batch.Entries ?? new List<StandardRvtChangeCandidateEntry>());
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static List<FamilyBrowserOperationLogEntry> GetPendingOperationEntries(string workspaceRoot, Document doc)
	{
		List<FamilyBrowserOperationLogEntry> snapshot = new List<FamilyBrowserOperationLogEntry>();
		string key = BuildPendingOperationKey(workspaceRoot, doc);
		if (string.IsNullOrWhiteSpace(key))
		{
			return snapshot;
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			PendingOperationBatch batch = null;
			if (PendingOperationEntriesByDocument.TryGetValue(key, out batch) && batch != null)
			{
				snapshot = (batch.Entries ?? new List<FamilyBrowserOperationLogEntry>()).Where([SpecialName] (FamilyBrowserOperationLogEntry x) => ShouldVerifyCommittedOperation(x)).Select([SpecialName] (FamilyBrowserOperationLogEntry x) => CloneOperationEntry(x)).ToList();
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		return snapshot.Where([SpecialName] (FamilyBrowserOperationLogEntry x) => IsPendingOperationPresentAtSync(doc, x)).ToList();
	}

	public static bool HasPendingOperationEntries(string workspaceRoot, Document doc)
	{
		return GetPendingOperationEntries(workspaceRoot, doc).Count > 0;
	}

	public static void HandleDocumentClosing(Document doc, int documentId)
	{
		string key = BuildPendingOperationKey(HostWorkspacePathResolver.ResolveRoot(), doc);
		if (documentId < 0 || string.IsNullOrWhiteSpace(key))
		{
			return;
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			PendingOperationCloseKeysByDocumentId.Remove(documentId);
			PendingCandidateCloseKeysByDocumentId.Remove(documentId);
			if (PendingOperationEntriesByDocument.ContainsKey(key))
			{
				PendingOperationCloseKeysByDocumentId[documentId] = key;
			}
			if (PendingCandidateSaveBatchesByDocument.ContainsKey(key))
			{
				PendingCandidateCloseKeysByDocumentId[documentId] = key;
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static void HandleDocumentClosed(int documentId)
	{
		if (documentId < 0)
		{
			return;
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			string key = string.Empty;
			if (PendingOperationCloseKeysByDocumentId.TryGetValue(documentId, out key))
			{
				PendingOperationCloseKeysByDocumentId.Remove(documentId);
				PendingOperationEntriesByDocument.Remove(key);
			}
			string candidateKey = string.Empty;
			if (PendingCandidateCloseKeysByDocumentId.TryGetValue(documentId, out candidateKey))
			{
				PendingCandidateCloseKeysByDocumentId.Remove(documentId);
				PendingCandidateSaveBatchesByDocument.Remove(candidateKey);
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static void ClearPendingOperationState()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			PendingOperationEntriesByDocument.Clear();
			PendingOperationCloseKeysByDocumentId.Clear();
			PendingCandidateSaveBatchesByDocument.Clear();
			PendingCandidateCloseKeysByDocumentId.Clear();
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static void RecordLoadableFamilyOperation(string workspaceRoot, Document doc, StandardLibraryRegistrationRecord registration, LoadableFamilySyncExecutionReport execution, string operationKind)
	{
		if (execution != null && execution.Items != null)
		{
			List<FamilyBrowserOperationLogEntry> entries = BuildLoadableFamilyOperationEntries(doc, registration, execution, operationKind);
			QueueOperationEntries(workspaceRoot, doc, entries);
		}
	}

	private static List<FamilyBrowserOperationLogEntry> BuildLoadableFamilyOperationEntries(Document doc, StandardLibraryRegistrationRecord registration, LoadableFamilySyncExecutionReport execution, string operationKind)
	{
		List<FamilyBrowserOperationLogEntry> entries = new List<FamilyBrowserOperationLogEntry>();
		foreach (LoadableFamilySyncExecutionItem item in execution?.Items ?? new List<LoadableFamilySyncExecutionItem>())
		{
			if (item == null)
			{
				continue;
			}
			FamilyBrowserOperationLogEntry familyEntry = new FamilyBrowserOperationLogEntry
			{
				EntryId = Guid.NewGuid().ToString("N"),
				RecordedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
				UserName = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity(),
				OperationKind = (operationKind ?? "LoadableFamilyOperation"),
				DocumentTitle = (doc?.Title ?? execution.ProjectDocumentTitle),
				DocumentPath = (doc?.PathName ?? string.Empty),
				SourceId = (registration?.SourceId ?? string.Empty),
				StandardSourceKey = ((registration == null) ? string.Empty : FamilyBrowserDataLoader.BuildSourceKey(registration.SourceId, registration.LastSnapshotPath)),
				StandardDisplayName = (registration?.DisplayName ?? execution.StandardDisplayName),
				CandidateKind = "LoadableFamily",
				CategoryName = (item.CategoryName ?? string.Empty),
				FamilyName = (item.FamilyName ?? string.Empty),
				PlannedAction = (item.PlannedAction ?? string.Empty),
				Outcome = (item.Outcome ?? string.Empty),
				Details = (item.Details ?? string.Empty),
				CommitState = "PendingSaveOrSync",
				CommitKind = string.Empty,
				CommittedAtUtc = string.Empty
			};
			entries.Add(familyEntry);
			if (!ShouldVerifyCommittedOperation(familyEntry))
			{
				continue;
			}
			foreach (string typeName in ResolveLoadedFamilyTypeNames(doc, item.CategoryName, item.FamilyName))
			{
				FamilyBrowserOperationLogEntry typeEntry = CloneOperationEntry(familyEntry);
				typeEntry.EntryId = Guid.NewGuid().ToString("N");
				typeEntry.TypeName = typeName;
				entries.Add(typeEntry);
			}
		}
		return entries;
	}

	private static List<string> ResolveLoadedFamilyTypeNames(Document doc, string categoryName, string familyName)
	{
		List<string> result = new List<string>();
		if (doc == null || string.IsNullOrWhiteSpace(familyName))
		{
			return result;
		}
		try
		{
			foreach (Family family in new FilteredElementCollector(doc).OfClass(typeof(Family)).Cast<Family>())
			{
				if (family == null || !string.Equals(Normalize(family.Name), Normalize(familyName), StringComparison.Ordinal))
				{
					continue;
				}
				string actualCategory = family.FamilyCategory?.Name ?? string.Empty;
				if (!string.IsNullOrWhiteSpace(categoryName) && !string.Equals(Normalize(actualCategory), Normalize(categoryName), StringComparison.Ordinal))
				{
					continue;
				}
				foreach (ElementId symbolId in family.GetFamilySymbolIds())
				{
					FamilySymbol symbol = doc.GetElement(symbolId) as FamilySymbol;
					string typeName = ResolveElementName(symbol);
					if (!string.IsNullOrWhiteSpace(typeName))
					{
						result.Add(typeName);
					}
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return result.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
	}

	public static void RecordSystemTypeOperation(string workspaceRoot, Document doc, StandardLibraryRegistrationRecord registration, SystemTypeApplyExecutionReport execution, string operationKind)
	{
		if (execution != null && execution.Items != null)
		{
			List<FamilyBrowserOperationLogEntry> entries = (from x in execution.Items
				where x != null
				select new FamilyBrowserOperationLogEntry
				{
					EntryId = Guid.NewGuid().ToString("N"),
					RecordedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
					UserName = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity(),
					OperationKind = (operationKind ?? "SystemTypeOperation"),
					DocumentTitle = (doc?.Title ?? execution.ProjectDocumentTitle),
					DocumentPath = (doc?.PathName ?? string.Empty),
					SourceId = (registration?.SourceId ?? string.Empty),
					StandardSourceKey = ((registration == null) ? string.Empty : FamilyBrowserDataLoader.BuildSourceKey(registration.SourceId, registration.LastSnapshotPath)),
					StandardDisplayName = (registration?.DisplayName ?? execution.StandardDisplayName),
					CandidateKind = "SystemType",
					CategoryName = (x.CategoryName ?? string.Empty),
					TypeName = (x.SystemTypeName ?? string.Empty),
					SystemFamilyKind = (x.SystemFamilyKind ?? string.Empty),
					PlannedAction = (x.SyncAction ?? string.Empty),
					Outcome = (x.Outcome ?? string.Empty),
					Details = (x.Details ?? string.Empty),
					CommitState = "PendingSaveOrSync",
					CommitKind = string.Empty,
					CommittedAtUtc = string.Empty
				}).ToList();
			QueueOperationEntries(workspaceRoot, doc, entries);
		}
	}

	public static List<StandardRvtChangeCandidateEntry> LoadRecent(string workspaceRoot, string sourceId, int limit = 50)
	{
		if (string.IsNullOrWhiteSpace(sourceId))
		{
			return new List<StandardRvtChangeCandidateEntry>();
		}
		FamilyBrowserTrackingPersistenceService.FlushPending(workspaceRoot);
		List<StandardRvtChangeCandidateEntry> entries = (LoadLog(workspaceRoot, sourceId)?.Entries ?? new List<StandardRvtChangeCandidateEntry>()).Where([SpecialName] (StandardRvtChangeCandidateEntry x) => x != null).ToList();
		entries.AddRange(LoadImmutableHistory(workspaceRoot, sourceId, Math.Max(200, checked(limit * 5))));
		return entries.GroupBy<StandardRvtChangeCandidateEntry, string>([SpecialName] (StandardRvtChangeCandidateEntry x) => string.IsNullOrWhiteSpace(x.EntryId) ? BuildCandidateIdentity(x) : x.EntryId, StringComparer.OrdinalIgnoreCase).Select([SpecialName] (IGrouping<string, StandardRvtChangeCandidateEntry> x) => x.OrderByDescending([SpecialName] (StandardRvtChangeCandidateEntry y) => y.CommittedAtUtc ?? y.RecordedAtUtc ?? string.Empty, StringComparer.Ordinal).First()).OrderByDescending<StandardRvtChangeCandidateEntry, string>([SpecialName] (StandardRvtChangeCandidateEntry x) => string.IsNullOrWhiteSpace(x.CommittedAtUtc) ? x.RecordedAtUtc ?? string.Empty : x.CommittedAtUtc, StringComparer.Ordinal).Take(Math.Max(1, limit)).ToList();
	}

	private static List<StandardRvtChangeCandidateEntry> LoadImmutableHistory(string workspaceRoot, string sourceId, int limit)
	{
		List<StandardRvtChangeCandidateEntry> entries = new List<StandardRvtChangeCandidateEntry>();
		try
		{
			string folder = Path.Combine(GetCandidateFolder(workspaceRoot), "History", SafeFileName(sourceId));
			if (!Directory.Exists(folder))
			{
				return entries;
			}
			foreach (string path in Directory.EnumerateFiles(folder, "*.json", SearchOption.AllDirectories).Select([SpecialName] (string x) => new FileInfo(x)).OrderByDescending([SpecialName] (FileInfo x) => x.LastWriteTimeUtc).Take(Math.Max(1, limit)).Select([SpecialName] (FileInfo x) => x.FullName))
			{
				try
				{
					StandardRvtChangeCandidateEntry entry = DataContractJsonFileStore.Load<StandardRvtChangeCandidateEntry>(path);
					if (entry != null)
					{
						entries.Add(entry);
					}
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		return entries;
	}

	private static string BuildCandidateIdentity(StandardRvtChangeCandidateEntry entry)
	{
		if (entry == null)
		{
			return string.Empty;
		}
		return (entry.RecordedAtUtc ?? string.Empty) + "|" + Normalize(entry.CandidateKind) + "|" + Normalize(entry.CategoryName) + "|" + Normalize(entry.FamilyName) + "|" + Normalize(entry.SystemFamilyKind) + "|" + Normalize(entry.TypeName) + "|" + Normalize(entry.ChangeKind);
	}

	public static string BuildRecentLoadableFamilyNameText(string workspaceRoot, string sourceId, int limit = 40)
	{
		List<string> names = (from x in LoadRecent(workspaceRoot, sourceId, Math.Max(checked(limit * 2), 50))
			where string.Equals(x.CandidateKind ?? string.Empty, "LoadableFamily", StringComparison.OrdinalIgnoreCase)
			select (x.FamilyName ?? string.Empty).Trim() into x
			where x.Length > 0
			select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).Take(Math.Max(1, limit)).ToList();
		return string.Join(Environment.NewLine, names);
	}

	public static string BuildRecentSummary(string workspaceRoot, string sourceId, int limit = 12, bool isKorean = false)
	{
		List<StandardRvtChangeCandidateEntry> recent = LoadRecent(workspaceRoot, sourceId, 100);
		if (recent.Count == 0)
		{
			return isKorean ? "아직 표준 RVT 변경 후보 기록이 없습니다." : "No standard RVT change candidates have been recorded yet.";
		}
		List<string> loadableNames = (from x in recent
			where string.Equals(x.CandidateKind ?? string.Empty, "LoadableFamily", StringComparison.OrdinalIgnoreCase)
			select (x.FamilyName ?? string.Empty).Trim() into x
			where x.Length > 0
			select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).Take(limit).ToList();
		List<string> systemNames = (from x in recent
			where string.Equals(x.CandidateKind ?? string.Empty, "SystemType", StringComparison.OrdinalIgnoreCase)
			select (x.SystemFamilyKind ?? string.Empty).Trim() + " / " + (x.TypeName ?? string.Empty).Trim() into x
			where x.Trim().Length > 3
			select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).Take(limit).ToList();
		StringBuilder builder = new StringBuilder();
		builder.Append(isKorean ? "최근 후보: " : "Recent candidates: ");
		builder.Append(recent.Count.ToString(CultureInfo.InvariantCulture));
		if (loadableNames.Count > 0)
		{
			builder.Append(isKorean ? " | 패밀리: " : " | Families: ");
			builder.Append(string.Join(", ", loadableNames));
		}
		if (systemNames.Count > 0)
		{
			builder.Append(isKorean ? " | 시스템 타입: " : " | System types: ");
			builder.Append(string.Join(", ", systemNames));
		}
		return builder.ToString();
	}

	private static void AddPendingCandidates(Document doc, StandardRvtTarget target, IEnumerable<StandardRvtChangeCandidateEntry> candidates)
	{
		List<StandardRvtChangeCandidateEntry> list = (candidates ?? Enumerable.Empty<StandardRvtChangeCandidateEntry>()).Where([SpecialName] (StandardRvtChangeCandidateEntry x) => x != null).ToList();
		if (doc == null || target == null || list.Count == 0)
		{
			return;
		}
		string documentKey = BuildDocumentPendingKey(doc, target);
		if (string.IsNullOrWhiteSpace(documentKey))
		{
			return;
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (!PendingCandidatesByDocument.TryGetValue(documentKey, out var pending))
			{
				pending = new Dictionary<string, StandardRvtChangeCandidateEntry>(StringComparer.Ordinal);
				PendingCandidatesByDocument[documentKey] = pending;
			}
			foreach (StandardRvtChangeCandidateEntry entry in list)
			{
				AddOrMerge(pending, entry);
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	private static List<StandardRvtChangeCandidateEntry> DrainPendingCandidates(Document doc, StandardRvtTarget target)
	{
		List<StandardRvtChangeCandidateEntry> result = new List<StandardRvtChangeCandidateEntry>();
		string documentKey = BuildDocumentPendingKey(doc, target);
		if (string.IsNullOrWhiteSpace(documentKey))
		{
			return result;
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (PendingCandidatesByDocument.TryGetValue(documentKey, out var pending))
			{
				result = pending.Values.Where([SpecialName] (StandardRvtChangeCandidateEntry x) => x != null).ToList();
				PendingCandidatesByDocument.Remove(documentKey);
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		return result;
	}

	private static string BuildDocumentPendingKey(Document doc, StandardRvtTarget target)
	{
		string path = NormalizePathForCompare(doc?.PathName);
		if (string.IsNullOrWhiteSpace(path))
		{
			return string.Empty;
		}
		return Normalize(target?.SourceId) + "|" + path;
	}

	private static void CollectSnapshotDeltaCandidates(IDictionary<string, StandardRvtChangeCandidateEntry> entries, Document doc, StandardRvtTarget target)
	{
		StandardLibrarySnapshot snapshot = LoadSnapshot(target?.SnapshotPath);
		if (entries == null || doc == null || target == null || snapshot == null)
		{
			return;
		}
		CollectLoadableFamilySnapshotDeltaCandidates(entries, doc, target, snapshot);
		CollectSystemTypeSnapshotDeltaCandidates(entries, doc, target, snapshot);
		EnrichCandidateFingerprintSummaries(entries, snapshot);
	}

	private static void EnrichCandidateFingerprintSummaries(IDictionary<string, StandardRvtChangeCandidateEntry> entries, StandardLibrarySnapshot snapshot)
	{
		if (entries == null || snapshot == null)
		{
			return;
		}
		Dictionary<string, StandardLoadableFamilySnapshotItem> families = (snapshot.LoadableFamilies ?? new List<StandardLoadableFamilySnapshotItem>()).Where([SpecialName] (StandardLoadableFamilySnapshotItem x) => x != null).GroupBy<StandardLoadableFamilySnapshotItem, string>([SpecialName] (StandardLoadableFamilySnapshotItem x) => BuildLoadableFamilyKey(x.CategoryName, x.FamilyName), StringComparer.Ordinal).ToDictionary<IGrouping<string, StandardLoadableFamilySnapshotItem>, string, StandardLoadableFamilySnapshotItem>([SpecialName] (IGrouping<string, StandardLoadableFamilySnapshotItem> x) => x.Key, [SpecialName] (IGrouping<string, StandardLoadableFamilySnapshotItem> x) => x.First(), StringComparer.Ordinal);
		Dictionary<string, StandardSystemTypeSnapshotItem> systems = (snapshot.SystemTypes ?? new List<StandardSystemTypeSnapshotItem>()).Where([SpecialName] (StandardSystemTypeSnapshotItem x) => x != null).GroupBy<StandardSystemTypeSnapshotItem, string>([SpecialName] (StandardSystemTypeSnapshotItem x) => BuildSystemTypeKey(x.TypeClassName, x.CategoryName, x.TypeName), StringComparer.Ordinal).ToDictionary<IGrouping<string, StandardSystemTypeSnapshotItem>, string, StandardSystemTypeSnapshotItem>([SpecialName] (IGrouping<string, StandardSystemTypeSnapshotItem> x) => x.Key, [SpecialName] (IGrouping<string, StandardSystemTypeSnapshotItem> x) => x.First(), StringComparer.Ordinal);
		foreach (StandardRvtChangeCandidateEntry entry in entries.Values.Where([SpecialName] (StandardRvtChangeCandidateEntry x) => x != null))
		{
			if (string.Equals(entry.ChangeKind, "Added", StringComparison.OrdinalIgnoreCase))
			{
				entry.BeforeFingerprint = "missing";
				entry.AfterFingerprint = "present; precise rescan required";
				continue;
			}
			if (string.Equals(entry.CandidateKind, "LoadableFamily", StringComparison.OrdinalIgnoreCase))
			{
				StandardLoadableFamilySnapshotItem family;
				if (families.TryGetValue(BuildLoadableFamilyKey(entry.CategoryName, entry.FamilyName), out family))
				{
					entry.BeforeFingerprint = CompactFingerprint(family.ContentFingerprint);
				}
			}
			else if (string.Equals(entry.CandidateKind, "SystemType", StringComparison.OrdinalIgnoreCase))
			{
				StandardSystemTypeSnapshotItem system;
				if (systems.TryGetValue(BuildSystemTypeKey(entry.SystemFamilyKind, entry.CategoryName, entry.TypeName), out system))
				{
					entry.BeforeFingerprint = CompactFingerprint(FirstNonEmpty(system.SemanticFingerprint, system.DetailedComponentSignature, system.RoutingPreferenceSignature, system.CompoundStructureSignature));
				}
			}
			entry.AfterFingerprint = string.Equals(entry.ChangeKind, "Deleted", StringComparison.OrdinalIgnoreCase) ? "missing" : "changed; precise rescan required";
		}
	}

	private static string FirstNonEmpty(params string[] values)
	{
		return (values ?? new string[0]).FirstOrDefault([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)) ?? string.Empty;
	}

	private static string CompactFingerprint(string value)
	{
		string text = (value ?? string.Empty).Trim();
		if (text.Length == 0)
		{
			return "last-scan fingerprint unavailable";
		}
		return text.Length <= 24 ? text : text.Substring(0, 24) + "...";
	}

	private static void CollectLoadableFamilySnapshotDeltaCandidates(IDictionary<string, StandardRvtChangeCandidateEntry> entries, Document doc, StandardRvtTarget target, StandardLibrarySnapshot snapshot)
	{
		Dictionary<string, StandardLoadableFamilySnapshotItem> standardFamilies = (snapshot.LoadableFamilies ?? new List<StandardLoadableFamilySnapshotItem>()).Where([SpecialName] (StandardLoadableFamilySnapshotItem x) => x != null && !string.IsNullOrWhiteSpace(x.FamilyName)).GroupBy<StandardLoadableFamilySnapshotItem, string>([SpecialName] (StandardLoadableFamilySnapshotItem x) => BuildLoadableFamilyKey(x.CategoryName, x.FamilyName), StringComparer.Ordinal).ToDictionary<IGrouping<string, StandardLoadableFamilySnapshotItem>, string, StandardLoadableFamilySnapshotItem>([SpecialName] (IGrouping<string, StandardLoadableFamilySnapshotItem> g) => g.Key, [SpecialName] (IGrouping<string, StandardLoadableFamilySnapshotItem> g) => g.First(), StringComparer.Ordinal);
		Dictionary<string, Family> currentFamilies = new Dictionary<string, Family>(StringComparer.Ordinal);
		foreach (Family family in new FilteredElementCollector(doc).OfClass(typeof(Family)).Cast<Family>())
		{
			if (family != null && FamilyBrowserFamilyClassificationService.IsBrowserLoadableFamily(family))
			{
				currentFamilies[BuildLoadableFamilyKey(FamilyBrowserFamilyClassificationService.ResolveCategoryName(family), family.Name)] = family;
			}
		}
		foreach (KeyValuePair<string, Family> pair in currentFamilies)
		{
			if (!standardFamilies.ContainsKey(pair.Key))
			{
				Family family = pair.Value;
				StandardRvtChangeCandidateEntry entry = CreateBaseEntry(doc, target, "LoadableFamily", "Added", "FamilyAddedAfterLastScan");
				entry.CategoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family);
				entry.CategoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family);
				entry.FamilyName = family.Name ?? string.Empty;
				entry.Details = "SavedStandardRvtSnapshotDelta";
				AddOrMerge(entries, entry);
			}
		}
		foreach (KeyValuePair<string, StandardLoadableFamilySnapshotItem> pair2 in standardFamilies)
		{
			if (!currentFamilies.ContainsKey(pair2.Key))
			{
				StandardLoadableFamilySnapshotItem item = pair2.Value;
				StandardRvtChangeCandidateEntry entry2 = CreateBaseEntry(doc, target, "LoadableFamily", "Deleted", "FamilyDeletedAfterLastScan");
				entry2.CategoryName = item.CategoryName ?? string.Empty;
				entry2.CategoryId = item.CategoryId ?? string.Empty;
				entry2.FamilyName = item.FamilyName ?? string.Empty;
				entry2.Details = "SavedStandardRvtSnapshotDelta";
				AddOrMerge(entries, entry2);
			}
		}
		CollectLoadableTypeSnapshotDeltaCandidates(entries, doc, target, standardFamilies, currentFamilies);
	}

	private static void CollectLoadableTypeSnapshotDeltaCandidates(IDictionary<string, StandardRvtChangeCandidateEntry> entries, Document doc, StandardRvtTarget target, IDictionary<string, StandardLoadableFamilySnapshotItem> standardFamilies, IDictionary<string, Family> currentFamilies)
	{
		Dictionary<string, StandardLoadableFamilySnapshotItem> standardTypes = new Dictionary<string, StandardLoadableFamilySnapshotItem>(StringComparer.Ordinal);
		Dictionary<string, string> standardTypeNames = new Dictionary<string, string>(StringComparer.Ordinal);
		foreach (StandardLoadableFamilySnapshotItem family in standardFamilies.Values)
		{
			foreach (string typeName in family.TypeNames ?? new List<string>())
			{
				if (!string.IsNullOrWhiteSpace(typeName))
				{
					string key = BuildLoadableTypeKey(family.CategoryName, family.FamilyName, typeName);
					standardTypes[key] = family;
					standardTypeNames[key] = typeName;
				}
			}
		}
		Dictionary<string, FamilySymbol> currentTypes = new Dictionary<string, FamilySymbol>(StringComparer.Ordinal);
		foreach (FamilySymbol symbol in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>())
		{
			Family family = symbol?.Family;
			if (family != null && FamilyBrowserFamilyClassificationService.IsBrowserLoadableFamily(family))
			{
				currentTypes[BuildLoadableTypeKey(FamilyBrowserFamilyClassificationService.ResolveCategoryName(family), family.Name, ResolveElementName(symbol))] = symbol;
			}
		}
		foreach (KeyValuePair<string, FamilySymbol> pair in currentTypes)
		{
			Family family2 = pair.Value?.Family;
			if (family2 != null && currentFamilies.ContainsKey(BuildLoadableFamilyKey(FamilyBrowserFamilyClassificationService.ResolveCategoryName(family2), family2.Name)) && standardFamilies.ContainsKey(BuildLoadableFamilyKey(FamilyBrowserFamilyClassificationService.ResolveCategoryName(family2), family2.Name)) && !standardTypes.ContainsKey(pair.Key))
			{
				StandardRvtChangeCandidateEntry entry = CreateBaseEntry(doc, target, "LoadableFamily", "Added", "TypeAddedAfterLastScan");
				entry.CategoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family2);
				entry.CategoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family2);
				entry.FamilyName = family2.Name ?? string.Empty;
				entry.TypeName = ResolveElementName(pair.Value);
				entry.Details = "SavedStandardRvtSnapshotDelta";
				AddOrMerge(entries, entry);
			}
		}
		foreach (KeyValuePair<string, StandardLoadableFamilySnapshotItem> pair2 in standardTypes)
		{
			if (!currentTypes.ContainsKey(pair2.Key))
			{
				StandardLoadableFamilySnapshotItem item = pair2.Value;
				if (!currentFamilies.ContainsKey(BuildLoadableFamilyKey(item.CategoryName, item.FamilyName)))
				{
					continue;
				}
				StandardRvtChangeCandidateEntry entry2 = CreateBaseEntry(doc, target, "LoadableFamily", "Deleted", "TypeDeletedAfterLastScan");
				entry2.CategoryName = item.CategoryName ?? string.Empty;
				entry2.CategoryId = item.CategoryId ?? string.Empty;
				entry2.FamilyName = item.FamilyName ?? string.Empty;
				entry2.TypeName = standardTypeNames.TryGetValue(pair2.Key, out var deletedTypeName) ? deletedTypeName : ExtractLastKeyPart(pair2.Key);
				entry2.Details = "SavedStandardRvtSnapshotDelta";
				AddOrMerge(entries, entry2);
			}
		}
	}

	private static void CollectSystemTypeSnapshotDeltaCandidates(IDictionary<string, StandardRvtChangeCandidateEntry> entries, Document doc, StandardRvtTarget target, StandardLibrarySnapshot snapshot)
	{
		Dictionary<string, StandardSystemTypeSnapshotItem> standardTypes = (snapshot.SystemTypes ?? new List<StandardSystemTypeSnapshotItem>()).Where([SpecialName] (StandardSystemTypeSnapshotItem x) => x != null && !string.IsNullOrWhiteSpace(x.TypeName)).GroupBy<StandardSystemTypeSnapshotItem, string>([SpecialName] (StandardSystemTypeSnapshotItem x) => BuildSystemTypeKey(x.TypeClassName, x.CategoryName, x.TypeName), StringComparer.Ordinal).ToDictionary<IGrouping<string, StandardSystemTypeSnapshotItem>, string, StandardSystemTypeSnapshotItem>([SpecialName] (IGrouping<string, StandardSystemTypeSnapshotItem> g) => g.Key, [SpecialName] (IGrouping<string, StandardSystemTypeSnapshotItem> g) => g.First(), StringComparer.Ordinal);
		Dictionary<string, ElementType> currentTypes = new Dictionary<string, ElementType>(StringComparer.Ordinal);
		foreach (ElementType elementType in new FilteredElementCollector(doc).WhereElementIsElementType().Cast<ElementType>())
		{
			if (elementType != null && (!(elementType is FamilySymbol) || SystemTypeDetailedComponentSnapshotService.SupportsRequiredCurtainPanelComponents(elementType.GetType().Name)) && AllowedSystemTypeNames.Contains(elementType.GetType().Name))
			{
				currentTypes[BuildSystemTypeKey(elementType.GetType().Name, ResolveCategoryName(elementType), ResolveElementName(elementType))] = elementType;
			}
		}
		foreach (KeyValuePair<string, ElementType> pair in currentTypes)
		{
			if (!standardTypes.ContainsKey(pair.Key))
			{
				ElementType elementType2 = pair.Value;
				StandardRvtChangeCandidateEntry entry = CreateBaseEntry(doc, target, "SystemType", "Added", "SystemTypeAddedAfterLastScan");
				entry.CategoryName = ResolveCategoryName(elementType2);
				entry.CategoryId = ResolveCategoryId(elementType2);
				entry.SystemFamilyKind = elementType2.GetType().Name;
				entry.TypeName = ResolveElementName(elementType2);
				entry.Details = "SavedStandardRvtSnapshotDelta";
				AddOrMerge(entries, entry);
			}
		}
		foreach (KeyValuePair<string, StandardSystemTypeSnapshotItem> pair2 in standardTypes)
		{
			if (!currentTypes.ContainsKey(pair2.Key))
			{
				StandardSystemTypeSnapshotItem item = pair2.Value;
				StandardRvtChangeCandidateEntry entry2 = CreateBaseEntry(doc, target, "SystemType", "Deleted", "SystemTypeDeletedAfterLastScan");
				entry2.CategoryName = item.CategoryName ?? string.Empty;
				entry2.CategoryId = item.CategoryId ?? string.Empty;
				entry2.SystemFamilyKind = item.TypeClassName ?? string.Empty;
				entry2.TypeName = item.TypeName ?? string.Empty;
				entry2.Details = "SavedStandardRvtSnapshotDelta";
				AddOrMerge(entries, entry2);
			}
		}
	}

	private static void CollectCandidates(IDictionary<string, StandardRvtChangeCandidateEntry> entries, Document doc, ICollection<ElementId> elementIds, StandardRvtTarget target, string changeKind)
	{
		if (entries == null || doc == null || elementIds == null || elementIds.Count == 0)
		{
			return;
		}
		foreach (ElementId elementId in elementIds)
		{
			Element element = null;
			try
			{
				element = doc.GetElement(elementId);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
				continue;
			}
			if (element != null)
			{
				AddLoadableFamilyCandidate(entries, doc, target, element, changeKind);
				AddSystemTypeCandidate(entries, doc, target, element, changeKind);
			}
		}
	}

	private static void AddLoadableFamilyCandidate(IDictionary<string, StandardRvtChangeCandidateEntry> entries, Document doc, StandardRvtTarget target, Element element, string changeKind)
	{
		Family family = element as Family;
		FamilySymbol symbol = element as FamilySymbol;
		FamilyInstance familyInstance = element as FamilyInstance;
		if (family == null && symbol != null)
		{
			family = symbol.Family;
		}
		if (family == null && familyInstance != null && familyInstance.Symbol != null)
		{
			family = familyInstance.Symbol.Family;
		}
		if (family != null && FamilyBrowserFamilyClassificationService.IsBrowserLoadableFamily(family))
		{
			string typeName = ((symbol == null) ? string.Empty : ResolveElementName(symbol));
			if (string.IsNullOrWhiteSpace(typeName) && familyInstance != null && familyInstance.Symbol != null)
			{
				typeName = ResolveElementName(familyInstance.Symbol);
			}
			string reason = ((element is FamilySymbol) ? (string.Equals(changeKind, "Added", StringComparison.OrdinalIgnoreCase) ? "TypeAdded" : "TypeChanged") : ((!(element is Family)) ? "FamilyReferenceChanged" : (string.Equals(changeKind, "Added", StringComparison.OrdinalIgnoreCase) ? "FamilyAdded" : "FamilyChanged")));
			StandardRvtChangeCandidateEntry entry = CreateBaseEntry(doc, target, "LoadableFamily", changeKind, reason);
			entry.CategoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family);
			entry.CategoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family);
			entry.FamilyName = family.Name ?? string.Empty;
			entry.TypeName = typeName;
			entry.Details = element.GetType().Name;
			AddOrMerge(entries, entry);
		}
	}

	private static void AddSystemTypeCandidate(IDictionary<string, StandardRvtChangeCandidateEntry> entries, Document doc, StandardRvtTarget target, Element element, string changeKind)
	{
		if (element is ElementType elementType && (!(elementType is FamilySymbol) || SystemTypeDetailedComponentSnapshotService.SupportsRequiredCurtainPanelComponents(elementType.GetType().Name)))
		{
			string typeClassName = elementType.GetType().Name;
			if (AllowedSystemTypeNames.Contains(typeClassName))
			{
				StandardRvtChangeCandidateEntry entry = CreateBaseEntry(doc, target, "SystemType", changeKind, string.Equals(changeKind, "Added", StringComparison.OrdinalIgnoreCase) ? "SystemTypeAdded" : "SystemTypeChanged");
				entry.CategoryName = ResolveCategoryName(elementType);
				entry.CategoryId = ResolveCategoryId(elementType);
				entry.SystemFamilyKind = typeClassName;
				entry.TypeName = ResolveElementName(elementType);
				entry.Details = element.GetType().Name;
				AddOrMerge(entries, entry);
			}
		}
	}

	private static StandardRvtChangeCandidateEntry CreateBaseEntry(Document doc, StandardRvtTarget target, string candidateKind, string changeKind, string reason)
	{
		StandardRvtChangeCandidateEntry standardRvtChangeCandidateEntry = new StandardRvtChangeCandidateEntry();
		standardRvtChangeCandidateEntry.EntryId = Guid.NewGuid().ToString("N");
		standardRvtChangeCandidateEntry.RecordedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		standardRvtChangeCandidateEntry.UserName = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		try
		{
			standardRvtChangeCandidateEntry.RevitUserName = doc == null || doc.Application == null ? string.Empty : doc.Application.Username ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			standardRvtChangeCandidateEntry.RevitUserName = string.Empty;
			ProjectData.ClearProjectError();
		}
		standardRvtChangeCandidateEntry.MachineName = Environment.MachineName ?? string.Empty;
		standardRvtChangeCandidateEntry.DocumentTitle = doc?.Title ?? string.Empty;
		standardRvtChangeCandidateEntry.DocumentPath = doc?.PathName ?? string.Empty;
		standardRvtChangeCandidateEntry.CanonicalDocumentIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(standardRvtChangeCandidateEntry.DocumentPath);
		standardRvtChangeCandidateEntry.SourceId = target?.SourceId ?? string.Empty;
		standardRvtChangeCandidateEntry.SlotKey = target?.SlotKey ?? string.Empty;
		standardRvtChangeCandidateEntry.DisciplineKey = target?.DisciplineKey ?? string.Empty;
		standardRvtChangeCandidateEntry.DisciplineLabel = target?.DisciplineLabel ?? string.Empty;
		standardRvtChangeCandidateEntry.CandidateKind = candidateKind;
		standardRvtChangeCandidateEntry.ChangeKind = changeKind;
		standardRvtChangeCandidateEntry.Reason = reason;
		standardRvtChangeCandidateEntry.CommitState = "Pending";
		return standardRvtChangeCandidateEntry;
	}

	private static void AddOrMerge(IDictionary<string, StandardRvtChangeCandidateEntry> entries, StandardRvtChangeCandidateEntry entry)
	{
		if (entry != null)
		{
			string key = Normalize(entry.CandidateKind) + "|" + Normalize(entry.CategoryId) + "|" + Normalize(entry.CategoryName) + "|" + Normalize(entry.FamilyName) + "|" + Normalize(entry.SystemFamilyKind) + "|" + Normalize(entry.TypeName) + "|" + Normalize(entry.Reason);
			if (!string.IsNullOrWhiteSpace(key.Replace("|", string.Empty)) && !entries.ContainsKey(key))
			{
				entries.Add(key, entry);
			}
		}
	}

	private static StandardRvtTarget ResolveTargetForDocument(string workspaceRoot, Document doc)
	{
		List<string> documentPaths = BuildDocumentComparePaths(doc);
		if (documentPaths.Count == 0)
		{
			return null;
		}
		return GetCachedTargets(workspaceRoot).FirstOrDefault((StandardRvtTarget x) => documentPaths.Any((string path) => string.Equals(path, NormalizePathForCompare(x.StandardRvtPath), StringComparison.OrdinalIgnoreCase)));
	}

	private static List<string> BuildDocumentComparePaths(Document doc)
	{
		List<string> result = new List<string>();
		if (doc == null)
		{
			return result;
		}
		try
		{
			AddDocumentComparePath(result, doc.PathName);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			if (doc.IsWorkshared)
			{
				ModelPath modelPath = doc.GetWorksharingCentralModelPath();
				AddDocumentComparePath(result, (modelPath == null) ? string.Empty : ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath));
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static void AddDocumentComparePath(IList<string> paths, string path)
	{
		if (paths == null)
		{
			return;
		}
		string normalized = NormalizePathForCompare(path);
		if (!string.IsNullOrWhiteSpace(normalized) && !paths.Any((string x) => string.Equals(x, normalized, StringComparison.OrdinalIgnoreCase)))
		{
			paths.Add(normalized);
		}
	}

	private static List<StandardRvtTarget> GetCachedTargets(string workspaceRoot)
	{
		DateTime nowUtc = DateTime.UtcNow;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (string.Equals(CachedWorkspaceRoot, workspaceRoot ?? string.Empty, StringComparison.OrdinalIgnoreCase) && CachedTargets != null && CachedTargets.Count > 0 && (nowUtc - CachedAtUtc).TotalSeconds < 30.0)
			{
				return CachedTargets.ToList();
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		List<StandardRvtTarget> loaded = LoadTargets(workspaceRoot);
		object syncRoot2 = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot2);
		bool lockTaken2 = false;
		try
		{
			Monitor.Enter(syncRoot2, ref lockTaken2);
			CachedWorkspaceRoot = workspaceRoot ?? string.Empty;
			CachedAtUtc = nowUtc;
			CachedTargets = loaded;
			return CachedTargets.ToList();
		}
		finally
		{
			if (lockTaken2)
			{
				Monitor.Exit(syncRoot2);
			}
		}
	}

	private static List<StandardRvtTarget> LoadTargets(string workspaceRoot)
	{
		List<StandardRvtTarget> result = new List<StandardRvtTarget>();
		try
		{
			FamilyBrowserStandardPolicy policy = FamilyBrowserStandardPolicyStore.LoadOrCreate(workspaceRoot, Environment.UserName);
			AddTarget(result, policy, policy?.IntegratedLibrary);
			foreach (FamilyBrowserStandardLibrarySlot slot in FamilyBrowserStandardPolicyStore.GetDisciplineSlots(policy))
			{
				AddTarget(result, policy, slot);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return (from g in result.Where([SpecialName] (StandardRvtTarget x) => x != null && !string.IsNullOrWhiteSpace(x.StandardRvtPath)).GroupBy<StandardRvtTarget, string>([SpecialName] (StandardRvtTarget x) => NormalizePathForCompare(x.StandardRvtPath), StringComparer.OrdinalIgnoreCase)
			select g.First()).ToList();
	}

	private static void AddTarget(IList<StandardRvtTarget> result, FamilyBrowserStandardPolicy policy, FamilyBrowserStandardLibrarySlot slot)
	{
		if (result != null && slot != null && !string.IsNullOrWhiteSpace(slot.StandardRvtPath))
		{
			result.Add(new StandardRvtTarget
			{
				SourceId = (slot.SourceId ?? string.Empty),
				StandardRvtPath = slot.StandardRvtPath,
				SnapshotPath = (slot.SnapshotPath ?? string.Empty),
				SlotKey = (slot.SlotKey ?? string.Empty),
				DisciplineKey = (slot.Discipline ?? string.Empty),
				DisciplineLabel = FamilyBrowserStandardPolicyStore.ResolveSlotDisplayName(slot, korean: false)
			});
		}
	}

	private static bool AppendCandidates(string workspaceRoot, StandardRvtTarget target, IEnumerable<StandardRvtChangeCandidateEntry> candidates)
	{
		List<StandardRvtChangeCandidateEntry> list = (candidates ?? Enumerable.Empty<StandardRvtChangeCandidateEntry>()).Where([SpecialName] (StandardRvtChangeCandidateEntry x) => x != null).ToList();
		if (list.Count == 0 || target == null || string.IsNullOrWhiteSpace(target.SourceId))
		{
			return false;
		}
		if (!AppendImmutableCandidates(workspaceRoot, target.SourceId, list))
		{
			return false;
		}
		if (!FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
		{
			return true;
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			StandardRvtChangeCandidateLog log = LoadLog(workspaceRoot, target.SourceId);
			log.SourceId = target.SourceId;
			log.StandardRvtPath = target.StandardRvtPath ?? string.Empty;
			log.UpdatedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
			log.Entries.AddRange(list);
			log.Entries = log.Entries.Where([SpecialName] (StandardRvtChangeCandidateEntry x) => x != null).GroupBy<StandardRvtChangeCandidateEntry, string>([SpecialName] (StandardRvtChangeCandidateEntry x) => string.IsNullOrWhiteSpace(x.EntryId) ? BuildCandidateIdentity(x) : x.EntryId, StringComparer.OrdinalIgnoreCase).Select([SpecialName] (IGrouping<string, StandardRvtChangeCandidateEntry> x) => x.OrderByDescending([SpecialName] (StandardRvtChangeCandidateEntry y) => y.CommittedAtUtc ?? y.RecordedAtUtc ?? string.Empty, StringComparer.Ordinal).First()).OrderByDescending<StandardRvtChangeCandidateEntry, string>([SpecialName] (StandardRvtChangeCandidateEntry x) => string.IsNullOrWhiteSpace(x.CommittedAtUtc) ? x.RecordedAtUtc ?? string.Empty : x.CommittedAtUtc, StringComparer.Ordinal).Take(MaxLogEntriesPerSource)
				.OrderBy<StandardRvtChangeCandidateEntry, string>([SpecialName] (StandardRvtChangeCandidateEntry x) => x.RecordedAtUtc ?? string.Empty, StringComparer.Ordinal)
				.ToList();
			SaveLog(workspaceRoot, log);
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		return true;
	}

	private static bool AppendImmutableCandidates(string workspaceRoot, string sourceId, IEnumerable<StandardRvtChangeCandidateEntry> entries)
	{
		return FamilyBrowserTrackingPersistenceService.PersistStandardCandidateEntries(workspaceRoot, sourceId, entries);
	}

	private static bool AppendOperationEntries(string workspaceRoot, IEnumerable<FamilyBrowserOperationLogEntry> entries)
	{
		List<FamilyBrowserOperationLogEntry> list = (entries ?? Enumerable.Empty<FamilyBrowserOperationLogEntry>()).Where([SpecialName] (FamilyBrowserOperationLogEntry x) => x != null).ToList();
		if (list.Count == 0)
		{
			return true;
		}
		if (!FamilyBrowserTrackingPersistenceService.PersistOperationEntries(workspaceRoot, list))
		{
			return false;
		}
		if (!FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
		{
			return true;
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			string logPath = BuildOperationLogPath(workspaceRoot, DateTime.UtcNow);
			FamilyBrowserOperationLog log = LoadOperationLog(logPath);
			log.LogDate = DateTime.UtcNow.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
			log.UpdatedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
			log.Entries.AddRange(list);
			if (log.Entries.Count > 1000)
			{
				log.Entries = log.Entries.OrderByDescending<FamilyBrowserOperationLogEntry, string>([SpecialName] (FamilyBrowserOperationLogEntry x) => x.RecordedAtUtc ?? string.Empty, StringComparer.Ordinal).Take(1000).OrderBy<FamilyBrowserOperationLogEntry, string>([SpecialName] (FamilyBrowserOperationLogEntry x) => x.RecordedAtUtc ?? string.Empty, StringComparer.Ordinal)
					.ToList();
			}
			try
			{
				Directory.CreateDirectory(Path.GetDirectoryName(logPath));
				File.WriteAllText(logPath, PlainJsonReportWriter.Serialize(log), Encoding.UTF8);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		return true;
	}

	private static void QueueOperationEntries(string workspaceRoot, Document doc, IEnumerable<FamilyBrowserOperationLogEntry> entries)
	{
		List<FamilyBrowserOperationLogEntry> list = (entries ?? Enumerable.Empty<FamilyBrowserOperationLogEntry>()).Where([SpecialName] (FamilyBrowserOperationLogEntry x) => x != null).ToList();
		if (list.Count == 0 || doc == null)
		{
			return;
		}
		string key = BuildPendingOperationKey(workspaceRoot, doc);
		if (string.IsNullOrWhiteSpace(key))
		{
			return;
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			PendingOperationBatch batch = null;
			if (!PendingOperationEntriesByDocument.TryGetValue(key, out batch) || batch == null)
			{
				batch = new PendingOperationBatch
				{
					WorkspaceRoot = (workspaceRoot ?? string.Empty),
					DocumentKey = key
				};
				PendingOperationEntriesByDocument[key] = batch;
			}
			if (batch.Entries == null)
			{
				batch.Entries = new List<FamilyBrowserOperationLogEntry>();
			}
			batch.Entries.AddRange(list);
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	private static bool IsSuccessfulRevitApiEventStatus(object status)
	{
		if (status == null)
		{
			return false;
		}
		return string.Equals(status.ToString(), "Succeeded", StringComparison.OrdinalIgnoreCase);
	}

	private static bool CommitPendingOperationEntries(Document doc, string commitKind)
	{
		if (doc == null)
		{
			return false;
		}
		string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
		string key = BuildPendingOperationKey(workspaceRoot, doc);
		if (string.IsNullOrWhiteSpace(key))
		{
			return false;
		}
		PendingOperationBatch batch = null;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (!PendingOperationEntriesByDocument.TryGetValue(key, out batch) || batch == null)
			{
				return false;
			}
			PendingOperationEntriesByDocument.Remove(key);
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		List<FamilyBrowserOperationLogEntry> list = (batch.Entries ?? new List<FamilyBrowserOperationLogEntry>()).Where([SpecialName] (FamilyBrowserOperationLogEntry x) => x != null).ToList();
		if (list.Count == 0)
		{
			return false;
		}
		string committedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		string documentPath = ResolveOperationDocumentIdentity(doc);
		foreach (FamilyBrowserOperationLogEntry entry in list)
		{
			bool notCommitted = ShouldVerifyCommittedOperation(entry) && !IsPendingOperationPresentAtSync(doc, entry);
			entry.CommitState = notCommitted ? "NotCommitted" : "Committed";
			entry.CommitKind = commitKind ?? string.Empty;
			entry.CommittedAtUtc = committedAtUtc;
			if (notCommitted)
			{
				entry.Outcome = "NotCommitted";
				entry.Details = AppendOperationDetail(entry.Details, "The loaded/applied item was not present when the save or synchronization was confirmed.");
			}
			if (!string.IsNullOrWhiteSpace(documentPath))
			{
				entry.DocumentPath = documentPath;
			}
		}
		if (AppendOperationEntries(batch.WorkspaceRoot, list))
		{
			return true;
		}
		RestorePendingOperationBatch(key, batch);
		return false;
	}

	private static void RestorePendingOperationBatch(string key, PendingOperationBatch batch)
	{
		if (string.IsNullOrWhiteSpace(key) || batch == null)
		{
			return;
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			PendingOperationBatch current = null;
			if (!PendingOperationEntriesByDocument.TryGetValue(key, out current) || current == null)
			{
				PendingOperationEntriesByDocument[key] = batch;
				return;
			}
			current.Entries.AddRange(batch.Entries ?? new List<FamilyBrowserOperationLogEntry>());
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	private static bool ShouldVerifyCommittedOperation(FamilyBrowserOperationLogEntry entry)
	{
		string outcome = Normalize(entry?.Outcome);
		if (string.IsNullOrWhiteSpace(outcome))
		{
			return false;
		}
		return outcome.Contains("loaded") || outcome.Contains("created") || outcome.Contains("overwritten") || outcome.Contains("updated") || outcome.Contains("applied") || outcome.Contains("copied") || outcome.Contains("consolidated");
	}

	private static FamilyBrowserOperationLogEntry CloneOperationEntry(FamilyBrowserOperationLogEntry source)
	{
		if (source == null)
		{
			return new FamilyBrowserOperationLogEntry();
		}
		return new FamilyBrowserOperationLogEntry
		{
			EntryId = source.EntryId,
			RecordedAtUtc = source.RecordedAtUtc,
			UserName = source.UserName,
			OperationKind = source.OperationKind,
			DocumentTitle = source.DocumentTitle,
			DocumentPath = source.DocumentPath,
			SourceId = source.SourceId,
			StandardSourceKey = source.StandardSourceKey,
			StandardDisplayName = source.StandardDisplayName,
			CandidateKind = source.CandidateKind,
			CategoryName = source.CategoryName,
			FamilyName = source.FamilyName,
			TypeName = source.TypeName,
			SystemFamilyKind = source.SystemFamilyKind,
			PlannedAction = source.PlannedAction,
			Outcome = source.Outcome,
			Details = source.Details,
			CommitState = source.CommitState,
			CommitKind = source.CommitKind,
			CommittedAtUtc = source.CommittedAtUtc
		};
	}

	private static bool IsPendingOperationPresentAtSync(Document doc, FamilyBrowserOperationLogEntry entry)
	{
		if (doc == null || entry == null)
		{
			return false;
		}
		try
		{
			if (string.Equals(entry.CandidateKind ?? string.Empty, "LoadableFamily", StringComparison.OrdinalIgnoreCase))
			{
				if (string.IsNullOrWhiteSpace(entry.FamilyName))
				{
					return true;
				}
				foreach (Family family in new FilteredElementCollector(doc).OfClass(typeof(Family)).Cast<Family>())
				{
					if (family == null || !string.Equals(Normalize(family.Name), Normalize(entry.FamilyName), StringComparison.Ordinal))
					{
						continue;
					}
					if (!string.IsNullOrWhiteSpace(entry.CategoryName) && !string.Equals(Normalize(family.FamilyCategory?.Name ?? string.Empty), Normalize(entry.CategoryName), StringComparison.Ordinal))
					{
						continue;
					}
					if (string.IsNullOrWhiteSpace(entry.TypeName))
					{
						return true;
					}
					foreach (ElementId symbolId in family.GetFamilySymbolIds())
					{
						FamilySymbol symbol = doc.GetElement(symbolId) as FamilySymbol;
						if (string.Equals(Normalize(ResolveElementName(symbol)), Normalize(entry.TypeName), StringComparison.Ordinal))
						{
							return true;
						}
					}
				}
				return false;
			}
			if (string.Equals(entry.CandidateKind ?? string.Empty, "SystemType", StringComparison.OrdinalIgnoreCase))
			{
				if (string.IsNullOrWhiteSpace(entry.TypeName))
				{
					return true;
				}
				foreach (ElementType elementType in new FilteredElementCollector(doc).WhereElementIsElementType().Cast<ElementType>())
				{
					if (elementType == null || (elementType is FamilySymbol && !SystemTypeDetailedComponentSnapshotService.SupportsRequiredCurtainPanelComponents(elementType.GetType().Name)))
					{
						continue;
					}
					if (!string.IsNullOrWhiteSpace(entry.SystemFamilyKind) && !string.Equals(elementType.GetType().Name, entry.SystemFamilyKind, StringComparison.OrdinalIgnoreCase))
					{
						continue;
					}
					if (!string.Equals(Normalize(ResolveElementName(elementType)), Normalize(entry.TypeName), StringComparison.Ordinal))
					{
						continue;
					}
					if (!string.IsNullOrWhiteSpace(entry.CategoryName) && !string.Equals(Normalize(ResolveCategoryName(elementType)), Normalize(entry.CategoryName), StringComparison.Ordinal))
					{
						continue;
					}
					return true;
				}
				return false;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
			return false;
		}
		return true;
	}

	private static string AppendOperationDetail(string details, string message)
	{
		string safeDetails = details ?? string.Empty;
		string safeMessage = message ?? string.Empty;
		if (string.IsNullOrWhiteSpace(safeMessage))
		{
			return safeDetails;
		}
		if (string.IsNullOrWhiteSpace(safeDetails))
		{
			return safeMessage;
		}
		return safeDetails.Trim() + " | " + safeMessage;
	}

	private static string BuildPendingOperationKey(string workspaceRoot, Document doc)
	{
		if (doc == null)
		{
			return string.Empty;
		}
		// The path changes after Save As and the managed root can change after policy refresh.
		// Pending state belongs to this open Revit document instance only.
		return "runtime:" + RuntimeHelpers.GetHashCode(doc).ToString(CultureInfo.InvariantCulture);
	}

	private static string ResolveOperationDocumentIdentity(Document doc)
	{
		if (doc == null)
		{
			return string.Empty;
		}
		try
		{
			string identityPath = ProjectSnapshotStore.ResolveProjectIdentityPath(doc);
			if (!string.IsNullOrWhiteSpace(identityPath))
			{
				return identityPath;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			return doc.PathName ?? string.Empty;
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
			return string.Empty;
		}
	}

	private static FamilyBrowserOperationLog LoadOperationLog(string logPath)
	{
		FamilyBrowserOperationLog LoadOperationLog;
		if (string.IsNullOrWhiteSpace(logPath) || !File.Exists(logPath))
		{
			LoadOperationLog = new FamilyBrowserOperationLog();
		}
		else
		{
			try
			{
				FamilyBrowserOperationLog log = DataContractJsonFileStore.Load<FamilyBrowserOperationLog>(logPath);
				if (log == null)
				{
					LoadOperationLog = new FamilyBrowserOperationLog();
				}
				else
				{
					if (log.Entries == null)
					{
						log.Entries = new List<FamilyBrowserOperationLogEntry>();
					}
					LoadOperationLog = log;
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				LoadOperationLog = new FamilyBrowserOperationLog();
				ProjectData.ClearProjectError();
			}
		}
		return LoadOperationLog;
	}

	private static string BuildOperationLogPath(string workspaceRoot, DateTime utcDate)
	{
		return Path.Combine(GetOperationLogFolder(workspaceRoot), "family-browser-operations-" + utcDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + ".json");
	}

	private static string GetOperationLogFolder(string workspaceRoot)
	{
		return FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "OperationLogs");
	}

	private static bool IsSameCandidate(StandardRvtChangeCandidateEntry left, StandardRvtChangeCandidateEntry right)
	{
		if (left == null || right == null)
		{
			return false;
		}
		return string.Equals(Normalize(left.CandidateKind), Normalize(right.CandidateKind), StringComparison.Ordinal) && string.Equals(Normalize(left.CategoryId), Normalize(right.CategoryId), StringComparison.Ordinal) && string.Equals(Normalize(left.CategoryName), Normalize(right.CategoryName), StringComparison.Ordinal) && string.Equals(Normalize(left.FamilyName), Normalize(right.FamilyName), StringComparison.Ordinal) && string.Equals(Normalize(left.SystemFamilyKind), Normalize(right.SystemFamilyKind), StringComparison.Ordinal) && string.Equals(Normalize(left.TypeName), Normalize(right.TypeName), StringComparison.Ordinal) && string.Equals(Normalize(left.Reason), Normalize(right.Reason), StringComparison.Ordinal);
	}

	private static StandardRvtChangeCandidateLog LoadLog(string workspaceRoot, string sourceId)
	{
		string path = BuildLogPath(workspaceRoot, sourceId);
		StandardRvtChangeCandidateLog LoadLog;
		if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
		{
			LoadLog = new StandardRvtChangeCandidateLog
			{
				SourceId = (sourceId ?? string.Empty)
			};
		}
		else
		{
			try
			{
				StandardRvtChangeCandidateLog log = DataContractJsonFileStore.Load<StandardRvtChangeCandidateLog>(path);
				if (log == null)
				{
					LoadLog = new StandardRvtChangeCandidateLog
					{
						SourceId = (sourceId ?? string.Empty)
					};
				}
				else
				{
					if (log.Entries == null)
					{
						log.Entries = new List<StandardRvtChangeCandidateEntry>();
					}
					LoadLog = log;
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				LoadLog = new StandardRvtChangeCandidateLog
				{
					SourceId = (sourceId ?? string.Empty)
				};
				ProjectData.ClearProjectError();
			}
		}
		return LoadLog;
	}

	private static void SaveLog(string workspaceRoot, StandardRvtChangeCandidateLog log)
	{
		if (log != null && !string.IsNullOrWhiteSpace(log.SourceId) && FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
		{
			try
			{
				string path = BuildLogPath(workspaceRoot, log.SourceId);
				Directory.CreateDirectory(Path.GetDirectoryName(path));
				File.WriteAllText(path, PlainJsonReportWriter.Serialize(log), Encoding.UTF8);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
	}

	private static string BuildLogPath(string workspaceRoot, string sourceId)
	{
		if (string.IsNullOrWhiteSpace(sourceId))
		{
			return string.Empty;
		}
		return Path.Combine(GetCandidateFolder(workspaceRoot), "standard-rvt-change-candidates-" + SafeFileName(sourceId) + ".json");
	}

	private static string GetCandidateFolder(string workspaceRoot)
	{
		return FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "StandardChangeCandidates");
	}

	private static StandardLibrarySnapshot LoadSnapshot(string snapshotPath)
	{
		if (string.IsNullOrWhiteSpace(snapshotPath) || !File.Exists(snapshotPath))
		{
			return null;
		}
		try
		{
			return DataContractJsonFileStore.Load<StandardLibrarySnapshot>(snapshotPath);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
			return null;
		}
	}

	private static string BuildLoadableFamilyKey(string categoryName, string familyName)
	{
		return Normalize(categoryName) + "|" + Normalize(familyName);
	}

	private static string BuildLoadableTypeKey(string categoryName, string familyName, string typeName)
	{
		return BuildLoadableFamilyKey(categoryName, familyName) + "|" + Normalize(typeName);
	}

	private static string BuildSystemTypeKey(string typeClassName, string categoryName, string typeName)
	{
		return Normalize(typeClassName) + "|" + Normalize(categoryName) + "|" + Normalize(typeName);
	}

	private static string ExtractLastKeyPart(string key)
	{
		if (string.IsNullOrWhiteSpace(key))
		{
			return string.Empty;
		}
		string[] parts = key.Split(new char[1] { '|' });
		if (parts.Length == 0)
		{
			return key;
		}
		return parts[parts.Length - 1];
	}

	private static string ResolveElementName(Element element)
	{
		string ResolveElementName;
		if (element == null)
		{
			ResolveElementName = string.Empty;
		}
		else
		{
			try
			{
				ResolveElementName = element.Name ?? string.Empty;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ResolveElementName = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return ResolveElementName;
	}

	private static string ResolveCategoryName(Element element)
	{
		string ResolveCategoryName;
		try
		{
			ResolveCategoryName = element?.Category?.Name ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveCategoryName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveCategoryName;
	}

	private static string ResolveCategoryId(Element element)
	{
		try
		{
			if (element != null && element.Category != null && (object)element.Category.Id != null)
			{
				return RevitElementIdCompat.CompatIntegerValue(element.Category.Id).ToString(CultureInfo.InvariantCulture);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Empty;
	}

	private static string NormalizePathForCompare(string value)
	{
		return FamilyBrowserPathIdentityService.GetComparableIdentity(value);
	}

	private static string SafeFileName(string value)
	{
		string text = value ?? string.Empty;
		if (text.Length == 0)
		{
			return "standard";
		}
		char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
		char[] chars = text.Select([SpecialName] (char ch) => (!Enumerable.Contains(invalidFileNameChars, ch)) ? ch : '_').ToArray();
		return new string(chars);
	}

	private static string Normalize(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().ToLowerInvariant();
	}
}
