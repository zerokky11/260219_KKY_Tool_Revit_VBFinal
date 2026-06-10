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

		public string SlotKey { get; set; }

		public string DisciplineKey { get; set; }

		public string DisciplineLabel { get; set; }

		public StandardRvtTarget()
		{
			SourceId = string.Empty;
			StandardRvtPath = string.Empty;
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

	private const int MaxLogEntriesPerSource = 500;

	private static readonly object SyncRoot = RuntimeHelpers.GetObjectValue(new object());

	private static string CachedWorkspaceRoot = string.Empty;

	private static DateTime CachedAtUtc = DateTime.MinValue;

	private static List<StandardRvtTarget> CachedTargets = new List<StandardRvtTarget>();

	private static readonly HashSet<string> AllowedSystemTypeNames = new HashSet<string>(new string[21]
	{
		"WallType", "FloorType", "RoofType", "CeilingType", "StairsType", "RailingType", "DuctType", "PipeType", "FlexDuctType", "FlexPipeType",
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
			if (entries.Count != 0)
			{
				AppendCandidates(workspaceRoot, target, entries.Values);
			}
		}
	}

	public static void RecordLoadableFamilyOperation(string workspaceRoot, Document doc, StandardLibraryRegistrationRecord registration, LoadableFamilySyncExecutionReport execution, string operationKind)
	{
		if (execution != null && execution.Items != null)
		{
			List<FamilyBrowserOperationLogEntry> entries = execution.Items.Where([SpecialName] (LoadableFamilySyncExecutionItem x) => x != null).Select([SpecialName] (LoadableFamilySyncExecutionItem x) =>
			{
				FamilyBrowserOperationLogEntry familyBrowserOperationLogEntry = new FamilyBrowserOperationLogEntry
				{
					EntryId = Guid.NewGuid().ToString("N"),
					RecordedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
					UserName = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity(),
					OperationKind = (operationKind ?? "LoadableFamilyOperation")
				};
				Document obj = doc;
				familyBrowserOperationLogEntry.DocumentTitle = ((obj != null) ? obj.Title : null) ?? execution.ProjectDocumentTitle;
				Document obj2 = doc;
				familyBrowserOperationLogEntry.DocumentPath = ((obj2 != null) ? obj2.PathName : null) ?? string.Empty;
				familyBrowserOperationLogEntry.SourceId = registration?.SourceId ?? string.Empty;
				familyBrowserOperationLogEntry.StandardDisplayName = registration?.DisplayName ?? execution.StandardDisplayName;
				familyBrowserOperationLogEntry.CandidateKind = "LoadableFamily";
				familyBrowserOperationLogEntry.CategoryName = x.CategoryName ?? string.Empty;
				familyBrowserOperationLogEntry.FamilyName = x.FamilyName ?? string.Empty;
				familyBrowserOperationLogEntry.PlannedAction = x.PlannedAction ?? string.Empty;
				familyBrowserOperationLogEntry.Outcome = x.Outcome ?? string.Empty;
				familyBrowserOperationLogEntry.Details = x.Details ?? string.Empty;
				return familyBrowserOperationLogEntry;
			}).ToList();
			AppendOperationEntries(workspaceRoot, entries);
		}
	}

	public static void RecordSystemTypeOperation(string workspaceRoot, Document doc, StandardLibraryRegistrationRecord registration, SystemTypeApplyExecutionReport execution, string operationKind)
	{
		if (execution != null && execution.Items != null)
		{
			List<FamilyBrowserOperationLogEntry> entries = execution.Items.Where([SpecialName] (SystemTypeApplyExecutionItem x) => x != null).Select([SpecialName] (SystemTypeApplyExecutionItem x) =>
			{
				FamilyBrowserOperationLogEntry familyBrowserOperationLogEntry = new FamilyBrowserOperationLogEntry
				{
					EntryId = Guid.NewGuid().ToString("N"),
					RecordedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
					UserName = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity(),
					OperationKind = (operationKind ?? "SystemTypeOperation")
				};
				Document obj = doc;
				familyBrowserOperationLogEntry.DocumentTitle = ((obj != null) ? obj.Title : null) ?? execution.ProjectDocumentTitle;
				Document obj2 = doc;
				familyBrowserOperationLogEntry.DocumentPath = ((obj2 != null) ? obj2.PathName : null) ?? string.Empty;
				familyBrowserOperationLogEntry.SourceId = registration?.SourceId ?? string.Empty;
				familyBrowserOperationLogEntry.StandardDisplayName = registration?.DisplayName ?? execution.StandardDisplayName;
				familyBrowserOperationLogEntry.CandidateKind = "SystemType";
				familyBrowserOperationLogEntry.CategoryName = x.CategoryName ?? string.Empty;
				familyBrowserOperationLogEntry.TypeName = x.SystemTypeName ?? string.Empty;
				familyBrowserOperationLogEntry.SystemFamilyKind = x.SystemFamilyKind ?? string.Empty;
				familyBrowserOperationLogEntry.PlannedAction = x.SyncAction ?? string.Empty;
				familyBrowserOperationLogEntry.Outcome = x.Outcome ?? string.Empty;
				familyBrowserOperationLogEntry.Details = x.Details ?? string.Empty;
				return familyBrowserOperationLogEntry;
			}).ToList();
			AppendOperationEntries(workspaceRoot, entries);
		}
	}

	public static List<StandardRvtChangeCandidateEntry> LoadRecent(string workspaceRoot, string sourceId, int limit = 50)
	{
		if (string.IsNullOrWhiteSpace(sourceId))
		{
			return new List<StandardRvtChangeCandidateEntry>();
		}
		return (LoadLog(workspaceRoot, sourceId)?.Entries ?? new List<StandardRvtChangeCandidateEntry>()).Where([SpecialName] (StandardRvtChangeCandidateEntry x) => x != null).OrderByDescending([SpecialName] (StandardRvtChangeCandidateEntry x) => x.RecordedAtUtc ?? string.Empty, StringComparer.Ordinal).Take(Math.Max(1, limit))
			.ToList();
	}

	public static string BuildRecentLoadableFamilyNameText(string workspaceRoot, string sourceId, int limit = 40)
	{
		List<string> names = (from x in LoadRecent(workspaceRoot, sourceId, Math.Max(checked(limit * 2), 50))
			where string.Equals(x.CandidateKind ?? string.Empty, "LoadableFamily", StringComparison.OrdinalIgnoreCase)
			select (x.FamilyName ?? string.Empty).Trim() into x
			where x.Length > 0
			select x).Distinct(StringComparer.OrdinalIgnoreCase).Take(Math.Max(1, limit)).ToList();
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
			select x).Distinct(StringComparer.OrdinalIgnoreCase).Take(limit).ToList();
		List<string> systemNames = (from x in recent
			where string.Equals(x.CandidateKind ?? string.Empty, "SystemType", StringComparison.OrdinalIgnoreCase)
			select (x.SystemFamilyKind ?? string.Empty).Trim() + " / " + (x.TypeName ?? string.Empty).Trim() into x
			where x.Trim().Length > 3
			select x).Distinct(StringComparer.OrdinalIgnoreCase).Take(limit).ToList();
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
		Family family = (Family)(object)((element is Family) ? element : null);
		FamilySymbol symbol = (FamilySymbol)(object)((element is FamilySymbol) ? element : null);
		FamilyInstance familyInstance = (FamilyInstance)(object)((element is FamilyInstance) ? element : null);
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
			string typeName = ((symbol == null) ? string.Empty : ResolveElementName((Element)(object)symbol));
			if (string.IsNullOrWhiteSpace(typeName) && familyInstance != null && familyInstance.Symbol != null)
			{
				typeName = ResolveElementName((Element)(object)familyInstance.Symbol);
			}
			string reason = ((element is FamilySymbol) ? (string.Equals(changeKind, "Added", StringComparison.OrdinalIgnoreCase) ? "TypeAdded" : "TypeChanged") : ((!(element is Family)) ? "FamilyReferenceChanged" : (string.Equals(changeKind, "Added", StringComparison.OrdinalIgnoreCase) ? "FamilyAdded" : "FamilyChanged")));
			StandardRvtChangeCandidateEntry entry = CreateBaseEntry(doc, target, "LoadableFamily", changeKind, reason);
			entry.CategoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family);
			entry.CategoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family);
			entry.FamilyName = ((Element)family).Name ?? string.Empty;
			entry.TypeName = typeName;
			entry.Details = ((object)element).GetType().Name;
			AddOrMerge(entries, entry);
		}
	}

	private static void AddSystemTypeCandidate(IDictionary<string, StandardRvtChangeCandidateEntry> entries, Document doc, StandardRvtTarget target, Element element, string changeKind)
	{
		ElementType elementType = (ElementType)(object)((element is ElementType) ? element : null);
		if (elementType != null && !(elementType is FamilySymbol))
		{
			string typeClassName = ((object)elementType).GetType().Name;
			if (AllowedSystemTypeNames.Contains(typeClassName))
			{
				StandardRvtChangeCandidateEntry entry = CreateBaseEntry(doc, target, "SystemType", changeKind, string.Equals(changeKind, "Added", StringComparison.OrdinalIgnoreCase) ? "SystemTypeAdded" : "SystemTypeChanged");
				entry.CategoryName = ResolveCategoryName((Element)(object)elementType);
				entry.CategoryId = ResolveCategoryId((Element)(object)elementType);
				entry.SystemFamilyKind = typeClassName;
				entry.TypeName = ResolveElementName((Element)(object)elementType);
				entry.Details = ((object)element).GetType().Name;
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
		standardRvtChangeCandidateEntry.DocumentTitle = ((doc != null) ? doc.Title : null) ?? string.Empty;
		standardRvtChangeCandidateEntry.DocumentPath = ((doc != null) ? doc.PathName : null) ?? string.Empty;
		standardRvtChangeCandidateEntry.SourceId = target?.SourceId ?? string.Empty;
		standardRvtChangeCandidateEntry.SlotKey = target?.SlotKey ?? string.Empty;
		standardRvtChangeCandidateEntry.DisciplineKey = target?.DisciplineKey ?? string.Empty;
		standardRvtChangeCandidateEntry.DisciplineLabel = target?.DisciplineLabel ?? string.Empty;
		standardRvtChangeCandidateEntry.CandidateKind = candidateKind;
		standardRvtChangeCandidateEntry.ChangeKind = changeKind;
		standardRvtChangeCandidateEntry.Reason = reason;
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
		string text = NormalizePathForCompare(doc.PathName);
		if (string.IsNullOrWhiteSpace(text))
		{
			return null;
		}
		return GetCachedTargets(workspaceRoot).FirstOrDefault([SpecialName] (StandardRvtTarget x) => string.Equals(NormalizePathForCompare(x.StandardRvtPath), text, StringComparison.OrdinalIgnoreCase));
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
		return (from g in result.Where([SpecialName] (StandardRvtTarget x) => x != null && !string.IsNullOrWhiteSpace(x.StandardRvtPath)).GroupBy([SpecialName] (StandardRvtTarget x) => NormalizePathForCompare(x.StandardRvtPath), StringComparer.OrdinalIgnoreCase)
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
				SlotKey = (slot.SlotKey ?? string.Empty),
				DisciplineKey = (slot.Discipline ?? string.Empty),
				DisciplineLabel = FamilyBrowserStandardPolicyStore.ResolveSlotDisplayName(slot, korean: false)
			});
		}
	}

	private static void AppendCandidates(string workspaceRoot, StandardRvtTarget target, IEnumerable<StandardRvtChangeCandidateEntry> candidates)
	{
		List<StandardRvtChangeCandidateEntry> list = (candidates ?? Enumerable.Empty<StandardRvtChangeCandidateEntry>()).Where([SpecialName] (StandardRvtChangeCandidateEntry x) => x != null).ToList();
		if (list.Count == 0 || target == null || string.IsNullOrWhiteSpace(target.SourceId))
		{
			return;
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
			using (List<StandardRvtChangeCandidateEntry>.Enumerator enumerator = list.GetEnumerator())
			{
				_Closure_0024__25_002D0 closure_0024__25_002D = default(_Closure_0024__25_002D0);
				while (enumerator.MoveNext())
				{
					closure_0024__25_002D = new _Closure_0024__25_002D0(closure_0024__25_002D);
					closure_0024__25_002D._0024VB_0024Local_entry = enumerator.Current;
					log.Entries.RemoveAll(closure_0024__25_002D._Lambda_0024__1);
					log.Entries.Add(closure_0024__25_002D._0024VB_0024Local_entry);
				}
			}
			log.Entries = log.Entries.Where([SpecialName] (StandardRvtChangeCandidateEntry x) => x != null).OrderByDescending([SpecialName] (StandardRvtChangeCandidateEntry x) => x.RecordedAtUtc ?? string.Empty, StringComparer.Ordinal).Take(500)
				.OrderBy([SpecialName] (StandardRvtChangeCandidateEntry x) => x.RecordedAtUtc ?? string.Empty, StringComparer.Ordinal)
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
	}

	private static void AppendOperationEntries(string workspaceRoot, IEnumerable<FamilyBrowserOperationLogEntry> entries)
	{
		List<FamilyBrowserOperationLogEntry> list = (entries ?? Enumerable.Empty<FamilyBrowserOperationLogEntry>()).Where([SpecialName] (FamilyBrowserOperationLogEntry x) => x != null).ToList();
		if (list.Count == 0 || !FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
		{
			return;
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
				log.Entries = log.Entries.OrderByDescending([SpecialName] (FamilyBrowserOperationLogEntry x) => x.RecordedAtUtc ?? string.Empty, StringComparer.Ordinal).Take(1000).OrderBy([SpecialName] (FamilyBrowserOperationLogEntry x) => x.RecordedAtUtc ?? string.Empty, StringComparer.Ordinal)
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
			object obj;
			if (element == null)
			{
				obj = null;
			}
			else
			{
				Category category = element.Category;
				obj = ((category != null) ? category.Name : null);
			}
			if (obj == null)
			{
				obj = string.Empty;
			}
			ResolveCategoryName = (string)obj;
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
			if (element != null && element.Category != null && element.Category.Id != null)
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
		string NormalizePathForCompare;
		if (string.IsNullOrWhiteSpace(value))
		{
			NormalizePathForCompare = string.Empty;
		}
		else
		{
			try
			{
				NormalizePathForCompare = Path.GetFullPath(Environment.ExpandEnvironmentVariables(value.Trim())).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				NormalizePathForCompare = value.Trim().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
				ProjectData.ClearProjectError();
			}
		}
		return NormalizePathForCompare;
	}

	private static string SafeFileName(string value)
	{
		string text = value ?? string.Empty;
		if (text.Length == 0)
		{
			return "standard";
		}
		char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
		char[] chars = text.Select([SpecialName] (char ch) => (!invalidFileNameChars.Contains(ch)) ? ch : '_').ToArray();
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
