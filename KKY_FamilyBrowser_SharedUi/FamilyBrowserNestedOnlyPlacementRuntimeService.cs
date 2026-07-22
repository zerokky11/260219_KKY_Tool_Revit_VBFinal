using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

public enum FamilyBrowserNestedOnlyPlacementMatchState
{
	None,
	PendingVerification,
	VerificationUnavailable,
	ExactMatch
}

public sealed class FamilyBrowserNestedOnlyPlacementMatchResult
{
	public FamilyBrowserNestedOnlyPlacementMatchState State { get; set; }

	public FamilyBrowserNestedOnlyPlacementEntry Entry { get; set; }

	public string ProjectContentFingerprint { get; set; }

	public string Detail { get; set; }

	public FamilyBrowserNestedOnlyPlacementMatchResult()
	{
		ProjectContentFingerprint = string.Empty;
		Detail = string.Empty;
	}
}

public static class FamilyBrowserNestedOnlyPlacementRuntimeService
{
	private sealed class ProjectFamilyFingerprintEvidence
	{
		public int FamilyElementId { get; set; }

		public string FamilyUniqueId { get; set; }

		public string FamilyName { get; set; }

		public string CategoryName { get; set; }

		public string CategoryId { get; set; }

		public bool IsShared { get; set; }

		public string ContentFingerprint { get; set; }

		public string FastFingerprint { get; set; }

		public string FailureReason { get; set; }

		public DateTime CapturedUtc { get; set; }

		public ProjectFamilyFingerprintEvidence()
		{
			FamilyUniqueId = string.Empty;
			FamilyName = string.Empty;
			CategoryName = string.Empty;
			CategoryId = string.Empty;
			ContentFingerprint = string.Empty;
			FastFingerprint = string.Empty;
			FailureReason = string.Empty;
			CapturedUtc = DateTime.MinValue;
		}
	}

	private sealed class DocumentFingerprintState
	{
		public int DocumentToken { get; set; }

		public string DocumentKey { get; set; }

		public string CatalogStamp { get; set; }

		public List<FamilyBrowserNestedOnlyPlacementEntry> Candidates { get; set; }

		public Dictionary<string, ProjectFamilyFingerprintEvidence> EvidenceByUniqueId { get; set; }

		public Dictionary<int, string> UniqueIdByElementId { get; set; }

		public Queue<int> PendingFamilyIds { get; set; }

		public HashSet<int> PendingFamilyIdSet { get; set; }

		public int VerifiedCount { get; set; }

		public int FailedCount { get; set; }

		public DocumentFingerprintState()
		{
			DocumentKey = string.Empty;
			CatalogStamp = string.Empty;
			Candidates = new List<FamilyBrowserNestedOnlyPlacementEntry>();
			EvidenceByUniqueId = new Dictionary<string, ProjectFamilyFingerprintEvidence>(StringComparer.Ordinal);
			UniqueIdByElementId = new Dictionary<int, string>();
			PendingFamilyIds = new Queue<int>();
			PendingFamilyIdSet = new HashSet<int>();
		}
	}

	private sealed class CandidateSet
	{
		public string Stamp { get; set; }

		public List<FamilyBrowserNestedOnlyPlacementEntry> Entries { get; set; }

		public int CatalogCount { get; set; }

		public CandidateSet()
		{
			Stamp = string.Empty;
			Entries = new List<FamilyBrowserNestedOnlyPlacementEntry>();
		}
	}

	private sealed class CandidateCacheEntry
	{
		public CandidateSet CandidateSet { get; set; }

		public DateTime CachedUtc { get; set; }
	}

	private static readonly object SyncRoot = new object();

	private static readonly Dictionary<int, DocumentFingerprintState> States = new Dictionary<int, DocumentFingerprintState>();

	private static readonly Dictionary<string, CandidateCacheEntry> CandidateSetCache = new Dictionary<string, CandidateCacheEntry>(StringComparer.OrdinalIgnoreCase);

	private static readonly TimeSpan CandidateCacheDuration = TimeSpan.FromSeconds(15.0);

	private static readonly TimeSpan FailedEvidenceRetryDelay = TimeSpan.FromSeconds(30.0);

	public static void ResetAll()
	{
		lock (SyncRoot)
		{
			States.Clear();
			CandidateSetCache.Clear();
		}
	}

	public static void InvalidateCatalogs()
	{
		lock (SyncRoot)
		{
			CandidateSetCache.Clear();
			foreach (DocumentFingerprintState state in States.Values)
			{
				state.CatalogStamp = string.Empty;
			}
		}
	}

	public static void ScheduleRefresh(Document document, FamilyBrowserStandardPolicy policy)
	{
		ScheduleRefresh(document, policy, string.Empty);
	}

	public static void ScheduleRefresh(Document document, FamilyBrowserStandardPolicy policy, string discipline)
	{
		if (!CanInspectDocument(document))
		{
			return;
		}
		CandidateSet candidateSet = ResolveCandidateSet(policy, discipline);
		DocumentFingerprintState state = GetOrCreateState(document, candidateSet);
		List<Family> families;
		try
		{
			families = new FilteredElementCollector(document).OfClass(typeof(Family)).Cast<Family>().ToList();
		}
		catch
		{
			return;
		}
		HashSet<string> liveUniqueIds = new HashSet<string>(StringComparer.Ordinal);
		foreach (Family family in families)
		{
			if (!HasCandidate(state.Candidates, family))
			{
				continue;
			}
			string uniqueId = SafeFamilyUniqueId(family);
			int familyId = SafeElementId(family?.Id);
			if (string.IsNullOrWhiteSpace(uniqueId) || familyId <= 0)
			{
				continue;
			}
			liveUniqueIds.Add(uniqueId);
			state.UniqueIdByElementId[familyId] = uniqueId;
			ProjectFamilyFingerprintEvidence evidence;
			if (!state.EvidenceByUniqueId.TryGetValue(uniqueId, out evidence) || ShouldRetryFailedEvidence(evidence))
			{
				QueueFamily(state, familyId);
			}
		}
		foreach (string staleUniqueId in state.EvidenceByUniqueId.Keys.Where(delegate(string value)
		{
			return !liveUniqueIds.Contains(value);
		}).ToList())
		{
			state.EvidenceByUniqueId.Remove(staleUniqueId);
		}
	}

	public static bool HasPending(Document document)
	{
		DocumentFingerprintState state = TryGetState(document);
		return state != null && state.PendingFamilyIds.Count > 0;
	}

	public static int SeedFromProjectSnapshot(Document document, FamilyBrowserStandardPolicy policy, string discipline, ProjectContentSnapshot snapshot)
	{
		if (!CanInspectDocument(document) || snapshot == null || snapshot.LoadableFamilies == null)
		{
			return 0;
		}
		ScheduleRefresh(document, policy, discipline);
		DocumentFingerprintState state = TryGetState(document);
		if (state == null)
		{
			return 0;
		}
		int seeded = 0;
		foreach (ProjectLoadableFamilySnapshotItem item in snapshot.LoadableFamilies)
		{
			if (item == null || string.IsNullOrWhiteSpace(item.UniqueId) || string.IsNullOrWhiteSpace(item.ContentFingerprint))
			{
				continue;
			}
			Family family = null;
			try
			{
				family = document.GetElement(item.UniqueId) as Family;
			}
			catch
			{
			}
			if (family == null || !HasCandidate(state.Candidates, family))
			{
				continue;
			}
			int familyId = SafeElementId(family.Id);
			string uniqueId = SafeFamilyUniqueId(family);
			if (familyId <= 0 || string.IsNullOrWhiteSpace(uniqueId))
			{
				continue;
			}
			bool replacing = state.EvidenceByUniqueId.ContainsKey(uniqueId);
			state.EvidenceByUniqueId[uniqueId] = new ProjectFamilyFingerprintEvidence
			{
				FamilyElementId = familyId,
				FamilyUniqueId = uniqueId,
				FamilyName = SafeFamilyName(family),
				CategoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family),
				CategoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family),
				IsShared = item.IsShared,
				ContentFingerprint = item.ContentFingerprint ?? string.Empty,
				FailureReason = string.Empty,
				CapturedUtc = DateTime.UtcNow
			};
			state.UniqueIdByElementId[familyId] = uniqueId;
			RemovePendingFamily(state, familyId);
			if (!replacing)
			{
				state.VerifiedCount++;
			}
			seeded++;
		}
		return seeded;
	}

	public static bool ProcessNextPending(UIApplication uiApplication, Document document, FamilyBrowserStandardPolicy policy)
	{
		return ProcessNextPending(uiApplication, document, policy, string.Empty);
	}

	public static bool ProcessNextPending(UIApplication uiApplication, Document document, FamilyBrowserStandardPolicy policy, string discipline)
	{
		if (!CanInspectDocument(document))
		{
			return false;
		}
		CandidateSet candidateSet = ResolveCandidateSet(policy, discipline);
		DocumentFingerprintState state = GetOrCreateState(document, candidateSet);
		if (state.PendingFamilyIds.Count == 0 && state.EvidenceByUniqueId.Count == 0)
		{
			ScheduleRefresh(document, policy, discipline);
			state = TryGetState(document);
		}
		if (state == null || state.PendingFamilyIds.Count == 0)
		{
			return false;
		}
		if (SafeIsModifiable(document))
		{
			return true;
		}
		int familyId = state.PendingFamilyIds.Dequeue();
		state.PendingFamilyIdSet.Remove(familyId);
		Family family = null;
		try
		{
			family = document.GetElement(new ElementId(familyId)) as Family;
		}
		catch
		{
		}
		if (family == null || !HasCandidate(state.Candidates, family))
		{
			return state.PendingFamilyIds.Count > 0;
		}
		string familyName = SafeFamilyName(family);
		string categoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family);
		string categoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family);
		string uniqueId = SafeFamilyUniqueId(family);
		bool isShared = SafeIsShared(family);
		ProjectFamilyFingerprintEvidence evidence = new ProjectFamilyFingerprintEvidence
		{
			FamilyElementId = familyId,
			FamilyUniqueId = uniqueId,
			FamilyName = familyName,
			CategoryName = categoryName,
			CategoryId = categoryId,
			IsShared = isShared,
			CapturedUtc = DateTime.UtcNow
		};
		try
		{
			evidence.FastFingerprint = LoadableFamilyContentSignatureService.Build(document, family, includeDeepContent: false) ?? string.Empty;
			using (FamilyThumbnailConstraintDialogGuard dialogGuard = new FamilyThumbnailConstraintDialogGuard(uiApplication))
			{
				dialogGuard.SetCurrentFamily(categoryName, familyName);
				LoadableFamilyContentSignatureResult result = LoadableFamilyContentSignatureService.BuildResult(document, family, includeDeepContent: true);
				evidence.ContentFingerprint = result?.Fingerprint ?? string.Empty;
				evidence.FailureReason = result?.ErrorMessage ?? string.Empty;
			}
		}
		catch (Exception ex)
		{
			evidence.ContentFingerprint = string.Empty;
			evidence.FailureReason = ex.GetType().Name + ": " + ex.Message;
		}
		if (string.IsNullOrWhiteSpace(evidence.ContentFingerprint) && string.IsNullOrWhiteSpace(evidence.FailureReason))
		{
			evidence.FailureReason = "Precise project family fingerprint was empty.";
		}
		if (!string.IsNullOrWhiteSpace(uniqueId))
		{
			state.EvidenceByUniqueId[uniqueId] = evidence;
			state.UniqueIdByElementId[familyId] = uniqueId;
		}
		if (string.IsNullOrWhiteSpace(evidence.ContentFingerprint))
		{
			state.FailedCount++;
			TryWriteVerificationFailureLog(document, evidence);
		}
		else
		{
			state.VerifiedCount++;
		}
		return state.PendingFamilyIds.Count > 0;
	}

	public static void InvalidateChangedFamilies(Document document, ICollection<ElementId> addedIds, ICollection<ElementId> modifiedIds, ICollection<ElementId> deletedIds)
	{
		DocumentFingerprintState state = TryGetState(document);
		if (state == null)
		{
			return;
		}
		InvalidateDeletedFamilies(state, deletedIds);
		InvalidateLiveFamilies(document, state, addedIds, forceFamilyInvalidation: true);
		InvalidateLiveFamilies(document, state, modifiedIds, forceFamilyInvalidation: true);
	}

	public static FamilyBrowserNestedOnlyPlacementMatchResult EvaluatePlacement(Document document, Family family, FamilyBrowserStandardPolicy policy)
	{
		return EvaluatePlacement(document, family, policy, string.Empty);
	}

	public static FamilyBrowserNestedOnlyPlacementMatchResult EvaluatePlacement(Document document, Family family, FamilyBrowserStandardPolicy policy, string discipline)
	{
		FamilyBrowserNestedOnlyPlacementMatchResult noMatch = new FamilyBrowserNestedOnlyPlacementMatchResult
		{
			State = FamilyBrowserNestedOnlyPlacementMatchState.None
		};
		if (!CanInspectDocument(document) || family == null)
		{
			return noMatch;
		}
		CandidateSet candidateSet = ResolveCandidateSet(policy, discipline);
		DocumentFingerprintState state = GetOrCreateState(document, candidateSet);
		if (state.PendingFamilyIds.Count == 0 && state.EvidenceByUniqueId.Count == 0)
		{
			ScheduleRefresh(document, policy, discipline);
			state = TryGetState(document);
		}
		if (state == null)
		{
			return noMatch;
		}
		List<FamilyBrowserNestedOnlyPlacementEntry> candidates = FindCandidates(state.Candidates, family);
		if (candidates.Count == 0)
		{
			return noMatch;
		}
		string uniqueId = SafeFamilyUniqueId(family);
		int familyId = SafeElementId(family.Id);
		ProjectFamilyFingerprintEvidence evidence;
		if (string.IsNullOrWhiteSpace(uniqueId) || !state.EvidenceByUniqueId.TryGetValue(uniqueId, out evidence))
		{
			QueueFamily(state, familyId);
			return BuildPendingResult(candidates[0], "Project family fingerprint has not been verified yet.");
		}
		string currentFastFingerprint = string.Empty;
		try
		{
			currentFastFingerprint = LoadableFamilyContentSignatureService.Build(document, family, includeDeepContent: false) ?? string.Empty;
		}
		catch
		{
		}
		if (!string.IsNullOrWhiteSpace(evidence.FastFingerprint) && !string.IsNullOrWhiteSpace(currentFastFingerprint) && !string.Equals(evidence.FastFingerprint, currentFastFingerprint, StringComparison.OrdinalIgnoreCase))
		{
			state.EvidenceByUniqueId.Remove(uniqueId);
			QueueFamily(state, familyId);
			return BuildPendingResult(candidates[0], "Project family changed after fingerprint verification.");
		}
		if (string.IsNullOrWhiteSpace(evidence.ContentFingerprint))
		{
			if (ShouldRetryFailedEvidence(evidence))
			{
				QueueFamily(state, familyId);
			}
			return new FamilyBrowserNestedOnlyPlacementMatchResult
			{
				State = FamilyBrowserNestedOnlyPlacementMatchState.VerificationUnavailable,
				Entry = candidates[0],
				Detail = string.IsNullOrWhiteSpace(evidence.FailureReason) ? "Project family fingerprint verification failed." : evidence.FailureReason
			};
		}
		foreach (FamilyBrowserNestedOnlyPlacementEntry candidate in candidates)
		{
			if (FamilyBrowserNestedOnlyPlacementFingerprintPolicy.IsExactMatch(candidate, evidence.ContentFingerprint, evidence.IsShared))
			{
				return new FamilyBrowserNestedOnlyPlacementMatchResult
				{
					State = FamilyBrowserNestedOnlyPlacementMatchState.ExactMatch,
					Entry = candidate,
					ProjectContentFingerprint = evidence.ContentFingerprint,
					Detail = "Name, category, shared state, and precise content fingerprint match the registered standard."
				};
			}
		}
		return noMatch;
	}

	public static string BuildDiagnostic(Document document)
	{
		DocumentFingerprintState state = TryGetState(document);
		if (state == null)
		{
			return "nestedOnlyFingerprintState=not-scheduled";
		}
		return "nestedOnlyCandidates=" + state.Candidates.Count.ToString(CultureInfo.InvariantCulture) +
			";nestedOnlyVerified=" + state.VerifiedCount.ToString(CultureInfo.InvariantCulture) +
			";nestedOnlyVerificationFailed=" + state.FailedCount.ToString(CultureInfo.InvariantCulture) +
			";nestedOnlyVerificationPending=" + state.PendingFamilyIds.Count.ToString(CultureInfo.InvariantCulture);
	}

	private static FamilyBrowserNestedOnlyPlacementMatchResult BuildPendingResult(FamilyBrowserNestedOnlyPlacementEntry entry, string detail)
	{
		return new FamilyBrowserNestedOnlyPlacementMatchResult
		{
			State = FamilyBrowserNestedOnlyPlacementMatchState.PendingVerification,
			Entry = entry,
			Detail = detail ?? string.Empty
		};
	}

	private static void RemovePendingFamily(DocumentFingerprintState state, int familyId)
	{
		if (state == null || familyId <= 0 || !state.PendingFamilyIdSet.Remove(familyId))
		{
			return;
		}
		state.PendingFamilyIds = new Queue<int>(state.PendingFamilyIds.Where(delegate(int value)
		{
			return value != familyId;
		}));
	}

	private static DocumentFingerprintState GetOrCreateState(Document document, CandidateSet candidateSet)
	{
		int documentToken = RuntimeHelpers.GetHashCode(document);
		string documentKey = BuildDocumentKey(document);
		lock (SyncRoot)
		{
			DocumentFingerprintState state;
			if (!States.TryGetValue(documentToken, out state) || state == null || !string.Equals(state.DocumentKey, documentKey, StringComparison.OrdinalIgnoreCase) || !string.Equals(state.CatalogStamp, candidateSet.Stamp, StringComparison.Ordinal))
			{
				state = new DocumentFingerprintState
				{
					DocumentToken = documentToken,
					DocumentKey = documentKey,
					CatalogStamp = candidateSet.Stamp,
					Candidates = candidateSet.Entries.ToList()
				};
				States[documentToken] = state;
			}
			if (States.Count > 16)
			{
				foreach (int staleToken in States.Keys.Where(delegate(int value)
				{
					return value != documentToken;
				}).Take(States.Count - 16).ToList())
				{
					States.Remove(staleToken);
				}
			}
			return state;
		}
	}

	private static DocumentFingerprintState TryGetState(Document document)
	{
		if (document == null)
		{
			return null;
		}
		int token = RuntimeHelpers.GetHashCode(document);
		lock (SyncRoot)
		{
			DocumentFingerprintState state;
			return States.TryGetValue(token, out state) ? state : null;
		}
	}

	private static CandidateSet ResolveCandidateSet(FamilyBrowserStandardPolicy policy, string discipline)
	{
		List<FamilyBrowserStandardLibrarySlot> slots = ResolveSlots(policy, discipline).Where(delegate(FamilyBrowserStandardLibrarySlot slot) { return slot != null; }).ToList();
		string cacheKey = BuildCandidateCacheKey(discipline, slots);
		lock (SyncRoot)
		{
			CandidateCacheEntry cached;
			if (CandidateSetCache.TryGetValue(cacheKey, out cached) && cached != null && cached.CandidateSet != null && DateTime.UtcNow - cached.CachedUtc < CandidateCacheDuration)
			{
				return cached.CandidateSet;
			}
		}
		CandidateSet result = new CandidateSet();
		List<string> stampParts = new List<string>();
		Dictionary<string, FamilyBrowserNestedOnlyPlacementEntry> entries = new Dictionary<string, FamilyBrowserNestedOnlyPlacementEntry>(StringComparer.Ordinal);
		foreach (FamilyBrowserStandardLibrarySlot slot in slots)
		{
			string snapshotPath = ResolveSnapshotPath(slot);
			if (string.IsNullOrWhiteSpace(snapshotPath))
			{
				continue;
			}
			FamilyBrowserNestedOnlyPlacementCatalog catalog = FamilyBrowserNestedOnlyPlacementCatalogStore.TryLoadForSnapshot(snapshotPath);
			if (catalog == null || !catalog.IsComplete)
			{
				stampParts.Add(snapshotPath + "|unavailable");
				continue;
			}
			result.CatalogCount++;
			stampParts.Add(snapshotPath + "|" + (catalog.SourceSnapshotLastWriteUtc ?? string.Empty) + "|" + catalog.SourceSnapshotLength.ToString(CultureInfo.InvariantCulture));
			foreach (FamilyBrowserNestedOnlyPlacementEntry entry in catalog.Entries ?? new List<FamilyBrowserNestedOnlyPlacementEntry>())
			{
				if (entry == null || !entry.IsShared || string.IsNullOrWhiteSpace(entry.ContentFingerprint))
				{
					continue;
				}
				string key = BuildCandidateKey(entry);
				FamilyBrowserNestedOnlyPlacementEntry existing;
				if (entries.TryGetValue(key, out existing))
				{
					existing.ParentFamilyNames = (existing.ParentFamilyNames ?? new List<string>()).Concat(entry.ParentFamilyNames ?? new List<string>()).Where(delegate(string value)
					{
						return !string.IsNullOrWhiteSpace(value);
					}).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(delegate(string value)
					{
						return value;
					}, StringComparer.OrdinalIgnoreCase).ToList();
				}
				else
				{
					entries.Add(key, entry);
				}
			}
		}
		result.Entries = entries.Values.OrderBy(delegate(FamilyBrowserNestedOnlyPlacementEntry item)
		{
			return Normalize(item.CategoryName) + "|" + Normalize(item.FamilyName) + "|" + Normalize(item.ContentFingerprint);
		}, StringComparer.Ordinal).ToList();
		result.Stamp = string.Join(";", stampParts.OrderBy(delegate(string value)
		{
			return value;
		}, StringComparer.OrdinalIgnoreCase)) + "|entries=" + result.Entries.Count.ToString(CultureInfo.InvariantCulture);
		lock (SyncRoot)
		{
			CandidateSetCache[cacheKey] = new CandidateCacheEntry
			{
				CandidateSet = result,
				CachedUtc = DateTime.UtcNow
			};
		}
		return result;
	}

	private static IEnumerable<FamilyBrowserStandardLibrarySlot> ResolveSlots(FamilyBrowserStandardPolicy policy, string discipline)
	{
		if (policy == null)
		{
			return Enumerable.Empty<FamilyBrowserStandardLibrarySlot>();
		}
		if (string.Equals(policy.Mode, "Integrated", StringComparison.OrdinalIgnoreCase))
		{
			return policy.IntegratedLibrary != null && policy.IntegratedLibrary.Enabled ? new[] { policy.IntegratedLibrary } : Enumerable.Empty<FamilyBrowserStandardLibrarySlot>();
		}
		if (!string.IsNullOrWhiteSpace(discipline))
		{
			FamilyBrowserStandardLibrarySlot assigned = FamilyBrowserFileGuardDisciplineService.ResolveSlot(policy, discipline);
			return assigned == null
				? Enumerable.Empty<FamilyBrowserStandardLibrarySlot>()
				: new List<FamilyBrowserStandardLibrarySlot> { assigned };
		}
		return FamilyBrowserStandardPolicyStore.GetDisciplineSlots(policy);
	}

	private static string BuildCandidateCacheKey(string discipline, IEnumerable<FamilyBrowserStandardLibrarySlot> slots)
	{
		List<string> parts = new List<string>
		{
			FamilyBrowserPolicyKey.Normalize(discipline)
		};
		foreach (FamilyBrowserStandardLibrarySlot slot in slots ?? Enumerable.Empty<FamilyBrowserStandardLibrarySlot>())
		{
			parts.Add(FamilyBrowserPolicyKey.Normalize(slot == null ? string.Empty : slot.SlotKey));
			parts.Add(FamilyBrowserPolicyKey.Normalize(slot == null ? string.Empty : slot.Discipline));
			parts.Add(slot == null ? string.Empty : (slot.RegistrationPath ?? string.Empty));
			parts.Add(slot == null ? string.Empty : (slot.SnapshotPath ?? string.Empty));
		}
		return string.Join("|", parts.ToArray());
	}

	private static string ResolveSnapshotPath(FamilyBrowserStandardLibrarySlot slot)
	{
		if (slot == null)
		{
			return string.Empty;
		}
		try
		{
			string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
			StandardLibraryRegistrationRecord registration = null;
			string registrationPath = FamilyBrowserStandardPolicyStore.ResolveSlotRegistrationPath(workspaceRoot, slot);
			if (!string.IsNullOrWhiteSpace(registrationPath) && File.Exists(registrationPath))
			{
				registration = DataContractJsonFileStore.Load<StandardLibraryRegistrationRecord>(registrationPath);
			}
			string snapshotPath = FamilyBrowserStandardPolicyStore.ResolveSlotSnapshotPath(workspaceRoot, slot, registration);
			if (string.IsNullOrWhiteSpace(snapshotPath) && !string.IsNullOrWhiteSpace(slot.SnapshotPath) && File.Exists(slot.SnapshotPath))
			{
				snapshotPath = slot.SnapshotPath;
			}
			return snapshotPath ?? string.Empty;
		}
		catch
		{
			return string.Empty;
		}
	}

	private static void InvalidateDeletedFamilies(DocumentFingerprintState state, ICollection<ElementId> ids)
	{
		if (state == null || ids == null)
		{
			return;
		}
		foreach (ElementId id in ids)
		{
			int elementId = SafeElementId(id);
			string uniqueId;
			if (elementId > 0 && state.UniqueIdByElementId.TryGetValue(elementId, out uniqueId))
			{
				state.UniqueIdByElementId.Remove(elementId);
				state.EvidenceByUniqueId.Remove(uniqueId);
			}
		}
	}

	private static void InvalidateLiveFamilies(Document document, DocumentFingerprintState state, ICollection<ElementId> ids, bool forceFamilyInvalidation)
	{
		if (document == null || state == null || ids == null)
		{
			return;
		}
		foreach (ElementId id in ids)
		{
			Element element = null;
			try
			{
				element = document.GetElement(id);
			}
			catch
			{
			}
			Family family = element as Family;
			bool isFamilyElement = family != null;
			if (family == null && element is FamilySymbol symbol)
			{
				family = symbol.Family;
			}
			if (family == null || !HasCandidate(state.Candidates, family))
			{
				continue;
			}
			string uniqueId = SafeFamilyUniqueId(family);
			int familyId = SafeElementId(family.Id);
			ProjectFamilyFingerprintEvidence evidence;
			bool invalidate = isFamilyElement && forceFamilyInvalidation;
			if (!invalidate && state.EvidenceByUniqueId.TryGetValue(uniqueId, out evidence) && !string.IsNullOrWhiteSpace(evidence.FastFingerprint))
			{
				try
				{
					string currentFast = LoadableFamilyContentSignatureService.Build(document, family, includeDeepContent: false);
					invalidate = string.IsNullOrWhiteSpace(currentFast) || !string.Equals(currentFast, evidence.FastFingerprint, StringComparison.OrdinalIgnoreCase);
				}
				catch
				{
					invalidate = true;
				}
			}
			else if (!state.EvidenceByUniqueId.ContainsKey(uniqueId))
			{
				invalidate = true;
			}
			if (invalidate)
			{
				state.EvidenceByUniqueId.Remove(uniqueId);
				QueueFamily(state, familyId);
			}
		}
	}

	private static void QueueFamily(DocumentFingerprintState state, int familyId)
	{
		if (state == null || familyId <= 0 || !state.PendingFamilyIdSet.Add(familyId))
		{
			return;
		}
		state.PendingFamilyIds.Enqueue(familyId);
	}

	private static bool ShouldRetryFailedEvidence(ProjectFamilyFingerprintEvidence evidence)
	{
		return evidence != null && string.IsNullOrWhiteSpace(evidence.ContentFingerprint) && DateTime.UtcNow - evidence.CapturedUtc >= FailedEvidenceRetryDelay;
	}

	private static bool HasCandidate(IEnumerable<FamilyBrowserNestedOnlyPlacementEntry> candidates, Family family)
	{
		return FindCandidates(candidates, family).Count > 0;
	}

	private static List<FamilyBrowserNestedOnlyPlacementEntry> FindCandidates(IEnumerable<FamilyBrowserNestedOnlyPlacementEntry> candidates, Family family)
	{
		if (family == null)
		{
			return new List<FamilyBrowserNestedOnlyPlacementEntry>();
		}
		string familyName = SafeFamilyName(family);
		string categoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family);
		string categoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family);
		bool isShared = SafeIsShared(family);
		return (candidates ?? Enumerable.Empty<FamilyBrowserNestedOnlyPlacementEntry>()).Where(delegate(FamilyBrowserNestedOnlyPlacementEntry entry)
		{
			return EntryIdentityMatches(entry, categoryId, categoryName, familyName, isShared);
		}).ToList();
	}

	private static bool EntryIdentityMatches(FamilyBrowserNestedOnlyPlacementEntry entry, string categoryId, string categoryName, string familyName, bool isShared)
	{
		if (entry == null || !entry.IsShared || !isShared || !string.Equals(Normalize(entry.FamilyName), Normalize(familyName), StringComparison.Ordinal))
		{
			return false;
		}
		string entryCategoryId = Normalize(entry.CategoryId);
		string projectCategoryId = Normalize(categoryId);
		if (!string.IsNullOrWhiteSpace(entryCategoryId) && !string.IsNullOrWhiteSpace(projectCategoryId))
		{
			return string.Equals(entryCategoryId, projectCategoryId, StringComparison.Ordinal);
		}
		string entryCategoryName = Normalize(entry.CategoryName);
		string projectCategoryName = Normalize(categoryName);
		return string.IsNullOrWhiteSpace(entryCategoryName) || string.IsNullOrWhiteSpace(projectCategoryName) || string.Equals(entryCategoryName, projectCategoryName, StringComparison.Ordinal);
	}

	private static string BuildCandidateKey(FamilyBrowserNestedOnlyPlacementEntry entry)
	{
		return Normalize(entry.CategoryId) + "|" + Normalize(entry.CategoryName) + "|" + Normalize(entry.FamilyName) + "|" + Normalize(entry.ContentFingerprint) + "|" + entry.IsShared.ToString();
	}

	private static bool CanInspectDocument(Document document)
	{
		try
		{
			return document != null && document.IsValidObject && !document.IsFamilyDocument;
		}
		catch
		{
			return false;
		}
	}

	private static bool SafeIsModifiable(Document document)
	{
		try
		{
			return document != null && document.IsModifiable;
		}
		catch
		{
			return true;
		}
	}

	private static bool SafeIsShared(Family family)
	{
		try
		{
			Parameter parameter = ((Element)family)?.get_Parameter(BuiltInParameter.FAMILY_SHARED);
			return parameter != null && parameter.StorageType == StorageType.Integer && parameter.AsInteger() != 0;
		}
		catch
		{
			return false;
		}
	}

	private static string SafeFamilyName(Family family)
	{
		try
		{
			return family?.Name ?? string.Empty;
		}
		catch
		{
			return string.Empty;
		}
	}

	private static string SafeFamilyUniqueId(Family family)
	{
		try
		{
			return family?.UniqueId ?? string.Empty;
		}
		catch
		{
			return string.Empty;
		}
	}

	private static int SafeElementId(ElementId id)
	{
		try
		{
			return id == null ? 0 : RevitElementIdCompat.CompatIntegerValue(id);
		}
		catch
		{
			return 0;
		}
	}

	private static string BuildDocumentKey(Document document)
	{
		try
		{
			string identity = ProjectSnapshotStore.ResolveProjectIdentityPath(document);
			return Normalize(string.IsNullOrWhiteSpace(identity) ? document?.Title : identity);
		}
		catch
		{
			return Normalize(document?.Title);
		}
	}

	private static void TryWriteVerificationFailureLog(Document document, ProjectFamilyFingerprintEvidence evidence)
	{
		try
		{
			string detail = "Document=" + (document?.Title ?? string.Empty) + Environment.NewLine +
				"Category=" + (evidence?.CategoryName ?? string.Empty) + Environment.NewLine +
				"Family=" + (evidence?.FamilyName ?? string.Empty) + Environment.NewLine +
				"Reason=" + (evidence?.FailureReason ?? string.Empty);
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Nested-only project family fingerprint verification failed", new InvalidOperationException(evidence?.FailureReason ?? "Fingerprint verification failed."), detail);
		}
		catch
		{
		}
	}

	private static string Normalize(string value)
	{
		return (value ?? string.Empty).Trim().ToUpperInvariant();
	}
}
