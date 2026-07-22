using System;
using System.Collections.Generic;
using System.Linq;

public static class NestedLoadableFamilyDifferencePropagationService
{
	public static void Apply(IList<LoadableFamilyComparisonItem> items)
	{
		if (items == null || items.Count == 0)
		{
			return;
		}

		List<LoadableFamilyComparisonItem> candidates = items.Where(x => x != null && !string.IsNullOrWhiteSpace(x.FamilyName)).ToList();
		Dictionary<LoadableFamilyComparisonItem, List<LoadableFamilyComparisonItem>> parentsByChild = BuildParentMap(candidates);
		Queue<LoadableFamilyComparisonItem> pending = new Queue<LoadableFamilyComparisonItem>();
		HashSet<LoadableFamilyComparisonItem> inQueue = new HashSet<LoadableFamilyComparisonItem>();

		foreach (LoadableFamilyComparisonItem child in parentsByChild.Keys)
		{
			if (!StatusRequiresNestedReview(child.Status, parentsByChild[child]))
			{
				continue;
			}
			PrepareNestedDifferenceChild(child);
			pending.Enqueue(child);
			inQueue.Add(child);
		}

		int propagationSteps = 0;
		int maximumPropagationSteps = Math.Max(32, candidates.Count * candidates.Count * 4);
		while (pending.Count > 0 && propagationSteps < maximumPropagationSteps)
		{
			propagationSteps++;
			LoadableFamilyComparisonItem child = pending.Dequeue();
			inQueue.Remove(child);
			List<LoadableFamilyComparisonItem> parents = null;
			if (!parentsByChild.TryGetValue(child, out parents) || parents == null)
			{
				continue;
			}
			foreach (LoadableFamilyComparisonItem parent in parents)
			{
				if (parent == null || ReferenceEquals(parent, child))
				{
					continue;
				}
				bool appended = AppendNestedDifferenceToParent(parent, child);
				if (appended && parent.IsNestedLoadableChild && StatusRequiresNestedReview(parent.Status, null) && inQueue.Add(parent))
				{
					parent.IsNestedLoadableDifference = true;
					pending.Enqueue(parent);
				}
			}
		}
	}

	private static Dictionary<LoadableFamilyComparisonItem, List<LoadableFamilyComparisonItem>> BuildParentMap(IList<LoadableFamilyComparisonItem> items)
	{
		Dictionary<LoadableFamilyComparisonItem, List<LoadableFamilyComparisonItem>> result = new Dictionary<LoadableFamilyComparisonItem, List<LoadableFamilyComparisonItem>>();
		foreach (LoadableFamilyComparisonItem parent in items)
		{
			List<StandardNestedLoadableFamilySnapshotItem> dependencies = new List<StandardNestedLoadableFamilySnapshotItem>();
			if (parent.NestedLoadableFamilies != null)
			{
				dependencies.AddRange(parent.NestedLoadableFamilies.Where(x => x != null));
			}
			if (parent.ProjectNestedLoadableFamilies != null)
			{
				dependencies.AddRange(parent.ProjectNestedLoadableFamilies.Where(x => x != null));
			}
			HashSet<LoadableFamilyComparisonItem> parentChildren = new HashSet<LoadableFamilyComparisonItem>();
			foreach (StandardNestedLoadableFamilySnapshotItem dependency in dependencies)
			{
				LoadableFamilyComparisonItem child = ResolveNestedComparisonItem(items, dependency);
				if (child == null || ReferenceEquals(child, parent) || !parentChildren.Add(child))
				{
					continue;
				}
				child.IsNestedLoadableChild = true;
				EnsureLists(child);
				AddDistinctText(child.NestedParentFamilyNames, parent.FamilyName);
				List<LoadableFamilyComparisonItem> parents = null;
				if (!result.TryGetValue(child, out parents))
				{
					parents = new List<LoadableFamilyComparisonItem>();
					result.Add(child, parents);
				}
				if (!parents.Contains(parent))
				{
					parents.Add(parent);
				}
			}
		}
		return result;
	}

	private static LoadableFamilyComparisonItem ResolveNestedComparisonItem(IList<LoadableFamilyComparisonItem> items, StandardNestedLoadableFamilySnapshotItem dependency)
	{
		if (items == null || dependency == null || string.IsNullOrWhiteSpace(dependency.FamilyName))
		{
			return null;
		}
		string familyToken = Normalize(dependency.FamilyName);
		string categoryToken = NormalizeCategory(dependency.CategoryName);
		List<LoadableFamilyComparisonItem> familyMatches = items.Where(x => x != null && string.Equals(Normalize(x.FamilyName), familyToken, StringComparison.Ordinal)).ToList();
		if (familyMatches.Count == 0)
		{
			return null;
		}
		if (!string.IsNullOrWhiteSpace(categoryToken))
		{
			List<LoadableFamilyComparisonItem> exact = familyMatches.Where(x => string.Equals(NormalizeCategory(x.CategoryName), categoryToken, StringComparison.Ordinal)).ToList();
			if (exact.Count == 1)
			{
				return exact[0];
			}
		}
		if (familyMatches.Count == 1)
		{
			return familyMatches[0];
		}
		List<LoadableFamilyComparisonItem> alreadyMarked = familyMatches.Where(x => x.IsNestedLoadableChild).ToList();
		return alreadyMarked.Count == 1 ? alreadyMarked[0] : null;
	}

	private static void PrepareNestedDifferenceChild(LoadableFamilyComparisonItem child)
	{
		EnsureLists(child);
		child.IsNestedLoadableChild = true;
		child.IsNestedLoadableDifference = true;
		string originalStatus = Normalize(child.Status);
		string familyLabel = BuildFamilyLabel(child);
		string parents = FormatNameList(child.NestedParentFamilyNames, 8);
		string parentContext = string.IsNullOrWhiteSpace(parents) ? string.Empty : " Parent family: " + parents + ".";
		if (string.Equals(originalStatus, "loadavailable", StringComparison.Ordinal))
		{
			child.Status = "NestedMissingFromParent";
			AddDistinctText(child.FingerprintDifferenceSummary, "Nested family missing from parent family: " + familyLabel + "." + parentContext);
			PrependDifferenceDetail(child, new LoadableFingerprintDifferenceDetailItem
			{
				Area = "nested family",
				DifferenceKind = "missing",
				StandardValue = familyLabel,
				ProjectValue = "-",
				Details = "The approved nested family is missing from the current parent family." + parentContext + " Update the parent family; this nested helper is not loaded independently."
			});
		}
		else if (string.Equals(originalStatus, "projectonly", StringComparison.Ordinal))
		{
			child.Status = "NestedExtraInParent";
			AddDistinctText(child.FingerprintDifferenceSummary, "Extra nested family in current parent family: " + familyLabel + "." + parentContext);
			PrependDifferenceDetail(child, new LoadableFingerprintDifferenceDetailItem
			{
				Area = "nested family",
				DifferenceKind = "project-only",
				StandardValue = "-",
				ProjectValue = familyLabel,
				Details = "The current parent family contains a nested family that is not present in the approved standard composition." + parentContext
			});
		}
		else if (child.FingerprintDifferenceSummary.Count == 0)
		{
			AddDistinctText(child.FingerprintDifferenceSummary, "Nested family differs from standard: " + familyLabel + ".");
		}

		if (!string.IsNullOrWhiteSpace(parents))
		{
			string instruction = string.Equals(Normalize(child.Status), "nestedmissingfromparent", StringComparison.Ordinal)
				? " Update the parent family to restore the approved nested composition; nested helper rows are not loaded independently."
				: " Review/update the parent family; nested helper rows are not loaded independently.";
			child.Notes = "Nested family used by parent families: " + parents + "." + instruction;
		}
	}

	private static bool AppendNestedDifferenceToParent(LoadableFamilyComparisonItem parent, LoadableFamilyComparisonItem child)
	{
		EnsureLists(parent);
		string childLabel = BuildFamilyLabel(child);
		bool changed = !parent.IsNestedLoadableDifference;
		parent.IsNestedLoadableDifference = true;
		if (!parent.NestedDifferenceFamilyNames.Any(x => string.Equals(Normalize(x), Normalize(childLabel), StringComparison.Ordinal)))
		{
			parent.NestedDifferenceFamilyNames.Add(childLabel);
			changed = true;
		}
		string reason = BuildDifferenceReason(child);
		changed = UpsertNestedDifferenceSummary(parent.FingerprintDifferenceSummary, childLabel, reason) || changed;

		changed = UpsertNestedDependencyDetail(parent.FingerprintDifferenceDetails, childLabel, child, reason) || changed;
		foreach (LoadableFingerprintDifferenceDetailItem detail in child.FingerprintDifferenceDetails.Where(x => x != null).Take(4))
		{
			changed = AddDistinctNestedDetail(parent.FingerprintDifferenceDetails, new LoadableFingerprintDifferenceDetailItem
			{
				Area = string.IsNullOrWhiteSpace(detail.Area) ? "family detail" : detail.Area,
				DifferenceKind = string.IsNullOrWhiteSpace(detail.DifferenceKind) ? "dependency-different" : detail.DifferenceKind,
				StandardValue = detail.StandardValue ?? string.Empty,
				ProjectValue = detail.ProjectValue ?? string.Empty,
				Details = childLabel + " / " + (string.IsNullOrWhiteSpace(detail.Area) ? "detail" : detail.Area) + ": " + (detail.Details ?? string.Empty)
			}) || changed;
		}

		string parentStatus = Normalize(parent.Status);
		string nestedReviewNote = "Nested dependency requires review: " + childLabel + " - " + reason;
		if (string.Equals(parentStatus, "loadedlatest", StringComparison.Ordinal) || string.Equals(parentStatus, "stampnormalizationneeded", StringComparison.Ordinal))
		{
			parent.Status = string.Equals(Normalize(child.Status), "manualreview", StringComparison.Ordinal) ? "ManualReview" : "DifferentFromStandard";
			parent.Notes = nestedReviewNote;
			changed = true;
		}
		else
		{
			string updatedNotes = AppendNote(parent.Notes, nestedReviewNote);
			if (!string.Equals(parent.Notes ?? string.Empty, updatedNotes, StringComparison.Ordinal))
			{
				parent.Notes = updatedNotes;
				changed = true;
			}
		}
		return changed;
	}

	private static bool StatusRequiresNestedReview(string status, IEnumerable<LoadableFamilyComparisonItem> parents)
	{
		string token = Normalize(status);
		if (string.IsNullOrWhiteSpace(token))
		{
			return false;
		}
		if (string.Equals(token, "loadedlatest", StringComparison.Ordinal) || string.Equals(token, "stampnormalizationneeded", StringComparison.Ordinal))
		{
			return false;
		}
		if (string.Equals(token, "loadavailable", StringComparison.Ordinal) && parents != null)
		{
			return parents.Any(ParentExistsInCurrentProject);
		}
		return true;
	}

	private static bool ParentExistsInCurrentProject(LoadableFamilyComparisonItem parent)
	{
		return parent != null && !string.Equals(Normalize(parent.Status), "loadavailable", StringComparison.Ordinal);
	}

	private static bool UpsertNestedDifferenceSummary(IList<string> summaries, string childLabel, string reason)
	{
		if (summaries == null)
		{
			return false;
		}
		string prefix = "Nested family differs: " + childLabel + " - ";
		string value = prefix + reason;
		for (int i = 0; i < summaries.Count; i++)
		{
			if (!(summaries[i] ?? string.Empty).StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
			{
				continue;
			}
			if (string.Equals(summaries[i], value, StringComparison.Ordinal))
			{
				return false;
			}
			summaries[i] = value;
			return true;
		}
		summaries.Add(value);
		return true;
	}

	private static bool UpsertNestedDependencyDetail(IList<LoadableFingerprintDifferenceDetailItem> details, string childLabel, LoadableFamilyComparisonItem child, string reason)
	{
		if (details == null)
		{
			return false;
		}
		string standardValue = BuildFingerprintValue(childLabel, child.StandardContentFingerprint, child.StandardFingerprint);
		string projectValue = BuildFingerprintValue(childLabel, child.ProjectContentFingerprint, child.ProjectFingerprint);
		LoadableFingerprintDifferenceDetailItem existing = details.FirstOrDefault(x => x != null &&
			string.Equals(Normalize(x.Area), "nested family", StringComparison.Ordinal) &&
			string.Equals(Normalize(x.DifferenceKind), "dependency-different", StringComparison.Ordinal) &&
			string.Equals(x.StandardValue ?? string.Empty, standardValue, StringComparison.OrdinalIgnoreCase) &&
			string.Equals(x.ProjectValue ?? string.Empty, projectValue, StringComparison.OrdinalIgnoreCase));
		if (existing != null)
		{
			if (string.Equals(existing.Details ?? string.Empty, reason, StringComparison.Ordinal))
			{
				return false;
			}
			existing.Details = reason;
			return true;
		}
		if (details.Count >= 24)
		{
			return false;
		}
		details.Insert(0, new LoadableFingerprintDifferenceDetailItem
		{
			Area = "nested family",
			DifferenceKind = "dependency-different",
			StandardValue = standardValue,
			ProjectValue = projectValue,
			Details = reason
		});
		return true;
	}

	private static bool AddDistinctNestedDetail(IList<LoadableFingerprintDifferenceDetailItem> details, LoadableFingerprintDifferenceDetailItem candidate)
	{
		if (details == null || candidate == null || details.Count >= 24)
		{
			return false;
		}
		bool exists = details.Any(x => x != null &&
			string.Equals(x.Area ?? string.Empty, candidate.Area ?? string.Empty, StringComparison.OrdinalIgnoreCase) &&
			string.Equals(x.DifferenceKind ?? string.Empty, candidate.DifferenceKind ?? string.Empty, StringComparison.OrdinalIgnoreCase) &&
			string.Equals(x.StandardValue ?? string.Empty, candidate.StandardValue ?? string.Empty, StringComparison.OrdinalIgnoreCase) &&
			string.Equals(x.ProjectValue ?? string.Empty, candidate.ProjectValue ?? string.Empty, StringComparison.OrdinalIgnoreCase) &&
			string.Equals(x.Details ?? string.Empty, candidate.Details ?? string.Empty, StringComparison.OrdinalIgnoreCase));
		if (exists)
		{
			return false;
		}
		details.Add(candidate);
		return true;
	}

	private static void EnsureLists(LoadableFamilyComparisonItem item)
	{
		if (item.FingerprintDifferenceSummary == null)
		{
			item.FingerprintDifferenceSummary = new List<string>();
		}
		if (item.FingerprintDifferenceDetails == null)
		{
			item.FingerprintDifferenceDetails = new List<LoadableFingerprintDifferenceDetailItem>();
		}
		if (item.NestedParentFamilyNames == null)
		{
			item.NestedParentFamilyNames = new List<string>();
		}
		if (item.NestedDifferenceFamilyNames == null)
		{
			item.NestedDifferenceFamilyNames = new List<string>();
		}
	}

	private static void PrependDifferenceDetail(LoadableFamilyComparisonItem item, LoadableFingerprintDifferenceDetailItem detail)
	{
		EnsureLists(item);
		item.FingerprintDifferenceDetails.Insert(0, detail);
	}

	private static void AddDistinctText(IList<string> values, string value)
	{
		if (values == null || string.IsNullOrWhiteSpace(value))
		{
			return;
		}
		string trimmed = value.Trim();
		if (!values.Any(x => string.Equals((x ?? string.Empty).Trim(), trimmed, StringComparison.OrdinalIgnoreCase)))
		{
			values.Add(trimmed);
		}
	}

	private static string BuildDifferenceReason(LoadableFamilyComparisonItem item)
	{
		if (item != null && item.FingerprintDifferenceSummary != null)
		{
			List<string> reasons = item.FingerprintDifferenceSummary.Where(x => !string.IsNullOrWhiteSpace(x)).Take(3).ToList();
			if (reasons.Count > 0)
			{
				return LimitText(string.Join(" | ", reasons), 560);
			}
		}
		return "Nested family status: " + ((item == null || string.IsNullOrWhiteSpace(item.Status)) ? "review required" : item.Status) + ".";
	}

	private static string BuildFamilyLabel(LoadableFamilyComparisonItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		string family = item.FamilyName ?? string.Empty;
		string category = NormalizeCategory(item.CategoryName);
		return string.IsNullOrWhiteSpace(category) ? family : family + " (" + (item.CategoryName ?? string.Empty).Trim() + ")";
	}

	private static string BuildFingerprintValue(string label, string contentFingerprint, string fallbackFingerprint)
	{
		string fingerprint = string.IsNullOrWhiteSpace(contentFingerprint) ? fallbackFingerprint : contentFingerprint;
		return "Nested Family: " + label + " | Fingerprint: " + (string.IsNullOrWhiteSpace(fingerprint) ? "-" : ShortHash(fingerprint));
	}

	private static string ShortHash(string value)
	{
		string text = (value ?? string.Empty).Trim();
		return text.Length <= 12 ? text : text.Substring(0, 12);
	}

	private static string FormatNameList(IEnumerable<string> values, int limit)
	{
		List<string> names = (values ?? Enumerable.Empty<string>()).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
		if (names.Count == 0)
		{
			return string.Empty;
		}
		string result = string.Join(", ", names.Take(Math.Max(1, limit)));
		return names.Count > limit ? result + " ..." : result;
	}

	private static string AppendNote(string current, string addition)
	{
		if (string.IsNullOrWhiteSpace(addition))
		{
			return current ?? string.Empty;
		}
		if (string.IsNullOrWhiteSpace(current))
		{
			return addition.Trim();
		}
		if ((current ?? string.Empty).IndexOf(addition.Trim(), StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return current;
		}
		return current.Trim() + " " + addition.Trim();
	}

	private static string LimitText(string value, int maxLength)
	{
		string text = (value ?? string.Empty).Trim();
		return text.Length <= maxLength ? text : text.Substring(0, maxLength - 3).TrimEnd() + "...";
	}

	private static string NormalizeCategory(string value)
	{
		string token = Normalize(value);
		return string.Equals(token, "-", StringComparison.Ordinal) ? string.Empty : token;
	}

	private static string Normalize(string value)
	{
		return (value ?? string.Empty).Trim().ToLowerInvariant();
	}
}
