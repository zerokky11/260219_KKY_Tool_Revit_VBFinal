using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

public sealed class SystemTypeSyncService
{
	[CompilerGenerated]
	internal sealed class _Closure_0024__3_002D0
	{
		public SystemTypeSemanticSnapshot _0024VB_0024Local_sourceType;

		public _Closure_0024__3_002D0(_Closure_0024__3_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_sourceType = arg0._0024VB_0024Local_sourceType;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__6(SystemTypeSemanticSnapshot x)
		{
			return !string.Equals(Normalize(x.CategoryName), Normalize(_0024VB_0024Local_sourceType.CategoryName), StringComparison.Ordinal);
		}

		[SpecialName]
		internal bool _Lambda_0024__9(SystemTypeSemanticSnapshot x)
		{
			return string.Equals(Normalize(x.SystemFamilyKind), Normalize(_0024VB_0024Local_sourceType.SystemFamilyKind), StringComparison.Ordinal);
		}

		[SpecialName]
		internal bool _Lambda_0024__10(SystemTypeSemanticSnapshot x)
		{
			return string.Equals(Normalize(x.CategoryName), Normalize(_0024VB_0024Local_sourceType.CategoryName), StringComparison.Ordinal);
		}

		[SpecialName]
		internal bool _Lambda_0024__11(SystemTypeSemanticSnapshot x)
		{
			return IsDuplicateNameForSource(x.TypeName, _0024VB_0024Local_sourceType.TypeName);
		}

		[SpecialName]
		internal bool _Lambda_0024__12(SystemTypeSemanticSnapshot x)
		{
			return !string.Equals(Normalize(x.TypeName), Normalize(_0024VB_0024Local_sourceType.TypeName), StringComparison.Ordinal);
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__4_002D0
	{
		public string _0024VB_0024Local_key;

		public _Closure_0024__4_002D0(_Closure_0024__4_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_key = arg0._0024VB_0024Local_key;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__3(KeyValuePair<string, string> x)
		{
			return string.Equals(Normalize(x.Key), _0024VB_0024Local_key, StringComparison.Ordinal);
		}

		[SpecialName]
		internal bool _Lambda_0024__4(KeyValuePair<string, string> x)
		{
			return string.Equals(Normalize(x.Key), _0024VB_0024Local_key, StringComparison.Ordinal);
		}
	}

	private static readonly Regex DuplicateSuffixPattern = new Regex("^(.*?)(?:\\s+\\d+|\\(\\d+\\))$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

	private SystemTypeSyncService()
	{
	}

	public static SystemTypeSyncPlan BuildPlan(SystemTypeCatalogSnapshot sourceCatalog, SystemTypeCatalogSnapshot targetCatalog)
	{
		SystemTypeSyncPlan plan = new SystemTypeSyncPlan
		{
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
		};
		if (sourceCatalog == null)
		{
			return plan;
		}
		List<SystemTypeSemanticSnapshot> targetTypes = targetCatalog?.Types ?? new List<SystemTypeSemanticSnapshot>();
		Dictionary<string, SystemTypeSemanticSnapshot> targetExactMap = BuildFirstMap(targetTypes, [SpecialName] (SystemTypeSemanticSnapshot x) => BuildExactKey(x.SystemFamilyKind, x.CategoryName, x.TypeName), StringComparer.Ordinal);
		Dictionary<string, List<SystemTypeSemanticSnapshot>> targetNameGroups = targetTypes.GroupBy([SpecialName] (SystemTypeSemanticSnapshot x) => BuildNameKey(x.SystemFamilyKind, x.TypeName), StringComparer.Ordinal).ToDictionary([SpecialName] (IGrouping<string, SystemTypeSemanticSnapshot> x) => x.Key, [SpecialName] (IGrouping<string, SystemTypeSemanticSnapshot> x) => x.ToList(), StringComparer.Ordinal);
		using (IEnumerator<SystemTypeSemanticSnapshot> enumerator = sourceCatalog.Types.OrderBy([SpecialName] (SystemTypeSemanticSnapshot x) => Normalize(x.SystemFamilyKind), StringComparer.Ordinal).ThenBy([SpecialName] (SystemTypeSemanticSnapshot x) => Normalize(x.TypeName), StringComparer.Ordinal).GetEnumerator())
		{
			_Closure_0024__3_002D0 closure_0024__3_002D = default(_Closure_0024__3_002D0);
			while (enumerator.MoveNext())
			{
				closure_0024__3_002D = new _Closure_0024__3_002D0(closure_0024__3_002D);
				closure_0024__3_002D._0024VB_0024Local_sourceType = enumerator.Current;
				SystemTypeSyncPlanItem item = new SystemTypeSyncPlanItem
				{
					SystemFamilyKind = closure_0024__3_002D._0024VB_0024Local_sourceType.SystemFamilyKind,
					CategoryName = closure_0024__3_002D._0024VB_0024Local_sourceType.CategoryName,
					SourceTypeName = closure_0024__3_002D._0024VB_0024Local_sourceType.TypeName,
					SourceFingerprint = SystemTypeFingerprintService.Compute(closure_0024__3_002D._0024VB_0024Local_sourceType)
				};
				string exactKey = BuildExactKey(closure_0024__3_002D._0024VB_0024Local_sourceType.SystemFamilyKind, closure_0024__3_002D._0024VB_0024Local_sourceType.CategoryName, closure_0024__3_002D._0024VB_0024Local_sourceType.TypeName);
				string nameKey = BuildNameKey(closure_0024__3_002D._0024VB_0024Local_sourceType.SystemFamilyKind, closure_0024__3_002D._0024VB_0024Local_sourceType.TypeName);
				SystemTypeSemanticSnapshot exactTarget = null;
				targetExactMap.TryGetValue(exactKey, out exactTarget);
				List<SystemTypeSemanticSnapshot> categoryMismatchTargets = null;
				if (!targetNameGroups.TryGetValue(nameKey, out categoryMismatchTargets) || categoryMismatchTargets == null)
				{
					categoryMismatchTargets = new List<SystemTypeSemanticSnapshot>();
				}
				categoryMismatchTargets = categoryMismatchTargets.Where(closure_0024__3_002D._Lambda_0024__6).OrderBy([SpecialName] (SystemTypeSemanticSnapshot x) => Normalize(x.CategoryName), StringComparer.Ordinal).ThenBy([SpecialName] (SystemTypeSemanticSnapshot x) => Normalize(x.TypeName), StringComparer.Ordinal)
					.ToList();
				List<SystemTypeSemanticSnapshot> relatedDuplicates = targetTypes.Where(closure_0024__3_002D._Lambda_0024__9).Where(closure_0024__3_002D._Lambda_0024__10).Where(closure_0024__3_002D._Lambda_0024__11)
					.Where(closure_0024__3_002D._Lambda_0024__12)
					.ToList();
				item.RelatedDuplicateNames.AddRange(relatedDuplicates.Select([SpecialName] (SystemTypeSemanticSnapshot x) => x.TypeName).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase));
				if (exactTarget == null)
				{
					if (categoryMismatchTargets.Count > 0)
					{
						item.Action = "ManualReview";
						item.DestinationTypeName = string.Join(", ", categoryMismatchTargets.Select([SpecialName] (SystemTypeSemanticSnapshot x) => x.TypeName).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase));
						item.DiffSummary.Add("Category mismatch: standard category " + (closure_0024__3_002D._0024VB_0024Local_sourceType.CategoryName ?? string.Empty) + " / project category " + string.Join(", ", categoryMismatchTargets.Select([SpecialName] (SystemTypeSemanticSnapshot x) => x.CategoryName).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase)));
						item.Reason = "System type name exists in the target project, but the category differs from the registered standard. Review the category before sync.";
					}
					else
					{
						item.Action = ((item.RelatedDuplicateNames.Count > 0) ? "ManualReview" : "SkipMissingType");
						item.Reason = ((item.RelatedDuplicateNames.Count > 0) ? "Exact destination type is missing, but numbered duplicate names already exist. Promote one duplicate or clean them before sync." : "No destination type with the same system family kind and type name exists. Missing system types are left uncreated; only existing project system types are updated.");
					}
					plan.Items.Add(item);
				}
				else
				{
					item.DestinationTypeName = exactTarget.TypeName;
					item.DestinationFingerprint = SystemTypeFingerprintService.Compute(exactTarget);
					item.DiffSummary.AddRange(BuildDiffSummary(closure_0024__3_002D._0024VB_0024Local_sourceType, exactTarget));
					if (string.Equals(item.SourceFingerprint, item.DestinationFingerprint, StringComparison.OrdinalIgnoreCase))
					{
						item.Action = ((item.RelatedDuplicateNames.Count > 0) ? "ConsolidateDuplicateSuffixTypes" : "KeepDestination");
						item.Reason = ((item.RelatedDuplicateNames.Count > 0) ? "Canonical type already matches the source, but numbered duplicates still exist and should be merged or deleted." : "Destination type already matches the registered source type.");
						plan.Items.Add(item);
					}
					else
					{
						item.Action = "OverwriteDestination";
						item.Reason = ((item.RelatedDuplicateNames.Count > 0) ? "Destination type differs from the source and numbered duplicates also exist. Update the canonical type first, then clean up duplicates." : "Destination type differs from the registered source type and should be updated in place.");
						plan.Items.Add(item);
					}
				}
			}
		}
		return plan;
	}

	private static List<string> BuildDiffSummary(SystemTypeSemanticSnapshot sourceType, SystemTypeSemanticSnapshot destinationType)
	{
		List<string> diffs = new List<string>();
		if (sourceType == null || destinationType == null)
		{
			return diffs;
		}
		AppendValueDiff(diffs, "Classification code", sourceType.ClassificationCode, destinationType.ClassificationCode);
		AppendValueDiff(diffs, "Segment", sourceType.SegmentName, destinationType.SegmentName);
		AppendValueDiff(diffs, "Material", sourceType.MaterialName, destinationType.MaterialName);
		AppendValueDiff(diffs, "Shape", sourceType.Shape, destinationType.Shape);
		AppendValueDiff(diffs, "Routing preference", sourceType.RoutingPreferenceSignature, destinationType.RoutingPreferenceSignature);
		AppendValueDiff(diffs, "Compound structure", sourceType.CompoundStructureSignature, destinationType.CompoundStructureSignature);
		HashSet<string> sourceKeys = new HashSet<string>(sourceType.Parameters.Keys.Select(Normalize), StringComparer.Ordinal);
		HashSet<string> destinationKeys = new HashSet<string>(destinationType.Parameters.Keys.Select(Normalize), StringComparer.Ordinal);
		foreach (string key in sourceKeys.Except(destinationKeys, StringComparer.Ordinal).OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal))
		{
			diffs.Add("Type parameter missing in destination: " + key);
		}
		foreach (string key2 in destinationKeys.Except(sourceKeys, StringComparer.Ordinal).OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal))
		{
			diffs.Add("Extra destination type parameter: " + key2);
		}
		using (IEnumerator<string> enumerator3 = sourceKeys.Intersect(destinationKeys, StringComparer.Ordinal).OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal).GetEnumerator())
		{
			_Closure_0024__4_002D0 closure_0024__4_002D = default(_Closure_0024__4_002D0);
			while (enumerator3.MoveNext())
			{
				closure_0024__4_002D = new _Closure_0024__4_002D0(closure_0024__4_002D);
				closure_0024__4_002D._0024VB_0024Local_key = enumerator3.Current;
				string value = sourceType.Parameters.First(closure_0024__4_002D._Lambda_0024__3).Value;
				if (!string.Equals(b: Normalize(destinationType.Parameters.First(closure_0024__4_002D._Lambda_0024__4).Value), a: Normalize(value), comparisonType: StringComparison.Ordinal))
				{
					diffs.Add("Type parameter changed: " + closure_0024__4_002D._0024VB_0024Local_key);
				}
			}
		}
		return diffs;
	}

	private static Dictionary<string, T> BuildFirstMap<T>(IEnumerable<T> items, Func<T, string> keySelector, IEqualityComparer<string> comparer)
	{
		Dictionary<string, T> result = new Dictionary<string, T>(comparer);
		if (items == null)
		{
			return result;
		}
		foreach (T item in items)
		{
			string key = keySelector(item);
			if (!result.ContainsKey(key))
			{
				result.Add(key, item);
			}
		}
		return result;
	}

	private static void AppendValueDiff(ICollection<string> diffs, string label, string sourceValue, string destinationValue)
	{
		if (!string.Equals(Normalize(sourceValue), Normalize(destinationValue), StringComparison.Ordinal))
		{
			diffs.Add(label + " changed.");
		}
	}

	private static string BuildExactKey(string systemFamilyKind, string categoryName, string typeName)
	{
		return Normalize(systemFamilyKind) + "|" + Normalize(categoryName) + "|" + Normalize(typeName);
	}

	private static string BuildNameKey(string systemFamilyKind, string typeName)
	{
		return Normalize(systemFamilyKind) + "|" + Normalize(typeName);
	}

	private static bool IsDuplicateNameForSource(string candidateName, string sourceName)
	{
		if (string.IsNullOrWhiteSpace(candidateName) || string.IsNullOrWhiteSpace(sourceName))
		{
			return false;
		}
		if (string.Equals(Normalize(candidateName), Normalize(sourceName), StringComparison.Ordinal))
		{
			return false;
		}
		return string.Equals(Normalize(RemoveDuplicateSuffix(candidateName)), Normalize(sourceName), StringComparison.Ordinal);
	}

	private static string RemoveDuplicateSuffix(string value)
	{
		string input = (value ?? string.Empty).Trim();
		Match match = DuplicateSuffixPattern.Match(input);
		if (match.Success)
		{
			return match.Groups[1].Value.Trim();
		}
		return input;
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
