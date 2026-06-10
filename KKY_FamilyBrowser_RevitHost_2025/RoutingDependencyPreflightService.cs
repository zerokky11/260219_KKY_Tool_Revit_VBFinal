using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

public sealed class RoutingDependencyPreflightService
{
	[CompilerGenerated]
	internal sealed class _Closure_0024__3_002D0
	{
		public RoutingDependencySnapshot _0024VB_0024Local_dependency;

		public _Closure_0024__3_002D0(_Closure_0024__3_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_dependency = arg0._0024VB_0024Local_dependency;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__12(RoutingFamilyTypeSnapshot x)
		{
			return string.Equals(Normalize(x.TypeName), Normalize(_0024VB_0024Local_dependency.TypeName), StringComparison.Ordinal);
		}

		[SpecialName]
		internal bool _Lambda_0024__13(RoutingFamilyTypeSnapshot x)
		{
			return !string.Equals(Normalize(x.TypeName), Normalize(_0024VB_0024Local_dependency.TypeName), StringComparison.Ordinal);
		}

		[SpecialName]
		internal bool _Lambda_0024__14(RoutingFamilyTypeSnapshot x)
		{
			return IsDuplicateNameForSource(x.TypeName, _0024VB_0024Local_dependency.TypeName);
		}
	}

	private static readonly Regex DuplicateSuffixPattern = new Regex("^(.*?)(?:\\s+\\d+|\\(\\d+\\))$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

	private RoutingDependencyPreflightService()
	{
	}

	public static RoutingDependencyPreflightPlan BuildPlan(SystemTypeCatalogSnapshot sourceCatalog, RoutingFamilyCatalogSnapshot targetCatalog)
	{
		RoutingDependencyPreflightPlan plan = new RoutingDependencyPreflightPlan
		{
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
		};
		if (sourceCatalog == null)
		{
			return plan;
		}
		List<RoutingFamilyCatalogEntry> source = targetCatalog?.Families ?? new List<RoutingFamilyCatalogEntry>();
		Dictionary<string, RoutingFamilyCatalogEntry> byLibraryId = source.Where([SpecialName] (RoutingFamilyCatalogEntry x) => !string.IsNullOrWhiteSpace(x.LibraryFamilyId)).GroupBy<RoutingFamilyCatalogEntry, string>([SpecialName] (RoutingFamilyCatalogEntry x) => Normalize(x.LibraryFamilyId), StringComparer.Ordinal).ToDictionary<IGrouping<string, RoutingFamilyCatalogEntry>, string, RoutingFamilyCatalogEntry>([SpecialName] (IGrouping<string, RoutingFamilyCatalogEntry> x) => x.Key, [SpecialName] (IGrouping<string, RoutingFamilyCatalogEntry> x) => x.First(), StringComparer.Ordinal);
		Dictionary<string, List<RoutingFamilyCatalogEntry>> byFamilyName = source.GroupBy<RoutingFamilyCatalogEntry, string>([SpecialName] (RoutingFamilyCatalogEntry x) => Normalize(x.FamilyName), StringComparer.Ordinal).ToDictionary<IGrouping<string, RoutingFamilyCatalogEntry>, string, List<RoutingFamilyCatalogEntry>>([SpecialName] (IGrouping<string, RoutingFamilyCatalogEntry> x) => x.Key, [SpecialName] (IGrouping<string, RoutingFamilyCatalogEntry> x) => x.ToList(), StringComparer.Ordinal);
		_Closure_0024__3_002D0 closure_0024__3_002D = default(_Closure_0024__3_002D0);
		foreach (SystemTypeSemanticSnapshot sourceType in sourceCatalog.Types.OrderBy<SystemTypeSemanticSnapshot, string>([SpecialName] (SystemTypeSemanticSnapshot x) => Normalize(x.SystemFamilyKind), StringComparer.Ordinal).ThenBy<SystemTypeSemanticSnapshot, string>([SpecialName] (SystemTypeSemanticSnapshot x) => Normalize(x.TypeName), StringComparer.Ordinal))
		{
			using IEnumerator<RoutingDependencySnapshot> enumerator2 = sourceType.RoutingDependencies.OrderBy<RoutingDependencySnapshot, string>([SpecialName] (RoutingDependencySnapshot x) => Normalize(x.DependencyRole), StringComparer.Ordinal).ThenBy<RoutingDependencySnapshot, string>([SpecialName] (RoutingDependencySnapshot x) => Normalize(x.FamilyName), StringComparer.Ordinal).ThenBy<RoutingDependencySnapshot, string>([SpecialName] (RoutingDependencySnapshot x) => Normalize(x.TypeName), StringComparer.Ordinal)
				.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				closure_0024__3_002D = new _Closure_0024__3_002D0(closure_0024__3_002D);
				closure_0024__3_002D._0024VB_0024Local_dependency = enumerator2.Current;
				RoutingDependencyPreflightItem item = new RoutingDependencyPreflightItem
				{
					SystemFamilyKind = sourceType.SystemFamilyKind,
					SystemTypeName = sourceType.TypeName,
					DependencyRole = closure_0024__3_002D._0024VB_0024Local_dependency.DependencyRole,
					SourceLibraryFamilyId = closure_0024__3_002D._0024VB_0024Local_dependency.LibraryFamilyId,
					SourceFamilyName = closure_0024__3_002D._0024VB_0024Local_dependency.FamilyName,
					SourceTypeName = closure_0024__3_002D._0024VB_0024Local_dependency.TypeName,
					SourceFamilyFingerprint = closure_0024__3_002D._0024VB_0024Local_dependency.FamilyFingerprint,
					SourceTypeFingerprint = closure_0024__3_002D._0024VB_0024Local_dependency.TypeFingerprint
				};
				RoutingFamilyCatalogEntry targetFamily = null;
				if (!string.IsNullOrWhiteSpace(closure_0024__3_002D._0024VB_0024Local_dependency.LibraryFamilyId))
				{
					byLibraryId.TryGetValue(Normalize(closure_0024__3_002D._0024VB_0024Local_dependency.LibraryFamilyId), out targetFamily);
				}
				if (targetFamily == null)
				{
					List<RoutingFamilyCatalogEntry> nameMatches = null;
					if (byFamilyName.TryGetValue(Normalize(closure_0024__3_002D._0024VB_0024Local_dependency.FamilyName), out nameMatches))
					{
						item.TargetFamilyName = nameMatches[0].FamilyName;
						item.Action = "ManualReviewNameOnlyMatch";
						item.Reason = "A family with the same name exists, but the canonical family identity does not match.";
						item.RelatedTypeNames.AddRange(CollectAllTypeNames(nameMatches));
						plan.Items.Add(item);
					}
					else
					{
						item.Action = "LoadMissingDependencyFamily";
						item.Reason = "No target family with the required canonical identity is loaded.";
						plan.Items.Add(item);
					}
					continue;
				}
				item.TargetLibraryFamilyId = targetFamily.LibraryFamilyId;
				item.TargetFamilyName = targetFamily.FamilyName;
				item.TargetFamilyFingerprint = targetFamily.FamilyFingerprint;
				RoutingFamilyTypeSnapshot targetType = targetFamily.Types.FirstOrDefault(closure_0024__3_002D._Lambda_0024__12);
				List<string> duplicateTypeNames = (from x in targetFamily.Types.Where(closure_0024__3_002D._Lambda_0024__13).Where(closure_0024__3_002D._Lambda_0024__14)
					select x.TypeName).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
				item.RelatedTypeNames.AddRange(duplicateTypeNames);
				if (targetType == null)
				{
					item.Action = ((duplicateTypeNames.Count > 0) ? "PromoteOrRenameDependencyType" : "ReloadFamilyOverwrite");
					item.Reason = ((duplicateTypeNames.Count > 0) ? "The canonical dependency type is missing, but duplicate-suffix types exist in the loaded family." : "The family is loaded, but the required dependency type is missing.");
					plan.Items.Add(item);
					continue;
				}
				item.TargetTypeName = targetType.TypeName;
				item.TargetTypeFingerprint = targetType.TypeFingerprint;
				bool familyMatches = FingerprintEquals(closure_0024__3_002D._0024VB_0024Local_dependency.FamilyFingerprint, targetFamily.FamilyFingerprint);
				bool typeMatches = FingerprintEquals(closure_0024__3_002D._0024VB_0024Local_dependency.TypeFingerprint, targetType.TypeFingerprint);
				if (familyMatches && typeMatches && duplicateTypeNames.Count == 0)
				{
					item.Action = "ReuseLoadedDependency";
					item.Reason = "The loaded dependency family and type already match the canonical source.";
					plan.Items.Add(item);
				}
				else if (familyMatches && typeMatches && duplicateTypeNames.Count > 0)
				{
					item.Action = "ReuseAndCleanupDuplicateTypes";
					item.Reason = "The canonical dependency matches, but duplicate-suffix types should be remapped and cleaned.";
					plan.Items.Add(item);
				}
				else
				{
					item.Action = "ReloadFamilyOverwrite";
					item.Reason = BuildReloadReason(familyMatches, typeMatches, duplicateTypeNames.Count > 0);
					plan.Items.Add(item);
				}
			}
		}
		return plan;
	}

	private static string BuildReloadReason(bool familyMatches, bool typeMatches, bool hasDuplicates)
	{
		if (!familyMatches && !typeMatches)
		{
			return hasDuplicates ? "The loaded dependency family and type both differ from the canonical source, and duplicate-suffix types also exist." : "The loaded dependency family and type both differ from the canonical source.";
		}
		if (!familyMatches)
		{
			return hasDuplicates ? "The loaded dependency family differs from the canonical source, and duplicate-suffix types also exist." : "The loaded dependency family differs from the canonical source.";
		}
		if (!typeMatches)
		{
			return hasDuplicates ? "The loaded dependency type differs from the canonical source, and duplicate-suffix types also exist." : "The loaded dependency type differs from the canonical source.";
		}
		return "The dependency should be reloaded and reconciled.";
	}

	private static IEnumerable<string> CollectAllTypeNames(IEnumerable<RoutingFamilyCatalogEntry> entries)
	{
		return entries.SelectMany([SpecialName] (RoutingFamilyCatalogEntry x) => x.Types.Select([SpecialName] (RoutingFamilyTypeSnapshot t) => t.TypeName)).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase);
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
		string input = candidateName.Trim();
		Match match = DuplicateSuffixPattern.Match(input);
		if (!match.Success)
		{
			return false;
		}
		return string.Equals(Normalize(match.Groups[1].Value), Normalize(sourceName), StringComparison.Ordinal);
	}

	private static bool FingerprintEquals(string left, string right)
	{
		return string.Equals(Normalize(left), Normalize(right), StringComparison.Ordinal);
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
