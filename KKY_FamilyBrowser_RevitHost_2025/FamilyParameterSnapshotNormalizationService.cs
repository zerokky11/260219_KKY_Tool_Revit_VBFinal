using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

public sealed class FamilyParameterSnapshotNormalizationService
{
	private FamilyParameterSnapshotNormalizationService()
	{
	}

	public static List<StandardFamilyParameterSnapshotItem> DeduplicateDefinitions(IEnumerable<StandardFamilyParameterSnapshotItem> parameters)
	{
		return (from x in (parameters ?? Enumerable.Empty<StandardFamilyParameterSnapshotItem>()).Where([SpecialName] (StandardFamilyParameterSnapshotItem x) => x != null && !string.IsNullOrWhiteSpace(x.Name)).GroupBy([SpecialName] (StandardFamilyParameterSnapshotItem x) => BuildDefinitionIdentityKey(x), StringComparer.Ordinal)
			select NormalizeSelectedParameter(SelectPreferredParameterSnapshot(x)) into x
			where x != null
			select x).OrderBy([SpecialName] (StandardFamilyParameterSnapshotItem x) => Normalize(NormalizedRole(x)), StringComparer.Ordinal).ThenBy([SpecialName] (StandardFamilyParameterSnapshotItem x) => Normalize(x.Name), StringComparer.Ordinal).ThenBy([SpecialName] (StandardFamilyParameterSnapshotItem x) => Normalize(x.TypeName), StringComparer.Ordinal)
			.ToList();
	}

	public static List<StandardFamilyParameterSnapshotItem> DeduplicateDefinitionsAndTypeValues(IEnumerable<StandardFamilyParameterSnapshotItem> parameters)
	{
		return (from x in (parameters ?? Enumerable.Empty<StandardFamilyParameterSnapshotItem>()).Where([SpecialName] (StandardFamilyParameterSnapshotItem x) => x != null && !string.IsNullOrWhiteSpace(x.Name)).GroupBy([SpecialName] (StandardFamilyParameterSnapshotItem x) => BuildDefinitionAndTypeValueIdentityKey(x), StringComparer.Ordinal)
			select NormalizeSelectedParameter(SelectPreferredParameterSnapshot(x)) into x
			where x != null
			select x).OrderBy([SpecialName] (StandardFamilyParameterSnapshotItem x) => Normalize(NormalizedRole(x)), StringComparer.Ordinal).ThenBy([SpecialName] (StandardFamilyParameterSnapshotItem x) => Normalize(x.Name), StringComparer.Ordinal).ThenBy([SpecialName] (StandardFamilyParameterSnapshotItem x) => Normalize(x.TypeName), StringComparer.Ordinal)
			.ToList();
	}

	private static string BuildDefinitionAndTypeValueIdentityKey(StandardFamilyParameterSnapshotItem item)
	{
		string key = BuildDefinitionIdentityKey(item);
		if (item != null && string.Equals(NormalizedRole(item), "type", StringComparison.Ordinal))
		{
			key += "|typevalue|" + Normalize(item.TypeName);
		}
		return key;
	}

	public static string BuildDefinitionIdentityKey(StandardFamilyParameterSnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		string role = NormalizedRole(item);
		string name = Normalize(item.Name);
		string storage = Normalize(item.StorageType);
		string portableIdentity = Normalize(ResolvePortableParameterIdentity(item));
		if (portableIdentity.Length > 0)
		{
			return "def|" + role + "|id|" + portableIdentity + "|" + name + "|" + storage;
		}
		string kind = ((item.IsShared || !string.IsNullOrWhiteSpace(item.ExternalGuid)) ? "shared" : "local");
		return "def|" + role + "|" + kind + "|" + name + "|" + storage;
	}

	public static string BuildComparableDefinitionSignature(IEnumerable<StandardFamilyParameterSnapshotItem> parameters)
	{
		IOrderedEnumerable<string> parts = (from x in DeduplicateDefinitions(parameters)
			select NormalizedRole(x) + ":" + Normalize(x.Name) + ":" + Normalize(x.StorageType) + ":" + Normalize(x.Formula) + ":" + x.IsShared + ":" + Normalize(ResolvePortableParameterIdentity(x))).Distinct(StringComparer.Ordinal).OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal);
		return string.Join("|", parts);
	}

	public static string NormalizedRole(StandardFamilyParameterSnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		string scope = Normalize(item.Scope);
		if (item.IsInstance || string.Equals(scope, "instance", StringComparison.Ordinal))
		{
			return "instance";
		}
		return "type";
	}

	public static string ResolvePortableParameterIdentity(StandardFamilyParameterSnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		if (!string.IsNullOrWhiteSpace(item.ExternalGuid))
		{
			return "guid:" + item.ExternalGuid.Trim();
		}
		if (!string.IsNullOrWhiteSpace(item.ParameterId) && item.ParameterId.Trim().StartsWith("-", StringComparison.Ordinal))
		{
			return "builtin:" + item.ParameterId.Trim();
		}
		return string.Empty;
	}

	private static StandardFamilyParameterSnapshotItem SelectPreferredParameterSnapshot(IEnumerable<StandardFamilyParameterSnapshotItem> items)
	{
		return (from x in items ?? Enumerable.Empty<StandardFamilyParameterSnapshotItem>()
			where x != null
			orderby ParameterSnapshotPreferenceRank(x), !string.IsNullOrWhiteSpace(x.ExternalGuid) descending, x.IsShared descending, !string.IsNullOrWhiteSpace(x.Formula) descending, !string.IsNullOrWhiteSpace(x.ValuePreview) descending, !string.IsNullOrWhiteSpace(x.TypeName) descending
			select x).FirstOrDefault();
	}

	private static int ParameterSnapshotPreferenceRank(StandardFamilyParameterSnapshotItem item)
	{
		if (item == null)
		{
			return int.MaxValue;
		}
		if (string.Equals(NormalizedRole(item), "instance", StringComparison.Ordinal))
		{
			return 0;
		}
		string scope = Normalize(item.Scope);
		if (string.Equals(scope, "type", StringComparison.Ordinal))
		{
			return 1;
		}
		if (string.Equals(scope, "family", StringComparison.Ordinal))
		{
			return 2;
		}
		return 3;
	}

	private static StandardFamilyParameterSnapshotItem NormalizeSelectedParameter(StandardFamilyParameterSnapshotItem item)
	{
		if (item == null)
		{
			return null;
		}
		string role = NormalizedRole(item);
		return new StandardFamilyParameterSnapshotItem
		{
			Scope = (string.Equals(role, "instance", StringComparison.Ordinal) ? "Instance" : "Type"),
			TypeName = (item.TypeName ?? string.Empty),
			Name = (item.Name ?? string.Empty),
			StorageType = (item.StorageType ?? string.Empty),
			ValuePreview = (item.ValuePreview ?? string.Empty),
			Formula = (item.Formula ?? string.Empty),
			IsInstance = string.Equals(role, "instance", StringComparison.Ordinal),
			IsReadOnly = item.IsReadOnly,
			IsShared = (item.IsShared || !string.IsNullOrWhiteSpace(item.ExternalGuid)),
			ParameterId = (item.ParameterId ?? string.Empty),
			ExternalGuid = (item.ExternalGuid ?? string.Empty)
		};
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
