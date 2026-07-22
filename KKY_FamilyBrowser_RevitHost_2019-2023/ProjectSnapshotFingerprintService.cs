using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;

public sealed class ProjectSnapshotFingerprintService
{
	private ProjectSnapshotFingerprintService()
	{
	}

	public static string BuildLoadableFingerprint(StandardLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return BuildLoadableFingerprintCore(item.CategoryName, item.CategoryGroup, item.FamilyName, item.TypeCount, item.TypeNames, item.IsShared, item.ContentFingerprint, item.Parameters);
	}

	public static string BuildLoadableFingerprint(ProjectLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return BuildLoadableFingerprintCore(item.CategoryName, item.CategoryGroup, item.FamilyName, item.TypeCount, item.TypeNames, item.IsShared, item.ContentFingerprint, item.Parameters);
	}

	public static string BuildSystemFingerprint(StandardSystemTypeSnapshotItem item, bool includeDetailedComponents = true)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return BuildSystemFingerprintCore(item.TypeClassName, item.CategoryName, item.TypeName, item.SupportsRoutingDependencies, item.ClassificationCode, item.SegmentName, item.MaterialName, item.Shape, item.RoutingPreferenceSignature, item.CompoundStructureSignature, includeDetailedComponents, item.DetailedComponentsCaptured, item.DetailedComponentSignature, item.DetailedComponents);
	}

	public static string BuildSystemFingerprint(ProjectSystemTypeSnapshotItem item, bool includeDetailedComponents = true)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return BuildSystemFingerprintCore(item.TypeClassName, item.CategoryName, item.TypeName, item.SupportsRoutingDependencies, item.ClassificationCode, item.SegmentName, item.MaterialName, item.Shape, item.RoutingPreferenceSignature, item.CompoundStructureSignature, includeDetailedComponents, item.DetailedComponentsCaptured, item.DetailedComponentSignature, item.DetailedComponents);
	}

	private static string HashString(string value)
	{
		using SHA256 sha = SHA256.Create();
		byte[] bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
		return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant();
	}

	private static string BuildLoadableFingerprintCore(string categoryName, string categoryGroup, string familyName, int typeCount, IEnumerable<string> typeNames, bool isShared, string contentFingerprint, IEnumerable<StandardFamilyParameterSnapshotItem> parameters)
	{
		string normalizedContentFingerprint = Normalize(contentFingerprint);
		List<string> parts = new List<string>
		{
			"loadable-v3",
			Normalize(categoryName),
			Normalize(categoryGroup),
			Normalize(familyName),
			typeCount.ToString(CultureInfo.InvariantCulture),
			BuildTypeNameSignature(typeNames),
			isShared.ToString()
		};
		if (normalizedContentFingerprint.Length > 0)
		{
			parts.Add("content=" + normalizedContentFingerprint);
		}
		else
		{
			parts.Add("params=" + BuildParameterSignature(parameters));
		}
		return HashString(string.Join("|", parts));
	}

	private static string BuildSystemFingerprintCore(string typeClassName, string categoryName, string typeName, bool supportsRoutingDependencies, string classificationCode, string segmentName, string materialName, string shape, string routingPreferenceSignature, string compoundStructureSignature, bool includeDetailedComponents, bool detailedComponentsCaptured, string detailedComponentSignature, IEnumerable<SystemTypeDetailedComponentSnapshotItem> detailedComponents)
	{
		bool useDetailedComponents = includeDetailedComponents && detailedComponentsCaptured && SystemTypeDetailedComponentSnapshotService.SupportsDetailedComponents(typeClassName);
		bool useRequiredCurtainPanelComponents = detailedComponentsCaptured && SystemTypeDetailedComponentSnapshotService.HasRequiredCurtainPanelComponents(detailedComponents);
		List<string> parts = new List<string>
		{
			useRequiredCurtainPanelComponents ? "system-v5" : (useDetailedComponents ? "system-v4" : "system-v3"),
			Normalize(typeClassName),
			Normalize(categoryName),
			Normalize(typeName),
			supportsRoutingDependencies.ToString().ToLowerInvariant(),
			"classification=" + Normalize(classificationCode),
			"segment=" + Normalize(segmentName),
			"material=" + Normalize(materialName),
			"shape=" + Normalize(shape),
			"routing=" + NormalizeRoutingPreferenceSignature(routingPreferenceSignature),
			"compound=" + Normalize(compoundStructureSignature)
		};
		if (useDetailedComponents)
		{
			string optionalSignature = SystemTypeDetailedComponentSnapshotService.BuildOptionalDetailedComponentSignature(detailedComponents);
			parts.Add("detailed-components=" + Normalize(string.IsNullOrWhiteSpace(optionalSignature) ? detailedComponentSignature : optionalSignature));
		}
		if (useRequiredCurtainPanelComponents)
		{
			parts.Add("curtain-panel-dependencies=" + Normalize(SystemTypeDetailedComponentSnapshotService.BuildRequiredCurtainPanelSignature(detailedComponents)));
		}
		return HashString(string.Join("|", parts));
	}

	public static string NormalizeRoutingPreferenceSignature(string signature)
	{
		if (string.IsNullOrWhiteSpace(signature))
		{
			return string.Empty;
		}
		List<string> lines = new List<string>();
		string[] array = signature.Replace("\r", "\n").Split(new string[1] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
		for (int i = 0; i < array.Length; i = checked(i + 1))
		{
			string line = Normalize(array[i]);
			if (!string.IsNullOrWhiteSpace(line))
			{
				lines.Add(NormalizeRoutingPreferenceRuleLine(line));
			}
		}
		return string.Join("\n", lines.OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal));
	}

	private static string NormalizeRoutingPreferenceRuleLine(string line)
	{
		string[] parts = (line ?? string.Empty).Split('|');
		if (parts.Length < 11)
		{
			return Normalize(line);
		}
		List<string> stableParts = new List<string>
		{
			Normalize(parts[0]),
			Normalize(parts[1]),
			Normalize(parts[2]),
			Normalize(parts[3]),
			Normalize(parts[4]),
			Normalize(parts[5]),
			Normalize(parts[6]),
			NormalizeRoutingCriteriaSignature(string.Join("|", parts.Skip(10)))
		};
		return string.Join("|", stableParts);
	}

	private static string NormalizeRoutingCriteriaSignature(string criteriaSignature)
	{
		if (string.IsNullOrWhiteSpace(criteriaSignature))
		{
			return string.Empty;
		}
		List<string> criteria = new List<string>();
		string[] array = criteriaSignature.Split('&');
		for (int i = 0; i < array.Length; i = checked(i + 1))
		{
			string criterion = Normalize(array[i]);
			if (criterion.Length != 0)
			{
				List<string> tokens = (from x in criterion.Split(':')
					select NormalizeRoutingCriteriaToken(x) into x
					where !string.IsNullOrWhiteSpace(x)
					select x).ToList();
				if (tokens.Count > 0)
				{
					criteria.Add(string.Join(":", tokens));
				}
			}
		}
		return string.Join("&", criteria);
	}

	private static string NormalizeRoutingCriteriaToken(string token)
	{
		string normalizedToken = Normalize(token);
		if (normalizedToken.Length == 0)
		{
			return string.Empty;
		}
		int equalsIndex = normalizedToken.IndexOf('=');
		checked
		{
			if (equalsIndex <= 0 || equalsIndex >= normalizedToken.Length - 1)
			{
				return normalizedToken;
			}
			string name = normalizedToken.Substring(0, equalsIndex);
			if (double.TryParse(normalizedToken.Substring(equalsIndex + 1), NumberStyles.Float, CultureInfo.InvariantCulture, out var numericValue))
			{
				return name + "=" + Math.Round(numericValue, 9).ToString("G12", CultureInfo.InvariantCulture);
			}
			return normalizedToken;
		}
	}

	private static string BuildTypeNameSignature(IEnumerable<string> typeNames)
	{
		if (typeNames == null)
		{
			return string.Empty;
		}
		return Normalize(string.Join("|", typeNames.OrderBy([SpecialName] (string x) => Normalize(x), StringComparer.Ordinal)));
	}

	private static string BuildParameterSignature(IEnumerable<StandardFamilyParameterSnapshotItem> parameters)
	{
		return FamilyParameterSnapshotNormalizationService.BuildComparableDefinitionSignature(parameters);
	}

	private static string BuildPortableParameterIdentity(StandardFamilyParameterSnapshotItem parameter)
	{
		if (parameter == null)
		{
			return string.Empty;
		}
		if (!string.IsNullOrWhiteSpace(parameter.ExternalGuid))
		{
			return "guid:" + parameter.ExternalGuid.Trim();
		}
		if (!string.IsNullOrWhiteSpace(parameter.ParameterId) && parameter.ParameterId.Trim().StartsWith("-", StringComparison.Ordinal))
		{
			return "builtin:" + parameter.ParameterId.Trim();
		}
		return string.Empty;
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
