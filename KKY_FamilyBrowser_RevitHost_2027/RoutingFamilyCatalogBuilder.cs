using System;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;

public sealed class RoutingFamilyCatalogBuilder
{
	private RoutingFamilyCatalogBuilder()
	{
	}

	public static RoutingFamilyCatalogSnapshot Build(string sourceId, ProjectContentSnapshot projectSnapshot)
	{
		RoutingFamilyCatalogSnapshot routingFamilyCatalogSnapshot = new RoutingFamilyCatalogSnapshot();
		routingFamilyCatalogSnapshot.SourceId = sourceId ?? string.Empty;
		routingFamilyCatalogSnapshot.DocumentTitle = projectSnapshot?.DocumentTitle ?? string.Empty;
		routingFamilyCatalogSnapshot.DocumentPath = projectSnapshot?.DocumentPath ?? string.Empty;
		routingFamilyCatalogSnapshot.CapturedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		RoutingFamilyCatalogSnapshot catalog = routingFamilyCatalogSnapshot;
		if (projectSnapshot == null || projectSnapshot.LoadableFamilies == null)
		{
			return catalog;
		}
		foreach (ProjectLoadableFamilySnapshotItem family in projectSnapshot.LoadableFamilies.Where([SpecialName] (ProjectLoadableFamilySnapshotItem x) => !FamilyBrowserFamilyClassificationService.IsTypeManagedFamilyLike(x.CategoryName, x.CategoryId, x.FamilyName)).OrderBy<ProjectLoadableFamilySnapshotItem, string>([SpecialName] (ProjectLoadableFamilySnapshotItem x) => BuildFamilyKey(x.CategoryName, x.FamilyName), StringComparer.Ordinal))
		{
			string familyKey = BuildFamilyKey(family.CategoryName, family.FamilyName);
			RoutingFamilyCatalogEntry entry = new RoutingFamilyCatalogEntry
			{
				LibraryFamilyId = familyKey,
				FamilyName = family.FamilyName,
				FamilyFingerprint = ProjectSnapshotFingerprintService.BuildLoadableFingerprint(family)
			};
			foreach (string familyTypeName in family.TypeNames.OrderBy<string, string>([SpecialName] (string x) => Normalize(x), StringComparer.Ordinal))
			{
				entry.Types.Add(new RoutingFamilyTypeSnapshot
				{
					TypeName = familyTypeName,
					TypeFingerprint = SystemTypeFingerprintService.ComputeSimpleTypeFingerprint(familyKey, familyTypeName)
				});
			}
			catalog.Families.Add(entry);
		}
		return catalog;
	}

	public static string BuildFamilyKey(string categoryName, string familyName)
	{
		return Normalize(categoryName) + "|" + Normalize(familyName);
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
