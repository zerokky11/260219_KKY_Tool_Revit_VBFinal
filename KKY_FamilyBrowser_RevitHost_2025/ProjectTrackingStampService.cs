using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;

public sealed class ProjectTrackingStampService
{
	private ProjectTrackingStampService()
	{
	}

	public static ProjectTrackingCatalog BuildCatalog(StandardLibraryRegistrationRecord registration, StandardLibrarySnapshot standardSnapshot, ProjectContentSnapshot projectSnapshot)
	{
		ProjectTrackingCatalog catalog = new ProjectTrackingCatalog
		{
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			SourceId = registration.SourceId,
			SourceDisplayName = registration.DisplayName,
			ApprovedStandardStamp = standardSnapshot.CapturedAtUtc,
			ProjectDocumentTitle = projectSnapshot.DocumentTitle,
			ProjectDocumentPath = projectSnapshot.DocumentPath
		};
		Dictionary<string, ProjectLoadableFamilySnapshotItem> projectLoadableMap = BuildFirstMap(projectSnapshot.LoadableFamilies.Where([SpecialName] (ProjectLoadableFamilySnapshotItem x) => !FamilyBrowserFamilyClassificationService.IsTypeManagedFamilyLike(x.CategoryName, x.CategoryId, x.FamilyName)), [SpecialName] (ProjectLoadableFamilySnapshotItem x) => Normalize(x.CategoryName) + "|" + Normalize(x.FamilyName), StringComparer.Ordinal);
		foreach (StandardLoadableFamilySnapshotItem standardFamily in standardSnapshot.LoadableFamilies.Where([SpecialName] (StandardLoadableFamilySnapshotItem x) => !FamilyBrowserFamilyClassificationService.IsTypeManagedFamilyLike(x.CategoryName, x.CategoryId, x.FamilyName)))
		{
			string key = Normalize(standardFamily.CategoryName) + "|" + Normalize(standardFamily.FamilyName);
			ProjectLoadableFamilySnapshotItem projectFamily = null;
			if (projectLoadableMap.TryGetValue(key, out projectFamily))
			{
				string standardFingerprint = ProjectSnapshotFingerprintService.BuildLoadableFingerprint(standardFamily);
				string projectFingerprint = ProjectSnapshotFingerprintService.BuildLoadableFingerprint(projectFamily);
				if (string.Equals(standardFingerprint, projectFingerprint, StringComparison.Ordinal))
				{
					catalog.LoadableFamilies.Add(new TrackedLoadableFamilyState
					{
						IdentityKey = key,
						FamilyName = standardFamily.FamilyName,
						CategoryName = standardFamily.CategoryName,
						ApprovedFingerprint = standardFingerprint,
						ApprovedStandardStamp = standardSnapshot.CapturedAtUtc,
						ApprovedAtUtc = catalog.GeneratedAtUtc
					});
				}
			}
		}
		Dictionary<string, ProjectSystemTypeSnapshotItem> projectSystemMap = BuildFirstMap(projectSnapshot.SystemTypes, [SpecialName] (ProjectSystemTypeSnapshotItem x) => Normalize(x.TypeClassName) + "|" + Normalize(x.CategoryName) + "|" + Normalize(x.TypeName), StringComparer.Ordinal);
		foreach (StandardSystemTypeSnapshotItem standardType in standardSnapshot.SystemTypes)
		{
			string key2 = Normalize(standardType.TypeClassName) + "|" + Normalize(standardType.CategoryName) + "|" + Normalize(standardType.TypeName);
			ProjectSystemTypeSnapshotItem projectType = null;
			if (projectSystemMap.TryGetValue(key2, out projectType))
			{
				string standardFingerprint2 = ProjectSnapshotFingerprintService.BuildSystemFingerprint(standardType);
				string projectFingerprint2 = ProjectSnapshotFingerprintService.BuildSystemFingerprint(projectType);
				if (string.Equals(standardFingerprint2, projectFingerprint2, StringComparison.Ordinal))
				{
					catalog.SystemTypes.Add(new TrackedSystemTypeState
					{
						IdentityKey = key2,
						TypeName = standardType.TypeName,
						CategoryName = standardType.CategoryName,
						TypeClassName = standardType.TypeClassName,
						ApprovedFingerprint = standardFingerprint2,
						ApprovedSemanticFingerprint = standardFingerprint2,
						ApprovedStandardStamp = standardSnapshot.CapturedAtUtc,
						ApprovedAtUtc = catalog.GeneratedAtUtc
					});
				}
			}
		}
		return catalog;
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

	private static string Normalize(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().ToLowerInvariant();
	}
}
