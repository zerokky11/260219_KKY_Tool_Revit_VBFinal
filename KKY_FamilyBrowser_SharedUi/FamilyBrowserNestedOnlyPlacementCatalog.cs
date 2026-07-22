using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

public sealed class FamilyBrowserNestedOnlyPlacementCatalog
{
	public int SchemaVersion { get; set; }

	public bool IsComplete { get; set; }

	public string IncompleteReason { get; set; }

	public string SourceId { get; set; }

	public string SourceSnapshotPath { get; set; }

	public string SourceSnapshotLastWriteUtc { get; set; }

	public long SourceSnapshotLength { get; set; }

	public string SnapshotMode { get; set; }

	public string GeneratedAtUtc { get; set; }

	public bool FingerprintCoverageComplete { get; set; }

	public int SkippedFingerprintCount { get; set; }

	public List<FamilyBrowserNestedOnlyPlacementEntry> Entries { get; set; }

	public FamilyBrowserNestedOnlyPlacementCatalog()
	{
		SchemaVersion = 2;
		IncompleteReason = string.Empty;
		SourceId = string.Empty;
		SourceSnapshotPath = string.Empty;
		SourceSnapshotLastWriteUtc = string.Empty;
		SnapshotMode = string.Empty;
		GeneratedAtUtc = string.Empty;
		Entries = new List<FamilyBrowserNestedOnlyPlacementEntry>();
	}
}

public sealed class FamilyBrowserNestedOnlyPlacementEntry
{
	public string FamilyName { get; set; }

	public string CategoryName { get; set; }

	public string CategoryId { get; set; }

	public string ContentFingerprint { get; set; }

	public bool IsShared { get; set; }

	public string SourceId { get; set; }

	public List<string> ParentFamilyNames { get; set; }

	public FamilyBrowserNestedOnlyPlacementEntry()
	{
		FamilyName = string.Empty;
		CategoryName = string.Empty;
		CategoryId = string.Empty;
		ContentFingerprint = string.Empty;
		SourceId = string.Empty;
		ParentFamilyNames = new List<string>();
	}
}

public static class FamilyBrowserNestedOnlyPlacementFingerprintPolicy
{
	public static bool IsExactMatch(FamilyBrowserNestedOnlyPlacementEntry entry, string projectContentFingerprint, bool projectIsShared)
	{
		return entry != null &&
			entry.IsShared &&
			projectIsShared &&
			!string.IsNullOrWhiteSpace(entry.ContentFingerprint) &&
			!string.IsNullOrWhiteSpace(projectContentFingerprint) &&
			string.Equals(entry.ContentFingerprint.Trim(), projectContentFingerprint.Trim(), StringComparison.OrdinalIgnoreCase);
	}
}

public static class FamilyBrowserNestedOnlyPlacementCatalogStore
{
	private const string SidecarSuffix = ".nested-only-placement-v2.json";

	public static string GetSidecarPath(string snapshotPath)
	{
		return string.IsNullOrWhiteSpace(snapshotPath) ? string.Empty : snapshotPath.Trim() + SidecarSuffix;
	}

	public static FamilyBrowserNestedOnlyPlacementCatalog Build(StandardLibrarySnapshot snapshot, string snapshotPath)
	{
		FamilyBrowserNestedOnlyPlacementCatalog catalog = new FamilyBrowserNestedOnlyPlacementCatalog
		{
			SourceId = snapshot?.SourceId ?? string.Empty,
			SourceSnapshotPath = snapshotPath ?? string.Empty,
			SnapshotMode = snapshot?.SnapshotMode ?? string.Empty,
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
		};
		ApplySnapshotFileStamp(catalog, snapshotPath);
		List<StandardLoadableFamilySnapshotItem> families = snapshot?.LoadableFamilies ?? new List<StandardLoadableFamilySnapshotItem>();
		bool isPrecise = string.Equals(snapshot?.SnapshotMode, "Precise", StringComparison.OrdinalIgnoreCase);
		bool usageCaptured = families.All(delegate(StandardLoadableFamilySnapshotItem item)
		{
			return item == null || item.StandalonePlacementUsageCaptured;
		});
		Dictionary<string, HashSet<string>> parentNamesByChild = BuildParentLookup(families);
		bool nestedMetadataConsistent = families.Where(delegate(StandardLoadableFamilySnapshotItem item)
		{
			return item != null && item.IsNestedLoadableChild;
		}).All(delegate(StandardLoadableFamilySnapshotItem item)
		{
			return HasParentReference(parentNamesByChild, item.CategoryId, item.CategoryName, item.FamilyName);
		});
		catalog.IsComplete = isPrecise && usageCaptured && nestedMetadataConsistent;
		if (!isPrecise)
		{
			catalog.IncompleteReason = "PreciseStandardScanRequired";
		}
		else if (!usageCaptured)
		{
			catalog.IncompleteReason = "StandalonePlacementUsageNotCaptured";
		}
		else if (!nestedMetadataConsistent)
		{
			catalog.IncompleteReason = "NestedParentMetadataIncomplete";
		}
		if (!catalog.IsComplete)
		{
			return catalog;
		}
		Dictionary<string, FamilyBrowserNestedOnlyPlacementEntry> entries = new Dictionary<string, FamilyBrowserNestedOnlyPlacementEntry>(StringComparer.Ordinal);
		int skippedFingerprintCount = 0;
		foreach (StandardLoadableFamilySnapshotItem item in families)
		{
			if (item == null || !item.IsNestedLoadableChild || !item.IsShared || !item.StandalonePlacementUsageCaptured || item.StandaloneInstanceCount > 0)
			{
				continue;
			}
			if (string.IsNullOrWhiteSpace(item.ContentFingerprint))
			{
				skippedFingerprintCount++;
				continue;
			}
			string key = BuildIdentityKey(item.CategoryId, item.CategoryName, item.FamilyName);
			if (string.IsNullOrWhiteSpace(key) || entries.ContainsKey(key))
			{
				continue;
			}
			HashSet<string> parentNames = ResolveParentNames(parentNamesByChild, item.CategoryId, item.CategoryName, item.FamilyName);
			if (parentNames.Count == 0)
			{
				continue;
			}
			entries.Add(key, new FamilyBrowserNestedOnlyPlacementEntry
			{
				FamilyName = item.FamilyName ?? string.Empty,
				CategoryName = item.CategoryName ?? string.Empty,
				CategoryId = item.CategoryId ?? string.Empty,
				ContentFingerprint = item.ContentFingerprint ?? string.Empty,
				IsShared = item.IsShared,
				SourceId = snapshot?.SourceId ?? string.Empty,
				ParentFamilyNames = parentNames.OrderBy(delegate(string value)
				{
					return value;
				}, StringComparer.OrdinalIgnoreCase).ToList()
			});
		}
		catalog.Entries = entries.Values.OrderBy(delegate(FamilyBrowserNestedOnlyPlacementEntry item)
		{
			return Normalize(item.CategoryName) + "|" + Normalize(item.FamilyName);
		}, StringComparer.Ordinal).ToList();
		catalog.SkippedFingerprintCount = skippedFingerprintCount;
		catalog.FingerprintCoverageComplete = skippedFingerprintCount == 0;
		return catalog;
	}

	public static string SaveForSnapshot(string snapshotPath, StandardLibrarySnapshot snapshot)
	{
		if (string.IsNullOrWhiteSpace(snapshotPath))
		{
			return string.Empty;
		}
		FamilyBrowserNestedOnlyPlacementCatalog catalog = Build(snapshot, snapshotPath);
		string outputPath = GetSidecarPath(snapshotPath);
		WriteAtomically(outputPath, PlainJsonReportWriter.Serialize(catalog));
		TryWriteLocalCache(snapshotPath, catalog);
		return outputPath;
	}

	public static FamilyBrowserNestedOnlyPlacementCatalog TryLoadForSnapshot(string snapshotPath)
	{
		if (string.IsNullOrWhiteSpace(snapshotPath))
		{
			return null;
		}
		bool snapshotAvailable = File.Exists(snapshotPath);
		string sidecarPath = GetSidecarPath(snapshotPath);
		FamilyBrowserNestedOnlyPlacementCatalog catalog = TryLoadCatalog(sidecarPath);
		if (catalog != null && (!snapshotAvailable || MatchesSnapshotFile(catalog, snapshotPath)))
		{
			TryWriteLocalCache(snapshotPath, catalog);
			return catalog;
		}
		if (snapshotAvailable)
		{
			try
			{
				StandardLibrarySnapshot snapshot = DataContractJsonFileStore.Load<StandardLibrarySnapshot>(snapshotPath);
				catalog = Build(snapshot, snapshotPath);
				try
				{
					WriteAtomically(sidecarPath, PlainJsonReportWriter.Serialize(catalog));
				}
				catch
				{
				}
				TryWriteLocalCache(snapshotPath, catalog);
				return catalog;
			}
			catch
			{
				return null;
			}
		}
		return TryLoadCatalog(GetLocalCachePath(snapshotPath));
	}

	public static bool Contains(FamilyBrowserNestedOnlyPlacementCatalog catalog, string categoryId, string categoryName, string familyName)
	{
		return FindEntry(catalog, categoryId, categoryName, familyName) != null;
	}

	public static FamilyBrowserNestedOnlyPlacementEntry FindEntry(FamilyBrowserNestedOnlyPlacementCatalog catalog, string categoryId, string categoryName, string familyName)
	{
		return FindEntries(catalog, categoryId, categoryName, familyName).FirstOrDefault();
	}

	public static List<FamilyBrowserNestedOnlyPlacementEntry> FindEntries(FamilyBrowserNestedOnlyPlacementCatalog catalog, string categoryId, string categoryName, string familyName)
	{
		List<FamilyBrowserNestedOnlyPlacementEntry> matches = new List<FamilyBrowserNestedOnlyPlacementEntry>();
		if (catalog == null || !catalog.IsComplete || string.IsNullOrWhiteSpace(familyName))
		{
			return matches;
		}
		string normalizedFamily = Normalize(familyName);
		string normalizedCategoryId = Normalize(categoryId);
		string normalizedCategoryName = Normalize(categoryName);
		foreach (FamilyBrowserNestedOnlyPlacementEntry entry in catalog.Entries ?? new List<FamilyBrowserNestedOnlyPlacementEntry>())
		{
			if (entry == null || !string.Equals(Normalize(entry.FamilyName), normalizedFamily, StringComparison.Ordinal))
			{
				continue;
			}
			string entryCategoryId = Normalize(entry.CategoryId);
			if (!string.IsNullOrWhiteSpace(entryCategoryId) && !string.IsNullOrWhiteSpace(normalizedCategoryId))
			{
				if (string.Equals(entryCategoryId, normalizedCategoryId, StringComparison.Ordinal))
				{
					matches.Add(entry);
				}
				continue;
			}
			string entryCategoryName = Normalize(entry.CategoryName);
			if (string.IsNullOrWhiteSpace(entryCategoryName) || string.IsNullOrWhiteSpace(normalizedCategoryName) || string.Equals(entryCategoryName, normalizedCategoryName, StringComparison.Ordinal))
			{
				matches.Add(entry);
			}
		}
		return matches;
	}

	private static Dictionary<string, HashSet<string>> BuildParentLookup(IEnumerable<StandardLoadableFamilySnapshotItem> families)
	{
		Dictionary<string, HashSet<string>> result = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
		foreach (StandardLoadableFamilySnapshotItem parent in families ?? Enumerable.Empty<StandardLoadableFamilySnapshotItem>())
		{
			if (parent == null)
			{
				continue;
			}
			foreach (StandardNestedLoadableFamilySnapshotItem child in parent.NestedLoadableFamilies ?? new List<StandardNestedLoadableFamilySnapshotItem>())
			{
				if (child == null || string.IsNullOrWhiteSpace(child.FamilyName))
				{
					continue;
				}
				AddParentReference(result, BuildIdentityKey(child.CategoryId, child.CategoryName, child.FamilyName), parent.FamilyName);
				AddParentReference(result, BuildFamilyOnlyKey(child.FamilyName), parent.FamilyName);
			}
		}
		return result;
	}

	private static void AddParentReference(Dictionary<string, HashSet<string>> lookup, string key, string parentFamilyName)
	{
		if (lookup == null || string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(parentFamilyName))
		{
			return;
		}
		HashSet<string> names;
		if (!lookup.TryGetValue(key, out names))
		{
			names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
			lookup.Add(key, names);
		}
		names.Add(parentFamilyName.Trim());
	}

	private static bool HasParentReference(Dictionary<string, HashSet<string>> lookup, string categoryId, string categoryName, string familyName)
	{
		return ResolveParentNames(lookup, categoryId, categoryName, familyName).Count > 0;
	}

	private static HashSet<string> ResolveParentNames(Dictionary<string, HashSet<string>> lookup, string categoryId, string categoryName, string familyName)
	{
		HashSet<string> result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		if (lookup == null)
		{
			return result;
		}
		HashSet<string> exact;
		if (lookup.TryGetValue(BuildIdentityKey(categoryId, categoryName, familyName), out exact))
		{
			result.UnionWith(exact);
		}
		if (result.Count == 0)
		{
			HashSet<string> byName;
			if (lookup.TryGetValue(BuildFamilyOnlyKey(familyName), out byName))
			{
				result.UnionWith(byName);
			}
		}
		return result;
	}

	private static string BuildIdentityKey(string categoryId, string categoryName, string familyName)
	{
		string family = Normalize(familyName);
		if (string.IsNullOrWhiteSpace(family))
		{
			return string.Empty;
		}
		string category = Normalize(categoryId);
		if (string.IsNullOrWhiteSpace(category))
		{
			category = Normalize(categoryName);
		}
		return category + "|" + family;
	}

	private static string BuildFamilyOnlyKey(string familyName)
	{
		string family = Normalize(familyName);
		return string.IsNullOrWhiteSpace(family) ? string.Empty : "*|" + family;
	}

	private static FamilyBrowserNestedOnlyPlacementCatalog TryLoadCatalog(string path)
	{
		if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
		{
			return null;
		}
		try
		{
			FamilyBrowserNestedOnlyPlacementCatalog catalog = DataContractJsonFileStore.Load<FamilyBrowserNestedOnlyPlacementCatalog>(path);
			return catalog != null && catalog.SchemaVersion == 2 ? catalog : null;
		}
		catch
		{
			return null;
		}
	}

	private static void ApplySnapshotFileStamp(FamilyBrowserNestedOnlyPlacementCatalog catalog, string snapshotPath)
	{
		if (catalog == null || string.IsNullOrWhiteSpace(snapshotPath) || !File.Exists(snapshotPath))
		{
			return;
		}
		FileInfo info = new FileInfo(snapshotPath);
		catalog.SourceSnapshotLength = info.Length;
		catalog.SourceSnapshotLastWriteUtc = info.LastWriteTimeUtc.ToString("O", CultureInfo.InvariantCulture);
	}

	private static bool MatchesSnapshotFile(FamilyBrowserNestedOnlyPlacementCatalog catalog, string snapshotPath)
	{
		if (catalog == null || string.IsNullOrWhiteSpace(snapshotPath) || !File.Exists(snapshotPath))
		{
			return false;
		}
		FileInfo info = new FileInfo(snapshotPath);
		return catalog.SourceSnapshotLength == info.Length && string.Equals(catalog.SourceSnapshotLastWriteUtc ?? string.Empty, info.LastWriteTimeUtc.ToString("O", CultureInfo.InvariantCulture), StringComparison.Ordinal);
	}

	private static void TryWriteLocalCache(string snapshotPath, FamilyBrowserNestedOnlyPlacementCatalog catalog)
	{
		try
		{
			WriteAtomically(GetLocalCachePath(snapshotPath), PlainJsonReportWriter.Serialize(catalog));
		}
		catch
		{
		}
	}

	private static string GetLocalCachePath(string snapshotPath)
	{
		string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KKY", "FamilyBrowser", "Cache", "Guard", "NestedOnly");
		byte[] bytes = Encoding.UTF8.GetBytes((snapshotPath ?? string.Empty).Trim().ToUpperInvariant());
		using (SHA256 sha = SHA256.Create())
		{
			string hash = BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", string.Empty);
			return Path.Combine(root, hash + ".json");
		}
	}

	private static void WriteAtomically(string path, string content)
	{
		if (string.IsNullOrWhiteSpace(path))
		{
			return;
		}
		string directory = Path.GetDirectoryName(path);
		if (!string.IsNullOrWhiteSpace(directory))
		{
			Directory.CreateDirectory(directory);
		}
		string tempPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
		try
		{
			File.WriteAllText(tempPath, content ?? string.Empty, new UTF8Encoding(false));
			FamilyBrowserAtomicFileService.Promote(tempPath, path);
		}
		finally
		{
			if (File.Exists(tempPath))
			{
				File.Delete(tempPath);
			}
		}
	}

	private static string Normalize(string value)
	{
		return (value ?? string.Empty).Trim().ToUpperInvariant();
	}
}
