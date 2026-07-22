using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

public sealed class FamilyBrowserDataLoader : IFamilyBrowserDataLoader
{
	private const int SchemaVersion = 2;
	private const string ManifestFileName = "family-browser-manifest-v2.json";
	private const string IndexFileName = "standard-browser-index-v2.json";
	private const string DetailsCatalogFileName = "standard-browser-details-v2.json";
	private const string ThumbnailIndexFileName = "thumbnail-index-v2.json";
	private const string ProjectStateFileName = "project-browser-state-v2.json";
	private const string ManifestMutationLockFileName = "family-browser-manifest-v2.lock";
	private static readonly object PreparedCatalogLock = new object();
	private static readonly Dictionary<string, FamilyBrowserStandardListCatalog> PreparedCatalogs = new Dictionary<string, FamilyBrowserStandardListCatalog>(StringComparer.OrdinalIgnoreCase);
	private static readonly object ThumbnailLookupLock = new object();
	private static readonly Dictionary<string, Dictionary<string, string>> ThumbnailLookups = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
	private static readonly HashSet<string> ThumbnailLookupBuilds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

	public static readonly FamilyBrowserDataLoader Default = new FamilyBrowserDataLoader();

	public FamilyBrowserCompactSnapshotLoadResult LoadSnapshotProjection(string workspaceRoot, StandardLibraryRegistrationRecord registration, bool allowSourceFallback)
	{
		Stopwatch sw = Stopwatch.StartNew();
		FamilyBrowserCompactSnapshotLoadResult result = new FamilyBrowserCompactSnapshotLoadResult();
		if (registration == null || string.IsNullOrWhiteSpace(registration.LastSnapshotPath))
		{
			result.Reason = "Standard snapshot registration is missing.";
			return result;
		}

		string sourceKey = BuildSourceKey(registration.SourceId, registration.LastSnapshotPath);
		string managedSourceFolder = GetManagedSourceFolder(workspaceRoot, sourceKey);
		string remoteManifestPath = Path.Combine(managedSourceFolder, ManifestFileName);
		string localSourceFolder = GetLocalSourceFolder(sourceKey);
		string localManifestPath = Path.Combine(localSourceFolder, ManifestFileName);
		FamilyBrowserManifestV2 manifest = TryLoad<FamilyBrowserManifestV2>(remoteManifestPath);
		bool remoteManifestAvailable = manifest != null && manifest.SchemaVersion == SchemaVersion;
		if (remoteManifestAvailable)
		{
			TryWriteJsonAtomic(localManifestPath, manifest);
		}
		else
		{
			manifest = TryLoad<FamilyBrowserManifestV2>(localManifestPath);
			result.OfflineCacheUsed = manifest != null && manifest.SchemaVersion == SchemaVersion;
		}

		if (manifest != null && manifest.SchemaVersion == SchemaVersion && manifest.StandardIndex != null && !string.IsNullOrWhiteSpace(manifest.StandardIndex.RelativePath))
		{
			string localIndexPath;
			bool localCacheHit;
			if (TryResolveArtifact(managedSourceFolder, localSourceFolder, manifest.StandardIndex, remoteManifestAvailable, out localIndexPath, out localCacheHit))
			{
				StandardBrowserIndexV2 index = TryLoad<StandardBrowserIndexV2>(localIndexPath);
				if (index != null && index.SchemaVersion == SchemaVersion)
				{
					result.Success = true;
					result.LocalCacheHit = localCacheHit;
					result.Manifest = manifest;
					result.Index = index;
					result.Revision = manifest.SourceSnapshotRevision ?? string.Empty;
					result.SourcePath = localIndexPath;
					result.BytesRead = SafeLength(localIndexPath);
					result.Snapshot = BuildProjection(index, registration, manifest.SourceKey);
					RegisterThumbnailIndex(workspaceRoot, managedSourceFolder, localSourceFolder, manifest, remoteManifestAvailable);
					sw.Stop();
					result.ElapsedMilliseconds = sw.ElapsedMilliseconds;
					RecordPerformance("compact-index-load", sw.ElapsedMilliseconds, result.BytesRead, index.Items == null ? 0 : index.Items.Count, localCacheHit, localIndexPath, result.OfflineCacheUsed ? "offline-cache" : "ready");
					return result;
				}
			}
		}

		if (!allowSourceFallback || !File.Exists(registration.LastSnapshotPath))
		{
			result.Reason = "Browser V2 index is unavailable.";
			sw.Stop();
			result.ElapsedMilliseconds = sw.ElapsedMilliseconds;
			RecordPerformance("compact-index-miss", sw.ElapsedMilliseconds, 0L, 0, false, registration.LastSnapshotPath, result.Reason);
			return result;
		}

		try
		{
			Stopwatch fallbackSw = Stopwatch.StartNew();
			StandardLibrarySnapshot snapshot = DataContractJsonFileStore.Load<StandardLibrarySnapshot>(registration.LastSnapshotPath);
			fallbackSw.Stop();
			if (snapshot == null)
			{
				result.Reason = "Standard snapshot fallback was empty.";
				return result;
			}
			PublishStandardArtifacts(workspaceRoot, registration.LastSnapshotPath, snapshot);
			FamilyBrowserCompactSnapshotLoadResult published = LoadSnapshotProjection(workspaceRoot, registration, false);
			published.SourceFallbackUsed = true;
			published.ElapsedMilliseconds += fallbackSw.ElapsedMilliseconds;
			RecordPerformance("snapshot-source-fallback", fallbackSw.ElapsedMilliseconds, SafeLength(registration.LastSnapshotPath), (snapshot.LoadableFamilies == null ? 0 : snapshot.LoadableFamilies.Count) + (snapshot.SystemTypes == null ? 0 : snapshot.SystemTypes.Count), false, registration.LastSnapshotPath, published.Success ? "published-v2" : "publish-failed");
			return published;
		}
		catch (Exception ex)
		{
			result.Reason = "Standard snapshot fallback failed: " + ex.Message;
			sw.Stop();
			result.ElapsedMilliseconds = sw.ElapsedMilliseconds;
			RecordPerformance("snapshot-source-fallback-failed", sw.ElapsedMilliseconds, 0L, 0, false, registration.LastSnapshotPath, ex.Message);
			return result;
		}
	}

	public BrowserDetailRecord LoadDetail(string workspaceRoot, FamilyBrowserManifestV2 manifest, string detailKey)
	{
		if (manifest == null || string.IsNullOrWhiteSpace(manifest.SourceKey) || string.IsNullOrWhiteSpace(detailKey))
		{
			return null;
		}
		string managedSourceFolder = GetManagedSourceFolder(workspaceRoot, manifest.SourceKey);
		string localSourceFolder = GetLocalSourceFolder(manifest.SourceKey);
		FamilyBrowserArtifactReferenceV2 detailReference = new FamilyBrowserArtifactReferenceV2 { RelativePath = detailKey };
		string localPath;
		bool localCacheHit;
		if (!TryResolveArtifact(managedSourceFolder, localSourceFolder, detailReference, Directory.Exists(managedSourceFolder), out localPath, out localCacheHit))
		{
			return null;
		}
		return TryLoad<BrowserDetailRecord>(localPath);
	}

	public bool TryLoadRowCache(string cacheKey, out FamilyBrowserRowCacheV2 cache)
	{
		cache = null;
		if (string.IsNullOrWhiteSpace(cacheKey))
		{
			return false;
		}
		string keyHash = HashText(cacheKey);
		string path = Path.Combine(GetLocalCacheRoot(), "rows", "browser-row-cache-v2-" + keyHash.Substring(0, 32) + ".json");
		Stopwatch sw = Stopwatch.StartNew();
		FamilyBrowserRowCacheV2 loaded = TryLoad<FamilyBrowserRowCacheV2>(path);
		sw.Stop();
		if (loaded == null || loaded.SchemaVersion != SchemaVersion || !string.Equals(loaded.CacheKeyHash, keyHash, StringComparison.Ordinal))
		{
			RecordPerformance("row-cache-miss", sw.ElapsedMilliseconds, SafeLength(path), 0, false, path, "missing-or-stale");
			return false;
		}
		cache = loaded;
		int rowCount = (loaded.Families == null ? 0 : loaded.Families.Count) + (loaded.Systems == null ? 0 : loaded.Systems.Count);
		RecordPerformance("row-cache-hit", sw.ElapsedMilliseconds, SafeLength(path), rowCount, true, path, "ready");
		return true;
	}

	public void SaveRowCache(string cacheKey, FamilyBrowserRowCacheV2 cache)
	{
		if (string.IsNullOrWhiteSpace(cacheKey) || cache == null)
		{
			return;
		}
		string keyHash = HashText(cacheKey);
		cache.SchemaVersion = SchemaVersion;
		cache.CacheKeyHash = keyHash;
		cache.SavedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		string path = Path.Combine(GetLocalCacheRoot(), "rows", "browser-row-cache-v2-" + keyHash.Substring(0, 32) + ".json");
		Stopwatch sw = Stopwatch.StartNew();
		TryWriteJsonAtomic(path, cache);
		sw.Stop();
		int rowCount = (cache.Families == null ? 0 : cache.Families.Count) + (cache.Systems == null ? 0 : cache.Systems.Count);
		RecordPerformance("row-cache-save", sw.ElapsedMilliseconds, SafeLength(path), rowCount, true, path, "saved");
	}

	public static FamilyBrowserSyntheticPerformanceAuditResult RunSyntheticPerformanceAudit(int familyCount, int systemCount)
	{
		FamilyBrowserSyntheticPerformanceAuditResult result = new FamilyBrowserSyntheticPerformanceAuditResult
		{
			FamilyCount = Math.Max(0, familyCount),
			SystemCount = Math.Max(0, systemCount)
		};
		string cacheKey = "quality-gate|" + Guid.NewGuid().ToString("N") + "|" + result.FamilyCount.ToString(CultureInfo.InvariantCulture) + "|" + result.SystemCount.ToString(CultureInfo.InvariantCulture);
		string keyHash = HashText(cacheKey);
		string path = Path.Combine(GetLocalCacheRoot(), "rows", "browser-row-cache-v2-" + keyHash.Substring(0, 32) + ".json");
		try
		{
			FamilyBrowserRowCacheV2 cache = new FamilyBrowserRowCacheV2();
			for (int i = 0; i < result.FamilyCount; i++)
			{
				cache.Families.Add(new FamilyBrowserCachedFamilyRowV2
				{
					Status = "Load Available",
					RawStatus = "LoadAvailable",
					DisciplineKey = i % 2 == 0 ? "Mechanical" : "Electrical",
					DisciplineLabel = i % 2 == 0 ? "Mechanical" : "Electrical",
					Name = "PERF_FAMILY_" + i.ToString("D4", CultureInfo.InvariantCulture),
					Category = i % 2 == 0 ? "Mechanical Equipment" : "Electrical Fixtures",
					CategoryGroup = "Model",
					Action = "Load",
					Notes = "Synthetic row cache audit"
				});
			}
			for (int i = 0; i < result.SystemCount; i++)
			{
				cache.Systems.Add(new FamilyBrowserCachedSystemRowV2
				{
					Status = "Apply Available",
					RawStatus = "LoadAvailable",
					DisciplineKey = i % 2 == 0 ? "Mechanical" : "FireProtection",
					DisciplineLabel = i % 2 == 0 ? "Mechanical" : "Fire Protection",
					Name = "PERF_SYSTEM_" + i.ToString("D4", CultureInfo.InvariantCulture),
					Category = i % 2 == 0 ? "Duct System" : "Piping System",
					SystemFamilyKind = i % 2 == 0 ? "Duct" : "Pipe",
					Action = "Apply",
					Notes = "Synthetic row cache audit"
				});
			}

			Stopwatch sw = Stopwatch.StartNew();
			Default.SaveRowCache(cacheKey, cache);
			sw.Stop();
			result.SaveMilliseconds = sw.ElapsedMilliseconds;
			result.CacheBytes = SafeLength(path);

			FamilyBrowserRowCacheV2 loaded;
			sw.Restart();
			bool coldLoaded = Default.TryLoadRowCache(cacheKey, out loaded);
			sw.Stop();
			result.ColdLoadMilliseconds = sw.ElapsedMilliseconds;

			sw.Restart();
			bool warmLoaded = Default.TryLoadRowCache(cacheKey, out loaded);
			sw.Stop();
			result.WarmLoadMilliseconds = sw.ElapsedMilliseconds;

			sw.Restart();
			bool offlineLoaded = Default.TryLoadRowCache(cacheKey, out loaded);
			sw.Stop();
			result.OfflineLoadMilliseconds = sw.ElapsedMilliseconds;
			int loadedFamilyCount = loaded == null || loaded.Families == null ? 0 : loaded.Families.Count;
			int loadedSystemCount = loaded == null || loaded.Systems == null ? 0 : loaded.Systems.Count;
			result.Success = coldLoaded && warmLoaded && offlineLoaded && loadedFamilyCount == result.FamilyCount && loadedSystemCount == result.SystemCount;
			RecordPerformance("synthetic-cache-audit", result.ColdLoadMilliseconds + result.WarmLoadMilliseconds + result.OfflineLoadMilliseconds, result.CacheBytes, result.FamilyCount + result.SystemCount, true, path, result.Success ? "pass" : "row-count-mismatch");
		}
		catch (Exception ex)
		{
			result.ErrorMessage = ex.Message;
			result.Success = false;
		}
		finally
		{
			TryDelete(path);
		}
		return result;
	}

	public static void PublishStandardArtifacts(string workspaceRoot, string snapshotPath, StandardLibrarySnapshot snapshot)
	{
		if (snapshot == null || string.IsNullOrWhiteSpace(snapshotPath))
		{
			return;
		}
		Stopwatch sw = Stopwatch.StartNew();
		try
		{
			string sourceKey = BuildSourceKey(snapshot.SourceId, snapshotPath);
			string sourceFolder = GetManagedSourceFolder(workspaceRoot, sourceKey);
			string revision = BuildSnapshotRevision(snapshot, snapshotPath);
			string revisionFolder = Path.Combine(sourceFolder, revision);
			string detailsFolder = Path.Combine(revisionFolder, "details");
			Directory.CreateDirectory(detailsFolder);

			StandardBrowserIndexV2 index = BuildIndex(snapshot);
			StandardBrowserDetailsCatalogV2 detailsCatalog = new StandardBrowserDetailsCatalogV2 { SourceId = snapshot.SourceId ?? string.Empty };
			foreach (BrowserIndexItem item in index.Items)
			{
				BrowserDetailRecord detail = BuildDetailRecord(snapshot, item);
				if (detail == null)
				{
					continue;
				}
				string detailPath = Path.Combine(detailsFolder, item.ItemId + ".json");
				TryWriteJsonAtomic(detailPath, detail);
				item.DetailKey = MakeRelativePath(sourceFolder, detailPath);
				detailsCatalog.Items.Add(new BrowserDetailCatalogEntryV2
				{
					ItemId = item.ItemId,
					DetailKey = item.DetailKey,
					Sha256 = HashFile(detailPath),
					Length = SafeLength(detailPath)
				});
			}

			string indexPath = Path.Combine(revisionFolder, IndexFileName);
			string detailsCatalogPath = Path.Combine(revisionFolder, DetailsCatalogFileName);
			TryWriteJsonAtomic(indexPath, index);
			TryWriteJsonAtomic(detailsCatalogPath, detailsCatalog);
			ThumbnailIndexV2 thumbnailIndex = BuildThumbnailIndex(workspaceRoot, snapshot.SourceId, revision);
			string thumbnailIndexPath = Path.Combine(revisionFolder, ThumbnailIndexFileName);
			TryWriteJsonAtomic(thumbnailIndexPath, thumbnailIndex);
			RegisterThumbnailIndex(thumbnailIndex);

			using (FileStream manifestLock = AcquireManifestMutationLock(sourceFolder))
			{
				FamilyBrowserManifestV2 previous = TryLoad<FamilyBrowserManifestV2>(Path.Combine(sourceFolder, ManifestFileName));
				FamilyBrowserManifestV2 manifest = new FamilyBrowserManifestV2
				{
					SourceKey = sourceKey,
					SourceId = snapshot.SourceId ?? string.Empty,
					SourceSnapshotPath = snapshotPath,
					SourceSnapshotRevision = revision,
					SourceSnapshotFileLastWriteUtc = File.Exists(snapshotPath) ? File.GetLastWriteTimeUtc(snapshotPath).ToString("O", CultureInfo.InvariantCulture) : string.Empty,
					SourceSnapshotFileLength = SafeLength(snapshotPath),
					SourceFileLastWriteUtc = snapshot.SourceFileLastWriteUtc ?? string.Empty,
					SourceFileLength = snapshot.SourceFileLength,
					SnapshotCapturedAtUtc = snapshot.CapturedAtUtc ?? string.Empty,
					StandardListRevision = previous == null ? string.Empty : previous.StandardListRevision,
					ProjectScanRevision = previous == null ? string.Empty : previous.ProjectScanRevision,
					PublishedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
					StandardIndex = BuildArtifactReference(sourceFolder, indexPath),
					StandardDetails = BuildArtifactReference(sourceFolder, detailsCatalogPath),
					StandardList = previous == null ? new FamilyBrowserArtifactReferenceV2() : previous.StandardList,
					ThumbnailIndex = BuildArtifactReference(sourceFolder, thumbnailIndexPath),
					ProjectState = previous == null ? new FamilyBrowserArtifactReferenceV2() : previous.ProjectState
				};
				TryWriteJsonAtomic(Path.Combine(sourceFolder, ManifestFileName), manifest);
				TryWriteJsonAtomic(Path.Combine(GetLocalSourceFolder(sourceKey), ManifestFileName), manifest);
			}
			sw.Stop();
			RecordPerformance("v2-artifacts-publish", sw.ElapsedMilliseconds, SafeLength(indexPath), index.Items.Count, false, sourceFolder, revision);
		}
		catch (Exception ex)
		{
			sw.Stop();
			RecordPerformance("v2-artifacts-publish-failed", sw.ElapsedMilliseconds, 0L, 0, false, snapshotPath, ex.Message);
		}
	}

	public static bool ValidateCurrentSourceRevision(string workspaceRoot, StandardLibraryRegistrationRecord registration, out string reason)
	{
		reason = string.Empty;
		if (registration == null)
		{
			reason = "Standard RVT registration is missing.";
			return false;
		}
		if (string.IsNullOrWhiteSpace(registration.LastSnapshotPath) || !File.Exists(registration.LastSnapshotPath))
		{
			reason = "The registered standard snapshot is missing. Run the standard scan again.";
			return false;
		}

		FamilyBrowserStandardRevisionState revisionState = FamilyBrowserStandardRevisionService.Probe(workspaceRoot, registration, true);
		if (revisionState == null || revisionState.BlocksStandardUse)
		{
			reason = revisionState == null ? "The Standard RVT revision could not be verified." : (string.IsNullOrWhiteSpace(revisionState.Reason) ? revisionState.ErrorMessage : revisionState.Reason);
			return false;
		}

		string sourcePath = !string.IsNullOrWhiteSpace(registration.ResolvedPath) ? registration.ResolvedPath : registration.Locator;
		if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
		{
			reason = "The registered Standard RVT source is unavailable. Model changes are disabled while the source cannot be verified.";
			return false;
		}

		long currentSourceLength = SafeLength(sourcePath);
		string currentSourceWriteUtc = File.GetLastWriteTimeUtc(sourcePath).ToString("O", CultureInfo.InvariantCulture);
		if ((registration.SourceFileLength > 0L && registration.SourceFileLength != currentSourceLength) ||
			!UtcFileStampMatches(registration.SourceFileLastWriteUtc, currentSourceWriteUtc))
		{
			reason = "The Standard RVT changed after the registered scan. Scan the standard again before loading or applying content.";
			return false;
		}

		string sourceKey = BuildSourceKey(registration.SourceId, registration.LastSnapshotPath);
		string manifestPath = Path.Combine(GetManagedSourceFolder(workspaceRoot, sourceKey), ManifestFileName);
		FamilyBrowserManifestV2 manifest = TryLoad<FamilyBrowserManifestV2>(manifestPath);
		if (manifest == null || manifest.SchemaVersion != SchemaVersion)
		{
			// Legacy registrations do not have a V2 manifest yet. The source stamp check above
			// still prevents a stale standard from mutating the current project.
			reason = "legacy-source-stamp-verified";
			return true;
		}

		if (!string.Equals(manifest.SourceKey ?? string.Empty, sourceKey, StringComparison.OrdinalIgnoreCase) ||
			!PathsEqual(manifest.SourceSnapshotPath, registration.LastSnapshotPath))
		{
			reason = "The browser manifest does not match the registered standard snapshot. Refresh the browser or scan the standard again.";
			return false;
		}
		if ((manifest.SourceFileLength > 0L && manifest.SourceFileLength != currentSourceLength) ||
			!UtcFileStampMatches(manifest.SourceFileLastWriteUtc, currentSourceWriteUtc))
		{
			reason = "The Standard RVT revision no longer matches the browser manifest. Scan the standard again before changing the model.";
			return false;
		}

		long currentSnapshotLength = SafeLength(registration.LastSnapshotPath);
		string currentSnapshotWriteUtc = File.GetLastWriteTimeUtc(registration.LastSnapshotPath).ToString("O", CultureInfo.InvariantCulture);
		if ((manifest.SourceSnapshotFileLength > 0L && manifest.SourceSnapshotFileLength != currentSnapshotLength) ||
			!UtcFileStampMatches(manifest.SourceSnapshotFileLastWriteUtc, currentSnapshotWriteUtc))
		{
			reason = "The registered standard snapshot changed after the browser manifest was published. Refresh or scan the standard again.";
			return false;
		}

		reason = "current-source-revision-verified";
		return true;
	}

	public static void PublishStandardList(string workspaceRoot, StandardLibraryRegistrationRecord registration, FamilyBrowserStandardListCatalog catalog)
	{
		if (registration == null || catalog == null || string.IsNullOrWhiteSpace(registration.LastSnapshotPath) || string.IsNullOrWhiteSpace(catalog.SourcePath) || !File.Exists(catalog.SourcePath))
		{
			return;
		}
		try
		{
			string sourceKey = BuildSourceKey(registration.SourceId, registration.LastSnapshotPath);
			string sourceFolder = GetManagedSourceFolder(workspaceRoot, sourceKey);
			using (FileStream manifestLock = AcquireManifestMutationLock(sourceFolder))
			{
				string manifestPath = Path.Combine(sourceFolder, ManifestFileName);
				FamilyBrowserManifestV2 manifest = TryLoad<FamilyBrowserManifestV2>(manifestPath);
				if (manifest == null || manifest.SchemaVersion != SchemaVersion)
				{
					return;
				}
				string revisionFolder = Path.Combine(sourceFolder, manifest.SourceSnapshotRevision ?? "current");
				string standardListPath = Path.Combine(revisionFolder, "standard-list-v2.json");
				CopyFileAtomic(catalog.SourcePath, standardListPath);
				manifest.StandardListRevision = BuildFileRevision(catalog.SourcePath);
				manifest.StandardList = BuildArtifactReference(sourceFolder, standardListPath);
				manifest.PublishedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
				TryWriteJsonAtomic(manifestPath, manifest);
				TryWriteJsonAtomic(Path.Combine(GetLocalSourceFolder(sourceKey), ManifestFileName), manifest);
			}
		}
		catch (Exception ex)
		{
			RecordPerformance("standard-list-v2-publish-failed", 0L, 0L, 0, false, catalog.SourcePath, ex.Message);
		}
	}

	public static void PublishProjectState(string workspaceRoot, StandardLibraryRegistrationRecord registration, ProjectScanCacheRecord record, ProjectContentSnapshot snapshot, ProjectStandardComparisonReport report)
	{
		if (registration == null || record == null || string.IsNullOrWhiteSpace(registration.LastSnapshotPath))
		{
			return;
		}
		try
		{
			string sourceKey = BuildSourceKey(registration.SourceId, registration.LastSnapshotPath);
			string sourceFolder = GetManagedSourceFolder(workspaceRoot, sourceKey);
			using (FileStream manifestLock = AcquireManifestMutationLock(sourceFolder))
			{
				string manifestPath = Path.Combine(sourceFolder, ManifestFileName);
				FamilyBrowserManifestV2 manifest = TryLoad<FamilyBrowserManifestV2>(manifestPath);
				if (manifest == null || manifest.SchemaVersion != SchemaVersion)
				{
					return;
				}
				string projectKey = string.IsNullOrWhiteSpace(record.ProjectKey) ? "unknown" : HashText(record.ProjectKey).Substring(0, 32);
				string projectFolder = Path.Combine(sourceFolder, manifest.SourceSnapshotRevision ?? "current", "projects", projectKey);
				string projectStatePath = Path.Combine(projectFolder, ProjectStateFileName);
				ProjectBrowserStateV2 state = new ProjectBrowserStateV2
				{
					ProjectKey = record.ProjectKey ?? string.Empty,
					ProjectTitle = record.ProjectTitle ?? string.Empty,
					ProjectDocumentPath = record.ProjectDocumentPath ?? string.Empty,
					ProjectRevision = (record.ProjectFileLastWriteUtc ?? string.Empty) + "|" + record.ProjectFileLength.ToString(CultureInfo.InvariantCulture),
					StandardSourceId = registration.SourceId ?? string.Empty,
					StandardSnapshotRevision = manifest.SourceSnapshotRevision ?? string.Empty,
					ProjectSnapshotPath = record.ProjectSnapshotPath ?? string.Empty,
					ComparisonReportPath = record.ComparisonReportPath ?? string.Empty,
					SavedAtUtc = record.SavedAtUtc ?? DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
					LoadableFamilyCount = report == null || report.LoadableFamilies == null ? 0 : report.LoadableFamilies.Count,
					SystemTypeCount = report == null || report.SystemTypes == null ? 0 : report.SystemTypes.Count
				};
				TryWriteJsonAtomic(projectStatePath, state);
				manifest.ProjectScanRevision = state.ProjectRevision;
				manifest.ProjectState = BuildArtifactReference(sourceFolder, projectStatePath);
				manifest.PublishedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
				TryWriteJsonAtomic(manifestPath, manifest);
				TryWriteJsonAtomic(Path.Combine(GetLocalSourceFolder(sourceKey), ManifestFileName), manifest);
			}
		}
		catch (Exception ex)
		{
			RecordPerformance("project-state-v2-publish-failed", 0L, 0L, 0, false, record.ProjectDocumentPath, ex.Message);
		}
	}

	public static FamilyBrowserStandardListCatalog PrepareStandardListCatalog(FamilyBrowserStandardLibrarySlot slot, StandardLibraryRegistrationRecord registration)
	{
		string key = BuildCatalogKey(slot, registration);
		lock (PreparedCatalogLock)
		{
			FamilyBrowserStandardListCatalog cached;
			if (PreparedCatalogs.TryGetValue(key, out cached) && cached != null)
			{
				return cached;
			}
		}
		FamilyBrowserStandardListCatalog catalog = FamilyBrowserStandardListService.LoadForSlot(slot, registration);
		lock (PreparedCatalogLock)
		{
			PreparedCatalogs[key] = catalog;
		}
		return catalog;
	}

	public static bool TryGetPreparedStandardListCatalog(FamilyBrowserStandardLibrarySlot slot, StandardLibraryRegistrationRecord registration, out FamilyBrowserStandardListCatalog catalog)
	{
		lock (PreparedCatalogLock)
		{
			return PreparedCatalogs.TryGetValue(BuildCatalogKey(slot, registration), out catalog) && catalog != null;
		}
	}

	public static void ClearPreparedCaches()
	{
		lock (PreparedCatalogLock)
		{
			PreparedCatalogs.Clear();
		}
		lock (ThumbnailLookupLock)
		{
			ThumbnailLookups.Clear();
			ThumbnailLookupBuilds.Clear();
		}
	}

	public static string ResolveThumbnailPath(string workspaceRoot, string sourceId, string categoryName, string familyName, string expectedPath)
	{
		if (string.IsNullOrWhiteSpace(sourceId) || string.IsNullOrWhiteSpace(familyName))
		{
			return string.Empty;
		}
		string lookupKey = Normalize(sourceId);
		Dictionary<string, string> lookup;
		lock (ThumbnailLookupLock)
		{
			ThumbnailLookups.TryGetValue(lookupKey, out lookup);
		}
		if (lookup == null)
		{
			lookup = BuildThumbnailLookupOnDemand(workspaceRoot, sourceId);
		}
		string path;
		string exactKey = BuildThumbnailLookupKey(categoryName, familyName);
		if (lookup != null && lookup.TryGetValue(exactKey, out path))
		{
			return path;
		}
		string familyOnlyKey = BuildThumbnailLookupKey(string.Empty, familyName);
		if (lookup != null && lookup.TryGetValue(familyOnlyKey, out path))
		{
			return path;
		}
		return expectedPath ?? string.Empty;
	}

	public static string BuildSourceKey(string sourceId, string snapshotPath)
	{
		return HashText(Normalize(sourceId) + "|" + NormalizePath(snapshotPath)).Substring(0, 32);
	}

	public static void RecordPerformance(string stage, long elapsedMilliseconds, long bytes, int rowCount, bool cacheHit, string sourcePath, string detail)
	{
		try
		{
			FamilyBrowserPerformanceEvent item = new FamilyBrowserPerformanceEvent
			{
				Stage = stage ?? string.Empty,
				ElapsedMilliseconds = elapsedMilliseconds,
				Bytes = bytes,
				RowCount = rowCount,
				CacheHit = cacheHit,
				NetworkPath = IsNetworkPath(sourcePath),
				SourcePath = sourcePath ?? string.Empty,
				Detail = detail ?? string.Empty
			};
			string folder = Path.Combine(GetLocalCacheRoot(), "diagnostics");
			Directory.CreateDirectory(folder);
			string path = Path.Combine(folder, "family-browser-performance-" + DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + ".jsonl");
			File.AppendAllText(path, PlainJsonReportWriter.Serialize(item) + Environment.NewLine, new UTF8Encoding(false));
		}
		catch
		{
		}
	}

	private static StandardBrowserIndexV2 BuildIndex(StandardLibrarySnapshot snapshot)
	{
		StandardBrowserIndexV2 index = new StandardBrowserIndexV2
		{
			SourceId = snapshot.SourceId ?? string.Empty,
			DisplayName = snapshot.DisplayName ?? string.Empty,
			SnapshotMode = snapshot.SnapshotMode ?? string.Empty,
			CapturedAtUtc = snapshot.CapturedAtUtc ?? string.Empty,
			SourceFileLastWriteUtc = snapshot.SourceFileLastWriteUtc ?? string.Empty,
			SourceFileLength = snapshot.SourceFileLength,
			RevitVersion = snapshot.RevitVersion ?? string.Empty
		};
		foreach (StandardLoadableFamilySnapshotItem family in snapshot.LoadableFamilies ?? new List<StandardLoadableFamilySnapshotItem>())
		{
			if (family == null)
			{
				continue;
			}
			string itemId = BuildItemId("family", family.CategoryName, family.FamilyName, string.Empty);
			index.Items.Add(new BrowserIndexItem
			{
				ItemId = itemId,
				ItemKind = "family",
				Name = family.FamilyName ?? string.Empty,
				CategoryName = family.CategoryName ?? string.Empty,
				CategoryId = family.CategoryId ?? string.Empty,
				CategoryGroup = family.CategoryGroup ?? string.Empty,
				TypeCount = family.TypeCount,
				Fingerprint = family.ContentFingerprint ?? string.Empty,
				ThumbnailKey = BuildThumbnailLookupKey(family.CategoryName, family.FamilyName),
				IsNestedLoadableChild = family.IsNestedLoadableChild
			});
		}
		foreach (StandardSystemTypeSnapshotItem systemType in snapshot.SystemTypes ?? new List<StandardSystemTypeSnapshotItem>())
		{
			if (systemType == null)
			{
				continue;
			}
			index.Items.Add(new BrowserIndexItem
			{
				ItemId = BuildItemId("system", systemType.CategoryName, systemType.TypeName, systemType.TypeClassName),
				ItemKind = "system",
				Name = systemType.TypeName ?? string.Empty,
				CategoryName = systemType.CategoryName ?? string.Empty,
				CategoryId = systemType.CategoryId ?? string.Empty,
				TypeClassName = systemType.TypeClassName ?? string.Empty,
				TypeCount = 1,
				Fingerprint = systemType.SemanticFingerprint ?? string.Empty,
				SupportsRoutingDependencies = systemType.SupportsRoutingDependencies
			});
		}
		index.Items = index.Items.OrderBy(x => Normalize(x.ItemKind) + "|" + Normalize(x.CategoryName) + "|" + Normalize(x.Name), StringComparer.Ordinal).ToList();
		return index;
	}

	private static BrowserDetailRecord BuildDetailRecord(StandardLibrarySnapshot snapshot, BrowserIndexItem item)
	{
		if (item == null)
		{
			return null;
		}
		BrowserDetailRecord detail = new BrowserDetailRecord
		{
			ItemId = item.ItemId,
			ItemKind = item.ItemKind,
			Name = item.Name,
			CategoryName = item.CategoryName
		};
		if (string.Equals(item.ItemKind, "family", StringComparison.OrdinalIgnoreCase))
		{
			detail.Family = (snapshot.LoadableFamilies ?? new List<StandardLoadableFamilySnapshotItem>()).FirstOrDefault(x => x != null && string.Equals(BuildItemId("family", x.CategoryName, x.FamilyName, string.Empty), item.ItemId, StringComparison.Ordinal));
		}
		else
		{
			detail.SystemType = (snapshot.SystemTypes ?? new List<StandardSystemTypeSnapshotItem>()).FirstOrDefault(x => x != null && string.Equals(BuildItemId("system", x.CategoryName, x.TypeName, x.TypeClassName), item.ItemId, StringComparison.Ordinal));
		}
		return detail;
	}

	private static StandardLibrarySnapshot BuildProjection(StandardBrowserIndexV2 index, StandardLibraryRegistrationRecord registration, string manifestSourceKey)
	{
		StandardLibrarySnapshot snapshot = new StandardLibrarySnapshot
		{
			SourceId = index.SourceId ?? registration.SourceId ?? string.Empty,
			DisplayName = index.DisplayName ?? registration.DisplayName ?? string.Empty,
			SnapshotMode = index.SnapshotMode ?? registration.SnapshotMode ?? string.Empty,
			SourceFileLastWriteUtc = index.SourceFileLastWriteUtc ?? registration.SourceFileLastWriteUtc ?? string.Empty,
			SourceFileLength = index.SourceFileLength,
			CapturedAtUtc = index.CapturedAtUtc ?? registration.LastSnapshotAtUtc ?? string.Empty,
			RevitVersion = index.RevitVersion ?? registration.RevitVersion ?? string.Empty,
			ResolvedPath = registration.ResolvedPath ?? string.Empty,
			Locator = registration.Locator ?? string.Empty,
				SourceKind = registration.SourceKind ?? string.Empty,
				BrowserManifestSourceKey = manifestSourceKey ?? string.Empty
		};
		foreach (BrowserIndexItem item in index.Items ?? new List<BrowserIndexItem>())
		{
			if (string.Equals(item.ItemKind, "family", StringComparison.OrdinalIgnoreCase))
			{
				snapshot.LoadableFamilies.Add(new StandardLoadableFamilySnapshotItem
				{
					FamilyName = item.Name ?? string.Empty,
					CategoryName = item.CategoryName ?? string.Empty,
					CategoryId = item.CategoryId ?? string.Empty,
					CategoryGroup = item.CategoryGroup ?? string.Empty,
					TypeCount = item.TypeCount,
					ContentFingerprint = item.Fingerprint ?? string.Empty,
					IsNestedLoadableChild = item.IsNestedLoadableChild,
					MetadataMode = "BrowserIndexV2",
					BrowserDetailKey = item.DetailKey ?? string.Empty
				});
			}
			else if (string.Equals(item.ItemKind, "system", StringComparison.OrdinalIgnoreCase))
			{
				snapshot.SystemTypes.Add(new StandardSystemTypeSnapshotItem
				{
					TypeName = item.Name ?? string.Empty,
					CategoryName = item.CategoryName ?? string.Empty,
					CategoryId = item.CategoryId ?? string.Empty,
					TypeClassName = item.TypeClassName ?? string.Empty,
					SemanticFingerprint = item.Fingerprint ?? string.Empty,
					SupportsRoutingDependencies = item.SupportsRoutingDependencies,
					BrowserDetailKey = item.DetailKey ?? string.Empty
				});
			}
		}
		snapshot.Summary = new StandardLibrarySnapshotSummary
		{
			LoadableFamilyCount = snapshot.LoadableFamilies.Count,
			LoadableTypeCount = snapshot.LoadableFamilies.Sum(x => x == null ? 0 : x.TypeCount),
			SystemTypeCount = snapshot.SystemTypes.Count,
			SystemTypeClassCount = snapshot.SystemTypes.Select(x => Normalize(x.TypeClassName)).Distinct(StringComparer.Ordinal).Count()
		};
		return snapshot;
	}

	private static ThumbnailIndexV2 BuildThumbnailIndex(string workspaceRoot, string sourceId, string revision)
	{
		ThumbnailIndexV2 index = new ThumbnailIndexV2
		{
			SourceId = sourceId ?? string.Empty,
			SourceRevision = revision ?? string.Empty,
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
		};
		if (string.IsNullOrWhiteSpace(sourceId))
		{
			return index;
		}
		string sourceFolder = Path.Combine(FamilyBrowserStandardPolicyStore.GetThumbnailFolder(workspaceRoot), SafeFileName(sourceId));
		index.ThumbnailRoot = sourceFolder;
		if (!Directory.Exists(sourceFolder))
		{
			return index;
		}
		foreach (string metadataPath in SafeEnumerateFiles(sourceFolder, "*.png.meta.json"))
		{
			FamilyThumbnailCacheMetadata metadata = TryLoad<FamilyThumbnailCacheMetadata>(metadataPath);
			if (metadata == null || string.IsNullOrWhiteSpace(metadata.FamilyName))
			{
				continue;
			}
			string imagePath = metadataPath.Substring(0, metadataPath.Length - ".meta.json".Length);
			if (!File.Exists(imagePath))
			{
				continue;
			}
			FileInfo imageInfo = new FileInfo(imagePath);
			index.Items.Add(new ThumbnailIndexEntry
			{
				ThumbnailKey = BuildThumbnailLookupKey(metadata.CategoryName, metadata.FamilyName),
				SourceId = metadata.SourceId ?? sourceId,
				CategoryName = metadata.CategoryName ?? string.Empty,
				FamilyName = metadata.FamilyName ?? string.Empty,
				RelativeImagePath = MakeRelativePath(sourceFolder, imagePath),
				MetadataLastWriteUtc = File.GetLastWriteTimeUtc(metadataPath).ToString("O", CultureInfo.InvariantCulture),
				ImageLength = imageInfo.Length
			});
		}
		return index;
	}

	private static Dictionary<string, string> BuildThumbnailLookupOnDemand(string workspaceRoot, string sourceId)
	{
		string sourceKey = Normalize(sourceId);
		lock (ThumbnailLookupLock)
		{
			Dictionary<string, string> cached;
			if (ThumbnailLookups.TryGetValue(sourceKey, out cached))
			{
				return cached;
			}
			if (ThumbnailLookupBuilds.Contains(sourceKey))
			{
				return new Dictionary<string, string>(StringComparer.Ordinal);
			}
			ThumbnailLookupBuilds.Add(sourceKey);
		}
		Stopwatch sw = Stopwatch.StartNew();
		ThumbnailIndexV2 index = BuildThumbnailIndex(workspaceRoot, sourceId, "on-demand");
		RegisterThumbnailIndex(index);
		sw.Stop();
		RecordPerformance("thumbnail-index-build", sw.ElapsedMilliseconds, 0L, index.Items.Count, false, index.ThumbnailRoot, "metadata-enumeration=1");
		lock (ThumbnailLookupLock)
		{
			Dictionary<string, string> result;
			return ThumbnailLookups.TryGetValue(sourceKey, out result) ? result : new Dictionary<string, string>(StringComparer.Ordinal);
		}
	}

	private static void RegisterThumbnailIndex(string workspaceRoot, string managedSourceFolder, string localSourceFolder, FamilyBrowserManifestV2 manifest, bool remoteAvailable)
	{
		if (manifest == null || manifest.ThumbnailIndex == null || string.IsNullOrWhiteSpace(manifest.ThumbnailIndex.RelativePath))
		{
			return;
		}
		string path;
		bool localCacheHit;
		if (TryResolveArtifact(managedSourceFolder, localSourceFolder, manifest.ThumbnailIndex, remoteAvailable, out path, out localCacheHit))
		{
			ThumbnailIndexV2 index = TryLoad<ThumbnailIndexV2>(path);
			if (index != null)
			{
				RegisterThumbnailIndex(index);
			}
		}
	}

	private static void RegisterThumbnailIndex(ThumbnailIndexV2 index)
	{
		if (index == null || string.IsNullOrWhiteSpace(index.SourceId))
		{
			return;
		}
		Dictionary<string, string> lookup = new Dictionary<string, string>(StringComparer.Ordinal);
		foreach (ThumbnailIndexEntry item in index.Items ?? new List<ThumbnailIndexEntry>())
		{
			if (item == null || string.IsNullOrWhiteSpace(item.FamilyName))
			{
				continue;
			}
			string imagePath = Path.Combine(index.ThumbnailRoot ?? string.Empty, item.RelativeImagePath ?? string.Empty);
			lookup[BuildThumbnailLookupKey(item.CategoryName, item.FamilyName)] = imagePath;
			string familyOnly = BuildThumbnailLookupKey(string.Empty, item.FamilyName);
			if (!lookup.ContainsKey(familyOnly))
			{
				lookup[familyOnly] = imagePath;
			}
		}
		lock (ThumbnailLookupLock)
		{
			ThumbnailLookups[Normalize(index.SourceId)] = lookup;
			ThumbnailLookupBuilds.Add(Normalize(index.SourceId));
		}
	}

	private static bool TryResolveArtifact(string managedSourceFolder, string localSourceFolder, FamilyBrowserArtifactReferenceV2 reference, bool remoteAvailable, out string localPath, out bool localCacheHit)
	{
		localPath = string.Empty;
		localCacheHit = false;
		if (reference == null || string.IsNullOrWhiteSpace(reference.RelativePath))
		{
			return false;
		}
		localPath = Path.Combine(localSourceFolder, NormalizeRelativePath(reference.RelativePath));
		if (File.Exists(localPath) && (string.IsNullOrWhiteSpace(reference.Sha256) || string.Equals(HashFile(localPath), reference.Sha256, StringComparison.OrdinalIgnoreCase)))
		{
			localCacheHit = true;
			return true;
		}
		if (!remoteAvailable)
		{
			return false;
		}
		string remotePath = Path.Combine(managedSourceFolder, NormalizeRelativePath(reference.RelativePath));
		if (!File.Exists(remotePath))
		{
			return false;
		}
		try
		{
			CopyFileAtomic(remotePath, localPath);
			if (!string.IsNullOrWhiteSpace(reference.Sha256) && !string.Equals(HashFile(localPath), reference.Sha256, StringComparison.OrdinalIgnoreCase))
			{
				TryDelete(localPath);
				return false;
			}
			return true;
		}
		catch
		{
			return false;
		}
	}

	private static FamilyBrowserArtifactReferenceV2 BuildArtifactReference(string sourceFolder, string path)
	{
		return new FamilyBrowserArtifactReferenceV2
		{
			RelativePath = MakeRelativePath(sourceFolder, path),
			Sha256 = HashFile(path),
			Length = SafeLength(path),
			LastWriteUtc = File.Exists(path) ? File.GetLastWriteTimeUtc(path).ToString("O", CultureInfo.InvariantCulture) : string.Empty
		};
	}

	private static string GetManagedSourceFolder(string workspaceRoot, string sourceKey)
	{
		string root = FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "BrowserCacheV2");
		return Path.Combine(root, sourceKey ?? "unknown");
	}

	private static string GetLocalCacheRoot()
	{
		string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
		return Path.Combine(local, "KKY", "FamilyBrowser", "Cache", "v2");
	}

	private static string GetLocalSourceFolder(string sourceKey)
	{
		return Path.Combine(GetLocalCacheRoot(), "sources", sourceKey ?? "unknown");
	}

	private static string BuildCatalogKey(FamilyBrowserStandardLibrarySlot slot, StandardLibraryRegistrationRecord registration)
	{
		return Normalize(slot == null ? string.Empty : slot.SlotKey) + "|" + NormalizePath(slot == null ? string.Empty : slot.StandardListPath) + "|" + Normalize(slot == null ? string.Empty : slot.StandardListSheetName) + "|" + Normalize(registration == null ? string.Empty : registration.SourceId);
	}

	private static string BuildSnapshotRevision(StandardLibrarySnapshot snapshot, string snapshotPath)
	{
		string raw = (snapshot.SourceId ?? string.Empty) + "|" + (snapshot.SourceFileLastWriteUtc ?? string.Empty) + "|" + snapshot.SourceFileLength.ToString(CultureInfo.InvariantCulture) + "|" + (snapshot.CapturedAtUtc ?? string.Empty) + "|" + BuildFileRevision(snapshotPath);
		return HashText(raw).Substring(0, 24);
	}

	private static string BuildFileRevision(string path)
	{
		try
		{
			FileInfo info = new FileInfo(path);
			return info.Length.ToString(CultureInfo.InvariantCulture) + "|" + info.LastWriteTimeUtc.Ticks.ToString(CultureInfo.InvariantCulture);
		}
		catch
		{
			return "missing";
		}
	}

	private static string BuildItemId(string kind, string category, string name, string typeClass)
	{
		return HashText(Normalize(kind) + "|" + Normalize(category) + "|" + Normalize(name) + "|" + Normalize(typeClass)).Substring(0, 40);
	}

	private static string BuildThumbnailLookupKey(string categoryName, string familyName)
	{
		return Normalize(categoryName) + "|" + Normalize(familyName);
	}

	private static string HashText(string value)
	{
		using (SHA256 sha = SHA256.Create())
		{
			byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty));
			StringBuilder sb = new StringBuilder(hash.Length * 2);
			foreach (byte b in hash)
			{
				sb.Append(b.ToString("x2", CultureInfo.InvariantCulture));
			}
			return sb.ToString();
		}
	}

	private static string HashFile(string path)
	{
		if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
		{
			return string.Empty;
		}
		using (SHA256 sha = SHA256.Create())
		using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
		{
			byte[] hash = sha.ComputeHash(stream);
			StringBuilder sb = new StringBuilder(hash.Length * 2);
			foreach (byte b in hash)
			{
				sb.Append(b.ToString("x2", CultureInfo.InvariantCulture));
			}
			return sb.ToString();
		}
	}

	private static T TryLoad<T>(string path) where T : class
	{
		try
		{
			return string.IsNullOrWhiteSpace(path) || !File.Exists(path) ? null : DataContractJsonFileStore.Load<T>(path);
		}
		catch
		{
			return null;
		}
	}

	private static void TryWriteJsonAtomic(string path, object value)
	{
		if (string.IsNullOrWhiteSpace(path) || value == null)
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
			File.WriteAllText(tempPath, PlainJsonReportWriter.Serialize(value), new UTF8Encoding(false));
			ReplaceFileAtomic(tempPath, path);
		}
		finally
		{
			TryDelete(tempPath);
		}
	}

	private static FileStream AcquireManifestMutationLock(string sourceFolder)
	{
		Directory.CreateDirectory(sourceFolder);
		string lockPath = Path.Combine(sourceFolder, ManifestMutationLockFileName);
		Stopwatch wait = Stopwatch.StartNew();
		while (true)
		{
			try
			{
				return new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.WriteThrough);
			}
			catch (IOException)
			{
				if (wait.ElapsedMilliseconds >= 15000L)
				{
					throw new IOException("Timed out waiting for the Family Browser manifest publication lock: " + lockPath);
				}
				Thread.Sleep(100);
			}
		}
	}

	private static void CopyFileAtomic(string sourcePath, string destinationPath)
	{
		string directory = Path.GetDirectoryName(destinationPath);
		if (!string.IsNullOrWhiteSpace(directory))
		{
			Directory.CreateDirectory(directory);
		}
		string tempPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(destinationPath);
		try
		{
			File.Copy(sourcePath, tempPath, true);
			ReplaceFileAtomic(tempPath, destinationPath);
		}
		finally
		{
			TryDelete(tempPath);
		}
	}

	private static void ReplaceFileAtomic(string tempPath, string destinationPath)
	{
		FamilyBrowserAtomicFileService.Promote(tempPath, destinationPath);
	}

	private static IEnumerable<string> SafeEnumerateFiles(string folder, string pattern)
	{
		try
		{
			return Directory.EnumerateFiles(folder, pattern, SearchOption.AllDirectories).ToList();
		}
		catch
		{
			return new List<string>();
		}
	}

	private static string MakeRelativePath(string root, string path)
	{
		try
		{
			Uri rootUri = new Uri(AppendDirectorySeparator(Path.GetFullPath(root)));
			Uri pathUri = new Uri(Path.GetFullPath(path));
			return Uri.UnescapeDataString(rootUri.MakeRelativeUri(pathUri).ToString()).Replace('/', Path.DirectorySeparatorChar);
		}
		catch
		{
			return Path.GetFileName(path) ?? string.Empty;
		}
	}

	private static string AppendDirectorySeparator(string path)
	{
		if (string.IsNullOrWhiteSpace(path) || path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
		{
			return path;
		}
		return path + Path.DirectorySeparatorChar;
	}

	private static string NormalizeRelativePath(string path)
	{
		return (path ?? string.Empty).Replace('/', Path.DirectorySeparatorChar).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
	}

	private static bool PathsEqual(string left, string right)
	{
		try
		{
			return string.Equals(Path.GetFullPath(left ?? string.Empty).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), Path.GetFullPath(right ?? string.Empty).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), StringComparison.OrdinalIgnoreCase);
		}
		catch
		{
			return string.Equals((left ?? string.Empty).Trim(), (right ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase);
		}
	}

	private static bool UtcFileStampMatches(string expected, string actual)
	{
		if (string.IsNullOrWhiteSpace(expected))
		{
			return true;
		}
		DateTime expectedUtc;
		DateTime actualUtc;
		if (DateTime.TryParse(expected, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out expectedUtc) &&
			DateTime.TryParse(actual, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out actualUtc))
		{
			return Math.Abs((expectedUtc.ToUniversalTime() - actualUtc.ToUniversalTime()).TotalSeconds) <= 2.0;
		}
		return string.Equals(expected.Trim(), (actual ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase);
	}

	private static string NormalizePath(string path)
	{
		if (string.IsNullOrWhiteSpace(path))
		{
			return string.Empty;
		}
		try
		{
			return Path.GetFullPath(Environment.ExpandEnvironmentVariables(path.Trim())).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).ToLowerInvariant();
		}
		catch
		{
			return path.Trim().ToLowerInvariant();
		}
	}

	private static string Normalize(string value)
	{
		return (value ?? string.Empty).Trim().ToLowerInvariant();
	}

	private static string SafeFileName(string value)
	{
		string text = value ?? string.Empty;
		foreach (char invalid in Path.GetInvalidFileNameChars())
		{
			text = text.Replace(invalid, '_');
		}
		return string.IsNullOrWhiteSpace(text) ? "unknown" : text.Trim();
	}

	private static long SafeLength(string path)
	{
		try
		{
			return string.IsNullOrWhiteSpace(path) || !File.Exists(path) ? 0L : new FileInfo(path).Length;
		}
		catch
		{
			return 0L;
		}
	}

	private static void TryDelete(string path)
	{
		try
		{
			if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
			{
				File.Delete(path);
			}
		}
		catch
		{
		}
	}

	private static bool IsNetworkPath(string path)
	{
		if (string.IsNullOrWhiteSpace(path))
		{
			return false;
		}
		try
		{
			string fullPath = Path.GetFullPath(path);
			if (fullPath.StartsWith("\\\\", StringComparison.Ordinal))
			{
				return true;
			}
			string root = Path.GetPathRoot(fullPath);
			if (string.IsNullOrWhiteSpace(root))
			{
				return false;
			}
			DriveInfo drive = new DriveInfo(root);
			return drive.DriveType == DriveType.Network;
		}
		catch
		{
			return path.StartsWith("\\\\", StringComparison.Ordinal);
		}
	}
}
