using System;
using System.Collections.Generic;
using System.Threading;

public interface IFamilyBrowserDataLoader
{
	FamilyBrowserCompactSnapshotLoadResult LoadSnapshotProjection(string workspaceRoot, StandardLibraryRegistrationRecord registration, bool allowSourceFallback);

	BrowserDetailRecord LoadDetail(string workspaceRoot, FamilyBrowserManifestV2 manifest, string detailKey);

	bool TryLoadRowCache(string cacheKey, out FamilyBrowserRowCacheV2 cache);

	void SaveRowCache(string cacheKey, FamilyBrowserRowCacheV2 cache);
}

public sealed class FamilyBrowserArtifactReferenceV2
{
	public string RelativePath { get; set; }

	public string Sha256 { get; set; }

	public long Length { get; set; }

	public string LastWriteUtc { get; set; }

	public FamilyBrowserArtifactReferenceV2()
	{
		RelativePath = string.Empty;
		Sha256 = string.Empty;
		LastWriteUtc = string.Empty;
	}
}

public sealed class FamilyBrowserManifestV2
{
	public int SchemaVersion { get; set; }

	public string SourceKey { get; set; }

	public string SourceId { get; set; }

	public string SourceSnapshotPath { get; set; }

	public string SourceSnapshotRevision { get; set; }

	public string SourceSnapshotFileLastWriteUtc { get; set; }

	public long SourceSnapshotFileLength { get; set; }

	public string SourceFileLastWriteUtc { get; set; }

	public long SourceFileLength { get; set; }

	public string SnapshotCapturedAtUtc { get; set; }

	public string StandardListRevision { get; set; }

	public string ProjectScanRevision { get; set; }

	public string PublishedAtUtc { get; set; }

	public FamilyBrowserArtifactReferenceV2 StandardIndex { get; set; }

	public FamilyBrowserArtifactReferenceV2 StandardDetails { get; set; }

	public FamilyBrowserArtifactReferenceV2 StandardList { get; set; }

	public FamilyBrowserArtifactReferenceV2 ThumbnailIndex { get; set; }

	public FamilyBrowserArtifactReferenceV2 ProjectState { get; set; }

	public FamilyBrowserManifestV2()
	{
		SchemaVersion = 2;
		SourceKey = string.Empty;
		SourceId = string.Empty;
		SourceSnapshotPath = string.Empty;
		SourceSnapshotRevision = string.Empty;
		SourceSnapshotFileLastWriteUtc = string.Empty;
		SourceFileLastWriteUtc = string.Empty;
		SnapshotCapturedAtUtc = string.Empty;
		StandardListRevision = string.Empty;
		ProjectScanRevision = string.Empty;
		PublishedAtUtc = string.Empty;
		StandardIndex = new FamilyBrowserArtifactReferenceV2();
		StandardDetails = new FamilyBrowserArtifactReferenceV2();
		StandardList = new FamilyBrowserArtifactReferenceV2();
		ThumbnailIndex = new FamilyBrowserArtifactReferenceV2();
		ProjectState = new FamilyBrowserArtifactReferenceV2();
	}
}

public sealed class BrowserIndexItem
{
	public string ItemId { get; set; }

	public string ItemKind { get; set; }

	public string Name { get; set; }

	public string CategoryName { get; set; }

	public string CategoryId { get; set; }

	public string CategoryGroup { get; set; }

	public string TypeClassName { get; set; }

	public int TypeCount { get; set; }

	public string Fingerprint { get; set; }

	public string DetailKey { get; set; }

	public string ThumbnailKey { get; set; }

	public bool IsNestedLoadableChild { get; set; }

	public bool SupportsRoutingDependencies { get; set; }

	public BrowserIndexItem()
	{
		ItemId = string.Empty;
		ItemKind = string.Empty;
		Name = string.Empty;
		CategoryName = string.Empty;
		CategoryId = string.Empty;
		CategoryGroup = string.Empty;
		TypeClassName = string.Empty;
		Fingerprint = string.Empty;
		DetailKey = string.Empty;
		ThumbnailKey = string.Empty;
	}
}

public sealed class StandardBrowserIndexV2
{
	public int SchemaVersion { get; set; }

	public string SourceId { get; set; }

	public string DisplayName { get; set; }

	public string SnapshotMode { get; set; }

	public string CapturedAtUtc { get; set; }

	public string SourceFileLastWriteUtc { get; set; }

	public long SourceFileLength { get; set; }

	public string RevitVersion { get; set; }

	public List<BrowserIndexItem> Items { get; set; }

	public StandardBrowserIndexV2()
	{
		SchemaVersion = 2;
		SourceId = string.Empty;
		DisplayName = string.Empty;
		SnapshotMode = string.Empty;
		CapturedAtUtc = string.Empty;
		SourceFileLastWriteUtc = string.Empty;
		RevitVersion = string.Empty;
		Items = new List<BrowserIndexItem>();
	}
}

public sealed class BrowserDetailRecord
{
	public int SchemaVersion { get; set; }

	public string ItemId { get; set; }

	public string ItemKind { get; set; }

	public string Name { get; set; }

	public string CategoryName { get; set; }

	public StandardLoadableFamilySnapshotItem Family { get; set; }

	public StandardSystemTypeSnapshotItem SystemType { get; set; }

	public BrowserDetailRecord()
	{
		SchemaVersion = 2;
		ItemId = string.Empty;
		ItemKind = string.Empty;
		Name = string.Empty;
		CategoryName = string.Empty;
	}
}

public sealed class BrowserDetailCatalogEntryV2
{
	public string ItemId { get; set; }

	public string DetailKey { get; set; }

	public string Sha256 { get; set; }

	public long Length { get; set; }

	public BrowserDetailCatalogEntryV2()
	{
		ItemId = string.Empty;
		DetailKey = string.Empty;
		Sha256 = string.Empty;
	}
}

public sealed class StandardBrowserDetailsCatalogV2
{
	public int SchemaVersion { get; set; }

	public string SourceId { get; set; }

	public List<BrowserDetailCatalogEntryV2> Items { get; set; }

	public StandardBrowserDetailsCatalogV2()
	{
		SchemaVersion = 2;
		SourceId = string.Empty;
		Items = new List<BrowserDetailCatalogEntryV2>();
	}
}

public sealed class ThumbnailIndexEntry
{
	public string ThumbnailKey { get; set; }

	public string SourceId { get; set; }

	public string CategoryName { get; set; }

	public string FamilyName { get; set; }

	public string RelativeImagePath { get; set; }

	public string MetadataLastWriteUtc { get; set; }

	public long ImageLength { get; set; }

	public ThumbnailIndexEntry()
	{
		ThumbnailKey = string.Empty;
		SourceId = string.Empty;
		CategoryName = string.Empty;
		FamilyName = string.Empty;
		RelativeImagePath = string.Empty;
		MetadataLastWriteUtc = string.Empty;
	}
}

public sealed class ThumbnailIndexV2
{
	public int SchemaVersion { get; set; }

	public string SourceId { get; set; }

	public string SourceRevision { get; set; }

	public string ThumbnailRoot { get; set; }

	public string GeneratedAtUtc { get; set; }

	public List<ThumbnailIndexEntry> Items { get; set; }

	public ThumbnailIndexV2()
	{
		SchemaVersion = 2;
		SourceId = string.Empty;
		SourceRevision = string.Empty;
		ThumbnailRoot = string.Empty;
		GeneratedAtUtc = string.Empty;
		Items = new List<ThumbnailIndexEntry>();
	}
}

public sealed class ProjectBrowserStateV2
{
	public int SchemaVersion { get; set; }

	public string ProjectKey { get; set; }

	public string ProjectTitle { get; set; }

	public string ProjectDocumentPath { get; set; }

	public string ProjectRevision { get; set; }

	public string StandardSourceId { get; set; }

	public string StandardSnapshotRevision { get; set; }

	public string ProjectSnapshotPath { get; set; }

	public string ComparisonReportPath { get; set; }

	public string SavedAtUtc { get; set; }

	public int LoadableFamilyCount { get; set; }

	public int SystemTypeCount { get; set; }

	public ProjectBrowserStateV2()
	{
		SchemaVersion = 2;
		ProjectKey = string.Empty;
		ProjectTitle = string.Empty;
		ProjectDocumentPath = string.Empty;
		ProjectRevision = string.Empty;
		StandardSourceId = string.Empty;
		StandardSnapshotRevision = string.Empty;
		ProjectSnapshotPath = string.Empty;
		ComparisonReportPath = string.Empty;
		SavedAtUtc = string.Empty;
	}
}

public sealed class FamilyBrowserCompactSnapshotLoadResult
{
	public bool Success { get; set; }

	public bool LocalCacheHit { get; set; }

	public bool SourceFallbackUsed { get; set; }

	public bool OfflineCacheUsed { get; set; }

	public string Reason { get; set; }

	public string Revision { get; set; }

	public string SourcePath { get; set; }

	public long BytesRead { get; set; }

	public long ElapsedMilliseconds { get; set; }

	public StandardLibrarySnapshot Snapshot { get; set; }

	public FamilyBrowserManifestV2 Manifest { get; set; }

	public StandardBrowserIndexV2 Index { get; set; }

	public FamilyBrowserCompactSnapshotLoadResult()
	{
		Reason = string.Empty;
		Revision = string.Empty;
		SourcePath = string.Empty;
	}
}

public sealed class FamilyBrowserCachedFamilyRowV2
{
	public string Status { get; set; }
	public string RawStatus { get; set; }
	public string DisciplineKey { get; set; }
	public string DisciplineLabel { get; set; }
	public string Name { get; set; }
	public string Category { get; set; }
	public string CategoryGroup { get; set; }
	public string Action { get; set; }
	public string Notes { get; set; }
	public string ApprovedRev { get; set; }
	public string LoadedRev { get; set; }
	public string ChangeSummary { get; set; }
	public string DifferenceSummaryTable { get; set; }
	public string TypeSummary { get; set; }
	public string ParameterSummary { get; set; }
	public string TypeParameterSummary { get; set; }
	public string NestedSummary { get; set; }
	public bool IsNestedLoadableChild { get; set; }
	public string PreviewImagePath { get; set; }
	public string PreviewDiagnostic { get; set; }
	public string DetailKey { get; set; }
	public string DetailSourceKey { get; set; }

	public FamilyBrowserCachedFamilyRowV2()
	{
		Status = string.Empty;
		RawStatus = string.Empty;
		DisciplineKey = string.Empty;
		DisciplineLabel = string.Empty;
		Name = string.Empty;
		Category = string.Empty;
		CategoryGroup = string.Empty;
		Action = string.Empty;
		Notes = string.Empty;
		ApprovedRev = string.Empty;
		LoadedRev = string.Empty;
		ChangeSummary = string.Empty;
		DifferenceSummaryTable = string.Empty;
		TypeSummary = string.Empty;
		ParameterSummary = string.Empty;
		TypeParameterSummary = string.Empty;
		NestedSummary = string.Empty;
		PreviewImagePath = string.Empty;
		PreviewDiagnostic = string.Empty;
		DetailKey = string.Empty;
		DetailSourceKey = string.Empty;
	}
}

public sealed class FamilyBrowserCachedSystemRowV2
{
	public string Status { get; set; }
	public string RawStatus { get; set; }
	public string DisciplineKey { get; set; }
	public string DisciplineLabel { get; set; }
	public string Name { get; set; }
	public string Category { get; set; }
	public string SystemFamilyKind { get; set; }
	public string Action { get; set; }
	public string Notes { get; set; }
	public string ParameterSummary { get; set; }
	public string LayerSummary { get; set; }
	public string DifferenceSummaryTable { get; set; }
	public string DetailKey { get; set; }
	public string DetailSourceKey { get; set; }

	public FamilyBrowserCachedSystemRowV2()
	{
		Status = string.Empty;
		RawStatus = string.Empty;
		DisciplineKey = string.Empty;
		DisciplineLabel = string.Empty;
		Name = string.Empty;
		Category = string.Empty;
		SystemFamilyKind = string.Empty;
		Action = string.Empty;
		Notes = string.Empty;
		ParameterSummary = string.Empty;
		LayerSummary = string.Empty;
		DifferenceSummaryTable = string.Empty;
		DetailKey = string.Empty;
		DetailSourceKey = string.Empty;
	}
}

public sealed class FamilyBrowserCachedUnregisteredRowV2
{
	public string ItemKind { get; set; }
	public string Name { get; set; }
	public string CategoryName { get; set; }
	public string TypeClassName { get; set; }
	public string Notes { get; set; }
	public string Source { get; set; }

	public FamilyBrowserCachedUnregisteredRowV2()
	{
		ItemKind = string.Empty;
		Name = string.Empty;
		CategoryName = string.Empty;
		TypeClassName = string.Empty;
		Notes = string.Empty;
		Source = string.Empty;
	}
}

public sealed class FamilyBrowserRowCacheV2
{
	public int SchemaVersion { get; set; }

	public string CacheKeyHash { get; set; }

	public string SavedAtUtc { get; set; }

	public bool UsedProjectScan { get; set; }

	public string UnregisteredListStatusText { get; set; }

	public List<FamilyBrowserCachedFamilyRowV2> Families { get; set; }

	public List<FamilyBrowserCachedSystemRowV2> Systems { get; set; }

	public List<FamilyBrowserCachedSystemRowV2> SystemComparisons { get; set; }

	public List<FamilyBrowserCachedUnregisteredRowV2> UnregisteredFamilies { get; set; }

	public List<FamilyBrowserCachedUnregisteredRowV2> UnregisteredSystems { get; set; }

	public FamilyBrowserRowCacheV2()
	{
		SchemaVersion = 2;
		CacheKeyHash = string.Empty;
		SavedAtUtc = string.Empty;
		UnregisteredListStatusText = string.Empty;
		Families = new List<FamilyBrowserCachedFamilyRowV2>();
		Systems = new List<FamilyBrowserCachedSystemRowV2>();
		SystemComparisons = new List<FamilyBrowserCachedSystemRowV2>();
		UnregisteredFamilies = new List<FamilyBrowserCachedUnregisteredRowV2>();
		UnregisteredSystems = new List<FamilyBrowserCachedUnregisteredRowV2>();
	}
}

public sealed class FamilyBrowserPreparedSlotData
{
	public string SlotKey { get; set; }

	public StandardLibraryRegistrationRecord Registration { get; set; }

	public StandardLibrarySnapshot SnapshotProjection { get; set; }

	public FamilyBrowserStandardListCatalog StandardListCatalog { get; set; }

	public FamilyBrowserManifestV2 Manifest { get; set; }

	public FamilyBrowserStandardRevisionState StandardRevisionState { get; set; }

	public ProjectScanCacheLoadResult ProjectScan { get; set; }

	public string ProjectScanCacheStamp { get; set; }

	public string StandardListCacheStamp { get; set; }

	public string Diagnostic { get; set; }

	public FamilyBrowserPreparedSlotData()
	{
		SlotKey = string.Empty;
		ProjectScanCacheStamp = string.Empty;
		StandardListCacheStamp = string.Empty;
		Diagnostic = string.Empty;
	}
}

public sealed class FamilyBrowserStartupPreloadResult
{
	public int Generation { get; set; }

	public FamilyBrowserStandardPolicy Policy { get; set; }

	public List<FamilyBrowserPreparedSlotData> Slots { get; set; }

	public string StartedAtUtc { get; set; }

	public string CompletedAtUtc { get; set; }

	public long ElapsedMilliseconds { get; set; }

	public int LocalCacheHitCount { get; set; }

	public int SourceFallbackCount { get; set; }

	public string ErrorMessage { get; set; }

	public FamilyBrowserStartupPreloadResult()
	{
		Slots = new List<FamilyBrowserPreparedSlotData>();
		StartedAtUtc = string.Empty;
		CompletedAtUtc = string.Empty;
		ErrorMessage = string.Empty;
	}
}

public sealed class FamilyBrowserLoadGeneration : IDisposable
{
	private readonly CancellationTokenSource _cancellation;

	public int Value { get; private set; }

	public CancellationToken Token
	{
		get { return _cancellation.Token; }
	}

	public FamilyBrowserLoadGeneration(int value)
	{
		Value = value;
		_cancellation = new CancellationTokenSource();
	}

	public void Cancel()
	{
		if (!_cancellation.IsCancellationRequested)
		{
			_cancellation.Cancel();
		}
	}

	public void Dispose()
	{
		_cancellation.Dispose();
	}
}

public sealed class FamilyBrowserPerformanceEvent
{
	public string TimeUtc { get; set; }
	public string Stage { get; set; }
	public long ElapsedMilliseconds { get; set; }
	public long Bytes { get; set; }
	public int RowCount { get; set; }
	public bool CacheHit { get; set; }
	public bool NetworkPath { get; set; }
	public string SourcePath { get; set; }
	public string Detail { get; set; }

	public FamilyBrowserPerformanceEvent()
	{
		TimeUtc = DateTime.UtcNow.ToString("O");
		Stage = string.Empty;
		SourcePath = string.Empty;
		Detail = string.Empty;
	}
}

public sealed class FamilyBrowserSyntheticPerformanceAuditResult
{
	public bool Success { get; set; }
	public int FamilyCount { get; set; }
	public int SystemCount { get; set; }
	public long CacheBytes { get; set; }
	public long SaveMilliseconds { get; set; }
	public long ColdLoadMilliseconds { get; set; }
	public long WarmLoadMilliseconds { get; set; }
	public long OfflineLoadMilliseconds { get; set; }
	public string ErrorMessage { get; set; }

	public FamilyBrowserSyntheticPerformanceAuditResult()
	{
		ErrorMessage = string.Empty;
	}
}
