using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyThumbnailPreviewService
{
	private class PreviewExportFileState
	{
		public string FullName { get; set; }

		public long Length { get; set; }

		public DateTime LastWriteTimeUtc { get; set; }

		public PreviewExportFileState()
		{
			FullName = string.Empty;
		}
	}

	private class PreviewBoundsResult
	{
		public BoundingBoxXYZ Bounds { get; set; }

		public BoundingBoxXYZ PhysicalBounds { get; set; }

		public BoundingBoxXYZ ConnectorBounds { get; set; }

		public bool ConnectorExtentsIncluded { get; set; }

		public bool ConnectorExtentsClamped { get; set; }
	}

	private class FamilyThumbnailStageException : Exception
	{
		public FamilyThumbnailStageException(string stageName, Exception innerException)
			: base("Snapshot failed at " + stageName + ": " + ((innerException == null) ? "Unknown error." : innerException.Message), innerException)
		{
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__19_002D0
	{
		public ISet<string> _0024VB_0024Local_selectedNameSet;

		public _Closure_0024__19_002D0(_Closure_0024__19_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_selectedNameSet = arg0._0024VB_0024Local_selectedNameSet;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__1(Family x)
		{
			if (_0024VB_0024Local_selectedNameSet.Count != 0)
			{
				return _0024VB_0024Local_selectedNameSet.Contains(Normalize(x.Name ?? string.Empty));
			}
			return true;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__20_002D0
	{
		public ISet<string> _0024VB_0024Local_selectedNameSet;

		public _Closure_0024__20_002D0(_Closure_0024__20_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_selectedNameSet = arg0._0024VB_0024Local_selectedNameSet;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__2(Family x)
		{
			if (_0024VB_0024Local_selectedNameSet.Count != 0)
			{
				return _0024VB_0024Local_selectedNameSet.Contains(Normalize(x.Name ?? string.Empty));
			}
			return true;
		}
	}

	private const int PreviewCacheFileNameMaxLength = 120;

	private const double PreviewPaddingRatio = 1.15;

	private const double PreviewMinimumPaddingFeet = 1.5;

	private const double PreviewMinimumExtentFeet = 0.5;

	private const double PreviewConnectorMaxDiagonalRatio = 12.0;

	private const double PreviewConnectorMaxExtraFeet = 40.0;

	private const bool PreviewIncludeConnectorExtentsInBounds = true;

	private const bool PreviewUseSectionBox = false;

	private const bool PreviewShowConnectorGraphics = true;

	private const int PreviewConnectorSurfaceTransparency = 100;

	private const string PreviewCacheVersionFolderName = "preview-v14-white-background-centered-fit";

	private const int PreviewExportPixelSize = 768;

	private const double PreviewSafeRecenterScale = 0.94;

	private const double PreviewMinimumFrameMarginRatio = 0.08;

	private const double PreviewMinimumFrameMarginPixels = 40.0;

	private const string PreviewTempFileStemPrefix = "kky_thumb_";

	private FamilyThumbnailPreviewService()
	{
	}

	public static string GetCacheFolder(string workspaceRoot, string sourceId)
	{
		return Path.Combine(FamilyBrowserStandardPolicyStore.GetThumbnailFolder(workspaceRoot), SafeFileName(sourceId), PreviewCacheVersionFolderName);
	}

	public static string GetCachedImagePath(string workspaceRoot, string sourceId, string categoryName, string familyName)
	{
		return Path.Combine(GetCacheFolder(workspaceRoot, sourceId), BuildCacheFileName(categoryName, familyName));
	}

	public static string ResolveExistingCachedImagePath(string workspaceRoot, string sourceId, string categoryName, string familyName)
	{
		if (string.IsNullOrWhiteSpace(sourceId) || string.IsNullOrWhiteSpace(familyName))
		{
			return string.Empty;
		}
		try
		{
			string expected = GetCachedImagePath(workspaceRoot, sourceId, categoryName, familyName);
			return FamilyBrowserDataLoader.ResolveThumbnailPath(workspaceRoot, sourceId, categoryName, familyName, expected);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
			return string.Empty;
		}
	}

	public static string GetCachedFailureMessage(string imagePath)
	{
		string markerPath = BuildFailureMarkerPath(imagePath);
		string GetCachedFailureMessage;
		if (string.IsNullOrWhiteSpace(markerPath) || !File.Exists(markerPath))
		{
			GetCachedFailureMessage = string.Empty;
		}
		else
		{
			try
			{
				GetCachedFailureMessage = File.ReadAllText(markerPath, Encoding.UTF8);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				GetCachedFailureMessage = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return GetCachedFailureMessage;
	}

	public static FamilyThumbnailBatchUpdateResult UpdateStandardCache(Document standardDocument, string workspaceRoot, StandardLibraryRegistrationRecord registration, Action<int, int, string> progress, UIApplication uiApplication = null, ISet<string> selectedFamilyNames = null)
	{
		_Closure_0024__19_002D0 arg = default(_Closure_0024__19_002D0);
		_Closure_0024__19_002D0 CS_0024_003C_003E8__locals3 = new _Closure_0024__19_002D0(arg);
		if (standardDocument == null)
		{
			throw new ArgumentNullException("standardDocument");
		}
		if (registration == null)
		{
			throw new ArgumentNullException("registration");
		}
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Update standard 3D image cache", "표준 3D 이미지 캐시 갱신"));
		FamilyThumbnailBatchUpdateResult result = new FamilyThumbnailBatchUpdateResult();
		string outputFolder = GetCacheFolder(workspaceRoot, registration.SourceId);
		Directory.CreateDirectory(outputFolder);
		result.OutputFolder = outputFolder;
		StandardLibrarySnapshot standardSnapshot = LoadStandardSnapshot(registration.LastSnapshotPath);
		Dictionary<string, StandardLoadableFamilySnapshotItem> snapshotFamilyMap = BuildSnapshotFamilyMap(standardSnapshot);
		CS_0024_003C_003E8__locals3._0024VB_0024Local_selectedNameSet = NormalizeFamilyNameSet(selectedFamilyNames);
		List<Family> families = (from Family x in new FilteredElementCollector(standardDocument).OfClass(typeof(Family))
			where x != null && x.FamilyCategory != null
			where CS_0024_003C_003E8__locals3._0024VB_0024Local_selectedNameSet.Count == 0 || CS_0024_003C_003E8__locals3._0024VB_0024Local_selectedNameSet.Contains(Normalize(x.Name ?? string.Empty))
			select x).OrderBy<Family, string>([SpecialName] (Family x) => Normalize(x.FamilyCategory.Name) + "|" + Normalize(x.Name), StringComparer.Ordinal).ToList();
		int total = families.Count;
		int current = 0;
		checked
		{
			using (FamilyThumbnailConstraintDialogGuard dialogGuard = new FamilyThumbnailConstraintDialogGuard(uiApplication))
			{
				foreach (Family family in families)
				{
					current++;
					string categoryName = ((family.FamilyCategory == null) ? string.Empty : family.FamilyCategory.Name);
					string messageName = categoryName + " / " + family.Name;
					progress?.Invoke(current, total, messageName);
					FamilyThumbnailBatchUpdateItem item = new FamilyThumbnailBatchUpdateItem
					{
						FamilyName = family.Name,
						CategoryName = categoryName,
						ImagePath = GetCachedImagePath(workspaceRoot, registration.SourceId, categoryName, family.Name)
					};
					StandardLoadableFamilySnapshotItem snapshotItem = null;
					snapshotFamilyMap?.TryGetValue(BuildSnapshotFamilyKey(categoryName, family.Name), out snapshotItem);
					string cacheStamp = BuildFamilyThumbnailCacheStamp(registration, standardSnapshot, snapshotItem, categoryName, family.Name);
					int dialogRecordStart = dialogGuard.RecordCount;
					dialogGuard.SetCurrentFamily(categoryName, family.Name);
					try
					{
						if (IsCachedThumbnailCurrent(item.ImagePath, cacheStamp))
						{
							item.Skipped = true;
							item.Message = "Snapshot skipped: cached image matches the registered standard snapshot metadata.";
							result.SkippedCount++;
						}
						else if (!IsFamilyEditable(family))
						{
							item.Skipped = true;
							item.Message = "Family is not editable.";
							result.SkippedCount++;
						}
						else
						{
							FamilyThumbnailGenerationResult generation = GenerateAccurate3DPreview(standardDocument, family, item.ImagePath);
							item.Success = true;
							item.Message = generation.Message;
							result.SuccessCount++;
							WriteThumbnailCacheMetadata(item.ImagePath, registration, standardSnapshot, categoryName, family.Name, cacheStamp);
						}
					}
					catch (FamilyThumbnailStageException ex)
					{
						ProjectData.SetProjectError(ex);
						FamilyThumbnailStageException ex2 = ex;
						item.Success = false;
						item.Message = ex2.Message;
						result.FailedCount++;
						ProjectData.ClearProjectError();
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						item.Success = false;
						item.Message = "Snapshot failed: " + ex4.Message;
						result.FailedCount++;
						ProjectData.ClearProjectError();
					}
					finally
					{
						dialogGuard.ClearCurrentFamily();
					}
					List<FamilyThumbnailAutoConfirmedDialogRecord> newDialogRecords = dialogGuard.GetRecordsSince(dialogRecordStart);
					if (newDialogRecords.Count > 0)
					{
						result.AutoConfirmedDialogs.AddRange(newDialogRecords);
						item.AutoConfirmedDialogs.AddRange(newDialogRecords.Select([SpecialName] (FamilyThumbnailAutoConfirmedDialogRecord x) => BuildAutoConfirmedDialogSummary(x)));
						item.Message = AppendAutoConfirmedDialogMessage(item.Message, newDialogRecords);
					}
					if (item.Success)
					{
						DeleteFailureMarker(item.ImagePath);
					}
					else if (item.Skipped && File.Exists(item.ImagePath))
					{
						DeleteFailureMarker(item.ImagePath);
					}
					else
					{
						WriteFailureMarker(item.ImagePath, item.Message);
					}
					result.Items.Add(item);
				}
			}
			return result;
		}
	}

	public static FamilyThumbnailBatchUpdateResult UpdateProjectCache(Document projectDocument, string workspaceRoot, string projectThumbnailSourceId, ProjectContentSnapshot projectSnapshot, Action<int, int, string> progress, UIApplication uiApplication = null, ISet<string> selectedFamilyNames = null)
	{
		_Closure_0024__20_002D0 arg = default(_Closure_0024__20_002D0);
		_Closure_0024__20_002D0 CS_0024_003C_003E8__locals3 = new _Closure_0024__20_002D0(arg);
		if (projectDocument == null)
		{
			throw new ArgumentNullException("projectDocument");
		}
		if (string.IsNullOrWhiteSpace(projectThumbnailSourceId))
		{
			throw new ArgumentException("A project thumbnail source id is required.", "projectThumbnailSourceId");
		}
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Update project 3D image cache", "프로젝트 3D 이미지 캐시 갱신"));
		FamilyThumbnailBatchUpdateResult result = new FamilyThumbnailBatchUpdateResult();
		string outputFolder = GetCacheFolder(workspaceRoot, projectThumbnailSourceId);
		Directory.CreateDirectory(outputFolder);
		result.OutputFolder = outputFolder;
		Dictionary<string, ProjectLoadableFamilySnapshotItem> snapshotFamilyMap = BuildProjectSnapshotFamilyMap(projectSnapshot);
		CS_0024_003C_003E8__locals3._0024VB_0024Local_selectedNameSet = NormalizeFamilyNameSet(selectedFamilyNames);
		List<Family> families = (from Family x in new FilteredElementCollector(projectDocument).OfClass(typeof(Family))
			where x != null && x.FamilyCategory != null
			where FamilyBrowserFamilyClassificationService.IsBrowserLoadableFamily(x)
			where CS_0024_003C_003E8__locals3._0024VB_0024Local_selectedNameSet.Count == 0 || CS_0024_003C_003E8__locals3._0024VB_0024Local_selectedNameSet.Contains(Normalize(x.Name ?? string.Empty))
			select x).OrderBy<Family, string>([SpecialName] (Family x) => Normalize(x.FamilyCategory.Name) + "|" + Normalize(x.Name), StringComparer.Ordinal).ToList();
		int total = families.Count;
		int current = 0;
		checked
		{
			using (FamilyThumbnailConstraintDialogGuard dialogGuard = new FamilyThumbnailConstraintDialogGuard(uiApplication))
			{
				foreach (Family family in families)
				{
					current++;
					string categoryName = ((family.FamilyCategory == null) ? string.Empty : family.FamilyCategory.Name);
					string messageName = categoryName + " / " + family.Name;
					progress?.Invoke(current, total, messageName);
					FamilyThumbnailBatchUpdateItem item = new FamilyThumbnailBatchUpdateItem
					{
						FamilyName = family.Name,
						CategoryName = categoryName,
						ImagePath = GetCachedImagePath(workspaceRoot, projectThumbnailSourceId, categoryName, family.Name)
					};
					ProjectLoadableFamilySnapshotItem snapshotItem = null;
					snapshotFamilyMap?.TryGetValue(BuildSnapshotFamilyKey(categoryName, family.Name), out snapshotItem);
					string cacheStamp = BuildProjectFamilyThumbnailCacheStamp(projectThumbnailSourceId, projectSnapshot, snapshotItem, categoryName, family.Name);
					int dialogRecordStart = dialogGuard.RecordCount;
					dialogGuard.SetCurrentFamily(categoryName, family.Name);
					try
					{
						if (IsCachedThumbnailCurrent(item.ImagePath, cacheStamp))
						{
							item.Skipped = true;
							item.Message = "Snapshot skipped: cached image matches the current project scan metadata.";
							result.SkippedCount++;
						}
						else if (!IsFamilyEditable(family))
						{
							item.Skipped = true;
							item.Message = "Family is not editable.";
							result.SkippedCount++;
						}
						else
						{
							FamilyThumbnailGenerationResult generation = GenerateAccurate3DPreview(projectDocument, family, item.ImagePath);
							item.Success = true;
							item.Message = generation.Message;
							result.SuccessCount++;
							WriteThumbnailCacheMetadata(item.ImagePath, projectThumbnailSourceId, string.Empty, 0L, "Project", projectSnapshot?.CapturedAtUtc ?? string.Empty, categoryName, family.Name, cacheStamp);
						}
					}
					catch (FamilyThumbnailStageException ex)
					{
						ProjectData.SetProjectError(ex);
						FamilyThumbnailStageException ex2 = ex;
						item.Success = false;
						item.Message = ex2.Message;
						result.FailedCount++;
						ProjectData.ClearProjectError();
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						item.Success = false;
						item.Message = "Snapshot failed: " + ex4.Message;
						result.FailedCount++;
						ProjectData.ClearProjectError();
					}
					finally
					{
						dialogGuard.ClearCurrentFamily();
					}
					List<FamilyThumbnailAutoConfirmedDialogRecord> newDialogRecords = dialogGuard.GetRecordsSince(dialogRecordStart);
					if (newDialogRecords.Count > 0)
					{
						result.AutoConfirmedDialogs.AddRange(newDialogRecords);
						item.AutoConfirmedDialogs.AddRange(newDialogRecords.Select([SpecialName] (FamilyThumbnailAutoConfirmedDialogRecord x) => BuildAutoConfirmedDialogSummary(x)));
						item.Message = AppendAutoConfirmedDialogMessage(item.Message, newDialogRecords);
					}
					if (item.Success)
					{
						DeleteFailureMarker(item.ImagePath);
					}
					else if (item.Skipped && File.Exists(item.ImagePath))
					{
						DeleteFailureMarker(item.ImagePath);
					}
					else
					{
						WriteFailureMarker(item.ImagePath, item.Message);
					}
					result.Items.Add(item);
				}
			}
			return result;
		}
	}

	private static ISet<string> NormalizeFamilyNameSet(ISet<string> values)
	{
		HashSet<string> result = new HashSet<string>(StringComparer.Ordinal);
		if (values == null)
		{
			return result;
		}
		foreach (string value in values)
		{
			string normalized = Normalize(value);
			if (normalized.Length > 0)
			{
				result.Add(normalized);
			}
		}
		return result;
	}

	private static StandardLibrarySnapshot LoadStandardSnapshot(string snapshotPath)
	{
		StandardLibrarySnapshot LoadStandardSnapshot;
		if (string.IsNullOrWhiteSpace(snapshotPath) || !File.Exists(snapshotPath))
		{
			LoadStandardSnapshot = null;
		}
		else
		{
			try
			{
				LoadStandardSnapshot = DataContractJsonFileStore.Load<StandardLibrarySnapshot>(snapshotPath);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				LoadStandardSnapshot = null;
				ProjectData.ClearProjectError();
			}
		}
		return LoadStandardSnapshot;
	}

	private static Dictionary<string, StandardLoadableFamilySnapshotItem> BuildSnapshotFamilyMap(StandardLibrarySnapshot snapshot)
	{
		Dictionary<string, StandardLoadableFamilySnapshotItem> result = new Dictionary<string, StandardLoadableFamilySnapshotItem>(StringComparer.Ordinal);
		if (snapshot == null || snapshot.LoadableFamilies == null)
		{
			return result;
		}
		foreach (StandardLoadableFamilySnapshotItem item in snapshot.LoadableFamilies)
		{
			if (item != null && !string.IsNullOrWhiteSpace(item.FamilyName))
			{
				string key = BuildSnapshotFamilyKey(item.CategoryName, item.FamilyName);
				if (!string.IsNullOrWhiteSpace(key) && !result.ContainsKey(key))
				{
					result.Add(key, item);
				}
			}
		}
		return result;
	}

	private static Dictionary<string, ProjectLoadableFamilySnapshotItem> BuildProjectSnapshotFamilyMap(ProjectContentSnapshot snapshot)
	{
		Dictionary<string, ProjectLoadableFamilySnapshotItem> result = new Dictionary<string, ProjectLoadableFamilySnapshotItem>(StringComparer.Ordinal);
		if (snapshot == null || snapshot.LoadableFamilies == null)
		{
			return result;
		}
		foreach (ProjectLoadableFamilySnapshotItem item in snapshot.LoadableFamilies)
		{
			if (item != null && !string.IsNullOrWhiteSpace(item.FamilyName))
			{
				string key = BuildSnapshotFamilyKey(item.CategoryName, item.FamilyName);
				if (!string.IsNullOrWhiteSpace(key) && !result.ContainsKey(key))
				{
					result.Add(key, item);
				}
			}
		}
		return result;
	}

	private static string BuildSnapshotFamilyKey(string categoryName, string familyName)
	{
		return Normalize(categoryName) + "|" + Normalize(familyName);
	}

	private static string BuildFamilyThumbnailCacheStamp(StandardLibraryRegistrationRecord registration, StandardLibrarySnapshot snapshot, StandardLoadableFamilySnapshotItem snapshotItem, string categoryName, string familyName)
	{
		List<string> list = new List<string>();
		list.Add("thumbnail-cache-stamp-v9-white-background-centered-fit");
		list.Add("source-id=" + Normalize(registration?.SourceId ?? string.Empty));
		list.Add("snapshot-mode=" + Normalize(snapshot?.SnapshotMode ?? registration?.SnapshotMode));
		list.Add("category=" + Normalize(categoryName));
		list.Add("family=" + Normalize(familyName));
		List<string> parts = list;
		bool num = snapshotItem != null && (IsPreciseSnapshot(snapshot) || string.Equals(snapshotItem.MetadataMode ?? string.Empty, "Precise", StringComparison.OrdinalIgnoreCase)) && !string.IsNullOrWhiteSpace(snapshotItem.ContentFingerprint);
		if (snapshotItem != null)
		{
			parts.Add("types=" + BuildOrderedStringStamp(snapshotItem.TypeNames));
			parts.Add("shared=" + snapshotItem.IsShared);
			parts.Add("content=" + Normalize(snapshotItem.ContentFingerprint));
			parts.Add("parameters=" + BuildParameterCacheStamp(snapshotItem.Parameters));
			parts.Add("nested=" + BuildNestedFamilyCacheStamp(snapshotItem.NestedLoadableFamilies));
		}
		if (!num)
		{
			parts.Add("source-lastwrite=" + Normalize(registration?.SourceFileLastWriteUtc ?? string.Empty));
			parts.Add("source-length=" + (registration?.SourceFileLength ?? 0).ToString(CultureInfo.InvariantCulture));
		}
		return HashString(string.Join("\n", parts));
	}

	private static string BuildProjectFamilyThumbnailCacheStamp(string projectThumbnailSourceId, ProjectContentSnapshot snapshot, ProjectLoadableFamilySnapshotItem snapshotItem, string categoryName, string familyName)
	{
		List<string> list = new List<string>();
		list.Add("project-thumbnail-cache-stamp-v4-white-background-centered-fit");
		list.Add("source-id=" + Normalize(projectThumbnailSourceId));
		list.Add("captured-at=" + Normalize(snapshot?.CapturedAtUtc ?? string.Empty));
		list.Add("document-path=" + Normalize(snapshot?.DocumentPath ?? string.Empty));
		list.Add("category=" + Normalize(categoryName));
		list.Add("family=" + Normalize(familyName));
		List<string> parts = list;
		if (snapshotItem != null)
		{
			parts.Add("types=" + BuildOrderedStringStamp(snapshotItem.TypeNames));
			parts.Add("shared=" + snapshotItem.IsShared);
			parts.Add("content=" + Normalize(snapshotItem.ContentFingerprint));
			parts.Add("parameters=" + BuildParameterCacheStamp(snapshotItem.Parameters));
			parts.Add("unique-id=" + Normalize(snapshotItem.UniqueId));
			parts.Add("type-count=" + snapshotItem.TypeCount.ToString(CultureInfo.InvariantCulture));
			parts.Add("instance-count=" + snapshotItem.InstanceCount.ToString(CultureInfo.InvariantCulture));
		}
		return HashString(string.Join("\n", parts));
	}

	private static bool IsPreciseSnapshot(StandardLibrarySnapshot snapshot)
	{
		if (snapshot != null)
		{
			return string.Equals(snapshot.SnapshotMode ?? string.Empty, "Precise", StringComparison.OrdinalIgnoreCase);
		}
		return false;
	}

	private static string BuildOrderedStringStamp(IEnumerable<string> values)
	{
		return string.Join("|", (from x in values ?? Enumerable.Empty<string>()
			select Normalize(x) into x
			where x.Length > 0
			select x).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal));
	}

	private static string BuildParameterCacheStamp(IEnumerable<StandardFamilyParameterSnapshotItem> parameters)
	{
		return string.Join("|", (from x in parameters ?? Enumerable.Empty<StandardFamilyParameterSnapshotItem>()
			where x != null
			select Normalize(x.Scope) + ":" + Normalize(x.TypeName) + ":" + Normalize(x.Name) + ":" + Normalize(x.StorageType) + ":" + Normalize(x.ValuePreview) + ":" + x.IsShared + ":" + Normalize(x.ExternalGuid) + ":" + Normalize(x.ParameterId)).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal));
	}

	private static string BuildNestedFamilyCacheStamp(IEnumerable<StandardNestedLoadableFamilySnapshotItem> items)
	{
		return string.Join("|", (from x in items ?? Enumerable.Empty<StandardNestedLoadableFamilySnapshotItem>()
			where x != null
			select Normalize(x.CategoryName) + ":" + Normalize(x.FamilyName) + ":" + x.TypeCount.ToString(CultureInfo.InvariantCulture) + ":" + BuildOrderedStringStamp(x.TypeNames) + ":" + x.IsShared).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal));
	}

	private static bool IsCachedThumbnailCurrent(string imagePath, string cacheStamp)
	{
		bool IsCachedThumbnailCurrent;
		if (string.IsNullOrWhiteSpace(imagePath) || string.IsNullOrWhiteSpace(cacheStamp) || !File.Exists(imagePath))
		{
			IsCachedThumbnailCurrent = false;
		}
		else
		{
			string metadataPath = BuildThumbnailMetadataPath(imagePath);
			if (string.IsNullOrWhiteSpace(metadataPath) || !File.Exists(metadataPath))
			{
				IsCachedThumbnailCurrent = false;
			}
			else
			{
				try
				{
					FamilyThumbnailCacheMetadata metadata = DataContractJsonFileStore.Load<FamilyThumbnailCacheMetadata>(metadataPath);
					IsCachedThumbnailCurrent = metadata != null && metadata.SchemaVersion == 1 && string.Equals(metadata.FamilyCacheStamp ?? string.Empty, cacheStamp, StringComparison.Ordinal);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					IsCachedThumbnailCurrent = false;
					ProjectData.ClearProjectError();
				}
			}
		}
		return IsCachedThumbnailCurrent;
	}

	private static string BuildThumbnailMetadataPath(string imagePath)
	{
		if (string.IsNullOrWhiteSpace(imagePath))
		{
			return string.Empty;
		}
		return imagePath + ".meta.json";
	}

	private static string FindCachedImageByMetadata(string sourceFolder, string sourceId, string categoryName, string familyName, bool requireCategoryMatch)
	{
		if (string.IsNullOrWhiteSpace(sourceFolder) || string.IsNullOrWhiteSpace(familyName) || !Directory.Exists(sourceFolder))
		{
			return string.Empty;
		}
		try
		{
			foreach (string metadataPath in Directory.EnumerateFiles(sourceFolder, "*.png.meta.json", SearchOption.AllDirectories))
			{
				FamilyThumbnailCacheMetadata metadata = null;
				try
				{
					metadata = DataContractJsonFileStore.Load<FamilyThumbnailCacheMetadata>(metadataPath);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					metadata = null;
					ProjectData.ClearProjectError();
				}
				if (metadata != null && (string.IsNullOrWhiteSpace(sourceId) || string.Equals(Normalize(metadata.SourceId), Normalize(sourceId), StringComparison.Ordinal)) && string.Equals(Normalize(metadata.FamilyName), Normalize(familyName), StringComparison.Ordinal) && (!requireCategoryMatch || string.Equals(Normalize(metadata.CategoryName), Normalize(categoryName), StringComparison.Ordinal)))
				{
					string imagePath = metadataPath.Substring(0, checked(metadataPath.Length - ".meta.json".Length));
					if (File.Exists(imagePath))
					{
						return imagePath;
					}
				}
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		return string.Empty;
	}

	private static void WriteThumbnailCacheMetadata(string imagePath, StandardLibraryRegistrationRecord registration, StandardLibrarySnapshot snapshot, string categoryName, string familyName, string cacheStamp)
	{
		WriteThumbnailCacheMetadata(imagePath, registration?.SourceId ?? string.Empty, registration?.SourceFileLastWriteUtc ?? string.Empty, registration?.SourceFileLength ?? 0, snapshot?.SnapshotMode ?? registration?.SnapshotMode, snapshot?.CapturedAtUtc ?? string.Empty, categoryName, familyName, cacheStamp);
	}

	private static void WriteThumbnailCacheMetadata(string imagePath, string sourceId, string sourceFileLastWriteUtc, long sourceFileLength, string snapshotMode, string snapshotCapturedAtUtc, string categoryName, string familyName, string cacheStamp)
	{
		if (!string.IsNullOrWhiteSpace(imagePath) && !string.IsNullOrWhiteSpace(cacheStamp))
		{
			try
			{
				FamilyThumbnailCacheMetadata metadata = new FamilyThumbnailCacheMetadata
				{
					SourceId = (sourceId ?? string.Empty),
					SourceFileLastWriteUtc = (sourceFileLastWriteUtc ?? string.Empty),
					SourceFileLength = sourceFileLength,
					SnapshotMode = (snapshotMode ?? string.Empty),
					SnapshotCapturedAtUtc = (snapshotCapturedAtUtc ?? string.Empty),
					CategoryName = (categoryName ?? string.Empty),
					FamilyName = (familyName ?? string.Empty),
					FamilyCacheStamp = cacheStamp,
					ImageGeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
				};
				string path = BuildThumbnailMetadataPath(imagePath);
				Directory.CreateDirectory(Path.GetDirectoryName(path));
				string temporaryPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
				try
				{
					using (FileStream stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
					using (StreamWriter writer = new StreamWriter(stream, new UTF8Encoding(false)))
					{
						writer.Write(PlainJsonReportWriter.Serialize(metadata));
						writer.Flush();
						stream.Flush(true);
					}
					FamilyBrowserAtomicFileService.Promote(temporaryPath, path);
					temporaryPath = string.Empty;
				}
				finally
				{
					if (!string.IsNullOrWhiteSpace(temporaryPath) && File.Exists(temporaryPath))
					{
						File.Delete(temporaryPath);
					}
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
	}

	public static bool IsCachedStandardThumbnailCurrent(string imagePath, StandardLibraryRegistrationRecord registration, StandardLibrarySnapshot snapshot, StandardLoadableFamilySnapshotItem snapshotItem, string categoryName, string familyName)
	{
		string cacheStamp = BuildFamilyThumbnailCacheStamp(registration, snapshot, snapshotItem, categoryName, familyName);
		return IsCachedThumbnailCurrent(imagePath, cacheStamp);
	}

	public static void WriteStandardThumbnailCacheMetadata(string imagePath, StandardLibraryRegistrationRecord registration, StandardLibrarySnapshot snapshot, StandardLoadableFamilySnapshotItem snapshotItem, string categoryName, string familyName)
	{
		string cacheStamp = BuildFamilyThumbnailCacheStamp(registration, snapshot, snapshotItem, categoryName, familyName);
		WriteThumbnailCacheMetadata(imagePath, registration, snapshot, categoryName, familyName, cacheStamp);
	}

	public static FamilyThumbnailPreviewResult Generate(Document projectDocument, string workspaceRoot, string familyName, string categoryName)
	{
		if (projectDocument == null)
		{
			throw new ArgumentNullException("projectDocument");
		}
		Family family = FindFamily(projectDocument, familyName, categoryName);
		FamilyThumbnailPreviewResult Generate;
		if (family == null)
		{
			Generate = new FamilyThumbnailPreviewResult
			{
				Success = false,
				Message = "Loaded family was not found in the active project."
			};
		}
		else
		{
			try
			{
				FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Generate family 3D image", "패밀리 3D 이미지 생성"));
				string outputFolder = FamilyBrowserStandardPolicyStore.GetThumbnailFolder(workspaceRoot);
				Directory.CreateDirectory(outputFolder);
				Bitmap previewBitmap = GetFastPreviewBitmap(projectDocument, family);
				if (previewBitmap == null)
				{
					Generate = new FamilyThumbnailPreviewResult
					{
						Success = false,
						Message = "Fast preview was not available for this loaded family."
					};
				}
				else
				{
					string fileStem = SafeFileName("fast_" + categoryName + "_" + familyName + "_" + DateTime.UtcNow.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture));
					string imagePath = Path.Combine(outputFolder, fileStem + ".png");
					using (previewBitmap)
					{
						previewBitmap.Save(imagePath, ImageFormat.Png);
					}
					Generate = new FamilyThumbnailPreviewResult
					{
						Success = true,
						ImagePath = imagePath,
						Message = "Fast loaded-family preview generated."
					};
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Generate = new FamilyThumbnailPreviewResult
				{
					Success = false,
					Message = ex2.Message
				};
				ProjectData.ClearProjectError();
			}
		}
		return Generate;
	}

	public static FamilyThumbnailGenerationResult GenerateFromOpenFamilyDocument(Document familyDocument, string targetImagePath)
	{
		if (familyDocument == null)
		{
			throw new ArgumentNullException("familyDocument");
		}
		if (string.IsNullOrWhiteSpace(targetImagePath))
		{
			throw new ArgumentException(FamilyBrowserLanguageService.Text("Target image path is required.", "대상 이미지 경로가 필요합니다."), "targetImagePath");
		}
		FamilyThumbnailGenerationResult generation = new FamilyThumbnailGenerationResult();
		ElementId viewId = ElementId.InvalidElementId;
		PreviewBoundsResult previewBoundsResult = null;
		RunPreviewStage(generation, "Create new 3D preview view", [SpecialName] () =>
		{
			using Transaction transaction = new Transaction(familyDocument, "KKY Family Browser Preview View");
			transaction.Start();
			try
			{
				FailureHandlingOptions failureHandlingOptions = transaction.GetFailureHandlingOptions();
				failureHandlingOptions.SetFailuresPreprocessor(new FamilyThumbnailPreviewFailuresPreprocessor());
				failureHandlingOptions.SetClearAfterRollback(bFlag: true);
				transaction.SetFailureHandlingOptions(failureHandlingOptions);
				View3D view3D = CreatePreviewView(familyDocument);
				previewBoundsResult = PreparePreviewView(familyDocument, view3D);
				viewId = view3D.Id;
				transaction.Commit();
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				try
				{
					if (transaction.GetStatus() == TransactionStatus.Started)
					{
						transaction.RollBack();
					}
				}
				catch (Exception projectError3)
				{
					ProjectData.SetProjectError(projectError3);
					ProjectData.ClearProjectError();
				}
				throw;
			}
		});
		if (previewBoundsResult != null && previewBoundsResult.ConnectorExtentsClamped)
		{
			generation.ConnectorExtentsClamped = true;
		}
		string directoryName = Path.GetDirectoryName(targetImagePath);
		RunPreviewStage(generation, "Prepare thumbnail output folder", [SpecialName] () =>
		{
			Directory.CreateDirectory(directoryName);
		});
		string text = BuildTempExportStem();
		string basePath = Path.Combine(directoryName, text);
		DateTime exportStartedUtc = DateTime.MinValue;
		RunPreviewStage(generation, "Export PNG snapshot", [SpecialName] () =>
		{
			exportStartedUtc = DateTime.UtcNow.AddSeconds(-2.0);
			ExportViewImage(familyDocument, viewId, basePath);
		});
		string text2 = string.Empty;
		RunPreviewStage(generation, "Find exported PNG", [SpecialName] () =>
		{
			text2 = FindExportedImage(directoryName, text, null, exportStartedUtc);
		});
		RunPreviewStage(generation, "Validate exported PNG", [SpecialName] () =>
		{
			if (string.IsNullOrWhiteSpace(text2) || !File.Exists(text2))
			{
				throw new InvalidOperationException(BuildMissingExportDiagnostic(directoryName, text, null, exportStartedUtc));
			}
		});
		RunPreviewStage(generation, "Save centered white-background thumbnail cache file", [SpecialName] () =>
		{
			if (!SaveWhiteBackgroundPreviewImage(text2, targetImagePath))
			{
				File.Copy(text2, targetImagePath, overwrite: true);
			}
		});
		try
		{
			File.Delete(text2);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		generation.Success = true;
		generation.Message = "Snapshot OK: new 3D view, shaded/fine, preview visibility on, annotations/reference datum hidden, connector graphics shown with thin/transparent overrides, element/category thin lines, low-DPI export for thinner connector strokes, focused home view, robust PNG export, centered safety margin restored, white background normalized.";
		if (generation.ConnectorExtentsClamped)
		{
			generation.Message += " Connector extents were clamped to keep the family geometry readable.";
		}
		return generation;
	}

	private static FamilyThumbnailGenerationResult GenerateAccurate3DPreview(Document projectDocument, Family family, string targetImagePath)
	{
		Document document = null;
		try
		{
			RunPreviewStage(new FamilyThumbnailGenerationResult(), "Open family in background", [SpecialName] () =>
			{
				document = projectDocument.EditFamily(family);
			});
			return GenerateFromOpenFamilyDocument(document, targetImagePath);
		}
		finally
		{
			if (document != null)
			{
				try
				{
					document.Close(saveModified: false);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
	}

	private static Bitmap GetFastPreviewBitmap(Document projectDocument, Family family)
	{
		foreach (ElementId symbolId in family.GetFamilySymbolIds())
		{
			Bitmap bitmap = TryGetPreviewImage(projectDocument.GetElement(symbolId));
			if (bitmap != null)
			{
				return bitmap;
			}
		}
		return TryGetPreviewImage(family);
	}

	private static Bitmap TryGetPreviewImage(Element element)
	{
		Bitmap TryGetPreviewImage;
		if (element == null)
		{
			TryGetPreviewImage = null;
		}
		else if (!(element is ElementType elementType))
		{
			TryGetPreviewImage = null;
		}
		else
		{
			try
			{
				TryGetPreviewImage = elementType.GetPreviewImage(new Size(512, 512));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				TryGetPreviewImage = null;
				ProjectData.ClearProjectError();
			}
		}
		return TryGetPreviewImage;
	}

	private static Family FindFamily(Document projectDocument, string familyName, string categoryName)
	{
		string normalizedFamilyName = Normalize(familyName);
		string normalizedCategoryName = Normalize(categoryName);
		foreach (Family family in new FilteredElementCollector(projectDocument).OfClass(typeof(Family)))
		{
			if (string.Equals(Normalize(family.Name), normalizedFamilyName, StringComparison.Ordinal))
			{
				if (normalizedCategoryName.Length == 0)
				{
					return family;
				}
				if (string.Equals(Normalize((family.FamilyCategory == null) ? string.Empty : family.FamilyCategory.Name), normalizedCategoryName, StringComparison.Ordinal))
				{
					return family;
				}
			}
		}
		return null;
	}

	private static View3D CreatePreviewView(Document familyDocument)
	{
		ViewFamilyType viewType = new FilteredElementCollector(familyDocument).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault([SpecialName] (ViewFamilyType x) => x.ViewFamily == ViewFamily.ThreeDimensional);
		if (viewType == null)
		{
			throw new InvalidOperationException(FamilyBrowserLanguageService.Text("No 3D view family type was found in the family document.", "패밀리 문서에서 3D 뷰 패밀리 타입을 찾지 못했습니다."));
		}
		View3D view = View3D.CreateIsometric(familyDocument, viewType.Id);
		try
		{
			view.Name = "KKY Family Browser Preview 3D " + DateTime.UtcNow.ToString("HHmmssfff", CultureInfo.InvariantCulture);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return view;
	}

	private static PreviewBoundsResult PreparePreviewView(Document familyDocument, View3D view)
	{
		if (view == null)
		{
			throw new ArgumentNullException("view");
		}
		try
		{
			view.DetailLevel = ViewDetailLevel.Fine;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		SetPreviewDisplayStyle(view, "Shading", "ShadingWithEdges", "Realistic", "HLR", "HiddenLine", "Wireframe");
		try
		{
			view.Scale = 1;
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		try
		{
			view.AreAnnotationCategoriesHidden = true;
		}
		catch (Exception projectError3)
		{
			ProjectData.SetProjectError(projectError3);
			ProjectData.ClearProjectError();
		}
		EnablePreviewVisibilityParameters(familyDocument);
		ConfigurePreviewCategories(familyDocument, view);
		PreviewBoundsResult previewBounds = BuildPreviewBounds(familyDocument);
		ApplyPreviewFocus(view, previewBounds?.Bounds);
		return previewBounds;
	}

	private static void ConfigurePreviewCategories(Document familyDocument, View3D view)
	{
		HideReferenceDatumCategories(familyDocument, view);
		foreach (Category category in familyDocument.Settings.Categories)
		{
			ConfigurePreviewCategory(view, category);
		}
		ApplyThinLineOverridesToModelElements(familyDocument, view);
	}

	private static void EnablePreviewVisibilityParameters(Document familyDocument)
	{
		try
		{
			foreach (Element element in new FilteredElementCollector(familyDocument).WhereElementIsNotElementType())
			{
				if (element == null || element.Parameters == null)
				{
					continue;
				}
				foreach (Parameter parameter in element.Parameters)
				{
					try
					{
						if (parameter != null && !parameter.IsReadOnly && parameter.StorageType == StorageType.Integer && IsPreviewVisibilityParameterName((parameter.Definition == null) ? string.Empty : parameter.Definition.Name))
						{
							parameter.Set(1);
						}
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						ProjectData.ClearProjectError();
					}
				}
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
	}

	private static bool IsPreviewVisibilityParameterName(string value)
	{
		string compact = Normalize(value).Replace(" ", string.Empty);
		if (Operators.CompareString(compact, "previewvisibility", TextCompare: false) != 0 && Operators.CompareString(compact, "previewvisible", TextCompare: false) != 0)
		{
			return compact.Contains("previewvisibility");
		}
		return true;
	}

	private static void ConfigurePreviewCategory(View3D view, Category category)
	{
		if (category == null)
		{
			return;
		}
		if (category.CategoryType == CategoryType.Annotation || IsReferenceDatumCategory(category))
		{
			HideCategory(view, category);
		}
		else
		{
			ShowCategory(view, category);
			if (IsConnectorCategory(category))
			{
				ApplyConnectorThinLineOverride(view, category);
			}
			else
			{
				ApplyThinLineOverride(view, category);
			}
		}
		try
		{
			foreach (Category subCategory in category.SubCategories)
			{
				ConfigurePreviewCategory(view, subCategory);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void HideReferenceDatumCategories(Document familyDocument, View3D view)
	{
		HideElementClassCategories<ReferencePlane>(familyDocument, view);
		HideElementClassCategories<Level>(familyDocument, view);
	}

	private static void HideElementClassCategories<T>(Document familyDocument, View3D view) where T : Element
	{
		try
		{
			foreach (T item in new FilteredElementCollector(familyDocument).OfClass(typeof(T)))
			{
				T element = item;
				if (element != null)
				{
					HideCategory(view, element.Category);
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void HideCategory(View3D view, Category category)
	{
		if (category == null)
		{
			return;
		}
		try
		{
			if (view.CanCategoryBeHidden(category.Id))
			{
				view.SetCategoryHidden(category.Id, hide: true);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void ShowCategory(View3D view, Category category)
	{
		if (category == null)
		{
			return;
		}
		try
		{
			if (view.CanCategoryBeHidden(category.Id))
			{
				view.SetCategoryHidden(category.Id, hide: false);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void ApplyThinLineOverride(View3D view, Category category)
	{
		if (category != null)
		{
			try
			{
				view.SetCategoryOverrides(category.Id, CreateThinLineOverride());
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
	}

	private static OverrideGraphicSettings CreateThinLineOverride()
	{
		OverrideGraphicSettings graphicsOverride = new OverrideGraphicSettings();
		try
		{
			graphicsOverride.SetProjectionLineWeight(1);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			graphicsOverride.SetCutLineWeight(1);
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		return graphicsOverride;
	}

	private static OverrideGraphicSettings CreateConnectorThinLineOverride()
	{
		OverrideGraphicSettings graphicsOverride = CreateThinLineOverride();
		try
		{
			graphicsOverride.SetSurfaceTransparency(100);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			graphicsOverride.SetHalftone(halftone: true);
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		return graphicsOverride;
	}

	private static void ApplyConnectorThinLineOverride(View3D view, Category category)
	{
		if (category != null)
		{
			try
			{
				view.SetCategoryOverrides(category.Id, CreateConnectorThinLineOverride());
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
	}

	private static void SetPreviewDisplayStyle(View3D view, params string[] preferredStyleNames)
	{
		if (view == null || preferredStyleNames == null)
		{
			return;
		}
		foreach (string styleName in preferredStyleNames)
		{
			if (!string.IsNullOrWhiteSpace(styleName))
			{
				try
				{
					view.DisplayStyle = (DisplayStyle)Enum.Parse(typeof(DisplayStyle), styleName, ignoreCase: true);
					break;
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
	}

	private static void ApplyThinLineOverridesToModelElements(Document familyDocument, View3D view)
	{
		if (familyDocument == null || view == null)
		{
			return;
		}
		try
		{
			OverrideGraphicSettings graphicsOverride = CreateThinLineOverride();
			OverrideGraphicSettings connectorGraphicsOverride = CreateConnectorThinLineOverride();
			foreach (Element element in new FilteredElementCollector(familyDocument).WhereElementIsNotElementType())
			{
				if (element == null || element.Category == null || element.Category.CategoryType != CategoryType.Model || element is View || element is ReferencePlane || element is Level)
				{
					continue;
				}
				try
				{
					if (IsConnectorElement(element))
					{
						view.SetElementOverrides(element.Id, connectorGraphicsOverride);
					}
					else
					{
						view.SetElementOverrides(element.Id, graphicsOverride);
					}
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
	}

	private static PreviewBoundsResult BuildPreviewBounds(Document familyDocument)
	{
		BoundingBoxXYZ physicalBounds = null;
		BoundingBoxXYZ connectorBounds = null;
		BoundingBoxXYZ fallbackModelBounds = null;
		BoundingBoxXYZ geometryBounds = null;
		foreach (Element element in new FilteredElementCollector(familyDocument).WhereElementIsNotElementType())
		{
			if (element == null || element.Category == null || element.Category.CategoryType != CategoryType.Model)
			{
				continue;
			}
			BoundingBoxXYZ elementBounds = GetUsableElementBounds(element);
			BoundingBoxXYZ elementGeometryBounds = GetUsableElementGeometryBounds(element);
			if (elementGeometryBounds != null)
			{
				elementBounds = UnionBounds(elementBounds, elementGeometryBounds);
				geometryBounds = UnionBounds(geometryBounds, elementGeometryBounds);
			}
			if (elementBounds != null)
			{
				fallbackModelBounds = UnionBounds(fallbackModelBounds, elementBounds);
				if (IsConnectorElement(element))
				{
					connectorBounds = UnionBounds(connectorBounds, elementBounds);
				}
				else if (IsPhysicalPreviewElement(element))
				{
					physicalBounds = UnionBounds(physicalBounds, elementBounds);
				}
			}
		}
		BoundingBoxXYZ baseBounds = physicalBounds ?? geometryBounds ?? fallbackModelBounds;
		if (baseBounds == null)
		{
			return null;
		}
		BoundingBoxXYZ displayBounds = UnionBounds(CloneBounds(baseBounds), geometryBounds);
		bool connectorIncluded = false;
		bool connectorClamped = false;
		if (connectorBounds != null)
		{
			BoundingBoxXYZ boundsWithConnectors = UnionBounds(displayBounds, connectorBounds);
			double baseDiagonal = BoundsDiagonal(baseBounds);
			double connectorDiagonal = BoundsDiagonal(boundsWithConnectors);
			double maxAllowedDiagonal = Math.Max(baseDiagonal * 12.0, baseDiagonal + 40.0);
			if (baseDiagonal <= 0.0 || connectorDiagonal <= maxAllowedDiagonal)
			{
				displayBounds = boundsWithConnectors;
				connectorIncluded = true;
			}
			else
			{
				connectorClamped = true;
			}
		}
		return new PreviewBoundsResult
		{
			Bounds = ExpandBounds(displayBounds),
			PhysicalBounds = physicalBounds,
			ConnectorBounds = connectorBounds,
			ConnectorExtentsIncluded = connectorIncluded,
			ConnectorExtentsClamped = connectorClamped
		};
	}

	private static BoundingBoxXYZ GetUsableElementBounds(Element element)
	{
		try
		{
			BoundingBoxXYZ bounds = element.get_BoundingBox((View)null);
			if (IsUsableBounds(bounds))
			{
				return NormalizeBoundsToModelCoordinates(bounds);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return null;
	}

	private static BoundingBoxXYZ GetUsableElementGeometryBounds(Element element)
	{
		BoundingBoxXYZ GetUsableElementGeometryBounds;
		if (element == null)
		{
			GetUsableElementGeometryBounds = null;
		}
		else
		{
			try
			{
				Options options = new Options
				{
					ComputeReferences = false,
					IncludeNonVisibleObjects = false
				};
				GetUsableElementGeometryBounds = BuildGeometryBounds(element.get_Geometry(options), Transform.Identity);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				GetUsableElementGeometryBounds = null;
				ProjectData.ClearProjectError();
			}
		}
		return GetUsableElementGeometryBounds;
	}

	private static BoundingBoxXYZ BuildGeometryBounds(GeometryElement geometry, Transform transform)
	{
		if ((object)geometry == null)
		{
			return null;
		}
		BoundingBoxXYZ result = null;
		foreach (GeometryObject geometryObject in geometry)
		{
			result = UnionBounds(result, BuildGeometryObjectBounds(geometryObject, transform));
		}
		return result;
	}

	private static BoundingBoxXYZ BuildGeometryObjectBounds(GeometryObject geometryObject, Transform transform)
	{
		if ((object)geometryObject == null)
		{
			return null;
		}
		Transform activeTransform = transform ?? Transform.Identity;
		try
		{
			if (geometryObject is GeometryInstance instance)
			{
				Transform instanceTransform = activeTransform.Multiply(instance.Transform);
				return BuildGeometryBounds(instance.GetInstanceGeometry(), instanceTransform);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			if (geometryObject is Solid { Faces: not null } solid && solid.Faces.Size > 0)
			{
				return TransformBounds(NormalizeBoundsToModelCoordinates(solid.GetBoundingBox()), activeTransform);
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		try
		{
			if (geometryObject is Curve curve)
			{
				return TransformBounds(BuildCurveBounds(curve), activeTransform);
			}
		}
		catch (Exception projectError3)
		{
			ProjectData.SetProjectError(projectError3);
			ProjectData.ClearProjectError();
		}
		try
		{
			if (geometryObject is Mesh { NumTriangles: >0 } mesh)
			{
				return TransformBounds(BuildMeshBounds(mesh), activeTransform);
			}
		}
		catch (Exception projectError4)
		{
			ProjectData.SetProjectError(projectError4);
			ProjectData.ClearProjectError();
		}
		return null;
	}

	private static BoundingBoxXYZ BuildCurveBounds(Curve curve)
	{
		if ((object)curve == null)
		{
			return null;
		}
		BoundingBoxXYZ result = null;
		try
		{
			foreach (XYZ point in curve.Tessellate())
			{
				result = UnionPoint(result, point);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static BoundingBoxXYZ BuildMeshBounds(Mesh mesh)
	{
		if ((object)mesh == null)
		{
			return null;
		}
		BoundingBoxXYZ result = null;
		try
		{
			PropertyInfo verticesProperty = mesh.GetType().GetProperty("Vertices", BindingFlags.Instance | BindingFlags.Public);
			if ((object)verticesProperty != null && verticesProperty.GetValue(mesh, null) is IEnumerable<XYZ> vertices)
			{
				foreach (XYZ point in vertices)
				{
					result = UnionPoint(result, point);
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		if (result != null)
		{
			return result;
		}
		checked
		{
			try
			{
				MethodInfo triangleMethod = mesh.GetType().GetMethod("get_Triangle", BindingFlags.Instance | BindingFlags.Public);
				if ((object)triangleMethod == null)
				{
					triangleMethod = mesh.GetType().GetMethod("Triangle", BindingFlags.Instance | BindingFlags.Public);
				}
				if ((object)triangleMethod != null)
				{
					int num = mesh.NumTriangles - 1;
					for (int triangleIndex = 0; triangleIndex <= num; triangleIndex++)
					{
						object triangle = RuntimeHelpers.GetObjectValue(triangleMethod.Invoke(mesh, new object[1] { triangleIndex }));
						if (triangle == null)
						{
							continue;
						}
						MethodInfo vertexMethod = triangle.GetType().GetMethod("get_Vertex", BindingFlags.Instance | BindingFlags.Public);
						if ((object)vertexMethod == null)
						{
							vertexMethod = triangle.GetType().GetMethod("Vertex", BindingFlags.Instance | BindingFlags.Public);
						}
						if ((object)vertexMethod != null)
						{
							int vertexIndex = 0;
							do
							{
								XYZ point2 = vertexMethod.Invoke(RuntimeHelpers.GetObjectValue(triangle), new object[1] { vertexIndex }) as XYZ;
								result = UnionPoint(result, point2);
								vertexIndex++;
							}
							while (vertexIndex <= 2);
						}
					}
				}
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				ProjectData.ClearProjectError();
			}
			return result;
		}
	}

	private static BoundingBoxXYZ TransformBounds(BoundingBoxXYZ bounds, Transform transform)
	{
		if (bounds == null || !IsUsableBounds(bounds))
		{
			return null;
		}
		Transform activeTransform = transform ?? Transform.Identity;
		XYZ min = bounds.Min;
		XYZ max = bounds.Max;
		XYZ[] obj = new XYZ[8]
		{
			new XYZ(min.X, min.Y, min.Z),
			new XYZ(min.X, min.Y, max.Z),
			new XYZ(min.X, max.Y, min.Z),
			new XYZ(min.X, max.Y, max.Z),
			new XYZ(max.X, min.Y, min.Z),
			new XYZ(max.X, min.Y, max.Z),
			new XYZ(max.X, max.Y, min.Z),
			new XYZ(max.X, max.Y, max.Z)
		};
		BoundingBoxXYZ result = null;
		XYZ[] array = obj;
		foreach (XYZ point in array)
		{
			result = UnionPoint(result, activeTransform.OfPoint(point));
		}
		return result;
	}

	private static bool IsPhysicalPreviewElement(Element element)
	{
		if (element == null)
		{
			return false;
		}
		if (element is ReferencePlane || element is Level || element is View)
		{
			return false;
		}
		if (IsConnectorElement(element) || IsReferenceDatumCategory(element.Category))
		{
			return false;
		}
		return true;
	}

	private static bool IsConnectorElement(Element element)
	{
		if (element == null)
		{
			return false;
		}
		if (element.GetType().Name.IndexOf("Connector", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return true;
		}
		return IsConnectorCategory(element.Category);
	}

	private static bool IsConnectorCategory(Category category)
	{
		if (category == null)
		{
			return false;
		}
		return (category.Name ?? string.Empty).IndexOf("connector", StringComparison.OrdinalIgnoreCase) >= 0 || IsBuiltInCategoryName(category, "OST_ConnectorElem");
	}

	private static bool IsReferenceDatumCategory(Category category)
	{
		if (category == null)
		{
			return false;
		}
		string name = category.Name ?? string.Empty;
		return name.IndexOf("reference plane", StringComparison.OrdinalIgnoreCase) >= 0 || name.IndexOf("reference line", StringComparison.OrdinalIgnoreCase) >= 0 || name.IndexOf("level", StringComparison.OrdinalIgnoreCase) >= 0;
	}

	private static bool IsBuiltInCategoryName(Category category, params string[] builtInNames)
	{
		bool IsBuiltInCategoryName;
		if (category == null || builtInNames == null || builtInNames.Length == 0)
		{
			IsBuiltInCategoryName = false;
		}
		else
		{
			try
			{
				string a = ((BuiltInCategory)RevitElementIdCompat.CompatIntegerValue(category.Id)/*cast due to .constrained prefix*/).ToString();
				IsBuiltInCategoryName = builtInNames.Any([SpecialName] (string x) => string.Equals(a, x, StringComparison.OrdinalIgnoreCase));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				IsBuiltInCategoryName = false;
				ProjectData.ClearProjectError();
			}
		}
		return IsBuiltInCategoryName;
	}

	private static bool IsUsableBounds(BoundingBoxXYZ bounds)
	{
		if (bounds == null || bounds.Min == null || bounds.Max == null)
		{
			return false;
		}
		return bounds.Max.X > bounds.Min.X && bounds.Max.Y > bounds.Min.Y && bounds.Max.Z > bounds.Min.Z;
	}

	private static BoundingBoxXYZ CloneBounds(BoundingBoxXYZ bounds)
	{
		if (bounds == null)
		{
			return null;
		}
		return new BoundingBoxXYZ
		{
			Min = new XYZ(bounds.Min.X, bounds.Min.Y, bounds.Min.Z),
			Max = new XYZ(bounds.Max.X, bounds.Max.Y, bounds.Max.Z)
		};
	}

	private static BoundingBoxXYZ NormalizeBoundsToModelCoordinates(BoundingBoxXYZ bounds)
	{
		BoundingBoxXYZ NormalizeBoundsToModelCoordinates;
		if (bounds == null || !IsUsableBounds(bounds))
		{
			NormalizeBoundsToModelCoordinates = null;
		}
		else
		{
			Transform transform = null;
			try
			{
				transform = bounds.Transform;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				transform = null;
				ProjectData.ClearProjectError();
			}
			if (transform == null)
			{
				NormalizeBoundsToModelCoordinates = CloneBounds(bounds);
			}
			else
			{
				try
				{
					XYZ min = bounds.Min;
					XYZ max = bounds.Max;
					XYZ[] obj = new XYZ[8]
					{
						new XYZ(min.X, min.Y, min.Z),
						new XYZ(min.X, min.Y, max.Z),
						new XYZ(min.X, max.Y, min.Z),
						new XYZ(min.X, max.Y, max.Z),
						new XYZ(max.X, min.Y, min.Z),
						new XYZ(max.X, min.Y, max.Z),
						new XYZ(max.X, max.Y, min.Z),
						new XYZ(max.X, max.Y, max.Z)
					};
					BoundingBoxXYZ normalized = null;
					XYZ[] array = obj;
					foreach (XYZ point in array)
					{
						normalized = UnionPoint(normalized, transform.OfPoint(point));
					}
					NormalizeBoundsToModelCoordinates = normalized ?? CloneBounds(bounds);
				}
				catch (Exception projectError2)
				{
					ProjectData.SetProjectError(projectError2);
					NormalizeBoundsToModelCoordinates = CloneBounds(bounds);
					ProjectData.ClearProjectError();
				}
			}
		}
		return NormalizeBoundsToModelCoordinates;
	}

	private static BoundingBoxXYZ UnionPoint(BoundingBoxXYZ bounds, XYZ point)
	{
		if (point == null)
		{
			return bounds;
		}
		if (bounds == null)
		{
			return new BoundingBoxXYZ
			{
				Min = new XYZ(point.X, point.Y, point.Z),
				Max = new XYZ(point.X, point.Y, point.Z)
			};
		}
		return new BoundingBoxXYZ
		{
			Min = new XYZ(Math.Min(bounds.Min.X, point.X), Math.Min(bounds.Min.Y, point.Y), Math.Min(bounds.Min.Z, point.Z)),
			Max = new XYZ(Math.Max(bounds.Max.X, point.X), Math.Max(bounds.Max.Y, point.Y), Math.Max(bounds.Max.Z, point.Z))
		};
	}

	private static BoundingBoxXYZ UnionBounds(BoundingBoxXYZ first, BoundingBoxXYZ second)
	{
		if (first == null)
		{
			return CloneBounds(second);
		}
		if (second == null)
		{
			return CloneBounds(first);
		}
		return new BoundingBoxXYZ
		{
			Min = new XYZ(Math.Min(first.Min.X, second.Min.X), Math.Min(first.Min.Y, second.Min.Y), Math.Min(first.Min.Z, second.Min.Z)),
			Max = new XYZ(Math.Max(first.Max.X, second.Max.X), Math.Max(first.Max.Y, second.Max.Y), Math.Max(first.Max.Z, second.Max.Z))
		};
	}

	private static BoundingBoxXYZ ExpandBounds(BoundingBoxXYZ bounds)
	{
		if (bounds == null)
		{
			return null;
		}
		XYZ min = bounds.Min;
		XYZ max = bounds.Max;
		double sizeX = Math.Max(max.X - min.X, 0.5);
		double sizeY = Math.Max(max.Y - min.Y, 0.5);
		double sizeZ = Math.Max(max.Z - min.Z, 0.5);
		XYZ center = new XYZ((min.X + max.X) / 2.0, (min.Y + max.Y) / 2.0, (min.Z + max.Z) / 2.0);
		double padding = Math.Max(Math.Max(sizeX, Math.Max(sizeY, sizeZ)) * 1.15, 1.5);
		return new BoundingBoxXYZ
		{
			Min = new XYZ(center.X - sizeX / 2.0 - padding, center.Y - sizeY / 2.0 - padding, center.Z - sizeZ / 2.0 - padding),
			Max = new XYZ(center.X + sizeX / 2.0 + padding, center.Y + sizeY / 2.0 + padding, center.Z + sizeZ / 2.0 + padding)
		};
	}

	private static double BoundsDiagonal(BoundingBoxXYZ bounds)
	{
		if (bounds == null || !IsUsableBounds(bounds))
		{
			return 0.0;
		}
		double num = bounds.Max.X - bounds.Min.X;
		double sizeY = bounds.Max.Y - bounds.Min.Y;
		double sizeZ = bounds.Max.Z - bounds.Min.Z;
		return Math.Sqrt(num * num + sizeY * sizeY + sizeZ * sizeZ);
	}

	private static void ApplyPreviewFocus(View3D view, BoundingBoxXYZ previewBounds)
	{
		try
		{
			view.IsSectionBoxActive = false;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			view.CropBoxActive = false;
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		try
		{
			SetDefaultIsometricOrientation(view, previewBounds);
		}
		catch (Exception projectError3)
		{
			ProjectData.SetProjectError(projectError3);
			ProjectData.ClearProjectError();
		}
	}

	private static void SetDefaultIsometricOrientation(View3D view, BoundingBoxXYZ previewBounds)
	{
		XYZ target = new XYZ(0.0, 0.0, 0.0);
		double distance = 10.0;
		if (previewBounds != null && IsUsableBounds(previewBounds))
		{
			target = new XYZ((previewBounds.Min.X + previewBounds.Max.X) / 2.0, (previewBounds.Min.Y + previewBounds.Max.Y) / 2.0, (previewBounds.Min.Z + previewBounds.Max.Z) / 2.0);
			double num = previewBounds.Max.X - previewBounds.Min.X;
			double sizeY = previewBounds.Max.Y - previewBounds.Min.Y;
			double sizeZ = previewBounds.Max.Z - previewBounds.Min.Z;
			distance = Math.Max(Math.Sqrt(num * num + sizeY * sizeY + sizeZ * sizeZ) * 1.4, 4.0);
		}
		XYZ eye = new XYZ(target.X + distance, target.Y - distance, target.Z + distance * 0.8);
		XYZ forward = target.Subtract(eye).Normalize();
		XYZ worldUp = new XYZ(0.0, 0.0, 1.0);
		XYZ up = forward.CrossProduct(worldUp).Normalize().CrossProduct(forward)
			.Normalize();
		view.SetOrientation(new ViewOrientation3D(eye, up, forward));
	}

	private static void ExportViewImage(Document familyDocument, ElementId viewId, string basePath)
	{
		ExportViewImage(familyDocument, viewId, basePath, FitDirectionType.Horizontal);
	}

	private static void ExportViewImage(Document familyDocument, ElementId viewId, string basePath, FitDirectionType fitDirection)
	{
		List<ElementId> viewIds = new List<ElementId> { viewId };
		ImageExportOptions options = new ImageExportOptions
		{
			ExportRange = ExportRange.SetOfViews,
			FilePath = basePath,
			FitDirection = fitDirection,
			HLRandWFViewsFileType = ImageFileType.PNG,
			ImageResolution = ResolveImageResolution("DPI_72", ImageResolution.DPI_72),
			PixelSize = 768,
			ShadowViewsFileType = ImageFileType.PNG,
			ZoomType = ZoomFitType.FitToPage
		};
		options.SetViewsAndSheets(viewIds);
		Autodesk.Revit.DB.Color originalBackground;
		bool backgroundChanged = TrySetRevitApplicationBackground(familyDocument, new Autodesk.Revit.DB.Color(byte.MaxValue, byte.MaxValue, byte.MaxValue), out originalBackground);
		try
		{
			familyDocument.ExportImage(options);
		}
		finally
		{
			if (backgroundChanged)
			{
				TryRestoreRevitApplicationBackground(familyDocument, originalBackground);
			}
		}
	}

	private static bool TrySetRevitApplicationBackground(Document document, Autodesk.Revit.DB.Color color, out Autodesk.Revit.DB.Color originalBackground)
	{
		originalBackground = null;
		if (document == null || color == null)
		{
			return false;
		}
		try
		{
			Autodesk.Revit.ApplicationServices.Application application = document.Application;
			if (application == null)
			{
				return false;
			}
			originalBackground = application.BackgroundColor;
			application.BackgroundColor = color;
			return true;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return false;
	}

	private static void TryRestoreRevitApplicationBackground(Document document, Autodesk.Revit.DB.Color originalBackground)
	{
		if (document == null || originalBackground == null)
		{
			return;
		}
		try
		{
			Autodesk.Revit.ApplicationServices.Application application = document.Application;
			if (application != null)
			{
				application.BackgroundColor = originalBackground;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static ImageResolution ResolveImageResolution(string preferredName, ImageResolution fallback)
	{
		if (!string.IsNullOrWhiteSpace(preferredName))
		{
			try
			{
				return (ImageResolution)Enum.Parse(typeof(ImageResolution), preferredName, ignoreCase: true);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		return fallback;
	}

	private static bool SaveWhiteBackgroundPreviewImage(string sourcePath, string targetPath)
	{
		checked
		{
			bool SaveWhiteBackgroundPreviewImage;
			if (string.IsNullOrWhiteSpace(sourcePath) || string.IsNullOrWhiteSpace(targetPath))
			{
				SaveWhiteBackgroundPreviewImage = false;
			}
			else
			{
				try
				{
					using (Bitmap source = new Bitmap(sourcePath))
					{
						int canvasWidth = source.Width;
						int canvasHeight = source.Height;
						if (canvasWidth > 0 && canvasHeight > 0)
						{
							using (Bitmap normalizedSource = CreateCenteredWhiteBackgroundPreviewBitmap(source))
							{
								normalizedSource.Save(targetPath, ImageFormat.Png);
							}
							goto end_IL_001f;
						}
						SaveWhiteBackgroundPreviewImage = false;
						goto end_IL_0018;
						end_IL_001f:;
					}
					SaveWhiteBackgroundPreviewImage = true;
					end_IL_0018:;
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					SaveWhiteBackgroundPreviewImage = false;
					ProjectData.ClearProjectError();
				}
			}
			return SaveWhiteBackgroundPreviewImage;
		}
	}

	private static double ClampDouble(double value, double minimum, double maximum)
	{
		if (minimum > maximum)
		{
			double num = minimum;
			minimum = maximum;
			maximum = num;
		}
		return Math.Max(minimum, Math.Min(maximum, value));
	}

	private static bool PreviewImageTouchesEdge(string imagePath)
	{
		if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
		{
			return false;
		}
		try
		{
			using Bitmap image = new Bitmap(imagePath);
			return PreviewContentTouchesImageEdge(FindPreviewContentBounds(image), image.Width, image.Height);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return false;
	}

	private static double ComputePreviewFrameMargin(string imagePath)
	{
		if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
		{
			return -1.0;
		}
		try
		{
			using Bitmap image = new Bitmap(imagePath);
			System.Drawing.Rectangle contentBounds = FindPreviewContentBounds(image);
			if (contentBounds.IsEmpty)
			{
				return -1.0;
			}
			double left = contentBounds.Left;
			double top = contentBounds.Top;
			double right = image.Width - contentBounds.Right;
			double bottom = image.Height - contentBounds.Bottom;
			return Math.Min(Math.Min(left, top), Math.Min(right, bottom));
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return -1.0;
	}

	private static double ComputePreviewFrameMargin(System.Drawing.Rectangle contentBounds, int imageWidth, int imageHeight)
	{
		if (contentBounds.IsEmpty || imageWidth <= 0 || imageHeight <= 0)
		{
			return -1.0;
		}
		double left = contentBounds.Left;
		double top = contentBounds.Top;
		double right = imageWidth - contentBounds.Right;
		double bottom = imageHeight - contentBounds.Bottom;
		return Math.Min(Math.Min(left, top), Math.Min(right, bottom));
	}

	private static double ComputeRequiredPreviewFrameMargin(string imagePath)
	{
		if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
		{
			return PreviewMinimumFrameMarginPixels;
		}
		try
		{
			using Bitmap image = new Bitmap(imagePath);
			return ComputeRequiredPreviewFrameMargin(image.Width, image.Height);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return PreviewMinimumFrameMarginPixels;
	}

	private static double ComputeRequiredPreviewFrameMargin(int imageWidth, int imageHeight)
	{
		if (imageWidth <= 0 || imageHeight <= 0)
		{
			return PreviewMinimumFrameMarginPixels;
		}
		return Math.Max(PreviewMinimumFrameMarginPixels, (double)Math.Min(imageWidth, imageHeight) * PreviewMinimumFrameMarginRatio);
	}

	private static void TryDeletePreviewFile(string imagePath)
	{
		if (string.IsNullOrWhiteSpace(imagePath))
		{
			return;
		}
		try
		{
			if (File.Exists(imagePath))
			{
				File.Delete(imagePath);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static Bitmap CreateCenteredWhiteBackgroundPreviewBitmap(Bitmap source)
	{
		int canvasWidth = source.Width;
		int canvasHeight = source.Height;
		System.Drawing.Rectangle contentBounds = FindPreviewContentBounds(source);
		double scale = 0.94;
		int drawWidth = Math.Max(1, (int)Math.Round((double)canvasWidth * scale));
		int drawHeight = Math.Max(1, (int)Math.Round((double)canvasHeight * scale));
		double baseX = (double)(canvasWidth - drawWidth) / 2.0;
		double baseY = (double)(canvasHeight - drawHeight) / 2.0;
		double shiftX = 0.0;
		double shiftY = 0.0;
		if (contentBounds.Width > 0 && contentBounds.Height > 0)
		{
			double contentCenterX = (double)contentBounds.Left + (double)contentBounds.Width / 2.0;
			double contentCenterY = (double)contentBounds.Top + (double)contentBounds.Height / 2.0;
			shiftX = (double)canvasWidth / 2.0 - (baseX + contentCenterX * scale);
			shiftY = (double)canvasHeight / 2.0 - (baseY + contentCenterY * scale);
			shiftX = ClampDouble(shiftX, 0.0 - baseX, baseX);
			shiftY = ClampDouble(shiftY, 0.0 - baseY, baseY);
		}
		int drawX = (int)Math.Round(baseX + shiftX);
		int drawY = (int)Math.Round(baseY + shiftY);
		Bitmap normalized = new Bitmap(canvasWidth, canvasHeight, PixelFormat.Format32bppArgb);
		using (Graphics graphics = Graphics.FromImage(normalized))
		{
			graphics.Clear(System.Drawing.Color.White);
			graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
			graphics.SmoothingMode = SmoothingMode.HighQuality;
			graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
			graphics.DrawImage(source, new System.Drawing.Rectangle(drawX, drawY, drawWidth, drawHeight), new System.Drawing.Rectangle(0, 0, canvasWidth, canvasHeight), GraphicsUnit.Pixel);
		}
		return normalized;
	}

	private static System.Drawing.Rectangle FindPreviewContentBounds(Bitmap image)
	{
		if (image == null || image.Width <= 0 || image.Height <= 0)
		{
			return System.Drawing.Rectangle.Empty;
		}
		System.Drawing.Color background = EstimatePreviewBackgroundColor(image);
		int left = image.Width;
		int top = image.Height;
		int right = -1;
		int bottom = -1;
		checked
		{
			using (Bitmap scanImage = new Bitmap(image.Width, image.Height, PixelFormat.Format32bppArgb))
			{
				using (Graphics graphics = Graphics.FromImage(scanImage))
				{
					graphics.DrawImageUnscaled(image, 0, 0);
				}
				BitmapData data = null;
				try
				{
					System.Drawing.Rectangle rect = new System.Drawing.Rectangle(0, 0, scanImage.Width, scanImage.Height);
					data = scanImage.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
					int stride = Math.Abs(data.Stride);
					byte[] bytes = new byte[stride * scanImage.Height - 1 + 1];
					Marshal.Copy(data.Scan0, bytes, 0, bytes.Length);
					int num = scanImage.Height - 1;
					for (int y = 0; y <= num; y++)
					{
						int rowOffset = ((data.Stride >= 0) ? (y * data.Stride) : ((scanImage.Height - 1 - y) * stride));
						int num2 = scanImage.Width - 1;
						for (int x = 0; x <= num2; x++)
						{
							int offset = rowOffset + x * 4;
							if (IsPreviewContentPixel(System.Drawing.Color.FromArgb(bytes[offset + 3], bytes[offset + 2], bytes[offset + 1], bytes[offset]), background))
							{
								if (x < left)
								{
									left = x;
								}
								if (x > right)
								{
									right = x;
								}
								if (y < top)
								{
									top = y;
								}
								if (y > bottom)
								{
									bottom = y;
								}
							}
						}
					}
				}
				finally
				{
					if (data != null)
					{
						try
						{
							scanImage.UnlockBits(data);
						}
						catch (Exception projectError)
						{
							ProjectData.SetProjectError(projectError);
							ProjectData.ClearProjectError();
						}
					}
				}
			}
			if (right < left || bottom < top)
			{
				return System.Drawing.Rectangle.Empty;
			}
			int inflate = Math.Max(24, (int)Math.Round((double)Math.Min(image.Width, image.Height) * 0.1));
			left = Math.Max(0, left - inflate);
			top = Math.Max(0, top - inflate);
			right = Math.Min(image.Width - 1, right + inflate);
			bottom = Math.Min(image.Height - 1, bottom + inflate);
			return System.Drawing.Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
		}
	}

	private static bool PreviewContentTouchesImageEdge(System.Drawing.Rectangle contentBounds, int imageWidth, int imageHeight)
	{
		if (imageWidth <= 0 || imageHeight <= 0 || contentBounds.IsEmpty)
		{
			return false;
		}
		checked
		{
			int tolerance = Math.Max(8, (int)Math.Round((double)Math.Min(imageWidth, imageHeight) * 0.025));
			return contentBounds.Left <= tolerance || contentBounds.Top <= tolerance || contentBounds.Right >= imageWidth - tolerance || contentBounds.Bottom >= imageHeight - tolerance;
		}
	}

	private static System.Drawing.Color EstimatePreviewBackgroundColor(Bitmap image)
	{
		if (image == null || image.Width <= 0 || image.Height <= 0)
		{
			return System.Drawing.Color.White;
		}
		try
		{
			Dictionary<int, int> histogram = new Dictionary<int, int>();
			int step = Math.Max(1, Math.Min(image.Width, image.Height) / 160);
			for (int x = 0; x < image.Width; x += step)
			{
				AddPreviewBackgroundSample(histogram, image.GetPixel(x, 0));
				AddPreviewBackgroundSample(histogram, image.GetPixel(x, image.Height - 1));
			}
			for (int y = 0; y < image.Height; y += step)
			{
				AddPreviewBackgroundSample(histogram, image.GetPixel(0, y));
				AddPreviewBackgroundSample(histogram, image.GetPixel(image.Width - 1, y));
			}
			if (histogram.Count > 0)
			{
				int key = histogram.OrderByDescending([SpecialName] (KeyValuePair<int, int> x) => x.Value).First().Key;
				return ColorFromPreviewBackgroundKey(key);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return System.Drawing.Color.White;
	}

	private static void AddPreviewBackgroundSample(Dictionary<int, int> histogram, System.Drawing.Color color)
	{
		if (histogram == null || color.A <= 12)
		{
			return;
		}
		int a = QuantizePreviewColorChannel(color.A);
		int r = QuantizePreviewColorChannel(color.R);
		int g = QuantizePreviewColorChannel(color.G);
		int b = QuantizePreviewColorChannel(color.B);
		int key = unchecked((a << 24) | (r << 16) | (g << 8) | b);
		int count;
		histogram.TryGetValue(key, out count);
		histogram[key] = count + 1;
	}

	private static int QuantizePreviewColorChannel(int value)
	{
		value = Math.Max(0, Math.Min(255, value));
		return Math.Max(0, Math.Min(255, value / 16 * 16));
	}

	private static System.Drawing.Color ColorFromPreviewBackgroundKey(int key)
	{
		int a = (key >> 24) & 0xFF;
		int r = (key >> 16) & 0xFF;
		int g = (key >> 8) & 0xFF;
		int b = key & 0xFF;
		return System.Drawing.Color.FromArgb(a, r, g, b);
	}

	private static bool IsPreviewContentPixel(System.Drawing.Color pixel, System.Drawing.Color background)
	{
		if (pixel.A <= 12)
		{
			return false;
		}
		if (checked(Math.Abs(pixel.R - background.R) <= 8 && Math.Abs(pixel.G - background.G) <= 8 && Math.Abs(pixel.B - background.B) <= 8))
		{
			return false;
		}
		return true;
	}

	private static string FindExportedImage(string outputFolder, string fileStem, IDictionary<string, PreviewExportFileState> beforeState, DateTime exportStartedUtc)
	{
		if (string.IsNullOrWhiteSpace(outputFolder) || !Directory.Exists(outputFolder))
		{
			return string.Empty;
		}
		string directPath = Path.Combine(outputFolder, fileStem + ".png");
		if (IsUsablePngFile(directPath))
		{
			return directPath;
		}
		FileInfo stemMatch = (from x in (from x in Directory.GetFiles(outputFolder, fileStem + "*.png")
				select new FileInfo(x) into x
				where x.Exists && x.Length > 0
				select x).ToList()
			where Path.GetFileNameWithoutExtension(x.Name).StartsWith(fileStem, StringComparison.OrdinalIgnoreCase)
			orderby x.LastWriteTimeUtc descending
			select x).FirstOrDefault();
		if (stemMatch != null)
		{
			return stemMatch.FullName;
		}
		if (beforeState == null)
		{
			return string.Empty;
		}
		FileInfo changedExport = (from x in (from x in Directory.GetFiles(outputFolder, "*.png")
				select new FileInfo(x) into x
				where x.Exists && x.Length > 0
				select x).ToList()
			where IsNewOrUpdatedPng(x, beforeState, exportStartedUtc)
			orderby x.LastWriteTimeUtc descending
			select x).FirstOrDefault();
		if (changedExport != null)
		{
			return changedExport.FullName;
		}
		return string.Empty;
	}

	private static Dictionary<string, PreviewExportFileState> CapturePngFileState(string outputFolder)
	{
		Dictionary<string, PreviewExportFileState> result = new Dictionary<string, PreviewExportFileState>(StringComparer.OrdinalIgnoreCase);
		if (string.IsNullOrWhiteSpace(outputFolder) || !Directory.Exists(outputFolder))
		{
			return result;
		}
		string[] files = Directory.GetFiles(outputFolder, "*.png");
		foreach (string filePath in files)
		{
			try
			{
				FileInfo info = new FileInfo(filePath);
				result[info.FullName] = new PreviewExportFileState
				{
					FullName = info.FullName,
					Length = info.Length,
					LastWriteTimeUtc = info.LastWriteTimeUtc
				};
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		return result;
	}

	private static bool IsUsablePngFile(string filePath)
	{
		bool IsUsablePngFile;
		if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
		{
			IsUsablePngFile = false;
		}
		else
		{
			try
			{
				IsUsablePngFile = new FileInfo(filePath).Length > 0;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				IsUsablePngFile = false;
				ProjectData.ClearProjectError();
			}
		}
		return IsUsablePngFile;
	}

	private static bool IsNewOrUpdatedPng(FileInfo fileInfo, IDictionary<string, PreviewExportFileState> beforeState, DateTime exportStartedUtc)
	{
		if (fileInfo == null || !fileInfo.Exists || fileInfo.Length <= 0)
		{
			return false;
		}
		PreviewExportFileState previous = null;
		if (beforeState == null || !beforeState.TryGetValue(fileInfo.FullName, out previous))
		{
			return DateTime.Compare(exportStartedUtc, DateTime.MinValue) == 0 || DateTime.Compare(fileInfo.LastWriteTimeUtc, exportStartedUtc) >= 0;
		}
		return fileInfo.Length != previous.Length || DateTime.Compare(fileInfo.LastWriteTimeUtc, previous.LastWriteTimeUtc) != 0;
	}

	private static string BuildMissingExportDiagnostic(string outputFolder, string fileStem, IDictionary<string, PreviewExportFileState> beforeState, DateTime exportStartedUtc)
	{
		Dictionary<string, PreviewExportFileState> afterState = CapturePngFileState(outputFolder);
		List<string> recentPngs = (from x in afterState.Values.OrderByDescending([SpecialName] (PreviewExportFileState x) => x.LastWriteTimeUtc).Take(8)
			select Path.GetFileName(x.FullName) + " (" + x.Length.ToString(CultureInfo.InvariantCulture) + " bytes, " + x.LastWriteTimeUtc.ToString("u", CultureInfo.InvariantCulture) + ")").ToList();
		StringBuilder message = new StringBuilder();
		message.Append("Preview export finished, but no PNG file was produced or discovered.");
		message.Append(" Output folder=" + (outputFolder ?? string.Empty));
		message.Append("; tempStem=" + (fileStem ?? string.Empty));
		message.Append("; pngBefore=" + (beforeState?.Count ?? 0).ToString(CultureInfo.InvariantCulture));
		message.Append("; pngAfter=" + afterState.Count.ToString(CultureInfo.InvariantCulture));
		if (DateTime.Compare(exportStartedUtc, DateTime.MinValue) != 0)
		{
			message.Append("; exportStartedUtc=" + exportStartedUtc.ToString("u", CultureInfo.InvariantCulture));
		}
		if (recentPngs.Count > 0)
		{
			message.Append("; recentPngs=" + string.Join(", ", recentPngs));
		}
		return message.ToString();
	}

	private static void RunPreviewStage(FamilyThumbnailGenerationResult generation, string stageName, Action action)
	{
		try
		{
			action();
			generation?.Steps.Add(stageName);
		}
		catch (FamilyThumbnailStageException ex)
		{
			ProjectData.SetProjectError(ex);
			FamilyThumbnailStageException ex2 = ex;
			throw;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			throw new FamilyThumbnailStageException(stageName, ex4);
		}
	}

	private static string AppendAutoConfirmedDialogMessage(string message, IEnumerable<FamilyThumbnailAutoConfirmedDialogRecord> records)
	{
		List<FamilyThumbnailAutoConfirmedDialogRecord> list = (records ?? new List<FamilyThumbnailAutoConfirmedDialogRecord>()).ToList();
		if (list.Count == 0)
		{
			return message ?? string.Empty;
		}
		List<string> reasons = (from x in list.Select([SpecialName] (FamilyThumbnailAutoConfirmedDialogRecord x) =>
			{
				string text;
				if (x != null)
				{
					text = x.Reason;
					if (text == null)
					{
						return string.Empty;
					}
				}
				else
				{
					text = string.Empty;
				}
				return text;
			})
			where !string.IsNullOrWhiteSpace(x)
			select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
		string suffix = " Auto-confirmed Revit warning dialog(s): " + list.Count.ToString(CultureInfo.InvariantCulture) + ((reasons.Count == 0) ? string.Empty : (" (" + string.Join(", ", reasons) + ")")) + ".";
		if (string.IsNullOrWhiteSpace(message))
		{
			return suffix.Trim();
		}
		return message.TrimEnd() + suffix;
	}

	private static string BuildAutoConfirmedDialogSummary(FamilyThumbnailAutoConfirmedDialogRecord record)
	{
		if (record == null)
		{
			return string.Empty;
		}
		return (string.IsNullOrWhiteSpace(record.Reason) ? "RevitWarning" : record.Reason) + " / result=" + (string.IsNullOrWhiteSpace(record.OverrideResult) ? "-" : record.OverrideResult);
	}

	public static string SaveBatchDiagnosticReport(FamilyThumbnailBatchUpdateResult result, string outputPath)
	{
		if (result == null)
		{
			throw new ArgumentNullException("result");
		}
		if (string.IsNullOrWhiteSpace(outputPath))
		{
			throw new ArgumentException(FamilyBrowserLanguageService.Text("Output path is empty.", "출력 경로가 비어 있습니다."), "outputPath");
		}
		string outputFolder = Path.GetDirectoryName(outputPath);
		if (!string.IsNullOrWhiteSpace(outputFolder))
		{
			Directory.CreateDirectory(outputFolder);
		}
		File.WriteAllText(outputPath, BuildBatchDiagnosticText(result), Encoding.UTF8);
		return outputPath;
	}

	private static string BuildBatchDiagnosticText(FamilyThumbnailBatchUpdateResult result)
	{
		StringBuilder sb = new StringBuilder();
		sb.AppendLine("KKY Family Browser 3D Thumbnail Diagnostics");
		sb.AppendLine("Created UTC: " + DateTime.UtcNow.ToString("u", CultureInfo.InvariantCulture));
		sb.AppendLine("Process: open standard RVT in background; open family with EditFamily; auto-confirm OK-continuable Revit family geometry/constraint warnings only during thumbnail generation; create new 3D preview view; set shaded/fine preview; show model and connector preview categories; hide annotations/reference datum; apply category and element thin line overrides; apply transparent connector fill override; focus home/isometric view around physical model geometry; export PNG with short temp names and discover Revit-renamed/new PNG files.");
		sb.AppendLine("Success: " + (result?.SuccessCount ?? 0).ToString(CultureInfo.InvariantCulture));
		sb.AppendLine("Failed: " + (result?.FailedCount ?? 0).ToString(CultureInfo.InvariantCulture));
		sb.AppendLine("Skipped: " + (result?.SkippedCount ?? 0).ToString(CultureInfo.InvariantCulture));
		sb.AppendLine("Auto-confirmed dialogs: " + ((result != null && result.AutoConfirmedDialogs != null) ? result.AutoConfirmedDialogs.Count : 0).ToString(CultureInfo.InvariantCulture));
		sb.AppendLine();
		if (result != null && result.Items != null)
		{
			foreach (FamilyThumbnailBatchUpdateItem item in result.Items)
			{
				sb.AppendLine((item.Success ? "OK" : (item.Skipped ? "SKIP" : "FAIL")) + " | " + item.CategoryName + " | " + item.FamilyName);
				if (!string.IsNullOrWhiteSpace(item.Message))
				{
					sb.AppendLine("  " + item.Message.Replace(Environment.NewLine, " "));
				}
				if (item.AutoConfirmedDialogs != null && item.AutoConfirmedDialogs.Count > 0)
				{
					foreach (string dialogSummary in item.AutoConfirmedDialogs)
					{
						sb.AppendLine("  Auto-confirmed: " + (dialogSummary ?? string.Empty));
					}
				}
				if (!string.IsNullOrWhiteSpace(item.ImagePath))
				{
					sb.AppendLine("  Image: " + item.ImagePath);
				}
			}
		}
		if (result != null && result.AutoConfirmedDialogs != null && result.AutoConfirmedDialogs.Count > 0)
		{
			sb.AppendLine();
			sb.AppendLine("Auto-confirmed Revit dialogs");
			foreach (FamilyThumbnailAutoConfirmedDialogRecord record in result.AutoConfirmedDialogs)
			{
				if (record != null)
				{
					sb.AppendLine("- " + (string.IsNullOrWhiteSpace(record.CategoryName) ? "-" : record.CategoryName) + " | " + (string.IsNullOrWhiteSpace(record.FamilyName) ? "-" : record.FamilyName) + " | " + (string.IsNullOrWhiteSpace(record.Reason) ? "-" : record.Reason) + " | result=" + (string.IsNullOrWhiteSpace(record.OverrideResult) ? "-" : record.OverrideResult) + " | utc=" + (string.IsNullOrWhiteSpace(record.ConfirmedAtUtc) ? "-" : record.ConfirmedAtUtc));
					if (!string.IsNullOrWhiteSpace(record.DialogText))
					{
						sb.AppendLine("  Dialog: " + CompactDiagnosticText(record.DialogText));
					}
				}
			}
		}
		return sb.ToString();
	}

	private static string CompactDiagnosticText(string value)
	{
		string text = (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
		while (text.Contains("  "))
		{
			text = text.Replace("  ", " ");
		}
		if (text.Length > 900)
		{
			return text.Substring(0, 900) + "...";
		}
		return text;
	}

	private static void WriteFailureMarker(string imagePath, string message)
	{
		string markerPath = BuildFailureMarkerPath(imagePath);
		if (!string.IsNullOrWhiteSpace(markerPath))
		{
			try
			{
				Directory.CreateDirectory(Path.GetDirectoryName(markerPath));
				File.WriteAllText(markerPath, message ?? string.Empty, Encoding.UTF8);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
	}

	private static void DeleteFailureMarker(string imagePath)
	{
		string markerPath = BuildFailureMarkerPath(imagePath);
		if (!string.IsNullOrWhiteSpace(markerPath) && File.Exists(markerPath))
		{
			try
			{
				File.Delete(markerPath);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
	}

	private static string BuildFailureMarkerPath(string imagePath)
	{
		if (string.IsNullOrWhiteSpace(imagePath))
		{
			return string.Empty;
		}
		return imagePath + ".fail.txt";
	}

	private static bool IsFamilyEditable(Family family)
	{
		bool IsFamilyEditable;
		try
		{
			IsFamilyEditable = family.IsEditable;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			IsFamilyEditable = true;
			ProjectData.ClearProjectError();
		}
		return IsFamilyEditable;
	}

	private static string BuildCacheFileName(string categoryName, string familyName)
	{
		return SafeFileName(categoryName + "__" + familyName) + ".png";
	}

	private static string SafeFileName(string value, int maxLength = 120)
	{
		if (string.IsNullOrWhiteSpace(value))
		{
			return "family_preview";
		}
		char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
		string safe = new string(value.Select([SpecialName] (char ch) => (!Enumerable.Contains(invalidFileNameChars, ch)) ? ch : '_').ToArray()).Trim().TrimEnd(new char[2] { '.', ' ' });
		if (string.IsNullOrWhiteSpace(safe))
		{
			safe = "family_preview";
		}
		if (maxLength <= 0 || safe.Length <= maxLength)
		{
			return safe;
		}
		string suffix = "_" + StableFileNameHash(safe);
		int prefixLength = Math.Max(1, checked(maxLength - suffix.Length));
		string prefix = safe.Substring(0, Math.Min(prefixLength, safe.Length)).Trim().TrimEnd(new char[2] { '.', ' ' });
		if (string.IsNullOrWhiteSpace(prefix))
		{
			prefix = "family";
		}
		return prefix + suffix;
	}

	private static string BuildTempExportStem()
	{
		return "kky_thumb_" + DateTime.UtcNow.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture) + "_" + Guid.NewGuid().ToString("N").Substring(0, 8);
	}

	private static string StableFileNameHash(string value)
	{
		using SHA256 sha = SHA256.Create();
		byte[] bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
		return BitConverter.ToString(sha.ComputeHash(bytes), 0, 4).Replace("-", string.Empty);
	}

	private static string HashString(string value)
	{
		using SHA256 sha = SHA256.Create();
		byte[] bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
		return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant();
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
