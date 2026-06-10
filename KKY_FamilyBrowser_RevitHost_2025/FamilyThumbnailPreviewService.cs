using System;
using System.Collections;
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
				return _0024VB_0024Local_selectedNameSet.Contains(Normalize(((Element)x).Name ?? string.Empty));
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
				return _0024VB_0024Local_selectedNameSet.Contains(Normalize(((Element)x).Name ?? string.Empty));
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

	private const string PreviewCacheVersionFolderName = "preview-v6-no-section-safe-recenter";

	private const int PreviewExportPixelSize = 768;

	private const double PreviewSafeRecenterScale = 0.94;

	private const string PreviewTempFileStemPrefix = "kky_thumb_";

	private FamilyThumbnailPreviewService()
	{
	}

	public static string GetCacheFolder(string workspaceRoot, string sourceId)
	{
		return Path.Combine(FamilyBrowserStandardPolicyStore.GetThumbnailFolder(workspaceRoot), SafeFileName(sourceId), "preview-v6-no-section-safe-recenter");
	}

	public static string GetCachedImagePath(string workspaceRoot, string sourceId, string categoryName, string familyName)
	{
		return Path.Combine(GetCacheFolder(workspaceRoot, sourceId), BuildCacheFileName(categoryName, familyName));
	}

	public static string ResolveExistingCachedImagePath(string workspaceRoot, string sourceId, string categoryName, string familyName)
	{
		string ResolveExistingCachedImagePath;
		string expected;
		if (string.IsNullOrWhiteSpace(sourceId) || string.IsNullOrWhiteSpace(familyName))
		{
			ResolveExistingCachedImagePath = string.Empty;
		}
		else
		{
			try
			{
				expected = GetCachedImagePath(workspaceRoot, sourceId, categoryName, familyName);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ResolveExistingCachedImagePath = string.Empty;
				ProjectData.ClearProjectError();
				goto IL_0105;
			}
			if (File.Exists(expected))
			{
				ResolveExistingCachedImagePath = expected;
			}
			else
			{
				try
				{
					string sourceFolder = Path.Combine(FamilyBrowserStandardPolicyStore.GetThumbnailFolder(workspaceRoot), SafeFileName(sourceId));
					if (!Directory.Exists(sourceFolder))
					{
						ResolveExistingCachedImagePath = expected;
					}
					else
					{
						string exactFileName = BuildCacheFileName(categoryName, familyName);
						foreach (string item in Directory.EnumerateDirectories(sourceFolder, "preview-*", SearchOption.TopDirectoryOnly))
						{
							string candidate = Path.Combine(item, exactFileName);
							if (!File.Exists(candidate))
							{
								continue;
							}
							ResolveExistingCachedImagePath = candidate;
							goto end_IL_004d;
						}
						string familyOnlyMatch = FindCachedImageByMetadata(sourceFolder, sourceId, categoryName, familyName, requireCategoryMatch: true);
						if (!string.IsNullOrWhiteSpace(familyOnlyMatch))
						{
							ResolveExistingCachedImagePath = familyOnlyMatch;
						}
						else
						{
							familyOnlyMatch = FindCachedImageByMetadata(sourceFolder, sourceId, categoryName, familyName, requireCategoryMatch: false);
							if (string.IsNullOrWhiteSpace(familyOnlyMatch))
							{
								goto IL_0103;
							}
							ResolveExistingCachedImagePath = familyOnlyMatch;
						}
					}
					end_IL_004d:;
				}
				catch (Exception projectError2)
				{
					ProjectData.SetProjectError(projectError2);
					ProjectData.ClearProjectError();
					goto IL_0103;
				}
			}
		}
		goto IL_0105;
		IL_0105:
		return ResolveExistingCachedImagePath;
		IL_0103:
		ResolveExistingCachedImagePath = expected;
		goto IL_0105;
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
		//IL_007b: Unknown result type (might be due to invalid IL or missing references)
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
		List<Family> families = (from Family x in (IEnumerable)new FilteredElementCollector(standardDocument).OfClass(typeof(Family))
			where x != null && x.FamilyCategory != null
			where CS_0024_003C_003E8__locals3._0024VB_0024Local_selectedNameSet.Count == 0 || CS_0024_003C_003E8__locals3._0024VB_0024Local_selectedNameSet.Contains(Normalize(((Element)x).Name ?? string.Empty))
			select x).OrderBy<Family, string>([SpecialName] (Family x) => Normalize(x.FamilyCategory.Name) + "|" + Normalize(((Element)x).Name), StringComparer.Ordinal).ToList();
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
					string messageName = categoryName + " / " + ((Element)family).Name;
					progress?.Invoke(current, total, messageName);
					FamilyThumbnailBatchUpdateItem item = new FamilyThumbnailBatchUpdateItem
					{
						FamilyName = ((Element)family).Name,
						CategoryName = categoryName,
						ImagePath = GetCachedImagePath(workspaceRoot, registration.SourceId, categoryName, ((Element)family).Name)
					};
					StandardLoadableFamilySnapshotItem snapshotItem = null;
					snapshotFamilyMap?.TryGetValue(BuildSnapshotFamilyKey(categoryName, ((Element)family).Name), out snapshotItem);
					string cacheStamp = BuildFamilyThumbnailCacheStamp(registration, standardSnapshot, snapshotItem, categoryName, ((Element)family).Name);
					int dialogRecordStart = dialogGuard.RecordCount;
					dialogGuard.SetCurrentFamily(categoryName, ((Element)family).Name);
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
							WriteThumbnailCacheMetadata(item.ImagePath, registration, standardSnapshot, categoryName, ((Element)family).Name, cacheStamp);
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
			result.DiagnosticReportPath = WriteBatchDiagnosticReport(result);
			return result;
		}
	}

	public static FamilyThumbnailBatchUpdateResult UpdateProjectCache(Document projectDocument, string workspaceRoot, string projectThumbnailSourceId, ProjectContentSnapshot projectSnapshot, Action<int, int, string> progress, UIApplication uiApplication = null, ISet<string> selectedFamilyNames = null)
	{
		//IL_0073: Unknown result type (might be due to invalid IL or missing references)
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
		List<Family> families = (from Family x in (IEnumerable)new FilteredElementCollector(projectDocument).OfClass(typeof(Family))
			where x != null && x.FamilyCategory != null
			where FamilyBrowserFamilyClassificationService.IsBrowserLoadableFamily(x)
			where CS_0024_003C_003E8__locals3._0024VB_0024Local_selectedNameSet.Count == 0 || CS_0024_003C_003E8__locals3._0024VB_0024Local_selectedNameSet.Contains(Normalize(((Element)x).Name ?? string.Empty))
			select x).OrderBy<Family, string>([SpecialName] (Family x) => Normalize(x.FamilyCategory.Name) + "|" + Normalize(((Element)x).Name), StringComparer.Ordinal).ToList();
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
					string messageName = categoryName + " / " + ((Element)family).Name;
					progress?.Invoke(current, total, messageName);
					FamilyThumbnailBatchUpdateItem item = new FamilyThumbnailBatchUpdateItem
					{
						FamilyName = ((Element)family).Name,
						CategoryName = categoryName,
						ImagePath = GetCachedImagePath(workspaceRoot, projectThumbnailSourceId, categoryName, ((Element)family).Name)
					};
					ProjectLoadableFamilySnapshotItem snapshotItem = null;
					snapshotFamilyMap?.TryGetValue(BuildSnapshotFamilyKey(categoryName, ((Element)family).Name), out snapshotItem);
					string cacheStamp = BuildProjectFamilyThumbnailCacheStamp(projectThumbnailSourceId, projectSnapshot, snapshotItem, categoryName, ((Element)family).Name);
					int dialogRecordStart = dialogGuard.RecordCount;
					dialogGuard.SetCurrentFamily(categoryName, ((Element)family).Name);
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
							WriteThumbnailCacheMetadata(item.ImagePath, projectThumbnailSourceId, string.Empty, 0L, "Project", projectSnapshot?.CapturedAtUtc ?? string.Empty, categoryName, ((Element)family).Name, cacheStamp);
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
			result.DiagnosticReportPath = WriteBatchDiagnosticReport(result);
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
		list.Add("thumbnail-cache-stamp-v2-shaded-thinline");
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
		list.Add("project-thumbnail-cache-stamp-v1-shaded-thinline");
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
				File.WriteAllText(path, PlainJsonReportWriter.Serialize(metadata), Encoding.UTF8);
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
			//IL_000b: Unknown result type (might be due to invalid IL or missing references)
			//IL_0011: Expected O, but got Unknown
			//IL_0012: Unknown result type (might be due to invalid IL or missing references)
			//IL_0073: Unknown result type (might be due to invalid IL or missing references)
			//IL_0079: Invalid comparison between Unknown and I4
			//IL_0065: Unknown result type (might be due to invalid IL or missing references)
			//IL_007c: Unknown result type (might be due to invalid IL or missing references)
			Transaction val = new Transaction(familyDocument, "KKY Family Browser Preview View");
			try
			{
				val.Start();
				try
				{
					FailureHandlingOptions failureHandlingOptions = val.GetFailureHandlingOptions();
					failureHandlingOptions.SetFailuresPreprocessor((IFailuresPreprocessor)(object)new FamilyThumbnailPreviewFailuresPreprocessor());
					failureHandlingOptions.SetClearAfterRollback(true);
					val.SetFailureHandlingOptions(failureHandlingOptions);
					View3D val2 = CreatePreviewView(familyDocument);
					previewBoundsResult = PreparePreviewView(familyDocument, val2);
					viewId = ((Element)val2).Id;
					val.Commit();
				}
				catch (Exception projectError2)
				{
					ProjectData.SetProjectError(projectError2);
					try
					{
						if ((int)val.GetStatus() == 1)
						{
							val.RollBack();
						}
					}
					catch (Exception projectError3)
					{
						ProjectData.SetProjectError(projectError3);
						ProjectData.ClearProjectError();
					}
					throw;
				}
			}
			finally
			{
				((IDisposable)val)?.Dispose();
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
		RunPreviewStage(generation, "Save centered thumbnail cache file", [SpecialName] () =>
		{
			if (!SaveCenteredPreviewImage(text2, targetImagePath))
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
		generation.Message = "Snapshot OK: new 3D view, shaded/fine, preview visibility on, annotations/reference datum hidden, connector graphics shown with thin/transparent overrides, element/category thin lines, safe padded section box, low-DPI export for thinner connector strokes, focused home view, robust PNG export, centered thumbnail content.";
		if (generation.ConnectorExtentsClamped)
		{
			generation.Message += " Connector extents were clamped to keep the family geometry readable.";
		}
		return generation;
	}

	private static FamilyThumbnailGenerationResult GenerateAccurate3DPreview(Document projectDocument, Family family, string targetImagePath)
	{
		Document val = null;
		try
		{
			RunPreviewStage(new FamilyThumbnailGenerationResult(), "Open family in background", [SpecialName] () =>
			{
				val = projectDocument.EditFamily(family);
			});
			return GenerateFromOpenFamilyDocument(val, targetImagePath);
		}
		finally
		{
			if (val != null)
			{
				try
				{
					val.Close(false);
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
		return TryGetPreviewImage((Element)(object)family);
	}

	private static Bitmap TryGetPreviewImage(Element element)
	{
		Bitmap TryGetPreviewImage;
		if (element == null)
		{
			TryGetPreviewImage = null;
		}
		else
		{
			ElementType elementType = (ElementType)(object)((element is ElementType) ? element : null);
			if (elementType == null)
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
		}
		return TryGetPreviewImage;
	}

	private static Family FindFamily(Document projectDocument, string familyName, string categoryName)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0031: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		string normalizedFamilyName = Normalize(familyName);
		string normalizedCategoryName = Normalize(categoryName);
		foreach (Family item in new FilteredElementCollector(projectDocument).OfClass(typeof(Family)))
		{
			Family family = item;
			if (string.Equals(Normalize(((Element)family).Name), normalizedFamilyName, StringComparison.Ordinal))
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
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		ViewFamilyType viewType = ((IEnumerable)new FilteredElementCollector(familyDocument).OfClass(typeof(ViewFamilyType))).Cast<ViewFamilyType>().FirstOrDefault([SpecialName] (ViewFamilyType x) => (int)x.ViewFamily == 102);
		if (viewType == null)
		{
			throw new InvalidOperationException(FamilyBrowserLanguageService.Text("No 3D view family type was found in the family document.", "패밀리 문서에서 3D 뷰 패밀리 타입을 찾지 못했습니다."));
		}
		View3D view = View3D.CreateIsometric(familyDocument, ((Element)viewType).Id);
		try
		{
			((Element)view).Name = "KKY Family Browser Preview 3D " + DateTime.UtcNow.ToString("HHmmssfff", CultureInfo.InvariantCulture);
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
			((View)view).DetailLevel = (ViewDetailLevel)3;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		SetPreviewDisplayStyle(view, "Shading", "ShadingWithEdges", "Realistic", "HLR", "HiddenLine", "Wireframe");
		try
		{
			((View)view).Scale = 1;
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		try
		{
			((View)view).AreAnnotationCategoriesHidden = true;
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
		//IL_0020: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Expected O, but got Unknown
		HideReferenceDatumCategories(familyDocument, view);
		foreach (Category category2 in familyDocument.Settings.Categories)
		{
			Category category = category2;
			ConfigurePreviewCategory(view, category);
		}
		ApplyThinLineOverridesToModelElements(familyDocument, view);
	}

	private static void EnablePreviewVisibilityParameters(Document familyDocument)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0043: Unknown result type (might be due to invalid IL or missing references)
		//IL_0049: Expected O, but got Unknown
		//IL_0055: Unknown result type (might be due to invalid IL or missing references)
		//IL_005b: Invalid comparison between Unknown and I4
		try
		{
			foreach (Element element in new FilteredElementCollector(familyDocument).WhereElementIsNotElementType())
			{
				if (element == null || element.Parameters == null)
				{
					continue;
				}
				foreach (Parameter parameter2 in element.Parameters)
				{
					Parameter parameter = parameter2;
					try
					{
						if (parameter != null && !parameter.IsReadOnly && (int)parameter.StorageType == 1 && IsPreviewVisibilityParameterName((parameter.Definition == null) ? string.Empty : parameter.Definition.Name))
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
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Invalid comparison between Unknown and I4
		//IL_0054: Unknown result type (might be due to invalid IL or missing references)
		//IL_005a: Expected O, but got Unknown
		if (category == null)
		{
			return;
		}
		if ((int)category.CategoryType == 2 || IsReferenceDatumCategory(category))
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
			foreach (Category subCategory2 in category.SubCategories)
			{
				Category subCategory = subCategory2;
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
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			foreach (T element in new FilteredElementCollector(familyDocument).OfClass(typeof(T)))
			{
				if (element != null)
				{
					HideCategory(view, ((Element)element).Category);
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
			if (((View)view).CanCategoryBeHidden(category.Id))
			{
				((View)view).SetCategoryHidden(category.Id, true);
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
			if (((View)view).CanCategoryBeHidden(category.Id))
			{
				((View)view).SetCategoryHidden(category.Id, false);
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
				((View)view).SetCategoryOverrides(category.Id, CreateThinLineOverride());
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
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Expected O, but got Unknown
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
			graphicsOverride.SetHalftone(true);
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
				((View)view).SetCategoryOverrides(category.Id, CreateConnectorThinLineOverride());
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
		//IL_002b: Unknown result type (might be due to invalid IL or missing references)
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
					((View)view).DisplayStyle = (DisplayStyle)Enum.Parse(typeof(DisplayStyle), styleName, ignoreCase: true);
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
		//IL_001a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004a: Invalid comparison between Unknown and I4
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
				if (element == null || element.Category == null || (int)element.Category.CategoryType != 1 || element is View || element is ReferencePlane || element is Level)
				{
					continue;
				}
				try
				{
					if (IsConnectorElement(element))
					{
						((View)view).SetElementOverrides(element.Id, connectorGraphicsOverride);
					}
					else
					{
						((View)view).SetElementOverrides(element.Id, graphicsOverride);
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
		//IL_000a: Unknown result type (might be due to invalid IL or missing references)
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0043: Invalid comparison between Unknown and I4
		BoundingBoxXYZ physicalBounds = null;
		BoundingBoxXYZ connectorBounds = null;
		BoundingBoxXYZ fallbackModelBounds = null;
		BoundingBoxXYZ geometryBounds = null;
		foreach (Element element in new FilteredElementCollector(familyDocument).WhereElementIsNotElementType())
		{
			if (element == null || element.Category == null || (int)element.Category.CategoryType != 1)
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
			BoundingBoxXYZ bounds = element[(View)null];
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
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Expected O, but got Unknown
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
				GetUsableElementGeometryBounds = BuildGeometryBounds(element[options], Transform.Identity);
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
		if (geometry == null)
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
		if (geometryObject == null)
		{
			return null;
		}
		Transform activeTransform = transform ?? Transform.Identity;
		try
		{
			GeometryInstance instance = (GeometryInstance)(object)((geometryObject is GeometryInstance) ? geometryObject : null);
			if (instance != null)
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
			Solid solid = (Solid)(object)((geometryObject is Solid) ? geometryObject : null);
			if (solid != null && solid.Faces != null && solid.Faces.Size > 0)
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
			Curve curve = (Curve)(object)((geometryObject is Curve) ? geometryObject : null);
			if (curve != null)
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
			Mesh mesh = (Mesh)(object)((geometryObject is Mesh) ? geometryObject : null);
			if (mesh != null && mesh.NumTriangles > 0)
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
		if (curve == null)
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
		if (mesh == null)
		{
			return null;
		}
		BoundingBoxXYZ result = null;
		try
		{
			PropertyInfo verticesProperty = ((object)mesh).GetType().GetProperty("Vertices", BindingFlags.Instance | BindingFlags.Public);
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
				MethodInfo triangleMethod = ((object)mesh).GetType().GetMethod("get_Triangle", BindingFlags.Instance | BindingFlags.Public);
				if ((object)triangleMethod == null)
				{
					triangleMethod = ((object)mesh).GetType().GetMethod("Triangle", BindingFlags.Instance | BindingFlags.Public);
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
								object? obj = vertexMethod.Invoke(RuntimeHelpers.GetObjectValue(triangle), new object[1] { vertexIndex });
								XYZ point2 = (XYZ)((obj is XYZ) ? obj : null);
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
		//IL_0045: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		//IL_0079: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		//IL_0093: Unknown result type (might be due to invalid IL or missing references)
		//IL_0099: Expected O, but got Unknown
		//IL_00ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b3: Expected O, but got Unknown
		//IL_00c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cd: Expected O, but got Unknown
		//IL_00e1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e7: Expected O, but got Unknown
		//IL_00fb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0101: Expected O, but got Unknown
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
		XYZ[] array = (XYZ[])(object)obj;
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
		if (((object)element).GetType().Name.IndexOf("Connector", StringComparison.OrdinalIgnoreCase) >= 0)
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
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		bool IsBuiltInCategoryName;
		if (category == null || builtInNames == null || builtInNames.Length == 0)
		{
			IsBuiltInCategoryName = false;
		}
		else
		{
			try
			{
				string a = ((Enum)(BuiltInCategory)RevitElementIdCompat.CompatIntegerValue(category.Id)/*cast due to .constrained prefix*/).ToString();
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
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_0038: Unknown result type (might be due to invalid IL or missing references)
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0064: Expected O, but got Unknown
		//IL_0065: Expected O, but got Unknown
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
		//IL_0063: Unknown result type (might be due to invalid IL or missing references)
		//IL_0069: Expected O, but got Unknown
		//IL_007d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0083: Expected O, but got Unknown
		//IL_0097: Unknown result type (might be due to invalid IL or missing references)
		//IL_009d: Expected O, but got Unknown
		//IL_00b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b7: Expected O, but got Unknown
		//IL_00cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d1: Expected O, but got Unknown
		//IL_00e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00eb: Expected O, but got Unknown
		//IL_00ff: Unknown result type (might be due to invalid IL or missing references)
		//IL_0105: Expected O, but got Unknown
		//IL_0119: Unknown result type (might be due to invalid IL or missing references)
		//IL_011f: Expected O, but got Unknown
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
					XYZ[] array = (XYZ[])(object)obj;
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
		//IL_0052: Unknown result type (might be due to invalid IL or missing references)
		//IL_0057: Unknown result type (might be due to invalid IL or missing references)
		//IL_009a: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a4: Expected O, but got Unknown
		//IL_00a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f1: Expected O, but got Unknown
		//IL_00f2: Expected O, but got Unknown
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0025: Unknown result type (might be due to invalid IL or missing references)
		//IL_002f: Expected O, but got Unknown
		//IL_002f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0042: Unknown result type (might be due to invalid IL or missing references)
		//IL_004c: Expected O, but got Unknown
		//IL_004d: Expected O, but got Unknown
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
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		//IL_0075: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		//IL_007f: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00db: Expected O, but got Unknown
		//IL_00dc: Expected O, but got Unknown
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
		//IL_00b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ba: Expected O, but got Unknown
		//IL_00e3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e8: Unknown result type (might be due to invalid IL or missing references)
		//IL_012d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0137: Expected O, but got Unknown
		//IL_0137: Unknown result type (might be due to invalid IL or missing references)
		//IL_017c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0186: Expected O, but got Unknown
		//IL_0187: Expected O, but got Unknown
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
			((View)view).CropBoxActive = false;
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
		//IL_001b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0021: Expected O, but got Unknown
		//IL_013c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0142: Expected O, but got Unknown
		//IL_016a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0171: Expected O, but got Unknown
		//IL_0190: Unknown result type (might be due to invalid IL or missing references)
		//IL_019a: Expected O, but got Unknown
		//IL_009f: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a5: Expected O, but got Unknown
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
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_0020: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0035: Unknown result type (might be due to invalid IL or missing references)
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		//IL_004a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0051: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Expected O, but got Unknown
		List<ElementId> viewIds = new List<ElementId> { viewId };
		ImageExportOptions options = new ImageExportOptions
		{
			ExportRange = (ExportRange)2,
			FilePath = basePath,
			FitDirection = (FitDirectionType)0,
			HLRandWFViewsFileType = (ImageFileType)4,
			ImageResolution = ResolveImageResolution("DPI_72", (ImageResolution)0),
			PixelSize = 768,
			ShadowViewsFileType = (ImageFileType)4,
			ZoomType = (ZoomFitType)0
		};
		options.SetViewsAndSheets((IList<ElementId>)viewIds);
		familyDocument.ExportImage(options);
	}

	private static ImageResolution ResolveImageResolution(string preferredName, ImageResolution fallback)
	{
		//IL_002d: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_002f: Unknown result type (might be due to invalid IL or missing references)
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

	private static bool SaveCenteredPreviewImage(string sourcePath, string targetPath)
	{
		checked
		{
			bool SaveCenteredPreviewImage;
			if (string.IsNullOrWhiteSpace(sourcePath) || string.IsNullOrWhiteSpace(targetPath))
			{
				SaveCenteredPreviewImage = false;
			}
			else
			{
				try
				{
					using (Bitmap source = new Bitmap(sourcePath))
					{
						Rectangle contentBounds = FindPreviewContentBounds(source);
						int canvasWidth = source.Width;
						int canvasHeight = source.Height;
						if (canvasWidth > 0 && canvasHeight > 0)
						{
							double scale = Math.Min(1.0, Math.Max(0.75, 0.94));
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
							using (Bitmap output = new Bitmap(canvasWidth, canvasHeight, PixelFormat.Format32bppArgb))
							{
								using (Graphics graphics = Graphics.FromImage(output))
								{
									graphics.Clear(EstimatePreviewBackgroundColor(source));
									graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
									graphics.SmoothingMode = SmoothingMode.HighQuality;
									graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
									graphics.DrawImage(source, new Rectangle(drawX, drawY, drawWidth, drawHeight), new Rectangle(0, 0, canvasWidth, canvasHeight), GraphicsUnit.Pixel);
								}
								output.Save(targetPath, ImageFormat.Png);
							}
							goto end_IL_001f;
						}
						SaveCenteredPreviewImage = false;
						goto end_IL_0018;
						end_IL_001f:;
					}
					SaveCenteredPreviewImage = true;
					end_IL_0018:;
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					SaveCenteredPreviewImage = false;
					ProjectData.ClearProjectError();
				}
			}
			return SaveCenteredPreviewImage;
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

	private static Rectangle FindPreviewContentBounds(Bitmap image)
	{
		if (image == null || image.Width <= 0 || image.Height <= 0)
		{
			return Rectangle.Empty;
		}
		Color background = EstimatePreviewBackgroundColor(image);
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
					Rectangle rect = new Rectangle(0, 0, scanImage.Width, scanImage.Height);
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
							if (IsPreviewContentPixel(Color.FromArgb(bytes[offset + 3], bytes[offset + 2], bytes[offset + 1], bytes[offset]), background))
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
				return Rectangle.Empty;
			}
			int inflate = Math.Max(24, (int)Math.Round((double)Math.Min(image.Width, image.Height) * 0.1));
			left = Math.Max(0, left - inflate);
			top = Math.Max(0, top - inflate);
			right = Math.Min(image.Width - 1, right + inflate);
			bottom = Math.Min(image.Height - 1, bottom + inflate);
			return Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
		}
	}

	private static bool PreviewContentTouchesImageEdge(Rectangle contentBounds, int imageWidth, int imageHeight)
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

	private static Color EstimatePreviewBackgroundColor(Bitmap image)
	{
		if (image == null || image.Width <= 0 || image.Height <= 0)
		{
			return Color.White;
		}
		checked
		{
			Color[] source = new Color[4]
			{
				image.GetPixel(0, 0),
				image.GetPixel(image.Width - 1, 0),
				image.GetPixel(0, image.Height - 1),
				image.GetPixel(image.Width - 1, image.Height - 1)
			};
			int a = (int)Math.Round(source.Average([SpecialName] (Color c) => c.A));
			int r = (int)Math.Round(source.Average([SpecialName] (Color c) => c.R));
			int g = (int)Math.Round(source.Average([SpecialName] (Color c) => c.G));
			int b = (int)Math.Round(source.Average([SpecialName] (Color c) => c.B));
			return Color.FromArgb(a, r, g, b);
		}
	}

	private static bool IsPreviewContentPixel(Color pixel, Color background)
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

	private static string WriteBatchDiagnosticReport(FamilyThumbnailBatchUpdateResult result)
	{
		string WriteBatchDiagnosticReport;
		try
		{
			if (result == null || string.IsNullOrWhiteSpace(result.OutputFolder))
			{
				WriteBatchDiagnosticReport = string.Empty;
			}
			else
			{
				Directory.CreateDirectory(result.OutputFolder);
				string text = Path.Combine(result.OutputFolder, "thumbnail-diagnostics-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + ".txt");
				File.WriteAllText(text, BuildBatchDiagnosticText(result), Encoding.UTF8);
				WriteBatchDiagnosticReport = text;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			WriteBatchDiagnosticReport = string.Empty;
			ProjectData.ClearProjectError();
		}
		return WriteBatchDiagnosticReport;
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
		string safe = new string(value.Select([SpecialName] (char ch) => (!Enumerable.Contains(invalidFileNameChars, ch)) ? ch : '_').ToArray()).Trim().TrimEnd('.', ' ');
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
		string prefix = safe.Substring(0, Math.Min(prefixLength, safe.Length)).Trim().TrimEnd('.', ' ');
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
