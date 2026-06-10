using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Microsoft.VisualBasic.CompilerServices;

public sealed class ProjectSnapshotStore
{
	private sealed class FileStamp
	{
		public string LastWriteUtc { get; set; }

		public long Length { get; set; }

		public FileStamp()
		{
			LastWriteUtc = string.Empty;
		}
	}

	private ProjectSnapshotStore()
	{
	}

	public static string Save(string workspaceRoot, ProjectContentSnapshot snapshot, Document doc = null)
	{
		FamilyBrowserRevitVersionContext.SetCurrentVersion(doc);
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Save project scan snapshot", "프로젝트 스캔 스냅샷 저장"));
		string outputDir = Path.Combine(ResolveProjectHistoryFolder(workspaceRoot, snapshot, doc), "Snapshots");
		Directory.CreateDirectory(outputDir);
		string safeDocumentName = MakeSafeFileName(snapshot.DocumentTitle ?? "Untitled");
		string fileName = "project-snapshot-" + safeDocumentName + "-" + DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + ".json";
		string text = Path.Combine(outputDir, fileName);
		File.WriteAllText(text, PlainJsonReportWriter.Serialize(snapshot));
		return text;
	}

	public static ProjectScanCacheRecord SaveLatestProjectScan(string workspaceRoot, Document doc, StandardLibraryRegistrationRecord registration, StandardLibrarySnapshot standardSnapshot, string projectSnapshotPath, string comparisonReportPath, ProjectContentSnapshot projectSnapshot, ProjectStandardComparisonReport report, string thumbnailSourceId)
	{
		FamilyBrowserRevitVersionContext.SetCurrentVersion(doc);
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Save latest project scan cache", "프로젝트 최신 스캔 캐시 저장"));
		string projectKey = ResolveProjectCacheKey(workspaceRoot, doc);
		if (string.IsNullOrWhiteSpace(projectKey))
		{
			return null;
		}
		string identityPath = ResolveProjectIdentityPath(doc);
		FileStamp projectStamp = BuildFileStamp(identityPath);
		string thumbnailFolder = (string.IsNullOrWhiteSpace(thumbnailSourceId) ? string.Empty : FamilyThumbnailPreviewService.GetCacheFolder(workspaceRoot, thumbnailSourceId));
		if (report != null && report.Project != null)
		{
			report.Project.ThumbnailSourceId = thumbnailSourceId ?? string.Empty;
			report.Project.ThumbnailFolder = thumbnailFolder;
		}
		ProjectScanCacheRecord projectScanCacheRecord = new ProjectScanCacheRecord();
		projectScanCacheRecord.ProjectKey = projectKey;
		projectScanCacheRecord.ProjectTitle = SafeDocumentTitle(doc);
		projectScanCacheRecord.ProjectDocumentPath = SafeDocumentPath(doc);
		projectScanCacheRecord.ProjectCentralPath = ResolveCentralPath(doc);
		projectScanCacheRecord.ProjectIdentityPath = identityPath;
		projectScanCacheRecord.ProjectFileLastWriteUtc = projectStamp.LastWriteUtc;
		projectScanCacheRecord.ProjectFileLength = projectStamp.Length;
		projectScanCacheRecord.RevitVersion = SafeRevitVersion(doc);
		projectScanCacheRecord.StandardSourceId = registration?.SourceId ?? string.Empty;
		projectScanCacheRecord.StandardSnapshotPath = registration?.LastSnapshotPath ?? string.Empty;
		projectScanCacheRecord.StandardSnapshotAtUtc = standardSnapshot?.CapturedAtUtc ?? registration?.LastSnapshotAtUtc;
		projectScanCacheRecord.StandardSnapshotMode = standardSnapshot?.SnapshotMode ?? registration?.SnapshotMode;
		projectScanCacheRecord.StandardSourceFileLastWriteUtc = registration?.SourceFileLastWriteUtc ?? string.Empty;
		projectScanCacheRecord.StandardSourceFileLength = registration?.SourceFileLength ?? 0;
		projectScanCacheRecord.ProjectSnapshotPath = projectSnapshotPath ?? string.Empty;
		projectScanCacheRecord.ComparisonReportPath = comparisonReportPath ?? string.Empty;
		projectScanCacheRecord.ThumbnailSourceId = thumbnailSourceId ?? string.Empty;
		projectScanCacheRecord.ThumbnailFolder = thumbnailFolder;
		projectScanCacheRecord.CapturedAtUtc = projectSnapshot?.CapturedAtUtc ?? DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		projectScanCacheRecord.SavedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		ProjectScanCacheRecord record = projectScanCacheRecord;
		string path = BuildProjectScanRecordPath(workspaceRoot, projectKey);
		Directory.CreateDirectory(Path.GetDirectoryName(path));
		File.WriteAllText(path, PlainJsonReportWriter.Serialize(record), Encoding.UTF8);
		SaveProjectScanAliasRecords(workspaceRoot, projectKey, record);
		return record;
	}

	public static ProjectScanCacheLoadResult TryLoadLatestProjectScan(string workspaceRoot, Document doc, StandardLibraryRegistrationRecord registration, StandardLibrarySnapshot standardSnapshot)
	{
		FamilyBrowserRevitVersionContext.SetCurrentVersion(doc);
		ProjectScanCacheLoadResult result = new ProjectScanCacheLoadResult();
		string projectKey = ResolveProjectCacheKey(workspaceRoot, doc);
		ProjectScanCacheLoadResult TryLoadLatestProjectScan;
		ProjectScanCacheRecord record;
		if (string.IsNullOrWhiteSpace(projectKey))
		{
			result.Reason = "No project key.";
			TryLoadLatestProjectScan = result;
		}
		else
		{
			string recordPath = BuildProjectScanRecordPath(workspaceRoot, projectKey);
			if (!File.Exists(recordPath))
			{
				result.Reason = "No cached project scan record.";
				TryLoadLatestProjectScan = result;
			}
			else
			{
				record = null;
				try
				{
					record = DataContractJsonFileStore.Load<ProjectScanCacheRecord>(recordPath);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					result.Reason = "Cached project scan record could not be read: " + ex2.Message;
					TryLoadLatestProjectScan = result;
					ProjectData.ClearProjectError();
					goto IL_03fd;
				}
				if (record == null)
				{
					result.Reason = "Cached project scan record was empty.";
					TryLoadLatestProjectScan = result;
				}
				else if (record.SchemaVersion < 2)
				{
					result.Reason = "Cached project scan schema is stale.";
					TryLoadLatestProjectScan = result;
				}
				else if (!string.Equals(record.ProjectKey ?? string.Empty, projectKey, StringComparison.Ordinal))
				{
					result.Reason = "Cached project key does not match.";
					TryLoadLatestProjectScan = result;
				}
				else
				{
					if (registration == null)
					{
						goto IL_019d;
					}
					if (!string.Equals(record.StandardSourceId ?? string.Empty, registration.SourceId ?? string.Empty, StringComparison.OrdinalIgnoreCase))
					{
						result.Reason = "Cached standard source id does not match.";
						TryLoadLatestProjectScan = result;
					}
					else if (!string.Equals(NormalizePath(record.StandardSnapshotPath), NormalizePath(registration.LastSnapshotPath), StringComparison.OrdinalIgnoreCase))
					{
						result.Reason = "Cached standard snapshot path does not match.";
						TryLoadLatestProjectScan = result;
					}
					else
					{
						if (string.Equals(record.StandardSourceFileLastWriteUtc ?? string.Empty, registration.SourceFileLastWriteUtc ?? string.Empty, StringComparison.Ordinal) && record.StandardSourceFileLength == registration.SourceFileLength)
						{
							goto IL_019d;
						}
						result.Reason = "Cached standard RVT stamp is stale.";
						TryLoadLatestProjectScan = result;
					}
				}
			}
		}
		goto IL_03fd;
		IL_019d:
		if (standardSnapshot != null && !string.IsNullOrWhiteSpace(standardSnapshot.CapturedAtUtc) && !string.Equals(record.StandardSnapshotAtUtc ?? string.Empty, standardSnapshot.CapturedAtUtc, StringComparison.Ordinal))
		{
			result.Reason = "Cached standard snapshot timestamp is stale.";
			TryLoadLatestProjectScan = result;
		}
		else
		{
			string currentStampPath = ResolveProjectIdentityPath(doc);
			FileStamp currentStamp = BuildFileStamp(currentStampPath);
			int num;
			if (IsProjectAliasMatch(record, doc))
			{
				num = ((!string.Equals(NormalizePath(record.ProjectIdentityPath), NormalizePath(currentStampPath), StringComparison.OrdinalIgnoreCase)) ? 1 : 0);
				if (num != 0 && !string.IsNullOrWhiteSpace(record.ProjectIdentityPath))
				{
					FileStamp canonicalStamp = BuildFileStamp(record.ProjectIdentityPath);
					if (!string.IsNullOrWhiteSpace(canonicalStamp.LastWriteUtc) || canonicalStamp.Length > 0)
					{
						currentStamp = canonicalStamp;
					}
				}
			}
			else
			{
				num = 0;
			}
			if (num != 0 && string.IsNullOrWhiteSpace(currentStamp.LastWriteUtc))
			{
				goto IL_02d1;
			}
			if (!string.IsNullOrWhiteSpace(record.ProjectFileLastWriteUtc) && !string.Equals(record.ProjectFileLastWriteUtc, currentStamp.LastWriteUtc, StringComparison.Ordinal))
			{
				result.Reason = "Cached project scan is stale: file timestamp changed.";
				TryLoadLatestProjectScan = result;
			}
			else
			{
				if (record.ProjectFileLength <= 0 || currentStamp.Length <= 0 || record.ProjectFileLength == currentStamp.Length)
				{
					goto IL_02d1;
				}
				result.Reason = "Cached project scan is stale: file length changed.";
				TryLoadLatestProjectScan = result;
			}
		}
		goto IL_03fd;
		IL_03fd:
		return TryLoadLatestProjectScan;
		IL_02d1:
		if (string.IsNullOrWhiteSpace(record.ProjectSnapshotPath) || !File.Exists(record.ProjectSnapshotPath))
		{
			result.Reason = "Cached project snapshot file is missing.";
			TryLoadLatestProjectScan = result;
		}
		else if (string.IsNullOrWhiteSpace(record.ComparisonReportPath) || !File.Exists(record.ComparisonReportPath))
		{
			result.Reason = "Cached comparison report file is missing.";
			TryLoadLatestProjectScan = result;
		}
		else
		{
			try
			{
				result.Snapshot = DataContractJsonFileStore.Load<ProjectContentSnapshot>(record.ProjectSnapshotPath);
				result.Report = DataContractJsonFileStore.Load<ProjectStandardComparisonReport>(record.ComparisonReportPath);
				if (result.Report != null && result.Report.Project != null)
				{
					result.Report.Project.ThumbnailSourceId = record.ThumbnailSourceId ?? string.Empty;
					result.Report.Project.ThumbnailFolder = record.ThumbnailFolder ?? string.Empty;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				result.Reason = "Cached project scan data could not be read: " + ex4.Message;
				TryLoadLatestProjectScan = result;
				ProjectData.ClearProjectError();
				goto IL_03fd;
			}
			result.Record = record;
			result.Success = result.Report != null;
			if (!result.Success)
			{
				result.Reason = "Cached comparison report was empty.";
			}
			TryLoadLatestProjectScan = result;
		}
		goto IL_03fd;
	}

	public static string BuildProjectThumbnailSourceId(Document doc)
	{
		string key = BuildProjectCacheKey(doc);
		if (string.IsNullOrWhiteSpace(key))
		{
			return string.Empty;
		}
		return "project-" + HashText(key).Substring(0, 24);
	}

	public static string BuildProjectThumbnailSourceId(string workspaceRoot, Document doc)
	{
		string key = ResolveProjectCacheKey(workspaceRoot, doc);
		if (string.IsNullOrWhiteSpace(key))
		{
			return string.Empty;
		}
		return "project-" + HashText(key).Substring(0, 24);
	}

	public static string GetProjectHistoryFolder(string workspaceRoot, Document doc)
	{
		FamilyBrowserRevitVersionContext.SetCurrentVersion(doc);
		string projectKey = ResolveProjectCacheKey(workspaceRoot, doc);
		return GetProjectHistoryFolderFromKey(workspaceRoot, projectKey);
	}

	public static string GetProjectHistoryFolder(string workspaceRoot, string documentPath, string documentTitle)
	{
		string projectKey = BuildProjectCacheKey(documentPath, documentTitle);
		return GetProjectHistoryFolderFromKey(workspaceRoot, projectKey);
	}

	private static string MakeSafeFileName(string value)
	{
		char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
		string normalized = new string(value.Select([SpecialName] (char ch) => (!Enumerable.Contains(invalidFileNameChars, ch)) ? ch : '_').ToArray()).Trim();
		if (normalized.Length == 0)
		{
			return "Untitled";
		}
		return normalized;
	}

	private static string BuildProjectScanRecordPath(string workspaceRoot, string projectKey)
	{
		return Path.Combine(GetProjectHistoryFolderFromKey(workspaceRoot, projectKey), "project-scan-latest.json");
	}

	public static string GetProjectScanCacheFolder(string workspaceRoot)
	{
		return FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "Projects");
	}

	public static string GetLatestProjectScanRecordCacheStamp(string workspaceRoot, Document doc)
	{
		FamilyBrowserRevitVersionContext.SetCurrentVersion(doc);
		string projectKey = ResolveProjectCacheKey(workspaceRoot, doc);
		string GetLatestProjectScanRecordCacheStamp;
		if (string.IsNullOrWhiteSpace(projectKey))
		{
			GetLatestProjectScanRecordCacheStamp = string.Empty;
		}
		else
		{
			string recordPath = BuildProjectScanRecordPath(workspaceRoot, projectKey);
			if (string.IsNullOrWhiteSpace(recordPath))
			{
				GetLatestProjectScanRecordCacheStamp = string.Empty;
			}
			else
			{
				try
				{
					FileInfo info = new FileInfo(recordPath);
					GetLatestProjectScanRecordCacheStamp = (info.Exists ? (recordPath + "|" + info.Length.ToString(CultureInfo.InvariantCulture) + "|" + info.LastWriteTimeUtc.Ticks.ToString(CultureInfo.InvariantCulture)) : (recordPath + "|missing"));
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					GetLatestProjectScanRecordCacheStamp = recordPath;
					ProjectData.ClearProjectError();
				}
			}
		}
		return GetLatestProjectScanRecordCacheStamp;
	}

	private static string BuildProjectCacheKey(Document doc)
	{
		string identityPath = ResolveProjectIdentityPath(doc);
		if (!string.IsNullOrWhiteSpace(identityPath))
		{
			return "path|" + NormalizePath(identityPath);
		}
		string title = SafeDocumentTitle(doc);
		if (string.IsNullOrWhiteSpace(title))
		{
			return string.Empty;
		}
		return "title|" + title.Trim().ToLowerInvariant();
	}

	private static string ResolveProjectCacheKey(string workspaceRoot, Document doc)
	{
		string primaryKey = BuildProjectCacheKey(doc);
		if (string.IsNullOrWhiteSpace(workspaceRoot))
		{
			return primaryKey;
		}
		List<string> candidateKeys = BuildProjectCacheKeyCandidates(doc);
		foreach (string candidateKey in candidateKeys)
		{
			ProjectScanCacheRecord record = TryReadProjectScanRecord(workspaceRoot, candidateKey);
			if (record != null && !string.IsNullOrWhiteSpace(record.ProjectKey))
			{
				return record.ProjectKey;
			}
		}
		return primaryKey;
	}

	private static List<string> BuildProjectCacheKeyCandidates(Document doc)
	{
		List<string> list = new List<string>();
		AddProjectKey(list, BuildProjectCacheKey(doc));
		string centralPath = ResolveCentralPath(doc);
		string documentPath = SafeDocumentPath(doc);
		string title = SafeDocumentTitle(doc);
		AddProjectNameAliasKeys(list, centralPath);
		AddProjectNameAliasKeys(list, documentPath);
		AddProjectNameAliasKeys(list, title);
		return list;
	}

	private static List<string> BuildProjectAliasKeys(ProjectScanCacheRecord record)
	{
		List<string> result = new List<string>();
		if (record == null)
		{
			return result;
		}
		AddProjectNameAliasKeys(result, record.ProjectCentralPath);
		AddProjectNameAliasKeys(result, record.ProjectIdentityPath);
		AddProjectNameAliasKeys(result, record.ProjectDocumentPath);
		AddProjectNameAliasKeys(result, record.ProjectTitle);
		return result;
	}

	private static void AddProjectNameAliasKeys(List<string> keys, string value)
	{
		string token = CanonicalProjectFileToken(value);
		if (!string.IsNullOrWhiteSpace(token))
		{
			AddProjectKey(keys, "name|" + token);
		}
	}

	private static void AddProjectKey(List<string> keys, string key)
	{
		if (keys != null && !string.IsNullOrWhiteSpace(key) && !keys.Any([SpecialName] (string x) => string.Equals(x, key, StringComparison.OrdinalIgnoreCase)))
		{
			keys.Add(key);
		}
	}

	private static void SaveProjectScanAliasRecords(string workspaceRoot, string canonicalProjectKey, ProjectScanCacheRecord record)
	{
		if (string.IsNullOrWhiteSpace(workspaceRoot) || string.IsNullOrWhiteSpace(canonicalProjectKey) || record == null)
		{
			return;
		}
		foreach (string aliasKey in BuildProjectAliasKeys(record))
		{
			if (!string.IsNullOrWhiteSpace(aliasKey) && !string.Equals(aliasKey, canonicalProjectKey, StringComparison.OrdinalIgnoreCase))
			{
				try
				{
					ProjectScanCacheRecord aliasRecord = CloneProjectScanRecord(record);
					aliasRecord.ProjectKey = canonicalProjectKey;
					string path = BuildProjectScanRecordPath(workspaceRoot, aliasKey);
					Directory.CreateDirectory(Path.GetDirectoryName(path));
					File.WriteAllText(path, PlainJsonReportWriter.Serialize(aliasRecord), Encoding.UTF8);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
	}

	private static ProjectScanCacheRecord CloneProjectScanRecord(ProjectScanCacheRecord record)
	{
		if (record == null)
		{
			return null;
		}
		return new ProjectScanCacheRecord
		{
			SchemaVersion = record.SchemaVersion,
			ProjectKey = record.ProjectKey,
			ProjectTitle = record.ProjectTitle,
			ProjectDocumentPath = record.ProjectDocumentPath,
			ProjectCentralPath = record.ProjectCentralPath,
			ProjectIdentityPath = record.ProjectIdentityPath,
			ProjectFileLastWriteUtc = record.ProjectFileLastWriteUtc,
			ProjectFileLength = record.ProjectFileLength,
			RevitVersion = record.RevitVersion,
			StandardSourceId = record.StandardSourceId,
			StandardSnapshotPath = record.StandardSnapshotPath,
			StandardSnapshotAtUtc = record.StandardSnapshotAtUtc,
			StandardSnapshotMode = record.StandardSnapshotMode,
			StandardSourceFileLastWriteUtc = record.StandardSourceFileLastWriteUtc,
			StandardSourceFileLength = record.StandardSourceFileLength,
			ProjectSnapshotPath = record.ProjectSnapshotPath,
			ComparisonReportPath = record.ComparisonReportPath,
			ThumbnailSourceId = record.ThumbnailSourceId,
			ThumbnailFolder = record.ThumbnailFolder,
			CapturedAtUtc = record.CapturedAtUtc,
			SavedAtUtc = record.SavedAtUtc
		};
	}

	private static ProjectScanCacheRecord TryReadProjectScanRecord(string workspaceRoot, string projectKey)
	{
		ProjectScanCacheRecord TryReadProjectScanRecord;
		if (string.IsNullOrWhiteSpace(workspaceRoot) || string.IsNullOrWhiteSpace(projectKey))
		{
			TryReadProjectScanRecord = null;
		}
		else
		{
			try
			{
				string recordPath = BuildProjectScanRecordPath(workspaceRoot, projectKey);
				TryReadProjectScanRecord = ((!string.IsNullOrWhiteSpace(recordPath) && File.Exists(recordPath)) ? DataContractJsonFileStore.Load<ProjectScanCacheRecord>(recordPath) : null);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				TryReadProjectScanRecord = null;
				ProjectData.ClearProjectError();
			}
		}
		return TryReadProjectScanRecord;
	}

	private static bool IsProjectAliasMatch(ProjectScanCacheRecord record, Document doc)
	{
		if (record == null || doc == null)
		{
			return false;
		}
		List<string> list = (from x in BuildProjectCacheKeyCandidates(doc)
			where x.StartsWith("name|", StringComparison.OrdinalIgnoreCase)
			select x.Substring("name|".Length) into x
			where !string.IsNullOrWhiteSpace(x)
			select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).ToList();
		if (list.Count == 0)
		{
			return false;
		}
		return (from x in BuildProjectAliasKeys(record)
			where x.StartsWith("name|", StringComparison.OrdinalIgnoreCase)
			select x.Substring("name|".Length) into x
			where !string.IsNullOrWhiteSpace(x)
			select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).ToList().Any([SpecialName] (string x) => list.Contains<string>(x, StringComparer.OrdinalIgnoreCase));
	}

	private static string CanonicalProjectFileToken(string value)
	{
		if (string.IsNullOrWhiteSpace(value))
		{
			return string.Empty;
		}
		string text = value.Trim();
		try
		{
			text = Path.GetFileNameWithoutExtension(text);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		if (string.IsNullOrWhiteSpace(text))
		{
			return string.Empty;
		}
		string token = text.Trim();
		string[] suffixes = new string[12]
		{
			"_detached", "-detached", ".detached", " detached", "(detached)", " - detached", " _ detached", "_detached copy", "_detached_copy", "-detached copy",
			"-detached-copy", " detached copy"
		};
		bool changed = true;
		while (changed)
		{
			changed = false;
			string[] array = suffixes;
			foreach (string suffix in array)
			{
				if (token.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
				{
					token = token.Substring(0, checked(token.Length - suffix.Length)).Trim();
					changed = true;
				}
			}
		}
		return token.ToLowerInvariant().Replace(" ", string.Empty).Replace("_", string.Empty)
			.Replace("-", string.Empty)
			.Replace(".", string.Empty)
			.Replace(",", string.Empty);
	}

	private static string BuildProjectCacheKey(string documentPath, string documentTitle)
	{
		if (!string.IsNullOrWhiteSpace(documentPath))
		{
			return "path|" + NormalizePath(documentPath);
		}
		if (string.IsNullOrWhiteSpace(documentTitle))
		{
			return string.Empty;
		}
		return "title|" + documentTitle.Trim().ToLowerInvariant();
	}

	private static string ResolveProjectHistoryFolder(string workspaceRoot, ProjectContentSnapshot snapshot, Document doc)
	{
		if (doc != null)
		{
			return GetProjectHistoryFolder(workspaceRoot, doc);
		}
		string documentPath = snapshot?.DocumentPath ?? string.Empty;
		string documentTitle = snapshot?.DocumentTitle ?? string.Empty;
		return GetProjectHistoryFolder(workspaceRoot, documentPath, documentTitle);
	}

	private static string GetProjectHistoryFolderFromKey(string workspaceRoot, string projectKey)
	{
		string effectiveKey = projectKey ?? string.Empty;
		if (string.IsNullOrWhiteSpace(effectiveKey))
		{
			effectiveKey = "untitled|" + Guid.NewGuid().ToString("N");
		}
		return Path.Combine(GetProjectScanCacheFolder(workspaceRoot), HashText(effectiveKey).Substring(0, 32));
	}

	public static string ResolveProjectIdentityPath(Document doc)
	{
		string centralPath = ResolveCentralPath(doc);
		if (!string.IsNullOrWhiteSpace(centralPath))
		{
			return centralPath;
		}
		return SafeDocumentPath(doc);
	}

	private static string ResolveCentralPath(Document doc)
	{
		string ResolveCentralPath;
		if (doc == null)
		{
			ResolveCentralPath = string.Empty;
		}
		else
		{
			try
			{
				if (!doc.IsWorkshared)
				{
					ResolveCentralPath = string.Empty;
				}
				else
				{
					ModelPath modelPath = doc.GetWorksharingCentralModelPath();
					ResolveCentralPath = ((modelPath != null) ? ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath) : string.Empty);
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ResolveCentralPath = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return ResolveCentralPath;
	}

	private static string SafeDocumentTitle(Document doc)
	{
		string SafeDocumentTitle;
		try
		{
			SafeDocumentTitle = ((doc != null) ? doc.Title : null) ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			SafeDocumentTitle = string.Empty;
			ProjectData.ClearProjectError();
		}
		return SafeDocumentTitle;
	}

	private static string SafeDocumentPath(Document doc)
	{
		string SafeDocumentPath;
		try
		{
			SafeDocumentPath = ((doc != null) ? doc.PathName : null) ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			SafeDocumentPath = string.Empty;
			ProjectData.ClearProjectError();
		}
		return SafeDocumentPath;
	}

	private static string SafeRevitVersion(Document doc)
	{
		string SafeRevitVersion;
		try
		{
			object obj;
			if (doc == null)
			{
				obj = null;
			}
			else
			{
				Application application = doc.Application;
				obj = ((application != null) ? application.VersionNumber : null);
			}
			if (obj == null)
			{
				obj = string.Empty;
			}
			SafeRevitVersion = (string)obj;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			SafeRevitVersion = string.Empty;
			ProjectData.ClearProjectError();
		}
		return SafeRevitVersion;
	}

	private static FileStamp BuildFileStamp(string path)
	{
		FileStamp stamp = new FileStamp();
		if (string.IsNullOrWhiteSpace(path))
		{
			return stamp;
		}
		try
		{
			FileInfo info = new FileInfo(path);
			if (info.Exists)
			{
				stamp.LastWriteUtc = info.LastWriteTimeUtc.ToString("O", CultureInfo.InvariantCulture);
				stamp.Length = info.Length;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return stamp;
	}

	private static string NormalizePath(string path)
	{
		string NormalizePath;
		if (string.IsNullOrWhiteSpace(path))
		{
			NormalizePath = string.Empty;
		}
		else
		{
			try
			{
				NormalizePath = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).ToUpperInvariant();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				NormalizePath = path.Trim().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).ToUpperInvariant();
				ProjectData.ClearProjectError();
			}
		}
		return NormalizePath;
	}

	private static string HashText(string value)
	{
		using SHA256 sha = SHA256.Create();
		return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty))).Replace("-", string.Empty).ToLowerInvariant();
	}
}
