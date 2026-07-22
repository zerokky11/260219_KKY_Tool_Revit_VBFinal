using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
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

	private sealed class ProjectLookupIdentity
	{
		public string ProjectTitle { get; set; }

		public string DocumentPath { get; set; }

		public string CentralPath { get; set; }

		public string IdentityPath { get; set; }

		public string DocumentRevisionToken { get; set; }

		public bool HasLiveDocument { get; set; }

		public bool IsDocumentModified { get; set; }

		public ProjectLookupIdentity()
		{
			ProjectTitle = string.Empty;
			DocumentPath = string.Empty;
			CentralPath = string.Empty;
			IdentityPath = string.Empty;
			DocumentRevisionToken = string.Empty;
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
		string fileName = "project-snapshot-" + safeDocumentName + "-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-fffffff", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".json";
		string text = Path.Combine(outputDir, fileName);
		WriteJsonAtomic(text, snapshot);
		return text;
	}

	public static ProjectScanCacheRecord SaveLatestProjectScan(string workspaceRoot, Document doc, StandardLibraryRegistrationRecord registration, StandardLibrarySnapshot standardSnapshot, string projectSnapshotPath, string comparisonReportPath, ProjectContentSnapshot projectSnapshot, ProjectStandardComparisonReport report, string thumbnailSourceId)
	{
		FamilyBrowserRevitVersionContext.SetCurrentVersion(doc);
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Save latest project scan cache", "프로젝트 최신 스캔 캐시 저장"));
		string publicationReason;
		if (!CanPublishSharedProjectState(doc, out publicationReason))
		{
			throw new InvalidOperationException(publicationReason);
		}
		FamilyBrowserStandardRevisionState standardRevision = FamilyBrowserStandardRevisionService.Probe(workspaceRoot, registration, computeRevisionHash: true);
		string expectedSnapshotPath = registration?.LastSnapshotPath ?? string.Empty;
		string expectedSnapshotAtUtc = standardSnapshot?.CapturedAtUtc ?? registration?.LastSnapshotAtUtc ?? string.Empty;
		if (!FamilyBrowserStandardRevisionService.MatchesSnapshotGeneration(standardRevision, expectedSnapshotPath, expectedSnapshotAtUtc))
		{
			throw new InvalidOperationException("The Standard RVT or its registered snapshot changed before the project result was published. Run Current Model Check again with the latest standard snapshot.");
		}
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
		bool documentModified;
		projectScanCacheRecord.ProjectDocumentRevisionToken = BuildProjectDocumentRevisionToken(doc, out documentModified);
		projectScanCacheRecord.CapturedFromModifiedDocument = documentModified;
		if (documentModified)
		{
			throw new InvalidOperationException("A modified Revit document cannot be published as the latest shared project scan. Save or synchronize the project, then run Current Model Check again.");
		}
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
		string path = BuildProjectScanRecordPath(workspaceRoot, projectKey, registration);
		Directory.CreateDirectory(Path.GetDirectoryName(path));
		WriteJsonAtomic(path, record);
		string legacyPath = BuildProjectScanRecordPath(workspaceRoot, projectKey);
		if (!string.Equals(NormalizePath(path), NormalizePath(legacyPath), StringComparison.OrdinalIgnoreCase))
		{
			Directory.CreateDirectory(Path.GetDirectoryName(legacyPath));
			WriteJsonAtomic(legacyPath, record);
		}
		SaveProjectScanAliasRecords(workspaceRoot, projectKey, record, registration);
		FamilyBrowserDataLoader.PublishProjectState(workspaceRoot, registration, record, projectSnapshot, report);
		return record;
	}

	public static ProjectScanCacheLoadResult TryLoadLatestProjectScan(string workspaceRoot, Document doc, StandardLibraryRegistrationRecord registration, StandardLibrarySnapshot standardSnapshot)
	{
		FamilyBrowserRevitVersionContext.SetCurrentVersion(doc);
		string publicationReason;
		if (!CanPublishSharedProjectState(doc, out publicationReason))
		{
			return new ProjectScanCacheLoadResult
			{
				Reason = publicationReason
			};
		}
		return TryLoadLatestProjectScanCore(workspaceRoot, BuildProjectLookupIdentity(doc), registration, standardSnapshot);
	}

	public static ProjectScanCacheLoadResult TryLoadLatestProjectScan(string workspaceRoot, FamilyBrowserDeploymentProjectIdentity projectIdentity, StandardLibraryRegistrationRecord registration, StandardLibrarySnapshot standardSnapshot)
	{
		return TryLoadLatestProjectScanCore(workspaceRoot, BuildProjectLookupIdentity(projectIdentity), registration, standardSnapshot);
	}

	public static string BuildProjectScanCacheStamp(ProjectScanCacheLoadResult result)
	{
		if (result == null)
		{
			return "project-scan:not-preloaded";
		}
		if (!result.Success || result.Record == null)
		{
			return "project-scan:miss:" + (result.Reason ?? string.Empty).Replace("|", "/");
		}
		ProjectScanCacheRecord record = result.Record;
		string value = string.Join("|", new string[12]
		{
			record.ProjectKey ?? string.Empty,
			record.ProjectDocumentRevisionToken ?? string.Empty,
			record.SavedAtUtc ?? string.Empty,
			record.CapturedAtUtc ?? string.Empty,
			record.StandardSourceId ?? string.Empty,
			record.StandardSnapshotAtUtc ?? string.Empty,
			record.ProjectFileLastWriteUtc ?? string.Empty,
			record.ProjectFileLength.ToString(CultureInfo.InvariantCulture),
			record.StandardSourceFileLastWriteUtc ?? string.Empty,
			record.StandardSourceFileLength.ToString(CultureInfo.InvariantCulture),
			record.ProjectSnapshotPath ?? string.Empty,
			record.ComparisonReportPath ?? string.Empty
		});
		return "project-scan:hit:" + HashText(value);
	}

	private static ProjectScanCacheLoadResult TryLoadLatestProjectScanCore(string workspaceRoot, ProjectLookupIdentity projectIdentity, StandardLibraryRegistrationRecord registration, StandardLibrarySnapshot standardSnapshot)
	{
		ProjectScanCacheLoadResult result = new ProjectScanCacheLoadResult();
		string projectKey = ResolveProjectCacheKey(workspaceRoot, projectIdentity);
		if (string.IsNullOrWhiteSpace(projectKey))
		{
			result.Reason = "No project key.";
			return result;
		}

		string recordPath = BuildProjectScanRecordPath(workspaceRoot, projectKey, registration);
		if (!File.Exists(recordPath))
		{
			string legacyRecordPath = BuildProjectScanRecordPath(workspaceRoot, projectKey);
			if (!File.Exists(legacyRecordPath))
			{
				result.Reason = "No cached project scan record.";
				return result;
			}
			recordPath = legacyRecordPath;
		}

		ProjectScanCacheRecord record;
		try
		{
			record = DataContractJsonFileStore.Load<ProjectScanCacheRecord>(recordPath);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			result.Reason = "Cached project scan record could not be read: " + ex.Message;
			ProjectData.ClearProjectError();
			return result;
		}
		if (record == null)
		{
			result.Reason = "Cached project scan record was empty.";
			return result;
		}
		if (record.SchemaVersion < 4)
		{
			result.Reason = "Cached project scan predates dirty-document validation and must be rebuilt.";
			return result;
		}
		if (record.CapturedFromModifiedDocument)
		{
			result.Reason = "Cached project scan was captured from a document with unsaved changes.";
			return result;
		}
		if (!string.Equals(record.ProjectKey ?? string.Empty, projectKey, StringComparison.Ordinal))
		{
			result.Reason = "Cached project key does not match.";
			return result;
		}
		if (projectIdentity != null && projectIdentity.HasLiveDocument)
		{
			if (projectIdentity.IsDocumentModified)
			{
				result.Reason = "Cached project scan is stale: the live document has unsaved changes.";
				return result;
			}
			if (string.IsNullOrWhiteSpace(record.ProjectDocumentRevisionToken))
			{
				result.Reason = "Cached project scan has no live document revision token.";
				return result;
			}
			if (string.IsNullOrWhiteSpace(projectIdentity.DocumentRevisionToken) || !string.Equals(record.ProjectDocumentRevisionToken, projectIdentity.DocumentRevisionToken, StringComparison.Ordinal))
			{
				result.Reason = "Cached project scan is stale: the Revit document revision changed.";
				return result;
			}
		}
		if (registration != null)
		{
			if (!string.Equals(record.StandardSourceId ?? string.Empty, registration.SourceId ?? string.Empty, StringComparison.OrdinalIgnoreCase))
			{
				result.Reason = "Cached standard source id does not match.";
				return result;
			}
			if (!string.Equals(NormalizePath(record.StandardSnapshotPath), NormalizePath(registration.LastSnapshotPath), StringComparison.OrdinalIgnoreCase))
			{
				result.Reason = "Cached standard snapshot path does not match.";
				return result;
			}
			if (!string.Equals(record.StandardSourceFileLastWriteUtc ?? string.Empty, registration.SourceFileLastWriteUtc ?? string.Empty, StringComparison.Ordinal) || record.StandardSourceFileLength != registration.SourceFileLength)
			{
				result.Reason = "Cached standard RVT stamp is stale.";
				return result;
			}
		}
		if (standardSnapshot != null && !string.IsNullOrWhiteSpace(standardSnapshot.CapturedAtUtc) && !string.Equals(record.StandardSnapshotAtUtc ?? string.Empty, standardSnapshot.CapturedAtUtc, StringComparison.Ordinal))
		{
			result.Reason = "Cached standard snapshot timestamp is stale.";
			return result;
		}

		string currentStampPath = projectIdentity == null ? string.Empty : projectIdentity.IdentityPath;
		FileStamp currentStamp = BuildFileStamp(currentStampPath);
		bool identityChanged = IsProjectAliasMatch(record, projectIdentity) && !string.Equals(NormalizePath(record.ProjectIdentityPath), NormalizePath(currentStampPath), StringComparison.OrdinalIgnoreCase);
		if (identityChanged && !string.IsNullOrWhiteSpace(record.ProjectIdentityPath))
		{
			FileStamp canonicalStamp = BuildFileStamp(record.ProjectIdentityPath);
			if (!string.IsNullOrWhiteSpace(canonicalStamp.LastWriteUtc) || canonicalStamp.Length > 0)
			{
				currentStamp = canonicalStamp;
			}
		}
		bool unavailableAliasStamp = identityChanged && string.IsNullOrWhiteSpace(currentStamp.LastWriteUtc);
		if (!unavailableAliasStamp && !string.IsNullOrWhiteSpace(record.ProjectFileLastWriteUtc) && !string.Equals(record.ProjectFileLastWriteUtc, currentStamp.LastWriteUtc, StringComparison.Ordinal))
		{
			result.Reason = "Cached project scan is stale: file timestamp changed.";
			return result;
		}
		if (!unavailableAliasStamp && record.ProjectFileLength > 0 && currentStamp.Length > 0 && record.ProjectFileLength != currentStamp.Length)
		{
			result.Reason = "Cached project scan is stale: file length changed.";
			return result;
		}
		if (string.IsNullOrWhiteSpace(record.ProjectSnapshotPath) || !File.Exists(record.ProjectSnapshotPath))
		{
			result.Reason = "Cached project snapshot file is missing.";
			return result;
		}
		if (string.IsNullOrWhiteSpace(record.ComparisonReportPath) || !File.Exists(record.ComparisonReportPath))
		{
			result.Reason = "Cached comparison report file is missing.";
			return result;
		}
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			result.Reason = "Cached project scan data could not be read: " + ex.Message;
			ProjectData.ClearProjectError();
			return result;
		}
		result.Record = record;
		result.Success = result.Report != null;
		if (!result.Success)
		{
			result.Reason = "Cached comparison report was empty.";
		}
		return result;
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

	public static string GetProjectCoordinationFolder(string workspaceRoot, Document doc)
	{
		FamilyBrowserRevitVersionContext.SetCurrentVersion(doc);
		ProjectLookupIdentity identity = BuildProjectLookupIdentity(doc);
		string coordinationKey = BuildProjectCoordinationKey(identity);
		if (string.IsNullOrWhiteSpace(coordinationKey))
		{
			coordinationKey = BuildProjectCacheKey(identity);
		}
		return Path.Combine(GetProjectScanCacheFolder(workspaceRoot), "_coordination", HashText(coordinationKey ?? string.Empty).Substring(0, 32));
	}

	public static FileStream TryAcquireProjectPublicationLock(string workspaceRoot, Document doc)
	{
		try
		{
			string folder = GetProjectCoordinationFolder(workspaceRoot, doc);
			Directory.CreateDirectory(folder);
			return new FileStream(Path.Combine(folder, "project-scan-publication.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.WriteThrough);
		}
		catch (IOException)
		{
			return null;
		}
		catch (UnauthorizedAccessException)
		{
			return null;
		}
	}

	public static bool CanPublishSharedProjectState(Document doc, out string reason)
	{
		reason = string.Empty;
		if (doc == null)
		{
			reason = FamilyBrowserLanguageService.Text("No active Revit project is available for shared publication.", "공용 결과를 저장할 활성 Revit 프로젝트가 없습니다.");
			return false;
		}
		if (SafeIsModified(doc))
		{
			reason = FamilyBrowserLanguageService.Text("Save or synchronize the project before publishing shared Family Browser data.", "Family Browser 공용 데이터를 저장하기 전에 프로젝트를 저장하거나 동기화하세요.");
			return false;
		}

		bool workshared;
		try
		{
			workshared = doc.IsWorkshared;
		}
		catch (Exception ex)
		{
			reason = FamilyBrowserLanguageService.Text("The project worksharing state could not be verified: ", "프로젝트 작업공유 상태를 확인하지 못했습니다: ") + ex.Message;
			return false;
		}
		if (!workshared)
		{
			return true;
		}

		string documentPath = SafeDocumentPath(doc);
		string centralPath = ResolveCentralPath(doc);
		if (LooksLikeNonFileModelPath(documentPath) || LooksLikeNonFileModelPath(centralPath))
		{
			// Revit cloud/model-server paths do not expose BasicFileInfo through the file system.
			// The live IsModified check above remains the strongest available publication boundary.
			return true;
		}
		if (string.IsNullOrWhiteSpace(documentPath) || !File.Exists(documentPath))
		{
			reason = FamilyBrowserLanguageService.Text("The workshared local file could not be verified. Reopen it from a reachable local/central path before publishing shared data.", "작업공유 로컬 파일을 확인하지 못했습니다. 접근 가능한 로컬/센트럴 경로에서 다시 연 뒤 공용 데이터를 저장하세요.");
			return false;
		}

		BasicFileInfo localInfo = null;
		BasicFileInfo centralInfo = null;
		try
		{
			localInfo = BasicFileInfo.Extract(documentPath);
			if (localInfo == null || !localInfo.IsWorkshared)
			{
				reason = FamilyBrowserLanguageService.Text("The local RVT worksharing metadata could not be verified.", "로컬 RVT의 작업공유 정보를 확인하지 못했습니다.");
				return false;
			}
			if (localInfo.IsLocal && !localInfo.AllLocalChangesSavedToCentral)
			{
				reason = FamilyBrowserLanguageService.Text("The local RVT contains changes that have not been synchronized with Central. Synchronize before publishing a shared check or catalog.", "로컬 RVT에 센트럴로 동기화되지 않은 변경사항이 있습니다. 공용 검사 또는 카탈로그를 저장하기 전에 동기화하세요.");
				return false;
			}
			if (localInfo.IsCentral)
			{
				return true;
			}
			if (!localInfo.IsLocal)
			{
				reason = FamilyBrowserLanguageService.Text("The workshared RVT is neither a verified local file nor the Central file.", "작업공유 RVT가 확인된 로컬 파일이나 센트럴 파일 상태가 아닙니다.");
				return false;
			}
			if (string.IsNullOrWhiteSpace(centralPath) || !File.Exists(centralPath))
			{
				reason = FamilyBrowserLanguageService.Text("The Central RVT is not reachable, so the local revision cannot be proven current.", "센트럴 RVT에 접근할 수 없어 로컬 revision이 최신인지 확인하지 못했습니다.");
				return false;
			}

			centralInfo = BasicFileInfo.Extract(centralPath);
			if (centralInfo == null || !centralInfo.IsWorkshared || !centralInfo.IsCentral)
			{
				reason = FamilyBrowserLanguageService.Text("The registered Central path does not expose a valid Central RVT revision.", "등록된 센트럴 경로에서 유효한 센트럴 RVT revision을 확인하지 못했습니다.");
				return false;
			}
			Guid localEpisode = localInfo.LatestCentralEpisodeGUID;
			Guid centralEpisode = centralInfo.LatestCentralEpisodeGUID;
			if (localEpisode != Guid.Empty && centralEpisode != Guid.Empty && localEpisode != centralEpisode)
			{
				reason = FamilyBrowserLanguageService.Text("The local RVT belongs to a different Central episode. Recreate or update the local file before publishing shared data.", "로컬 RVT의 센트럴 episode가 현재 센트럴과 다릅니다. 로컬 파일을 다시 만들거나 갱신한 뒤 공용 데이터를 저장하세요.");
				return false;
			}
			if (centralInfo.LatestCentralVersion > localInfo.LatestCentralVersion)
			{
				reason = FamilyBrowserLanguageService.Text("The local RVT is behind the current Central revision. Reload Latest or synchronize before publishing shared data.", "로컬 RVT가 현재 센트럴 revision보다 오래되었습니다. Reload Latest 또는 동기화 후 공용 데이터를 저장하세요.");
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			reason = FamilyBrowserLanguageService.Text("The local/Central RVT revision check failed: ", "로컬/센트럴 RVT revision 확인에 실패했습니다: ") + ex.Message;
			return false;
		}
		finally
		{
			if (centralInfo != null)
			{
				centralInfo.Dispose();
			}
			if (localInfo != null)
			{
				localInfo.Dispose();
			}
		}
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

	private static void WriteJsonAtomic(string path, object value)
	{
		string folder = Path.GetDirectoryName(path);
		if (!string.IsNullOrWhiteSpace(folder))
		{
			Directory.CreateDirectory(folder);
		}
		string temporaryPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
		try
		{
			File.WriteAllText(temporaryPath, PlainJsonReportWriter.Serialize(value), new UTF8Encoding(false));
			FamilyBrowserAtomicFileService.Promote(temporaryPath, path);
		}
		finally
		{
			if (File.Exists(temporaryPath))
			{
				File.Delete(temporaryPath);
			}
		}
	}

	private static string BuildProjectScanRecordPath(string workspaceRoot, string projectKey)
	{
		return Path.Combine(GetProjectHistoryFolderFromKey(workspaceRoot, projectKey), "project-scan-latest.json");
	}

	private static string BuildProjectScanRecordPath(string workspaceRoot, string projectKey, StandardLibraryRegistrationRecord registration)
	{
		string standardKey = BuildProjectScanStandardCacheKey(registration);
		if (string.IsNullOrWhiteSpace(standardKey))
		{
			return BuildProjectScanRecordPath(workspaceRoot, projectKey);
		}
		return Path.Combine(GetProjectHistoryFolderFromKey(workspaceRoot, projectKey), "project-scan-latest-" + HashText(standardKey).Substring(0, 16) + ".json");
	}

	private static string BuildProjectScanStandardCacheKey(StandardLibraryRegistrationRecord registration)
	{
		if (registration == null)
		{
			return string.Empty;
		}
		string sourceId = registration.SourceId ?? string.Empty;
		string resolvedPath = NormalizePath(registration.ResolvedPath ?? string.Empty);
		string locator = NormalizePath(registration.Locator ?? string.Empty);
		string snapshotPath = NormalizePath(registration.LastSnapshotPath ?? string.Empty);
		string key = string.Join("|", new string[4] { sourceId, resolvedPath, locator, snapshotPath });
		return string.IsNullOrWhiteSpace(key.Replace("|", string.Empty)) ? string.Empty : key;
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
				string folder = GetProjectHistoryFolderFromKey(workspaceRoot, projectKey);
				if (Directory.Exists(folder))
				{
					List<string> stamps = Directory.GetFiles(folder, "project-scan-latest*.json").OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).Select([SpecialName] (string path) =>
					{
						FileInfo info = new FileInfo(path);
						return path + "|" + info.Length.ToString(CultureInfo.InvariantCulture) + "|" + info.LastWriteTimeUtc.Ticks.ToString(CultureInfo.InvariantCulture);
					}).ToList();
					GetLatestProjectScanRecordCacheStamp = ((stamps.Count == 0) ? (recordPath + "|missing") : string.Join(";", stamps));
				}
				else
				{
					GetLatestProjectScanRecordCacheStamp = recordPath + "|missing";
				}
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
		return BuildProjectCacheKey(BuildProjectLookupIdentity(doc));
	}

	private static string BuildProjectCacheKey(ProjectLookupIdentity projectIdentity)
	{
		if (projectIdentity == null)
		{
			return string.Empty;
		}
		return BuildProjectCacheKey(projectIdentity.IdentityPath, projectIdentity.ProjectTitle);
	}

	private static string BuildProjectCoordinationKey(ProjectLookupIdentity projectIdentity)
	{
		if (projectIdentity == null)
		{
			return string.Empty;
		}
		string identity = FamilyBrowserPathIdentityService.GetComparableIdentity(projectIdentity.IdentityPath);
		if (!string.IsNullOrWhiteSpace(identity))
		{
			return "physical|" + identity;
		}
		return BuildProjectCacheKey(projectIdentity);
	}

	private static string ResolveProjectCacheKey(string workspaceRoot, Document doc)
	{
		return ResolveProjectCacheKey(workspaceRoot, BuildProjectLookupIdentity(doc));
	}

	private static string ResolveProjectCacheKey(string workspaceRoot, ProjectLookupIdentity projectIdentity)
	{
		string primaryKey = BuildProjectCacheKey(projectIdentity);
		if (string.IsNullOrWhiteSpace(GetProjectScanCacheFolder(workspaceRoot)))
		{
			return primaryKey;
		}
		List<string> candidateKeys = BuildProjectCacheKeyCandidates(projectIdentity);
		foreach (string candidateKey in candidateKeys)
		{
			ProjectScanCacheRecord record = TryReadProjectScanRecord(workspaceRoot, candidateKey);
			bool primaryCandidate = string.Equals(candidateKey, primaryKey, StringComparison.OrdinalIgnoreCase);
			if (record != null && !string.IsNullOrWhiteSpace(record.ProjectKey) && (primaryCandidate || IsSafeProjectAliasRecord(record, projectIdentity)))
			{
				return record.ProjectKey;
			}
		}
		return primaryKey;
	}

	private static List<string> BuildProjectCacheKeyCandidates(Document doc)
	{
		return BuildProjectCacheKeyCandidates(BuildProjectLookupIdentity(doc));
	}

	private static List<string> BuildProjectCacheKeyCandidates(ProjectLookupIdentity projectIdentity)
	{
		List<string> list = new List<string>();
		AddProjectKey(list, BuildProjectCacheKey(projectIdentity));
		if (projectIdentity != null)
		{
			AddProjectNameAliasKeys(list, projectIdentity.CentralPath);
			AddProjectNameAliasKeys(list, projectIdentity.DocumentPath);
			AddProjectNameAliasKeys(list, projectIdentity.ProjectTitle);
		}
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

	private static void SaveProjectScanAliasRecords(string workspaceRoot, string canonicalProjectKey, ProjectScanCacheRecord record, StandardLibraryRegistrationRecord registration)
	{
		if (string.IsNullOrWhiteSpace(canonicalProjectKey) || record == null || string.IsNullOrWhiteSpace(GetProjectScanCacheFolder(workspaceRoot)))
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
					string path = BuildProjectScanRecordPath(workspaceRoot, aliasKey, registration);
					Directory.CreateDirectory(Path.GetDirectoryName(path));
					WriteJsonAtomic(path, aliasRecord);
					string legacyPath = BuildProjectScanRecordPath(workspaceRoot, aliasKey);
					if (!string.Equals(NormalizePath(path), NormalizePath(legacyPath), StringComparison.OrdinalIgnoreCase))
					{
						Directory.CreateDirectory(Path.GetDirectoryName(legacyPath));
						WriteJsonAtomic(legacyPath, aliasRecord);
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
			ProjectDocumentRevisionToken = record.ProjectDocumentRevisionToken,
			CapturedFromModifiedDocument = record.CapturedFromModifiedDocument,
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
		if (string.IsNullOrWhiteSpace(projectKey) || string.IsNullOrWhiteSpace(GetProjectScanCacheFolder(workspaceRoot)))
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
		return IsProjectAliasMatch(record, BuildProjectLookupIdentity(doc));
	}

	private static bool IsProjectAliasMatch(ProjectScanCacheRecord record, ProjectLookupIdentity projectIdentity)
	{
		if (record == null || projectIdentity == null)
		{
			return false;
		}
		List<string> list = (from x in BuildProjectCacheKeyCandidates(projectIdentity)
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

	private static bool IsSafeProjectAliasRecord(ProjectScanCacheRecord record, Document doc)
	{
		return IsSafeProjectAliasRecord(record, BuildProjectLookupIdentity(doc));
	}

	private static bool IsSafeProjectAliasRecord(ProjectScanCacheRecord record, ProjectLookupIdentity projectIdentity)
	{
		if (!IsProjectAliasMatch(record, projectIdentity))
		{
			return false;
		}
		string recordedIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(record == null ? string.Empty : record.ProjectIdentityPath);
		string currentIdentity = FamilyBrowserPathIdentityService.GetComparableIdentity(projectIdentity == null ? string.Empty : projectIdentity.IdentityPath);
		if (string.IsNullOrWhiteSpace(recordedIdentity) || string.IsNullOrWhiteSpace(currentIdentity) || !string.Equals(recordedIdentity, currentIdentity, StringComparison.OrdinalIgnoreCase))
		{
			return false;
		}
		FileStamp currentStamp = BuildFileStamp(projectIdentity == null ? string.Empty : projectIdentity.IdentityPath);
		if (string.IsNullOrWhiteSpace(currentStamp.LastWriteUtc) && currentStamp.Length <= 0)
		{
			return false;
		}
		if (!string.IsNullOrWhiteSpace(record.ProjectFileLastWriteUtc) && !string.Equals(record.ProjectFileLastWriteUtc, currentStamp.LastWriteUtc, StringComparison.Ordinal))
		{
			return false;
		}
		return record.ProjectFileLength <= 0 || currentStamp.Length <= 0 || record.ProjectFileLength == currentStamp.Length;
	}

	private static ProjectLookupIdentity BuildProjectLookupIdentity(Document doc)
	{
		string centralPath = ResolveCentralPath(doc);
		string documentPath = SafeDocumentPath(doc);
		bool isModified;
		string revisionToken = BuildProjectDocumentRevisionToken(doc, out isModified);
		return new ProjectLookupIdentity
		{
			ProjectTitle = SafeDocumentTitle(doc),
			DocumentPath = documentPath,
			CentralPath = centralPath,
			IdentityPath = string.IsNullOrWhiteSpace(centralPath) ? documentPath : centralPath,
			DocumentRevisionToken = revisionToken,
			HasLiveDocument = true,
			IsDocumentModified = isModified
		};
	}

	private static ProjectLookupIdentity BuildProjectLookupIdentity(FamilyBrowserDeploymentProjectIdentity projectIdentity)
	{
		string documentPath = projectIdentity == null ? string.Empty : projectIdentity.ModelPath ?? string.Empty;
		string centralPath = projectIdentity == null ? string.Empty : projectIdentity.CentralPath ?? string.Empty;
		return new ProjectLookupIdentity
		{
			ProjectTitle = projectIdentity == null ? string.Empty : projectIdentity.ProjectTitle ?? string.Empty,
			DocumentPath = documentPath,
			CentralPath = centralPath,
			IdentityPath = string.IsNullOrWhiteSpace(centralPath) ? documentPath : centralPath
		};
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
			SafeDocumentTitle = doc?.Title ?? string.Empty;
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
			SafeDocumentPath = doc?.PathName ?? string.Empty;
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
			SafeRevitVersion = doc?.Application?.VersionNumber ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			SafeRevitVersion = string.Empty;
			ProjectData.ClearProjectError();
		}
		return SafeRevitVersion;
	}

	private static string BuildProjectDocumentRevisionToken(Document doc, out bool isModified)
	{
		isModified = SafeIsModified(doc);
		if (doc == null)
		{
			return string.Empty;
		}
		string documentPath = SafeDocumentPath(doc);
		bool shareableCentralRevision;
		string basicFileRevision = BuildBasicFileRevisionToken(documentPath, out shareableCentralRevision);
		if (shareableCentralRevision && !string.IsNullOrWhiteSpace(basicFileRevision))
		{
			return "central|" + basicFileRevision;
		}
		string documentVersion = BuildLiveDocumentVersionToken(doc);
		FileStamp fileStamp = BuildFileStamp(documentPath);
		if (string.IsNullOrWhiteSpace(documentVersion) && string.IsNullOrWhiteSpace(basicFileRevision) && string.IsNullOrWhiteSpace(fileStamp.LastWriteUtc) && fileStamp.Length <= 0)
		{
			return string.Empty;
		}
		return "document|" + documentVersion + "|basic=" + basicFileRevision + "|write=" + fileStamp.LastWriteUtc + "|length=" + fileStamp.Length.ToString(CultureInfo.InvariantCulture);
	}

	private static bool LooksLikeNonFileModelPath(string value)
	{
		string text = (value ?? string.Empty).Trim();
		int schemeSeparator = text.IndexOf("://", StringComparison.Ordinal);
		return schemeSeparator > 0 && !text.StartsWith("file://", StringComparison.OrdinalIgnoreCase);
	}

	private static string BuildBasicFileRevisionToken(string documentPath, out bool shareableCentralRevision)
	{
		shareableCentralRevision = false;
		if (string.IsNullOrWhiteSpace(documentPath) || !File.Exists(documentPath))
		{
			return string.Empty;
		}
		BasicFileInfo info = null;
		try
		{
			info = BasicFileInfo.Extract(documentPath);
			if (info == null)
			{
				return string.Empty;
			}
			if (info.IsWorkshared && (info.IsCentral || (info.IsLocal && info.AllLocalChangesSavedToCentral)))
			{
				shareableCentralRevision = info.LatestCentralEpisodeGUID != Guid.Empty || info.LatestCentralVersion > 0;
				if (shareableCentralRevision)
				{
					return "episode=" + info.LatestCentralEpisodeGUID.ToString("D") + ";version=" + info.LatestCentralVersion.ToString(CultureInfo.InvariantCulture);
				}
			}
			DocumentVersion version = info.GetDocumentVersion();
			return BuildDocumentVersionToken(version);
		}
		catch
		{
			return string.Empty;
		}
		finally
		{
			if (info != null)
			{
				info.Dispose();
			}
		}
	}

	private static string BuildLiveDocumentVersionToken(Document doc)
	{
		try
		{
			System.Reflection.MethodInfo method = typeof(Document).GetMethod("GetDocumentVersion", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static, null, new Type[1] { typeof(Document) }, null);
			DocumentVersion version = method == null ? null : method.Invoke(null, new object[1] { doc }) as DocumentVersion;
			return BuildDocumentVersionToken(version);
		}
		catch
		{
			return string.Empty;
		}
	}

	private static string BuildDocumentVersionToken(DocumentVersion version)
	{
		return version == null ? string.Empty : version.VersionGUID.ToString("D") + ":" + version.NumberOfSaves.ToString(CultureInfo.InvariantCulture);
	}

	private static bool SafeIsModified(Document doc)
	{
		try
		{
			return doc != null && doc.IsModified;
		}
		catch
		{
			return true;
		}
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
