using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserRequestStore
{
	[CompilerGenerated]
	internal sealed class _Closure_0024__12_002D0
	{
		public string _0024VB_0024Local_requestId;

		public _Closure_0024__12_002D0(_Closure_0024__12_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_requestId = arg0._0024VB_0024Local_requestId;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(FamilyBrowserRequestRecord x)
		{
			return string.Equals(x.RequestId, _0024VB_0024Local_requestId.Trim(), StringComparison.OrdinalIgnoreCase);
		}
	}

	private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

	private FamilyBrowserRequestStore()
	{
	}

	public static FamilyBrowserRequestSaveResult Save(string workspaceRoot, FamilyBrowserRequestRecord record)
	{
		return Save(workspaceRoot, null, record);
	}

	public static FamilyBrowserRequestSaveResult Save(string workspaceRoot, FamilyBrowserStandardPolicy policy, FamilyBrowserRequestRecord record)
	{
		return Save(workspaceRoot, policy, record, Enumerable.Empty<string>());
	}

	public static FamilyBrowserRequestSaveResult Save(string workspaceRoot, FamilyBrowserStandardPolicy policy, FamilyBrowserRequestRecord record, IEnumerable<string> attachmentPaths)
	{
		if (record == null)
		{
			throw new ArgumentNullException("record");
		}
		FamilyBrowserRequestStoreInfo storeInfo = FamilyBrowserRequestStoreBackendService.ResolveInfo(workspaceRoot, policy);
		string outputDir = FamilyBrowserRequestStoreBackendService.ResolveWritableFolder(workspaceRoot, policy, requireWritable: true);
		Directory.CreateDirectory(outputDir);
		if (string.IsNullOrWhiteSpace(record.RequestId))
		{
			record.RequestId = CreateRequestId(record.RequestKind);
		}
		record.Status = NormalizeStatus(record.Status);
		if (string.IsNullOrWhiteSpace(record.CreatedAtUtc))
		{
			record.CreatedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		}
		if (string.IsNullOrWhiteSpace(record.UpdatedAtUtc))
		{
			record.UpdatedAtUtc = record.CreatedAtUtc;
		}
		if (record.Attachments == null)
		{
			record.Attachments = new List<string>();
		}
		if (record.AttachmentFiles == null)
		{
			record.AttachmentFiles = new List<FamilyBrowserRequestAttachmentFile>();
		}
		if (record.History == null)
		{
			record.History = new List<string>();
		}
		CopyAttachmentFiles(outputDir, record, attachmentPaths);
		string baseName = record.RequestId + "-" + MakeSafeFileName(record.ItemName ?? "Untitled");
		string requestPath = Path.Combine(outputDir, baseName + ".json");
		string mailDraftPath = Path.Combine(outputDir, baseName + "-mail.txt");
		record.SourcePath = requestPath;
		WriteAllTextAtomic(requestPath, PlainJsonReportWriter.Serialize(record));
		WriteAllTextAtomic(mailDraftPath, BuildMailDraft(record));
		FamilyBrowserRequestStoreBackendService.WriteStoreManifest(outputDir, storeInfo);
		WriteAttachmentManifest(record);
		return new FamilyBrowserRequestSaveResult
		{
			RequestPath = requestPath,
			MailDraftPath = mailDraftPath,
			AttachmentFolder = record.AttachmentFolder,
			AttachmentCount = ((record.AttachmentFiles != null) ? record.AttachmentFiles.Count : 0),
			StoreMode = ((storeInfo == null) ? string.Empty : storeInfo.Mode),
			StoreLocation = ((storeInfo == null) ? outputDir : storeInfo.StoreLocation),
			ConnectorNote = ((storeInfo == null) ? string.Empty : storeInfo.Detail)
		};
	}

	public static List<FamilyBrowserRequestRecord> List(string workspaceRoot)
	{
		return List(workspaceRoot, null);
	}

	public static List<FamilyBrowserRequestRecord> List(string workspaceRoot, FamilyBrowserStandardPolicy policy)
	{
		List<FamilyBrowserRequestRecord> records = new List<FamilyBrowserRequestRecord>();
		FamilyBrowserRequestStoreInfo storeInfo = FamilyBrowserRequestStoreBackendService.ResolveInfo(workspaceRoot, policy);
		string unavailableDetail = string.Empty;
		if (FamilyBrowserRequestStoreBackendService.IsObviouslyUnavailable(storeInfo, ref unavailableDetail))
		{
			return records;
		}
		string outputDir = ((storeInfo == null) ? string.Empty : storeInfo.StoreLocation);
		if (string.IsNullOrWhiteSpace(outputDir))
		{
			return records;
		}
		if (!FamilyBrowserRequestStoreBackendService.DirectoryExistsFast(storeInfo))
		{
			return records;
		}
		foreach (string requestPath in EnumerateRequestFiles(outputDir))
		{
			FamilyBrowserRequestRecord record = null;
			Exception loadError = null;
			if (TryLoadRequestRecord(requestPath, ref record, ref loadError))
			{
				EnsureRecordDefaults(record, requestPath);
				records.Add(record);
			}
			else
			{
				records.Add(BuildUnreadableRequestRecord(requestPath, loadError));
				WriteReadIssue(outputDir, requestPath, loadError);
			}
		}
		return records.OrderByDescending([SpecialName] (FamilyBrowserRequestRecord x) => ParseUtcOrMin(x.CreatedAtUtc)).ThenBy<FamilyBrowserRequestRecord, string>([SpecialName] (FamilyBrowserRequestRecord x) => x.RequestId, StringComparer.OrdinalIgnoreCase).ToList();
	}

	private static IEnumerable<string> EnumerateRequestFiles(string outputDir)
	{
		SortedSet<string> requestPaths = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
		string[] files = Directory.GetFiles(outputDir, "FBR-*.json");
		foreach (string requestPath in files)
		{
			requestPaths.Add(requestPath);
		}
		string[] files2 = Directory.GetFiles(outputDir, "*.json");
		foreach (string jsonPath in files2)
		{
			if (IsRequestRecordCandidate(jsonPath))
			{
				requestPaths.Add(jsonPath);
			}
		}
		return requestPaths;
	}

	private static bool IsRequestRecordCandidate(string requestPath)
	{
		string fileName = Path.GetFileName(requestPath);
		bool IsRequestRecordCandidate;
		if (fileName.StartsWith("FBR-", StringComparison.OrdinalIgnoreCase))
		{
			IsRequestRecordCandidate = true;
		}
		else
		{
			switch (fileName.ToLowerInvariant())
			{
			case "request-store-info.json":
			case "attachment-manifest.json":
			case "standard-policy.json":
			case "active-standard-library.json":
				IsRequestRecordCandidate = false;
				break;
			default:
				try
				{
					string json = File.ReadAllText(requestPath, Encoding.UTF8);
					IsRequestRecordCandidate = json.IndexOf("\"RequestId\"", StringComparison.OrdinalIgnoreCase) >= 0 || json.IndexOf("\"RequestKind\"", StringComparison.OrdinalIgnoreCase) >= 0 || json.IndexOf("\"CreatedAtUtc\"", StringComparison.OrdinalIgnoreCase) >= 0;
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					IsRequestRecordCandidate = false;
					ProjectData.ClearProjectError();
				}
				break;
			}
		}
		return IsRequestRecordCandidate;
	}

	public static FamilyBrowserRequestSaveResult UpdateStatus(string workspaceRoot, string requestId, string status, string updatedBy, string progressNote)
	{
		return UpdateStatus(workspaceRoot, null, requestId, status, updatedBy, progressNote);
	}

	public static FamilyBrowserRequestSaveResult UpdateStatus(string workspaceRoot, FamilyBrowserStandardPolicy policy, string requestId, string status, string updatedBy, string progressNote)
	{
		FamilyBrowserRequestRecord record = List(workspaceRoot, policy).FirstOrDefault([SpecialName] (FamilyBrowserRequestRecord x) => string.Equals(x.RequestId, requestId, StringComparison.OrdinalIgnoreCase));
		if (record == null)
		{
			throw new FileNotFoundException(FamilyBrowserLanguageService.Text("Request was not found.", "요청을 찾지 못했습니다."), requestId);
		}
		string previousStatus = record.Status;
		record.Status = NormalizeStatus(status);
		record.UpdatedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		record.LastUpdatedBy = updatedBy ?? string.Empty;
		record.ProgressNote = progressNote ?? string.Empty;
		if (record.History == null)
		{
			record.History = new List<string>();
		}
		record.History.Add(record.UpdatedAtUtc + " | " + record.LastUpdatedBy + " | " + previousStatus + " -> " + record.Status + " | " + record.ProgressNote);
		return Save(workspaceRoot, policy, record);
	}

	public static void Delete(string workspaceRoot, FamilyBrowserStandardPolicy policy, string requestId, IEnumerable<string> currentUserIdentities, bool allowAdminDelete)
	{
		_Closure_0024__12_002D0 arg = default(_Closure_0024__12_002D0);
		_Closure_0024__12_002D0 CS_0024_003C_003E8__locals4 = new _Closure_0024__12_002D0(arg);
		CS_0024_003C_003E8__locals4._0024VB_0024Local_requestId = requestId;
		if (string.IsNullOrWhiteSpace(CS_0024_003C_003E8__locals4._0024VB_0024Local_requestId))
		{
			throw new ArgumentException(FamilyBrowserLanguageService.Text("Request id is empty.", "요청 ID가 비어 있습니다."), "requestId");
		}
		FamilyBrowserRequestStoreInfo storeInfo = FamilyBrowserRequestStoreBackendService.ResolveInfo(workspaceRoot, policy);
		string obj = ((storeInfo == null) ? string.Empty : storeInfo.StoreLocation);
		if (string.IsNullOrWhiteSpace(obj))
		{
			throw new DirectoryNotFoundException("Request store folder is not configured.");
		}
		if (!Directory.Exists(obj))
		{
			throw new DirectoryNotFoundException("Request store folder was not found.");
		}
		string rootPath = NormalizeAbsolutePath(obj);
		FamilyBrowserRequestRecord record = List(workspaceRoot, policy).FirstOrDefault([SpecialName] (FamilyBrowserRequestRecord x) => string.Equals(x.RequestId, CS_0024_003C_003E8__locals4._0024VB_0024Local_requestId.Trim(), StringComparison.OrdinalIgnoreCase));
		if (record == null)
		{
			throw new FileNotFoundException(FamilyBrowserLanguageService.Text("Request was not found.", "요청을 찾지 못했습니다."), CS_0024_003C_003E8__locals4._0024VB_0024Local_requestId);
		}
		if (!allowAdminDelete && !RequestBelongsToAnyIdentity(record, currentUserIdentities))
		{
			throw new UnauthorizedAccessException("Only the request creator or an administrator can delete this request.");
		}
		string requestPath = NormalizeAbsolutePath(record.SourcePath);
		EnsurePathInsideRoot(requestPath, rootPath, "request file");
		string mailDraftPath = string.Empty;
		if (!string.IsNullOrWhiteSpace(requestPath))
		{
			string requestFolder = Path.GetDirectoryName(requestPath);
			string requestName = Path.GetFileNameWithoutExtension(requestPath);
			if (!string.IsNullOrWhiteSpace(requestFolder) && !string.IsNullOrWhiteSpace(requestName))
			{
				mailDraftPath = Path.Combine(requestFolder, requestName + "-mail.txt");
			}
		}
		DeleteFileInsideRoot(requestPath, rootPath);
		DeleteFileInsideRoot(mailDraftPath, rootPath);
		if (record.AttachmentFiles != null)
		{
			foreach (FamilyBrowserRequestAttachmentFile item in record.AttachmentFiles.Where([SpecialName] (FamilyBrowserRequestAttachmentFile x) => x != null))
			{
				DeleteFileInsideRoot(item.StoredPath, rootPath);
			}
		}
		if (IsRequestSpecificAttachmentFolder(record.AttachmentFolder, record.RequestId))
		{
			DeleteDirectoryInsideRoot(record.AttachmentFolder, rootPath);
		}
	}

	public static string GetRequestFolder(string workspaceRoot)
	{
		return GetRequestFolder(workspaceRoot, null);
	}

	public static string GetRequestFolder(string workspaceRoot, FamilyBrowserStandardPolicy policy)
	{
		FamilyBrowserRequestStoreInfo info = GetRequestStoreInfo(workspaceRoot, policy);
		if (info != null)
		{
			return info.StoreLocation;
		}
		return string.Empty;
	}

	public static FamilyBrowserRequestStoreInfo GetRequestStoreInfo(string workspaceRoot, FamilyBrowserStandardPolicy policy)
	{
		return FamilyBrowserRequestStoreBackendService.ResolveInfo(workspaceRoot, policy);
	}

	public static string Describe(string workspaceRoot, FamilyBrowserStandardPolicy policy)
	{
		return FamilyBrowserRequestStoreBackendService.BuildSummary(workspaceRoot, policy);
	}

	public static string CreateRequestId(string requestKind)
	{
		string kind = NormalizeKind(requestKind);
		return "FBR-" + kind + "-" + DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N").Substring(0, 8);
	}

	private static string BuildMailDraft(FamilyBrowserRequestRecord record)
	{
		StringBuilder builder = new StringBuilder();
		builder.AppendLine("Subject: [KKY Family Browser] " + ResolveKindTitle(record.RequestKind) + " - " + (record.ItemName ?? string.Empty));
		builder.AppendLine();
		builder.AppendLine("Request ID: " + record.RequestId);
		builder.AppendLine("Request Kind: " + record.RequestKind);
		builder.AppendLine("Status: " + record.Status);
		builder.AppendLine("Created: " + record.CreatedAtUtc);
		builder.AppendLine("Created By: " + record.CreatedBy);
		builder.AppendLine("Updated: " + record.UpdatedAtUtc);
		builder.AppendLine("Updated By: " + record.LastUpdatedBy);
		builder.AppendLine();
		builder.AppendLine("Project: " + record.ProjectTitle);
		builder.AppendLine("Project Path: " + record.ProjectPath);
		builder.AppendLine("Central Path: " + record.CentralPath);
		builder.AppendLine();
		builder.AppendLine("Standard Target: " + record.StandardTarget);
		builder.AppendLine("Standard Mode: " + record.StandardMode);
		builder.AppendLine("Standard RVT: " + record.StandardRvtPath);
		builder.AppendLine("Standard Source: " + record.StandardDisplayName);
		builder.AppendLine();
		builder.AppendLine("Item Name: " + record.ItemName);
		builder.AppendLine("Category: " + record.CategoryName);
		builder.AppendLine("Discipline: " + record.Discipline);
		builder.AppendLine("Suggested Action: " + record.SuggestedAction);
		builder.AppendLine();
		builder.AppendLine("Reason");
		builder.AppendLine(record.Reason);
		builder.AppendLine();
		builder.AppendLine("Notes");
		builder.AppendLine(record.Notes);
		builder.AppendLine();
		builder.AppendLine("Progress");
		builder.AppendLine(record.ProgressNote);
		builder.AppendLine();
		builder.AppendLine("Comparison: " + record.ComparisonPath);
		builder.AppendLine("Preflight: " + record.PreflightPath);
		builder.AppendLine("Tracking: " + record.TrackingState);
		if (record.Attachments != null && record.Attachments.Count > 0)
		{
			builder.AppendLine();
			builder.AppendLine("Attachments");
			foreach (string attachment in record.Attachments.Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)))
			{
				builder.AppendLine("- " + attachment);
			}
		}
		if (!string.IsNullOrWhiteSpace(record.AttachmentFolder))
		{
			builder.AppendLine();
			builder.AppendLine("Attachment Folder");
			builder.AppendLine(record.AttachmentFolder);
		}
		if (record.AttachmentFiles != null && record.AttachmentFiles.Count > 0)
		{
			builder.AppendLine();
			builder.AppendLine("Attachment Files");
			foreach (FamilyBrowserRequestAttachmentFile attachmentFile in record.AttachmentFiles.Where([SpecialName] (FamilyBrowserRequestAttachmentFile x) => x != null))
			{
				builder.AppendLine("- " + attachmentFile.DisplayName + " | " + attachmentFile.StoredPath);
			}
		}
		if (record.History != null && record.History.Count > 0)
		{
			builder.AppendLine();
			builder.AppendLine("History");
			foreach (string historyItem in record.History.Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)))
			{
				builder.AppendLine("- " + historyItem);
			}
		}
		return builder.ToString();
	}

	private static string ResolveKindTitle(string requestKind)
	{
		return NormalizeKind(requestKind) switch
		{
			"FAMILY" => "Family Request", 
			"SYSTEM" => "System Type Request", 
			"UPDATE" => "Update Approval Request", 
			_ => "Request", 
		};
	}

	private static string NormalizeKind(string requestKind)
	{
		string value = (requestKind ?? string.Empty).Trim().ToUpperInvariant();
		switch (value)
		{
		case "FAMILY":
		case "SYSTEM":
		case "UPDATE":
			return value;
		default:
			return "GENERAL";
		}
	}

	private static void CopyAttachmentFiles(string outputDir, FamilyBrowserRequestRecord record, IEnumerable<string> attachmentPaths)
	{
		if (attachmentPaths == null)
		{
			return;
		}
		List<string> validPaths = (from x in attachmentPaths
			where !string.IsNullOrWhiteSpace(x)
			select Environment.ExpandEnvironmentVariables(x.Trim()) into x
			where File.Exists(x)
			select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).ToList();
		if (validPaths.Count == 0)
		{
			return;
		}
		string attachmentFolder = BuildAttachmentFolder(outputDir, record);
		Directory.CreateDirectory(attachmentFolder);
		record.AttachmentFolder = attachmentFolder;
		string attachedAt = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		string attachedBy = ((!string.IsNullOrWhiteSpace(record.LastUpdatedBy)) ? record.LastUpdatedBy : record.CreatedBy);
		if (string.IsNullOrWhiteSpace(attachedBy))
		{
			attachedBy = Environment.UserName;
		}
		foreach (string sourcePath in validPaths)
		{
			FileInfo fileInfo = new FileInfo(sourcePath);
			string storedPath = MakeUniqueFilePath(Path.Combine(attachmentFolder, MakeSafeFileName(fileInfo.Name)));
			File.Copy(sourcePath, storedPath, overwrite: false);
			FamilyBrowserRequestAttachmentFile attachmentFile = new FamilyBrowserRequestAttachmentFile
			{
				DisplayName = fileInfo.Name,
				OriginalPath = sourcePath,
				StoredPath = storedPath,
				RelativePath = MakeRelativePath(outputDir, storedPath),
				SizeBytes = fileInfo.Length,
				AttachedAtUtc = attachedAt,
				AttachedBy = attachedBy
			};
			record.AttachmentFiles.Add(attachmentFile);
			record.Attachments.Add(storedPath);
		}
	}

	private static void WriteAttachmentManifest(FamilyBrowserRequestRecord record)
	{
		if (record != null && !string.IsNullOrWhiteSpace(record.AttachmentFolder) && record.AttachmentFiles != null && record.AttachmentFiles.Count != 0)
		{
			Directory.CreateDirectory(record.AttachmentFolder);
			WriteAllTextAtomic(Path.Combine(record.AttachmentFolder, "attachment-manifest.json"), PlainJsonReportWriter.Serialize(record.AttachmentFiles));
			string path = Path.Combine(record.AttachmentFolder, "request-context.txt");
			StringBuilder builder = new StringBuilder();
			builder.AppendLine("Request ID: " + record.RequestId);
			builder.AppendLine("Request Kind: " + record.RequestKind);
			builder.AppendLine("Status: " + record.Status);
			builder.AppendLine("Item: " + record.ItemName);
			builder.AppendLine("Category: " + record.CategoryName);
			builder.AppendLine("Discipline: " + record.Discipline);
			builder.AppendLine("Created By: " + record.CreatedBy);
			builder.AppendLine("Created At: " + record.CreatedAtUtc);
			builder.AppendLine("Request JSON: " + record.SourcePath);
			WriteAllTextAtomic(path, builder.ToString());
		}
	}

	private static void WriteAllTextAtomic(string path, string contents)
	{
		Directory.CreateDirectory(Path.GetDirectoryName(path));
		string tempPath = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
		File.WriteAllText(tempPath, contents ?? string.Empty, Utf8NoBom);
		if (File.Exists(path))
		{
			try
			{
				File.Replace(tempPath, path, null);
				return;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				File.Delete(path);
				ProjectData.ClearProjectError();
			}
		}
		File.Move(tempPath, path);
	}

	private static string BuildAttachmentFolder(string outputDir, FamilyBrowserRequestRecord record)
	{
		string discipline = ((!string.IsNullOrWhiteSpace(record.Discipline)) ? record.Discipline : record.StandardTarget);
		if (string.IsNullOrWhiteSpace(discipline))
		{
			discipline = "Unassigned";
		}
		string kind = ResolveKindTitle(record.RequestKind).Replace(" Request", string.Empty).Replace(" Approval", string.Empty);
		if (string.IsNullOrWhiteSpace(kind))
		{
			kind = "Request";
		}
		string itemName = (string.IsNullOrWhiteSpace(record.ItemName) ? "Untitled" : record.ItemName);
		return Path.Combine(outputDir, "Attachments", MakeSafeFileName(discipline), MakeSafeFileName(kind), MakeSafeFileName(itemName), MakeSafeFileName(record.RequestId));
	}

	private static string MakeUniqueFilePath(string candidatePath)
	{
		if (!File.Exists(candidatePath))
		{
			return candidatePath;
		}
		string folder = Path.GetDirectoryName(candidatePath);
		string name = Path.GetFileNameWithoutExtension(candidatePath);
		string extension = Path.GetExtension(candidatePath);
		int index = 2;
		string nextPath;
		while (true)
		{
			nextPath = Path.Combine(folder, name + "-" + index.ToString(CultureInfo.InvariantCulture) + extension);
			if (!File.Exists(nextPath))
			{
				break;
			}
			index = checked(index + 1);
		}
		return nextPath;
	}

	private static string MakeRelativePath(string rootPath, string fullPath)
	{
		if (string.IsNullOrWhiteSpace(rootPath) || string.IsNullOrWhiteSpace(fullPath))
		{
			return fullPath;
		}
		string root = Path.GetFullPath(rootPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Conversions.ToString(Path.DirectorySeparatorChar);
		string target = Path.GetFullPath(fullPath);
		if (target.StartsWith(root, StringComparison.OrdinalIgnoreCase))
		{
			return target.Substring(root.Length);
		}
		return fullPath;
	}

	private static string NormalizeAbsolutePath(string value)
	{
		if (string.IsNullOrWhiteSpace(value))
		{
			return string.Empty;
		}
		return Path.GetFullPath(Environment.ExpandEnvironmentVariables(value.Trim())).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
	}

	private static void EnsurePathInsideRoot(string candidatePath, string rootPath, string description)
	{
		if (string.IsNullOrWhiteSpace(candidatePath) || string.IsNullOrWhiteSpace(rootPath))
		{
			throw new InvalidOperationException(description + " path is empty.");
		}
		if (string.Equals(candidatePath, rootPath, StringComparison.OrdinalIgnoreCase))
		{
			throw new InvalidOperationException(description + " path points to the request store root.");
		}
		string rooted = rootPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Conversions.ToString(Path.DirectorySeparatorChar);
		if (!candidatePath.StartsWith(rooted, StringComparison.OrdinalIgnoreCase))
		{
			throw new InvalidOperationException(description + " path is outside the request store.");
		}
	}

	private static void DeleteFileInsideRoot(string filePath, string rootPath)
	{
		if (string.IsNullOrWhiteSpace(filePath))
		{
			return;
		}
		string normalizedPath = NormalizeAbsolutePath(filePath);
		if (!string.IsNullOrWhiteSpace(normalizedPath))
		{
			EnsurePathInsideRoot(normalizedPath, rootPath, "request file");
			if (File.Exists(normalizedPath))
			{
				File.Delete(normalizedPath);
			}
		}
	}

	private static void DeleteDirectoryInsideRoot(string folderPath, string rootPath)
	{
		if (string.IsNullOrWhiteSpace(folderPath))
		{
			return;
		}
		string normalizedPath = NormalizeAbsolutePath(folderPath);
		if (!string.IsNullOrWhiteSpace(normalizedPath))
		{
			EnsurePathInsideRoot(normalizedPath, rootPath, "request attachment folder");
			if (Directory.Exists(normalizedPath))
			{
				Directory.Delete(normalizedPath, recursive: true);
			}
		}
	}

	private static bool IsRequestSpecificAttachmentFolder(string folderPath, string requestId)
	{
		bool IsRequestSpecificAttachmentFolder;
		if (string.IsNullOrWhiteSpace(folderPath) || string.IsNullOrWhiteSpace(requestId))
		{
			IsRequestSpecificAttachmentFolder = false;
		}
		else
		{
			try
			{
				IsRequestSpecificAttachmentFolder = string.Equals(Path.GetFileName(NormalizeAbsolutePath(folderPath)), MakeSafeFileName(requestId), StringComparison.OrdinalIgnoreCase);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				IsRequestSpecificAttachmentFolder = false;
				ProjectData.ClearProjectError();
			}
		}
		return IsRequestSpecificAttachmentFolder;
	}

	private static bool RequestBelongsToAnyIdentity(FamilyBrowserRequestRecord record, IEnumerable<string> identities)
	{
		if (record == null)
		{
			return false;
		}
		HashSet<string> ownerKeys = BuildIdentityKeys(record.CreatedBy);
		if (ownerKeys.Count == 0)
		{
			return false;
		}
		if (identities == null)
		{
			return false;
		}
		foreach (string identity in identities)
		{
			foreach (string key in BuildIdentityKeys(identity))
			{
				if (ownerKeys.Contains(key))
				{
					return true;
				}
			}
		}
		return false;
	}

	private static HashSet<string> BuildIdentityKeys(string identity)
	{
		HashSet<string> keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		string value = (identity ?? string.Empty).Trim();
		if (value.Length == 0)
		{
			return keys;
		}
		keys.Add(value);
		int slashIndex = value.LastIndexOf('\\');
		checked
		{
			if (slashIndex >= 0 && slashIndex < value.Length - 1)
			{
				keys.Add(value.Substring(slashIndex + 1));
			}
			int atIndex = value.IndexOf('@');
			if (atIndex > 0)
			{
				keys.Add(value.Substring(0, atIndex));
			}
			return keys;
		}
	}

	private static string ResolveFileBackedRequestFolder(string workspaceRoot, FamilyBrowserStandardPolicy policy, bool requireWritable)
	{
		return FamilyBrowserRequestStoreBackendService.ResolveWritableFolder(workspaceRoot, policy, requireWritable);
	}

	private static string ResolveStoreMode(FamilyBrowserStandardPolicy policy)
	{
		if (policy == null || policy.RequestStore == null)
		{
			return "Local";
		}
		switch (FamilyBrowserPolicyKey.Normalize(policy.RequestStore.Mode))
		{
		case "networkshare":
		case "network-share":
			return "NetworkShare";
		case "sharepoint":
		case "share-point":
			return "SharePoint";
		case "cloudstorage":
		case "cloud-storage":
			return "CloudStorage";
		case "api":
			return "Api";
		default:
			return "Local";
		}
	}

	private static string ResolveConfiguredPath(FamilyBrowserStandardPolicy policy)
	{
		if (policy == null || policy.RequestStore == null)
		{
			return string.Empty;
		}
		return (policy.RequestStore.Path ?? string.Empty).Trim();
	}

	private static string ResolveConfiguredEndpoint(FamilyBrowserStandardPolicy policy)
	{
		if (policy == null || policy.RequestStore == null)
		{
			return string.Empty;
		}
		return (policy.RequestStore.Endpoint ?? string.Empty).Trim();
	}

	private static string GetLocalRequestFolder(string workspaceRoot)
	{
		return FamilyBrowserStandardPolicyStore.GetDataFolder(workspaceRoot, "Requests");
	}

	private static string NormalizeStatus(string status)
	{
		switch ((status ?? string.Empty).Trim().ToUpperInvariant())
		{
		case "SUBMITTED":
			return "Submitted";
		case "REVIEWING":
		case "INREVIEW":
		case "IN_REVIEW":
			return "Reviewing";
		case "APPROVED":
			return "Approved";
		case "REJECTED":
			return "Rejected";
		case "COMPLETED":
		case "DONE":
		case "CLOSED":
			return "Completed";
		default:
			return "Draft";
		}
	}

	private static void EnsureRecordDefaults(FamilyBrowserRequestRecord record, string requestPath)
	{
		if (record == null)
		{
			throw new InvalidDataException("Request record is empty.");
		}
		if (string.IsNullOrWhiteSpace(record.RequestId))
		{
			record.RequestId = Path.GetFileNameWithoutExtension(requestPath);
		}
		record.RequestKind = NormalizeKind(record.RequestKind);
		record.Status = NormalizeStatus(record.Status);
		record.SourcePath = requestPath;
		if (string.IsNullOrWhiteSpace(record.UpdatedAtUtc))
		{
			record.UpdatedAtUtc = record.CreatedAtUtc;
		}
		if (record.Attachments == null)
		{
			record.Attachments = new List<string>();
		}
		if (record.AttachmentFiles == null)
		{
			record.AttachmentFiles = new List<FamilyBrowserRequestAttachmentFile>();
		}
		if (string.IsNullOrWhiteSpace(record.AttachmentFolder) && record.AttachmentFiles.Count > 0)
		{
			record.AttachmentFolder = Path.GetDirectoryName(record.AttachmentFiles[0].StoredPath);
		}
		if (record.History == null)
		{
			record.History = new List<string>();
		}
	}

	private static bool TryLoadRequestRecord(string requestPath, ref FamilyBrowserRequestRecord record, ref Exception loadError)
	{
		record = null;
		loadError = null;
		try
		{
			record = DataContractJsonFileStore.Load<FamilyBrowserRequestRecord>(requestPath);
			return record != null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			loadError = ex2;
			ProjectData.ClearProjectError();
		}
		try
		{
			record = LoadRequestRecordFallback(requestPath);
			return record != null;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			if (loadError == null)
			{
				loadError = ex4;
			}
			ProjectData.ClearProjectError();
		}
		return false;
	}

	private static FamilyBrowserRequestRecord LoadRequestRecordFallback(string requestPath)
	{
		string json = File.ReadAllText(requestPath, Encoding.UTF8);
		FamilyBrowserRequestRecord record = new FamilyBrowserRequestRecord
		{
			RequestId = ReadJsonString(json, "RequestId"),
			RequestKind = ReadJsonString(json, "RequestKind"),
			Status = ReadJsonString(json, "Status"),
			CreatedAtUtc = ReadJsonString(json, "CreatedAtUtc"),
			CreatedBy = ReadJsonString(json, "CreatedBy"),
			UpdatedAtUtc = ReadJsonString(json, "UpdatedAtUtc"),
			LastUpdatedBy = ReadJsonString(json, "LastUpdatedBy"),
			ProjectTitle = ReadJsonString(json, "ProjectTitle"),
			ProjectPath = ReadJsonString(json, "ProjectPath"),
			CentralPath = ReadJsonString(json, "CentralPath"),
			StandardTarget = ReadJsonString(json, "StandardTarget"),
			StandardMode = ReadJsonString(json, "StandardMode"),
			StandardDisplayName = ReadJsonString(json, "StandardDisplayName"),
			StandardRvtPath = ReadJsonString(json, "StandardRvtPath"),
			StandardSourceId = ReadJsonString(json, "StandardSourceId"),
			ComparisonPath = ReadJsonString(json, "ComparisonPath"),
			PreflightPath = ReadJsonString(json, "PreflightPath"),
			TrackingState = ReadJsonString(json, "TrackingState"),
			ItemName = ReadJsonString(json, "ItemName"),
			CategoryName = ReadJsonString(json, "CategoryName"),
			Discipline = ReadJsonString(json, "Discipline"),
			SuggestedAction = ReadJsonString(json, "SuggestedAction"),
			Reason = ReadJsonString(json, "Reason"),
			Notes = ReadJsonString(json, "Notes"),
			ProgressNote = ReadJsonString(json, "ProgressNote"),
			SourcePath = ReadJsonString(json, "SourcePath"),
			AttachmentFolder = ReadJsonString(json, "AttachmentFolder")
		};
		if (string.IsNullOrWhiteSpace(record.RequestId) && string.IsNullOrWhiteSpace(record.ItemName) && string.IsNullOrWhiteSpace(record.CreatedAtUtc))
		{
			throw new InvalidDataException("Request JSON did not contain expected request fields.");
		}
		return record;
	}

	private static string ReadJsonString(string json, string propertyName)
	{
		if (string.IsNullOrWhiteSpace(json) || string.IsNullOrWhiteSpace(propertyName))
		{
			return string.Empty;
		}
		string pattern = "\"" + Regex.Escape(propertyName) + "\"\\s*:\\s*\"(?<value>(?:\\\\.|[^\"\\\\])*)\"";
		Match match = Regex.Match(json, pattern, RegexOptions.CultureInvariant);
		if (!match.Success)
		{
			return string.Empty;
		}
		return DecodeJsonString(match.Groups["value"].Value);
	}

	private static string DecodeJsonString(string value)
	{
		if (string.IsNullOrEmpty(value))
		{
			return string.Empty;
		}
		StringBuilder builder = new StringBuilder(value.Length);
		int index = 0;
		checked
		{
			while (index < value.Length)
			{
				char ch = value[index];
				if (ch != '\\' || index == value.Length - 1)
				{
					builder.Append(ch);
					index++;
					continue;
				}
				index++;
				char escaped = value[index];
				switch (escaped)
				{
				case '"':
					builder.Append('"');
					break;
				case '\\':
					builder.Append('\\');
					break;
				case '/':
					builder.Append('/');
					break;
				case 'b':
					builder.Append('\b');
					break;
				case 'f':
					builder.Append('\f');
					break;
				case 'n':
					builder.Append('\n');
					break;
				case 'r':
					builder.Append('\r');
					break;
				case 't':
					builder.Append('\t');
					break;
				case 'u':
				{
					if (index + 4 < value.Length && int.TryParse(value.Substring(index + 1, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var code))
					{
						builder.Append(Strings.ChrW(code));
						index += 4;
					}
					break;
				}
				default:
					builder.Append(escaped);
					break;
				}
				index++;
			}
			return builder.ToString();
		}
	}

	private static FamilyBrowserRequestRecord BuildUnreadableRequestRecord(string requestPath, Exception loadError)
	{
		string createdAt = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		return new FamilyBrowserRequestRecord
		{
			RequestId = Path.GetFileNameWithoutExtension(requestPath),
			RequestKind = "GENERAL",
			Status = "Reviewing",
			CreatedAtUtc = createdAt,
			UpdatedAtUtc = createdAt,
			ItemName = Path.GetFileNameWithoutExtension(requestPath),
			CategoryName = "Request file",
			Discipline = "Unassigned",
			SuggestedAction = "Ask the administrator to inspect this request file.",
			Reason = "The request file exists, but its JSON could not be read by the request board.",
			Notes = "The file is still visible so it is not lost. Check request-read-errors.log in the request store.",
			ProgressNote = ((loadError == null) ? string.Empty : loadError.Message),
			SourcePath = requestPath
		};
	}

	private static void WriteReadIssue(string outputDir, string requestPath, Exception loadError)
	{
		if (!string.IsNullOrWhiteSpace(outputDir))
		{
			try
			{
				Directory.CreateDirectory(outputDir);
				string logPath = Path.Combine(outputDir, "request-read-errors.log");
				string line = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture) + " | " + requestPath + " | " + ((loadError == null) ? "Unknown read error" : (loadError.GetType().Name + ": " + loadError.Message));
				File.AppendAllText(logPath, line + Environment.NewLine, Utf8NoBom);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
	}

	private static DateTime ParseUtcOrMin(string value)
	{
		if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var parsed))
		{
			return parsed.ToUniversalTime();
		}
		return DateTime.MinValue;
	}

	private static string MakeSafeFileName(string value)
	{
		char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
		string normalized = new string((value ?? string.Empty).Select([SpecialName] (char ch) => (!Enumerable.Contains(invalidFileNameChars, ch)) ? ch : '_').ToArray()).Trim();
		if (normalized.Length == 0)
		{
			return "Untitled";
		}
		if (normalized.Length > 80)
		{
			normalized = normalized.Substring(0, 80);
		}
		return normalized;
	}
}
