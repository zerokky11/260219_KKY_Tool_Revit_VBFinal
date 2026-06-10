using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;

public sealed class StandardLibraryRegistryStore
{
	private StandardLibraryRegistryStore()
	{
	}

	public static string SaveActiveRegistration(string workspaceRoot, StandardLibraryRegistrationRecord registration)
	{
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Save standard RVT registration", "표준 RVT 등록 정보 저장"));
		string registryFolder = FamilyBrowserStandardPolicyStore.GetRegistryFolder(workspaceRoot);
		Directory.CreateDirectory(registryFolder);
		string text = Path.Combine(registryFolder, "active-standard-library.json");
		File.WriteAllText(text, PlainJsonReportWriter.Serialize(registration));
		return text;
	}

	public static string SaveSnapshot(string workspaceRoot, StandardLibrarySnapshot snapshot)
	{
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Save standard RVT snapshot", "표준 RVT 스냅샷 저장"));
		string outputDir = FamilyBrowserStandardPolicyStore.GetSnapshotFolder(workspaceRoot);
		Directory.CreateDirectory(outputDir);
		string safeDisplayName = MakeSafeFileName(snapshot.DisplayName ?? "StandardLibrary");
		string fileName = "standard-library-snapshot-" + safeDisplayName + "-" + DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + ".json";
		string text = Path.Combine(outputDir, fileName);
		File.WriteAllText(text, PlainJsonReportWriter.Serialize(snapshot));
		return text;
	}

	public static StandardLibrarySnapshotCacheHit TryFindReusableSnapshot(string workspaceRoot, string sourceId, string resolvedPath, string sourceFileLastWriteUtc, long sourceFileLength, string requestedSnapshotMode)
	{
		string outputDir = FamilyBrowserStandardPolicyStore.GetSnapshotFolder(workspaceRoot);
		StandardLibrarySnapshotCacheHit TryFindReusableSnapshot;
		if (!Directory.Exists(outputDir))
		{
			TryFindReusableSnapshot = null;
		}
		else if (string.IsNullOrWhiteSpace(sourceId) || string.IsNullOrWhiteSpace(resolvedPath) || string.IsNullOrWhiteSpace(sourceFileLastWriteUtc) || sourceFileLength <= 0)
		{
			TryFindReusableSnapshot = null;
		}
		else
		{
			string requestedMode = NormalizeSnapshotMode(requestedSnapshotMode);
			string normalizedPath = NormalizePathForCompare(resolvedPath);
			IEnumerable<FileInfo> snapshotFiles;
			try
			{
				snapshotFiles = from x in new DirectoryInfo(outputDir).EnumerateFiles("standard-library-snapshot-*.json", SearchOption.TopDirectoryOnly)
					orderby x.LastWriteTimeUtc descending
					select x;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				TryFindReusableSnapshot = null;
				ProjectData.ClearProjectError();
				goto IL_017e;
			}
			foreach (FileInfo fileInfo in snapshotFiles)
			{
				StandardLibrarySnapshot snapshot = null;
				try
				{
					snapshot = DataContractJsonFileStore.Load<StandardLibrarySnapshot>(fileInfo.FullName);
				}
				catch (Exception projectError2)
				{
					ProjectData.SetProjectError(projectError2);
					ProjectData.ClearProjectError();
					continue;
				}
				if (snapshot == null || snapshot.SnapshotSchemaVersion < 5 || !string.Equals(snapshot.SourceId ?? string.Empty, sourceId, StringComparison.OrdinalIgnoreCase) || !string.Equals(NormalizePathForCompare(snapshot.ResolvedPath), normalizedPath, StringComparison.OrdinalIgnoreCase) || !string.Equals(snapshot.SourceFileLastWriteUtc ?? string.Empty, sourceFileLastWriteUtc, StringComparison.Ordinal) || snapshot.SourceFileLength != sourceFileLength || !IsSnapshotModeAcceptable(requestedMode, snapshot.SnapshotMode))
				{
					continue;
				}
				TryFindReusableSnapshot = new StandardLibrarySnapshotCacheHit
				{
					Snapshot = snapshot,
					SnapshotPath = fileInfo.FullName
				};
				goto IL_017e;
			}
			TryFindReusableSnapshot = null;
		}
		goto IL_017e;
		IL_017e:
		return TryFindReusableSnapshot;
	}

	private static string MakeSafeFileName(string value)
	{
		char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
		string normalized = new string(value.Select([SpecialName] (char ch) => (!invalidFileNameChars.Contains(ch)) ? ch : '_').ToArray()).Trim();
		if (normalized.Length == 0)
		{
			return "StandardLibrary";
		}
		return normalized;
	}

	private static string NormalizeSnapshotMode(string value)
	{
		if (string.Equals(value, "Precise", StringComparison.OrdinalIgnoreCase))
		{
			return "Precise";
		}
		return "Fast";
	}

	private static bool IsSnapshotModeAcceptable(string requestedMode, string candidateMode)
	{
		string a = NormalizeSnapshotMode(requestedMode);
		string candidate = NormalizeSnapshotMode(candidateMode);
		if (string.Equals(a, "Precise", StringComparison.OrdinalIgnoreCase))
		{
			return string.Equals(candidate, "Precise", StringComparison.OrdinalIgnoreCase);
		}
		return true;
	}

	private static string NormalizePathForCompare(string value)
	{
		string NormalizePathForCompare;
		if (string.IsNullOrWhiteSpace(value))
		{
			NormalizePathForCompare = string.Empty;
		}
		else
		{
			try
			{
				NormalizePathForCompare = Path.GetFullPath(value).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				NormalizePathForCompare = value.Trim();
				ProjectData.ClearProjectError();
			}
		}
		return NormalizePathForCompare;
	}
}
