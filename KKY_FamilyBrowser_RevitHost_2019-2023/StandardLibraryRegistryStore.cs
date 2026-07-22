using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;

public sealed class StandardLibraryPublicationResult
{
	public string SnapshotPath { get; set; }

	public string RegistrationPath { get; set; }

	public StandardLibraryPublicationResult()
	{
		SnapshotPath = string.Empty;
		RegistrationPath = string.Empty;
	}
}

public sealed class StandardLibraryRegistryStore
{
	private static readonly object SyncRoot = new object();

	private StandardLibraryRegistryStore()
	{
	}

	public static string SaveActiveRegistration(string workspaceRoot, StandardLibraryRegistrationRecord registration)
	{
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Save standard RVT registration", "표준 RVT 등록 정보 저장"));
		using (FileStream publicationLock = AcquirePublicationLock(workspaceRoot))
		{
			return SaveActiveRegistrationCore(workspaceRoot, registration);
		}
	}

	public static string SaveSnapshot(string workspaceRoot, StandardLibrarySnapshot snapshot)
	{
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Save standard RVT snapshot", "표준 RVT 스냅샷 저장"));
		using (FileStream publicationLock = AcquirePublicationLock(workspaceRoot))
		{
			return SaveSnapshotCore(workspaceRoot, snapshot, true);
		}
	}

	public static StandardLibraryPublicationResult PublishSnapshotAndActiveRegistration(string workspaceRoot, StandardLibrarySnapshot snapshot, StandardLibraryRegistrationRecord registration)
	{
		if (snapshot == null)
		{
			throw new ArgumentNullException("snapshot");
		}
		if (registration == null)
		{
			throw new ArgumentNullException("registration");
		}
		FamilyBrowserStandardPolicyStore.RequireManagedDataRootForWrite(workspaceRoot, FamilyBrowserLanguageService.Text("Publish standard RVT snapshot", "표준 RVT 스냅샷 발행"));
		using (FileStream publicationLock = AcquirePublicationLock(workspaceRoot))
		{
			string snapshotPath = SaveSnapshotCore(workspaceRoot, snapshot, false);
			registration.LastSnapshotPath = snapshotPath;
			registration.LastSnapshotAtUtc = snapshot.CapturedAtUtc ?? registration.LastSnapshotAtUtc ?? string.Empty;
			FamilyBrowserDataLoader.PublishStandardArtifacts(workspaceRoot, snapshotPath, snapshot);
			string registrationPath = SaveActiveRegistrationCore(workspaceRoot, registration);
			return new StandardLibraryPublicationResult
			{
				SnapshotPath = snapshotPath,
				RegistrationPath = registrationPath
			};
		}
	}

	private static string SaveActiveRegistrationCore(string workspaceRoot, StandardLibraryRegistrationRecord registration)
	{
		string registryFolder = FamilyBrowserStandardPolicyStore.GetRegistryFolder(workspaceRoot);
		Directory.CreateDirectory(registryFolder);
		string path = Path.Combine(registryFolder, "active-standard-library.json");
		WriteTextAtomic(path, PlainJsonReportWriter.Serialize(registration));
		FamilyBrowserStandardRevisionService.RecordBaseline(workspaceRoot, registration, registration == null ? string.Empty : registration.RegisteredBy);
		return path;
	}

	private static string SaveSnapshotCore(string workspaceRoot, StandardLibrarySnapshot snapshot, bool publishArtifacts)
	{
		string outputDir = FamilyBrowserStandardPolicyStore.GetSnapshotFolder(workspaceRoot);
		Directory.CreateDirectory(outputDir);
		string safeDisplayName = MakeSafeFileName(snapshot.DisplayName ?? "StandardLibrary");
		string fileName = "standard-library-snapshot-" + safeDisplayName + "-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-fffffff", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".json";
		string text = Path.Combine(outputDir, fileName);
		WriteTextAtomic(text, PlainJsonReportWriter.Serialize(snapshot));
		try
		{
			FamilyBrowserNestedOnlyPlacementCatalogStore.SaveForSnapshot(text, snapshot);
		}
		catch (Exception ex)
		{
			FamilyBrowserErrorHelp.WriteLog(workspaceRoot, "Nested-only placement catalog save failed", ex, "SnapshotPath=" + text);
		}
		if (publishArtifacts)
		{
			FamilyBrowserDataLoader.PublishStandardArtifacts(workspaceRoot, text, snapshot);
		}
		return text;
	}

	private static FileStream AcquirePublicationLock(string workspaceRoot)
	{
		string registryFolder = FamilyBrowserStandardPolicyStore.GetRegistryFolder(workspaceRoot);
		Directory.CreateDirectory(registryFolder);
		string lockPath = Path.Combine(registryFolder, ".standard-library-publication.lock");
		DateTime deadlineUtc = DateTime.UtcNow.AddMinutes(5.0);
		while (true)
		{
			try
			{
				return new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.WriteThrough);
			}
			catch (IOException)
			{
				if (DateTime.UtcNow >= deadlineUtc)
				{
					throw new IOException("Timed out waiting for the Standard RVT publication lock: " + lockPath);
				}
				System.Threading.Thread.Sleep(250);
			}
		}
	}

	private static void WriteTextAtomic(string path, string content)
	{
		lock (SyncRoot)
		{
			Directory.CreateDirectory(Path.GetDirectoryName(path));
			string temporary = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
			try
			{
				using (FileStream stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
				using (StreamWriter writer = new StreamWriter(stream))
				{
					writer.Write(content ?? string.Empty);
					writer.Flush();
					stream.Flush(true);
				}
				FamilyBrowserAtomicFileService.Promote(temporary, path);
			}
			finally
			{
				if (File.Exists(temporary))
				{
					try { File.Delete(temporary); } catch { }
				}
			}
		}
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
				if (snapshot == null || snapshot.SnapshotSchemaVersion < 9 || !string.Equals(snapshot.SourceId ?? string.Empty, sourceId, StringComparison.OrdinalIgnoreCase) || !string.Equals(NormalizePathForCompare(snapshot.ResolvedPath), normalizedPath, StringComparison.OrdinalIgnoreCase) || !string.Equals(snapshot.SourceFileLastWriteUtc ?? string.Empty, sourceFileLastWriteUtc, StringComparison.Ordinal) || snapshot.SourceFileLength != sourceFileLength || !IsSnapshotModeAcceptable(requestedMode, snapshot.SnapshotMode))
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
