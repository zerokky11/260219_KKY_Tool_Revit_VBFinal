using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

public sealed class FamilyBrowserManagedFolderSetupResult
{
	public bool Success { get; set; }

	public string RootPath { get; set; }

	public string PolicyPath { get; set; }

	public string RequestPath { get; set; }

	public string PointerPath { get; set; }

	public string Message { get; set; }

	public FamilyBrowserManagedFolderSetupResult()
	{
		RootPath = string.Empty;
		PolicyPath = string.Empty;
		RequestPath = string.Empty;
		PointerPath = string.Empty;
		Message = string.Empty;
	}
}

public sealed class FamilyBrowserHomepageManagedFolderProbeResult
{
	public bool Available { get; set; }

	public string Source { get; set; }

	public string BootstrapUrl { get; set; }

	public string ManagedPolicyPath { get; set; }

	public string ManagedRootPath { get; set; }

	public string Issue { get; set; }

	public string CheckedAtUtc { get; set; }

	public FamilyBrowserHomepageManagedFolderProbeResult()
	{
		Source = string.Empty;
		BootstrapUrl = string.Empty;
		ManagedPolicyPath = string.Empty;
		ManagedRootPath = string.Empty;
		Issue = string.Empty;
		CheckedAtUtc = string.Empty;
	}
}

public sealed class FamilyBrowserManagedFolderMigrationAnalysis
{
	public bool CanMigrate { get; set; }

	public bool SameRoot { get; set; }

	public string SourceRoot { get; set; }

	public string DestinationRoot { get; set; }

	public int SourceFileCount { get; set; }

	public int CopyFileCount { get; set; }

	public int AlreadyPresentCount { get; set; }

	public int BlockingConflictCount { get; set; }

	public int SkippedDiagnosticConflictCount { get; set; }

	public long CopyBytes { get; set; }

	public int RebasedJsonFileCount { get; set; }

	public List<string> BlockingConflicts { get; set; }

	public List<string> SkippedDiagnosticConflicts { get; set; }

	public string Issue { get; set; }

	public FamilyBrowserManagedFolderMigrationAnalysis()
	{
		SourceRoot = string.Empty;
		DestinationRoot = string.Empty;
		BlockingConflicts = new List<string>();
		SkippedDiagnosticConflicts = new List<string>();
		Issue = string.Empty;
	}
}

public sealed class FamilyBrowserManagedFolderMigrationResult
{
	public bool Success { get; set; }

	public string SourceRoot { get; set; }

	public string DestinationRoot { get; set; }

	public int CopiedFileCount { get; set; }

	public int AlreadyPresentCount { get; set; }

	public int SkippedDiagnosticConflictCount { get; set; }

	public int RebasedJsonFileCount { get; set; }

	public long CopiedBytes { get; set; }

	public int RolledBackFileCount { get; set; }

	public int RollbackFailedFileCount { get; set; }

	public bool SourceChangedDuringMigration { get; set; }

	public string Issue { get; set; }

	public FamilyBrowserManagedFolderMigrationResult()
	{
		SourceRoot = string.Empty;
		DestinationRoot = string.Empty;
		Issue = string.Empty;
	}
}

public static class FamilyBrowserManagedFolderSetupService
{
	private const string PointerFileName = "managed-folder-override.txt";

	private const string ReadmeFileName = "KKY_FAMILY_BROWSER_MANAGED_FOLDER_README.txt";

	private static readonly string[] ManagedSubfolders = new string[]
	{
		"Config",
		"RevitVersions",
		"StandardLists",
		"Requests",
		"Logs",
		"Diagnostics",
		"OperationLogs",
		"StandardChangeCandidates",
		"ElementChangeHistory",
		"ProjectCatalogs",
		"StandardRevisionManifests"
	};

	private static readonly string[] DiagnosticSubfolders = new string[]
	{
		"Logs",
		"Diagnostics",
		"OperationLogs"
	};

	public static string GetPointerPath()
	{
		string settingsRoot = FamilyBrowserUserSettingsStore.GetSettingsRoot();
		if (string.IsNullOrWhiteSpace(settingsRoot))
		{
			return string.Empty;
		}
		return Path.Combine(settingsRoot, PointerFileName);
	}

	public static string GetConfiguredOverrideRoot()
	{
		try
		{
			string pointerPath = GetPointerPath();
			if (string.IsNullOrWhiteSpace(pointerPath) || !File.Exists(pointerPath))
			{
				return string.Empty;
			}
			return NormalizeRoot(File.ReadAllText(pointerPath, Encoding.UTF8));
		}
		catch
		{
			return string.Empty;
		}
	}

	public static bool IsOverrideConfigured()
	{
		return !string.IsNullOrWhiteSpace(GetConfiguredOverrideRoot());
	}

	public static bool TryApplyPersistedOverride(string currentUser, out string issue)
	{
		issue = string.Empty;
		string root = GetConfiguredOverrideRoot();
		if (string.IsNullOrWhiteSpace(root))
		{
			return false;
		}
		if (!IsInternalNetworkShare(root))
		{
			issue = "The locally configured management folder is not an internal network share: " + root;
			return false;
		}
		if (!Directory.Exists(root))
		{
			issue = "The locally configured management folder is not reachable: " + root;
			return false;
		}
		FamilyBrowserMachineConfig previousMachineConfig = FamilyBrowserMachineConfigStore.Load();
		using (IDisposable managementContextLease = FamilyBrowserManagementContextLock.Acquire(TimeSpan.FromSeconds(30.0)))
		{
			try
			{
				string policyPath = BuildPolicyPath(root);
				FamilyBrowserMachineConfigStore.SetManagedPolicyPath(policyPath, currentUser);
				FamilyBrowserDeploymentBootstrapService.DisableBootstrap(currentUser);
				return true;
			}
			catch (Exception ex)
			{
				string rollbackIssue = string.Empty;
				try
				{
					RestoreMachineConfiguration(previousMachineConfig, currentUser);
				}
				catch (Exception rollbackError)
				{
					rollbackIssue = " Previous management configuration rollback failed: " + rollbackError.Message;
				}
				issue = ex.Message + rollbackIssue;
				return false;
			}
		}
	}

	public static bool HasUsableManagedFolder(out string issue)
	{
		issue = string.Empty;
		string policyPath = FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath();
		if (string.IsNullOrWhiteSpace(policyPath))
		{
			issue = "No managed policy path is configured.";
			return false;
		}
		try
		{
			string policyFolder = Path.GetDirectoryName(policyPath);
			if (string.IsNullOrWhiteSpace(policyFolder) || !Directory.Exists(policyFolder))
			{
				issue = "The configured management folder is not reachable: " + ResolveRootFromPolicyPath(policyPath);
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			issue = ex.Message;
			return false;
		}
	}

	public static FamilyBrowserManagedFolderSetupResult Configure(string workspaceRoot, string selectedRoot, string currentUser)
	{
		string root = NormalizeRoot(selectedRoot);
		if (string.IsNullOrWhiteSpace(root))
		{
			throw new InvalidOperationException("Select an existing internal network shared folder.");
		}
		if (!IsInternalNetworkShare(root))
		{
			throw new InvalidOperationException("The management folder must be a UNC path or a mapped network drive. Local disks and user folders cannot be used.");
		}
		if (!Directory.Exists(root))
		{
			throw new DirectoryNotFoundException("The selected network shared folder cannot be reached: " + root);
		}

		FamilyBrowserMachineConfig previousMachineConfig = FamilyBrowserMachineConfigStore.Load();
		string previousOverrideRoot = GetConfiguredOverrideRoot();
		using (IDisposable managementContextLease = FamilyBrowserManagementContextLock.Acquire(TimeSpan.FromSeconds(30.0)))
		{
			CreateAndProbeManagedFolders(root);
			string policyPath = BuildPolicyPath(root);
			string requestPath = Path.Combine(root, "Requests");
			WriteManagedFolderReadme(root, currentUser);
			try
			{
				FamilyBrowserMachineConfigStore.SetManagedPolicyPath(policyPath, currentUser);
				FamilyBrowserDeploymentBootstrapService.DisableBootstrap(currentUser);
				FamilyBrowserStandardPolicy policy = FamilyBrowserStandardPolicyStore.LoadOrCreate(workspaceRoot, currentUser);
				FamilyBrowserStandardPolicyStore.SetRequestStore(workspaceRoot, policy, "NetworkShare", requestPath, string.Empty, currentUser);
				SaveOverrideRoot(root);
				StandardRvtChangeCandidateService.NotifyPolicyChanged();
				FamilyBrowserNativeCommandGuardService.NotifyPolicyChanged(policy);
			}
			catch (Exception ex)
			{
				List<string> rollbackIssues = new List<string>();
				try
				{
					RestoreMachineConfiguration(previousMachineConfig, currentUser);
				}
				catch (Exception rollbackError)
				{
					rollbackIssues.Add("management configuration: " + rollbackError.Message);
				}
				try
				{
					RestoreOverrideRoot(previousOverrideRoot);
				}
				catch (Exception rollbackError)
				{
					rollbackIssues.Add("TEST pointer: " + rollbackError.Message);
				}
				string rollbackSuffix = rollbackIssues.Count == 0
					? " Previous management configuration was restored."
					: " Rollback requires administrator review (" + string.Join("; ", rollbackIssues.ToArray()) + ").";
				throw new InvalidOperationException(ex.Message + rollbackSuffix, ex);
			}

			return new FamilyBrowserManagedFolderSetupResult
			{
				Success = true,
				RootPath = root,
				PolicyPath = policyPath,
				RequestPath = requestPath,
				PointerPath = GetPointerPath(),
				Message = "The internal network management folder is ready."
			};
		}
	}

	public static FamilyBrowserManagedFolderMigrationAnalysis AnalyzeMigration(string sourceRoot, string destinationPolicyPath)
	{
		FamilyBrowserManagedFolderMigrationAnalysis analysis = new FamilyBrowserManagedFolderMigrationAnalysis();
		try
		{
			string source = NormalizeRoot(sourceRoot);
			string destination = NormalizeRoot(ResolveRootFromPolicyPath(destinationPolicyPath));
			analysis.SourceRoot = source;
			analysis.DestinationRoot = destination;
			ValidateMigrationRoots(source, destination);
			if (SamePath(source, destination))
			{
				analysis.SameRoot = true;
				analysis.CanMigrate = true;
				return analysis;
			}
			foreach (string sourceFile in EnumerateManagedFiles(source))
			{
				string relativePath = GetRelativeManagedPath(source, sourceFile);
				if (string.IsNullOrWhiteSpace(relativePath))
				{
					continue;
				}
				analysis.SourceFileCount++;
				int replacementCount;
				byte[] expectedBytes = ReadMigratedManagedFileBytes(sourceFile, relativePath, source, destination, out replacementCount);
				if (replacementCount > 0)
				{
					analysis.RebasedJsonFileCount++;
				}
				string destinationFile = Path.Combine(destination, relativePath);
				if (!File.Exists(destinationFile))
				{
					analysis.CopyFileCount++;
					analysis.CopyBytes += expectedBytes.LongLength;
					continue;
				}
				if (FileMatchesBytes(destinationFile, expectedBytes))
				{
					analysis.AlreadyPresentCount++;
					continue;
				}
				if (IsDiagnosticRelativePath(relativePath))
				{
					analysis.SkippedDiagnosticConflictCount++;
					AddLimited(analysis.SkippedDiagnosticConflicts, relativePath);
				}
				else
				{
					analysis.BlockingConflictCount++;
					AddLimited(analysis.BlockingConflicts, relativePath);
				}
			}
			analysis.CanMigrate = analysis.BlockingConflictCount == 0;
			if (!analysis.CanMigrate)
			{
				analysis.Issue = "The homepage management folder already contains different managed data. Nothing was copied.";
			}
		}
		catch (Exception ex)
		{
			analysis.CanMigrate = false;
			analysis.Issue = ex.Message;
		}
		return analysis;
	}

	public static FamilyBrowserManagedFolderMigrationResult MigrateToHomepage(string sourceRoot, string destinationPolicyPath, string currentUser)
	{
		FamilyBrowserManagedFolderMigrationResult result = new FamilyBrowserManagedFolderMigrationResult();
		FamilyBrowserManagedFolderMigrationAnalysis analysis = AnalyzeMigration(sourceRoot, destinationPolicyPath);
		result.SourceRoot = analysis.SourceRoot;
		result.DestinationRoot = analysis.DestinationRoot;
		result.AlreadyPresentCount = analysis.AlreadyPresentCount;
		result.SkippedDiagnosticConflictCount = analysis.SkippedDiagnosticConflictCount;
		if (!analysis.CanMigrate)
		{
			result.Issue = string.IsNullOrWhiteSpace(analysis.Issue) ? "Managed-folder migration preflight failed." : analysis.Issue;
			return result;
		}
		if (analysis.SameRoot)
		{
			result.Success = true;
			return result;
		}
		List<Tuple<string, byte[]>> copiedFiles = new List<Tuple<string, byte[]>>();
		try
		{
			string sourceFingerprintBefore = ComputeManagedSourceFingerprint(analysis.SourceRoot);
			foreach (string sourceFile in EnumerateManagedFiles(analysis.SourceRoot))
			{
				string relativePath = GetRelativeManagedPath(analysis.SourceRoot, sourceFile);
				if (string.IsNullOrWhiteSpace(relativePath))
				{
					continue;
				}
				int replacementCount;
				byte[] expectedBytes = ReadMigratedManagedFileBytes(sourceFile, relativePath, analysis.SourceRoot, analysis.DestinationRoot, out replacementCount);
				string destinationFile = Path.Combine(analysis.DestinationRoot, relativePath);
				if (File.Exists(destinationFile))
				{
					if (FileMatchesBytes(destinationFile, expectedBytes))
					{
						continue;
					}
					if (IsDiagnosticRelativePath(relativePath))
					{
						continue;
					}
					throw new IOException("A managed file changed after preflight. Nothing was overwritten: " + relativePath);
				}
				string destinationFolder = Path.GetDirectoryName(destinationFile);
				if (string.IsNullOrWhiteSpace(destinationFolder))
				{
					throw new IOException("The destination folder could not be resolved for: " + relativePath);
				}
				Directory.CreateDirectory(destinationFolder);
				string temporaryPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(destinationFile);
				try
				{
					File.WriteAllBytes(temporaryPath, expectedBytes);
					if (!FileMatchesBytes(temporaryPath, expectedBytes))
					{
						throw new IOException("Copied data verification failed: " + relativePath);
					}
					File.Move(temporaryPath, destinationFile);
					copiedFiles.Add(Tuple.Create(destinationFile, expectedBytes));
					try
					{
						File.SetLastWriteTimeUtc(destinationFile, File.GetLastWriteTimeUtc(sourceFile));
					}
					catch
					{
					}
				}
				finally
				{
					if (File.Exists(temporaryPath))
					{
						File.Delete(temporaryPath);
					}
				}
				result.CopiedFileCount++;
				result.CopiedBytes += expectedBytes.LongLength;
				if (replacementCount > 0)
				{
					result.RebasedJsonFileCount++;
				}
			}
			string sourceFingerprintAfter = ComputeManagedSourceFingerprint(analysis.SourceRoot);
			if (!string.Equals(sourceFingerprintBefore, sourceFingerprintAfter, StringComparison.OrdinalIgnoreCase))
			{
				result.SourceChangedDuringMigration = true;
				throw new IOException("The TEST management folder changed while it was being copied. The partial destination copy will be rolled back; retry after other users stop writing managed data.");
			}
			result.Success = true;
		}
		catch (Exception ex)
		{
			result.Success = false;
			for (int i = copiedFiles.Count - 1; i >= 0; i--)
			{
				Tuple<string, byte[]> copied = copiedFiles[i];
				try
				{
					if (File.Exists(copied.Item1) && FileMatchesBytes(copied.Item1, copied.Item2))
					{
						File.Delete(copied.Item1);
						result.RolledBackFileCount++;
					}
					else if (File.Exists(copied.Item1))
					{
						result.RollbackFailedFileCount++;
					}
				}
				catch
				{
					result.RollbackFailedFileCount++;
				}
			}
			result.Issue = ex.Message;
			if (result.RollbackFailedFileCount > 0)
			{
				result.Issue += " " + result.RollbackFailedFileCount.ToString(CultureInfo.InvariantCulture) + " copied file(s) could not be rolled back because they changed or remained locked; administrator review is required.";
			}
		}
		return result;
	}

	public static void ClearOverrideRoot()
	{
		string issue;
		TryClearOverrideRoot(out issue);
	}

	public static bool TryClearOverrideRoot(out string issue)
	{
		issue = string.Empty;
		try
		{
			string pointerPath = GetPointerPath();
			if (!string.IsNullOrWhiteSpace(pointerPath) && File.Exists(pointerPath))
			{
				File.Delete(pointerPath);
			}
			if (!string.IsNullOrWhiteSpace(pointerPath) && File.Exists(pointerPath))
			{
				issue = "The local TEST management-folder pointer could not be removed: " + pointerPath;
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			issue = ex.Message;
			return false;
		}
	}

	public static bool IsInternalNetworkShare(string path)
	{
		string expanded = Environment.ExpandEnvironmentVariables((path ?? string.Empty).Trim());
		if (string.IsNullOrWhiteSpace(expanded))
		{
			return false;
		}
		if (expanded.StartsWith("\\\\", StringComparison.Ordinal))
		{
			return true;
		}
		try
		{
			string root = Path.GetPathRoot(expanded);
			if (string.IsNullOrWhiteSpace(root))
			{
				return false;
			}
			DriveInfo drive = new DriveInfo(root);
			return drive.DriveType == DriveType.Network;
		}
		catch
		{
			return false;
		}
	}

	public static string ResolveRootFromPolicyPath(string policyPath)
	{
		if (string.IsNullOrWhiteSpace(policyPath))
		{
			return string.Empty;
		}
		try
		{
			string folder = Path.GetDirectoryName(policyPath);
			DirectoryInfo info = string.IsNullOrWhiteSpace(folder) ? null : new DirectoryInfo(folder);
			if (info != null && string.Equals(info.Name, "Config", StringComparison.OrdinalIgnoreCase) && info.Parent != null)
			{
				return info.Parent.FullName;
			}
			return folder ?? string.Empty;
		}
		catch
		{
			return string.Empty;
		}
	}

	private static string BuildPolicyPath(string root)
	{
		return Path.Combine(root, "Config", "standard-policy.json");
	}

	private static string NormalizeRoot(string value)
	{
		string expanded = Environment.ExpandEnvironmentVariables((value ?? string.Empty).Trim().Trim('"'));
		if (string.IsNullOrWhiteSpace(expanded))
		{
			return string.Empty;
		}
		try
		{
			return Path.GetFullPath(expanded).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
		}
		catch
		{
			return expanded.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
		}
	}

	private static void ValidateMigrationRoots(string source, string destination)
	{
		if (string.IsNullOrWhiteSpace(source) || !Directory.Exists(source))
		{
			throw new DirectoryNotFoundException("The TEST management folder is not reachable: " + source);
		}
		if (string.IsNullOrWhiteSpace(destination) || !Directory.Exists(destination))
		{
			throw new DirectoryNotFoundException("The homepage management folder is not reachable: " + destination);
		}
		if (!SamePath(source, destination) && (IsSameOrChildPath(source, destination) || IsSameOrChildPath(destination, source)))
		{
			throw new InvalidOperationException("The TEST and homepage management folders cannot be nested inside each other.");
		}
	}

	private static bool SamePath(string left, string right)
	{
		return string.Equals(NormalizeRoot(left), NormalizeRoot(right), StringComparison.OrdinalIgnoreCase);
	}

	private static bool IsSameOrChildPath(string candidate, string parent)
	{
		string child = NormalizeRoot(candidate);
		string root = NormalizeRoot(parent);
		if (string.IsNullOrWhiteSpace(child) || string.IsNullOrWhiteSpace(root))
		{
			return false;
		}
		return string.Equals(child, root, StringComparison.OrdinalIgnoreCase) || child.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) || child.StartsWith(root + Path.AltDirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
	}

	private static IEnumerable<string> EnumerateManagedFiles(string root)
	{
		List<string> files = new List<string>();
		foreach (string folderName in ManagedSubfolders)
		{
			string folder = Path.Combine(root, folderName);
			if (!Directory.Exists(folder))
			{
				continue;
			}
			Stack<string> pending = new Stack<string>();
			pending.Push(folder);
			while (pending.Count > 0)
			{
				string current = pending.Pop();
				foreach (string file in Directory.GetFiles(current, "*", SearchOption.TopDirectoryOnly))
				{
					if (!ShouldSkipMigrationFile(file))
					{
						files.Add(file);
					}
				}
				foreach (string child in Directory.GetDirectories(current, "*", SearchOption.TopDirectoryOnly))
				{
					FileAttributes attributes = File.GetAttributes(child);
					if ((attributes & FileAttributes.ReparsePoint) == 0)
					{
						pending.Push(child);
					}
				}
			}
		}
		files.Sort(StringComparer.OrdinalIgnoreCase);
		return files;
	}

	private static bool ShouldSkipMigrationFile(string filePath)
	{
		string name = Path.GetFileName(filePath) ?? string.Empty;
		return name.StartsWith(".kky-family-browser-write-test-", StringComparison.OrdinalIgnoreCase)
			|| name.StartsWith(".kky-w-", StringComparison.OrdinalIgnoreCase)
			|| name.StartsWith(".kky-t-", StringComparison.OrdinalIgnoreCase)
			|| name.StartsWith(".kky-b-", StringComparison.OrdinalIgnoreCase)
			|| name.StartsWith(".kky-r-", StringComparison.OrdinalIgnoreCase)
			|| name.EndsWith(".kky-lock", StringComparison.OrdinalIgnoreCase)
			|| name.IndexOf(".tmp-", StringComparison.OrdinalIgnoreCase) >= 0
			|| name.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase)
			|| name.IndexOf(".kky-migration-", StringComparison.OrdinalIgnoreCase) >= 0;
	}

	private static string GetRelativeManagedPath(string root, string filePath)
	{
		string normalizedRoot = NormalizeRoot(root);
		string fullPath = Path.GetFullPath(filePath);
		if (!fullPath.StartsWith(normalizedRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
		{
			return string.Empty;
		}
		return fullPath.Substring(normalizedRoot.Length + 1);
	}

	private static bool IsDiagnosticRelativePath(string relativePath)
	{
		string first = (relativePath ?? string.Empty).Split(new char[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries)[0];
		foreach (string folder in DiagnosticSubfolders)
		{
			if (string.Equals(first, folder, StringComparison.OrdinalIgnoreCase))
			{
				return true;
			}
		}
		return false;
	}

	private static void AddLimited(List<string> values, string value)
	{
		if (values != null && values.Count < 25)
		{
			values.Add(value ?? string.Empty);
		}
	}

	private static byte[] ReadMigratedManagedFileBytes(string sourceFile, string relativePath, string sourceRoot, string destinationRoot, out int replacementCount)
	{
		// Element-change commits are immutable evidence whose checksum covers the serialized payload.
		// Rewriting path-looking strings during a managed-root move would invalidate that evidence.
		if (IsElementChangeHistoryRelativePath(relativePath))
		{
			replacementCount = 0;
			return File.ReadAllBytes(sourceFile);
		}
		return ReadMigratedFileBytes(sourceFile, sourceRoot, destinationRoot, out replacementCount);
	}

	private static bool IsElementChangeHistoryRelativePath(string relativePath)
	{
		string[] parts = (relativePath ?? string.Empty).Split(new char[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);
		return parts.Length > 0 && string.Equals(parts[0], "ElementChangeHistory", StringComparison.OrdinalIgnoreCase);
	}

	private static byte[] ReadMigratedFileBytes(string sourceFile, string sourceRoot, string destinationRoot, out int replacementCount)
	{
		replacementCount = 0;
		byte[] original = File.ReadAllBytes(sourceFile);
		if (!string.Equals(Path.GetExtension(sourceFile), ".json", StringComparison.OrdinalIgnoreCase) || original.Length == 0)
		{
			return original;
		}
		bool hasBom = original.Length >= 3 && original[0] == 0xEF && original[1] == 0xBB && original[2] == 0xBF;
		string json = Encoding.UTF8.GetString(original, hasBom ? 3 : 0, original.Length - (hasBom ? 3 : 0));
		string rebased = RebaseJsonStringValues(json, sourceRoot, destinationRoot, out replacementCount);
		if (replacementCount == 0)
		{
			return original;
		}
		byte[] body = new UTF8Encoding(false).GetBytes(rebased);
		if (!hasBom)
		{
			return body;
		}
		byte[] withBom = new byte[body.Length + 3];
		withBom[0] = 0xEF;
		withBom[1] = 0xBB;
		withBom[2] = 0xBF;
		Buffer.BlockCopy(body, 0, withBom, 3, body.Length);
		return withBom;
	}

	private static string RebaseJsonStringValues(string json, string sourceRoot, string destinationRoot, out int replacementCount)
	{
		replacementCount = 0;
		if (string.IsNullOrEmpty(json))
		{
			return json ?? string.Empty;
		}
		StringBuilder output = new StringBuilder(json.Length + 128);
		int index = 0;
		while (index < json.Length)
		{
			if (json[index] != '"')
			{
				output.Append(json[index++]);
				continue;
			}
			int literalStart = index;
			index++;
			StringBuilder decoded = new StringBuilder();
			bool closed = false;
			while (index < json.Length)
			{
				char current = json[index++];
				if (current == '"')
				{
					closed = true;
					break;
				}
				if (current != '\\')
				{
					decoded.Append(current);
					continue;
				}
				if (index >= json.Length)
				{
					break;
				}
				char escape = json[index++];
				switch (escape)
				{
				case '"': decoded.Append('"'); break;
				case '\\': decoded.Append('\\'); break;
				case '/': decoded.Append('/'); break;
				case 'b': decoded.Append('\b'); break;
				case 'f': decoded.Append('\f'); break;
				case 'n': decoded.Append('\n'); break;
				case 'r': decoded.Append('\r'); break;
				case 't': decoded.Append('\t'); break;
				case 'u':
					if (index + 4 > json.Length)
					{
						index = json.Length;
						break;
					}
					int code;
					if (!int.TryParse(json.Substring(index, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out code))
					{
						index = json.Length;
						break;
					}
					decoded.Append((char)code);
					index += 4;
					break;
				default:
					decoded.Append(escape);
					break;
				}
			}
			if (!closed)
			{
				return json;
			}
			string rebased;
			if (TryRebaseManagedPath(decoded.ToString(), sourceRoot, destinationRoot, out rebased))
			{
				AppendJsonString(output, rebased);
				replacementCount++;
			}
			else
			{
				output.Append(json, literalStart, index - literalStart);
			}
		}
		return output.ToString();
	}

	private static bool TryRebaseManagedPath(string value, string sourceRoot, string destinationRoot, out string rebased)
	{
		rebased = value ?? string.Empty;
		if (string.IsNullOrWhiteSpace(value))
		{
			return false;
		}
		string source = NormalizeRoot(sourceRoot).Replace('/', '\\');
		string comparable = value.Replace('/', '\\');
		if (!comparable.StartsWith(source, StringComparison.OrdinalIgnoreCase))
		{
			return false;
		}
		if (comparable.Length > source.Length && comparable[source.Length] != '\\')
		{
			return false;
		}
		string suffix = comparable.Substring(source.Length);
		rebased = NormalizeRoot(destinationRoot) + suffix;
		return !string.Equals(value, rebased, StringComparison.Ordinal);
	}

	private static void AppendJsonString(StringBuilder output, string value)
	{
		output.Append('"');
		foreach (char current in value ?? string.Empty)
		{
			switch (current)
			{
			case '"': output.Append("\\\""); break;
			case '\\': output.Append("\\\\"); break;
			case '\b': output.Append("\\b"); break;
			case '\f': output.Append("\\f"); break;
			case '\n': output.Append("\\n"); break;
			case '\r': output.Append("\\r"); break;
			case '\t': output.Append("\\t"); break;
			default:
				if (current < ' ')
				{
					output.Append("\\u");
					output.Append(((int)current).ToString("x4", CultureInfo.InvariantCulture));
				}
				else
				{
					output.Append(current);
				}
				break;
			}
		}
		output.Append('"');
	}

	private static bool FileMatchesBytes(string filePath, byte[] expectedBytes)
	{
		FileInfo file = new FileInfo(filePath);
		if (!file.Exists || file.Length != expectedBytes.LongLength)
		{
			return false;
		}
		using (SHA256 sha = SHA256.Create())
		{
			byte[] expectedHash = sha.ComputeHash(expectedBytes);
			using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
			{
				byte[] actualHash = sha.ComputeHash(stream);
				if (actualHash.Length != expectedHash.Length)
				{
					return false;
				}
				for (int i = 0; i < actualHash.Length; i++)
				{
					if (actualHash[i] != expectedHash[i])
					{
						return false;
					}
				}
				return true;
			}
		}
	}

	private static string ComputeManagedSourceFingerprint(string root)
	{
		StringBuilder manifest = new StringBuilder();
		foreach (string filePath in EnumerateManagedFiles(root))
		{
			string relativePath = GetRelativeManagedPath(root, filePath);
			if (string.IsNullOrWhiteSpace(relativePath))
			{
				continue;
			}
			byte[] fileHash;
			long fileLength;
			using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
			using (SHA256 fileSha = SHA256.Create())
			{
				fileLength = stream.Length;
				fileHash = fileSha.ComputeHash(stream);
			}
			manifest.Append(relativePath.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar).ToUpperInvariant())
				.Append('|')
				.Append(fileLength.ToString(CultureInfo.InvariantCulture))
				.Append('|')
				.Append(BitConverter.ToString(fileHash).Replace("-", string.Empty))
				.Append('\n');
		}
		using (SHA256 manifestSha = SHA256.Create())
		{
			return BitConverter.ToString(manifestSha.ComputeHash(Encoding.UTF8.GetBytes(manifest.ToString()))).Replace("-", string.Empty);
		}
	}

	private static void CreateAndProbeManagedFolders(string root)
	{
		foreach (string folderName in ManagedSubfolders)
		{
			Directory.CreateDirectory(Path.Combine(root, folderName));
		}
		string probePath = Path.Combine(root, "Config", ".kky-w-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".tmp");
		try
		{
			File.WriteAllText(probePath, DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture), new UTF8Encoding(false));
		}
		finally
		{
			try
			{
				if (File.Exists(probePath))
				{
					File.Delete(probePath);
				}
			}
			catch
			{
			}
		}
	}

	private static void SaveOverrideRoot(string root)
	{
		string settingsRoot = FamilyBrowserUserSettingsStore.GetSettingsRoot(createFolder: true);
		if (string.IsNullOrWhiteSpace(settingsRoot))
		{
			throw new InvalidOperationException("The local Family Browser settings folder is unavailable.");
		}
		string pointerPath = Path.Combine(settingsRoot, PointerFileName);
		string temporaryPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(pointerPath);
		File.WriteAllText(temporaryPath, root, new UTF8Encoding(false));
		try
		{
			FamilyBrowserAtomicFileService.Promote(temporaryPath, pointerPath);
		}
		finally
		{
			if (File.Exists(temporaryPath))
			{
				File.Delete(temporaryPath);
			}
		}
	}

	private static void RestoreMachineConfiguration(FamilyBrowserMachineConfig previous, string currentUser)
	{
		if (previous != null && previous.UseManagedPolicy && !string.IsNullOrWhiteSpace(previous.ManagedPolicyPath))
		{
			FamilyBrowserMachineConfigStore.SetManagedPolicyPath(previous.ManagedPolicyPath, currentUser);
		}
		else
		{
			FamilyBrowserMachineConfigStore.ClearManagedPolicyPath(currentUser);
		}
		FamilyBrowserDeploymentBootstrapService.SetBootstrapUrl(previous == null ? string.Empty : previous.DeploymentBootstrapUrl, currentUser);
		FamilyBrowserDeploymentBootstrapService.ClearCache();
		StandardRvtChangeCandidateService.NotifyPolicyChanged();
		FamilyBrowserNativeCommandGuardService.NotifyPolicyChanged();
	}

	private static void RestoreOverrideRoot(string previousRoot)
	{
		if (!string.IsNullOrWhiteSpace(previousRoot))
		{
			SaveOverrideRoot(previousRoot);
			return;
		}
		string issue;
		if (!TryClearOverrideRoot(out issue))
		{
			throw new IOException(string.IsNullOrWhiteSpace(issue) ? "The previous TEST pointer state could not be restored." : issue);
		}
	}

	private static void WriteManagedFolderReadme(string root, string currentUser)
	{
		List<string> lines = new List<string>
		{
			"KKY Family Browser managed folder",
			"",
			"IMPORTANT / 중요",
			"Files and folders created here are managed by KKY Family Browser.",
			"Do not manually edit, move, rename, or delete generated files and folders.",
			"이 위치에 KKY Family Browser가 생성한 파일과 폴더를 수동으로 수정, 이동, 이름 변경 또는 삭제하지 마세요.",
			"",
			"Configured by: " + (currentUser ?? string.Empty),
			"Configured at (UTC): " + DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
		};
		File.WriteAllLines(Path.Combine(root, ReadmeFileName), lines.ToArray(), new UTF8Encoding(false));
	}
}
