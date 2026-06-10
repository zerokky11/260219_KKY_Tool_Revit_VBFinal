using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FingerprintDebugSignatureStore
{
	private FingerprintDebugSignatureStore()
	{
	}

	public static Dictionary<string, List<LoadableSignatureDebugRecord>> BuildLoadableSignatureIndex(IEnumerable<string> signaturePaths, IEnumerable<string> signatureRootFolders, int maxRunFoldersPerRoot = 6)
	{
		Dictionary<string, List<LoadableSignatureDebugRecord>> index = new Dictionary<string, List<LoadableSignatureDebugRecord>>(StringComparer.Ordinal);
		HashSet<string> seenPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		if (signaturePaths != null)
		{
			foreach (string signaturePath in signaturePaths)
			{
				AddLoadableSignatureFile(index, seenPaths, signaturePath);
			}
		}
		if (signatureRootFolders != null)
		{
			foreach (string rootFolder in signatureRootFolders)
			{
				AddLoadableSignatureRoot(index, seenPaths, rootFolder, maxRunFoldersPerRoot);
			}
		}
		return index;
	}

	public static LoadableSignatureDebugRecord FindBestLoadableSignatureRecord(Dictionary<string, List<LoadableSignatureDebugRecord>> index, string categoryName, string familyName)
	{
		if (index == null || string.IsNullOrWhiteSpace(familyName))
		{
			return null;
		}
		List<LoadableSignatureDebugRecord> records = null;
		string exactKey = "cf|" + FoldSignatureToken(categoryName) + "|" + FoldSignatureToken(familyName);
		if (index.TryGetValue(exactKey, out records) && records != null && records.Count > 0)
		{
			return SelectBestLoadableSignatureRecord(records);
		}
		string familyKey = "f|" + FoldSignatureToken(familyName);
		if (index.TryGetValue(familyKey, out records) && records != null && records.Count > 0)
		{
			return SelectBestLoadableSignatureRecord(records);
		}
		return null;
	}

	public static LoadableSignatureDebugRecord ReadLoadableSignatureRecord(string signaturePath)
	{
		LoadableSignatureDebugRecord ReadLoadableSignatureRecord;
		try
		{
			string expandedPath = Environment.ExpandEnvironmentVariables((signaturePath ?? string.Empty).Trim());
			if (string.IsNullOrWhiteSpace(expandedPath) || !File.Exists(expandedPath))
			{
				ReadLoadableSignatureRecord = null;
			}
			else
			{
				LoadableSignatureDebugRecord record = new LoadableSignatureDebugRecord
				{
					Path = expandedPath
				};
				try
				{
					record.LastWriteUtcTicks = new FileInfo(expandedPath).LastWriteTimeUtc.Ticks;
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				foreach (string line in File.ReadLines(expandedPath, Encoding.UTF8))
				{
					if (string.IsNullOrWhiteSpace(line))
					{
						continue;
					}
					if (line.StartsWith("-----", StringComparison.Ordinal))
					{
						break;
					}
					int separator = line.IndexOf('=');
					if (separator > 0)
					{
						string text = line.Substring(0, separator).Trim();
						string value = line.Substring(checked(separator + 1)).Trim();
						switch (text.ToLowerInvariant())
						{
						case "category":
							record.CategoryName = value;
							break;
						case "family":
							record.FamilyName = value;
							break;
						case "signature-mode":
							record.Mode = value;
							break;
						case "content-fingerprint":
							record.Fingerprint = value;
							break;
						case "error-message":
							record.ErrorMessage = value;
							break;
						}
					}
				}
				ReadLoadableSignatureRecord = ((!string.IsNullOrWhiteSpace(record.FamilyName)) ? record : null);
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ReadLoadableSignatureRecord = null;
			ProjectData.ClearProjectError();
		}
		return ReadLoadableSignatureRecord;
	}

	public static string CreateStandardRunFolder(string workspaceRoot, string sourceId, string displayName, string capturedAtUtc)
	{
		if (!FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
		{
			return string.Empty;
		}
		return CreateRunFolder(FamilyBrowserStandardPolicyStore.GetSnapshotFolder(workspaceRoot), "standard", FirstNonEmpty(displayName, sourceId, "StandardLibrary"), capturedAtUtc);
	}

	public static string CreateProjectRunFolder(string workspaceRoot, string documentTitle, string capturedAtUtc, string documentPath = "")
	{
		if (!FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
		{
			return string.Empty;
		}
		return CreateRunFolder(ProjectSnapshotStore.GetProjectHistoryFolder(workspaceRoot, documentPath, FirstNonEmpty(documentTitle, "Untitled")), "project", FirstNonEmpty(documentTitle, "Untitled"), capturedAtUtc);
	}

	public static string SaveLoadableSignature(string runFolder, string contextKind, string contextName, string categoryName, string familyName, string fingerprint, LoadableFamilyContentSignatureResult result)
	{
		string ignoredFailureReason = string.Empty;
		return SaveLoadableSignature(runFolder, contextKind, contextName, categoryName, familyName, fingerprint, result, ref ignoredFailureReason);
	}

	public static string SaveLoadableSignature(string runFolder, string contextKind, string contextName, string categoryName, string familyName, string fingerprint, LoadableFamilyContentSignatureResult result, ref string failureReason)
	{
		failureReason = string.Empty;
		string SaveLoadableSignature;
		if (string.IsNullOrWhiteSpace(runFolder) || result == null || string.IsNullOrWhiteSpace(result.Signature))
		{
			if (string.IsNullOrWhiteSpace(runFolder))
			{
				failureReason = "Signature diagnostics folder is empty.";
			}
			else if (result == null)
			{
				failureReason = "Signature result is empty.";
			}
			else
			{
				failureReason = "Signature source text is empty.";
			}
			SaveLoadableSignature = string.Empty;
		}
		else
		{
			StringBuilder builder = new StringBuilder();
			builder.AppendLine("context-kind=" + (contextKind ?? string.Empty));
			builder.AppendLine("context-name=" + (contextName ?? string.Empty));
			builder.AppendLine("category=" + (categoryName ?? string.Empty));
			builder.AppendLine("family=" + (familyName ?? string.Empty));
			builder.AppendLine("signature-mode=" + (result.Mode ?? string.Empty));
			builder.AppendLine("content-fingerprint=" + (fingerprint ?? string.Empty));
			if (!string.IsNullOrWhiteSpace(result.ErrorMessage))
			{
				builder.AppendLine("error-message=" + result.ErrorMessage);
			}
			builder.AppendLine("written-at-utc=" + DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
			builder.AppendLine();
			if (!string.IsNullOrWhiteSpace(result.DebugMetadata))
			{
				builder.AppendLine("----- signature-debug -----");
				builder.AppendLine(result.DebugMetadata);
				builder.AppendLine();
			}
			builder.AppendLine("----- signature-source -----");
			builder.AppendLine(result.Signature);
			string content = builder.ToString();
			try
			{
				string loadableFolder = Path.Combine(runFolder, "LoadableFamilies");
				Directory.CreateDirectory(loadableFolder);
				string fileName = SafeFileName(FirstNonEmpty(categoryName, "NoCategory") + "__" + FirstNonEmpty(familyName, "UnnamedFamily")) + ".signature.txt";
				string text = Path.Combine(loadableFolder, fileName);
				File.WriteAllText(text, content, Encoding.UTF8);
				SaveLoadableSignature = text;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception firstEx = ex;
				try
				{
					string loadableFolder2 = Path.Combine(runFolder, "LoadableFamilies");
					Directory.CreateDirectory(loadableFolder2);
					string shortName = "sig-" + ShortHash(FirstNonEmpty(contextKind, string.Empty) + "|" + FirstNonEmpty(contextName, string.Empty) + "|" + FirstNonEmpty(categoryName, string.Empty) + "|" + FirstNonEmpty(familyName, string.Empty)) + ".signature.txt";
					string text2 = Path.Combine(loadableFolder2, shortName);
					File.WriteAllText(text2, content, Encoding.UTF8);
					failureReason = "Signature diagnostic file used a short fallback name after the original name failed: " + firstEx.GetType().Name + " - " + firstEx.Message;
					SaveLoadableSignature = text2;
					ProjectData.ClearProjectError();
				}
				catch (Exception ex2)
				{
					ProjectData.SetProjectError(ex2);
					Exception fallbackEx = ex2;
					failureReason = "Signature diagnostic file write failed: " + firstEx.GetType().Name + " - " + firstEx.Message + "; fallback failed: " + fallbackEx.GetType().Name + " - " + fallbackEx.Message;
					SaveLoadableSignature = string.Empty;
					ProjectData.ClearProjectError();
				}
			}
		}
		return SaveLoadableSignature;
	}

	private static string CreateRunFolder(string baseFolder, string contextKind, string contextName, string capturedAtUtc)
	{
		string CreateRunFolder;
		if (string.IsNullOrWhiteSpace(baseFolder))
		{
			CreateRunFolder = string.Empty;
		}
		else
		{
			try
			{
				if (!DateTime.TryParse(capturedAtUtc, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var stamp))
				{
					stamp = DateTime.UtcNow;
				}
				string folderName = SafeFileName(FirstNonEmpty(contextKind, "context") + "-" + FirstNonEmpty(contextName, "Unnamed") + "-" + stamp.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture));
				string text = Path.Combine(baseFolder, "FingerprintDebug", folderName);
				Directory.CreateDirectory(text);
				CreateRunFolder = text;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				CreateRunFolder = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return CreateRunFolder;
	}

	private static string FirstNonEmpty(params string[] values)
	{
		if (values == null)
		{
			return string.Empty;
		}
		foreach (string value in values)
		{
			if (!string.IsNullOrWhiteSpace(value))
			{
				return value.Trim();
			}
		}
		return string.Empty;
	}

	private static void AddLoadableSignatureRoot(Dictionary<string, List<LoadableSignatureDebugRecord>> index, HashSet<string> seenPaths, string rootFolder, int maxRunFoldersPerRoot)
	{
		if (index == null || seenPaths == null || string.IsNullOrWhiteSpace(rootFolder))
		{
			return;
		}
		string expandedRoot = Environment.ExpandEnvironmentVariables(rootFolder.Trim());
		if (string.IsNullOrWhiteSpace(expandedRoot) || !Directory.Exists(expandedRoot))
		{
			return;
		}
		AddLoadableSignatureFolder(index, seenPaths, expandedRoot);
		AddLoadableSignatureFolder(index, seenPaths, Path.Combine(expandedRoot, "LoadableFamilies"));
		if (!string.Equals(Path.GetFileName(expandedRoot), "FingerprintDebug", StringComparison.OrdinalIgnoreCase))
		{
			return;
		}
		try
		{
			int limit = Math.Max(1, maxRunFoldersPerRoot);
			foreach (DirectoryInfo runFolder in (from x in new DirectoryInfo(expandedRoot).EnumerateDirectories("*", SearchOption.TopDirectoryOnly)
				orderby x.LastWriteTimeUtc descending
				select x).Take(limit))
			{
				AddLoadableSignatureFolder(index, seenPaths, Path.Combine(runFolder.FullName, "LoadableFamilies"));
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void AddLoadableSignatureFolder(Dictionary<string, List<LoadableSignatureDebugRecord>> index, HashSet<string> seenPaths, string loadableFolder)
	{
		if (index == null || seenPaths == null || string.IsNullOrWhiteSpace(loadableFolder) || !Directory.Exists(loadableFolder))
		{
			return;
		}
		try
		{
			string[] files = Directory.GetFiles(loadableFolder, "*.signature.txt", SearchOption.TopDirectoryOnly);
			foreach (string signaturePath in files)
			{
				AddLoadableSignatureFile(index, seenPaths, signaturePath);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void AddLoadableSignatureFile(Dictionary<string, List<LoadableSignatureDebugRecord>> index, HashSet<string> seenPaths, string signaturePath)
	{
		if (index == null || seenPaths == null || string.IsNullOrWhiteSpace(signaturePath))
		{
			return;
		}
		string expandedPath = Environment.ExpandEnvironmentVariables(signaturePath.Trim());
		if (string.IsNullOrWhiteSpace(expandedPath) || seenPaths.Contains(expandedPath))
		{
			return;
		}
		seenPaths.Add(expandedPath);
		LoadableSignatureDebugRecord record = ReadLoadableSignatureRecord(expandedPath);
		if (record != null && !string.IsNullOrWhiteSpace(record.FamilyName))
		{
			AddLoadableSignatureRecord(index, "f|" + FoldSignatureToken(record.FamilyName), record);
			if (!string.IsNullOrWhiteSpace(record.CategoryName))
			{
				AddLoadableSignatureRecord(index, "cf|" + FoldSignatureToken(record.CategoryName) + "|" + FoldSignatureToken(record.FamilyName), record);
			}
		}
	}

	private static void AddLoadableSignatureRecord(Dictionary<string, List<LoadableSignatureDebugRecord>> index, string key, LoadableSignatureDebugRecord record)
	{
		if (index != null && !string.IsNullOrWhiteSpace(key) && record != null)
		{
			List<LoadableSignatureDebugRecord> records = null;
			if (!index.TryGetValue(key, out records))
			{
				records = (index[key] = new List<LoadableSignatureDebugRecord>());
			}
			if (!records.Any([SpecialName] (LoadableSignatureDebugRecord x) => x != null && string.Equals(x.Path, record.Path, StringComparison.OrdinalIgnoreCase)))
			{
				records.Add(record);
			}
		}
	}

	private static LoadableSignatureDebugRecord SelectBestLoadableSignatureRecord(IEnumerable<LoadableSignatureDebugRecord> records)
	{
		return (from x in records?.Where([SpecialName] (LoadableSignatureDebugRecord x) => x != null)
			orderby !string.IsNullOrWhiteSpace(x.Fingerprint) descending, x.LastWriteUtcTicks descending
			select x).FirstOrDefault();
	}

	private static string FoldSignatureToken(string value)
	{
		return (value ?? string.Empty).Trim().ToLowerInvariant().Replace(" ", string.Empty)
			.Replace("_", string.Empty)
			.Replace("-", string.Empty)
			.Replace(".", string.Empty)
			.Replace(",", string.Empty);
	}

	private static string SafeFileName(string value)
	{
		char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
		string normalized = new string((value ?? string.Empty).Select([SpecialName] (char ch) => (!Enumerable.Contains(invalidFileNameChars, ch)) ? ch : '_').ToArray()).Trim();
		if (normalized.Length == 0)
		{
			return "Unnamed";
		}
		if (normalized.Length > 120)
		{
			normalized = normalized.Substring(0, 120);
		}
		return normalized;
	}

	private static string ShortHash(string value)
	{
		checked
		{
			string ShortHash;
			try
			{
				using SHA256 sha = SHA256.Create();
				byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty));
				StringBuilder builder = new StringBuilder();
				int num = Math.Min(7, bytes.Length - 1);
				for (int i = 0; i <= num; i++)
				{
					builder.Append(bytes[i].ToString("x2", CultureInfo.InvariantCulture));
				}
				ShortHash = builder.ToString();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ShortHash = Guid.NewGuid().ToString("N").Substring(0, 16);
				ProjectData.ClearProjectError();
			}
			return ShortHash;
		}
	}
}
