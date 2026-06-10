using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserUserSettingsStore
{
	private static readonly object SyncRoot = RuntimeHelpers.GetObjectValue(new object());

	private static string _runtimeLanguageCode = string.Empty;

	private static string _runtimeAdminModeValue = string.Empty;

	private FamilyBrowserUserSettingsStore()
	{
	}

	public static string GetSettingsRoot(bool createFolder = false)
	{
		string root = ResolveManagedUserSettingsRoot();
		if (string.IsNullOrWhiteSpace(root))
		{
			return string.Empty;
		}
		if (createFolder)
		{
			Directory.CreateDirectory(root);
		}
		return root;
	}

	public static string GetLanguageSettingsPath()
	{
		string root = GetSettingsRoot();
		if (string.IsNullOrWhiteSpace(root))
		{
			return string.Empty;
		}
		return Path.Combine(root, "language.txt");
	}

	public static string GetAdminModeSettingsPath()
	{
		string root = GetSettingsRoot();
		if (string.IsNullOrWhiteSpace(root))
		{
			return string.Empty;
		}
		return Path.Combine(root, "admin-mode.txt");
	}

	public static string LoadLanguageCode()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (string.Equals(_runtimeLanguageCode, "en", StringComparison.OrdinalIgnoreCase))
			{
				return "en";
			}
			if (string.Equals(_runtimeLanguageCode, "ko", StringComparison.OrdinalIgnoreCase))
			{
				return "ko";
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		if (string.Equals(ReadTrimmed(GetLanguageSettingsPath()), "en", StringComparison.OrdinalIgnoreCase))
		{
			return "en";
		}
		return "ko";
	}

	public static void SaveLanguageCode(string languageCode)
	{
		string normalized = (string.Equals((languageCode ?? string.Empty).Trim(), "en", StringComparison.OrdinalIgnoreCase) ? "en" : "ko");
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			_runtimeLanguageCode = normalized;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		WriteText(GetLanguageSettingsPathForWrite(), normalized);
	}

	public static bool LoadAdminModeEnabled()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (string.Equals(_runtimeAdminModeValue, "on", StringComparison.OrdinalIgnoreCase))
			{
				return true;
			}
			if (string.Equals(_runtimeAdminModeValue, "off", StringComparison.OrdinalIgnoreCase))
			{
				return false;
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		string value = ReadTrimmed(GetAdminModeSettingsPath()).ToLowerInvariant();
		return string.Equals(value, "on", StringComparison.Ordinal) || string.Equals(value, "true", StringComparison.Ordinal) || string.Equals(value, "1", StringComparison.Ordinal);
	}

	public static void SaveAdminModeEnabled(bool enabled)
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			_runtimeAdminModeValue = (enabled ? "on" : "off");
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		WriteText(GetAdminModeSettingsPathForWrite(), enabled ? "on" : "off");
	}

	private static string GetLanguageSettingsPathForWrite()
	{
		string root = GetSettingsRoot(createFolder: true);
		if (string.IsNullOrWhiteSpace(root))
		{
			return string.Empty;
		}
		return Path.Combine(root, "language.txt");
	}

	private static string GetAdminModeSettingsPathForWrite()
	{
		string root = GetSettingsRoot(createFolder: true);
		if (string.IsNullOrWhiteSpace(root))
		{
			return string.Empty;
		}
		return Path.Combine(root, "admin-mode.txt");
	}

	private static string ResolveManagedUserSettingsRoot()
	{
		string policyPath = FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath();
		if (string.IsNullOrWhiteSpace(policyPath))
		{
			return string.Empty;
		}
		string policyFolder = Path.GetDirectoryName(policyPath);
		if (string.IsNullOrWhiteSpace(policyFolder))
		{
			return string.Empty;
		}
		string root = policyFolder.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
		if (string.Equals(Path.GetFileName(root), "Config", StringComparison.OrdinalIgnoreCase))
		{
			DirectoryInfo parent = Directory.GetParent(root);
			if (parent != null && !string.IsNullOrWhiteSpace(parent.FullName))
			{
				root = parent.FullName;
			}
		}
		return Path.Combine(root, "UserSettings", SafeFolderName(Environment.UserName));
	}

	private static string ReadTrimmed(string path)
	{
		try
		{
			if (string.IsNullOrWhiteSpace(path))
			{
				return string.Empty;
			}
			if (File.Exists(path))
			{
				return File.ReadAllText(path, Encoding.UTF8).Trim();
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Empty;
	}

	private static void WriteText(string path, string value)
	{
		if (!string.IsNullOrWhiteSpace(path))
		{
			string folder = Path.GetDirectoryName(path);
			if (!string.IsNullOrWhiteSpace(folder))
			{
				Directory.CreateDirectory(folder);
			}
			File.WriteAllText(path, value ?? string.Empty, Encoding.UTF8);
		}
	}

	private static string SafeFolderName(string value)
	{
		string raw = (value ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(raw))
		{
			raw = "user";
		}
		char[] invalid = Path.GetInvalidFileNameChars();
		StringBuilder builder = new StringBuilder(raw.Length);
		string text = raw;
		foreach (char ch in text)
		{
			if (Array.IndexOf(invalid, ch) >= 0)
			{
				builder.Append('_');
			}
			else
			{
				builder.Append(ch);
			}
		}
		string safe = builder.ToString().Trim('.', ' ');
		if (string.IsNullOrWhiteSpace(safe))
		{
			return "user";
		}
		return safe;
	}
}
