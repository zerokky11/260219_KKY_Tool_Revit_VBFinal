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

	private static string _runtimeDetailedSystemTypeComparisonValue = string.Empty;

	private FamilyBrowserUserSettingsStore()
	{
	}

	public static string GetSettingsRoot(bool createFolder = false)
	{
		string root = ResolveLocalUserSettingsRoot();
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

	public static string GetDetailedSystemTypeComparisonSettingsPath()
	{
		string root = GetSettingsRoot();
		if (string.IsNullOrWhiteSpace(root))
		{
			return string.Empty;
		}
		return Path.Combine(root, "system-type-detail-components.txt");
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

	public static bool HasAdminModePreference()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (string.Equals(_runtimeAdminModeValue, "on", StringComparison.OrdinalIgnoreCase) || string.Equals(_runtimeAdminModeValue, "off", StringComparison.OrdinalIgnoreCase))
			{
				return true;
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
		return string.Equals(value, "on", StringComparison.Ordinal) || string.Equals(value, "true", StringComparison.Ordinal) || string.Equals(value, "1", StringComparison.Ordinal) || string.Equals(value, "off", StringComparison.Ordinal) || string.Equals(value, "false", StringComparison.Ordinal) || string.Equals(value, "0", StringComparison.Ordinal);
	}

	public static bool ResolveInitialAdminModeEnabled(bool canEnableAdminMode)
	{
		if (HasAdminModePreference())
		{
			return LoadAdminModeEnabled();
		}
		return canEnableAdminMode;
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

	public static bool ResolveDetailedSystemTypeComparisonEnabled(FamilyBrowserStandardPolicy policy)
	{
		string value = string.Empty;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			value = _runtimeDetailedSystemTypeComparisonValue;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		if (string.IsNullOrWhiteSpace(value))
		{
			value = ReadTrimmed(GetDetailedSystemTypeComparisonSettingsPath());
		}
		if (string.Equals(value, "on", StringComparison.OrdinalIgnoreCase) || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) || string.Equals(value, "1", StringComparison.Ordinal))
		{
			return true;
		}
		if (string.Equals(value, "off", StringComparison.OrdinalIgnoreCase) || string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) || string.Equals(value, "0", StringComparison.Ordinal))
		{
			return false;
		}
		return FamilyBrowserStandardPolicyStore.IsDetailedSystemTypeComparisonEnabled(policy);
	}

	public static void SaveDetailedSystemTypeComparisonEnabled(bool enabled)
	{
		string value = enabled ? "on" : "off";
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			_runtimeDetailedSystemTypeComparisonValue = value;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		WriteText(GetDetailedSystemTypeComparisonSettingsPathForWrite(), value);
	}

	public static bool ClearDetailedSystemTypeComparisonPreference()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			_runtimeDetailedSystemTypeComparisonValue = string.Empty;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		try
		{
			string path = GetDetailedSystemTypeComparisonSettingsPath();
			if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
			{
				File.Delete(path);
			}
			return true;
		}
		catch
		{
			return false;
		}
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

	private static string GetDetailedSystemTypeComparisonSettingsPathForWrite()
	{
		string root = GetSettingsRoot(createFolder: true);
		if (string.IsNullOrWhiteSpace(root))
		{
			return string.Empty;
		}
		return Path.Combine(root, "system-type-detail-components.txt");
	}

	private static string ResolveLocalUserSettingsRoot()
	{
		string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
		if (string.IsNullOrWhiteSpace(localAppData))
		{
			return string.Empty;
		}
		return Path.Combine(localAppData, "KKY", "FamilyBrowser", "Settings");
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
