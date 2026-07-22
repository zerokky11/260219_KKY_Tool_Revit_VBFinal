using System;
using System.IO;
using System.Text;

public static class FamilyBrowserMeasurementUnitPreferenceService
{
	private static readonly object SyncRoot = new object();

	private static string _runtimeUnit = string.Empty;

	public static string Normalize(string value)
	{
		return string.Equals((value ?? string.Empty).Trim(), "in", StringComparison.OrdinalIgnoreCase) ? "in" : "mm";
	}

	public static string Load()
	{
		lock (SyncRoot)
		{
			if (string.Equals(_runtimeUnit, "mm", StringComparison.Ordinal) || string.Equals(_runtimeUnit, "in", StringComparison.Ordinal))
			{
				return _runtimeUnit;
			}
			_runtimeUnit = ReadUnit(GetSettingsPath());
			return _runtimeUnit;
		}
	}

	public static bool Save(string value)
	{
		string normalized = Normalize(value);
		lock (SyncRoot)
		{
			_runtimeUnit = normalized;
		}
		return WriteUnit(GetSettingsPath(), normalized);
	}

	internal static string LoadFromPathForAudit(string path)
	{
		return ReadUnit(path);
	}

	internal static bool SaveToPathForAudit(string path, string value)
	{
		return WriteUnit(path, Normalize(value));
	}

	private static string GetSettingsPath()
	{
		string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
		if (string.IsNullOrWhiteSpace(localAppData))
		{
			return string.Empty;
		}
		return Path.Combine(localAppData, "KKY", "FamilyBrowser", "Settings", "measurement-unit.txt");
	}

	private static string ReadUnit(string path)
	{
		try
		{
			if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
			{
				return Normalize(File.ReadAllText(path, Encoding.UTF8));
			}
		}
		catch
		{
		}
		return "mm";
	}

	private static bool WriteUnit(string path, string value)
	{
		string tempPath = string.Empty;
		try
		{
			if (string.IsNullOrWhiteSpace(path))
			{
				return false;
			}
			string folder = Path.GetDirectoryName(path);
			if (string.IsNullOrWhiteSpace(folder))
			{
				return false;
			}
			Directory.CreateDirectory(folder);
			tempPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
			File.WriteAllText(tempPath, Normalize(value), new UTF8Encoding(false));
			FamilyBrowserAtomicFileService.Promote(tempPath, path);
			tempPath = string.Empty;
			return true;
		}
		catch
		{
			return false;
		}
		finally
		{
			try
			{
				if (!string.IsNullOrWhiteSpace(tempPath) && File.Exists(tempPath))
				{
					File.Delete(tempPath);
				}
			}
			catch
			{
			}
		}
	}
}
