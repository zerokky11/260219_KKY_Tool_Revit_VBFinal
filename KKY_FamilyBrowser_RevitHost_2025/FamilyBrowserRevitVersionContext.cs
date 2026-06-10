using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserRevitVersionContext
{
	private static readonly object SyncRoot = RuntimeHelpers.GetObjectValue(new object());

	private static string _versionNumber = string.Empty;

	private FamilyBrowserRevitVersionContext()
	{
	}

	public static void SetCurrentVersion(string versionNumber)
	{
		string normalized = NormalizeVersionNumber(versionNumber);
		if (string.IsNullOrWhiteSpace(normalized))
		{
			return;
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			_versionNumber = normalized;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static void SetCurrentVersion(object revitObject)
	{
		if (revitObject != null)
		{
			SetCurrentVersion(ResolveVersionNumber(RuntimeHelpers.GetObjectValue(revitObject)));
		}
	}

	public static string CurrentVersionNumber()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			return _versionNumber ?? string.Empty;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static string CurrentVersionFolderName()
	{
		string version = NormalizeVersionNumber(CurrentVersionNumber());
		if (string.IsNullOrWhiteSpace(version))
		{
			return "RevitUnknown";
		}
		return "Revit" + version;
	}

	public static string VersionedDataRoot(string dataRoot)
	{
		if (string.IsNullOrWhiteSpace(dataRoot))
		{
			return string.Empty;
		}
		return Path.Combine(dataRoot, "RevitVersions", CurrentVersionFolderName());
	}

	public static bool IsPathInCurrentVersionRoot(string path, string dataRoot)
	{
		if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(dataRoot))
		{
			return false;
		}
		return IsSameOrChildPath(path, VersionedDataRoot(dataRoot));
	}

	private static string NormalizeVersionNumber(string value)
	{
		string raw = (value ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(raw))
		{
			return string.Empty;
		}
		string digits = new string(raw.Where([SpecialName] (char ch) => char.IsDigit(ch)).ToArray());
		if (string.IsNullOrWhiteSpace(digits))
		{
			return string.Empty;
		}
		int fourDigitIndex = digits.IndexOf("20", StringComparison.Ordinal);
		if (fourDigitIndex >= 0 && digits.Length >= checked(fourDigitIndex + 4))
		{
			return digits.Substring(fourDigitIndex, 4);
		}
		if (digits.Length == 2)
		{
			return "20" + digits;
		}
		if (digits.Length >= 4)
		{
			return digits.Substring(0, 4);
		}
		return digits;
	}

	private static string ResolveVersionNumber(object revitObject)
	{
		string direct = TryReadStringProperty(RuntimeHelpers.GetObjectValue(revitObject), "VersionNumber");
		if (!string.IsNullOrWhiteSpace(direct))
		{
			return direct;
		}
		direct = TryReadStringProperty(RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(TryReadObjectProperty(RuntimeHelpers.GetObjectValue(revitObject), "Application"))), "VersionNumber");
		if (!string.IsNullOrWhiteSpace(direct))
		{
			return direct;
		}
		direct = TryReadStringProperty(RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(TryReadObjectProperty(RuntimeHelpers.GetObjectValue(revitObject), "ControlledApplication"))), "VersionNumber");
		if (!string.IsNullOrWhiteSpace(direct))
		{
			return direct;
		}
		direct = TryReadStringProperty(RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(TryReadObjectProperty(RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(TryReadObjectProperty(RuntimeHelpers.GetObjectValue(revitObject), "Document"))), "Application"))), "VersionNumber");
		if (!string.IsNullOrWhiteSpace(direct))
		{
			return direct;
		}
		return string.Empty;
	}

	private static object TryReadObjectProperty(object instance, string propertyName)
	{
		object TryReadObjectProperty;
		if (instance == null || string.IsNullOrWhiteSpace(propertyName))
		{
			TryReadObjectProperty = null;
		}
		else
		{
			try
			{
				TryReadObjectProperty = instance.GetType().GetProperty(propertyName)?.GetValue(RuntimeHelpers.GetObjectValue(instance), null);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				TryReadObjectProperty = null;
				ProjectData.ClearProjectError();
			}
		}
		return TryReadObjectProperty;
	}

	private static string TryReadStringProperty(object instance, string propertyName)
	{
		object value = RuntimeHelpers.GetObjectValue(TryReadObjectProperty(RuntimeHelpers.GetObjectValue(instance), propertyName));
		string TryReadStringProperty;
		if (value == null)
		{
			TryReadStringProperty = string.Empty;
		}
		else
		{
			try
			{
				TryReadStringProperty = Convert.ToString(RuntimeHelpers.GetObjectValue(value), CultureInfo.InvariantCulture);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				TryReadStringProperty = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return TryReadStringProperty;
	}

	private static bool IsSameOrChildPath(string candidate, string parent)
	{
		bool IsSameOrChildPath;
		if (string.IsNullOrWhiteSpace(candidate) || string.IsNullOrWhiteSpace(parent))
		{
			IsSameOrChildPath = false;
		}
		else
		{
			try
			{
				string candidateFull = Path.GetFullPath(Environment.ExpandEnvironmentVariables(candidate.Trim())).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
				string parentFull = Path.GetFullPath(Environment.ExpandEnvironmentVariables(parent.Trim())).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
				IsSameOrChildPath = string.Equals(candidateFull, parentFull, StringComparison.OrdinalIgnoreCase) || candidateFull.StartsWith(parentFull + Conversions.ToString(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase) || candidateFull.StartsWith(parentFull + Conversions.ToString(Path.AltDirectorySeparatorChar), StringComparison.OrdinalIgnoreCase);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				IsSameOrChildPath = false;
				ProjectData.ClearProjectError();
			}
		}
		return IsSameOrChildPath;
	}
}
