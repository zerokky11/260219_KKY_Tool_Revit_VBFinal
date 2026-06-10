using System;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserMachineConfigStore
{
	private const string RuntimeOnlyConfigPath = "(runtime only - managed shared folder required)";

	private static readonly object SyncRoot = RuntimeHelpers.GetObjectValue(new object());

	private static FamilyBrowserMachineConfig _runtimeConfig = FamilyBrowserMachineConfig.CreateDefault();

	private FamilyBrowserMachineConfigStore()
	{
	}

	public static string GetConfigPath()
	{
		return "(runtime only - managed shared folder required)";
	}

	public static FamilyBrowserMachineConfig Load()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			return CloneConfig(_runtimeConfig);
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static string Save(FamilyBrowserMachineConfig config, string currentUser)
	{
		if (config == null)
		{
			config = FamilyBrowserMachineConfig.CreateDefault();
		}
		Normalize(config);
		config.LastUpdatedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		config.LastUpdatedBy = currentUser ?? string.Empty;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			_runtimeConfig = CloneConfig(config);
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		return "(runtime only - managed shared folder required)";
	}

	public static string SetManagedPolicyPath(string policyPath, string currentUser)
	{
		FamilyBrowserMachineConfig familyBrowserMachineConfig = Load();
		string normalizedPolicyPath = Environment.ExpandEnvironmentVariables((policyPath ?? string.Empty).Trim());
		if (!IsAllowedManagedPath(normalizedPolicyPath))
		{
			throw new InvalidOperationException("Family Browser managed data must be stored in the homepage-managed shared folder, not a local C/AppData/ProgramData path.");
		}
		familyBrowserMachineConfig.UseManagedPolicy = true;
		familyBrowserMachineConfig.ManagedPolicyPath = normalizedPolicyPath;
		return Save(familyBrowserMachineConfig, currentUser);
	}

	public static string ClearManagedPolicyPath(string currentUser)
	{
		FamilyBrowserMachineConfig familyBrowserMachineConfig = Load();
		familyBrowserMachineConfig.UseManagedPolicy = false;
		familyBrowserMachineConfig.ManagedPolicyPath = string.Empty;
		return Save(familyBrowserMachineConfig, currentUser);
	}

	public static string ResolveManagedPolicyPath()
	{
		FamilyBrowserMachineConfig config = Load();
		if (!config.UseManagedPolicy || string.IsNullOrWhiteSpace(config.ManagedPolicyPath))
		{
			return string.Empty;
		}
		return Environment.ExpandEnvironmentVariables(config.ManagedPolicyPath);
	}

	public static string SetDeploymentBootstrapUrl(string bootstrapUrl, string currentUser)
	{
		FamilyBrowserMachineConfig familyBrowserMachineConfig = Load();
		familyBrowserMachineConfig.DeploymentBootstrapUrl = (bootstrapUrl ?? string.Empty).Trim();
		return Save(familyBrowserMachineConfig, currentUser);
	}

	public static string ResolveDeploymentBootstrapUrl()
	{
		return (Load().DeploymentBootstrapUrl ?? string.Empty).Trim();
	}

	public static void RecordBootstrapResult(FamilyBrowserDeploymentBootstrapResult result, string currentUser)
	{
		if (result != null)
		{
			FamilyBrowserMachineConfig familyBrowserMachineConfig = Load();
			familyBrowserMachineConfig.LastBootstrapCheckUtc = result.CheckedAtUtc;
			familyBrowserMachineConfig.LastBootstrapStatus = result.Message;
			familyBrowserMachineConfig.LastBootstrapSource = result.Source;
			Save(familyBrowserMachineConfig, currentUser);
		}
	}

	private static FamilyBrowserMachineConfig CloneConfig(FamilyBrowserMachineConfig source)
	{
		if (source == null)
		{
			return FamilyBrowserMachineConfig.CreateDefault();
		}
		return new FamilyBrowserMachineConfig
		{
			UseManagedPolicy = source.UseManagedPolicy,
			ManagedPolicyPath = (source.ManagedPolicyPath ?? string.Empty),
			DeploymentBootstrapUrl = (source.DeploymentBootstrapUrl ?? string.Empty),
			LastBootstrapCheckUtc = (source.LastBootstrapCheckUtc ?? string.Empty),
			LastBootstrapStatus = (source.LastBootstrapStatus ?? string.Empty),
			LastBootstrapSource = (source.LastBootstrapSource ?? string.Empty),
			LastUpdatedUtc = (source.LastUpdatedUtc ?? string.Empty),
			LastUpdatedBy = (source.LastUpdatedBy ?? string.Empty)
		};
	}

	private static void Normalize(FamilyBrowserMachineConfig config)
	{
		if (config != null)
		{
			config.ManagedPolicyPath = (config.ManagedPolicyPath ?? string.Empty).Trim();
			config.DeploymentBootstrapUrl = (config.DeploymentBootstrapUrl ?? string.Empty).Trim();
			config.LastBootstrapCheckUtc = (config.LastBootstrapCheckUtc ?? string.Empty).Trim();
			config.LastBootstrapStatus = (config.LastBootstrapStatus ?? string.Empty).Trim();
			config.LastBootstrapSource = (config.LastBootstrapSource ?? string.Empty).Trim();
			if (string.IsNullOrWhiteSpace(config.ManagedPolicyPath))
			{
				config.UseManagedPolicy = false;
			}
		}
	}

	private static bool IsAllowedManagedPath(string path)
	{
		bool IsAllowedManagedPath;
		if (string.IsNullOrWhiteSpace(path))
		{
			IsAllowedManagedPath = false;
		}
		else
		{
			try
			{
				if (path.StartsWith("\\\\", StringComparison.Ordinal))
				{
					IsAllowedManagedPath = true;
				}
				else
				{
					string root = Path.GetPathRoot(path);
					if (string.IsNullOrWhiteSpace(root))
					{
						IsAllowedManagedPath = false;
					}
					else
					{
						string windowsRoot = Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.Windows));
						IsAllowedManagedPath = (string.IsNullOrWhiteSpace(windowsRoot) || !string.Equals(root, windowsRoot, StringComparison.OrdinalIgnoreCase)) && !IsSameOrChildPath(path, Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)) && !IsSameOrChildPath(path, Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)) && !IsSameOrChildPath(path, Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)) && !IsSameOrChildPath(path, Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)) && !IsSameOrChildPath(path, Path.GetTempPath());
					}
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				IsAllowedManagedPath = false;
				ProjectData.ClearProjectError();
			}
		}
		return IsAllowedManagedPath;
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
				string candidateFull = Path.GetFullPath(candidate).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
				string parentFull = Path.GetFullPath(parent).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
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
