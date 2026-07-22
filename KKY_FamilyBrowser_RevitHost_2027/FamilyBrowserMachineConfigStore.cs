using System;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserMachineConfigStore
{
	private const string RuntimeOnlyConfigPath = "(runtime only - managed shared folder required)";

	private const string LastKnownManagedPolicyPathFileName = "last-known-managed-policy-path.txt";

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
		using (IDisposable managementContextLease = FamilyBrowserManagementContextLock.Acquire(TimeSpan.FromSeconds(30.0)))
		{
			return SetManagedPolicyPathNoLock(policyPath, currentUser);
		}
	}

	public static string GetLastKnownManagedPolicyPathCachePath()
	{
		string settingsRoot = FamilyBrowserUserSettingsStore.GetSettingsRoot();
		return string.IsNullOrWhiteSpace(settingsRoot) ? string.Empty : Path.Combine(settingsRoot, LastKnownManagedPolicyPathFileName);
	}

	public static string ResolveLastKnownManagedPolicyPath()
	{
		try
		{
			string cachePath = GetLastKnownManagedPolicyPathCachePath();
			if (string.IsNullOrWhiteSpace(cachePath) || !File.Exists(cachePath))
			{
				return string.Empty;
			}
			string policyPath = Environment.ExpandEnvironmentVariables(File.ReadAllText(cachePath, Encoding.UTF8).Trim());
			return IsUsableCachedManagedPolicyPath(policyPath) ? policyPath : string.Empty;
		}
		catch
		{
			return string.Empty;
		}
	}

	public static bool TryRestoreLastKnownManagedPolicyPath(string currentUser, out string issue)
	{
		issue = string.Empty;
		string activePath = ResolveManagedPolicyPath();
		if (IsUsableCachedManagedPolicyPath(activePath))
		{
			return true;
		}
		string cachedPath = ResolveLastKnownManagedPolicyPath();
		if (string.IsNullOrWhiteSpace(cachedPath))
		{
			return false;
		}
		try
		{
			SetManagedPolicyPath(cachedPath, currentUser);
			return true;
		}
		catch (Exception ex)
		{
			issue = ex.Message;
			return false;
		}
	}

	private static string SetManagedPolicyPathNoLock(string policyPath, string currentUser)
	{
		FamilyBrowserMachineConfig familyBrowserMachineConfig = Load();
		string normalizedPolicyPath = Environment.ExpandEnvironmentVariables((policyPath ?? string.Empty).Trim());
		if (!IsAllowedManagedPath(normalizedPolicyPath))
		{
			throw new InvalidOperationException(FamilyBrowserLanguageService.Text("Family Browser managed data must be stored in the homepage-managed shared folder, not a local C/AppData/ProgramData path.", "Family Browser 관리 데이터는 로컬 C/AppData/ProgramData 경로가 아니라 홈페이지에서 지정한 공용 관리 폴더에 저장해야 합니다."));
		}
		string previousPolicyPath = familyBrowserMachineConfig.UseManagedPolicy ? familyBrowserMachineConfig.ManagedPolicyPath : string.Empty;
		bool managementContextChanged = !SameManagedPolicyPath(previousPolicyPath, normalizedPolicyPath);
		if (managementContextChanged && FamilyBrowserElementChangeTrackingService.GetActiveUncommittedSessionCount() > 0)
		{
			throw new InvalidOperationException(FamilyBrowserLanguageService.Text("The management folder cannot change while this Revit session has tracked activity that has not reached a successful Save or Synchronize boundary. Save or synchronize the open project, then try again.", "현재 Revit 세션에 저장 또는 동기화 성공 경계에 도달하지 않은 추적 작업이 있어 관리폴더를 변경할 수 없습니다. 열려 있는 프로젝트를 저장하거나 동기화한 뒤 다시 시도하세요."));
		}
		if (managementContextChanged && !FamilyBrowserElementChangeTrackingService.IsManagedFolderTransitionAuthorized() &&
			(FamilyBrowserElementChangeTrackingService.GetProtectedRecoverySessionCount() > 0 ||
			 FamilyBrowserTrackingPersistenceService.HasBlockingElementSessionCheckpointForManagedPolicyPath(normalizedPolicyPath)))
		{
			throw new InvalidOperationException(FamilyBrowserLanguageService.Text("The management folder cannot change while a protected workshared local-save checkpoint is waiting for synchronization or history promotion. Complete synchronization, or use the verified Migrate Existing Data workflow.", "동기화 또는 관리 이력 승격을 기다리는 보호된 워크셰어링 로컬 저장 체크포인트가 있어 관리폴더를 변경할 수 없습니다. 동기화를 완료하거나 검증된 '기존 데이터 이관' 절차를 사용하세요."));
		}
		familyBrowserMachineConfig.UseManagedPolicy = true;
		familyBrowserMachineConfig.ManagedPolicyPath = normalizedPolicyPath;
		string savedPath = Save(familyBrowserMachineConfig, currentUser);
		RememberLastKnownManagedPolicyPath(normalizedPolicyPath);
		if (managementContextChanged)
		{
			FamilyBrowserElementChangeTrackingService.NotifyManagementContextChanged();
		}
		return savedPath;
	}

	public static string ClearManagedPolicyPath(string currentUser)
	{
		using (IDisposable managementContextLease = FamilyBrowserManagementContextLock.Acquire(TimeSpan.FromSeconds(30.0)))
		{
			return ClearManagedPolicyPathNoLock(currentUser);
		}
	}

	private static string ClearManagedPolicyPathNoLock(string currentUser)
	{
		FamilyBrowserMachineConfig familyBrowserMachineConfig = Load();
		bool managementContextChanged = familyBrowserMachineConfig.UseManagedPolicy || !string.IsNullOrWhiteSpace(familyBrowserMachineConfig.ManagedPolicyPath);
		if (managementContextChanged && FamilyBrowserElementChangeTrackingService.GetActiveUncommittedSessionCount() > 0)
		{
			throw new InvalidOperationException(FamilyBrowserLanguageService.Text("The management folder cannot be cleared while this Revit session has tracked activity that has not reached a successful Save or Synchronize boundary. Save or synchronize the open project, then try again.", "현재 Revit 세션에 저장 또는 동기화 성공 경계에 도달하지 않은 추적 작업이 있어 관리폴더를 해제할 수 없습니다. 열려 있는 프로젝트를 저장하거나 동기화한 뒤 다시 시도하세요."));
		}
		if (managementContextChanged && !FamilyBrowserElementChangeTrackingService.IsManagedFolderTransitionAuthorized() &&
			(FamilyBrowserElementChangeTrackingService.GetProtectedRecoverySessionCount() > 0 ||
			 FamilyBrowserTrackingPersistenceService.HasBlockingElementSessionCheckpointForManagedPolicyPath(string.Empty)))
		{
			throw new InvalidOperationException(FamilyBrowserLanguageService.Text("The management folder cannot be cleared while a protected workshared local-save checkpoint is waiting for synchronization or history promotion.", "동기화 또는 관리 이력 승격을 기다리는 보호된 워크셰어링 로컬 저장 체크포인트가 있어 관리폴더를 해제할 수 없습니다."));
		}
		familyBrowserMachineConfig.UseManagedPolicy = false;
		familyBrowserMachineConfig.ManagedPolicyPath = string.Empty;
		string savedPath = Save(familyBrowserMachineConfig, currentUser);
		if (managementContextChanged)
		{
			FamilyBrowserElementChangeTrackingService.NotifyManagementContextChanged();
		}
		return savedPath;
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

	private static bool IsUsableCachedManagedPolicyPath(string policyPath)
	{
		if (!IsAllowedManagedPath(policyPath))
		{
			return false;
		}
		try
		{
			string parent = Path.GetDirectoryName(policyPath);
			return !string.IsNullOrWhiteSpace(parent) && Directory.Exists(parent);
		}
		catch
		{
			return false;
		}
	}

	private static void RememberLastKnownManagedPolicyPath(string policyPath)
	{
		try
		{
			string normalized = Environment.ExpandEnvironmentVariables((policyPath ?? string.Empty).Trim());
			if (!IsAllowedManagedPath(normalized))
			{
				return;
			}
			string cachePath = GetLastKnownManagedPolicyPathCachePath();
			if (string.IsNullOrWhiteSpace(cachePath))
			{
				return;
			}
			Directory.CreateDirectory(Path.GetDirectoryName(cachePath));
			File.WriteAllText(cachePath, normalized, new UTF8Encoding(false));
		}
		catch
		{
		}
	}

	private static bool SameManagedPolicyPath(string left, string right)
	{
		return string.Equals(FamilyBrowserPathIdentityService.NormalizePath(left), FamilyBrowserPathIdentityService.NormalizePath(right), StringComparison.OrdinalIgnoreCase);
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
				string candidateFull = Path.GetFullPath(candidate).TrimEnd(new char[2]
				{
					Path.DirectorySeparatorChar,
					Path.AltDirectorySeparatorChar
				});
				string parentFull = Path.GetFullPath(parent).TrimEnd(new char[2]
				{
					Path.DirectorySeparatorChar,
					Path.AltDirectorySeparatorChar
				});
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
