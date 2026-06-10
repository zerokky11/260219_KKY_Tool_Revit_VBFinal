using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Cache;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserDeploymentBootstrapService
{
	private enum BootstrapPathTargetKind
	{
		File,
		Directory,
		FileOrParentDirectory
	}

	private sealed class ManagedPolicyCandidate
	{
		public string DisplayPath { get; set; }

		public string PolicyPath { get; set; }

		public bool IsRoot { get; set; }

		public ManagedPolicyCandidate()
		{
			DisplayPath = string.Empty;
			PolicyPath = string.Empty;
			IsRoot = false;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__34_002D0
	{
		public string _0024VB_0024Local_key;

		public _Closure_0024__34_002D0(_Closure_0024__34_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_key = arg0._0024VB_0024Local_key;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__1(FamilyBrowserStandardLibrarySlot x)
		{
			if (x != null)
			{
				return string.Equals(FamilyBrowserPolicyKey.Normalize(x.Discipline), _0024VB_0024Local_key, StringComparison.OrdinalIgnoreCase);
			}
			return false;
		}
	}

	public const string DefaultBootstrapUrl = "https://update.zerokky.com/family-browser/bootstrap.json";

	public const string DefaultBootstrapIndexUrl = "https://update.zerokky.com/family-browser/bootstrap-index.json";

	public const string DisabledBootstrapToken = "disabled";

	private const int PathProbeTimeoutMilliseconds = 800;

	private static readonly object SyncRoot = RuntimeHelpers.GetObjectValue(new object());

	private static DateTime _lastCheckedUtc = DateTime.MinValue;

	private static FamilyBrowserDeploymentBootstrapResult _lastResult;

	private static string _lastProjectIdentityKey = string.Empty;

	private static string _runtimeBootstrapCacheJson = string.Empty;

	private FamilyBrowserDeploymentBootstrapService()
	{
	}

	public static FamilyBrowserDeploymentBootstrapResult GetLastResult()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			return _lastResult;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static string ClearCache()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			_lastCheckedUtc = DateTime.MinValue;
			_lastResult = null;
			_lastProjectIdentityKey = string.Empty;
			_runtimeBootstrapCacheJson = string.Empty;
			return GetCachePath();
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static FamilyBrowserDeploymentBootstrapResult TryApply(string workspaceRoot, string currentUser, bool force = false, FamilyBrowserDeploymentProjectIdentity projectIdentity = null)
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			DateTime now = DateTime.UtcNow;
			string projectIdentityKey = BuildProjectIdentityKey(projectIdentity);
			if (!force && _lastResult != null)
			{
				int minutes = Math.Max(1, _lastResult.RefreshMinutes);
				if ((now - _lastCheckedUtc).TotalMinutes < (double)minutes && string.Equals(_lastProjectIdentityKey, projectIdentityKey, StringComparison.Ordinal))
				{
					return _lastResult;
				}
			}
			string bootstrapUrl = ResolveBootstrapUrl(projectIdentity);
			if (!force && _lastResult != null && string.Equals(_lastResult.BootstrapUrl ?? string.Empty, bootstrapUrl, StringComparison.OrdinalIgnoreCase) && string.Equals(_lastProjectIdentityKey, projectIdentityKey, StringComparison.Ordinal))
			{
				int minutes2 = Math.Max(1, _lastResult.RefreshMinutes);
				if ((now - _lastCheckedUtc).TotalMinutes < (double)minutes2)
				{
					return _lastResult;
				}
			}
			FamilyBrowserDeploymentBootstrapResult result = NewResult();
			result.CachePath = GetCachePath();
			try
			{
				result.BootstrapUrl = bootstrapUrl;
				if (IsBootstrapDisabledValue(bootstrapUrl))
				{
					result.Message = "Deployment bootstrap is disabled.";
					RememberResult(result, currentUser, projectIdentityKey);
					return result;
				}
				string json = string.Empty;
				string source = string.Empty;
				try
				{
					json = FetchText(bootstrapUrl);
					if (!string.IsNullOrWhiteSpace(json))
					{
						source = bootstrapUrl;
						WriteCache(json);
					}
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				if (string.IsNullOrWhiteSpace(json))
				{
					json = ReadCache();
					if (!string.IsNullOrWhiteSpace(json))
					{
						source = result.CachePath;
						result.UsedCache = true;
					}
				}
				if (string.IsNullOrWhiteSpace(json))
				{
					result.Message = "No deployment bootstrap was available.";
					RememberResult(result, currentUser, projectIdentityKey);
					return result;
				}
				FamilyBrowserDeploymentBootstrap bootstrap = DataContractJsonTextStore.Load<FamilyBrowserDeploymentBootstrap>(json);
				result.Source = source;
				Apply(workspaceRoot, currentUser, bootstrap, result);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				result.Message = "Deployment bootstrap was skipped: " + ex2.Message;
				ProjectData.ClearProjectError();
			}
			RememberResult(result, currentUser, projectIdentityKey);
			return result;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static FamilyBrowserDeploymentBootstrapSecurityRefreshResult RefreshSecurityFromHomepage(string workspaceRoot, string currentUser, FamilyBrowserDeploymentProjectIdentity projectIdentity = null)
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			FamilyBrowserDeploymentBootstrapSecurityRefreshResult result = new FamilyBrowserDeploymentBootstrapSecurityRefreshResult
			{
				CheckedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
				RefreshMinutes = 30
			};
			try
			{
				string bootstrapUrl = (result.BootstrapUrl = ResolveBootstrapUrl(projectIdentity));
				if (IsBootstrapDisabledValue(bootstrapUrl))
				{
					result.Message = "Homepage security refresh is disabled.";
					return result;
				}
				string json = FetchText(AddNoCacheQuery(bootstrapUrl));
				if (string.IsNullOrWhiteSpace(json))
				{
					result.Message = "Homepage security bootstrap was not reachable.";
					return result;
				}
				result.Source = bootstrapUrl;
				WriteCache(json);
				FamilyBrowserDeploymentBootstrap bootstrap = DataContractJsonTextStore.Load<FamilyBrowserDeploymentBootstrap>(json);
				if (bootstrap == null)
				{
					result.Message = "Homepage security bootstrap was empty.";
					return result;
				}
				if (bootstrap.Disabled)
				{
					result.Message = "Homepage security bootstrap is disabled.";
					return result;
				}
				result.RefreshMinutes = ((bootstrap.RefreshMinutes <= 0) ? 30 : bootstrap.RefreshMinutes);
				if (!HasSecurityValues(bootstrap.Security))
				{
					result.Success = true;
					result.Message = "Homepage security settings were not configured.";
					return result;
				}
				FamilyBrowserStandardPolicy policy = FamilyBrowserStandardPolicyStore.LoadOrCreate(workspaceRoot, currentUser);
				if (policy.Security == null)
				{
					policy.Security = FamilyBrowserSecurityPolicy.CreateDefault();
				}
				string beforeSignature = BuildSecurityRefreshSignature(policy.Security);
				ApplySecurity(policy, bootstrap.Security);
				string afterSignature = BuildSecurityRefreshSignature(policy.Security);
				result.Changed = !string.Equals(beforeSignature, afterSignature, StringComparison.Ordinal);
				result.Success = true;
				result.AdminProfileKeywords = (from x in policy.Security.AdminProfileKeywords ?? new List<string>()
					where !string.IsNullOrWhiteSpace(x)
					select x.Trim()).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
				if (result.Changed)
				{
					policy.Security.LastUpdatedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
					policy.Security.LastUpdatedBy = currentUser ?? string.Empty;
					FamilyBrowserStandardPolicyStore.Save(workspaceRoot, policy, currentUser);
					result.Message = "Homepage security refreshed.";
				}
				else
				{
					result.Message = "Homepage security unchanged.";
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				result.Message = "Homepage security refresh was skipped: " + ex2.Message;
				ProjectData.ClearProjectError();
			}
			return result;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static string SetBootstrapUrl(string bootstrapUrl, string currentUser)
	{
		string storedPath = FamilyBrowserMachineConfigStore.SetDeploymentBootstrapUrl(bootstrapUrl, currentUser);
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			_lastCheckedUtc = DateTime.MinValue;
			_lastResult = null;
			_lastProjectIdentityKey = string.Empty;
			return storedPath;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static string DisableBootstrap(string currentUser)
	{
		return SetBootstrapUrl("disabled", currentUser);
	}

	public static bool IsBootstrapDisabled(string bootstrapUrl)
	{
		return IsBootstrapDisabledValue(bootstrapUrl);
	}

	public static string GetBootstrapUrlPath()
	{
		return "(runtime only - default homepage bootstrap URL)";
	}

	public static FamilyBrowserDeploymentBootstrapProfileIndexResult ListBootstrapProfiles(bool force = false)
	{
		FamilyBrowserDeploymentBootstrapProfileIndexResult result = new FamilyBrowserDeploymentBootstrapProfileIndexResult
		{
			IndexUrl = ResolveBootstrapIndexUrl()
		};
		FamilyBrowserDeploymentBootstrapProfileIndexResult ListBootstrapProfiles;
		try
		{
			string json = FetchText(result.IndexUrl);
			if (string.IsNullOrWhiteSpace(json))
			{
				result.Message = "No bootstrap profile index was available.";
				ListBootstrapProfiles = result;
			}
			else
			{
				FamilyBrowserDeploymentBootstrapProfileIndex index = DataContractJsonTextStore.Load<FamilyBrowserDeploymentBootstrapProfileIndex>(json);
				if (index == null)
				{
					result.Message = "Bootstrap profile index was empty.";
					ListBootstrapProfiles = result;
				}
				else
				{
					result.Success = true;
					result.Source = result.IndexUrl;
					result.DefaultProfileId = (index.DefaultProfileId ?? string.Empty).Trim();
					result.Profiles = NormalizeProfiles(index.Profiles, result.IndexUrl);
					result.ProjectRules = (index.ProjectRules ?? new List<FamilyBrowserDeploymentBootstrapProjectRule>()).Where([SpecialName] (FamilyBrowserDeploymentBootstrapProjectRule x) => x != null).ToList();
					result.Message = "Bootstrap profile index loaded.";
					ListBootstrapProfiles = result;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result.Message = "Bootstrap profile index was skipped: " + ex2.Message;
			ListBootstrapProfiles = result;
			ProjectData.ClearProjectError();
		}
		return ListBootstrapProfiles;
	}

	private static void Apply(string workspaceRoot, string currentUser, FamilyBrowserDeploymentBootstrap bootstrap, FamilyBrowserDeploymentBootstrapResult result)
	{
		if (bootstrap == null)
		{
			result.Message = "Deployment bootstrap was empty.";
			return;
		}
		if (bootstrap.Disabled)
		{
			result.Message = "Deployment bootstrap is disabled.";
			return;
		}
		result.RefreshMinutes = ((bootstrap.RefreshMinutes <= 0) ? 30 : bootstrap.RefreshMinutes);
		bool policyChanged = false;
		bool hasManagedPolicyTarget = BootstrapHasManagedPolicyTarget(bootstrap);
		string managedPolicyPath = ResolveManagedPolicyPath(bootstrap, result);
		string managedRootPath = ResolveManagedRootFromPolicyPath(managedPolicyPath);
		if (!string.IsNullOrWhiteSpace(managedPolicyPath))
		{
			FamilyBrowserMachineConfigStore.SetManagedPolicyPath(managedPolicyPath, currentUser);
			PrepareManagedPolicyFolders(managedPolicyPath, bootstrap.RequestStore);
			result.ManagedPolicyPath = managedPolicyPath;
			result.Applied = true;
		}
		else if (hasManagedPolicyTarget)
		{
			FamilyBrowserMachineConfigStore.ClearManagedPolicyPath(currentUser);
			result.Applied = true;
		}
		if (bootstrap.SkipPolicyWrite)
		{
			result.Message = BuildAppliedMessage(bootstrap, result, policyValuesChecked: false);
			return;
		}
		FamilyBrowserStandardPolicy policy = FamilyBrowserStandardPolicyStore.LoadOrCreate(workspaceRoot, currentUser);
		if (!string.IsNullOrWhiteSpace(bootstrap.StandardMode))
		{
			if (string.Equals(FamilyBrowserPolicyKey.Normalize(bootstrap.StandardMode), "integrated", StringComparison.OrdinalIgnoreCase))
			{
				policy.Mode = "Integrated";
			}
			else
			{
				policy.Mode = "DisciplineSeparated";
			}
			policyChanged = true;
		}
		if (bootstrap.RequestStore != null)
		{
			string storeMode = (bootstrap.RequestStore.Mode ?? string.Empty).Trim();
			string storePath = ResolveCandidatePath(bootstrap.RequestStore.Path, bootstrap.RequestStore.PathCandidates, BootstrapPathTargetKind.Directory);
			if (!string.IsNullOrWhiteSpace(storePath) && !IsAllowedManagedRuntimePath(storePath))
			{
				storePath = string.Empty;
			}
			string endpoint = Expand(bootstrap.RequestStore.Endpoint);
			bool num = !string.IsNullOrWhiteSpace(storePath) || !string.IsNullOrWhiteSpace(endpoint);
			bool isLocalStoreMode = string.Equals(FamilyBrowserPolicyKey.Normalize(storeMode), FamilyBrowserPolicyKey.Normalize("Local"), StringComparison.OrdinalIgnoreCase);
			if (num || isLocalStoreMode)
			{
				FamilyBrowserStandardPolicyStore.SetRequestStore(workspaceRoot, policy, storeMode, storePath, endpoint, currentUser);
				policy = FamilyBrowserStandardPolicyStore.LoadOrCreate(workspaceRoot, currentUser);
				result.RequestStorePath = (string.IsNullOrWhiteSpace(storePath) ? endpoint : storePath);
				result.Applied = true;
				policyChanged = false;
			}
			else if (HasRequestStoreTarget(bootstrap.RequestStore))
			{
				FamilyBrowserStandardPolicyStore.SetRequestStore(workspaceRoot, policy, "Local", string.Empty, string.Empty, currentUser);
				policy = FamilyBrowserStandardPolicyStore.LoadOrCreate(workspaceRoot, currentUser);
				result.Applied = true;
				policyChanged = false;
			}
		}
		if (string.IsNullOrWhiteSpace(result.RequestStorePath) && !string.IsNullOrWhiteSpace(managedRootPath))
		{
			string defaultRequestPath = Path.Combine(managedRootPath, "Requests");
			if (IsAllowedManagedRuntimePath(defaultRequestPath))
			{
				try
				{
					Directory.CreateDirectory(defaultRequestPath);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
			FamilyBrowserStandardPolicyStore.SetRequestStore(workspaceRoot, policy, "NetworkShare", defaultRequestPath, string.Empty, currentUser);
			policy = FamilyBrowserStandardPolicyStore.LoadOrCreate(workspaceRoot, currentUser);
			result.RequestStorePath = defaultRequestPath;
			result.Applied = true;
			policyChanged = false;
		}
		if (bootstrap.StandardLibraries != null && bootstrap.StandardLibraries.Count > 0)
		{
			ApplyStandardLibraries(policy, bootstrap.StandardLibraries);
			policyChanged = true;
			result.Applied = true;
		}
		if (HasSecurityValues(bootstrap.Security))
		{
			ApplySecurity(policy, bootstrap.Security);
			policyChanged = true;
			result.Applied = true;
		}
		if (policyChanged)
		{
			FamilyBrowserStandardPolicyStore.Save(workspaceRoot, policy, currentUser);
		}
		result.Message = BuildAppliedMessage(bootstrap, result, policyValuesChecked: true);
	}

	private static string ResolveManagedPolicyPath(FamilyBrowserDeploymentBootstrap bootstrap, FamilyBrowserDeploymentBootstrapResult result)
	{
		List<ManagedPolicyCandidate> candidates = BuildManagedPolicyCandidates(bootstrap);
		if (candidates.Count == 0)
		{
			return string.Empty;
		}
		string preferredPath = FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath();
		if (!string.IsNullOrWhiteSpace(preferredPath))
		{
			string expandedPreferred = Expand(preferredPath);
			foreach (ManagedPolicyCandidate candidate in candidates)
			{
				if (SameText(candidate.PolicyPath, expandedPreferred) && ProbeManagedPolicyCandidate(candidate))
				{
					return candidate.PolicyPath;
				}
			}
		}
		foreach (ManagedPolicyCandidate candidate2 in candidates)
		{
			if (ProbeManagedPolicyCandidate(candidate2))
			{
				return candidate2.PolicyPath;
			}
		}
		if (result != null)
		{
			result.ManagedPolicyPathIssue = "management folder unavailable: " + string.Join(" / ", candidates.Select([SpecialName] (ManagedPolicyCandidate x) => x.DisplayPath));
		}
		return string.Empty;
	}

	private static List<ManagedPolicyCandidate> BuildManagedPolicyCandidates(FamilyBrowserDeploymentBootstrap bootstrap)
	{
		List<ManagedPolicyCandidate> candidates = new List<ManagedPolicyCandidate>();
		if (bootstrap == null)
		{
			return candidates;
		}
		AddManagedRootCandidate(candidates, bootstrap.ManagedRootPath);
		if (bootstrap.ManagedRootPathCandidates != null)
		{
			foreach (string path in bootstrap.ManagedRootPathCandidates)
			{
				AddManagedRootCandidate(candidates, path);
			}
		}
		AddManagedPolicyCandidate(candidates, bootstrap.ManagedPolicyPath);
		if (bootstrap.ManagedPolicyPathCandidates != null)
		{
			foreach (string path2 in bootstrap.ManagedPolicyPathCandidates)
			{
				AddManagedPolicyCandidate(candidates, path2);
			}
		}
		return candidates;
	}

	private static void AddManagedRootCandidate(List<ManagedPolicyCandidate> candidates, string rawPath)
	{
		string displayPath = Expand(rawPath);
		if (!string.IsNullOrWhiteSpace(displayPath))
		{
			AddManagedCandidate(candidates, displayPath, ResolvePolicyPathFromManagedRoot(displayPath), isRoot: true);
		}
	}

	private static void AddManagedPolicyCandidate(List<ManagedPolicyCandidate> candidates, string rawPath)
	{
		string policyPath = Expand(rawPath);
		if (!string.IsNullOrWhiteSpace(policyPath))
		{
			AddManagedCandidate(candidates, policyPath, policyPath, isRoot: false);
		}
	}

	private static void AddManagedCandidate(List<ManagedPolicyCandidate> candidates, string displayPath, string policyPath, bool isRoot)
	{
		if (candidates != null && !string.IsNullOrWhiteSpace(policyPath) && IsAllowedManagedRuntimePath(policyPath) && !candidates.Any([SpecialName] (ManagedPolicyCandidate x) => SameText(x.PolicyPath, policyPath)))
		{
			candidates.Add(new ManagedPolicyCandidate
			{
				DisplayPath = (displayPath ?? string.Empty).Trim(),
				PolicyPath = (policyPath ?? string.Empty).Trim(),
				IsRoot = isRoot
			});
		}
	}

	private static string ResolvePolicyPathFromManagedRoot(string rootPath)
	{
		string value = Expand(rootPath);
		if (string.IsNullOrWhiteSpace(value))
		{
			return string.Empty;
		}
		try
		{
			if (string.Equals(Path.GetExtension(value), ".json", StringComparison.OrdinalIgnoreCase))
			{
				return value;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return Path.Combine(value, "Config", "standard-policy.json");
	}

	private static bool ProbeManagedPolicyCandidate(ManagedPolicyCandidate candidate)
	{
		if (candidate == null)
		{
			return false;
		}
		if (candidate.IsRoot && ProbePathAvailable(candidate.DisplayPath, BootstrapPathTargetKind.Directory))
		{
			return true;
		}
		return ProbePathAvailable(candidate.PolicyPath, BootstrapPathTargetKind.FileOrParentDirectory);
	}

	private static string ResolveManagedRootFromPolicyPath(string managedPolicyPath)
	{
		string value = Expand(managedPolicyPath);
		string ResolveManagedRootFromPolicyPath;
		if (string.IsNullOrWhiteSpace(value))
		{
			ResolveManagedRootFromPolicyPath = string.Empty;
		}
		else
		{
			try
			{
				string policyFolder = Path.GetDirectoryName(value);
				if (string.IsNullOrWhiteSpace(policyFolder))
				{
					ResolveManagedRootFromPolicyPath = string.Empty;
				}
				else
				{
					DirectoryInfo policyFolderInfo = new DirectoryInfo(policyFolder);
					ResolveManagedRootFromPolicyPath = ((!string.Equals(policyFolderInfo.Name, "Config", StringComparison.OrdinalIgnoreCase) || policyFolderInfo.Parent == null) ? policyFolder : policyFolderInfo.Parent.FullName);
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ResolveManagedRootFromPolicyPath = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return ResolveManagedRootFromPolicyPath;
	}

	private static bool IsAllowedManagedRuntimePath(string path)
	{
		string value = Expand(path);
		bool IsAllowedManagedRuntimePath;
		if (string.IsNullOrWhiteSpace(value))
		{
			IsAllowedManagedRuntimePath = false;
		}
		else
		{
			try
			{
				if (value.StartsWith("\\\\", StringComparison.Ordinal))
				{
					IsAllowedManagedRuntimePath = true;
				}
				else
				{
					string root = Path.GetPathRoot(value);
					if (string.IsNullOrWhiteSpace(root))
					{
						IsAllowedManagedRuntimePath = false;
					}
					else
					{
						string windowsRoot = Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.Windows));
						if (!string.IsNullOrWhiteSpace(windowsRoot) && string.Equals(root, windowsRoot, StringComparison.OrdinalIgnoreCase))
						{
							IsAllowedManagedRuntimePath = false;
						}
						else
						{
							string commonAppData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
							string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
							string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
							string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
							string tempPath = Path.GetTempPath();
							IsAllowedManagedRuntimePath = !IsSameOrChildPath(value, commonAppData) && !IsSameOrChildPath(value, appData) && !IsSameOrChildPath(value, localAppData) && !IsSameOrChildPath(value, userProfile) && !IsSameOrChildPath(value, tempPath);
						}
					}
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				IsAllowedManagedRuntimePath = false;
				ProjectData.ClearProjectError();
			}
		}
		return IsAllowedManagedRuntimePath;
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

	private static void PrepareManagedPolicyFolders(string managedPolicyPath, FamilyBrowserBootstrapRequestStore requestStore)
	{
		if (string.IsNullOrWhiteSpace(managedPolicyPath) || !IsAllowedManagedRuntimePath(managedPolicyPath))
		{
			return;
		}
		try
		{
			string policyFolder = Path.GetDirectoryName(managedPolicyPath);
			if (string.IsNullOrWhiteSpace(policyFolder))
			{
				return;
			}
			Directory.CreateDirectory(policyFolder);
			DirectoryInfo rootFolder = Directory.GetParent(policyFolder);
			if (rootFolder != null)
			{
				Directory.CreateDirectory(Path.Combine(rootFolder.FullName, "RevitVersions"));
				Directory.CreateDirectory(Path.Combine(rootFolder.FullName, "StandardLists"));
				Directory.CreateDirectory(Path.Combine(rootFolder.FullName, "Requests"));
				Directory.CreateDirectory(Path.Combine(rootFolder.FullName, "Logs"));
				Directory.CreateDirectory(Path.Combine(rootFolder.FullName, "Diagnostics"));
				Directory.CreateDirectory(Path.Combine(rootFolder.FullName, "OperationLogs"));
				Directory.CreateDirectory(Path.Combine(rootFolder.FullName, "StandardChangeCandidates"));
			}
			if (requestStore != null)
			{
				string requestPath = Expand(requestStore.Path);
				if (!string.IsNullOrWhiteSpace(requestPath) && IsAllowedManagedRuntimePath(requestPath))
				{
					Directory.CreateDirectory(requestPath);
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void ApplyStandardLibraries(FamilyBrowserStandardPolicy policy, IEnumerable<FamilyBrowserBootstrapStandardLibrary> libraries)
	{
		if (policy.DisciplineLibraries == null)
		{
			policy.DisciplineLibraries = new List<FamilyBrowserStandardLibrarySlot>();
		}
		_Closure_0024__34_002D0 closure_0024__34_002D = default(_Closure_0024__34_002D0);
		foreach (FamilyBrowserBootstrapStandardLibrary library in libraries.Where([SpecialName] (FamilyBrowserBootstrapStandardLibrary x) => x != null))
		{
			string rawDiscipline = (library.Discipline ?? string.Empty).Trim();
			string displayName = (library.DisplayName ?? string.Empty).Trim();
			if (string.IsNullOrWhiteSpace(rawDiscipline))
			{
				rawDiscipline = displayName;
			}
			if (string.IsNullOrWhiteSpace(displayName))
			{
				displayName = rawDiscipline;
			}
			if (string.IsNullOrWhiteSpace(rawDiscipline))
			{
				rawDiscipline = "Other";
			}
			string discipline = FamilyBrowserStandardPolicyStore.ResolveDisciplineKey(rawDiscipline);
			if (string.IsNullOrWhiteSpace(discipline))
			{
				discipline = rawDiscipline;
			}
			FamilyBrowserStandardLibrarySlot slot = null;
			if (string.Equals(FamilyBrowserPolicyKey.Normalize(discipline), FamilyBrowserPolicyKey.Normalize("Integrated"), StringComparison.OrdinalIgnoreCase))
			{
				if (policy.IntegratedLibrary == null)
				{
					policy.IntegratedLibrary = FamilyBrowserStandardLibrarySlot.CreateIntegrated();
				}
				slot = policy.IntegratedLibrary;
				policy.Mode = "Integrated";
			}
			else
			{
				closure_0024__34_002D = new _Closure_0024__34_002D0(closure_0024__34_002D);
				closure_0024__34_002D._0024VB_0024Local_key = FamilyBrowserPolicyKey.Normalize(discipline);
				slot = policy.DisciplineLibraries.FirstOrDefault(closure_0024__34_002D._Lambda_0024__1);
				if (slot == null)
				{
					slot = FamilyBrowserStandardLibrarySlot.CreateDiscipline(discipline, displayName);
					policy.DisciplineLibraries.Add(slot);
				}
			}
			slot.Discipline = discipline;
			slot.DisplayName = (string.IsNullOrWhiteSpace(displayName) ? discipline : displayName);
			slot.Enabled = !library.Disabled;
			string standardRvtPath = ResolveCandidatePath(library.StandardRvtPath, library.StandardRvtPathCandidates, BootstrapPathTargetKind.File);
			if (!string.IsNullOrWhiteSpace(standardRvtPath))
			{
				slot.StandardRvtPath = standardRvtPath;
			}
			string registrationPath = ResolveCandidatePath(library.RegistrationPath, library.RegistrationPathCandidates, BootstrapPathTargetKind.FileOrParentDirectory);
			if (!string.IsNullOrWhiteSpace(registrationPath))
			{
				slot.RegistrationPath = registrationPath;
			}
			string snapshotPath = ResolveCandidatePath(library.SnapshotPath, library.SnapshotPathCandidates, BootstrapPathTargetKind.FileOrParentDirectory);
			if (!string.IsNullOrWhiteSpace(snapshotPath))
			{
				slot.SnapshotPath = snapshotPath;
			}
			string standardListPath = ResolveCandidatePath(library.StandardListPath, library.StandardListPathCandidates, BootstrapPathTargetKind.File);
			if (!string.IsNullOrWhiteSpace(standardListPath))
			{
				slot.StandardListPath = standardListPath;
			}
			string standardListSheetName = (library.StandardListSheetName ?? string.Empty).Trim();
			if (!string.IsNullOrWhiteSpace(standardListSheetName))
			{
				slot.StandardListSheetName = standardListSheetName;
			}
			string sourceId = (library.SourceId ?? string.Empty).Trim();
			if (!string.IsNullOrWhiteSpace(sourceId))
			{
				slot.SourceId = sourceId;
			}
		}
	}

	private static bool BootstrapHasManagedPolicyTarget(FamilyBrowserDeploymentBootstrap bootstrap)
	{
		if (bootstrap == null)
		{
			return false;
		}
		if (!string.IsNullOrWhiteSpace(bootstrap.ManagedRootPath) || !string.IsNullOrWhiteSpace(bootstrap.ManagedPolicyPath))
		{
			return true;
		}
		if (bootstrap.ManagedRootPathCandidates != null && bootstrap.ManagedRootPathCandidates.Any([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)))
		{
			return true;
		}
		return bootstrap.ManagedPolicyPathCandidates != null && bootstrap.ManagedPolicyPathCandidates.Any([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x));
	}

	private static bool HasRequestStoreTarget(FamilyBrowserBootstrapRequestStore requestStore)
	{
		if (requestStore == null)
		{
			return false;
		}
		if (!string.IsNullOrWhiteSpace(requestStore.Path) || !string.IsNullOrWhiteSpace(requestStore.Endpoint))
		{
			return true;
		}
		return requestStore.PathCandidates != null && requestStore.PathCandidates.Any([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x));
	}

	private static bool HasSecurityValues(FamilyBrowserBootstrapSecurity security)
	{
		if (security == null)
		{
			return false;
		}
		return (security.AdminUsers != null && security.AdminUsers.Count > 0) || (security.AdminProfileKeywords != null && security.AdminProfileKeywords.Count > 0) || (security.RequestApproverUsers != null && security.RequestApproverUsers.Count > 0) || (security.ReadOnlyUsers != null && security.ReadOnlyUsers.Count > 0);
	}

	private static string BuildSecurityRefreshSignature(FamilyBrowserSecurityPolicy security)
	{
		if (security == null)
		{
			security = FamilyBrowserSecurityPolicy.CreateDefault();
		}
		List<string> parts = new List<string> { "security" };
		AppendSecurityList(parts, security.AdminUsers);
		AppendSecurityList(parts, security.AdminProfileKeywords);
		AppendSecurityList(parts, security.RequestApproverUsers);
		AppendSecurityList(parts, security.ReadOnlyUsers);
		parts.Add(security.AllowUnlistedUsersAsModelers ? "true" : "false");
		parts.Add(security.AllowModelersToLoadFamilies ? "true" : "false");
		parts.Add(security.AllowModelersToApplySystemTypes ? "true" : "false");
		parts.Add(security.AllowModelersToSubmitRequests ? "true" : "false");
		return string.Join("|", parts.Select([SpecialName] (string x) => (x ?? string.Empty).Replace("|", "/")));
	}

	private static void AppendSecurityList(List<string> parts, IEnumerable<string> values)
	{
		parts?.Add(string.Join(",", (from x in values ?? new List<string>()
			where !string.IsNullOrWhiteSpace(x)
			select x.Trim()).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase)));
	}

	private static void ApplySecurity(FamilyBrowserStandardPolicy policy, FamilyBrowserBootstrapSecurity security)
	{
		if (policy.Security == null)
		{
			policy.Security = FamilyBrowserSecurityPolicy.CreateDefault();
		}
		if (security.AdminProfileKeywords != null && security.AdminProfileKeywords.Count > 0)
		{
			policy.Security.AdminUsers = FamilyBrowserSecurityPolicyService.ParseUserList(string.Join(";", security.AdminUsers ?? new List<string>()) ?? string.Empty);
			policy.Security.AdminProfileKeywords = FamilyBrowserSecurityPolicyService.ParseUserList(string.Join(";", security.AdminProfileKeywords));
		}
		else if (security.AdminUsers != null && security.AdminUsers.Count > 0)
		{
			policy.Security.AdminUsers = FamilyBrowserSecurityPolicyService.ParseUserList(string.Join(";", security.AdminUsers));
		}
		if (security.RequestApproverUsers != null && security.RequestApproverUsers.Count > 0)
		{
			policy.Security.RequestApproverUsers = FamilyBrowserSecurityPolicyService.ParseUserList(string.Join(";", security.RequestApproverUsers));
		}
		if (security.ReadOnlyUsers != null && security.ReadOnlyUsers.Count > 0)
		{
			policy.Security.ReadOnlyUsers = FamilyBrowserSecurityPolicyService.ParseUserList(string.Join(";", security.ReadOnlyUsers));
		}
	}

	private static string BuildAppliedMessage(FamilyBrowserDeploymentBootstrap bootstrap, FamilyBrowserDeploymentBootstrapResult result, bool policyValuesChecked)
	{
		List<string> parts = new List<string>();
		if (!string.IsNullOrWhiteSpace(bootstrap.Version))
		{
			parts.Add("version " + bootstrap.Version);
		}
		if (!string.IsNullOrWhiteSpace(result.ManagedPolicyPath))
		{
			parts.Add("managed policy connected");
		}
		else if (!string.IsNullOrWhiteSpace(result.ManagedPolicyPathIssue))
		{
			parts.Add(result.ManagedPolicyPathIssue);
		}
		if (!string.IsNullOrWhiteSpace(result.RequestStorePath))
		{
			parts.Add("request store connected");
		}
		if (!policyValuesChecked)
		{
			parts.Add("policy write skipped");
		}
		if (bootstrap.StandardLibraries != null && bootstrap.StandardLibraries.Count > 0)
		{
			parts.Add(bootstrap.StandardLibraries.Count.ToString(CultureInfo.InvariantCulture) + " standard target(s)");
		}
		if (result.UsedCache)
		{
			parts.Add("cache");
		}
		if (parts.Count == 0)
		{
			parts.Add("checked");
		}
		string adminMessage = (bootstrap.Message ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(adminMessage))
		{
			return "Deployment bootstrap " + string.Join(" / ", parts);
		}
		return adminMessage + " / " + string.Join(" / ", parts);
	}

	private static string ResolveBootstrapUrl(FamilyBrowserDeploymentProjectIdentity projectIdentity)
	{
		string stored = FamilyBrowserMachineConfigStore.ResolveDeploymentBootstrapUrl();
		if (!string.IsNullOrWhiteSpace(stored) && IsBootstrapDisabledValue(stored))
		{
			return "disabled";
		}
		string envValue = Environment.GetEnvironmentVariable("KKY_FAMILY_BROWSER_BOOTSTRAP_URL");
		if (!string.IsNullOrWhiteSpace(envValue))
		{
			if (IsBootstrapDisabledValue(envValue))
			{
				return "disabled";
			}
			return envValue.Trim();
		}
		if (!string.IsNullOrWhiteSpace(stored))
		{
			return stored;
		}
		string indexedUrl = ResolveBootstrapUrlFromIndex(projectIdentity);
		if (!string.IsNullOrWhiteSpace(indexedUrl))
		{
			return indexedUrl;
		}
		return "https://update.zerokky.com/family-browser/bootstrap.json";
	}

	private static string ResolveBootstrapIndexUrl()
	{
		string envValue = Environment.GetEnvironmentVariable("KKY_FAMILY_BROWSER_BOOTSTRAP_INDEX_URL");
		if (!string.IsNullOrWhiteSpace(envValue))
		{
			return envValue.Trim();
		}
		return "https://update.zerokky.com/family-browser/bootstrap-index.json";
	}

	private static string ResolveBootstrapUrlFromIndex(FamilyBrowserDeploymentProjectIdentity projectIdentity)
	{
		FamilyBrowserDeploymentBootstrapProfileIndexResult indexResult = ListBootstrapProfiles();
		if (indexResult == null || !indexResult.Success || indexResult.Profiles == null || indexResult.Profiles.Count == 0)
		{
			return string.Empty;
		}
		Dictionary<string, FamilyBrowserDeploymentBootstrapProfile> profilesById = (from x in indexResult.Profiles
			where x != null && !x.Disabled && !string.IsNullOrWhiteSpace(x.Id)
			group x by FamilyBrowserPolicyKey.Normalize(x.Id)).ToDictionary<IGrouping<string, FamilyBrowserDeploymentBootstrapProfile>, string, FamilyBrowserDeploymentBootstrapProfile>([SpecialName] (IGrouping<string, FamilyBrowserDeploymentBootstrapProfile> g) => g.Key, [SpecialName] (IGrouping<string, FamilyBrowserDeploymentBootstrapProfile> g) => g.First(), StringComparer.Ordinal);
		if (projectIdentity != null && indexResult.ProjectRules != null)
		{
			foreach (FamilyBrowserDeploymentBootstrapProjectRule rule in from x in indexResult.ProjectRules
				where x != null && !x.Disabled
				orderby x.Priority descending
				select x)
			{
				if (RuleMatchesProject(rule, projectIdentity))
				{
					string key = FamilyBrowserPolicyKey.Normalize(rule.ProfileId);
					if (profilesById.ContainsKey(key))
					{
						return profilesById[key].Url;
					}
				}
			}
		}
		string defaultKey = FamilyBrowserPolicyKey.Normalize(indexResult.DefaultProfileId);
		if (profilesById.ContainsKey(defaultKey))
		{
			return profilesById[defaultKey].Url;
		}
		FamilyBrowserDeploymentBootstrapProfile firstProfile = indexResult.Profiles.FirstOrDefault([SpecialName] (FamilyBrowserDeploymentBootstrapProfile x) => x != null && !x.Disabled && !string.IsNullOrWhiteSpace(x.Url));
		if (firstProfile != null)
		{
			return firstProfile.Url;
		}
		return string.Empty;
	}

	private static List<FamilyBrowserDeploymentBootstrapProfile> NormalizeProfiles(IEnumerable<FamilyBrowserDeploymentBootstrapProfile> profiles, string indexSource)
	{
		List<FamilyBrowserDeploymentBootstrapProfile> normalized = new List<FamilyBrowserDeploymentBootstrapProfile>();
		if (profiles == null)
		{
			return normalized;
		}
		foreach (FamilyBrowserDeploymentBootstrapProfile profile in profiles.Where([SpecialName] (FamilyBrowserDeploymentBootstrapProfile x) => x != null))
		{
			FamilyBrowserDeploymentBootstrapProfile item = new FamilyBrowserDeploymentBootstrapProfile
			{
				Id = (profile.Id ?? string.Empty).Trim(),
				Name = (profile.Name ?? string.Empty).Trim(),
				Description = (profile.Description ?? string.Empty).Trim(),
				Url = ResolveProfileUrl(indexSource, profile.Url),
				Disabled = profile.Disabled
			};
			if (string.IsNullOrWhiteSpace(item.Id))
			{
				item.Id = FamilyBrowserPolicyKey.Normalize(item.Name);
			}
			if (string.IsNullOrWhiteSpace(item.Name))
			{
				item.Name = item.Id;
			}
			if (!string.IsNullOrWhiteSpace(item.Url))
			{
				normalized.Add(item);
			}
		}
		return normalized;
	}

	private static string ResolveProfileUrl(string indexSource, string rawUrl)
	{
		string value = (rawUrl ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(value))
		{
			return string.Empty;
		}
		Uri absolute = null;
		if (Uri.TryCreate(value, UriKind.Absolute, out absolute))
		{
			return value;
		}
		if (Path.IsPathRooted(value))
		{
			return value;
		}
		Uri baseUri = null;
		if (Uri.TryCreate(indexSource, UriKind.Absolute, out baseUri) && (string.Equals(baseUri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) || string.Equals(baseUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) || string.Equals(baseUri.Scheme, Uri.UriSchemeFile, StringComparison.OrdinalIgnoreCase)))
		{
			return new Uri(baseUri, value).ToString();
		}
		try
		{
			string folder = Path.GetDirectoryName(indexSource);
			if (!string.IsNullOrWhiteSpace(folder))
			{
				return Path.GetFullPath(Path.Combine(folder, value));
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return value;
	}

	private static string ResolveCandidatePath(string primary, IEnumerable<string> candidates, BootstrapPathTargetKind targetKind)
	{
		List<string> candidateList = new List<string>();
		if (!string.IsNullOrWhiteSpace(primary))
		{
			candidateList.Add(primary);
		}
		if (candidates != null)
		{
			candidateList.AddRange(candidates.Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)));
		}
		if (candidateList.Count == 0)
		{
			return string.Empty;
		}
		List<string> uniqueCandidates = (from x in candidateList
			select Expand(x) into x
			where !string.IsNullOrWhiteSpace(x)
			select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).ToList();
		if (candidates == null || !candidates.Any([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)))
		{
			return uniqueCandidates.FirstOrDefault();
		}
		foreach (string candidate in uniqueCandidates)
		{
			if (ProbePathAvailable(candidate, targetKind))
			{
				return candidate;
			}
		}
		return string.Empty;
	}

	private static bool ProbePathAvailable(string path, BootstrapPathTargetKind targetKind)
	{
		bool ProbePathAvailable;
		if (string.IsNullOrWhiteSpace(path))
		{
			ProbePathAvailable = false;
		}
		else
		{
			Uri uri = null;
			if (Uri.TryCreate(path, UriKind.Absolute, out uri) && !string.Equals(uri.Scheme, Uri.UriSchemeFile, StringComparison.OrdinalIgnoreCase))
			{
				ProbePathAvailable = true;
			}
			else
			{
				Func<bool> check = [SpecialName] () =>
				{
					bool result;
					try
					{
						switch (targetKind)
						{
						case BootstrapPathTargetKind.Directory:
							result = Directory.Exists(path);
							break;
						case BootstrapPathTargetKind.File:
							result = File.Exists(path);
							break;
						default:
							if (File.Exists(path) || Directory.Exists(path))
							{
								result = true;
							}
							else
							{
								string directoryName = Path.GetDirectoryName(path);
								result = !string.IsNullOrWhiteSpace(directoryName) && Directory.Exists(directoryName);
							}
							break;
						}
					}
					catch (Exception projectError2)
					{
						ProjectData.SetProjectError(projectError2);
						result = false;
						ProjectData.ClearProjectError();
					}
					return result;
				};
				try
				{
					Task<bool> probeTask = Task.Run(check);
					ProbePathAvailable = probeTask.Wait(800) && probeTask.Result;
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProbePathAvailable = false;
					ProjectData.ClearProjectError();
				}
			}
		}
		return ProbePathAvailable;
	}

	private static bool RuleMatchesProject(FamilyBrowserDeploymentBootstrapProjectRule rule, FamilyBrowserDeploymentProjectIdentity projectIdentity)
	{
		if (rule == null || projectIdentity == null)
		{
			return false;
		}
		bool hasCriteria = false;
		if (!string.IsNullOrWhiteSpace(rule.ProjectNameContains))
		{
			hasCriteria = true;
			if (!ContainsText(projectIdentity.ProjectTitle, rule.ProjectNameContains))
			{
				return false;
			}
		}
		if (!string.IsNullOrWhiteSpace(rule.ModelPathContains))
		{
			hasCriteria = true;
			if (!ContainsText(projectIdentity.ModelPath, rule.ModelPathContains))
			{
				return false;
			}
		}
		if (!string.IsNullOrWhiteSpace(rule.CentralPathContains))
		{
			hasCriteria = true;
			if (!ContainsText(projectIdentity.CentralPath, rule.CentralPathContains))
			{
				return false;
			}
		}
		if (!string.IsNullOrWhiteSpace(rule.MatchValue))
		{
			hasCriteria = true;
			switch (FamilyBrowserPolicyKey.Normalize(rule.MatchMode))
			{
			case "exactprojectname":
				if (!SameText(projectIdentity.ProjectTitle, rule.MatchValue))
				{
					return false;
				}
				break;
			case "exactmodelpath":
				if (!SameText(projectIdentity.ModelPath, rule.MatchValue))
				{
					return false;
				}
				break;
			case "exactcentralpath":
				if (!SameText(projectIdentity.CentralPath, rule.MatchValue))
				{
					return false;
				}
				break;
			case "centralpathcontains":
				if (!ContainsText(projectIdentity.CentralPath, rule.MatchValue))
				{
					return false;
				}
				break;
			case "modelpathcontains":
				if (!ContainsText(projectIdentity.ModelPath, rule.MatchValue))
				{
					return false;
				}
				break;
			default:
				if (!ContainsText(projectIdentity.ProjectTitle, rule.MatchValue) && !ContainsText(projectIdentity.ModelPath, rule.MatchValue) && !ContainsText(projectIdentity.CentralPath, rule.MatchValue))
				{
					return false;
				}
				break;
			}
		}
		return hasCriteria;
	}

	private static string BuildProjectIdentityKey(FamilyBrowserDeploymentProjectIdentity projectIdentity)
	{
		if (projectIdentity == null)
		{
			return string.Empty;
		}
		return string.Join("|", (projectIdentity.ProjectTitle ?? string.Empty).Trim(), (projectIdentity.ModelPath ?? string.Empty).Trim(), (projectIdentity.CentralPath ?? string.Empty).Trim());
	}

	private static bool SameText(string left, string right)
	{
		return string.Equals((left ?? string.Empty).Trim(), (right ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase);
	}

	private static bool ContainsText(string source, string value)
	{
		string sourceText = source ?? string.Empty;
		string needle = (value ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(needle))
		{
			return true;
		}
		return sourceText.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
	}

	private static bool IsBootstrapDisabledValue(string value)
	{
		string normalized = (value ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(normalized))
		{
			return false;
		}
		return string.Equals(normalized, "disabled", StringComparison.OrdinalIgnoreCase) || string.Equals(normalized, "off", StringComparison.OrdinalIgnoreCase) || string.Equals(normalized, "none", StringComparison.OrdinalIgnoreCase) || string.Equals(normalized, "local-test", StringComparison.OrdinalIgnoreCase);
	}

	private static string FetchText(string source)
	{
		if (string.IsNullOrWhiteSpace(source))
		{
			return string.Empty;
		}
		string expanded = Expand(source);
		Uri sourceUri = null;
		if (Uri.TryCreate(expanded, UriKind.Absolute, out sourceUri))
		{
			if (string.Equals(sourceUri.Scheme, Uri.UriSchemeFile, StringComparison.OrdinalIgnoreCase))
			{
				return File.ReadAllText(sourceUri.LocalPath, Encoding.UTF8);
			}
			if (string.Equals(sourceUri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) || string.Equals(sourceUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
			{
				ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
				HttpWebRequest obj = (HttpWebRequest)WebRequest.Create(sourceUri);
				obj.Method = "GET";
				obj.Timeout = 3000;
				obj.ReadWriteTimeout = 3000;
				obj.UserAgent = "KKY-FamilyBrowser";
				obj.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);
				obj.Headers[HttpRequestHeader.CacheControl] = "no-cache";
				obj.Headers[HttpRequestHeader.Pragma] = "no-cache";
				using HttpWebResponse response = (HttpWebResponse)obj.GetResponse();
				if (response.StatusCode != HttpStatusCode.OK)
				{
					return string.Empty;
				}
				using Stream stream = response.GetResponseStream();
				if (stream == null)
				{
					return string.Empty;
				}
				using StreamReader reader = new StreamReader(stream, Encoding.UTF8);
				return reader.ReadToEnd();
			}
		}
		if (File.Exists(expanded))
		{
			return File.ReadAllText(expanded, Encoding.UTF8);
		}
		return string.Empty;
	}

	private static string AddNoCacheQuery(string source)
	{
		if (string.IsNullOrWhiteSpace(source))
		{
			return string.Empty;
		}
		Uri sourceUri = null;
		if (!Uri.TryCreate(Expand(source), UriKind.Absolute, out sourceUri))
		{
			return source;
		}
		if (!string.Equals(sourceUri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) && !string.Equals(sourceUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
		{
			return source;
		}
		UriBuilder uriBuilder = new UriBuilder(sourceUri);
		string query = (uriBuilder.Query ?? string.Empty).TrimStart('?');
		string token = "_kkyfbSecurityRefresh=" + DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture);
		uriBuilder.Query = (string.IsNullOrWhiteSpace(query) ? token : (query + "&" + token));
		return uriBuilder.Uri.AbsoluteUri;
	}

	private static string ReadCache()
	{
		return _runtimeBootstrapCacheJson ?? string.Empty;
	}

	private static void WriteCache(string json)
	{
		if (!string.IsNullOrWhiteSpace(json))
		{
			_runtimeBootstrapCacheJson = json;
		}
	}

	private static string GetCachePath()
	{
		return "(runtime only - no local bootstrap cache file)";
	}

	private static string GetProgramDataRoot()
	{
		return string.Empty;
	}

	private static string Expand(string value)
	{
		return Environment.ExpandEnvironmentVariables((value ?? string.Empty).Trim());
	}

	private static FamilyBrowserDeploymentBootstrapResult NewResult()
	{
		return new FamilyBrowserDeploymentBootstrapResult
		{
			CheckedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			RefreshMinutes = 30
		};
	}

	private static void RememberResult(FamilyBrowserDeploymentBootstrapResult result, string currentUser, string projectIdentityKey)
	{
		_lastCheckedUtc = DateTime.UtcNow;
		_lastResult = result;
		_lastProjectIdentityKey = projectIdentityKey ?? string.Empty;
		try
		{
			FamilyBrowserMachineConfigStore.RecordBootstrapResult(result, currentUser);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}
}
