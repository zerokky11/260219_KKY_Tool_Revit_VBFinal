using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserOperationalReadinessService
{
	private FamilyBrowserOperationalReadinessService()
	{
	}

	public static FamilyBrowserOperationalReadinessReport Build(string workspaceRoot, FamilyBrowserStandardPolicy policy, string currentUser)
	{
		if (string.IsNullOrWhiteSpace(workspaceRoot))
		{
			workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
		}
		if (policy == null)
		{
			policy = FamilyBrowserStandardPolicyStore.LoadOrCreate(workspaceRoot, currentUser);
		}
		FamilyBrowserOperationalReadinessReport obj = new FamilyBrowserOperationalReadinessReport
		{
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			PolicyPath = FamilyBrowserStandardPolicyStore.GetPolicyPath(workspaceRoot)
		};
		FamilyBrowserStandardLibrarySlot slot = FamilyBrowserStandardPolicyStore.GetEffectiveSlot(policy);
		StandardLibraryRegistrationRecord registration = TryLoadRegistration(workspaceRoot, policy);
		obj.StandardTarget = ResolveStandardTarget(policy, slot);
		AddPolicyFileCheck(obj);
		AddManagedPolicyCheck(obj);
		AddStandardSourceChecks(obj, workspaceRoot, slot, registration);
		AddRequestStoreCheck(obj, workspaceRoot, policy);
		AddPermissionExcelCheck(obj, policy);
		AddSecurityChecks(obj, policy, currentUser);
		AddProjectPolicyCheck(obj, policy);
		FinalizeCounts(obj);
		return obj;
	}

	private static void AddPolicyFileCheck(FamilyBrowserOperationalReadinessReport report)
	{
		if (string.IsNullOrWhiteSpace(report.PolicyPath))
		{
			Add(report, "Policy", "Blocked", "Policy path is empty.", "Set a local or managed policy path.");
			return;
		}
		string? directoryName = Path.GetDirectoryName(report.PolicyPath);
		string detail = string.Empty;
		if (IsWritableFolder(directoryName, ref detail))
		{
			Add(report, "Policy", "Ready", "Policy file location is writable.", report.PolicyPath);
		}
		else
		{
			Add(report, "Policy", "Blocked", "Policy file location is not writable.", detail);
		}
	}

	private static void AddManagedPolicyCheck(FamilyBrowserOperationalReadinessReport report)
	{
		string managedPolicyPath = FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath();
		if (string.IsNullOrWhiteSpace(managedPolicyPath))
		{
			Add(report, "Deployment", "Warning", "Local test policy is active.", "Use a managed policy path for company deployment.");
		}
		else if (FamilyBrowserStandardPolicyStore.IsConfiguredManagedPolicyAvailable())
		{
			Add(report, "Deployment", "Ready", "Managed policy is connected.", managedPolicyPath);
		}
		else
		{
			Add(report, "Deployment", "Warning", "Managed policy path is unavailable on this PC.", "Refresh the homepage path and connect the first reachable managed folder. Configured path: " + managedPolicyPath);
		}
	}

	private static void AddStandardSourceChecks(FamilyBrowserOperationalReadinessReport report, string workspaceRoot, FamilyBrowserStandardLibrarySlot slot, StandardLibraryRegistrationRecord registration)
	{
		if (slot == null)
		{
			Add(report, "Standard RVT", "Blocked", "No active standard library slot is available.", "Choose discipline separated or integrated mode and register a standard RVT.");
			return;
		}
		if (!slot.Enabled)
		{
			Add(report, "Standard RVT", "Blocked", "The active standard library slot is disabled.", "Enable the slot or choose another trade.");
			return;
		}
		string standardPath = FirstNonEmpty(slot.StandardRvtPath, (registration == null) ? string.Empty : registration.ResolvedPath);
		string snapshotPath = FamilyBrowserStandardPolicyStore.ResolveSlotSnapshotPath(workspaceRoot, slot, registration);
		if (string.IsNullOrWhiteSpace(standardPath))
		{
			Add(report, "Standard RVT", "Blocked", "Standard RVT is not registered for the active target.", "Register the approved standard RVT.");
		}
		else if (File.Exists(standardPath))
		{
			Add(report, "Standard RVT", "Ready", "Standard RVT file exists.", standardPath);
		}
		else
		{
			Add(report, "Standard RVT", "Blocked", "Standard RVT file cannot be found.", standardPath);
		}
		if (string.IsNullOrWhiteSpace(snapshotPath))
		{
			Add(report, "Snapshot", "Blocked", "Standard snapshot is missing.", "Register or refresh the standard RVT to rebuild the snapshot.");
		}
		else if (File.Exists(snapshotPath))
		{
			Add(report, "Snapshot", "Ready", "Standard snapshot exists.", snapshotPath);
		}
		else
		{
			Add(report, "Snapshot", "Blocked", "Saved snapshot file cannot be found.", snapshotPath);
		}
	}

	private static void AddRequestStoreCheck(FamilyBrowserOperationalReadinessReport report, string workspaceRoot, FamilyBrowserStandardPolicy policy)
	{
		FamilyBrowserRequestStoreInfo info = FamilyBrowserRequestStore.GetRequestStoreInfo(workspaceRoot, policy);
		if (info == null)
		{
			Add(report, "Request Store", "Blocked", "Request store is not configured.", "Choose Local for testing or NetworkShare for production.");
			return;
		}
		string detail = string.Empty;
		if (!FamilyBrowserRequestStoreBackendService.TestWritable(info, ref detail))
		{
			Add(report, "Request Store", "Blocked", "Request store is not writable.", detail);
			return;
		}
		string left = FamilyBrowserPolicyKey.Normalize(info.Mode);
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("Local"), TextCompare: false) == 0)
		{
			Add(report, "Request Store", "Warning", "Local request store is writable but is for single-PC testing.", "Use a network share or synced folder for team-wide request visibility.");
		}
		else if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("NetworkShare"), TextCompare: false) == 0)
		{
			Add(report, "Request Store", "Ready", "Network request store is writable.", detail);
		}
		else if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("SharePoint"), TextCompare: false) == 0 || Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("CloudStorage"), TextCompare: false) == 0)
		{
			if (info.UsesConnectorQueue)
			{
				Add(report, "Request Store", "Warning", info.DisplayName + " is using a local connector queue.", "Set a synced folder path so every installed user sees the same request board.");
			}
			else
			{
				Add(report, "Request Store", "Ready", info.DisplayName + " synced folder is writable.", detail);
			}
		}
		else if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("Api"), TextCompare: false) == 0)
		{
			Add(report, "Request Store", "Warning", "API request store queue is writable.", "A connector service is still required to publish queued requests to the server endpoint.");
		}
		else
		{
			Add(report, "Request Store", "Warning", "Request store is writable.", info.Detail);
		}
	}

	private static void AddPermissionExcelCheck(FamilyBrowserOperationalReadinessReport report, FamilyBrowserStandardPolicy policy)
	{
		FamilyBrowserFileGuardPolicy fileGuard = policy?.FileGuard ?? FamilyBrowserFileGuardPolicy.CreateDefault();
		if (fileGuard == null || !fileGuard.Enabled)
		{
			Add(report, "File-specific Guard", "Warning", "File-specific native command guard is not enabled.", "Non-admin users can still be controlled by default policy, but no RVT-specific native command block list is active.");
			return;
		}
		Add(report, "File-specific Guard", "Ready", (fileGuard.Targets ?? new List<FamilyBrowserFileGuardTarget>()).Where([SpecialName] (FamilyBrowserFileGuardTarget x) => x?.Enabled ?? false).Count().ToString(CultureInfo.InvariantCulture) + " guarded RVT target(s) configured.", fileGuard.RootFolder ?? string.Empty);
	}

	private static void AddSecurityChecks(FamilyBrowserOperationalReadinessReport report, FamilyBrowserStandardPolicy policy, string currentUser)
	{
		if (FamilyBrowserSecurityPolicyService.HasConfiguredAdmin(policy?.Security))
		{
			Add(report, "Security", "Ready", "Administrator users or profile keywords are configured.", "Current role: " + FamilyBrowserSecurityPolicyService.ResolveRole(policy, currentUser));
		}
		else
		{
			Add(report, "Security", "Warning", "Bootstrap security mode is active.", "Set administrator users or Autodesk profile keywords before team deployment; otherwise everyone is treated as admin.");
		}
		if (policy != null)
		{
			bool num = FamilyBrowserSecurityPolicyService.Can(policy, currentUser, "LoadFamilies");
			bool canApplySystems = FamilyBrowserSecurityPolicyService.Can(policy, currentUser, "ApplySystemTypes");
			bool canRequest = FamilyBrowserSecurityPolicyService.Can(policy, currentUser, "SubmitRequest");
			if (num || canApplySystems || canRequest)
			{
				Add(report, "Current User", "Ready", "Current user has at least one modeler workflow permission.", currentUser);
			}
			else
			{
				Add(report, "Current User", "Warning", "Current user is effectively read-only.", currentUser);
			}
		}
	}

	private static void AddProjectPolicyCheck(FamilyBrowserOperationalReadinessReport report, FamilyBrowserStandardPolicy policy)
	{
		int count = ((policy != null && policy.ProjectPolicyRules != null) ? policy.ProjectPolicyRules.Where([SpecialName] (FamilyBrowserProjectPolicyRule x) => x?.Enabled ?? false).Count() : 0);
		if (count == 0)
		{
			Add(report, "Project Policy", "Warning", "No project-specific policy rule is configured.", "Global policy will be used for every project.");
		}
		else
		{
			Add(report, "Project Policy", "Ready", count.ToString(CultureInfo.InvariantCulture) + " project-specific policy rule(s) are available.", "Matching is evaluated from top to bottom.");
		}
	}

	private static void Add(FamilyBrowserOperationalReadinessReport report, string area, string status, string message, string action)
	{
		report.Items.Add(new FamilyBrowserOperationalReadinessItem
		{
			Area = (area ?? string.Empty),
			Status = (status ?? "Warning"),
			Message = (message ?? string.Empty),
			Action = (action ?? string.Empty)
		});
	}

	private static void FinalizeCounts(FamilyBrowserOperationalReadinessReport report)
	{
		report.BlockingCount = report.Items.Where([SpecialName] (FamilyBrowserOperationalReadinessItem x) => string.Equals(x.Status, "Blocked", StringComparison.OrdinalIgnoreCase)).Count();
		report.WarningCount = report.Items.Where([SpecialName] (FamilyBrowserOperationalReadinessItem x) => string.Equals(x.Status, "Warning", StringComparison.OrdinalIgnoreCase)).Count();
		report.ReadyCount = report.Items.Where([SpecialName] (FamilyBrowserOperationalReadinessItem x) => string.Equals(x.Status, "Ready", StringComparison.OrdinalIgnoreCase)).Count();
	}

	private static StandardLibraryRegistrationRecord TryLoadRegistration(string workspaceRoot, FamilyBrowserStandardPolicy policy)
	{
		try
		{
			string path = FamilyBrowserStandardPolicyStore.ResolveEffectiveRegistrationPath(workspaceRoot, policy);
			if (File.Exists(path))
			{
				return DataContractJsonFileStore.Load<StandardLibraryRegistrationRecord>(path);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return null;
	}

	private static string ResolveStandardTarget(FamilyBrowserStandardPolicy policy, FamilyBrowserStandardLibrarySlot slot)
	{
		if (policy == null)
		{
			return "No policy";
		}
		if (string.Equals(policy.Mode, "Integrated", StringComparison.OrdinalIgnoreCase))
		{
			return "Integrated";
		}
		if (slot != null && !string.IsNullOrWhiteSpace(slot.Discipline))
		{
			return slot.Discipline;
		}
		return policy.ActiveDiscipline ?? string.Empty;
	}

	private static string ResolveRequestStoreMode(FamilyBrowserStandardPolicy policy)
	{
		if (policy == null || policy.RequestStore == null || string.IsNullOrWhiteSpace(policy.RequestStore.Mode))
		{
			return "Local";
		}
		return policy.RequestStore.Mode;
	}

	private static bool IsWritableFolder(string folderPath, ref string detail)
	{
		bool IsWritableFolder;
		if (string.IsNullOrWhiteSpace(folderPath))
		{
			detail = "Folder path is empty.";
			IsWritableFolder = false;
		}
		else
		{
			try
			{
				string expanded = Environment.ExpandEnvironmentVariables(folderPath);
				if (expanded.IndexOf("NoManagedDataRoot", StringComparison.OrdinalIgnoreCase) >= 0)
				{
					detail = "Managed shared folder is not connected.";
					IsWritableFolder = false;
				}
				else if (!IsAllowedSharedRuntimePath(expanded))
				{
					detail = "Local C/AppData/UserProfile paths are not valid Family Browser managed data folders.";
					IsWritableFolder = false;
				}
				else
				{
					Directory.CreateDirectory(expanded);
					string path = Path.Combine(expanded, ".kky-family-browser-write-test-" + Guid.NewGuid().ToString("N") + ".tmp");
					File.WriteAllText(path, DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
					File.Delete(path);
					detail = expanded;
					IsWritableFolder = true;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				detail = folderPath + " | " + ex2.Message;
				IsWritableFolder = false;
				ProjectData.ClearProjectError();
			}
		}
		return IsWritableFolder;
	}

	private static bool IsAllowedSharedRuntimePath(string path)
	{
		string value = Environment.ExpandEnvironmentVariables((path ?? string.Empty).Trim());
		bool IsAllowedSharedRuntimePath;
		if (string.IsNullOrWhiteSpace(value))
		{
			IsAllowedSharedRuntimePath = false;
		}
		else
		{
			try
			{
				if (value.StartsWith("\\\\", StringComparison.Ordinal))
				{
					IsAllowedSharedRuntimePath = true;
				}
				else
				{
					string root = Path.GetPathRoot(value);
					if (string.IsNullOrWhiteSpace(root))
					{
						IsAllowedSharedRuntimePath = false;
					}
					else
					{
						string windowsRoot = Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.Windows));
						IsAllowedSharedRuntimePath = (string.IsNullOrWhiteSpace(windowsRoot) || !string.Equals(root, windowsRoot, StringComparison.OrdinalIgnoreCase)) && !IsSameOrChildPath(value, Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)) && !IsSameOrChildPath(value, Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)) && !IsSameOrChildPath(value, Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)) && !IsSameOrChildPath(value, Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)) && !IsSameOrChildPath(value, Path.GetTempPath());
					}
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				IsAllowedSharedRuntimePath = false;
				ProjectData.ClearProjectError();
			}
		}
		return IsAllowedSharedRuntimePath;
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

	private static string FirstNonEmpty(params string[] values)
	{
		foreach (string value in values)
		{
			if (!string.IsNullOrWhiteSpace(value))
			{
				return value.Trim();
			}
		}
		return string.Empty;
	}
}
