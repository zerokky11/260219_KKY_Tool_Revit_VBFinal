using System;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserRequestStoreBackendService
{
	private FamilyBrowserRequestStoreBackendService()
	{
	}

	public static FamilyBrowserRequestStoreInfo ResolveInfo(string workspaceRoot, FamilyBrowserStandardPolicy policy)
	{
		string mode = ResolveStoreMode(policy);
		string configuredPath = ResolveConfiguredPath(policy);
		string endpoint = ResolveConfiguredEndpoint(policy);
		FamilyBrowserRequestStoreInfo info = new FamilyBrowserRequestStoreInfo
		{
			Mode = mode,
			Endpoint = endpoint
		};
		string left = FamilyBrowserPolicyKey.Normalize(mode);
		if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("NetworkShare"), TextCompare: false) == 0)
		{
			info.DisplayName = "Network share";
			info.StoreLocation = Environment.ExpandEnvironmentVariables(configuredPath);
			info.IsShared = true;
			info.IsFileBacked = true;
			info.Detail = (string.IsNullOrWhiteSpace(info.StoreLocation) ? "Network request store path is not configured." : "Team-visible request board through a shared Windows folder.");
		}
		else if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("SharePoint"), TextCompare: false) == 0)
		{
			info.DisplayName = "SharePoint / Microsoft 365";
			ApplySyncedOrQueuedStore(info, workspaceRoot, configuredPath, endpoint, "SharePoint", "Use a locally synced SharePoint folder for team-visible requests.");
		}
		else if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("CloudStorage"), TextCompare: false) == 0)
		{
			info.DisplayName = "Cloud storage";
			ApplySyncedOrQueuedStore(info, workspaceRoot, configuredPath, endpoint, "CloudStorage", "Use a locally synced cloud folder for team-visible requests.");
		}
		else if (Operators.CompareString(left, FamilyBrowserPolicyKey.Normalize("Api"), TextCompare: false) == 0)
		{
			info.DisplayName = "Internal API / DB";
			if (!string.IsNullOrWhiteSpace(configuredPath))
			{
				info.StoreLocation = Environment.ExpandEnvironmentVariables(configuredPath);
				info.IsShared = IsLikelySharedPath(info.StoreLocation);
				info.IsFileBacked = true;
				info.UsesConnectorQueue = true;
				info.RequiresConnectorSync = true;
				info.Detail = "Requests are staged in an API connector queue. A server connector can later post them to " + (string.IsNullOrWhiteSpace(endpoint) ? "the configured endpoint." : endpoint);
			}
			else
			{
				info.StoreLocation = GetLocalConnectorQueueFolder(workspaceRoot, mode);
				info.IsShared = false;
				info.IsFileBacked = true;
				info.UsesConnectorQueue = true;
				info.RequiresConnectorSync = true;
				info.Detail = "Requests are saved to a local API connector queue until a server connector is configured.";
			}
		}
		else
		{
			info.Mode = "Local";
			info.DisplayName = "Local test store";
			info.StoreLocation = FamilyBrowserStandardPolicyStore.GetDataFolder(ResolveWorkspaceRoot(workspaceRoot), "Requests");
			info.IsShared = false;
			info.IsFileBacked = true;
			info.Detail = "Local single-PC request store for demos and debugging.";
		}
		return info;
	}

	public static string ResolveWritableFolder(string workspaceRoot, FamilyBrowserStandardPolicy policy, bool requireWritable)
	{
		FamilyBrowserRequestStoreInfo info = ResolveInfo(workspaceRoot, policy);
		if (string.IsNullOrWhiteSpace(info.StoreLocation))
		{
			if (requireWritable)
			{
				throw new InvalidOperationException(info.Detail);
			}
			return string.Empty;
		}
		if (requireWritable)
		{
			if (!IsAllowedSharedRuntimePath(info.StoreLocation))
			{
				throw new InvalidOperationException(FamilyBrowserLanguageService.Text("Request data must be written only to the managed shared folder, not a local C/AppData/UserProfile path.", "요청 데이터는 로컬 C/AppData/사용자 폴더가 아니라 공용 관리 폴더에만 저장해야 합니다."));
			}
			if (!info.IsShared && !FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
			{
				throw new InvalidOperationException(FamilyBrowserLanguageService.Text("Request data is not written to the local C fallback folder. Refresh the homepage path and connect a managed shared folder first.", "요청 데이터는 로컬 C fallback 폴더에 저장하지 않습니다. 먼저 홈페이지 경로를 다시 확인해서 공용 관리 폴더를 연결하세요."));
			}
			string unavailableDetail = string.Empty;
			if (IsObviouslyUnavailable(info, ref unavailableDetail))
			{
				throw new InvalidOperationException(unavailableDetail);
			}
			if (info.IsShared && !DirectoryExistsFast(info))
			{
				throw new InvalidOperationException(info.StoreLocation + " | Request store is not reachable.");
			}
			Directory.CreateDirectory(info.StoreLocation);
			WriteStoreManifest(info.StoreLocation, info);
		}
		return info.StoreLocation;
	}

	public static bool IsObviouslyUnavailable(FamilyBrowserRequestStoreInfo info, ref string detail)
	{
		detail = string.Empty;
		bool IsObviouslyUnavailable;
		if (info == null || string.IsNullOrWhiteSpace(info.StoreLocation))
		{
			detail = ((info == null) ? "Request store is not configured." : info.Detail);
			IsObviouslyUnavailable = true;
		}
		else if (!info.IsFileBacked || !info.IsShared)
		{
			IsObviouslyUnavailable = false;
		}
		else
		{
			string location = Environment.ExpandEnvironmentVariables(info.StoreLocation).Trim();
			if (string.IsNullOrWhiteSpace(location))
			{
				detail = "Request store path is empty.";
				IsObviouslyUnavailable = true;
			}
			else
			{
				string root = string.Empty;
				try
				{
					root = Path.GetPathRoot(location);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					detail = location + " | " + ex2.Message;
					IsObviouslyUnavailable = true;
					ProjectData.ClearProjectError();
					goto IL_0115;
				}
				if (string.IsNullOrWhiteSpace(root) || root.StartsWith("\\\\", StringComparison.Ordinal))
				{
					IsObviouslyUnavailable = false;
				}
				else
				{
					if (root.Length >= 2 && root[1] == ':')
					{
						try
						{
							if (!new DriveInfo(root).IsReady)
							{
								detail = location + " | Drive is not ready.";
								IsObviouslyUnavailable = true;
								goto IL_0115;
							}
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							detail = location + " | " + ex4.Message;
							IsObviouslyUnavailable = true;
							ProjectData.ClearProjectError();
							goto IL_0115;
						}
					}
					IsObviouslyUnavailable = false;
				}
			}
		}
		goto IL_0115;
		IL_0115:
		return IsObviouslyUnavailable;
	}

	public static bool DirectoryExistsFast(FamilyBrowserRequestStoreInfo info, int timeoutMilliseconds = 1500)
	{
		if (info == null || string.IsNullOrWhiteSpace(info.StoreLocation))
		{
			return false;
		}
		string unavailableDetail = string.Empty;
		if (IsObviouslyUnavailable(info, ref unavailableDetail))
		{
			return false;
		}
		string storeLocation = info.StoreLocation;
		if (!info.IsShared && !IsLikelySharedPath(storeLocation))
		{
			return Directory.Exists(storeLocation);
		}
		try
		{
			Task<bool> probe = Task.Factory.StartNew([SpecialName] () => Directory.Exists(storeLocation));
			if (probe.Wait(Math.Max(250, timeoutMilliseconds)))
			{
				return probe.Result;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return false;
	}

	public static string ResolveOpenableFolder(string workspaceRoot, FamilyBrowserStandardPolicy policy)
	{
		FamilyBrowserRequestStoreInfo info = ResolveInfo(workspaceRoot, policy);
		if (info != null)
		{
			return info.StoreLocation;
		}
		return string.Empty;
	}

	public static string BuildSummary(string workspaceRoot, FamilyBrowserStandardPolicy policy)
	{
		FamilyBrowserRequestStoreInfo info = ResolveInfo(workspaceRoot, policy);
		if (info == null)
		{
			return string.Empty;
		}
		StringBuilder builder = new StringBuilder();
		builder.Append(info.DisplayName);
		if (!string.IsNullOrWhiteSpace(info.StoreLocation))
		{
			builder.Append(" | ");
			builder.Append(info.StoreLocation);
		}
		if (info.UsesConnectorQueue)
		{
			builder.Append(" | connector queue");
		}
		if (!string.IsNullOrWhiteSpace(info.Endpoint))
		{
			builder.Append(" | ");
			builder.Append(info.Endpoint);
		}
		return builder.ToString();
	}

	public static bool TestWritable(FamilyBrowserRequestStoreInfo info, ref string detail)
	{
		bool TestWritable;
		if (info == null || string.IsNullOrWhiteSpace(info.StoreLocation))
		{
			detail = ((info == null) ? "Request store is not configured." : info.Detail);
			TestWritable = false;
		}
		else
		{
			string unavailableDetail = string.Empty;
			if (IsObviouslyUnavailable(info, ref unavailableDetail))
			{
				detail = unavailableDetail;
				TestWritable = false;
			}
			else
			{
				try
				{
					if (!IsAllowedSharedRuntimePath(info.StoreLocation))
					{
						detail = FamilyBrowserLanguageService.Text("Request data must be written only to the managed shared folder, not a local C/AppData/UserProfile path.", "요청 데이터는 로컬 C/AppData/사용자 폴더가 아니라 공용 관리 폴더에만 저장해야 합니다.");
						TestWritable = false;
					}
					else if (info.IsShared && !DirectoryExistsFast(info))
					{
						detail = info.StoreLocation + " | Request store is not reachable.";
						TestWritable = false;
					}
					else
					{
						Directory.CreateDirectory(info.StoreLocation);
						string path = Path.Combine(info.StoreLocation, ".kky-w-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".tmp");
						File.WriteAllText(path, DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture), Encoding.UTF8);
						File.Delete(path);
						detail = info.StoreLocation;
						TestWritable = true;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					detail = info.StoreLocation + " | " + ex2.Message;
					TestWritable = false;
					ProjectData.ClearProjectError();
				}
			}
		}
		return TestWritable;
	}

	public static void WriteStoreManifest(string folder, FamilyBrowserRequestStoreInfo info)
	{
		if (info == null || string.IsNullOrWhiteSpace(folder))
		{
			return;
		}
		try
		{
			if (IsAllowedSharedRuntimePath(folder))
			{
				Directory.CreateDirectory(folder);
				File.WriteAllText(Path.Combine(folder, "request-store-info.json"), PlainJsonReportWriter.Serialize(info), Encoding.UTF8);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void ApplySyncedOrQueuedStore(FamilyBrowserRequestStoreInfo info, string workspaceRoot, string configuredPath, string endpoint, string connectorName, string syncedDetail)
	{
		if (!string.IsNullOrWhiteSpace(configuredPath))
		{
			info.StoreLocation = Environment.ExpandEnvironmentVariables(configuredPath);
			info.IsShared = true;
			info.IsFileBacked = true;
			info.UsesConnectorQueue = false;
			info.RequiresConnectorSync = false;
			info.Detail = syncedDetail;
		}
		else
		{
			info.StoreLocation = GetLocalConnectorQueueFolder(workspaceRoot, connectorName);
			info.IsShared = false;
			info.IsFileBacked = true;
			info.UsesConnectorQueue = true;
			info.RequiresConnectorSync = true;
			info.Detail = connectorName + " request store has no synced folder yet. Requests are held in a local connector queue.";
		}
	}

	private static string ResolveStoreMode(FamilyBrowserStandardPolicy policy)
	{
		if (policy == null || policy.RequestStore == null)
		{
			return "Local";
		}
		switch (FamilyBrowserPolicyKey.Normalize(policy.RequestStore.Mode))
		{
		case "networkshare":
		case "network-share":
		case "network":
		case "unc":
		case "share":
			return "NetworkShare";
		case "sharepoint":
		case "share-point":
		case "m365":
		case "office365":
			return "SharePoint";
		case "cloudstorage":
		case "cloud-storage":
		case "cloud":
			return "CloudStorage";
		case "api":
		case "server":
		case "database":
		case "db":
			return "Api";
		default:
			return "Local";
		}
	}

	private static string ResolveConfiguredPath(FamilyBrowserStandardPolicy policy)
	{
		if (policy == null || policy.RequestStore == null)
		{
			return string.Empty;
		}
		return (policy.RequestStore.Path ?? string.Empty).Trim();
	}

	private static string ResolveConfiguredEndpoint(FamilyBrowserStandardPolicy policy)
	{
		if (policy == null || policy.RequestStore == null)
		{
			return string.Empty;
		}
		return (policy.RequestStore.Endpoint ?? string.Empty).Trim();
	}

	private static string GetLocalConnectorQueueFolder(string workspaceRoot, string mode)
	{
		string safeMode = FamilyBrowserPolicyKey.Normalize(mode);
		if (string.IsNullOrWhiteSpace(safeMode))
		{
			safeMode = "connector";
		}
		return FamilyBrowserStandardPolicyStore.GetDataFolder(ResolveWorkspaceRoot(workspaceRoot), Path.Combine("Requests", "ConnectorQueue", safeMode));
	}

	private static string ResolveWorkspaceRoot(string workspaceRoot)
	{
		if (string.IsNullOrWhiteSpace(workspaceRoot))
		{
			return HostWorkspacePathResolver.ResolveRoot();
		}
		return workspaceRoot;
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

	private static bool IsLikelySharedPath(string path)
	{
		if (string.IsNullOrWhiteSpace(path))
		{
			return false;
		}
		string trimmed = path.Trim();
		if (trimmed.StartsWith("\\\\", StringComparison.OrdinalIgnoreCase))
		{
			return true;
		}
		return trimmed.IndexOf("OneDrive", StringComparison.OrdinalIgnoreCase) >= 0 || trimmed.IndexOf("SharePoint", StringComparison.OrdinalIgnoreCase) >= 0 || trimmed.IndexOf("Dropbox", StringComparison.OrdinalIgnoreCase) >= 0 || trimmed.IndexOf("Google Drive", StringComparison.OrdinalIgnoreCase) >= 0;
	}
}
