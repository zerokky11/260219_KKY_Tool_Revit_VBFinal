using System;
using System.Globalization;
using System.IO;
using System.Text;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserTestRootSetupService
{
	private FamilyBrowserTestRootSetupService()
	{
	}

	public static bool IsTestRootConfigured()
	{
		FamilyBrowserMachineConfig config = FamilyBrowserMachineConfigStore.Load();
		if (config == null || !config.UseManagedPolicy || string.IsNullOrWhiteSpace(config.ManagedPolicyPath))
		{
			return false;
		}
		if (!FamilyBrowserDeploymentBootstrapService.IsBootstrapDisabled(config.DeploymentBootstrapUrl))
		{
			return false;
		}
		string rootFolder = ResolveRootFolderFromPolicyPath(config.ManagedPolicyPath);
		if (string.IsNullOrWhiteSpace(rootFolder))
		{
			return false;
		}
		return File.Exists(Path.Combine(rootFolder, "KKY_FamilyBrowser_TEST_ROOT_README.txt"));
	}

	public static string ResolveConfiguredRootFolder()
	{
		string managedPolicyPath = FamilyBrowserMachineConfigStore.ResolveManagedPolicyPath();
		if (string.IsNullOrWhiteSpace(managedPolicyPath))
		{
			return string.Empty;
		}
		return ResolveRootFolderFromPolicyPath(managedPolicyPath);
	}

	private static string ResolveRootFolderFromPolicyPath(string managedPolicyPath)
	{
		string ResolveRootFolderFromPolicyPath;
		try
		{
			string configFolder = Path.GetDirectoryName(Environment.ExpandEnvironmentVariables((managedPolicyPath ?? string.Empty).Trim()));
			if (string.IsNullOrWhiteSpace(configFolder))
			{
				ResolveRootFolderFromPolicyPath = string.Empty;
			}
			else
			{
				DirectoryInfo parent = Directory.GetParent(configFolder.TrimEnd(new char[2]
				{
					Path.DirectorySeparatorChar,
					Path.AltDirectorySeparatorChar
				}));
				ResolveRootFolderFromPolicyPath = ((parent == null || string.IsNullOrWhiteSpace(parent.FullName)) ? configFolder : parent.FullName);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveRootFolderFromPolicyPath = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveRootFolderFromPolicyPath;
	}

	public static void DeactivateExternalBootstrap(string currentUser)
	{
		FamilyBrowserDeploymentBootstrapService.DisableBootstrap(currentUser);
		FamilyBrowserMachineConfigStore.ClearManagedPolicyPath(currentUser);
	}

	public static FamilyBrowserTestRootSetupResult Configure(string workspaceRoot, string selectedRoot, FamilyBrowserStandardPolicy sourcePolicy, string currentUser)
	{
		throw new InvalidOperationException(FamilyBrowserLanguageService.Text("Manual TEST root setup has been removed. Refresh the homepage path and use the managed shared folder only.", "수동 TEST 루트 설정은 제거되었습니다. 홈페이지 경로를 다시 확인하고 공용 관리 폴더만 사용하세요."));
	}

	private static string NormalizeFolder(string folder)
	{
		string expanded = Environment.ExpandEnvironmentVariables(folder ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(expanded))
		{
			return string.Empty;
		}
		return Path.GetFullPath(expanded.TrimEnd(new char[2]
		{
			Path.DirectorySeparatorChar,
			Path.AltDirectorySeparatorChar
		}));
	}

	private static string WriteGuide(FamilyBrowserTestRootSetupResult result)
	{
		string text = Path.Combine(result.RootFolder, "KKY_FamilyBrowser_TEST_ROOT_README.txt");
		StringBuilder builder = new StringBuilder();
		builder.AppendLine("KKY Family Browser Test Root");
		builder.AppendLine("Generated: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
		builder.AppendLine();
		builder.AppendLine("Use this root folder on every test PC when you want requests and standard RVT registration to be shared.");
		builder.AppendLine("The update-homepage bootstrap is disabled on this PC while this test root is active.");
		builder.AppendLine();
		builder.AppendLine("Root: " + result.RootFolder);
		builder.AppendLine("Policy: " + result.SharedPolicyPath);
		builder.AppendLine("Requests: " + result.RequestFolder);
		builder.AppendLine("Standards: " + result.StandardsFolder);
		builder.AppendLine("Registrations: " + result.RegistryFolder);
		builder.AppendLine("Snapshots: " + result.SnapshotFolder);
		builder.AppendLine("Thumbnails: " + result.ThumbnailFolder);
		builder.AppendLine();
		builder.AppendLine("How to test:");
		builder.AppendLine("1. On PC A and PC B, open Family Browser Settings.");
		builder.AppendLine("2. Choose Test Root Folder and select this same shared folder.");
		builder.AppendLine("3. Register a standard RVT from the add-in.");
		builder.AppendLine("4. Create a request on PC A, then refresh requests on PC B.");
		File.WriteAllText(text, builder.ToString(), Encoding.UTF8);
		return text;
	}

	private static bool IsLikelyTeamVisiblePath(string folderPath)
	{
		if (string.IsNullOrWhiteSpace(folderPath))
		{
			return false;
		}
		if (folderPath.StartsWith("\\\\", StringComparison.OrdinalIgnoreCase))
		{
			return true;
		}
		return !string.IsNullOrWhiteSpace(Path.GetPathRoot(folderPath)) && !folderPath.StartsWith(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), StringComparison.OrdinalIgnoreCase);
	}
}
