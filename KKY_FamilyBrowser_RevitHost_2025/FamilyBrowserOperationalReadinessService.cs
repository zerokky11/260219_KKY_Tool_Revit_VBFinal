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
			Area = LocalizeReadinessArea(area ?? string.Empty),
			Status = (status ?? "Warning"),
			Message = LocalizeReadinessText(message ?? string.Empty),
			Action = LocalizeReadinessText(action ?? string.Empty)
		});
	}

	private static string T(string englishText, string koreanText)
	{
		return FamilyBrowserLanguageService.Text(englishText, koreanText);
	}

	private static string LocalizeReadinessArea(string value)
	{
		if (!FamilyBrowserLanguageService.IsKorean())
		{
			return value ?? string.Empty;
		}
		switch (value ?? string.Empty)
		{
		case "Policy":
			return "정책";
		case "Deployment":
			return "배포";
		case "Standard RVT":
			return "표준 RVT";
		case "Snapshot":
			return "스냅샷";
		case "Request Store":
			return "요청 저장소";
		case "File-specific Guard":
			return "파일별 보호";
		case "Security":
			return "보안";
		case "Current User":
			return "현재 사용자";
		case "Project Policy":
			return "프로젝트 정책";
		default:
			return value ?? string.Empty;
		}
	}

	private static string LocalizeReadinessText(string value)
	{
		string text = value ?? string.Empty;
		if (!FamilyBrowserLanguageService.IsKorean() || string.IsNullOrWhiteSpace(text))
		{
			return text;
		}
		switch (text)
		{
		case "Policy path is empty.":
			return "정책 경로가 비어 있습니다.";
		case "Set a local or managed policy path.":
			return "로컬 또는 공용 관리 정책 경로를 설정하세요.";
		case "Policy file location is writable.":
			return "정책 파일 위치에 쓸 수 있습니다.";
		case "Policy file location is not writable.":
			return "정책 파일 위치에 쓸 수 없습니다.";
		case "Local test policy is active.":
			return "로컬 테스트 정책이 활성화되어 있습니다.";
		case "Use a managed policy path for company deployment.":
			return "회사 배포용 공용 관리 정책 경로를 사용하세요.";
		case "Managed policy is connected.":
			return "공용 관리 정책이 연결되어 있습니다.";
		case "Managed policy path is unavailable on this PC.":
			return "이 PC에서 공용 관리 정책 경로를 사용할 수 없습니다.";
		case "No active standard library slot is available.":
			return "활성 표준 라이브러리 슬롯이 없습니다.";
		case "Choose discipline separated or integrated mode and register a standard RVT.":
			return "공종 분리 또는 통합 모드를 선택하고 표준 RVT를 등록하세요.";
		case "The active standard library slot is disabled.":
			return "활성 표준 라이브러리 슬롯이 비활성화되어 있습니다.";
		case "Enable the slot or choose another trade.":
			return "슬롯을 활성화하거나 다른 공종을 선택하세요.";
		case "Standard RVT is not registered for the active target.":
			return "현재 대상에 표준 RVT가 등록되어 있지 않습니다.";
		case "Register the approved standard RVT.":
			return "승인된 표준 RVT를 등록하세요.";
		case "Standard RVT file exists.":
			return "표준 RVT 파일이 있습니다.";
		case "Standard RVT file cannot be found.":
			return "표준 RVT 파일을 찾을 수 없습니다.";
		case "Standard snapshot is missing.":
			return "표준 스냅샷이 없습니다.";
		case "Register or refresh the standard RVT to rebuild the snapshot.":
			return "표준 RVT를 등록하거나 새로고침해서 스냅샷을 다시 만드세요.";
		case "Standard snapshot exists.":
			return "표준 스냅샷이 있습니다.";
		case "Saved snapshot file cannot be found.":
			return "저장된 스냅샷 파일을 찾을 수 없습니다.";
		case "Request store is not configured.":
			return "요청 저장소가 설정되어 있지 않습니다.";
		case "Choose Local for testing or NetworkShare for production.":
			return "테스트는 Local, 운영은 NetworkShare를 선택하세요.";
		case "Request store is not writable.":
			return "요청 저장소에 쓸 수 없습니다.";
		case "Local request store is writable but is for single-PC testing.":
			return "로컬 요청 저장소에 쓸 수 있지만 단일 PC 테스트용입니다.";
		case "Use a network share or synced folder for team-wide request visibility.":
			return "팀 전체가 요청을 볼 수 있도록 네트워크 공유 또는 동기화 폴더를 사용하세요.";
		case "Network request store is writable.":
			return "네트워크 요청 저장소에 쓸 수 있습니다.";
		case "Set a synced folder path so every installed user sees the same request board.":
			return "모든 설치 사용자가 같은 요청 목록을 볼 수 있도록 동기화 폴더 경로를 설정하세요.";
		case "API request store queue is writable.":
			return "API 요청 저장소 큐에 쓸 수 있습니다.";
		case "A connector service is still required to publish queued requests to the server endpoint.":
			return "큐에 쌓인 요청을 서버 엔드포인트로 게시하려면 커넥터 서비스가 필요합니다.";
		case "Request store is writable.":
			return "요청 저장소에 쓸 수 있습니다.";
		case "File-specific native command guard is not enabled.":
			return "파일별 Revit 기본 명령 차단이 활성화되어 있지 않습니다.";
		case "Non-admin users can still be controlled by default policy, but no RVT-specific native command block list is active.":
			return "비관리자는 기본 정책으로 제어될 수 있지만 RVT별 기본 명령 차단 목록은 활성화되어 있지 않습니다.";
		case "Administrator users or profile keywords are configured.":
			return "관리자 사용자 또는 프로필 키워드가 설정되어 있습니다.";
		case "Bootstrap security mode is active.":
			return "초기 보안 모드가 활성화되어 있습니다.";
		case "Set administrator users or Autodesk profile keywords before team deployment; otherwise everyone is treated as admin.":
			return "팀 배포 전에 관리자 사용자 또는 Autodesk 프로필 키워드를 설정하세요. 설정하지 않으면 모든 사용자가 관리자로 처리됩니다.";
		case "Current user has at least one modeler workflow permission.":
			return "현재 사용자는 하나 이상의 모델러 작업 권한을 가지고 있습니다.";
		case "Current user is effectively read-only.":
			return "현재 사용자는 사실상 읽기 전용입니다.";
		case "No project-specific policy rule is configured.":
			return "프로젝트별 정책 규칙이 설정되어 있지 않습니다.";
		case "Global policy will be used for every project.":
			return "모든 프로젝트에 전역 정책을 사용합니다.";
		case "Matching is evaluated from top to bottom.":
			return "매칭은 위에서 아래 순서로 평가됩니다.";
		case "Folder path is empty.":
			return "폴더 경로가 비어 있습니다.";
		case "Managed shared folder is not connected.":
			return "공용 관리 폴더가 연결되어 있지 않습니다.";
		case "Local C/AppData/UserProfile paths are not valid Family Browser managed data folders.":
			return "로컬 C/AppData/UserProfile 경로는 Family Browser 관리 데이터 폴더로 사용할 수 없습니다.";
		default:
			break;
		}
		const string configuredPathPrefix = "Refresh the homepage path and connect the first reachable managed folder. Configured path: ";
		if (text.StartsWith(configuredPathPrefix, StringComparison.Ordinal))
		{
			return "홈페이지 경로를 새로고침하고 처음 접근 가능한 공용 관리 폴더를 연결하세요. 설정된 경로: " + text.Substring(configuredPathPrefix.Length);
		}
		const string connectorSuffix = " is using a local connector queue.";
		if (text.EndsWith(connectorSuffix, StringComparison.Ordinal))
		{
			return text.Substring(0, checked(text.Length - connectorSuffix.Length)) + "은(는) 로컬 커넥터 큐를 사용 중입니다.";
		}
		const string syncedSuffix = " synced folder is writable.";
		if (text.EndsWith(syncedSuffix, StringComparison.Ordinal))
		{
			return text.Substring(0, checked(text.Length - syncedSuffix.Length)) + " 동기화 폴더에 쓸 수 있습니다.";
		}
		const string currentRolePrefix = "Current role: ";
		if (text.StartsWith(currentRolePrefix, StringComparison.Ordinal))
		{
			return "현재 역할: " + text.Substring(currentRolePrefix.Length);
		}
		const string guardedSuffix = " guarded RVT target(s) configured.";
		if (text.EndsWith(guardedSuffix, StringComparison.Ordinal))
		{
			return text.Substring(0, checked(text.Length - guardedSuffix.Length)) + "개의 보호 RVT 대상이 설정되었습니다.";
		}
		const string projectRuleSuffix = " project-specific policy rule(s) are available.";
		if (text.EndsWith(projectRuleSuffix, StringComparison.Ordinal))
		{
			return text.Substring(0, checked(text.Length - projectRuleSuffix.Length)) + "개의 프로젝트별 정책 규칙을 사용할 수 있습니다.";
		}
		return text;
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
					string path = Path.Combine(expanded, ".kky-w-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".tmp");
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
