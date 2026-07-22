using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;

public partial class FamilyBrowserDashboardHtmlForm
{
	private FamilyBrowserHomepageManagedFolderProbeResult _homepageManagedFolderProbeResult;

	private bool? _auditManagedFolderTestOverride;

	private string _auditTestManagedFolderRoot = string.Empty;

	private bool IsTestManagedFolderOverrideActive()
	{
		if (_auditManagedFolderTestOverride.HasValue)
		{
			return _auditManagedFolderTestOverride.Value;
		}
		return FamilyBrowserManagedFolderSetupService.IsOverrideConfigured();
	}

	private string ResolveTestManagedFolderRoot()
	{
		if (_auditManagedFolderTestOverride.HasValue)
		{
			return _auditTestManagedFolderRoot ?? string.Empty;
		}
		return FamilyBrowserManagedFolderSetupService.GetConfiguredOverrideRoot();
	}

	private FamilyBrowserHomepageManagedFolderProbeResult ProbeHomepageManagedFolderForReturn()
	{
		_homepageManagedFolderProbeResult = FamilyBrowserDeploymentBootstrapService.ProbeHomepageManagedFolder(BuildDeploymentProjectIdentity(GetActiveDocument()));
		return _homepageManagedFolderProbeResult;
	}

	private bool ProbeHomepageReturnWhileTestFolderActive(bool showResult)
	{
		if (!IsTestManagedFolderOverrideActive())
		{
			return false;
		}
		FamilyBrowserHomepageManagedFolderProbeResult probe = ProbeHomepageManagedFolderForReturn();
		bool available = probe != null && probe.Available;
		_statusMessage = available
			? T("A homepage management folder is now available. Choose Switch or Migrate and Switch below.", "홈페이지 관리폴더를 사용할 수 있습니다. 아래에서 홈페이지 경로로 변경 또는 기존 데이터 이관 후 변경을 선택하세요.")
			: T("The TEST management folder remains active. No reachable homepage management folder was found.", "TEST 관리폴더를 계속 사용합니다. 접근 가능한 홈페이지 관리폴더를 찾지 못했습니다.");
		RenderDashboard();
		if (showResult)
		{
			string sourceRoot = ResolveTestManagedFolderRoot();
			if (available)
			{
				string message = T(
					"Homepage management folder found" + Environment.NewLine + probe.ManagedRootPath + Environment.NewLine + Environment.NewLine + "Current TEST management folder" + Environment.NewLine + sourceRoot + Environment.NewLine + Environment.NewLine + "Nothing has been switched yet. Use Switch to Homepage Folder, or use Migrate Existing Data and Switch in Admin Settings. The TEST source folder is retained.",
					"홈페이지 관리폴더 확인 완료" + Environment.NewLine + probe.ManagedRootPath + Environment.NewLine + Environment.NewLine + "현재 TEST 관리폴더" + Environment.NewLine + sourceRoot + Environment.NewLine + Environment.NewLine + "아직 경로를 변경하지 않았습니다. 홈페이지 경로로 변경을 사용하거나 관리자 설정에서 기존 데이터 이관 후 변경을 선택하세요. TEST 원본 폴더는 그대로 보존됩니다.");
				ShowDashboardMessage(this, message, T("Homepage Management Folder Available", "홈페이지 관리폴더 사용 가능"), MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			else
			{
				string issue = probe == null ? string.Empty : probe.Issue;
				string message = T(
					"No reachable homepage management folder was found." + Environment.NewLine + Environment.NewLine + "Current TEST management folder" + Environment.NewLine + sourceRoot + Environment.NewLine + Environment.NewLine + "Connect the company network or VPN, then refresh again." + (string.IsNullOrWhiteSpace(issue) ? string.Empty : Environment.NewLine + Environment.NewLine + "Details" + Environment.NewLine + issue),
					"접근 가능한 홈페이지 관리폴더를 찾지 못했습니다." + Environment.NewLine + Environment.NewLine + "현재 TEST 관리폴더" + Environment.NewLine + sourceRoot + Environment.NewLine + Environment.NewLine + "사내망 또는 VPN을 연결한 뒤 다시 새로고침하세요." + (string.IsNullOrWhiteSpace(issue) ? string.Empty : Environment.NewLine + Environment.NewLine + "상세" + Environment.NewLine + issue));
				ShowDashboardMessage(this, message, T("Homepage Management Folder Unavailable", "홈페이지 관리폴더 연결 안 됨"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
		}
		return true;
	}

	private void SwitchToHomepageManagedFolder(bool migrateExistingData)
	{
		IDisposable managementContextLease = null;
		IDisposable trackingTransitionAuthorization = null;
		try
		{
			managementContextLease = FamilyBrowserManagementContextLock.Acquire(TimeSpan.FromSeconds(5.0));
			trackingTransitionAuthorization = FamilyBrowserElementChangeTrackingService.AuthorizeManagedFolderTransition();
			if (!IsTestManagedFolderOverrideActive())
			{
				_statusMessage = T("No TEST management-folder override is active.", "현재 활성화된 TEST 관리폴더 설정이 없습니다.");
				RenderDashboard();
				return;
			}
			if (migrateExistingData && !EnsurePermission("ManagePolicy", T("Migrate Management Folder", "관리폴더 데이터 이관")))
			{
				return;
			}
			int otherRevitProcessCount = GetOtherRevitProcessCount();
			if (otherRevitProcessCount > 0)
			{
				ShowDashboardMessage(this,
					T(
						"The management folder cannot change while another Revit process is running on this PC. Close the other Revit sessions first so they cannot keep writing tracking records to the previous management folder.",
						"이 PC에서 다른 Revit 프로세스가 실행 중이라 관리폴더를 변경할 수 없습니다. 다른 Revit 세션이 이전 관리폴더에 추적 기록을 계속 쓰지 않도록 먼저 모두 닫으세요.")
					+ Environment.NewLine + Environment.NewLine
					+ T("Other Revit processes: ", "다른 Revit 프로세스: ") + otherRevitProcessCount.ToString(),
					T("Management Folder Switch Blocked", "관리폴더 변경 차단"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				RenderDashboard();
				return;
			}
			int uncommittedTrackingSessions = FamilyBrowserElementChangeTrackingService.GetActiveUncommittedSessionCount();
			if (uncommittedTrackingSessions > 0)
			{
				ShowDashboardMessage(this,
					T(
						"The management folder cannot change while an open Revit document still has tracked activity that has not reached a successful Save or Synchronize boundary. Save or synchronize every affected project, then try again. No management path or tracking record was changed.",
						"열려 있는 Revit 문서에 저장 또는 동기화 성공 경계에 도달하지 않은 추적 작업이 남아 있어 관리폴더를 변경할 수 없습니다. 해당 프로젝트를 모두 저장하거나 동기화한 뒤 다시 시도하세요. 관리 경로와 추적 기록은 변경하지 않았습니다.")
					+ Environment.NewLine + Environment.NewLine
					+ T("Affected open sessions: ", "해당 열린 세션: ") + uncommittedTrackingSessions.ToString(),
					T("Management Folder Switch Blocked", "관리폴더 변경 차단"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				RenderDashboard();
				return;
			}
			FamilyBrowserHomepageManagedFolderProbeResult probe = ProbeHomepageManagedFolderForReturn();
			if (probe == null || !probe.Available || string.IsNullOrWhiteSpace(probe.ManagedPolicyPath))
			{
				string issue = probe == null ? string.Empty : probe.Issue;
				ShowDashboardMessage(this,
					T("The homepage management folder is not reachable. Nothing was changed.", "홈페이지 관리폴더에 연결할 수 없어 아무것도 변경하지 않았습니다.") + (string.IsNullOrWhiteSpace(issue) ? string.Empty : Environment.NewLine + Environment.NewLine + issue),
					T("Management Folder Switch Stopped", "관리폴더 변경 중단"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				RenderDashboard();
				return;
			}
			string sourceRoot = ResolveTestManagedFolderRoot();
			string destinationRoot = probe.ManagedRootPath;
			FamilyBrowserTrackingFlushResult sourceTrackingFlush = FamilyBrowserTrackingPersistenceService.FlushPending(_workspaceRoot);
			if (sourceTrackingFlush == null || sourceTrackingFlush.FailedCount > 0 || sourceTrackingFlush.DestinationMismatchCount > 0)
			{
				ShowDashboardMessage(this,
					T(
						"The management folder cannot change because locally protected tracking records could not be settled in the current TEST folder. Resolve the pending tracking warning and try again. No management path or tracking record was changed.",
						"로컬에 안전 보관된 추적 기록을 현재 TEST 관리폴더에 확정하지 못해 관리폴더를 변경할 수 없습니다. 전송 대기 추적 경고를 해결한 뒤 다시 시도하세요. 관리 경로와 추적 기록은 변경하지 않았습니다.")
					+ Environment.NewLine + Environment.NewLine
					+ T("Write or recovery failures: ", "쓰기 또는 복구 실패: ") + (sourceTrackingFlush == null ? "1" : sourceTrackingFlush.FailedCount.ToString())
					+ T(" / Records bound to another management folder: ", " / 다른 관리폴더에 연결된 기록: ") + (sourceTrackingFlush == null ? "0" : sourceTrackingFlush.DestinationMismatchCount.ToString()),
					T("Management Folder Switch Blocked", "관리폴더 변경 차단"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				RenderDashboard();
				return;
			}
			int invalidLocalSaveCheckpoints = FamilyBrowserTrackingPersistenceService.GetInvalidElementSessionCheckpointCount();
			if (invalidLocalSaveCheckpoints > 0)
			{
				ShowDashboardMessage(this,
					T(
						"The management folder cannot change while a local-save tracking checkpoint is locked, corrupt, or otherwise requires recovery. Close the other Revit process if it is using the checkpoint, then refresh and try again. No management path or tracking record was changed.",
						"로컬 저장 추적 체크포인트가 잠겨 있거나 손상되었거나 복구가 필요한 상태라 관리폴더를 변경할 수 없습니다. 다른 Revit 프로세스가 체크포인트를 사용 중이면 닫고 새로고침한 뒤 다시 시도하세요. 관리 경로와 추적 기록은 변경하지 않았습니다.")
					+ Environment.NewLine + Environment.NewLine
					+ T("Recovery required: ", "복구 필요: ") + invalidLocalSaveCheckpoints.ToString(),
					T("Management Folder Switch Blocked", "관리폴더 변경 차단"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				RenderDashboard();
				return;
			}
			if (!migrateExistingData)
			{
				int pendingLocalSaveCheckpoints = FamilyBrowserTrackingPersistenceService.GetPendingElementSessionCheckpointCount(_workspaceRoot);
				if (pendingLocalSaveCheckpoints > 0)
				{
					ShowDashboardMessage(this,
						T(
							"Switch Only is blocked because this PC still has workshared local-save tracking checkpoints bound to the TEST management folder. Open the same local RVT files and complete Synchronize with Central, or choose Migrate Existing Data and Switch. No path or tracking record was changed.",
							"이 PC에 TEST 관리폴더와 연결된 워크셰어링 로컬 저장 추적 체크포인트가 남아 있어 경로만 변경할 수 없습니다. 같은 로컬 RVT 파일을 열어 센트럴과 동기화를 완료하거나 '기존 데이터 이관 후 변경'을 선택하세요. 경로와 추적 기록은 변경하지 않았습니다.")
						+ Environment.NewLine + Environment.NewLine
						+ T("Synchronization pending: ", "동기화 대기: ") + pendingLocalSaveCheckpoints.ToString(),
						T("Management Folder Switch Blocked", "관리폴더 변경 차단"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					RenderDashboard();
					return;
				}
			}
			FamilyBrowserManagedFolderMigrationAnalysis analysis = null;
			if (migrateExistingData)
			{
				analysis = FamilyBrowserManagedFolderSetupService.AnalyzeMigration(sourceRoot, probe.ManagedPolicyPath);
				if (analysis == null || !analysis.CanMigrate)
				{
					string conflictText = BuildManagedFolderConflictText(analysis);
					ShowDashboardMessage(this,
						T("Existing TEST data cannot be migrated because the homepage folder already contains different managed data. Nothing was overwritten. You can review the conflicts or use Switch Only to keep the existing homepage data.", "홈페이지 폴더에 서로 다른 관리 데이터가 이미 있어 TEST 데이터를 이관할 수 없습니다. 기존 파일은 덮어쓰지 않았습니다. 충돌 내용을 확인하거나 기존 홈페이지 데이터를 유지하는 경로만 변경을 사용하세요.") + conflictText,
						T("Management Data Conflict", "관리 데이터 충돌"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					RenderDashboard();
					return;
				}
			}
			string confirmMessage;
			if (migrateExistingData)
			{
				confirmMessage = T(
					"Copy TEST management data to the homepage folder, then switch this PC to the homepage path?",
					"TEST 관리 데이터를 홈페이지 폴더로 복사한 뒤 이 PC의 관리폴더를 홈페이지 경로로 변경할까요?")
					+ Environment.NewLine + Environment.NewLine + T("TEST source", "TEST 원본") + Environment.NewLine + sourceRoot
					+ Environment.NewLine + Environment.NewLine + T("Homepage destination", "홈페이지 대상") + Environment.NewLine + destinationRoot
					+ Environment.NewLine + Environment.NewLine + T("Files to copy: ", "복사할 파일: ") + analysis.CopyFileCount.ToString()
					+ T(" / Existing identical files: ", " / 이미 동일한 파일: ") + analysis.AlreadyPresentCount.ToString()
					+ Environment.NewLine + Environment.NewLine + T("The TEST source folder will not be deleted. Existing different homepage files will never be overwritten.", "TEST 원본 폴더는 삭제하지 않습니다. 홈페이지의 기존 다른 파일도 절대 덮어쓰지 않습니다.");
			}
			else
			{
				confirmMessage = T(
					"Switch this PC to the homepage management folder without copying TEST data?",
					"TEST 데이터를 복사하지 않고 이 PC의 관리폴더를 홈페이지 경로로 변경할까요?")
					+ Environment.NewLine + Environment.NewLine + T("Current TEST folder", "현재 TEST 폴더") + Environment.NewLine + sourceRoot
					+ Environment.NewLine + Environment.NewLine + T("Homepage folder", "홈페이지 폴더") + Environment.NewLine + destinationRoot
					+ Environment.NewLine + Environment.NewLine + T("The TEST folder remains unchanged but will no longer be the active management folder.", "TEST 폴더의 데이터는 그대로 남지만 더 이상 활성 관리폴더로 사용하지 않습니다.");
			}
			if (ShowDashboardChoiceMessage(this, confirmMessage, migrateExistingData ? T("Migrate and Switch", "데이터 이관 후 변경") : T("Switch Management Folder", "관리폴더 경로 변경"), MessageBoxIcon.Question, migrateExistingData ? T("Migrate and Switch", "이관 후 변경") : T("Switch", "변경"), T("Cancel", "취소")) != DialogResult.Yes)
			{
				return;
			}

			FamilyBrowserManagedFolderMigrationResult migrationResult = null;
			if (migrateExistingData)
			{
				migrationResult = FamilyBrowserManagedFolderSetupService.MigrateToHomepage(sourceRoot, probe.ManagedPolicyPath, Environment.UserName);
				if (migrationResult == null || !migrationResult.Success)
				{
					string issue = migrationResult == null ? string.Empty : migrationResult.Issue;
					ShowDashboardMessage(this,
						T("The migration did not complete, so the active management folder was not changed. The TEST source remains available.", "데이터 이관이 완료되지 않아 활성 관리폴더는 변경하지 않았습니다. TEST 원본은 그대로 사용할 수 있습니다.") + (string.IsNullOrWhiteSpace(issue) ? string.Empty : Environment.NewLine + Environment.NewLine + issue),
						T("Migration Stopped", "데이터 이관 중단"), MessageBoxButtons.OK, MessageBoxIcon.Hand);
					RenderDashboard();
					return;
				}
			}

			FamilyBrowserDeploymentBootstrapService.SetBootstrapUrl(string.Empty, Environment.UserName);
			FamilyBrowserDeploymentBootstrapService.ClearCache();
			FamilyBrowserDeploymentBootstrapResult applied = FamilyBrowserDeploymentBootstrapService.TryApply(_workspaceRoot, Environment.UserName, force: true, BuildDeploymentProjectIdentity(GetActiveDocument()));
			if (applied == null || !applied.Applied || !SameManagedFolderPath(applied.ManagedPolicyPath, probe.ManagedPolicyPath))
			{
				string rollbackIssue;
				FamilyBrowserManagedFolderSetupService.TryApplyPersistedOverride(Environment.UserName, out rollbackIssue);
				string details = applied == null ? string.Empty : applied.Message;
				if (!string.IsNullOrWhiteSpace(rollbackIssue))
				{
					details = details + Environment.NewLine + rollbackIssue;
				}
				ShowDashboardMessage(this,
					T("The homepage path could not be activated, so the TEST management folder was restored.", "홈페이지 경로를 활성화하지 못해 TEST 관리폴더 설정을 복원했습니다.") + (string.IsNullOrWhiteSpace(details) ? string.Empty : Environment.NewLine + Environment.NewLine + details),
					T("Management Folder Switch Failed", "관리폴더 변경 실패"), MessageBoxButtons.OK, MessageBoxIcon.Hand);
				RenderDashboard();
				return;
			}
			FamilyBrowserTrackingFlushResult trackingFlush = migrateExistingData
				? FamilyBrowserTrackingPersistenceService.FlushPendingForManagedFolderTransition(_workspaceRoot, sourceRoot)
				: FamilyBrowserTrackingPersistenceService.FlushPending(_workspaceRoot);
			if (trackingFlush == null || trackingFlush.FailedCount > 0 || trackingFlush.DestinationMismatchCount > 0 || trackingFlush.ElementSessionCheckpointRebindFailedCount > 0)
			{
				string rollbackIssue;
				bool rollbackApplied = FamilyBrowserManagedFolderSetupService.TryApplyPersistedOverride(Environment.UserName, out rollbackIssue);
				FamilyBrowserTrackingFlushResult rollbackTracking = null;
				if (rollbackApplied && migrateExistingData)
				{
					rollbackTracking = FamilyBrowserTrackingPersistenceService.FlushPendingForManagedFolderTransition(_workspaceRoot, destinationRoot);
				}
				string recoveryText = rollbackApplied && (rollbackTracking == null || (rollbackTracking.FailedCount == 0 && rollbackTracking.DestinationMismatchCount == 0))
					? T("The TEST management folder was restored. Local evidence remains protected.", "TEST 관리폴더를 복원했습니다. 로컬 증거 기록은 안전하게 보존되어 있습니다.")
					: T("Automatic rollback could not be fully verified. Do not delete the local KKY Family Browser tracking cache; contact the administrator with the Debug Log.", "자동 롤백을 완전히 확인하지 못했습니다. 로컬 KKY 패밀리 브라우저 추적 캐시를 삭제하지 말고 디버그 로그와 함께 관리자에게 문의하세요.");
				ShowDashboardMessage(this,
					T("Tracking recovery did not complete, so the homepage management-folder switch was stopped.", "추적 기록 복구가 완료되지 않아 홈페이지 관리폴더 변경을 중단했습니다.")
					+ Environment.NewLine + Environment.NewLine + recoveryText
					+ Environment.NewLine + Environment.NewLine
					+ T("Write or recovery failures: ", "쓰기 또는 복구 실패: ") + (trackingFlush == null ? "1" : trackingFlush.FailedCount.ToString())
					+ T(" / Destination mismatches: ", " / 대상 경로 불일치: ") + (trackingFlush == null ? "0" : trackingFlush.DestinationMismatchCount.ToString())
					+ T(" / Checkpoint rebind failures: ", " / 체크포인트 경로 변경 실패: ") + (trackingFlush == null ? "0" : trackingFlush.ElementSessionCheckpointRebindFailedCount.ToString())
					+ (string.IsNullOrWhiteSpace(rollbackIssue) ? string.Empty : Environment.NewLine + Environment.NewLine + rollbackIssue),
					T("Management Folder Switch Rolled Back", "관리폴더 변경 롤백"), MessageBoxButtons.OK, MessageBoxIcon.Hand);
				RenderDashboard();
				return;
			}
			string pointerIssue;
			if (!FamilyBrowserManagedFolderSetupService.TryClearOverrideRoot(out pointerIssue))
			{
				string rollbackIssue;
				bool rollbackApplied = FamilyBrowserManagedFolderSetupService.TryApplyPersistedOverride(Environment.UserName, out rollbackIssue);
				FamilyBrowserTrackingFlushResult rollbackTracking = null;
				if (rollbackApplied && migrateExistingData)
				{
					rollbackTracking = FamilyBrowserTrackingPersistenceService.FlushPendingForManagedFolderTransition(_workspaceRoot, destinationRoot);
				}
				string rollbackTrackingIssue = rollbackTracking != null && (rollbackTracking.FailedCount > 0 || rollbackTracking.DestinationMismatchCount > 0)
					? Environment.NewLine + T("Checkpoint rollback requires recovery. Do not delete the local tracking cache.", "체크포인트 롤백에 복구가 필요합니다. 로컬 추적 캐시를 삭제하지 마세요.")
					: string.Empty;
				ShowDashboardMessage(this,
					T("The homepage path was reached, but the local TEST pointer could not be removed. The TEST setting was restored to avoid an ambiguous startup state.", "홈페이지 경로에는 연결했지만 로컬 TEST 포인터를 제거하지 못했습니다. 다음 시작 상태가 꼬이지 않도록 TEST 설정을 복원했습니다.") + Environment.NewLine + Environment.NewLine + pointerIssue + (string.IsNullOrWhiteSpace(rollbackIssue) ? string.Empty : Environment.NewLine + rollbackIssue) + rollbackTrackingIssue,
					T("TEST Pointer Removal Failed", "TEST 포인터 제거 실패"), MessageBoxButtons.OK, MessageBoxIcon.Hand);
				RenderDashboard();
				return;
			}

			_deploymentBootstrapResult = applied;
			_homepageManagedFolderProbeResult = null;
			InvalidatePreparedDashboardDataAfterManagedPolicyChange(migrateExistingData ? "managed-folder-migrated-to-homepage" : "managed-folder-switched-to-homepage");
			_standardPolicy = LoadStandardPolicy();
			FamilyBrowserNativeCommandGuardService.NotifyPolicyChanged();
			RefreshManagedFolderAvailabilityState(queueOnboarding: false);
			RefreshDocumentShellOnly();
			_statusMessage = migrateExistingData
				? T("TEST data was migrated and the homepage management folder is now active.", "TEST 데이터를 이관하고 홈페이지 관리폴더로 변경했습니다.")
				: T("The homepage management folder is now active.", "홈페이지 관리폴더로 변경했습니다.");
			RenderDashboard();
			string successMessage = T("Active management folder", "활성 관리폴더") + Environment.NewLine + destinationRoot
				+ Environment.NewLine + Environment.NewLine + T("The TEST source folder was retained.", "TEST 원본 폴더는 그대로 보존했습니다.");
			if (migrationResult != null)
			{
				successMessage += Environment.NewLine + Environment.NewLine + T("Copied files: ", "복사한 파일: ") + migrationResult.CopiedFileCount.ToString()
					+ T(" / Rebased JSON files: ", " / 경로를 재작성한 JSON: ") + migrationResult.RebasedJsonFileCount.ToString()
					+ T(" / Skipped diagnostic conflicts: ", " / 건너뛴 진단 로그 충돌: ") + migrationResult.SkippedDiagnosticConflictCount.ToString();
			}
			if (trackingFlush != null && trackingFlush.ElementSessionCheckpointReboundCount > 0)
			{
				successMessage += Environment.NewLine + T("Rebound protected local-save checkpoints: ", "보호된 로컬 저장 체크포인트 경로 변경: ") + trackingFlush.ElementSessionCheckpointReboundCount.ToString();
			}
			ShowDashboardMessage(this, successMessage, migrateExistingData ? T("Migration and Switch Complete", "데이터 이관 및 변경 완료") : T("Management Folder Switched", "관리폴더 변경 완료"), MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
		catch (Exception ex)
		{
			ShowLoggedError(T("Switch Homepage Management Folder", "홈페이지 관리폴더로 변경"), ex);
		}
		finally
		{
			if (trackingTransitionAuthorization != null)
			{
				trackingTransitionAuthorization.Dispose();
			}
			if (managementContextLease != null)
			{
				managementContextLease.Dispose();
			}
		}
	}

	private static int GetOtherRevitProcessCount()
	{
		int currentProcessId;
		using (Process currentProcess = Process.GetCurrentProcess())
		{
			currentProcessId = currentProcess.Id;
		}
		int count = 0;
		Process[] processes = null;
		try
		{
			processes = Process.GetProcessesByName("Revit");
			foreach (Process process in processes)
			{
				try
				{
					if (process != null && process.Id != currentProcessId && !process.HasExited)
					{
						count++;
					}
				}
				catch
				{
					count++;
				}
			}
		}
		catch
		{
			return 1;
		}
		finally
		{
			if (processes != null)
			{
				foreach (Process process in processes)
				{
					if (process != null)
					{
						process.Dispose();
					}
				}
			}
		}
		return count;
	}

	private static string BuildManagedFolderConflictText(FamilyBrowserManagedFolderMigrationAnalysis analysis)
	{
		if (analysis == null)
		{
			return string.Empty;
		}
		StringBuilder sb = new StringBuilder();
		if (!string.IsNullOrWhiteSpace(analysis.Issue))
		{
			sb.AppendLine().AppendLine().Append(analysis.Issue);
		}
		if (analysis.BlockingConflicts != null && analysis.BlockingConflicts.Count > 0)
		{
			sb.AppendLine().AppendLine();
			int limit = Math.Min(8, analysis.BlockingConflicts.Count);
			for (int i = 0; i < limit; i++)
			{
				sb.AppendLine("- " + analysis.BlockingConflicts[i]);
			}
			if (analysis.BlockingConflicts.Count > limit)
			{
				sb.Append("... +").Append((analysis.BlockingConflicts.Count - limit).ToString());
			}
		}
		return sb.ToString();
	}

	private static bool SameManagedFolderPath(string left, string right)
	{
		try
		{
			string leftPath = Path.GetFullPath(Environment.ExpandEnvironmentVariables((left ?? string.Empty).Trim())).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
			string rightPath = Path.GetFullPath(Environment.ExpandEnvironmentVariables((right ?? string.Empty).Trim())).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
			return string.Equals(leftPath, rightPath, StringComparison.OrdinalIgnoreCase);
		}
		catch
		{
			return string.Equals((left ?? string.Empty).Trim(), (right ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase);
		}
	}

	private void AppendManagedFolderTransitionPanel(StringBuilder sb, string elementId, bool compact)
	{
		if (!IsTestManagedFolderOverrideActive())
		{
			return;
		}
		FamilyBrowserHomepageManagedFolderProbeResult probe = _homepageManagedFolderProbeResult;
		bool homepageAvailable = probe != null && probe.Available;
		string sourceRoot = ResolveTestManagedFolderRoot();
		string targetRoot = homepageAvailable ? probe.ManagedRootPath : string.Empty;
		sb.AppendLine("<div id=\"" + Attr(elementId) + "\" class=\"managed-folder-recovery managed-folder-transition" + (homepageAvailable ? " return-ready" : string.Empty) + (compact ? " compact" : string.Empty) + "\" role=\"status\">");
		sb.AppendLine("<div class=\"managed-folder-recovery-copy\"><span>" + Html(T("TEST MANAGEMENT FOLDER ACTIVE", "TEST 관리폴더 사용 중")) + "</span><strong>" + Html(homepageAvailable ? T("A homepage management folder is available.", "홈페이지 관리폴더를 사용할 수 있습니다.") : T("This PC is still using the TEST management folder.", "이 PC는 현재 TEST 관리폴더를 사용 중입니다.")) + "</strong><em>" + Html(homepageAvailable ? T("Switch only, or migrate the existing TEST data without overwriting homepage data and then switch.", "경로만 변경하거나, 홈페이지 데이터를 덮어쓰지 않고 기존 TEST 데이터를 이관한 뒤 변경할 수 있습니다.") : T("Refresh to check whether the homepage now provides a reachable management folder.", "새로고침하여 홈페이지에 접근 가능한 관리폴더가 새로 등록되었는지 확인하세요.")) + "</em></div>");
		sb.AppendLine("<div class=\"managed-folder-transition-paths\"><div><span>" + Html(T("TEST source", "TEST 원본")) + "</span><strong title=\"" + Attr(sourceRoot) + "\">" + Html(sourceRoot) + "</strong></div>" + (homepageAvailable ? "<div><span>" + Html(T("Homepage destination", "홈페이지 대상")) + "</span><strong title=\"" + Attr(targetRoot) + "\">" + Html(targetRoot) + "</strong></div>" : string.Empty) + "</div>");
		sb.AppendLine("<div class=\"managed-folder-recovery-actions\"><a class=\"tool\" href=\"kkyfb:managed-folder-retry\">" + Html(T("Check Homepage Path", "홈페이지 경로 확인")) + "</a>" + (homepageAvailable ? "<a class=\"tool primary\" href=\"kkyfb:managed-folder-switch-homepage\">" + Html(T("Switch to Homepage Folder", "홈페이지 경로로 변경")) + "</a>" : string.Empty) + (homepageAvailable && _adminModeEnabled ? "<a class=\"tool\" href=\"kkyfb:managed-folder-migrate-homepage\">" + Html(T("Migrate Existing Data and Switch", "기존 데이터 이관 후 변경")) + "</a>" : string.Empty) + "</div>");
		sb.AppendLine("<div class=\"managed-folder-recovery-warning\">" + Html(T("Migration never deletes the TEST source and never overwrites different files already stored in the homepage folder.", "이관 시 TEST 원본은 삭제하지 않으며 홈페이지 폴더에 이미 있는 다른 파일은 덮어쓰지 않습니다.")) + "</div>");
		sb.AppendLine("</div>");
	}
}
