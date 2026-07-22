using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserErrorHelp
{
	private FamilyBrowserErrorHelp()
	{
	}

	public static FamilyBrowserFriendlyError Build(string caption, Exception ex, string logPath, bool korean)
	{
		if (ex == null)
		{
			ex = new InvalidOperationException(korean ? "알 수 없는 오류가 발생했습니다." : "An unknown error occurred.");
		}
		string supportCode = "FBR-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-fff", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N").Substring(0, 4).ToUpperInvariant();
		string context = ((!string.IsNullOrWhiteSpace(caption)) ? caption.Trim() : (korean ? "작업" : "Action"));
		Exception mostRelevant = FindMostRelevantException(ex);
		string allText = BuildSearchText(ex);
		FamilyBrowserFriendlyError friendly = new FamilyBrowserFriendlyError
		{
			Title = (korean ? (context + " 작업을 완료하지 못했습니다.") : (context + " did not complete.")),
			TechnicalDetail = mostRelevant.GetType().Name + ": " + mostRelevant.Message,
			LogPath = (logPath ?? string.Empty),
			SupportCode = supportCode
		};
		if (ContainsAny(allText, "managed shared folder", "managed folder path", "local c fallback", "공용 관리 폴더", "홈페이지 경로를 다시 확인", "family browser 정책 저장"))
		{
			ApplyManagedFolderUnavailable(friendly, korean);
		}
		else if (mostRelevant is UnauthorizedAccessException || ContainsAny(allText, "access to the path is denied", "unauthorized", "permission denied", "권한"))
		{
			ApplyAccessDenied(friendly, korean);
		}
		else if (mostRelevant is FileNotFoundException || mostRelevant is DirectoryNotFoundException || ContainsAny(allText, "not found", "cannot be found", "찾을 수 없습니다", "path is empty"))
		{
			ApplyMissingPath(friendly, korean, allText);
		}
		else if (mostRelevant is TimeoutException || ContainsAny(allText, "timed out", "timeout", "pipe", "bridge is not connected", "not connected"))
		{
			ApplyConnectionIssue(friendly, korean);
		}
		else if (mostRelevant is TypeLoadException || mostRelevant is MissingMethodException || ContainsAny(allText, "could not load type", "could not load file or assembly", "missingmethodexception", "typeloadexception", "system.text.json"))
		{
			ApplyDependencyIssue(friendly, korean);
		}
		else if (mostRelevant is IOException || ContainsAny(allText, "being used by another process", "sharing violation", "network name", "i/o"))
		{
			ApplyFileBusyOrNetwork(friendly, korean);
		}
		else if (ContainsAny(allText, "request store", "network request", "sharepoint", "cloud", "connector", "api queue"))
		{
			ApplyRequestStoreIssue(friendly, korean);
		}
		else if (ContainsAny(allText, "standard_rvt_fingerprint_failed", "standard rvt fingerprint", "standard loadable family fingerprint", "standard family fingerprint", "standard content fingerprint"))
		{
			ApplyStandardFingerprintIssue(friendly, korean);
		}
		else if (ContainsAny(allText, "standard rvt", "standard library", "snapshot", "register the standard", "registered standard"))
		{
			ApplyStandardSourceIssue(friendly, korean);
		}
		else if (ContainsAny(allText, "system type", "routing", "dependency family", "canonical", "destination type"))
		{
			ApplySystemTypeIssue(friendly, korean);
		}
		else if (ContainsAny(allText, "transaction", "modifiable", "active document", "open a revit project", "document"))
		{
			ApplyRevitDocumentStateIssue(friendly, korean);
		}
		else
		{
			ApplyGenericIssue(friendly, korean);
		}
		return friendly;
	}

	public static string WriteLog(string workspaceRoot, string caption, Exception ex, string extraText = "")
	{
		string root = (string.IsNullOrWhiteSpace(workspaceRoot) ? HostWorkspacePathResolver.ResolveRoot() : workspaceRoot);
		try
		{
			string dataFolder;
			if (FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(root))
			{
				dataFolder = FamilyBrowserStandardPolicyStore.GetDataFolder(root, "Logs");
			}
			else
			{
				dataFolder = ResolveLocalDiagnosticLogFolder();
			}
			return WriteDiagnosticLogFile(dataFolder, root, caption, ex, extraText);
		}
		catch (Exception primaryError)
		{
			try
			{
				string fallbackContext = (extraText ?? string.Empty) + Environment.NewLine + "Primary log write failure: " + primaryError;
				return WriteDiagnosticLogFile(ResolveLocalDiagnosticLogFolder(), root, caption, ex, fallbackContext);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
				return "(log write failed)";
			}
		}
	}

	private static string ResolveLocalDiagnosticLogFolder()
	{
		string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
		return Path.Combine(localAppData, "KKY", "FamilyBrowser", "Diagnostics", "Errors");
	}

	private static string WriteDiagnosticLogFile(string dataFolder, string workspaceRoot, string caption, Exception ex, string extraText)
	{
		Directory.CreateDirectory(dataFolder);
		string fileName = DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-fffffff", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N").Substring(0, 8) + "-" + SafeFileName(caption) + ".log";
		string path = Path.Combine(dataFolder, fileName);
		StringBuilder builder = new StringBuilder();
		builder.AppendLine("Timestamp: " + DateTime.Now.ToString("O", CultureInfo.InvariantCulture));
		builder.AppendLine("Action: " + (caption ?? string.Empty));
		builder.AppendLine("Machine: " + Environment.MachineName);
		builder.AppendLine("User: " + Environment.UserName);
		builder.AppendLine("Workspace: " + (workspaceRoot ?? string.Empty));
		if (!string.IsNullOrWhiteSpace(extraText))
		{
			builder.AppendLine();
			builder.AppendLine("Context");
			builder.AppendLine(extraText);
		}
		builder.AppendLine();
		builder.AppendLine("Exception");
		builder.AppendLine((ex == null) ? "(no exception)" : ex.ToString());
		string temporaryPath = FamilyBrowserAtomicFileService.CreateSiblingTemporaryPath(path);
		try
		{
			byte[] payload = new UTF8Encoding(false).GetBytes(builder.ToString());
			using (FileStream stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
			{
				stream.Write(payload, 0, payload.Length);
				stream.Flush(true);
			}
			FamilyBrowserAtomicFileService.Promote(temporaryPath, path);
			return path;
		}
		finally
		{
			try
			{
				if (File.Exists(temporaryPath))
				{
					File.Delete(temporaryPath);
				}
			}
			catch
			{
			}
		}
	}

	private static void ApplyManagedFolderUnavailable(FamilyBrowserFriendlyError errorInfo, bool korean)
	{
		if (korean)
		{
			errorInfo.Summary = "공용 관리 폴더가 연결되지 않았습니다.";
			errorInfo.Cause = "홈페이지에서 지정한 Family Browser 관리 폴더를 현재 PC에서 열 수 없어 정책과 운영 데이터를 저장하지 못했습니다.";
			errorInfo.UserAction = "사내망/VPN과 네트워크 드라이브 연결을 확인한 뒤 새로고침 또는 홈페이지 경로 다시 확인을 실행하세요.";
			errorInfo.AdminAction = "홈페이지 관리 경로 후보가 모든 PC에서 접근 가능한 UNC 경로인지 확인하고, 현재 Windows 계정의 읽기/쓰기 권한을 점검하세요.";
		}
		else
		{
			errorInfo.Summary = "The managed shared folder is not connected.";
			errorInfo.Cause = "This PC cannot reach the Family Browser management folder provided by the homepage, so policy and operational data could not be saved.";
			errorInfo.UserAction = "Check the corporate network, VPN, and mapped drive, then refresh or run Check Homepage Path Again.";
			errorInfo.AdminAction = "Use a UNC path reachable from every PC and verify read/write access for the current Windows account.";
		}
	}

	private static void ApplyAccessDenied(FamilyBrowserFriendlyError errorInfo, bool korean)
	{
		if (korean)
		{
			errorInfo.Summary = "권한 또는 폴더 접근 문제입니다.";
			errorInfo.Cause = "현재 Windows 계정이 필요한 파일, 표준 RVT, 정책 파일, 요청 저장소, 또는 출력 폴더에 접근하지 못했습니다.";
			errorInfo.UserAction = "사내망/VPN 연결을 확인하고, 같은 파일이 열려 있다면 닫은 뒤 다시 실행하세요. 권한이 필요한 작업이면 관리자에게 접근 권한을 요청하세요.";
			errorInfo.AdminAction = "공유 폴더의 NTFS/공유 권한, 정책 Excel/JSON 경로, 요청 저장소 경로, 표준 RVT 경로에 현재 사용자 또는 그룹 권한이 있는지 확인하세요.";
		}
		else
		{
			errorInfo.Summary = "This looks like a permission or folder access issue.";
			errorInfo.Cause = "The current Windows account could not access a required file, standard RVT, policy file, request store, or output folder.";
			errorInfo.UserAction = "Check network/VPN access, close the same file if it is open, then try again. Ask an administrator for access if this is a restricted workflow.";
			errorInfo.AdminAction = "Check NTFS/share permissions for the policy path, request store, standard RVT path, and output folder for the current user or group.";
		}
	}

	private static void ApplyMissingPath(FamilyBrowserFriendlyError errorInfo, bool korean, string allText)
	{
		if (korean)
		{
			errorInfo.Summary = "필요한 파일 또는 폴더를 찾지 못했습니다.";
			errorInfo.Cause = ((allText.IndexOf("standard", StringComparison.OrdinalIgnoreCase) >= 0) ? "등록된 표준 RVT 또는 표준 스냅샷 경로가 현재 PC에서 유효하지 않을 수 있습니다." : "정책 파일, 요청 저장소, 첨부파일, 또는 출력 경로가 이동/삭제되었거나 현재 PC에서 보이지 않습니다.");
			errorInfo.UserAction = "파일 경로가 실제로 존재하는지 확인하고, 네트워크 드라이브나 동기화 폴더가 연결되어 있는지 확인한 뒤 다시 실행하세요.";
			errorInfo.AdminAction = "관리자 설정의 표준 RVT 경로, 공용 정책 JSON 경로, 요청 저장소 경로가 절대 경로 또는 UNC 경로로 유효한지 확인하세요.";
		}
		else
		{
			errorInfo.Summary = "A required file or folder was not found.";
			errorInfo.Cause = ((allText.IndexOf("standard", StringComparison.OrdinalIgnoreCase) >= 0) ? "The registered standard RVT or snapshot path may not be valid on this PC." : "A policy file, request store, attachment, or output path may have moved, been deleted, or be unavailable.");
			errorInfo.UserAction = "Confirm the path exists and that the network drive or sync folder is connected, then try again.";
			errorInfo.AdminAction = "Check that the standard RVT path, managed policy JSON path, and request store path are valid absolute or UNC paths.";
		}
	}

	private static void ApplyConnectionIssue(FamilyBrowserFriendlyError errorInfo, bool korean)
	{
		if (korean)
		{
			errorInfo.Summary = "Revit 연결 또는 응답 대기 문제가 발생했습니다.";
			errorInfo.Cause = "데스크톱 앱이 Revit 애드인 브리지와 연결되지 않았거나, Revit이 작업 처리 중이라 제한 시간 안에 응답하지 못했습니다.";
			errorInfo.UserAction = "Revit에서 프로젝트를 열고 Family Browser 애드인이 로드되어 있는지 확인한 뒤 다시 시도하세요. Revit이 바쁜 상태라면 잠시 기다린 뒤 실행하세요.";
			errorInfo.AdminAction = "애드인 설치 경로, 버전별 addin 파일, 방화벽/보안 프로그램이 로컬 named pipe 통신을 차단하는지 확인하세요.";
		}
		else
		{
			errorInfo.Summary = "Revit connection or response timeout.";
			errorInfo.Cause = "The desktop app is not connected to the Revit add-in bridge, or Revit did not respond before the timeout.";
			errorInfo.UserAction = "Open a project in Revit, confirm the Family Browser add-in is loaded, then retry. If Revit is busy, wait and try again.";
			errorInfo.AdminAction = "Check add-in installation paths, version-specific addin files, and whether security software blocks local named pipes.";
		}
	}

	private static void ApplyDependencyIssue(FamilyBrowserFriendlyError errorInfo, bool korean)
	{
		if (korean)
		{
			errorInfo.Summary = "애드인 DLL 또는 .NET 의존성 버전이 맞지 않을 가능성이 있습니다.";
			errorInfo.Cause = "Revit이 필요한 타입이나 메서드를 찾지 못했습니다. 보통 설치 파일 누락, 다른 버전 DLL 혼재, 또는 Revit 버전과 빌드 대상 불일치일 때 발생합니다.";
			errorInfo.UserAction = "Revit을 완전히 종료한 뒤 애드인을 다시 설치하고, 해당 Revit 버전에 맞는 빌드가 설치되었는지 확인하세요.";
			errorInfo.AdminAction = "설치 패키지에 필요한 DLL이 모두 포함되었는지, 2019/2021/2023/2025별 빌드 산출물이 올바른 addin 경로를 가리키는지 확인하세요.";
		}
		else
		{
			errorInfo.Summary = "The add-in DLL or .NET dependency version may not match.";
			errorInfo.Cause = "Revit could not find a required type or method. This usually means a missing DLL, mixed dependency versions, or a build installed for the wrong Revit version.";
			errorInfo.UserAction = "Close Revit completely, reinstall the add-in, and confirm the build matches the Revit version.";
			errorInfo.AdminAction = "Check that the installer includes all required DLLs and that 2019/2021/2023/2025 addin files point to the correct build output.";
		}
	}

	private static void ApplyFileBusyOrNetwork(FamilyBrowserFriendlyError errorInfo, bool korean)
	{
		if (korean)
		{
			errorInfo.Summary = "파일이 사용 중이거나 네트워크 저장소가 일시적으로 불안정합니다.";
			errorInfo.Cause = "요청 JSON, 정책 파일, 첨부파일, 표준 RVT, 또는 출력 파일을 다른 프로그램이 사용 중이거나 네트워크 연결이 끊겼을 수 있습니다.";
			errorInfo.UserAction = "Excel/Revit/파일 탐색기에서 같은 파일이 열려 있는지 확인하고 닫은 뒤 다시 실행하세요. 네트워크 저장소라면 연결 상태를 확인하세요.";
			errorInfo.AdminAction = "동기화 폴더 충돌, 파일 잠금, 공유 폴더 권한, 백신/보안 프로그램의 파일 쓰기 차단 여부를 확인하세요.";
		}
		else
		{
			errorInfo.Summary = "A file is busy or the network store is temporarily unstable.";
			errorInfo.Cause = "A request JSON, policy file, attachment, standard RVT, or output file may be open in another process, or the network connection may have dropped.";
			errorInfo.UserAction = "Close the same file in Excel/Revit/File Explorer, check network connectivity, then retry.";
			errorInfo.AdminAction = "Check sync conflicts, file locks, share permissions, and whether antivirus/security tools block file writes.";
		}
	}

	private static void ApplyRequestStoreIssue(FamilyBrowserFriendlyError errorInfo, bool korean)
	{
		if (korean)
		{
			errorInfo.Summary = "요청 저장소 설정 문제입니다.";
			errorInfo.Cause = "요청 목록, 첨부파일, 진행 상태를 저장할 위치가 없거나 현재 사용자가 쓸 수 없습니다.";
			errorInfo.UserAction = "관리자 설정에서 요청 저장소가 로컬/네트워크/SharePoint/클라우드/API 중 어떤 방식인지 확인하고, 저장소 열기가 되는지 확인하세요.";
			errorInfo.AdminAction = "요청 저장소 경로와 권한을 확인하세요. 운영 환경에서는 네트워크 공유 또는 동기화 폴더를 권장하며, API 방식은 별도 수집 커넥터가 필요합니다.";
		}
		else
		{
			errorInfo.Summary = "Request store configuration issue.";
			errorInfo.Cause = "The request list, attachments, or progress status cannot be saved because the store is missing or not writable.";
			errorInfo.UserAction = "Check the admin request store mode and confirm the store opens from the Family Browser.";
			errorInfo.AdminAction = "Verify request store path and permissions. Use NetworkShare or synced folder for production; API mode requires a connector to collect queued requests.";
		}
	}

	private static void ApplyStandardSourceIssue(FamilyBrowserFriendlyError errorInfo, bool korean)
	{
		if (korean)
		{
			errorInfo.Summary = "표준 RVT 또는 표준 스냅샷 준비 상태 문제입니다.";
			errorInfo.Cause = "현재 프로젝트를 비교하거나 로드할 기준 RVT가 등록되지 않았거나, 저장된 스냅샷이 없거나, 경로가 변경되었습니다.";
			errorInfo.UserAction = "관리자 설정에서 현재 공종의 표준 RVT가 연결되어 있는지 확인하고, 표준 기준 새로고침 또는 재등록을 실행하세요.";
			errorInfo.AdminAction = "공종별/통합 모드에 맞는 표준 RVT 경로, 스냅샷 파일, 정책 JSON의 등록 정보를 확인하세요.";
		}
		else
		{
			errorInfo.Summary = "Standard RVT or snapshot readiness issue.";
			errorInfo.Cause = "The standard RVT used for comparison/load is not registered, its snapshot is missing, or the path changed.";
			errorInfo.UserAction = "Check the active trade standard RVT in Admin settings and refresh or re-register the standard source.";
			errorInfo.AdminAction = "Verify standard RVT path, snapshot file, and policy JSON registration for trade-separated or integrated mode.";
		}
	}

	private static void ApplyStandardFingerprintIssue(FamilyBrowserFriendlyError errorInfo, bool korean)
	{
		if (korean)
		{
			errorInfo.Summary = "표준 RVT Fingerprint 생성 실패입니다.";
			errorInfo.Cause = "표준 RVT 스캔 중 하나 이상의 로더블 패밀리 Fingerprint가 생성되지 않았습니다. 이 상태의 표준 스냅샷은 기준 비교에 사용할 수 없습니다.";
			errorInfo.UserAction = "표준 RVT 관리에서 표시된 패밀리와 signature 진단 파일을 확인한 뒤 다시 스캔하세요. 같은 오류가 반복되면 해당 패밀리를 직접 열어 경고, 손상, 편집 불가 상태를 확인하세요.";
			errorInfo.AdminAction = "로그의 STANDARD_RVT_FINGERPRINT_FAILED 항목과 나열된 signature 경로를 확인하세요. 표준 RVT 캐시를 재사용 중이었다면 해당 표준 스냅샷을 초기화하고 정밀 스캔을 다시 실행하세요.";
		}
		else
		{
			errorInfo.Summary = "Standard RVT fingerprint creation failed.";
			errorInfo.Cause = "One or more standard loadable family fingerprints were not created during the standard RVT scan. That standard snapshot cannot be trusted for comparison.";
			errorInfo.UserAction = "Open Standard RVT management, inspect the listed families and signature diagnostics, then rescan. If the same item repeats, open that family directly and check warnings, corruption, or edit restrictions.";
			errorInfo.AdminAction = "Review the STANDARD_RVT_FINGERPRINT_FAILED log entry and the listed signature paths. If a cached standard snapshot was reused, reset that standard snapshot and run a precise scan again.";
		}
	}

	private static void ApplySystemTypeIssue(FamilyBrowserFriendlyError errorInfo, bool korean)
	{
		if (korean)
		{
			errorInfo.Summary = "시스템 타입 적용 중 기준 타입 또는 의존 패밀리 정리에 실패했습니다.";
			errorInfo.Cause = "표준 RVT의 시스템 타입, Routing Preference, 의존 로더블 패밀리, 또는 현재 프로젝트의 중복 타입 정리 단계에서 문제가 발생했습니다.";
			errorInfo.UserAction = "시스템 타입 검토를 먼저 실행하고, 표준 RVT가 올바른 공종으로 등록되어 있는지 확인한 뒤 선택 항목만 다시 적용해보세요.";
			errorInfo.AdminAction = "표준 RVT의 시스템 타입 이름, 의존 패밀리 이름/카테고리, 중복 타입, Routing Preference 규칙과 실패 로그 항목을 검토하세요.";
		}
		else
		{
			errorInfo.Summary = "System type apply failed while preparing the canonical type or dependency families.";
			errorInfo.Cause = "The failure happened around standard system types, routing preferences, dependency loadable families, or duplicate type consolidation.";
			errorInfo.UserAction = "Run system type review first, confirm the correct standard RVT is active, then retry only the selected item.";
			errorInfo.AdminAction = "Review standard RVT system type names, dependency family names/categories, duplicates, routing preference rules, and the failed log item.";
		}
	}

	private static void ApplyRevitDocumentStateIssue(FamilyBrowserFriendlyError errorInfo, bool korean)
	{
		if (korean)
		{
			errorInfo.Summary = "현재 Revit 문서 상태 때문에 작업을 실행할 수 없습니다.";
			errorInfo.Cause = "프로젝트가 열려 있지 않거나, 패밀리 문서/저장되지 않은 문서/편집 불가 상태에서 프로젝트용 작업을 실행했을 수 있습니다.";
			errorInfo.UserAction = "대상 프로젝트 RVT를 열고, 진행 중인 명령이나 편집 모드를 종료한 뒤 다시 실행하세요.";
			errorInfo.AdminAction = "사용자가 로컬, 센트럴, 분리 파일, 패밀리 문서, 저장되지 않은 문서 중 어떤 상태에서 실행했는지와 작업 가능 상태를 로그와 함께 확인하세요.";
		}
		else
		{
			errorInfo.Summary = "The current Revit document state does not allow this action.";
			errorInfo.Cause = "A project may not be open, or the command may have been run from a family document, unsaved document, or non-editable state.";
			errorInfo.UserAction = "Open the target project RVT, finish any active command/edit mode, then retry.";
			errorInfo.AdminAction = "Check whether the user ran the action from a local, central, detached, family, or unsaved document.";
		}
	}

	private static void ApplyGenericIssue(FamilyBrowserFriendlyError errorInfo, bool korean)
	{
		if (korean)
		{
			errorInfo.Summary = "예상하지 못한 오류가 발생했습니다.";
			errorInfo.Cause = "현재 작업 중 처리하지 못한 예외가 발생했습니다. 자세한 원인은 로그 파일에 저장되었습니다.";
			errorInfo.UserAction = "같은 작업을 한 번 더 시도해보고, 반복되면 이 창의 지원 코드와 로그 경로를 관리자에게 전달하세요.";
			errorInfo.AdminAction = "로그 파일의 Exception 전체 내용과 실행한 Revit 버전, 현재 프로젝트 경로, 표준 RVT 경로, 직전 사용자 동작을 함께 확인하세요.";
		}
		else
		{
			errorInfo.Summary = "An unexpected error occurred.";
			errorInfo.Cause = "The current workflow hit an exception that was not handled more specifically. Details were written to the log file.";
			errorInfo.UserAction = "Try the same action once more. If it repeats, send the support code and log path to the administrator.";
			errorInfo.AdminAction = "Review the full exception in the log with the Revit version, project path, standard RVT path, and the user action immediately before failure.";
		}
	}

	private static Exception FindMostRelevantException(Exception ex)
	{
		Exception current = ex;
		while (current != null && current.InnerException != null)
		{
			current = current.InnerException;
		}
		return current ?? ex;
	}

	private static string BuildSearchText(Exception ex)
	{
		List<string> parts = new List<string>();
		for (Exception current = ex; current != null; current = current.InnerException)
		{
			parts.Add(current.GetType().FullName);
			parts.Add(current.Message);
		}
		return string.Join(" | ", parts.Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x))).ToLowerInvariant();
	}

	private static bool ContainsAny(string value, params string[] tokens)
	{
		if (string.IsNullOrWhiteSpace(value) || tokens == null)
		{
			return false;
		}
		foreach (string token in tokens)
		{
			if (!string.IsNullOrWhiteSpace(token) && value.IndexOf(token.ToLowerInvariant(), StringComparison.OrdinalIgnoreCase) >= 0)
			{
				return true;
			}
		}
		return false;
	}

	private static string SafeFileName(string value)
	{
		char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
		string safe = new string((value ?? string.Empty).Select([SpecialName] (char ch) => (!Enumerable.Contains(invalidFileNameChars, ch)) ? ch : '_').ToArray()).Trim();
		if (string.IsNullOrWhiteSpace(safe))
		{
			return "FamilyBrowser";
		}
		if (safe.Length > 72)
		{
			safe = safe.Substring(0, 72);
		}
		return safe;
	}
}
