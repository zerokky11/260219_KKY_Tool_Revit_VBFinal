using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2025;

[Transaction(/*Could not decode attribute arguments.*/)]
[Regeneration(/*Could not decode attribute arguments.*/)]
public class CmdApplySystemTypes : IExternalCommand
{
	public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
	{
		//IL_06fb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0720: Unknown result type (might be due to invalid IL or missing references)
		//IL_002d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0068: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_024e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0255: Unknown result type (might be due to invalid IL or missing references)
		//IL_0279: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_02fc: Expected O, but got Unknown
		//IL_02fe: Unknown result type (might be due to invalid IL or missing references)
		//IL_0313: Unknown result type (might be due to invalid IL or missing references)
		//IL_06da: Unknown result type (might be due to invalid IL or missing references)
		//IL_06e1: Unknown result type (might be due to invalid IL or missing references)
		Document standardDoc = null;
		bool shouldCloseStandardDoc = false;
		Result Execute;
		try
		{
			UIDocument uiDoc = commandData.Application.ActiveUIDocument;
			if (uiDoc == null || uiDoc.Document == null)
			{
				message = T("Open a project document before applying system type updates.", "시스템 타입을 적용하기 전에 Revit 프로젝트를 열어주세요.");
				Execute = (Result)1;
			}
			else
			{
				string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
				string registryPath = Path.Combine(FamilyBrowserStandardPolicyStore.GetRegistryFolder(workspaceRoot), "active-standard-library.json");
				if (!File.Exists(registryPath))
				{
					message = T("No registered standard RVT was found. Register the approved standard RVT first.", "등록된 표준 RVT가 없습니다. 먼저 승인된 표준 RVT를 등록하세요.");
					Execute = (Result)(-1);
				}
				else
				{
					StandardLibraryRegistrationRecord registration = DataContractJsonFileStore.Load<StandardLibraryRegistrationRecord>(registryPath);
					if (string.IsNullOrWhiteSpace(registration.LastSnapshotPath) || !File.Exists(registration.LastSnapshotPath))
					{
						message = T("The registered standard snapshot could not be found. Re-register or refresh the standard RVT.", "등록된 표준 스냅샷을 찾을 수 없습니다. 표준 RVT를 다시 등록하거나 새로고침하세요.");
						Execute = (Result)(-1);
					}
					else
					{
						EnsureTargetDiffersFromStandard(uiDoc.Document, registration);
						StandardLibrarySnapshot standardSnapshot = DataContractJsonFileStore.Load<StandardLibrarySnapshot>(registration.LastSnapshotPath);
						standardDoc = StandardLibraryDocumentResolver.OpenRegisteredDocument(commandData.Application.Application, registration, ref shouldCloseStandardDoc);
						SystemTypePreflightReport preflightReport = SystemTypePreflightBuilderService.BuildReport(registration, standardDoc, uiDoc.Document);
						UpgradeStandardSnapshotSemantics(standardSnapshot, preflightReport.StandardCatalog);
						string preflightPath = SystemTypePreflightStore.Save(workspaceRoot, preflightReport);
						int executableCount = CountExecutableItems(preflightReport);
						if (executableCount == 0)
						{
							TaskDialog.Show(T("Apply System Types", "시스템 타입 적용"), T("No supported system type actions are ready to run right now.", "지금 바로 적용할 수 있는 지원 시스템 타입 항목이 없습니다.") + "\r\n\r\n" + T("Project: ", "프로젝트: ") + uiDoc.Document.Title + "\r\n" + T("Standard: ", "표준: ") + registration.DisplayName + "\r\n" + T("Ready: ", "준비: ") + preflightReport.Summary.ReadyCount + "\r\n" + T("Approval required: ", "승인 필요: ") + preflightReport.Summary.ApprovalRequiredCount + "\r\n" + T("Blocked: ", "차단: ") + preflightReport.Summary.BlockedCount + "\r\n" + T("Review report: ", "검토 보고서: ") + preflightPath);
							Execute = (Result)0;
						}
						else if (!ConfirmExecution(preflightReport, registration.DisplayName, uiDoc.Document.Title, executableCount))
						{
							Execute = (Result)1;
						}
						else
						{
							SystemTypeApplyExecutionReport executionReport = SystemTypeApplyExecutionService.Execute(uiDoc.Document, standardDoc, preflightReport, preflightPath);
							if (executionReport.Items.Any([SpecialName] (SystemTypeApplyExecutionItem x) => IsSuccessfulApplyOutcome(x.Outcome)))
							{
								ProjectContentSnapshot postProjectSnapshot = ProjectSnapshotCaptureService.Capture(uiDoc.Document);
								ProjectTrackingCatalog refreshedTrackingCatalog = ProjectTrackingStampService.BuildCatalog(registration, standardSnapshot, postProjectSnapshot);
								Transaction tx = new Transaction(uiDoc.Document, "KKY Family Browser Refresh Tracking");
								try
								{
									tx.Start();
									ProjectTrackingStoreService.Save(uiDoc.Document, refreshedTrackingCatalog);
									tx.Commit();
								}
								finally
								{
									((IDisposable)tx)?.Dispose();
								}
								executionReport.TrackingPath = ProjectTrackingOutputStore.Save(workspaceRoot, refreshedTrackingCatalog);
								executionReport.Summary.TrackingRefreshedCount = executionReport.Items.Count([SpecialName] (SystemTypeApplyExecutionItem x) => IsSuccessfulApplyOutcome(x.Outcome));
							}
							SystemTypePreflightReport postPreflightReport = SystemTypePreflightBuilderService.BuildReport(registration, standardDoc, uiDoc.Document);
							executionReport.PostPreflightPath = SystemTypePreflightStore.Save(workspaceRoot, postPreflightReport);
							string executionPath = SystemTypeApplyStore.Save(workspaceRoot, executionReport);
							TaskDialog.Show(T("Apply System Types", "시스템 타입 적용"), T("System type synchronization completed.", "시스템 타입 동기화를 완료했습니다.") + "\r\n\r\n" + T("Project: ", "프로젝트: ") + uiDoc.Document.Title + "\r\n" + T("Standard: ", "표준: ") + registration.DisplayName + "\r\n" + T("Created: ", "생성: ") + executionReport.Summary.CreatedCount + "\r\n" + T("Overwritten: ", "덮어쓰기: ") + executionReport.Summary.OverwrittenCount + "\r\n" + T("Consolidated: ", "통합: ") + executionReport.Summary.ConsolidatedCount + "\r\n" + T("Dependency families refreshed: ", "의존 패밀리 최신화: ") + executionReport.Summary.DependencyLoadedCount + "\r\n" + T("Retyped elements: ", "타입 재지정 요소: ") + executionReport.Summary.RetypedElementCount + "\r\n" + T("Deleted obsolete types: ", "불필요 타입 삭제: ") + executionReport.Summary.DeletedObsoleteTypeCount + "\r\n" + T("Tracking refreshed: ", "추적 갱신: ") + executionReport.Summary.TrackingRefreshedCount + "\r\n" + T("Blocked: ", "차단: ") + executionReport.Summary.BlockedCount + "\r\n" + T("Failed: ", "실패: ") + executionReport.Summary.FailedCount + "\r\n" + T("Skipped: ", "건너뜀: ") + executionReport.Summary.SkippedCount + "\r\n" + T("Review before apply: ", "적용 전 검토: ") + preflightPath + "\r\n" + T("Review after apply: ", "적용 후 검토: ") + executionReport.PostPreflightPath + "\r\n" + T("Tracking: ", "추적: ") + (string.IsNullOrWhiteSpace(executionReport.TrackingPath) ? T("(not written)", "(저장되지 않음)") : executionReport.TrackingPath) + "\r\n" + T("Execution report: ", "실행 보고서: ") + executionPath);
							Execute = (Result)0;
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			message = FamilyBrowserCommandError.ToExternalCommandMessage("Family Browser", ex2);
			Execute = (Result)(-1);
			ProjectData.ClearProjectError();
		}
		finally
		{
			if (shouldCloseStandardDoc && standardDoc != null)
			{
				try
				{
					standardDoc.Close(false);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
		return Execute;
	}

	private static bool ConfirmExecution(SystemTypePreflightReport preflightReport, string standardDisplayName, string projectTitle, int executableCount)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Expected O, but got Unknown
		//IL_017c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0182: Invalid comparison between Unknown and I4
		TaskDialog dialog = new TaskDialog(T("Apply System Types", "시스템 타입 적용"));
		dialog.MainInstruction = T("Apply approval-ready supported system type updates?", "승인 준비가 된 지원 시스템 타입을 적용할까요?");
		dialog.MainContent = T("Project: ", "프로젝트: ") + projectTitle + "\r\n" + T("Standard: ", "표준: ") + standardDisplayName + "\r\n" + T("Executable items now: ", "현재 실행 가능 항목: ") + executableCount + "\r\n" + T("Ready: ", "준비: ") + preflightReport.Summary.ReadyCount + "\r\n" + T("Approval required: ", "승인 필요: ") + preflightReport.Summary.ApprovalRequiredCount + "\r\n" + T("Blocked: ", "차단: ") + preflightReport.Summary.BlockedCount + "\r\n\r\n" + T("Supported policy: ", "지원 정책: ") + SystemTypeSupportPolicyService.SupportedApplySummary() + " " + T("Review-only types remain review-only.", "검토 전용 타입은 검토 대상으로만 유지됩니다.");
		dialog.CommonButtons = (TaskDialogCommonButtons)6;
		dialog.DefaultButton = (TaskDialogResult)6;
		return (int)dialog.Show() == 6;
	}

	private static string T(string englishText, string koreanText)
	{
		return FamilyBrowserLanguageService.Text(englishText, koreanText);
	}

	private static int CountExecutableItems(SystemTypePreflightReport preflightReport)
	{
		return (preflightReport?.ExecutionPlan?.Items ?? new List<SystemSyncExecutionItem>()).Count([SpecialName] (SystemSyncExecutionItem x) =>
		{
			SystemTypeIdentityService.Normalize(x.SystemFamilyKind);
			string a = SystemTypeIdentityService.Normalize(x.ExecutionStatus);
			return SystemTypeSupportPolicyService.CanApply(x.SystemFamilyKind) && !string.Equals(a, "blocked", StringComparison.Ordinal);
		});
	}

	private static bool IsSuccessfulApplyOutcome(string outcome)
	{
		switch (SystemTypeIdentityService.Normalize(outcome))
		{
		case "created":
		case "overwritten":
		case "consolidated":
			return true;
		default:
			return false;
		}
	}

	private static void EnsureTargetDiffersFromStandard(Document targetDocument, StandardLibraryRegistrationRecord registration)
	{
		if (targetDocument != null && !string.IsNullOrWhiteSpace(targetDocument.PathName))
		{
			string fullPath = Path.GetFullPath(targetDocument.PathName);
			string standardPath = Path.GetFullPath(registration.ResolvedPath);
			if (string.Equals(fullPath, standardPath, StringComparison.OrdinalIgnoreCase))
			{
				throw new InvalidOperationException(T("The active project is the registered standard RVT itself. Open a different target project before applying system types.", "현재 열린 프로젝트가 등록된 표준 RVT입니다. 시스템 타입을 적용할 대상 프로젝트를 따로 열어주세요."));
			}
		}
	}

	private static void UpgradeStandardSnapshotSemantics(StandardLibrarySnapshot standardSnapshot, SystemTypeCatalogSnapshot standardCatalog)
	{
		if (standardSnapshot == null || standardCatalog == null)
		{
			return;
		}
		Dictionary<string, string> semanticMap = standardCatalog.Types.ToDictionary<SystemTypeSemanticSnapshot, string, string>([SpecialName] (SystemTypeSemanticSnapshot x) => SystemTypeIdentityService.BuildKey(x.SystemFamilyKind, x.CategoryName, x.TypeName), [SpecialName] (SystemTypeSemanticSnapshot x) => SystemTypeFingerprintService.Compute(x), StringComparer.Ordinal);
		foreach (StandardSystemTypeSnapshotItem item in standardSnapshot.SystemTypes)
		{
			string key = SystemTypeIdentityService.BuildKey(item.TypeClassName, item.CategoryName, item.TypeName);
			string fingerprint = null;
			if (semanticMap.TryGetValue(key, out fingerprint))
			{
				item.SemanticFingerprint = fingerprint;
			}
		}
	}
}
