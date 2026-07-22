using System;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2019_2023;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CmdCompareProjectToStandard : IExternalCommand
{
	public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
	{
		Result Execute;
		try
		{
			UIDocument uiDoc = commandData.Application.ActiveUIDocument;
			if (uiDoc == null || uiDoc.Document == null)
			{
				message = T("Open a project document before checking the current model.", "현재 모델 검사를 실행하기 전에 프로젝트 RVT를 먼저 열어주세요.");
				Execute = Result.Cancelled;
			}
			else
			{
				string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
				FamilyBrowserStandardPolicy policy = FamilyBrowserStandardPolicyStore.LoadOrCreate(workspaceRoot, Environment.UserName);
				string registryPath = FamilyBrowserStandardPolicyStore.ResolveEffectiveRegistrationPath(workspaceRoot, policy);
				if (!File.Exists(registryPath))
				{
					message = T("No registered standard RVT was found. Connect the approved standard RVT first.", "등록된 표준 RVT를 찾을 수 없습니다. 승인된 표준 RVT를 먼저 연결하세요.");
					Execute = Result.Failed;
				}
				else
				{
					StandardLibraryRegistrationRecord registration = DataContractJsonFileStore.Load<StandardLibraryRegistrationRecord>(registryPath);
					if (string.IsNullOrWhiteSpace(registration.LastSnapshotPath) || !File.Exists(registration.LastSnapshotPath))
					{
						message = T("The standard snapshot could not be found. Refresh or reconnect the standard RVT.", "표준 스냅샷을 찾을 수 없습니다. 표준 RVT를 새로고침하거나 다시 연결하세요.");
						Execute = Result.Failed;
					}
					else
					{
						StandardLibrarySnapshot standardSnapshot = DataContractJsonFileStore.Load<StandardLibrarySnapshot>(registration.LastSnapshotPath);
						ProjectContentSnapshot projectSnapshot = ProjectSnapshotCaptureService.Capture(uiDoc.Document);
						string projectSnapshotPath = ProjectSnapshotStore.Save(workspaceRoot, projectSnapshot);
						ProjectTrackingCatalog trackingCatalog = ProjectTrackingStoreService.Load(uiDoc.Document);
						ProjectStandardComparisonReport report = ProjectStandardComparisonService.BuildReport(registration, registration.LastSnapshotPath, standardSnapshot, projectSnapshotPath, projectSnapshot, trackingCatalog, FamilyBrowserUserSettingsStore.ResolveDetailedSystemTypeComparisonEnabled(policy));
						ProjectStandardComparisonStore.Save(workspaceRoot, report);
						FamilyBrowserResultDialog.Show(T("Current Model Check", "현재 모델 검사"), T("Current model check completed.", "현재 모델 검사가 완료되었습니다.") + "\r\n\r\n" + T("Project", "프로젝트") + ": " + report.Project.DocumentTitle + "\r\n" + T("Standard", "표준") + ": " + report.Standard.DisplayName + "\r\n" + T("Tracking", "추적") + ": " + report.TrackingState + "\r\n" + T("Loadable latest", "로더블 기준 일치") + ": " + report.Summary.LoadableLatestCount + "\r\n" + T("Loadable available", "로더블 로드 가능") + ": " + report.Summary.LoadableLoadAvailableCount + "\r\n" + T("Loadable different", "로더블 표준과 다름") + ": " + report.Summary.LoadableDifferentCount + "\r\n" + T("Loadable project only", "로더블 프로젝트 전용") + ": " + report.Summary.LoadableProjectOnlyCount + "\r\n" + T("System latest", "시스템 기준 일치") + ": " + report.Summary.SystemLatestCount + "\r\n" + T("System available", "시스템 적용 가능") + ": " + report.Summary.SystemLoadAvailableCount + "\r\n" + T("System different", "시스템 표준과 다름") + ": " + report.Summary.SystemDifferentCount + "\r\n" + T("System project only", "시스템 프로젝트 전용") + ": " + report.Summary.SystemProjectOnlyCount + "\r\n" + T("Comparison report was saved for admin diagnostics.", "비교 보고서는 관리자 진단용으로 저장되었습니다."));
						Execute = Result.Succeeded;
					}
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			message = FamilyBrowserCommandError.ToExternalCommandMessage("Family Browser", ex2);
			Execute = Result.Failed;
			ProjectData.ClearProjectError();
		}
		return Execute;
	}

	Result IExternalCommand.Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Execute
		return this.Execute(commandData, ref message, elements);
	}

	private static string T(string englishText, string koreanText)
	{
		return FamilyBrowserLanguageService.Text(englishText, koreanText);
	}
}
