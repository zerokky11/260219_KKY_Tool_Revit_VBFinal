using System;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2019_2023;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CmdStampCurrentProjectState : IExternalCommand
{
	public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
	{
		Result Execute;
		try
		{
			UIDocument uiDoc = commandData.Application.ActiveUIDocument;
			if (uiDoc == null || uiDoc.Document == null)
			{
				message = T("Open a project document before stamping the current project state.", "현재 프로젝트 상태를 스탬프하기 전에 프로젝트 RVT를 먼저 열어주세요.");
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
						ProjectTrackingCatalog catalog = ProjectTrackingStampService.BuildCatalog(registration, standardSnapshot, projectSnapshot, FamilyBrowserUserSettingsStore.ResolveDetailedSystemTypeComparisonEnabled(policy));
						using (Transaction tx = new Transaction(uiDoc.Document, "KKY Family Browser Stamp Current State"))
						{
							tx.Start();
							ProjectTrackingStoreService.Save(uiDoc.Document, catalog);
							tx.Commit();
						}
						string outputPath = ProjectTrackingOutputStore.Save(workspaceRoot, catalog);
						FamilyBrowserResultDialog.Show(T("Project Tracking Stamped", "프로젝트 추적 스탬프 완료"), T("The current project state was stamped against the active standard snapshot.", "현재 프로젝트 상태를 활성 표준 스냅샷 기준으로 스탬프했습니다.") + "\r\n\r\n" + T("Project", "프로젝트") + ": " + uiDoc.Document.Title + "\r\n" + T("Standard", "표준") + ": " + registration.DisplayName + "\r\n" + T("Tracked loadable families", "추적된 로더블 패밀리") + ": " + catalog.LoadableFamilies.Count + "\r\n" + T("Tracked system types", "추적된 시스템 타입") + ": " + catalog.SystemTypes.Count + "\r\n" + T("Output", "결과 파일") + ": " + outputPath);
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
