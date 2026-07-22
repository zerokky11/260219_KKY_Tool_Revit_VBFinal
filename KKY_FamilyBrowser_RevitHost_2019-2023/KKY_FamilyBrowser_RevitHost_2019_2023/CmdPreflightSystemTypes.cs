using System;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2019_2023;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CmdPreflightSystemTypes : IExternalCommand
{
	public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
	{
		Document standardDoc = null;
		bool shouldCloseStandardDoc = false;
		Result Execute;
		try
		{
			UIDocument uiDoc = commandData.Application.ActiveUIDocument;
			if (uiDoc == null || uiDoc.Document == null)
			{
				message = T("Open a project document before reviewing system types.", "시스템 타입 검토를 실행하기 전에 프로젝트 RVT를 먼저 열어주세요.");
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
					standardDoc = StandardLibraryDocumentResolver.OpenRegisteredDocument(commandData.Application.Application, registration, ref shouldCloseStandardDoc);
					SystemTypePreflightReport report = SystemTypePreflightBuilderService.BuildReport(registration, standardDoc, uiDoc.Document);
					SystemTypePreflightStore.Save(workspaceRoot, report);
					int reviewCount = checked(report.Summary.ApprovalRequiredCount + report.Summary.DependencyManualReviewCount);
					FamilyBrowserResultDialog.Show(T("System Type Review", "시스템 타입 검토"), T("System type review completed.", "시스템 타입 검토가 완료되었습니다.") + "\r\n\r\n" + T("Project", "프로젝트") + ": " + uiDoc.Document.Title + "\r\n" + T("Standard", "표준") + ": " + registration.DisplayName + "\r\n" + T("No change", "변경 없음") + ": " + report.Summary.NoChangeCount + "\r\n" + T("Ready", "적용 가능") + ": " + report.Summary.ReadyCount + "\r\n" + T("Review", "검토") + ": " + reviewCount + "\r\n" + T("Blocked", "차단") + ": " + report.Summary.BlockedCount + "\r\n" + T("Missing dependency family", "의존 패밀리 없음") + ": " + report.Summary.MissingDependencyFamilyCount + "\r\n" + T("Dependency reload", "의존 패밀리 재로드") + ": " + report.Summary.DependencyReloadCount + "\r\n" + T("Review report was saved for admin diagnostics.", "검토 보고서는 관리자 진단용으로 저장되었습니다."));
					Execute = Result.Succeeded;
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
		finally
		{
			if (shouldCloseStandardDoc && standardDoc != null)
			{
				try
				{
					standardDoc.Close(saveModified: false);
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
