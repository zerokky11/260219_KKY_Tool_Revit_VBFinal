using System;
using System.Globalization;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2027;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CmdRunFamilyBrowserDiagnostics : IExternalCommand
{
	public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
	{
		Result Execute;
		try
		{
			UIDocument uiDoc = commandData.Application.ActiveUIDocument;
			if (uiDoc == null || uiDoc.Document == null)
			{
				message = T("Open a project document before running Family Browser diagnostics.", "Family Browser 진단을 실행하기 전에 프로젝트 RVT를 먼저 열어주세요.");
				Execute = Result.Cancelled;
			}
			else
			{
				DocumentFamilyDiagnosticsReport report = DocumentFamilyDiagnosticsBuilder.Build(uiDoc.Document);
				string outputPath = DiagnosticsOutputStore.Save(report);
				FamilyBrowserResultDialog.Show(T("Family Browser Diagnostics", "Family Browser 진단"), T("Diagnostics completed.", "진단이 완료되었습니다.") + "\r\n\r\n" + T("Document", "문서") + ": " + report.DocumentTitle + "\r\n" + T("Families", "패밀리") + ": " + report.Families.Count.ToString(CultureInfo.InvariantCulture) + "\r\n" + T("Output", "결과 파일") + ": " + outputPath);
				Execute = Result.Succeeded;
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
