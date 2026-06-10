using System;
using System.Globalization;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2027;

[Transaction(/*Could not decode attribute arguments.*/)]
[Regeneration(/*Could not decode attribute arguments.*/)]
public class CmdRunFamilyBrowserDiagnostics : IExternalCommand
{
	public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
	{
		//IL_0123: Unknown result type (might be due to invalid IL or missing references)
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		//IL_012b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0102: Unknown result type (might be due to invalid IL or missing references)
		//IL_0109: Unknown result type (might be due to invalid IL or missing references)
		Result Execute;
		try
		{
			UIDocument uiDoc = commandData.Application.ActiveUIDocument;
			if (uiDoc == null || uiDoc.Document == null)
			{
				message = T("Open a project document before running Family Browser diagnostics.", "Family Browser 진단을 실행하기 전에 프로젝트 RVT를 먼저 열어주세요.");
				Execute = (Result)1;
			}
			else
			{
				DocumentFamilyDiagnosticsReport report = DocumentFamilyDiagnosticsBuilder.Build(uiDoc.Document);
				string outputPath = DiagnosticsOutputStore.Save(report);
				TaskDialog.Show(T("Family Browser Diagnostics", "Family Browser 진단"), T("Diagnostics completed.", "진단이 완료되었습니다.") + "\r\n\r\n" + T("Document", "문서") + ": " + report.DocumentTitle + "\r\n" + T("Families", "패밀리") + ": " + report.Families.Count.ToString(CultureInfo.InvariantCulture) + "\r\n" + T("Output", "결과 파일") + ": " + outputPath);
				Execute = (Result)0;
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
		return Execute;
	}

	private static string T(string englishText, string koreanText)
	{
		return FamilyBrowserLanguageService.Text(englishText, koreanText);
	}
}
