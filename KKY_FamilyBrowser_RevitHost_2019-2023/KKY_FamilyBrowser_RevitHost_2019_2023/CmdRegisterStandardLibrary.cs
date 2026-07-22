using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2019_2023;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CmdRegisterStandardLibrary : IExternalCommand
{
	public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
	{
		Result Execute;
		try
		{
			string selectedPath = PromptForStandardLibraryPath(commandData);
			if (string.IsNullOrWhiteSpace(selectedPath))
			{
				Execute = Result.Cancelled;
			}
			else
			{
				StandardLibraryRegistrationResult registrationResult = StandardLibraryRegistrationService.Register(HostWorkspacePathResolver.ResolveRoot(), commandData.Application.Application, selectedPath, Environment.UserName, "Fast", commandData.Application);
				FamilyBrowserResultDialog.Show(T("Standard RVT Connected", K("7ZGc7KSAIFJWVCDsl7DqsrAg7JmE66OM")), T("The approved standard RVT has been connected and snapshotted.", K("7Iq57J2465CcIO2RnOykgCBSVlTrpbwg7Jew6rKw7ZWY6rOgIOyKpOuDheyDt+ydhCDsg53shLHtlojsirXri4jri6Qu")) + "\r\n\r\n" + T("Name", K("7J2066aE")) + ": " + registrationResult.Registration.DisplayName + "\r\n" + T("Source", K("7IaM7Iqk")) + ": " + registrationResult.Registration.SourceKind + "\r\n" + T("Loadable families", K("66Gc642U67iUIO2MqOuwgOumrA==")) + ": " + registrationResult.Snapshot.Summary.LoadableFamilyCount + "\r\n" + T("Loadable types", K("66Gc642U67iUIO2DgOyehQ==")) + ": " + registrationResult.Snapshot.Summary.LoadableTypeCount + "\r\n" + T("System types", K("7Iuc7Iqk7YWcIO2DgOyehQ==")) + ": " + registrationResult.Snapshot.Summary.SystemTypeCount + "\r\n" + T("Registry", K("65Ox66GdIOygleuztA==")) + ": " + registrationResult.RegistrationPath + "\r\n" + T("Snapshot", K("7Iqk64OF7IO3")) + ": " + registrationResult.SnapshotPath);
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

	private static string PromptForStandardLibraryPath(ExternalCommandData commandData)
	{
		using OpenFileDialog dialog = new OpenFileDialog();
		dialog.Title = T("Select Standard Family Library RVT", K("7ZGc7KSAIO2MqOuwgOumrCDrnbzsnbTruIzrn6zrpqwgUlZUIOyEoO2DnQ=="));
		dialog.Filter = "Revit Project (*.rvt)|*.rvt";
		dialog.CheckFileExists = true;
		dialog.Multiselect = false;
		dialog.RestoreDirectory = true;
		Document activeDoc = commandData.Application.ActiveUIDocument?.Document;
		if (activeDoc != null && !string.IsNullOrWhiteSpace(activeDoc.PathName))
		{
			try
			{
				dialog.InitialDirectory = Path.GetDirectoryName(activeDoc.PathName);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		if (dialog.ShowDialog() != DialogResult.OK)
		{
			return string.Empty;
		}
		return dialog.FileName;
	}

	private static string T(string englishText, string koreanText)
	{
		return FamilyBrowserLanguageService.Text(englishText, koreanText);
	}

	private static string K(string base64Text)
	{
		return Encoding.UTF8.GetString(Convert.FromBase64String(base64Text));
	}
}
