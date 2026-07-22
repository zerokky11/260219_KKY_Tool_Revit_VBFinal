using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2027;

[Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class CmdOpenFamilyBrowser : IExternalCommand
{
	public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
	{
		Result Execute;
		try
		{
			FamilyBrowserDashboardModelessRuntime.Show(commandData.Application);
			Execute = Result.Succeeded;
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
}
