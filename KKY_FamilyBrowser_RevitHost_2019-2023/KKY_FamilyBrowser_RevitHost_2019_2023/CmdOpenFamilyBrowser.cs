using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2019_2023;

[Transaction(/*Could not decode attribute arguments.*/)]
[Regeneration(/*Could not decode attribute arguments.*/)]
public class CmdOpenFamilyBrowser : IExternalCommand
{
	public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
	{
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_002c: Unknown result type (might be due to invalid IL or missing references)
		Result Execute;
		try
		{
			FamilyBrowserDashboardModelessRuntime.Show(commandData.Application);
			Execute = (Result)0;
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
}
