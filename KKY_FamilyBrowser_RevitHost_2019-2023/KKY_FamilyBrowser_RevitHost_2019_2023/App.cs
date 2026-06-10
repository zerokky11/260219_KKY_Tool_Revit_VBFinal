using System;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Microsoft.VisualBasic.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2019_2023;

public class App : IExternalApplication
{
	private const string TabName = "KKY Browser";

	private const string PanelName = "Family Browser";

	public Result OnStartup(UIControlledApplication application)
	{
		//IL_0098: Unknown result type (might be due to invalid IL or missing references)
		//IL_008f: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a0: Unknown result type (might be due to invalid IL or missing references)
		Result OnStartup;
		try
		{
			try
			{
				application.CreateRibbonTab("KKY Browser");
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			RibbonPanel panel = FindOrCreatePanel(application);
			string assemblyPath = Assembly.GetExecutingAssembly().Location;
			AddPushButton(panel, "KKYFamilyBrowserOpenDashboard", "Family\nBrowser", assemblyPath, "KKY_FamilyBrowser_RevitHost_2019_2023.CmdOpenFamilyBrowser");
			application.ControlledApplication.DocumentOpened += HandleDocumentOpened;
			application.ControlledApplication.DocumentChanged += HandleDocumentChanged;
			application.ViewActivated += HandleViewActivated;
			FamilyBrowserNativeCommandGuardService.Start(application);
			FamilyBrowserRevitBridgeRuntime.Start();
			OnStartup = (Result)0;
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			OnStartup = (Result)(-1);
			ProjectData.ClearProjectError();
		}
		return OnStartup;
	}

	public Result OnShutdown(UIControlledApplication application)
	{
		try
		{
			application.ControlledApplication.DocumentOpened -= HandleDocumentOpened;
			application.ControlledApplication.DocumentChanged -= HandleDocumentChanged;
			application.ViewActivated -= HandleViewActivated;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		FamilyBrowserNativeCommandGuardService.Stop();
		FamilyBrowserDashboardModelessRuntime.Stop();
		FamilyBrowserRevitBridgeRuntime.Stop();
		return (Result)0;
	}

	private void HandleDocumentOpened(object sender, DocumentOpenedEventArgs e)
	{
		try
		{
			FamilyBrowserNativeCommandGuardService.NotifyActiveDocumentChanged(((RevitAPIPostDocEventArgs)e).Document);
			FamilyBrowserDashboardModelessRuntime.NotifyActiveDocumentChanged(((RevitAPIPostDocEventArgs)e).Document);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private void HandleDocumentChanged(object sender, DocumentChangedEventArgs e)
	{
		try
		{
			FamilyBrowserNativeCommandGuardService.HandleDocumentChanged(e);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			StandardRvtChangeCandidateService.HandleDocumentChanged(e);
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		try
		{
			Document changedDocument = null;
			if (e != null)
			{
				changedDocument = e.GetDocument();
			}
			FamilyBrowserDashboardModelessRuntime.NotifyDocumentContentChanged(changedDocument);
		}
		catch (Exception projectError3)
		{
			ProjectData.SetProjectError(projectError3);
			ProjectData.ClearProjectError();
		}
	}

	private void HandleViewActivated(object sender, ViewActivatedEventArgs e)
	{
		try
		{
			if (e != null && e.CurrentActiveView != null)
			{
				FamilyBrowserNativeCommandGuardService.NotifyActiveDocumentChanged(((Element)e.CurrentActiveView).Document);
				FamilyBrowserDashboardModelessRuntime.NotifyActiveDocumentChanged(((Element)e.CurrentActiveView).Document);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private RibbonPanel FindOrCreatePanel(UIControlledApplication application)
	{
		RibbonPanel existing = application.GetRibbonPanels("KKY Browser").FirstOrDefault([SpecialName] (RibbonPanel x) => string.Equals(x.Name, "Family Browser", StringComparison.Ordinal));
		if (existing != null)
		{
			return existing;
		}
		return application.CreateRibbonPanel("KKY Browser", "Family Browser");
	}

	private void AddPushButton(RibbonPanel panel, string buttonName, string buttonText, string assemblyPath, string className)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0017: Expected O, but got Unknown
		PushButtonData val = new PushButtonData(buttonName, buttonText, assemblyPath, className);
		if (panel.GetItems().OfType<RibbonItem>().FirstOrDefault([SpecialName] (RibbonItem x) => string.Equals(x.Name, ((RibbonItemData)val).Name, StringComparison.Ordinal)) == null)
		{
			panel.AddItem((RibbonItemData)(object)val);
		}
	}
}
