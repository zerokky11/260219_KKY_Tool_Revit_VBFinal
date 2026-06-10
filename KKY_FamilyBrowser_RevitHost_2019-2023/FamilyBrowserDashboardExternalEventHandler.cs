using System;
using System.Runtime.CompilerServices;
using System.Threading;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserDashboardExternalEventHandler : IExternalEventHandler
{
	private readonly object _syncRoot;

	private FamilyBrowserDashboardHtmlForm _form;

	private string _pendingAction;

	public FamilyBrowserDashboardExternalEventHandler()
	{
		_syncRoot = RuntimeHelpers.GetObjectValue(new object());
	}

	public void SetForm(FamilyBrowserDashboardHtmlForm form)
	{
		object syncRoot = _syncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			_form = form;
			if (form == null)
			{
				_pendingAction = null;
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public bool Request(string action, ExternalEvent externalEvent)
	{
		//IL_0064: Unknown result type (might be due to invalid IL or missing references)
		//IL_0069: Unknown result type (might be due to invalid IL or missing references)
		if (externalEvent == null)
		{
			return false;
		}
		object syncRoot = _syncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (_form == null || _form.IsDisposed || _pendingAction != null)
			{
				return false;
			}
			_pendingAction = action ?? string.Empty;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		string resultName = ((Enum)externalEvent.Raise()/*cast due to .constrained prefix*/).ToString();
		if (string.Equals(resultName, "Accepted", StringComparison.OrdinalIgnoreCase) || string.Equals(resultName, "Pending", StringComparison.OrdinalIgnoreCase))
		{
			return true;
		}
		object syncRoot2 = _syncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot2);
		bool lockTaken2 = false;
		try
		{
			Monitor.Enter(syncRoot2, ref lockTaken2);
			_pendingAction = null;
		}
		finally
		{
			if (lockTaken2)
			{
				Monitor.Exit(syncRoot2);
			}
		}
		return false;
	}

	public void Execute(UIApplication app)
	{
		FamilyBrowserRevitVersionContext.SetCurrentVersion(app);
		string action = null;
		FamilyBrowserDashboardHtmlForm form = null;
		object syncRoot = _syncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			action = _pendingAction;
			_pendingAction = null;
			form = _form;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		if (form != null && !form.IsDisposed && !string.IsNullOrWhiteSpace(action))
		{
			try
			{
				form.ExecuteDashboardActionFromExternalEvent(action);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				form.ShowDashboardActionException("Family Browser", ex2);
				ProjectData.ClearProjectError();
			}
		}
	}

	public string GetName()
	{
		return "KKY Family Browser Dashboard";
	}
}
