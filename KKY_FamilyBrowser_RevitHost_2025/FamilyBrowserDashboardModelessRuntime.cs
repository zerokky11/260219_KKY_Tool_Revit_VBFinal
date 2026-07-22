using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserDashboardModelessRuntime
{
	private sealed class WindowHandleOwner : IWin32Window
	{
		private readonly nint _handle;

		public nint Handle => _handle;

		public WindowHandleOwner(nint handle)
		{
			_handle = handle;
		}
	}

	private static readonly object SyncRoot = RuntimeHelpers.GetObjectValue(new object());

	private static FamilyBrowserDashboardHtmlForm _form;

	private static FamilyBrowserDashboardExternalEventHandler _handler;

	private static ExternalEvent _externalEvent;

	private static string _lastAutoDocumentKey = string.Empty;

	private FamilyBrowserDashboardModelessRuntime()
	{
	}

	public static void Show(UIApplication application)
	{
		if (application == null)
		{
			throw new ArgumentNullException("application");
		}
		FamilyBrowserRevitVersionContext.SetCurrentVersion(application);
		FamilyBrowserDashboardHtmlForm existing = null;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (_form != null && !_form.IsDisposed)
			{
				existing = _form;
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		if (existing != null)
		{
			existing.BringModelessWindowToFront();
			return;
		}
		FamilyBrowserDashboardHtmlForm formToShow = null;
		object syncRoot2 = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot2);
		bool lockTaken2 = false;
		try
		{
			Monitor.Enter(syncRoot2, ref lockTaken2);
			if (_form != null && !_form.IsDisposed)
			{
				formToShow = _form;
			}
			else
			{
				_handler = new FamilyBrowserDashboardExternalEventHandler();
				_externalEvent = ExternalEvent.Create(_handler);
				formToShow = (_form = new FamilyBrowserDashboardHtmlForm(application, [SpecialName] (string action) => _handler.Request(action, _externalEvent)));
				_handler.SetForm(formToShow);
				formToShow.FormClosed += HandleFormClosed;
			}
		}
		finally
		{
			if (lockTaken2)
			{
				Monitor.Exit(syncRoot2);
			}
		}
		if (formToShow != null)
		{
			IWin32Window owner = ResolveRevitWindowOwner(application);
			if (owner == null)
			{
				formToShow.Show();
			}
			else
			{
				formToShow.Show(owner);
			}
			formToShow.Activate();
		}
	}

	public static void Stop()
	{
		FamilyBrowserDashboardHtmlForm formToClose = null;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			formToClose = _form;
			if (_handler != null)
			{
				_handler.SetForm(null);
			}
			_form = null;
			_externalEvent = null;
			_handler = null;
			_lastAutoDocumentKey = string.Empty;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		if (formToClose != null && !formToClose.IsDisposed)
		{
			try
			{
				formToClose.Close();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
	}

	public static void NotifyActiveDocumentChanged(Document document)
	{
		if (document == null)
		{
			return;
		}
		FamilyBrowserRevitVersionContext.SetCurrentVersion(document);
		string key = BuildDocumentKey(document);
		FamilyBrowserDashboardHtmlForm formToRefresh = null;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (_form == null || _form.IsDisposed)
			{
				_lastAutoDocumentKey = key;
				return;
			}
			if (string.Equals(_lastAutoDocumentKey, key, StringComparison.OrdinalIgnoreCase))
			{
				return;
			}
			_lastAutoDocumentKey = key;
			formToRefresh = _form;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		if (formToRefresh != null && !formToRefresh.IsDisposed)
		{
			formToRefresh.RefreshForActiveDocumentChanged(key);
		}
	}

	public static void NotifyDocumentContentChanged(Document document)
	{
		if (document == null)
		{
			return;
		}
		string key = BuildDocumentKey(document);
		FamilyBrowserDashboardHtmlForm formToNotify = null;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (_form == null || _form.IsDisposed)
			{
				return;
			}
			formToNotify = _form;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		formToNotify?.MarkCurrentProjectContentChanged(key);
	}

	public static void NotifyDocumentCommitFinalized(Document document, string commitKind)
	{
		if (document == null)
		{
			return;
		}
		string key = BuildDocumentKey(document);
		FamilyBrowserDashboardHtmlForm formToNotify = null;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (_form == null || _form.IsDisposed)
			{
				return;
			}
			formToNotify = _form;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		if (formToNotify != null && formToNotify.IsHandleCreated)
		{
			formToNotify.BeginInvoke((Action)([SpecialName] () => formToNotify.RefreshAfterDocumentCommit(key, commitKind)));
		}
	}

	private static void HandleFormClosed(object sender, FormClosedEventArgs e)
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (object.ReferenceEquals(RuntimeHelpers.GetObjectValue(sender), _form))
			{
				if (_handler != null)
				{
					_handler.SetForm(null);
				}
				_form = null;
				_externalEvent = null;
				_handler = null;
				_lastAutoDocumentKey = string.Empty;
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

	private static string BuildDocumentKey(Document document)
	{
		if (document == null)
		{
			return string.Empty;
		}
		string path = string.Empty;
		try
		{
			path = document.PathName;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		if (!string.IsNullOrWhiteSpace(path))
		{
			return path;
		}
		string title = string.Empty;
		try
		{
			title = document.Title;
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		return title + "|" + document.GetHashCode().ToString(CultureInfo.InvariantCulture);
	}

	private static IWin32Window ResolveRevitWindowOwner(UIApplication application)
	{
		if (application == null)
		{
			return null;
		}
		try
		{
			PropertyInfo propertyInfo = application.GetType().GetProperty("MainWindowHandle");
			if ((object)propertyInfo == null)
			{
				return null;
			}
			if (RuntimeHelpers.GetObjectValue(propertyInfo.GetValue(application, null)) is nint handle && handle != IntPtr.Zero)
			{
				return new WindowHandleOwner(handle);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return null;
	}
}
