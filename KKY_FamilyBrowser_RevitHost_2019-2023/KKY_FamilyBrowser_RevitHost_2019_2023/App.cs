using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Microsoft.VisualBasic.CompilerServices;

namespace KKY_FamilyBrowser_RevitHost_2019_2023;

public class App : IExternalApplication
{
	private const string TabName = "KKY Tools";

	private const string PanelName = "Family Browser";

	private const int ElementTrackingBaselineRetryLimit = 3;

	private static int NativeGuardPolicyPreloadStarted;

	private static int NativeGuardPolicyPreloadReady;

	private static int ElementTrackingBaselineRefreshPending;

	private static int ElementTrackingBaselineRetryCount;

	public Result OnStartup(UIControlledApplication application)
	{
		Result OnStartup;
		try
		{
			try
			{
				application.CreateRibbonTab(TabName);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			RibbonPanel panel = FindOrCreatePanel(application);
			string assemblyPath = Assembly.GetExecutingAssembly().Location;
			AddPushButton(panel, "KKYFamilyBrowserOpenDashboard", "Family\nBrowser", assemblyPath, "KKY_FamilyBrowser_RevitHost_2019_2023.CmdOpenFamilyBrowser");
			FamilyBrowserNativeCommandGuardService.Start(application);
			FamilyBrowserRevitBridgeRuntime.Start();
			FamilyBrowserAutomaticModelCheckService.ResetAll();
			FamilyBrowserElementChangeTrackingService.StartReloadLatestBridge(application.ControlledApplication, HostWorkspacePathResolver.ResolveRoot);
			application.Idling += HandleIdling;
			application.ViewActivated += HandleViewActivated;
			application.ControlledApplication.DocumentOpened += HandleDocumentOpened;
			application.ControlledApplication.DocumentChanged += HandleDocumentChanged;
			application.ControlledApplication.DocumentSaving += HandleDocumentSaving;
			application.ControlledApplication.DocumentSavingAs += HandleDocumentSavingAs;
			application.ControlledApplication.DocumentSaved += HandleDocumentSaved;
			application.ControlledApplication.DocumentSavedAs += HandleDocumentSavedAs;
			application.ControlledApplication.DocumentSynchronizingWithCentral += HandleDocumentSynchronizingWithCentral;
			application.ControlledApplication.DocumentSynchronizedWithCentral += HandleDocumentSynchronizedWithCentral;
			application.ControlledApplication.DocumentClosing += HandleDocumentClosing;
			application.ControlledApplication.DocumentClosed += HandleDocumentClosed;
			// Keep startup guards lightweight: bind native command blockers, but do
			// not scan or index project content while Revit opens a document.
			OnStartup = Result.Succeeded;
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			OnStartup = Result.Failed;
			ProjectData.ClearProjectError();
		}
		return OnStartup;
	}

	Result IExternalApplication.OnStartup(UIControlledApplication application)
	{
		//ILSpy generated this explicit interface implementation from .override directive in OnStartup
		return this.OnStartup(application);
	}

	public Result OnShutdown(UIControlledApplication application)
	{
		try
		{
			application.Idling -= HandleIdling;
			application.ViewActivated -= HandleViewActivated;
			application.ControlledApplication.DocumentOpened -= HandleDocumentOpened;
			application.ControlledApplication.DocumentChanged -= HandleDocumentChanged;
			application.ControlledApplication.DocumentSaving -= HandleDocumentSaving;
			application.ControlledApplication.DocumentSavingAs -= HandleDocumentSavingAs;
			application.ControlledApplication.DocumentSaved -= HandleDocumentSaved;
			application.ControlledApplication.DocumentSavedAs -= HandleDocumentSavedAs;
			application.ControlledApplication.DocumentSynchronizingWithCentral -= HandleDocumentSynchronizingWithCentral;
			application.ControlledApplication.DocumentSynchronizedWithCentral -= HandleDocumentSynchronizedWithCentral;
			application.ControlledApplication.DocumentClosing -= HandleDocumentClosing;
			application.ControlledApplication.DocumentClosed -= HandleDocumentClosed;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		StandardRvtChangeCandidateService.ClearPendingOperationState();
		FamilyBrowserAutomaticModelCheckService.ResetAll();
		FamilyBrowserElementChangeTrackingService.Stop();
		FamilyBrowserNativeCommandGuardService.Stop();
		FamilyBrowserDashboardModelessRuntime.Stop();
		FamilyBrowserRevitBridgeRuntime.Stop();
		return Result.Succeeded;
	}

	Result IExternalApplication.OnShutdown(UIControlledApplication application)
	{
		//ILSpy generated this explicit interface implementation from .override directive in OnShutdown
		return this.OnShutdown(application);
	}

	private void HandleViewActivated(object sender, ViewActivatedEventArgs e)
	{
		Document document;
		try
		{
			document = e?.CurrentActiveView?.Document;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "View activation document resolution failed", projectError);
			ProjectData.ClearProjectError();
			return;
		}
		RunEventHandlerSafely("Element tracking view-activation baseline failed", () => FamilyBrowserElementChangeTrackingService.BeginDocumentSession(HostWorkspacePathResolver.ResolveRoot(), document));
		RunEventHandlerSafely(() => FamilyBrowserNativeCommandGuardService.NotifyActiveDocumentChanged(document));
		RunEventHandlerSafely("Native guard view-activation policy preload failed", () => QueueNativeGuardPolicyPreload(document));
		RunEventHandlerSafely(() => FamilyBrowserAutomaticModelCheckService.Schedule(document, "ViewActivated"));
		RunEventHandlerSafely(() => FamilyBrowserDashboardModelessRuntime.NotifyActiveDocumentChanged(document));
	}

	private static void QueueNativeGuardPolicyPreload(Document document)
	{
		if (document == null ||
			System.Threading.Interlocked.CompareExchange(ref NativeGuardPolicyPreloadReady, 0, 0) == 1 ||
			System.Threading.Interlocked.CompareExchange(ref NativeGuardPolicyPreloadStarted, 1, 0) != 0)
		{
			return;
		}
		FamilyBrowserDeploymentProjectIdentity projectIdentity = BuildNativeGuardProjectIdentity(document);
		string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		Task.Run(delegate
		{
			bool preloadReady = false;
			try
			{
				string persistedOverrideIssue;
				FamilyBrowserManagedFolderSetupService.TryApplyPersistedOverride(currentUser, out persistedOverrideIssue);
				string lastKnownPathIssue;
				FamilyBrowserMachineConfigStore.TryRestoreLastKnownManagedPolicyPath(currentUser, out lastKnownPathIssue);
				string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
				if (FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
				{
					preloadReady = ApplyStartupPolicy(workspaceRoot, currentUser);
				}
				FamilyBrowserDeploymentBootstrapService.TryApplyManagedPathOnly(HostWorkspacePathResolver.ResolveRoot(), currentUser, projectIdentity);
				workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
				if (!FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
				{
					return;
				}
				preloadReady = ApplyStartupPolicy(workspaceRoot, currentUser) || preloadReady;
			}
			catch (Exception ex)
			{
				FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Native guard startup policy preload failed", ex);
			}
			finally
			{
				System.Threading.Interlocked.Exchange(ref NativeGuardPolicyPreloadReady, preloadReady ? 1 : 0);
				System.Threading.Interlocked.Exchange(ref NativeGuardPolicyPreloadStarted, 0);
			}
		});
	}

	private static bool ApplyStartupPolicy(string workspaceRoot, string currentUser)
	{
		if (!FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
		{
			return false;
		}
		FamilyBrowserStandardPolicy policy = FamilyBrowserStandardPolicyStore.LoadOrCreate(workspaceRoot, currentUser);
		bool isAdminProfile = string.Equals(FamilyBrowserSecurityPolicyService.ResolveRole(policy, currentUser, null), "Admin", StringComparison.OrdinalIgnoreCase);
		bool enabled = isAdminProfile && FamilyBrowserUserSettingsStore.LoadAdminModeEnabled();
		FamilyBrowserNativeCommandGuardService.NotifyAdminModeChanged(enabled, policy, refreshUiNow: false);
		bool trackingEnabled = FamilyBrowserStandardPolicyStore.IsProjectElementChangeTrackingEnabled(policy);
		FamilyBrowserElementChangeTrackingService.NotifyPolicyChanged(workspaceRoot, trackingEnabled);
		if (trackingEnabled)
		{
			RequestElementTrackingBaselineRefresh();
		}
		return true;
	}

	private static void EnsureManagedPolicyBeforeDocumentEditing(Document document)
	{
		if (document == null)
		{
			return;
		}
		string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		string persistedOverrideIssue;
		FamilyBrowserManagedFolderSetupService.TryApplyPersistedOverride(currentUser, out persistedOverrideIssue);
		string lastKnownPathIssue;
		FamilyBrowserMachineConfigStore.TryRestoreLastKnownManagedPolicyPath(currentUser, out lastKnownPathIssue);
		string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
		if (!FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
		{
			FamilyBrowserDeploymentBootstrapService.TryApplyManagedPathOnly(workspaceRoot, currentUser, BuildNativeGuardProjectIdentity(document));
			workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
		}
		ApplyStartupPolicy(workspaceRoot, currentUser);
	}

	private static void RequestElementTrackingBaselineRefresh()
	{
		System.Threading.Interlocked.Exchange(ref ElementTrackingBaselineRetryCount, 0);
		System.Threading.Interlocked.Exchange(ref ElementTrackingBaselineRefreshPending, 1);
	}

	private static void RequestElementTrackingBaselineRetry()
	{
		if (System.Threading.Interlocked.Increment(ref ElementTrackingBaselineRetryCount) <= ElementTrackingBaselineRetryLimit)
		{
			System.Threading.Interlocked.Exchange(ref ElementTrackingBaselineRefreshPending, 1);
		}
	}

	private static void HandleIdling(object sender, IdlingEventArgs e)
	{
		UIApplication uiApplication = sender as UIApplication;
		try
		{
			FamilyBrowserAutomaticModelCheckService.ProcessPending(uiApplication);
		}
		catch (Exception ex)
		{
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Automatic Current Model Check idling failed", ex);
		}
		bool appRequest = System.Threading.Interlocked.Exchange(ref ElementTrackingBaselineRefreshPending, 0) == 1;
		bool serviceRequest = FamilyBrowserElementChangeTrackingService.ConsumeDocumentSessionBaselineRefreshRequest();
		if (serviceRequest)
		{
			System.Threading.Interlocked.Exchange(ref ElementTrackingBaselineRetryCount, 0);
		}
		if (!appRequest && !serviceRequest)
		{
			return;
		}
		try
		{
			Document document = uiApplication?.ActiveUIDocument?.Document;
			if (document == null)
			{
				RequestElementTrackingBaselineRetry();
				return;
			}
			string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
			if (FamilyBrowserElementChangeTrackingService.IsEnabled(workspaceRoot))
			{
				FamilyBrowserElementChangeTrackingService.BeginDocumentSession(workspaceRoot, document);
				if (FamilyBrowserElementChangeTrackingService.HasDocumentSession(document))
				{
					System.Threading.Interlocked.Exchange(ref ElementTrackingBaselineRetryCount, 0);
				}
				else
				{
					RequestElementTrackingBaselineRetry();
				}
			}
			else
			{
				System.Threading.Interlocked.Exchange(ref ElementTrackingBaselineRetryCount, 0);
			}
		}
		catch (Exception ex)
		{
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Element tracking deferred baseline failed", ex);
			RequestElementTrackingBaselineRetry();
		}
	}

	private static FamilyBrowserDeploymentProjectIdentity BuildNativeGuardProjectIdentity(Document document)
	{
		FamilyBrowserDeploymentProjectIdentity identity = new FamilyBrowserDeploymentProjectIdentity();
		if (document == null)
		{
			return identity;
		}
		try
		{
			identity.ProjectTitle = document.Title ?? string.Empty;
			identity.ModelPath = document.PathName ?? string.Empty;
			if (document.IsWorkshared)
			{
				ModelPath centralPath = document.GetWorksharingCentralModelPath();
				if (centralPath != null)
				{
					identity.CentralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(centralPath) ?? string.Empty;
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return identity;
	}

	private void HandleDocumentOpened(object sender, DocumentOpenedEventArgs e)
	{
		Document document;
		try
		{
			document = e?.Document;
		}
		catch (Exception ex)
		{
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Document-open document resolution failed", ex);
			return;
		}
		RunEventHandlerSafely("Managed policy document-open preparation failed", () => EnsureManagedPolicyBeforeDocumentEditing(document));
		RunEventHandlerSafely("Element tracking document-open baseline failed", () => FamilyBrowserElementChangeTrackingService.BeginDocumentSession(HostWorkspacePathResolver.ResolveRoot(), document));
		RunEventHandlerSafely("Native guard document-open policy preload failed", () => QueueNativeGuardPolicyPreload(document));
		RunEventHandlerSafely("Automatic Current Model Check document-open scheduling failed", () => FamilyBrowserAutomaticModelCheckService.Schedule(document, "DocumentOpened", force: true));
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
			FamilyBrowserElementChangeTrackingService.HandleDocumentChanged(HostWorkspacePathResolver.ResolveRoot(), e);
		}
		catch (Exception projectError3)
		{
			ProjectData.SetProjectError(projectError3);
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Element tracking DocumentChanged failed", projectError3);
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
		catch (Exception projectError4)
		{
			ProjectData.SetProjectError(projectError4);
			ProjectData.ClearProjectError();
		}
	}

	private void HandleDocumentSaving(object sender, DocumentSavingEventArgs e)
	{
		RunEventHandlerSafely("Element tracking Save preparation failed", () => FamilyBrowserElementChangeTrackingService.PrepareDocumentCommit(HostWorkspacePathResolver.ResolveRoot(), e?.Document, "Save"));
		try
		{
			StandardRvtChangeCandidateService.HandleDocumentSaving(e?.Document);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private void HandleDocumentSavingAs(object sender, DocumentSavingAsEventArgs e)
	{
		RunEventHandlerSafely("Element tracking Save As preparation failed", () => FamilyBrowserElementChangeTrackingService.PrepareDocumentCommit(HostWorkspacePathResolver.ResolveRoot(), e?.Document, "SaveAs"));
		try
		{
			StandardRvtChangeCandidateService.HandleDocumentSaving(e?.Document);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private void HandleDocumentSaved(object sender, DocumentSavedEventArgs e)
	{
		Document document = null;
		object status;
		try
		{
			document = e?.Document;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			FamilyBrowserElementChangeTrackingService.HandleDocumentSaveCompletionFailure(HostWorkspacePathResolver.ResolveRoot(), null, "Save", projectError);
			ProjectData.ClearProjectError();
			return;
		}
		if (document == null)
		{
			FamilyBrowserElementChangeTrackingService.HandleDocumentSaveCompletionFailure(HostWorkspacePathResolver.ResolveRoot(), null, "Save", new InvalidOperationException("Save completed without a document."));
			return;
		}
		try
		{
			status = (e == null) ? null : (object)e.Status;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			FamilyBrowserElementChangeTrackingService.HandleDocumentSaveCompletionFailure(HostWorkspacePathResolver.ResolveRoot(), document, "Save", projectError);
			ProjectData.ClearProjectError();
			return;
		}
		if (status == null)
		{
			FamilyBrowserElementChangeTrackingService.HandleDocumentSaveCompletionFailure(HostWorkspacePathResolver.ResolveRoot(), document, "Save", new InvalidOperationException("Save completed without a status."));
			return;
		}
		RunEventHandlerSafely(() => StandardRvtChangeCandidateService.HandleDocumentSaved(document, status, "Save"));
		RunEventHandlerSafely("Element tracking Save commit failed", () => FamilyBrowserElementChangeTrackingService.HandleDocumentCommitted(HostWorkspacePathResolver.ResolveRoot(), document, status, "Save"));
		if (FamilyBrowserProjectCatalogService.IsSuccessfulRevitEventStatus(status))
		{
			RunEventHandlerSafely(() => ObserveProjectCatalogAfterCommit(document, "Save"));
			RunEventHandlerSafely(() => FamilyBrowserDashboardModelessRuntime.NotifyDocumentCommitFinalized(document, "Save"));
		}
	}

	private void HandleDocumentSavedAs(object sender, DocumentSavedAsEventArgs e)
	{
		Document document = null;
		object status;
		try
		{
			document = e?.Document;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			FamilyBrowserElementChangeTrackingService.HandleDocumentSaveCompletionFailure(HostWorkspacePathResolver.ResolveRoot(), null, "SaveAs", projectError);
			ProjectData.ClearProjectError();
			return;
		}
		if (document == null)
		{
			FamilyBrowserElementChangeTrackingService.HandleDocumentSaveCompletionFailure(HostWorkspacePathResolver.ResolveRoot(), null, "SaveAs", new InvalidOperationException("Save As completed without a document."));
			return;
		}
		try
		{
			status = (e == null) ? null : (object)e.Status;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			FamilyBrowserElementChangeTrackingService.HandleDocumentSaveCompletionFailure(HostWorkspacePathResolver.ResolveRoot(), document, "SaveAs", projectError);
			ProjectData.ClearProjectError();
			return;
		}
		if (status == null)
		{
			FamilyBrowserElementChangeTrackingService.HandleDocumentSaveCompletionFailure(HostWorkspacePathResolver.ResolveRoot(), document, "SaveAs", new InvalidOperationException("Save As completed without a status."));
			return;
		}
		RunEventHandlerSafely(() => StandardRvtChangeCandidateService.HandleDocumentSaved(document, status, "SaveAs"));
		RunEventHandlerSafely("Element tracking Save As commit failed", () => FamilyBrowserElementChangeTrackingService.HandleDocumentCommitted(HostWorkspacePathResolver.ResolveRoot(), document, status, "SaveAs"));
		if (FamilyBrowserProjectCatalogService.IsSuccessfulRevitEventStatus(status))
		{
			RunEventHandlerSafely(() => ObserveProjectCatalogAfterCommit(document, "SaveAs"));
			RunEventHandlerSafely(() => FamilyBrowserDashboardModelessRuntime.NotifyDocumentCommitFinalized(document, "SaveAs"));
		}
	}

	private void HandleDocumentSynchronizingWithCentral(object sender, DocumentSynchronizingWithCentralEventArgs e)
	{
		string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
		try
		{
			FamilyBrowserElementChangeTrackingService.HandleDocumentSynchronizingWithCentral(workspaceRoot, e?.Document);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			FamilyBrowserElementChangeTrackingService.HandleDocumentSynchronizationStartFailure(workspaceRoot, projectError);
			ProjectData.ClearProjectError();
		}
	}

	private void HandleDocumentSynchronizedWithCentral(object sender, DocumentSynchronizedWithCentralEventArgs e)
	{
		Document document;
		object status;
		try
		{
			document = e?.Document;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			FamilyBrowserElementChangeTrackingService.HandleDocumentSynchronizationCompletionFailure(HostWorkspacePathResolver.ResolveRoot(), null, projectError);
			ProjectData.ClearProjectError();
			return;
		}
		if (document == null)
		{
			FamilyBrowserElementChangeTrackingService.HandleDocumentSynchronizationCompletionFailure(HostWorkspacePathResolver.ResolveRoot(), null, new InvalidOperationException("Synchronize with Central completed without a document."));
			return;
		}
		try
		{
			status = (e == null) ? null : (object)e.Status;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			FamilyBrowserElementChangeTrackingService.HandleDocumentSynchronizationCompletionFailure(HostWorkspacePathResolver.ResolveRoot(), document, projectError);
			ProjectData.ClearProjectError();
			return;
		}
		if (status == null)
		{
			FamilyBrowserElementChangeTrackingService.HandleDocumentSynchronizationCompletionFailure(HostWorkspacePathResolver.ResolveRoot(), document, new InvalidOperationException("Synchronize with Central completed without a status."));
			return;
		}
		RunEventHandlerSafely(() => StandardRvtChangeCandidateService.HandleDocumentSynchronizedWithCentral(document, status));
		RunEventHandlerSafely("Element tracking synchronization commit failed", () => FamilyBrowserElementChangeTrackingService.HandleDocumentCommitted(HostWorkspacePathResolver.ResolveRoot(), document, status, "SynchronizeWithCentral"));
		if (FamilyBrowserProjectCatalogService.IsSuccessfulRevitEventStatus(status))
		{
			RunEventHandlerSafely(() => ObserveProjectCatalogAfterCommit(document, "SynchronizeWithCentral"));
			RunEventHandlerSafely(() => FamilyBrowserDashboardModelessRuntime.NotifyDocumentCommitFinalized(document, "SynchronizeWithCentral"));
		}
	}

	private void HandleDocumentClosing(object sender, DocumentClosingEventArgs e)
	{
		Document document = null;
		int documentId = -1;
		try
		{
			document = e?.Document;
		}
		catch (Exception ex)
		{
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Document-closing document resolution failed", ex);
		}
		try
		{
			documentId = (e == null) ? (-1) : e.DocumentId;
		}
		catch (Exception ex)
		{
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Document-closing identifier resolution failed", ex);
		}
		RunEventHandlerSafely("Automatic Current Model Check document-closing cleanup failed", () => FamilyBrowserAutomaticModelCheckService.Remove(document));
		RunEventHandlerSafely("Element tracking document-closing failed", () => FamilyBrowserElementChangeTrackingService.HandleDocumentClosing(document, documentId));
		RunEventHandlerSafely(() => StandardRvtChangeCandidateService.HandleDocumentClosing(document, documentId));
	}

	private void HandleDocumentClosed(object sender, DocumentClosedEventArgs e)
	{
		int documentId = (e == null) ? (-1) : e.DocumentId;
		RunEventHandlerSafely("Element tracking document-closed cleanup failed", () => FamilyBrowserElementChangeTrackingService.HandleDocumentClosed(documentId));
		RunEventHandlerSafely(() => StandardRvtChangeCandidateService.HandleDocumentClosed(documentId));
	}

	private static void ObserveProjectCatalogAfterCommit(Document document, string commitKind)
	{
		bool performed = FamilyBrowserElementChangeTrackingService.ConsumeProjectCatalogObservationRequired(document);
		bool published = !performed;
		Stopwatch stopwatch = Stopwatch.StartNew();
		try
		{
			if (performed)
			{
				FamilyBrowserProjectCatalogState state = FamilyBrowserProjectCatalogService.Observe(HostWorkspacePathResolver.ResolveRoot(), document, commitKind);
				published = FamilyBrowserProjectCatalogService.IsPublishedObservationState(state);
			}
		}
		finally
		{
			if (performed && !published)
			{
				FamilyBrowserElementChangeTrackingService.RestoreProjectCatalogObservationRequired(document);
			}
			stopwatch.Stop();
			FamilyBrowserElementChangeTrackingService.RecordProjectCatalogObservationPerformance(document, commitKind, performed, stopwatch.ElapsedMilliseconds);
		}
	}

	private static void RunEventHandlerSafely(Action action)
	{
		try
		{
			action?.Invoke();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void RunEventHandlerSafely(string caption, Action action)
	{
		try
		{
			action?.Invoke();
		}
		catch (Exception ex)
		{
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), caption, ex);
		}
	}

	private RibbonPanel FindOrCreatePanel(UIControlledApplication application)
	{
		RibbonPanel existing = application.GetRibbonPanels(TabName).FirstOrDefault([SpecialName] (RibbonPanel x) => string.Equals(x.Name, PanelName, StringComparison.Ordinal));
		if (existing != null)
		{
			return existing;
		}
		return application.CreateRibbonPanel(TabName, PanelName);
	}

	private void AddPushButton(RibbonPanel panel, string buttonName, string buttonText, string assemblyPath, string className)
	{
		PushButtonData pushButtonData = new PushButtonData(buttonName, buttonText, assemblyPath, className);
		RibbonItem existingItem = panel.GetItems().OfType<RibbonItem>().FirstOrDefault([SpecialName] (RibbonItem x) => string.Equals(x.Name, pushButtonData.Name, StringComparison.Ordinal));
		PushButton button = existingItem as PushButton;
		if (existingItem == null)
		{
			button = panel.AddItem(pushButtonData) as PushButton;
		}
		if (button != null)
		{
			button.ToolTip = "KKY Family Browser 열기";
			var smallIcon = FamilyBrowserRibbonIcon.LoadSmall();
			var largeIcon = FamilyBrowserRibbonIcon.LoadLarge();
			if (smallIcon != null)
			{
				button.Image = smallIcon;
			}
			if (largeIcon != null)
			{
				button.LargeImage = largeIcon;
			}
		}
	}
}
