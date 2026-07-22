using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserNativeCommandGuardService
{
	[DllImport("mpr.dll", CharSet = CharSet.Unicode)]
	private static extern int WNetGetConnection(string localName, StringBuilder remoteName, ref int length);

	private sealed class TrustedOperationScope : IDisposable
	{
		private bool Disposed;

		public TrustedOperationScope()
		{
			Disposed = false;
		}

		public void Dispose()
		{
			if (!Disposed)
			{
				Disposed = true;
				EndTrustedOperation();
			}
		}

		void IDisposable.Dispose()
		{
			//ILSpy generated this explicit interface implementation from .override directive in Dispose
			this.Dispose();
		}
	}

	private sealed class ProtectedContentChangeUpdater : IUpdater
	{
		private readonly UpdaterId UpdaterIdValue;

		public ProtectedContentChangeUpdater(AddInId addInId)
		{
			UpdaterIdValue = new UpdaterId(addInId, ProtectedChangeUpdaterGuid);
		}

		public void Execute(UpdaterData data)
		{
			try
			{
				HandleProtectedContentUpdaterExecute(data);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Protected native change updater execution failed", ex2, string.Empty);
				ProjectData.ClearProjectError();
			}
		}

		void IUpdater.Execute(UpdaterData data)
		{
			//ILSpy generated this explicit interface implementation from .override directive in Execute
			this.Execute(data);
		}

		public UpdaterId GetUpdaterId()
		{
			return UpdaterIdValue;
		}

		UpdaterId IUpdater.GetUpdaterId()
		{
			//ILSpy generated this explicit interface implementation from .override directive in GetUpdaterId
			return this.GetUpdaterId();
		}

		public string GetUpdaterName()
		{
			return "KKY Family Browser protected content and nested-only placement guard";
		}

		string IUpdater.GetUpdaterName()
		{
			//ILSpy generated this explicit interface implementation from .override directive in GetUpdaterName
			return this.GetUpdaterName();
		}

		public string GetAdditionalInformation()
		{
			return "Blocks protected family/type changes and direct placement of nested-only families before transaction commit for non-admin users.";
		}

		string IUpdater.GetAdditionalInformation()
		{
			//ILSpy generated this explicit interface implementation from .override directive in GetAdditionalInformation
			return this.GetAdditionalInformation();
		}

		public ChangePriority GetChangePriority()
		{
			return ChangePriority.FreeStandingComponents;
		}

		ChangePriority IUpdater.GetChangePriority()
		{
			//ILSpy generated this explicit interface implementation from .override directive in GetChangePriority
			return this.GetChangePriority();
		}
	}

	private sealed class ProtectedNativeCommandDefinition
	{
		public string Key { get; set; }

		public string DisplayNameEn { get; set; }

		public string DisplayNameKo { get; set; }

		public string RequiredPermission { get; set; }

		public List<string> PostableCommandNames { get; set; }

		public List<string> CommandIdNames { get; set; }

		public bool FamilyDocumentOnly { get; set; }

		public bool BlockDuringBlockedActionCooldownOnly { get; set; }

		public ProtectedNativeCommandDefinition()
		{
			Key = string.Empty;
			DisplayNameEn = string.Empty;
			DisplayNameKo = string.Empty;
			RequiredPermission = string.Empty;
			PostableCommandNames = new List<string>();
			CommandIdNames = new List<string>();
		}
	}

	private sealed class BoundNativeCommand
	{
		public AddInCommandBinding Binding { get; set; }

		public ProtectedNativeCommandDefinition Definition { get; set; }

		public string CommandIdName { get; set; }
	}

	private sealed class NativeCommandPermissionDecision
	{
		public bool Allowed { get; set; }

		public Document Document { get; set; }

		public string CurrentUser { get; set; }

		public string Role { get; set; }

		public NativeCommandPermissionDecision()
		{
			CurrentUser = string.Empty;
			Role = string.Empty;
		}
	}

	private sealed class ProtectedElementInfo
	{
		public string Kind { get; set; }

		public string Name { get; set; }

		public string ElementName { get; set; }

		public string CategoryName { get; set; }

		public ProtectedElementInfo()
		{
			Kind = string.Empty;
			Name = string.Empty;
			ElementName = string.Empty;
			CategoryName = string.Empty;
		}
	}

	private sealed class ProtectedChangeEvent
	{
		public string Action { get; set; }

		public string Kind { get; set; }

		public string Name { get; set; }

		public string OriginalName { get; set; }

		public string OriginalElementName { get; set; }

		public string CategoryName { get; set; }

		public string ElementIdText { get; set; }

		public string State { get; set; }

		public string RecoveryStatus { get; set; }

		public string RequiredAction { get; set; }

		public string PolicyReason { get; set; }

		public string ParentFamilyNames { get; set; }

		public string DetectedAtUtc { get; set; }

		public ProtectedChangeEvent()
		{
			Action = string.Empty;
			Kind = string.Empty;
			Name = string.Empty;
			OriginalName = string.Empty;
			OriginalElementName = string.Empty;
			CategoryName = string.Empty;
			ElementIdText = string.Empty;
			State = string.Empty;
			RecoveryStatus = string.Empty;
			RequiredAction = string.Empty;
			PolicyReason = string.Empty;
			ParentFamilyNames = string.Empty;
			DetectedAtUtc = string.Empty;
		}
	}

	private sealed class CachedNativeGuardDecision
	{
		public bool CanEditFamilies { get; set; }

		public bool CanAddDeleteTypes { get; set; }

		public DateTime CachedUtc { get; set; }
	}

	private sealed class CachedNestedOnlyPlacementCatalog
	{
		public FamilyBrowserNestedOnlyPlacementCatalog Catalog { get; set; }

		public DateTime CachedUtc { get; set; }
	}

	private static readonly object SyncRoot = RuntimeHelpers.GetObjectValue(new object());

	private static readonly List<BoundNativeCommand> BoundCommands = new List<BoundNativeCommand>();

	private static readonly Dictionary<string, Dictionary<int, ProtectedElementInfo>> ProtectedElementIndexes = new Dictionary<string, Dictionary<int, ProtectedElementInfo>>(StringComparer.OrdinalIgnoreCase);

	private static readonly Dictionary<string, int> CompleteProtectedElementIndexDocumentTokens = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

	private static readonly Dictionary<string, CachedNativeGuardDecision> NativeGuardDecisionCache = new Dictionary<string, CachedNativeGuardDecision>(StringComparer.OrdinalIgnoreCase);

	private static readonly Dictionary<string, CachedNestedOnlyPlacementCatalog> NestedOnlyPlacementCatalogCache = new Dictionary<string, CachedNestedOnlyPlacementCatalog>(StringComparer.OrdinalIgnoreCase);

	private static readonly Dictionary<string, string> CentralPathCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

	private static readonly Dictionary<string, string> MappedDriveRootCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

	private static UIControlledApplication ControlledApplication;

	private static Document LastActiveDocument;

	private static DateTime LastWarningUtc = DateTime.MinValue;

	private static DateTime LastBlockedNativeActionUtc = DateTime.MinValue;

	private static string LastBlockedNativeActionName = string.Empty;

	private static bool GuardInitialized = false;

	private static int TrustedOperationDepth = 0;

	private static DateTime TrustedOperationUntilUtc = DateTime.MinValue;

	private const double PolicyCacheSeconds = 60.0;

	private const double NativeGuardDecisionCacheSeconds = 30.0;

	private const double NestedOnlyPlacementCatalogCacheSeconds = 30.0;

	private const long SlowGuardTimingThresholdMilliseconds = 250L;

	private const double BlockedUndoCooldownSeconds = 10.0;

	private const int RibbonStateSettlePassCount = 3;

	private const int PostRollbackUiRefreshPassCount = 2;

	private static FamilyBrowserStandardPolicy CachedPolicy;

	private static string CachedPolicyUser = string.Empty;

	private static DateTime CachedPolicyLoadedUtc = DateTime.MinValue;

	private static string CachedPolicyStamp = string.Empty;

	private static bool DatabaseEventsAttached = false;

	private static ProtectedContentChangeUpdater ProtectedChangeUpdaterInstance;

	private static DateTime LastRibbonAvailabilityUpdateUtc = DateTime.MinValue;

	private static bool LastRibbonLoadFamilyAllowed = true;

	private static bool LastRibbonLoadFamilyKnown = false;

	private static readonly List<object> CachedLoadFamilyRibbonControls = new List<object>();

	private static int PendingRibbonRefreshPasses = 0;

	private static bool PendingProtectedElementBaselineRefresh = false;

	private static int PendingPostRollbackUiRefreshPasses = 0;

	private static bool UiEventsAttached = false;

	private static bool AdminModeStateKnown = false;

	private static bool AdminModeEnabledForNativeGuard = false;

	private static string AdminModeStateUser = string.Empty;

	private static readonly Guid ProtectedChangeUpdaterGuid = new Guid("49DDE62E-26D7-4A79-9D6D-CED0A197A23D");

	private static readonly FailureDefinitionId ProtectedChangeFailureId = new FailureDefinitionId(new Guid("11DB20D4-2631-43B3-B9D4-601E24C70C67"));

	private static bool ProtectedChangeFailureRegistered = false;

	private static readonly BuiltInParameter[] ProtectedNameChangeParameters = new BuiltInParameter[6]
	{
		BuiltInParameter.SYMBOL_NAME_PARAM,
		BuiltInParameter.ALL_MODEL_TYPE_NAME,
		BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM,
		BuiltInParameter.ALL_MODEL_FAMILY_NAME,
		BuiltInParameter.FAMILY_NAME_PSEUDO_PARAM,
		BuiltInParameter.ELEM_FAMILY_PARAM
	};

	private static readonly ProtectedNativeCommandDefinition FamilyLoadingEventDefinition = new ProtectedNativeCommandDefinition
	{
		Key = "native-family-loading-event",
		DisplayNameEn = "Revit family load",
		DisplayNameKo = "Revit 패밀리 로드",
		RequiredPermission = "EditFamilies"
	};

	private static readonly ProtectedNativeCommandDefinition FamilyDocumentSaveEventDefinition = new ProtectedNativeCommandDefinition
	{
		Key = "native-family-document-save-event",
		DisplayNameEn = "Family document save",
		DisplayNameKo = "패밀리 문서 저장",
		RequiredPermission = "EditFamilies",
		FamilyDocumentOnly = true
	};

	private static readonly ProtectedNativeCommandDefinition FamilyDocumentSaveAsEventDefinition = new ProtectedNativeCommandDefinition
	{
		Key = "native-family-document-save-as-event",
		DisplayNameEn = "Family document save as",
		DisplayNameKo = "패밀리 문서 다른 이름으로 저장",
		RequiredPermission = "EditFamilies",
		FamilyDocumentOnly = true
	};

	private FamilyBrowserNativeCommandGuardService()
	{
	}

	private static bool IsKoreanUi()
	{
		try
		{
			return !string.Equals(FamilyBrowserUserSettingsStore.LoadLanguageCode(), "en", StringComparison.OrdinalIgnoreCase);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return true;
	}

	private static string UiText(string englishText, string koreanText)
	{
		if (!IsKoreanUi())
		{
			return englishText;
		}
		return koreanText;
	}

	private static bool IsAdminModeEnabledForNativeGuard()
	{
		string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (AdminModeStateKnown && string.Equals(AdminModeStateUser, currentUser, StringComparison.OrdinalIgnoreCase))
			{
				return AdminModeEnabledForNativeGuard;
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		bool enabled = LoadAdminModeEnabledSetting();
		object syncRoot2 = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot2);
		bool lockTaken2 = false;
		try
		{
			Monitor.Enter(syncRoot2, ref lockTaken2);
			AdminModeEnabledForNativeGuard = enabled;
			AdminModeStateKnown = true;
			AdminModeStateUser = currentUser;
		}
		finally
		{
			if (lockTaken2)
			{
				Monitor.Exit(syncRoot2);
			}
		}
		return enabled;
	}

	private static bool LoadAdminModeEnabledSetting()
	{
		try
		{
			FamilyBrowserStandardPolicy policy = LoadPolicy();
			string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
			bool isAdminProfile = string.Equals(FamilyBrowserSecurityPolicyService.ResolveRole(policy, currentUser, null), "Admin", StringComparison.OrdinalIgnoreCase);
			return isAdminProfile && FamilyBrowserUserSettingsStore.LoadAdminModeEnabled();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
			return false;
		}
	}

	public static bool ResolveInitialAdminModeEnabled()
	{
		return IsAdminModeEnabledForNativeGuard();
	}

	private static bool CanNativeGuardPermission(FamilyBrowserStandardPolicy policy, string currentUser, string permission, FamilyBrowserProjectPolicyContext context)
	{
		if (FamilyBrowserPermissionExcelPolicyService.IsNativeGuardPermission(permission) && !IsFileGuardTargetedDocument(policy, context))
		{
			return true;
		}
		if (string.Equals(permission, "RenameFamilyOrType", StringComparison.OrdinalIgnoreCase))
		{
			return FamilyBrowserSecurityPolicyService.CanNativeGuard(policy, currentUser, "EditFamilies", context, IsAdminModeEnabledForNativeGuard()) && FamilyBrowserSecurityPolicyService.CanNativeGuard(policy, currentUser, "AddDeleteTypes", context, IsAdminModeEnabledForNativeGuard());
		}
		return FamilyBrowserSecurityPolicyService.CanNativeGuard(policy, currentUser, permission, context, IsAdminModeEnabledForNativeGuard());
	}

	private static string ResolveNativeGuardRoleLabel(FamilyBrowserStandardPolicy policy, string currentUser, FamilyBrowserProjectPolicyContext context)
	{
		string role = FamilyBrowserSecurityPolicyService.ResolveRole(policy, currentUser, context);
		if (string.Equals(role, "Admin", StringComparison.OrdinalIgnoreCase) && !IsAdminModeEnabledForNativeGuard())
		{
			return UiText("Admin profile (Admin Mode Off)", "관리자 프로필 (관리자 모드 OFF)");
		}
		return role;
	}

	public static void Start(UIControlledApplication application)
	{
		if (application == null)
		{
			return;
		}
		FamilyBrowserRevitVersionContext.SetCurrentVersion(application);
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (GuardInitialized)
			{
				return;
			}
			foreach (ProtectedNativeCommandDefinition definition in BuildCommandDefinitions())
			{
				BindCommand(application, definition);
			}
			ControlledApplication = application;
			AttachUiEvents(application);
			AttachDatabasePreEvents(application);
			// UpdaterRegistry registration is only valid while Revit is invoking the
			// external application. Register once during OnStartup and let Execute
			// decide whether the active document is a File Guard target.
			RegisterProtectedChangeUpdater(application);
			GuardInitialized = true;
			// Do not read shared policy or inspect the active document during add-in startup.
			// Native command bindings answer CanExecute on demand, and the dashboard/admin
			// toggles refresh the ribbon when a user actually changes policy state.
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static void Stop()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			foreach (BoundNativeCommand item in BoundCommands)
			{
				try
				{
					item.Binding.CanExecute -= HandleCanExecute;
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				try
				{
					item.Binding.BeforeExecuted -= HandleBeforeExecuted;
				}
				catch (Exception projectError2)
				{
					ProjectData.SetProjectError(projectError2);
					ProjectData.ClearProjectError();
				}
			}
			DetachUiEvents();
			DetachDatabasePreEvents();
			UnregisterProtectedChangeUpdater();
			BoundCommands.Clear();
			ProtectedElementIndexes.Clear();
			CompleteProtectedElementIndexDocumentTokens.Clear();
			CentralPathCache.Clear();
			MappedDriveRootCache.Clear();
			ControlledApplication = null;
			LastActiveDocument = null;
			GuardInitialized = false;
			LastBlockedNativeActionUtc = DateTime.MinValue;
			LastBlockedNativeActionName = string.Empty;
			TrustedOperationDepth = 0;
			TrustedOperationUntilUtc = DateTime.MinValue;
			CachedPolicy = null;
			CachedPolicyUser = string.Empty;
			CachedPolicyLoadedUtc = DateTime.MinValue;
			CachedPolicyStamp = string.Empty;
			NativeGuardDecisionCache.Clear();
			NestedOnlyPlacementCatalogCache.Clear();
			FamilyBrowserNestedOnlyPlacementRuntimeService.ResetAll();
			AdminModeStateKnown = false;
			AdminModeStateUser = string.Empty;
			DatabaseEventsAttached = false;
			ProtectedChangeUpdaterInstance = null;
			LastRibbonAvailabilityUpdateUtc = DateTime.MinValue;
			LastRibbonLoadFamilyAllowed = true;
			LastRibbonLoadFamilyKnown = false;
			CachedLoadFamilyRibbonControls.Clear();
			PendingRibbonRefreshPasses = 0;
			PendingProtectedElementBaselineRefresh = false;
			PendingPostRollbackUiRefreshPasses = 0;
			UiEventsAttached = false;
			AdminModeStateKnown = false;
			AdminModeEnabledForNativeGuard = false;
			AdminModeStateUser = string.Empty;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static IDisposable BeginTrustedOperation(string description)
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		checked
		{
			try
			{
				Monitor.Enter(syncRoot, ref lockTaken);
				TrustedOperationDepth++;
				TrustedOperationUntilUtc = DateTime.UtcNow.AddSeconds(2.0);
			}
			finally
			{
				if (lockTaken)
				{
					Monitor.Exit(syncRoot);
				}
			}
			return new TrustedOperationScope();
		}
	}

	public static void NotifyPolicyChanged()
	{
		InvalidatePolicyCacheOnly();
		RefreshActiveDocumentGuardState();
	}

	public static void NotifyStandardSnapshotChanged()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			NestedOnlyPlacementCatalogCache.Clear();
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		FamilyBrowserNestedOnlyPlacementRuntimeService.InvalidateCatalogs();
		if (LastActiveDocument != null)
		{
			ScheduleNestedOnlyFingerprintRefreshIfRequired(LastActiveDocument, LoadPolicy());
		}
	}

	private static void AttachUiEvents(UIControlledApplication application)
	{
		if (application == null || UiEventsAttached)
		{
			return;
		}
		application.Idling += HandleIdling;
		UiEventsAttached = true;
	}

	private static void DetachUiEvents()
	{
		if (!UiEventsAttached || ControlledApplication == null)
		{
			return;
		}
		try
		{
			ControlledApplication.Idling -= HandleIdling;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		UiEventsAttached = false;
	}

	private static void ScheduleProtectedRibbonRefresh()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			PendingRibbonRefreshPasses = Math.Max(PendingRibbonRefreshPasses, RibbonStateSettlePassCount);
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	private static void ScheduleProtectedElementBaselineRefresh()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			PendingProtectedElementBaselineRefresh = true;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	private static void SchedulePostRollbackUiRefresh()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			PendingPostRollbackUiRefreshPasses = Math.Max(PendingPostRollbackUiRefreshPasses, PostRollbackUiRefreshPassCount);
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	private static void HandleIdling(object sender, IdlingEventArgs e)
	{
		bool refreshPending = false;
		bool baselinePending = false;
		bool rollbackUiRefreshPending = false;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (PendingRibbonRefreshPasses > 0)
			{
				PendingRibbonRefreshPasses--;
				refreshPending = true;
			}
			baselinePending = PendingProtectedElementBaselineRefresh;
			PendingProtectedElementBaselineRefresh = false;
			if (PendingPostRollbackUiRefreshPasses > 0)
			{
				PendingPostRollbackUiRefreshPasses--;
				rollbackUiRefreshPending = true;
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		if (baselinePending)
		{
			EnsureProtectedElementIndexForGuard(LastActiveDocument);
		}
		if (refreshPending)
		{
			UpdateProtectedRibbonAvailability(force: true);
		}
		if (rollbackUiRefreshPending)
		{
			RefreshRevitUiAfterProtectedRollback(sender);
		}
		if (LastActiveDocument != null && FamilyBrowserNestedOnlyPlacementRuntimeService.HasPending(LastActiveDocument))
		{
			FamilyBrowserStandardPolicy policy = LoadPolicy();
			FamilyBrowserNestedOnlyPlacementRuntimeService.ProcessNextPending(sender as UIApplication, LastActiveDocument, policy, ResolveAssignedFileGuardDiscipline(LastActiveDocument, policy));
		}
	}

	private static void RefreshRevitUiAfterProtectedRollback(object sender)
	{
		try
		{
			UIApplication uiApplication = sender as UIApplication;
			UIDocument uiDocument = uiApplication?.ActiveUIDocument;
			if (uiDocument == null)
			{
				return;
			}
			LastActiveDocument = uiDocument.Document;
			uiDocument.RefreshActiveView();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Protected native change post-rollback UI refresh failed", ex, string.Empty);
			ProjectData.ClearProjectError();
		}
	}

	public static void NotifyPolicyChanged(FamilyBrowserStandardPolicy policy)
	{
		if (policy == null)
		{
			NotifyPolicyChanged();
			return;
		}
		SeedPolicyCache(policy);
		RefreshActiveDocumentGuardState();
	}

	public static void RefreshActiveDocumentGuardState()
	{
		Document document = LastActiveDocument;
		if (document != null)
		{
			try
			{
				SyncCurrentUserIdentity(document);
				RefreshProtectedChangeUpdaterRegistration(document);
				ScheduleNestedOnlyFingerprintRefreshIfRequired(document, LoadPolicy());
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			LastRibbonAvailabilityUpdateUtc = DateTime.MinValue;
			LastRibbonLoadFamilyKnown = false;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		ScheduleProtectedElementBaselineRefresh();
		ScheduleProtectedRibbonRefresh();
		UpdateProtectedRibbonAvailability(force: true);
	}

	public static void InvalidatePolicyCacheOnly()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			CachedPolicy = null;
			CachedPolicyUser = string.Empty;
			CachedPolicyLoadedUtc = DateTime.MinValue;
			CachedPolicyStamp = string.Empty;
			NativeGuardDecisionCache.Clear();
			NestedOnlyPlacementCatalogCache.Clear();
			MappedDriveRootCache.Clear();
			FamilyBrowserNestedOnlyPlacementRuntimeService.InvalidateCatalogs();
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	public static void NotifyAdminModeChanged(bool enabled, bool refreshUiNow = true)
	{
		NotifyAdminModeChanged(enabled, null, refreshUiNow);
	}

	public static void NotifyAdminModeChanged(bool enabled, FamilyBrowserStandardPolicy policy, bool refreshUiNow = true)
	{
		if (policy != null)
		{
			SeedPolicyCache(policy);
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			AdminModeEnabledForNativeGuard = enabled;
			AdminModeStateKnown = true;
			AdminModeStateUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
			LastRibbonAvailabilityUpdateUtc = DateTime.MinValue;
			LastRibbonLoadFamilyKnown = false;
			NativeGuardDecisionCache.Clear();
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		if (!enabled)
		{
			ScheduleProtectedElementBaselineRefresh();
		}
		if (refreshUiNow && LastActiveDocument != null)
		{
			try
			{
				RefreshProtectedChangeUpdaterRegistration(LastActiveDocument);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		ScheduleProtectedRibbonRefresh();
		if (refreshUiNow)
		{
			UpdateProtectedRibbonAvailability(force: true);
		}
	}

	public static void NotifyActiveDocumentChanged(Document document)
	{
		Stopwatch sw = Stopwatch.StartNew();
		LastActiveDocument = document;
		if (document != null)
		{
			SyncCurrentUserIdentity(document);
			ResolveCentralPath(document, allowResolve: true);
			// Activate the transaction-level guard before the first family/type edit.
			RefreshProtectedChangeUpdaterRegistration(document);
			ScheduleNestedOnlyFingerprintRefreshIfRequired(document, LoadPolicy());
		}
		ScheduleProtectedElementBaselineRefresh();
		ScheduleProtectedRibbonRefresh();
		UpdateProtectedRibbonAvailability(force: true);
		WriteGuardTiming("NotifyActiveDocumentChanged", sw, "Document=" + SafeDocumentTitle(document));
	}

	public static string BuildRuntimeGuardDiagnostic(Document document)
	{
		try
		{
			FamilyBrowserStandardPolicy policy = LoadPolicy();
			FamilyBrowserProjectPolicyContext context = BuildProjectContext(document, policy, includeCentralPath: true);
			string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
			FamilyBrowserFileGuardMatchResult fileGuardMatch = FamilyBrowserFileGuardPathMatcher.Resolve(policy?.FileGuard, context);
			bool targeted = fileGuardMatch.Target != null;
			bool canLoad = CanNativeGuardPermission(policy, currentUser, "LoadFamilies", context);
			bool canEdit = CanNativeGuardPermission(policy, currentUser, "EditFamilies", context);
			bool canTypes = CanNativeGuardPermission(policy, currentUser, "AddDeleteTypes", context);
			int bindingCount;
			bool updaterRegistered;
			int ribbonControlCount;
			bool ribbonStateKnown;
			bool ribbonLoadAllowed;
			bool projectBrowserRenameBound;
			int pendingRibbonRefreshPasses;
			bool protectedElementBaselineComplete;
			bool pendingProtectedElementBaselineRefresh;
			int pendingPostRollbackUiRefreshPasses;
			string documentKey = BuildDocumentKey(document);
			int documentToken = (document == null) ? 0 : RuntimeHelpers.GetHashCode(document);
			object syncRoot = SyncRoot;
			ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
			bool lockTaken = false;
			try
			{
				Monitor.Enter(syncRoot, ref lockTaken);
				bindingCount = BoundCommands.Count;
				updaterRegistered = ProtectedChangeUpdaterInstance != null;
				ribbonControlCount = CachedLoadFamilyRibbonControls.Count;
				ribbonStateKnown = LastRibbonLoadFamilyKnown;
				ribbonLoadAllowed = LastRibbonLoadFamilyAllowed;
				projectBrowserRenameBound = BoundCommands.Any([SpecialName] (BoundNativeCommand x) => string.Equals(x?.CommandIdName, "ID_PRJBROWSER_RENAME", StringComparison.OrdinalIgnoreCase));
				pendingRibbonRefreshPasses = PendingRibbonRefreshPasses;
				int completeDocumentToken;
				protectedElementBaselineComplete = !string.IsNullOrWhiteSpace(documentKey) && CompleteProtectedElementIndexDocumentTokens.TryGetValue(documentKey, out completeDocumentToken) && completeDocumentToken == documentToken;
				pendingProtectedElementBaselineRefresh = PendingProtectedElementBaselineRefresh;
				pendingPostRollbackUiRefreshPasses = PendingPostRollbackUiRefreshPasses;
			}
			finally
			{
				if (lockTaken)
				{
					Monitor.Exit(syncRoot);
				}
			}
			return "mode=" + (IsAdminModeEnabledForNativeGuard() ? "on" : "off") +
				";targeted=" + targeted.ToString() +
				";canLoad=" + canLoad.ToString() +
				";canEdit=" + canEdit.ToString() +
				";canTypes=" + canTypes.ToString() +
				";bindings=" + bindingCount.ToString(CultureInfo.InvariantCulture) +
				";updater=" + updaterRegistered.ToString() +
				";ribbonControls=" + ribbonControlCount.ToString(CultureInfo.InvariantCulture) +
				";ribbonState=" + (ribbonStateKnown ? (ribbonLoadAllowed ? "enabled" : "disabled") : "unknown") +
				";projectBrowserRenameBinding=" + projectBrowserRenameBound.ToString() +
				";pendingRibbonRefreshPasses=" + pendingRibbonRefreshPasses.ToString(CultureInfo.InvariantCulture) +
				";protectedElementBaselineComplete=" + protectedElementBaselineComplete.ToString() +
				";pendingProtectedElementBaselineRefresh=" + pendingProtectedElementBaselineRefresh.ToString() +
				";pendingPostRollbackUiRefreshPasses=" + pendingPostRollbackUiRefreshPasses.ToString(CultureInfo.InvariantCulture) +
				";" + FamilyBrowserNestedOnlyPlacementRuntimeService.BuildDiagnostic(document) +
				";document=" + SafeDocumentTitle(document) +
				";modelPath=" + ((context == null) ? string.Empty : (context.ModelPath ?? string.Empty)) +
				";centralPath=" + ((context == null) ? string.Empty : (context.CentralPath ?? string.Empty)) +
				";fileGuardMatch=" + FamilyBrowserFileGuardPathMatcher.Describe(fileGuardMatch);
		}
		catch (Exception ex)
		{
			return "diagnostic-error=" + ex.GetType().Name + ":" + (ex.Message ?? string.Empty).Replace("\r", " ").Replace("\n", " ");
		}
	}

	public static void HandleDocumentChanged(DocumentChangedEventArgs e)
	{
		if (e == null)
		{
			return;
		}
		Document doc = null;
		try
		{
			doc = e.GetDocument();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			doc = LastActiveDocument;
			ProjectData.ClearProjectError();
		}
		if (doc == null)
		{
			return;
		}
		LastActiveDocument = doc;
		SyncCurrentUserIdentity(doc);
		Stopwatch sw = Stopwatch.StartNew();
		FamilyBrowserStandardPolicy policy = LoadPolicy();
		FamilyBrowserProjectPolicyContext context = BuildProjectContext(doc, policy, includeCentralPath: true);
		if (!IsFileGuardTargetedDocument(policy, context))
		{
			UpdateProtectedRibbonAvailability();
			WriteGuardTiming("HandleDocumentChanged.not-targeted", sw, "Document=" + SafeDocumentTitle(doc));
			return;
		}
		FamilyBrowserNestedOnlyPlacementRuntimeService.InvalidateChangedFamilies(doc, e.GetAddedElementIds(), e.GetModifiedElementIds(), e.GetDeletedElementIds());
		if (IsTrustedOperationActive())
		{
			UpdateProtectedElementIndexFromChanges(doc, e.GetAddedElementIds(), e.GetModifiedElementIds(), e.GetDeletedElementIds());
			return;
		}
		string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		string role = ResolveNativeGuardRoleLabel(policy, currentUser, context);
		bool canEditFamilies;
		bool canAddDeleteTypes;
		ResolveNativeGuardPermissions(policy, currentUser, context, out canEditFamilies, out canAddDeleteTypes);
		WriteGuardTiming("HandleDocumentChanged.permission", sw, "CanEditFamilies=" + canEditFamilies.ToString() + Environment.NewLine + "CanAddDeleteTypes=" + canAddDeleteTypes.ToString() + Environment.NewLine + "Document=" + SafeDocumentTitle(doc));
		if (canEditFamilies && canAddDeleteTypes)
		{
			UpdateProtectedElementIndexFromChanges(doc, e.GetAddedElementIds(), e.GetModifiedElementIds(), e.GetDeletedElementIds());
			UpdateProtectedRibbonAvailability();
			return;
		}
		UpdateProtectedRibbonAvailability();
		string documentKey = BuildDocumentKey(doc);
		Dictionary<int, ProtectedElementInfo> previousIndex = null;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			ProtectedElementIndexes.TryGetValue(documentKey, out previousIndex);
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		List<ProtectedChangeEvent> events = new List<ProtectedChangeEvent>();
		if (previousIndex == null)
		{
			CollectAddedOrModifiedWithoutBaseline(events, doc, e.GetAddedElementIds(), "Added", canEditFamilies, canAddDeleteTypes);
			CollectAddedOrModifiedWithoutBaseline(events, doc, e.GetModifiedElementIds(), "Modified", canEditFamilies, canAddDeleteTypes);
			CollectDeletedFallback(events, e.GetDeletedElementIds(), canEditFamilies, canAddDeleteTypes);
			UpdateProtectedElementIndexFromChanges(doc, e.GetAddedElementIds(), e.GetModifiedElementIds(), e.GetDeletedElementIds());
			if (events.Count != 0)
			{
				WriteAudit(doc, currentUser, role, events);
				MarkDeletedProtectedContentDirty(doc, currentUser, events);
			}
			return;
		}
		CollectAddedOrModified(events, doc, previousIndex, e.GetAddedElementIds(), "Added", canEditFamilies, canAddDeleteTypes);
		CollectAddedOrModified(events, doc, previousIndex, e.GetModifiedElementIds(), "Modified", canEditFamilies, canAddDeleteTypes);
		CollectDeleted(events, previousIndex, e.GetDeletedElementIds(), canEditFamilies, canAddDeleteTypes);
		UpdateProtectedElementIndexFromChanges(doc, e.GetAddedElementIds(), e.GetModifiedElementIds(), e.GetDeletedElementIds());
		if (events.Count != 0)
		{
			WriteAudit(doc, currentUser, role, events);
			MarkDeletedProtectedContentDirty(doc, currentUser, events);
		}
	}

	private static bool IsTrustedOperationActive()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			return TrustedOperationDepth > 0 || DateTime.Compare(DateTime.UtcNow, TrustedOperationUntilUtc) <= 0;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	private static void EndTrustedOperation()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		checked
		{
			try
			{
				Monitor.Enter(syncRoot, ref lockTaken);
				if (TrustedOperationDepth > 0)
				{
					TrustedOperationDepth--;
				}
				TrustedOperationUntilUtc = DateTime.UtcNow.AddSeconds(2.0);
			}
			finally
			{
				if (lockTaken)
				{
					Monitor.Exit(syncRoot);
				}
			}
		}
	}

	private static void AttachDatabasePreEvents(UIControlledApplication application)
	{
		if (application != null && application.ControlledApplication != null && !DatabaseEventsAttached)
		{
			try
			{
				application.ControlledApplication.FamilyLoadingIntoDocument += HandleFamilyLoadingIntoDocument;
				application.ControlledApplication.DocumentSaving += HandleDocumentSaving;
				application.ControlledApplication.DocumentSavingAs += HandleDocumentSavingAs;
				application.ControlledApplication.FailuresProcessing += HandleFailuresProcessing;
				DatabaseEventsAttached = true;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				DatabaseEventsAttached = false;
				ProjectData.ClearProjectError();
			}
		}
	}

	private static void DetachDatabasePreEvents()
	{
		if (ControlledApplication != null && ControlledApplication.ControlledApplication != null && DatabaseEventsAttached)
		{
			try
			{
				ControlledApplication.ControlledApplication.FamilyLoadingIntoDocument -= HandleFamilyLoadingIntoDocument;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			try
			{
				ControlledApplication.ControlledApplication.DocumentSaving -= HandleDocumentSaving;
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				ProjectData.ClearProjectError();
			}
			try
			{
				ControlledApplication.ControlledApplication.DocumentSavingAs -= HandleDocumentSavingAs;
			}
			catch (Exception projectError3)
			{
				ProjectData.SetProjectError(projectError3);
				ProjectData.ClearProjectError();
			}
			try
			{
				ControlledApplication.ControlledApplication.FailuresProcessing -= HandleFailuresProcessing;
			}
			catch (Exception projectError4)
			{
				ProjectData.SetProjectError(projectError4);
				ProjectData.ClearProjectError();
			}
			DatabaseEventsAttached = false;
		}
	}

	private static void HandleFailuresProcessing(object sender, FailuresProcessingEventArgs e)
	{
		if (e == null)
		{
			return;
		}
		try
		{
			FailuresAccessor failuresAccessor = e.GetFailuresAccessor();
			if (failuresAccessor == null)
			{
				return;
			}
			IList<FailureMessageAccessor> messages = failuresAccessor.GetFailureMessages();
			if (messages == null || messages.Count == 0)
			{
				return;
			}
			bool hasProtectedChangeFailure = false;
			foreach (FailureMessageAccessor message in messages)
			{
				if (message != null && IsProtectedChangeFailureId(message.GetFailureDefinitionId()))
				{
					hasProtectedChangeFailure = true;
					break;
				}
			}
			if (hasProtectedChangeFailure)
			{
				ScheduleProtectedElementBaselineRefresh();
				SchedulePostRollbackUiRefresh();
				e.SetProcessingResult(FailureProcessingResult.ProceedWithRollBack);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Protected native change rollback processing failed", ex2, string.Empty);
			ProjectData.ClearProjectError();
		}
	}

	private static bool IsProtectedChangeFailureId(FailureDefinitionId failureId)
	{
		if ((object)failureId == null)
		{
			return false;
		}
		if (object.Equals(failureId, ProtectedChangeFailureId))
		{
			return true;
		}
		string actual = ReadFailureDefinitionGuidText(failureId);
		string expected = ReadFailureDefinitionGuidText(ProtectedChangeFailureId);
		return !string.IsNullOrWhiteSpace(actual) && string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
	}

	private static string ReadFailureDefinitionGuidText(FailureDefinitionId failureId)
	{
		if ((object)failureId == null)
		{
			return string.Empty;
		}
		string[] array = new string[2] { "Guid", "GUID" };
		foreach (string propertyName in array)
		{
			try
			{
				PropertyInfo propertyInfo = failureId.GetType().GetProperty(propertyName);
				if ((object)propertyInfo != null)
				{
					object value = RuntimeHelpers.GetObjectValue(propertyInfo.GetValue(failureId, null));
					if (value != null)
					{
						return Convert.ToString(RuntimeHelpers.GetObjectValue(value), CultureInfo.InvariantCulture);
					}
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		return failureId.ToString();
	}

	private static void HandleFamilyLoadingIntoDocument(object sender, FamilyLoadingIntoDocumentEventArgs e)
	{
		if (e == null || IsTrustedOperationActive())
		{
			return;
		}
		Document doc = e.Document;
		if (doc != null)
		{
			LastActiveDocument = doc;
			SyncCurrentUserIdentity(doc);
		}
		NativeCommandPermissionDecision decision = EvaluateNativeCommandPermission(FamilyLoadingEventDefinition, doc);
		if (!decision.Allowed)
		{
			RememberBlockedNativeAction(FamilyLoadingEventDefinition);
			if (TryCancelPreEvent(e.Cancellable, [SpecialName] () =>
			{
				e.Cancel();
			}))
			{
				WriteBlockedCommandAudit(decision.Document, decision.CurrentUser, decision.Role, FamilyLoadingEventDefinition, "FamilyName=" + (e.FamilyName ?? string.Empty) + Environment.NewLine + "FamilyPath=" + (e.FamilyPath ?? string.Empty));
				TaskDialog.Show("KKY Family Browser", BuildBlockedCommandMessage(FamilyLoadingEventDefinition, decision.CurrentUser, decision.Role));
			}
			else
			{
				WriteBlockedCommandAudit(decision.Document, decision.CurrentUser, decision.Role, FamilyLoadingEventDefinition, "Family load event was not cancellable." + Environment.NewLine + "FamilyName=" + (e.FamilyName ?? string.Empty) + Environment.NewLine + "FamilyPath=" + (e.FamilyPath ?? string.Empty));
			}
		}
	}

	private static void HandleDocumentSaving(object sender, DocumentSavingEventArgs e)
	{
		HandleFamilyDocumentSaveEvent(e, FamilyDocumentSaveEventDefinition);
	}

	private static void HandleDocumentSavingAs(object sender, DocumentSavingAsEventArgs e)
	{
		HandleFamilyDocumentSaveEvent(e, FamilyDocumentSaveAsEventDefinition);
	}

	private static void HandleFamilyDocumentSaveEvent(RevitAPIPreDocEventArgs e, ProtectedNativeCommandDefinition definition)
	{
		if (e == null || definition == null || IsTrustedOperationActive())
		{
			return;
		}
		Document doc = e.Document;
		if (doc == null || !doc.IsFamilyDocument)
		{
			return;
		}
		LastActiveDocument = doc;
		SyncCurrentUserIdentity(doc);
		NativeCommandPermissionDecision decision = EvaluateNativeCommandPermission(definition, doc);
		if (!decision.Allowed)
		{
			RememberBlockedNativeAction(definition);
			if (TryCancelPreEvent(e.Cancellable, [SpecialName] () =>
			{
				e.Cancel();
			}))
			{
				WriteBlockedCommandAudit(decision.Document, decision.CurrentUser, decision.Role, definition);
				TaskDialog.Show("KKY Family Browser", BuildBlockedCommandMessage(definition, decision.CurrentUser, decision.Role));
			}
			else
			{
				WriteBlockedCommandAudit(decision.Document, decision.CurrentUser, decision.Role, definition, "Family document save event was not cancellable.");
			}
		}
	}

	private static bool TryCancelPreEvent(bool cancellable, Action cancelAction)
	{
		bool TryCancelPreEvent;
		if (!cancellable || cancelAction == null)
		{
			TryCancelPreEvent = false;
		}
		else
		{
			try
			{
				cancelAction();
				TryCancelPreEvent = true;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				TryCancelPreEvent = false;
				ProjectData.ClearProjectError();
			}
		}
		return TryCancelPreEvent;
	}

	private static void RememberBlockedNativeAction(ProtectedNativeCommandDefinition definition)
	{
		RememberBlockedNativeAction((definition == null) ? string.Empty : NativeCommandDisplayName(definition));
	}

	private static void RememberBlockedNativeAction(string actionName)
	{
		LastBlockedNativeActionUtc = DateTime.UtcNow;
		LastBlockedNativeActionName = actionName ?? string.Empty;
	}

	private static bool IsBlockedNativeActionCooldownActive()
	{
		if (DateTime.Compare(LastBlockedNativeActionUtc, DateTime.MinValue) == 0)
		{
			return false;
		}
		return DateTime.UtcNow.Subtract(LastBlockedNativeActionUtc).TotalSeconds <= 10.0;
	}

	private static string BuildBlockedUndoMessage()
	{
		string actionName = (string.IsNullOrWhiteSpace(LastBlockedNativeActionName) ? UiText("the blocked Revit command", "차단된 Revit 명령") : LastBlockedNativeActionName);
		return UiText("The previous Revit command was already blocked before it was committed.", "직전 Revit 명령은 커밋되기 전에 이미 차단되었습니다.") + "\r\n\r\n" + UiText("Action: ", "작업: ") + actionName + "\r\n" + UiText("There is no automatic restore transaction to undo. Continue working through Family Browser.", "되돌릴 자동 복구 트랜잭션은 없습니다. Family Browser를 통해 작업을 계속하세요.");
	}
	private static void RegisterProtectedChangeUpdater(UIControlledApplication application)
	{
		if (application == null)
		{
			return;
		}
		try
		{
			EnsureProtectedChangeFailureDefinition();
			AddInId addInId = ResolveProtectedChangeAddInId(application);
			if (addInId == null)
			{
				FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Protected native change updater registration skipped", new InvalidOperationException("Family Browser could not resolve an AddInId for the protected native change updater."), string.Empty);
				return;
			}
			ProtectedContentChangeUpdater updater = new ProtectedContentChangeUpdater(addInId);
			UpdaterId updaterId = updater.GetUpdaterId();
			try
			{
				if (UpdaterRegistry.IsUpdaterRegistered(updaterId))
				{
					UpdaterRegistry.RemoveAllTriggers(updaterId);
					UpdaterRegistry.UnregisterUpdater(updaterId);
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			UpdaterRegistry.RegisterUpdater(updater);
			RegisterProtectedChangeTriggers(updaterId);
			ProtectedChangeUpdaterInstance = updater;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Protected native change updater registration failed", ex2, string.Empty);
			ProjectData.ClearProjectError();
		}
	}

	private static void RefreshProtectedChangeUpdaterRegistration(Document document)
	{
		// Registration is intentionally startup-scoped. Calling UpdaterRegistry
		// from this modeless dashboard path is rejected by Revit and used to leave
		// family/type changes completely unguarded.
		if (ProtectedChangeUpdaterInstance == null)
		{
			WriteGuardTiming("RefreshProtectedChangeUpdaterRegistration.missing", Stopwatch.StartNew(), "Document=" + SafeDocumentTitle(document));
		}
	}

	private static bool ShouldEnableProtectedChangeUpdater(Document doc)
	{
		if (doc == null || doc.IsFamilyDocument)
		{
			return false;
		}
		Stopwatch sw = Stopwatch.StartNew();
		try
		{
			FamilyBrowserStandardPolicy policy = LoadPolicy();
			FamilyBrowserProjectPolicyContext context = BuildProjectContext(doc, policy, includeCentralPath: true);
			bool shouldEnable = IsFileGuardTargetedDocument(policy, context);
			WriteGuardTiming("ShouldEnableProtectedChangeUpdater", sw, "ShouldEnable=" + shouldEnable.ToString() + Environment.NewLine + "Document=" + SafeDocumentTitle(doc));
			return shouldEnable;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return false;
	}

	private static AddInId ResolveProtectedChangeAddInId(UIControlledApplication application)
	{
		try
		{
			if (application?.ActiveAddInId != null)
			{
				return application.ActiveAddInId;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			string assemblyName = Assembly.GetExecutingAssembly().GetName().Name ?? string.Empty;
			Guid fallbackGuid;
			if (assemblyName.IndexOf("2027", StringComparison.OrdinalIgnoreCase) >= 0)
			{
				fallbackGuid = new Guid("8A0D0C84-8E3C-4A1B-B7D8-00C2CECB2027");
			}
			else if (assemblyName.IndexOf("2025", StringComparison.OrdinalIgnoreCase) >= 0)
			{
				fallbackGuid = new Guid("9E6E4A65-38B8-4EE7-9A1D-0D621E335F41");
			}
			else
			{
				fallbackGuid = new Guid("7D1E58FC-5B4C-43A1-9121-6D7B6D640201");
			}
			return new AddInId(fallbackGuid);
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		return null;
	}

	private static void EnsureProtectedChangeFailureDefinition()
	{
		if (!ProtectedChangeFailureRegistered)
		{
			try
			{
				FailureDefinition.CreateFailureDefinition(ProtectedChangeFailureId, FailureSeverity.Error, UiText("KKY Family Browser blocked a protected family/type change or direct placement of a nested-only family. Use Family Browser or ask an administrator.", "KKY Family Browser가 보호된 패밀리/타입 변경 또는 하위 전용 패밀리의 단독 배치를 차단했습니다. Family Browser를 사용하거나 관리자에게 문의하세요."));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			ProtectedChangeFailureRegistered = true;
		}
	}

	private static void RegisterProtectedChangeTriggers(UpdaterId updaterId)
	{
		if (updaterId != null)
		{
			RegisterProtectedChangeTriggers(updaterId, new ElementClassFilter(typeof(Family)));
			RegisterProtectedChangeTriggers(updaterId, new ElementClassFilter(typeof(FamilySymbol)));
			RegisterProtectedChangeTriggers(updaterId, new ElementClassFilter(typeof(ElementType)));
			RegisterNestedOnlyPlacementTrigger(updaterId);
		}
	}

	private static void RegisterNestedOnlyPlacementTrigger(UpdaterId updaterId)
	{
		if (updaterId == null)
		{
			return;
		}
		try
		{
			UpdaterRegistry.AddTrigger(updaterId, new ElementClassFilter(typeof(FamilyInstance)), Element.GetChangeTypeElementAddition());
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void RegisterProtectedChangeTriggers(UpdaterId updaterId, ElementFilter filter)
	{
		if (updaterId == null || filter == null)
		{
			return;
		}
		try
		{
			UpdaterRegistry.AddTrigger(updaterId, filter, Element.GetChangeTypeAny());
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			UpdaterRegistry.AddTrigger(updaterId, filter, Element.GetChangeTypeElementAddition());
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		try
		{
			UpdaterRegistry.AddTrigger(updaterId, filter, Element.GetChangeTypeElementDeletion());
		}
		catch (Exception projectError3)
		{
			ProjectData.SetProjectError(projectError3);
			ProjectData.ClearProjectError();
		}
		foreach (ElementId parameterId in BuildProtectedNameParameterIds())
		{
			try
			{
				UpdaterRegistry.AddTrigger(updaterId, filter, Element.GetChangeTypeParameter(parameterId));
			}
			catch (Exception projectError4)
			{
				ProjectData.SetProjectError(projectError4);
				ProjectData.ClearProjectError();
			}
		}
	}

	private static List<ElementId> BuildProtectedNameParameterIds()
	{
		List<ElementId> ids = new List<ElementId>();
		HashSet<int> seen = new HashSet<int>();
		BuiltInParameter[] protectedNameChangeParameters = ProtectedNameChangeParameters;
		for (int i = 0; i < protectedNameChangeParameters.Length; i = checked(i + 1))
		{
			int value = (int)protectedNameChangeParameters[i];
			if (seen.Add(value))
			{
				ids.Add(new ElementId(value));
			}
		}
		return ids;
	}

	private static void UnregisterProtectedChangeUpdater()
	{
		ProtectedContentChangeUpdater updater = ProtectedChangeUpdaterInstance;
		if (updater == null)
		{
			return;
		}
		try
		{
			UpdaterId updaterId = updater.GetUpdaterId();
			if (updaterId != null && UpdaterRegistry.IsUpdaterRegistered(updaterId))
			{
				UpdaterRegistry.RemoveAllTriggers(updaterId);
				UpdaterRegistry.UnregisterUpdater(updaterId);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		ProtectedChangeUpdaterInstance = null;
	}

	private static void BindCommand(UIControlledApplication application, ProtectedNativeCommandDefinition definition)
	{
		foreach (RevitCommandId commandId in ResolveCommandIds(definition))
		{
			if (commandId != null)
			{
				try
				{
					AddInCommandBinding binding = application.CreateAddInCommandBinding(commandId);
					binding.CanExecute += HandleCanExecute;
					binding.BeforeExecuted += HandleBeforeExecuted;
					BoundCommands.Add(new BoundNativeCommand
					{
						Binding = binding,
						Definition = definition,
						CommandIdName = commandId.Name ?? string.Empty
					});
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
	}

	private static List<RevitCommandId> ResolveCommandIds(ProtectedNativeCommandDefinition definition)
	{
		List<RevitCommandId> ids = new List<RevitCommandId>();
		if (definition == null)
		{
			return ids;
		}
		foreach (string commandName in definition.PostableCommandNames)
		{
			if (string.IsNullOrWhiteSpace(commandName))
			{
				continue;
			}
			try
			{
				if (Enum.TryParse<PostableCommand>(commandName, ignoreCase: true, out var parsed))
				{
					RevitCommandId commandId = RevitCommandId.LookupPostableCommandId(parsed);
					if (commandId != null)
					{
						ids.Add(commandId);
					}
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		foreach (string idName in definition.CommandIdNames)
		{
			if (string.IsNullOrWhiteSpace(idName))
			{
				continue;
			}
			try
			{
				RevitCommandId commandId2 = RevitCommandId.LookupCommandId(idName);
				if (commandId2 != null)
				{
					ids.Add(commandId2);
				}
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				ProjectData.ClearProjectError();
			}
		}
		return ids;
	}

	private static void UpdateProtectedRibbonAvailability(bool force = false)
	{
		try
		{
			DateTime now = DateTime.UtcNow;
			if (!force && (now - LastRibbonAvailabilityUpdateUtc).TotalSeconds < 1.0)
			{
				return;
			}
			LastRibbonAvailabilityUpdateUtc = now;
			// The pre-document FamilyLoading event uses the same permission as the
			// native Load Family command. Use it as a fail-closed fallback because
			// some Revit releases do not expose a bindable Load Family command id.
			ProtectedNativeCommandDefinition definition = FindBoundDefinition("native-load-family") ?? FamilyLoadingEventDefinition;
			if (definition != null)
			{
				bool allowed = EvaluateNativeCommandPermission(definition, LastActiveDocument, includeCentralPath: true).Allowed;
				if (force || !LastRibbonLoadFamilyKnown || allowed != LastRibbonLoadFamilyAllowed)
				{
					LastRibbonLoadFamilyAllowed = allowed;
					LastRibbonLoadFamilyKnown = true;
					ApplyLoadFamilyRibbonEnabledFast(allowed);
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static ProtectedNativeCommandDefinition FindBoundDefinition(string key)
	{
		if (string.IsNullOrWhiteSpace(key))
		{
			return null;
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			foreach (BoundNativeCommand item in BoundCommands)
			{
				if (item != null && item.Definition != null && string.Equals(item.Definition.Key, key, StringComparison.OrdinalIgnoreCase))
				{
					return item.Definition;
				}
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		return null;
	}

	private static void ApplyLoadFamilyRibbonEnabledFast(bool enabled)
	{
		Stopwatch sw = Stopwatch.StartNew();
		int appliedCount = 0;
		try
		{
			foreach (object control in ResolveLoadFamilyRibbonControlsFast())
			{
				SetBooleanProperty(control, "IsEnabled", enabled);
				SetBooleanProperty(control, "Enabled", enabled);
				appliedCount = checked(appliedCount + 1);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		WriteGuardTiming("ApplyLoadFamilyRibbonEnabledFast", sw, "Enabled=" + enabled.ToString() + Environment.NewLine + "Controls=" + appliedCount.ToString(CultureInfo.InvariantCulture));
	}

	private static object ResolveAutodeskWindowsRibbon()
	{
		object ResolveAutodeskWindowsRibbon;
		try
		{
			ResolveAutodeskWindowsRibbon = Type.GetType("Autodesk.Windows.ComponentManager, AdWindows")?.GetProperty("Ribbon")?.GetValue(null, null);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveAutodeskWindowsRibbon = null;
			ProjectData.ClearProjectError();
		}
		return ResolveAutodeskWindowsRibbon;
	}

	private static List<object> ResolveLoadFamilyRibbonControlsFast()
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (CachedLoadFamilyRibbonControls.Count > 0)
			{
				return CachedLoadFamilyRibbonControls.ToList();
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}

		Stopwatch sw = Stopwatch.StartNew();
		List<object> matches = new List<object>();
		HashSet<int> visited = new HashSet<int>();
		object ribbon = RuntimeHelpers.GetObjectValue(ResolveAutodeskWindowsRibbon());
		foreach (object tab in EnumerateRibbonProperty(ribbon, "Tabs"))
		{
			foreach (object panel in EnumerateRibbonProperty(tab, "Panels"))
			{
				object source = RuntimeHelpers.GetObjectValue(GetPropertyValue(panel, "Source"));
				CollectLoadFamilyRibbonControls(source ?? panel, matches, visited, 0);
			}
		}

		object syncRoot2 = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot2);
		bool lockTaken2 = false;
		try
		{
			Monitor.Enter(syncRoot2, ref lockTaken2);
			CachedLoadFamilyRibbonControls.Clear();
			CachedLoadFamilyRibbonControls.AddRange(matches);
		}
		finally
		{
			if (lockTaken2)
			{
				Monitor.Exit(syncRoot2);
			}
		}
		WriteGuardTiming("ResolveLoadFamilyRibbonControlsFast", sw, "Controls=" + matches.Count.ToString(CultureInfo.InvariantCulture));
		return matches;
	}

	private static void CollectLoadFamilyRibbonControls(object target, List<object> matches, HashSet<int> visited, int depth)
	{
		if (target == null || matches == null || visited == null || depth > 6)
		{
			return;
		}
		int key = RuntimeHelpers.GetHashCode(RuntimeHelpers.GetObjectValue(target));
		if (!visited.Add(key))
		{
			return;
		}
		if (IsLoadFamilyRibbonControl(RuntimeHelpers.GetObjectValue(target)))
		{
			matches.Add(RuntimeHelpers.GetObjectValue(target));
		}
		string[] childProperties = new string[4] { "Items", "LargeItems", "MediumItems", "SmallItems" };
		foreach (string propertyName in childProperties)
		{
			foreach (object child in EnumerateRibbonProperty(target, propertyName))
			{
				CollectLoadFamilyRibbonControls(child, matches, visited, checked(depth + 1));
			}
		}
	}

	private static List<object> EnumerateRibbonProperty(object target, string propertyName)
	{
		List<object> result = new List<object>();
		object value = RuntimeHelpers.GetObjectValue(GetPropertyValue(target, propertyName));
		if (value == null || value is string)
		{
			return result;
		}
		if (!(value is IEnumerable enumerable))
		{
			result.Add(RuntimeHelpers.GetObjectValue(value));
			return result;
		}
		try
		{
			foreach (object itemValue in enumerable)
			{
				object item = RuntimeHelpers.GetObjectValue(itemValue);
				if (item != null && !(item is string))
				{
					result.Add(RuntimeHelpers.GetObjectValue(item));
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static bool IsLoadFamilyRibbonControl(object target)
	{
		if (target == null)
		{
			return false;
		}
		List<string> values = new List<string>();
		string[] array = new string[12]
		{
			"Id", "Name", "Text", "AutomationName", "ItemText", "Title", "Caption", "CommandId", "CommandParameter", "Description",
			"ToolTip", "ToolTipText"
		};
		foreach (string propertyName in array)
		{
			object value = RuntimeHelpers.GetObjectValue(GetPropertyValue(RuntimeHelpers.GetObjectValue(target), propertyName));
			if (value != null)
			{
				if (value is string)
				{
					values.Add((string)value);
				}
				else
				{
					values.Add(Convert.ToString(RuntimeHelpers.GetObjectValue(value), CultureInfo.InvariantCulture));
				}
			}
		}
		foreach (string item in values)
		{
			string token = NormalizeRibbonToken(item);
			if (token.Length != 0 && (token.Contains("idloadfamily") || token.Contains("idfamilyload") || token.Contains("idinsertloadfamily") || token.Contains("idobjectsloadfamily") || token.Contains("idloadautodeskfamily") || token.Equals("loadfamily", StringComparison.Ordinal) || token.Equals("loadfamilysymbol", StringComparison.Ordinal) || token.Equals("loadautodeskfamily", StringComparison.Ordinal) || token.Equals(BuildKoreanLoadFamilyToken(), StringComparison.Ordinal)))
			{
				return true;
			}
		}
		return false;
	}

	private static string BuildKoreanLoadFamilyToken()
	{
		return "패밀리로드";
	}
	private static string NormalizeRibbonToken(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		StringBuilder builder = new StringBuilder();
		foreach (char ch in value)
		{
			if (char.IsLetterOrDigit(ch) || ch == '_' || ch == '-' || (ch >= '가' && ch <= '힣'))
			{
				builder.Append(char.ToLowerInvariant(ch));
			}
		}
		return builder.ToString().Replace("_", string.Empty).Replace("-", string.Empty);
	}

	private static object GetPropertyValue(object target, string propertyName)
	{
		object GetPropertyValue;
		if (target == null || string.IsNullOrWhiteSpace(propertyName))
		{
			GetPropertyValue = null;
		}
		else
		{
			try
			{
				GetPropertyValue = target.GetType().GetProperty(propertyName)?.GetValue(RuntimeHelpers.GetObjectValue(target), null);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				GetPropertyValue = null;
				ProjectData.ClearProjectError();
			}
		}
		return GetPropertyValue;
	}

	private static void SetBooleanProperty(object target, string propertyName, bool value)
	{
		if (target == null || string.IsNullOrWhiteSpace(propertyName))
		{
			return;
		}
		try
		{
			PropertyInfo propertyInfo = target.GetType().GetProperty(propertyName);
			if ((object)propertyInfo != null && propertyInfo.CanWrite)
			{
				propertyInfo.SetValue(RuntimeHelpers.GetObjectValue(target), value, null);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static ProtectedNativeCommandDefinition ResolveBoundDefinition(object sender)
	{
		if (!(sender is AddInCommandBinding binding))
		{
			return null;
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			foreach (BoundNativeCommand item in BoundCommands)
			{
				if (object.ReferenceEquals(item.Binding, binding))
				{
					return item.Definition;
				}
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		return null;
	}

	private static void HandleCanExecute(object sender, CanExecuteEventArgs e)
	{
		ProtectedNativeCommandDefinition definition = ResolveBoundDefinition(RuntimeHelpers.GetObjectValue(sender));
		if (definition == null)
		{
			return;
		}
		try
		{
			if (definition.BlockDuringBlockedActionCooldownOnly)
			{
				e.CanExecute = !IsBlockedNativeActionCooldownActive();
				return;
			}
			if (LastActiveDocument == null)
			{
				e.CanExecute = true;
				return;
			}
			NativeCommandPermissionDecision decision = EvaluateNativeCommandPermission(definition, LastActiveDocument, includeCentralPath: false);
			e.CanExecute = decision.Allowed;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void HandleBeforeExecuted(object sender, BeforeExecutedEventArgs e)
	{
		ProtectedNativeCommandDefinition definition = ResolveBoundDefinition(RuntimeHelpers.GetObjectValue(sender));
		if (definition == null)
		{
			return;
		}
		if (definition.BlockDuringBlockedActionCooldownOnly)
		{
			if (IsBlockedNativeActionCooldownActive())
			{
				try
				{
					e.Cancel = true;
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				TaskDialog.Show("KKY Family Browser", BuildBlockedUndoMessage());
			}
			return;
		}
		NativeCommandPermissionDecision decision = EvaluateNativeCommandPermission(definition);
		string currentUser = decision.CurrentUser;
		string role = decision.Role;
		if (string.Equals(definition.Key, "native-load-family", StringComparison.OrdinalIgnoreCase))
		{
			UpdateProtectedRibbonAvailability();
		}
		if (!decision.Allowed)
		{
			try
			{
				e.Cancel = true;
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				ProjectData.ClearProjectError();
			}
			RememberBlockedNativeAction(definition);
			WriteBlockedCommandAudit(decision.Document, decision.CurrentUser, decision.Role, definition);
			TaskDialog.Show("KKY Family Browser", BuildBlockedCommandMessage(definition, currentUser, role));
		}
	}

	private static string BuildBlockedCommandMessage(ProtectedNativeCommandDefinition definition, string currentUser, string role)
	{
		return UiText("This Revit command is blocked by the Family Browser operation policy.", "이 Revit 명령은 Family Browser 운영 정책으로 차단되었습니다.") + "\r\n\r\n" + UiText("Action: ", "작업: ") + NativeCommandDisplayName(definition) + "\r\n" + UiText("User: ", "사용자: ") + (currentUser ?? string.Empty) + "\r\n" + UiText("Current role: ", "현재 권한: ") + (role ?? string.Empty) + "\r\n\r\n" + UiText("Load standard families and apply system types through Family Browser.", "표준 패밀리 로드와 시스템 타입 적용은 Family Browser에서 실행하세요.") + "\r\n" + UiText("If this task requires administrator permission, send a request to the BIM manager.", "관리자 권한이 필요한 작업이면 BIM 관리자에게 요청하세요.");
	}
	private static NativeCommandPermissionDecision EvaluateNativeCommandPermission(ProtectedNativeCommandDefinition definition)
	{
		return EvaluateNativeCommandPermission(definition, LastActiveDocument);
	}

	private static NativeCommandPermissionDecision EvaluateNativeCommandPermission(ProtectedNativeCommandDefinition definition, Document doc)
	{
		return EvaluateNativeCommandPermission(definition, doc, includeCentralPath: true);
	}

	private static NativeCommandPermissionDecision EvaluateNativeCommandPermission(ProtectedNativeCommandDefinition definition, Document doc, bool includeCentralPath)
	{
		if (definition != null && definition.FamilyDocumentOnly)
		{
			string currentUserForFamilyDocumentCommand = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
			if (doc == null || !doc.IsFamilyDocument)
			{
				return new NativeCommandPermissionDecision
				{
					Allowed = true,
					Document = doc,
					CurrentUser = currentUserForFamilyDocumentCommand,
					Role = string.Empty
				};
			}
		}
		Stopwatch sw = Stopwatch.StartNew();
		FamilyBrowserStandardPolicy policy = LoadPolicy();
		FamilyBrowserProjectPolicyContext context = BuildProjectContext(doc, policy, includeCentralPath);
		string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		string role = ResolveNativeGuardRoleLabel(policy, currentUser, context);
		bool allowed = CanNativeGuardPermission(policy, currentUser, definition.RequiredPermission, context);
		WriteGuardTiming("EvaluateNativeCommandPermission", sw, "Action=" + ((definition == null) ? string.Empty : definition.Key) + Environment.NewLine + "IncludeCentralPath=" + includeCentralPath.ToString() + Environment.NewLine + "Allowed=" + allowed.ToString() + Environment.NewLine + "Document=" + SafeDocumentTitle(doc));
		return new NativeCommandPermissionDecision
		{
			Allowed = allowed,
			Document = doc,
			CurrentUser = currentUser,
			Role = role
		};
	}

	private static string NativeCommandDisplayName(ProtectedNativeCommandDefinition definition)
	{
		if (definition == null)
		{
			return UiText("Revit native command", "Revit 기본 명령");
		}
		switch (definition.Key)
		{
		case "native-load-family":
			return UiText("Revit native Load Family", "Revit 기본 패밀리 로드");
		case "native-edit-family":
			return UiText("Edit Family", "패밀리 편집");
		case "native-family-load-into-project":
			return UiText("Load Into Project", "프로젝트에 로드");
		case "native-family-save":
			return UiText("Family Save", "패밀리 저장");
		case "native-family-save-as":
			return UiText("Family Save As", "패밀리 다른 이름으로 저장");
		case "native-family-loading-event":
			return UiText("Revit family load", "Revit 패밀리 로드");
		case "native-family-document-save-event":
			return UiText("Family document save", "패밀리 문서 저장");
		case "native-family-document-save-as-event":
			return UiText("Family document save as", "패밀리 문서 다른 이름으로 저장");
		case "native-transfer-project-standards":
			return UiText("Transfer Project Standards", "프로젝트 표준 전송");
		case "native-purge-unused":
			return UiText("Purge Unused", "사용하지 않는 항목 소거");
		case "native-rename-family-or-type":
			return UiText("Rename family or type", "패밀리/타입 이름 변경");
		default:
			if (!IsKoreanUi() && !string.IsNullOrWhiteSpace(definition.DisplayNameEn))
			{
				return definition.DisplayNameEn;
			}
			if (IsKoreanUi() && !string.IsNullOrWhiteSpace(definition.DisplayNameKo))
			{
				return definition.DisplayNameKo;
			}
			return definition.Key;
		}
	}
	private static List<ProtectedNativeCommandDefinition> BuildCommandDefinitions()
	{
		return new List<ProtectedNativeCommandDefinition>
		{
			new ProtectedNativeCommandDefinition
			{
				Key = "native-undo-after-blocked-action",
				DisplayNameEn = "Undo after blocked command",
				DisplayNameKo = "차단 직후 되돌리기",
				PostableCommandNames = new List<string> { "Undo" },
				CommandIdNames = new List<string> { "ID_EDIT_UNDO", "ID_UNDO" },
				BlockDuringBlockedActionCooldownOnly = true
			},
			new ProtectedNativeCommandDefinition
			{
				Key = "native-redo-after-blocked-action",
				DisplayNameEn = "Redo after blocked command",
				DisplayNameKo = "차단 직후 다시 실행",
				PostableCommandNames = new List<string> { "Redo" },
				CommandIdNames = new List<string> { "ID_EDIT_REDO", "ID_REDO" },
				BlockDuringBlockedActionCooldownOnly = true
			},
			new ProtectedNativeCommandDefinition
			{
				Key = "native-load-family",
				DisplayNameEn = "Revit native Load Family",
				DisplayNameKo = "Revit 기본 패밀리 로드",
				RequiredPermission = "EditFamilies",
				PostableCommandNames = new List<string> { "LoadFamily", "LoadFamilySymbol", "LoadAutodeskFamily" },
				CommandIdNames = new List<string> { "ID_LOAD_FAMILY", "ID_FAMILY_LOAD", "ID_LOAD_FAMILY_SYMBOL", "ID_INSERT_LOAD_FAMILY_SYMBOL", "ID_LOAD_AUTODESK_FAMILY", "ID_OBJECTS_LOAD_FAMILY", "ID_INSERT_LOAD_FAMILY" }
			},
			new ProtectedNativeCommandDefinition
			{
				Key = "native-rename-family-or-type",
				DisplayNameEn = "Rename family or type",
				DisplayNameKo = "패밀리/타입 이름 변경",
				RequiredPermission = "RenameFamilyOrType",
				CommandIdNames = new List<string> { "ID_PRJBROWSER_RENAME", "ID_EDIT_RENAME", "ID_RENAME", "ID_PROJECT_BROWSER_RENAME", "ID_BROWSER_RENAME", "ID_OBJECTS_RENAME", "ID_ELEMENT_RENAME", "ID_TYPE_RENAME", "ID_FAMILY_RENAME" }
			},
			new ProtectedNativeCommandDefinition
			{
				Key = "native-edit-family",
				DisplayNameEn = "Edit Family",
				DisplayNameKo = "패밀리 편집",
				RequiredPermission = "EditFamilies",
				PostableCommandNames = new List<string> { "EditFamily" },
				CommandIdNames = new List<string> { "ID_EDIT_FAMILY", "ID_FAMILY_EDIT" }
			},
			new ProtectedNativeCommandDefinition
			{
				Key = "native-family-load-into-project",
				DisplayNameEn = "Load Into Project",
				DisplayNameKo = "프로젝트에 로드",
				RequiredPermission = "EditFamilies",
				PostableCommandNames = new List<string> { "LoadIntoProject", "LoadIntoProjectAndClose" },
				CommandIdNames = new List<string> { "ID_FAMILY_LOAD_INTO_PROJECT", "ID_FAMILY_LOAD_INTO_PROJECT_AND_CLOSE", "ID_LOAD_INTO_PROJECT", "ID_LOAD_INTO_PROJECT_AND_CLOSE", "ID_OBJECTS_LOAD_INTO_PROJECT", "ID_OBJECTS_LOAD_INTO_PROJECT_AND_CLOSE" },
				FamilyDocumentOnly = true
			},
			new ProtectedNativeCommandDefinition
			{
				Key = "native-family-save",
				DisplayNameEn = "Family Save",
				DisplayNameKo = "패밀리 저장",
				RequiredPermission = "EditFamilies",
				PostableCommandNames = new List<string> { "Save" },
				CommandIdNames = new List<string> { "ID_REVIT_FILE_SAVE", "ID_FILE_SAVE", "ID_SAVE", "ID_FAMILY_SAVE" },
				FamilyDocumentOnly = true
			},
			new ProtectedNativeCommandDefinition
			{
				Key = "native-family-save-as",
				DisplayNameEn = "Family Save As",
				DisplayNameKo = "패밀리 다른 이름으로 저장",
				RequiredPermission = "EditFamilies",
				PostableCommandNames = new List<string> { "SaveAs" },
				CommandIdNames = new List<string> { "ID_REVIT_FILE_SAVE_AS", "ID_FILE_SAVE_AS", "ID_SAVE_AS", "ID_FAMILY_SAVE_AS" },
				FamilyDocumentOnly = true
			},
			new ProtectedNativeCommandDefinition
			{
				Key = "native-transfer-project-standards",
				DisplayNameEn = "Transfer Project Standards",
				DisplayNameKo = "프로젝트 표준 전송",
				RequiredPermission = "AddDeleteTypes",
				PostableCommandNames = new List<string> { "TransferProjectStandards" },
				CommandIdNames = new List<string> { "ID_TRANSFER_PROJECT_STANDARDS", "ID_TRANSFER_STANDARDS" }
			},
			new ProtectedNativeCommandDefinition
			{
				Key = "native-purge-unused",
				DisplayNameEn = "Purge Unused",
				DisplayNameKo = "사용하지 않는 항목 소거",
				RequiredPermission = "AddDeleteTypes",
				PostableCommandNames = new List<string> { "PurgeUnused" },
				CommandIdNames = new List<string> { "ID_PURGE_UNUSED" }
			}
		};
	}

	private static void CollectAddedOrModified(List<ProtectedChangeEvent> events, Document doc, Dictionary<int, ProtectedElementInfo> previousIndex, ICollection<ElementId> ids, string action, bool canEditFamilies, bool canAddDeleteTypes)
	{
		if (ids == null)
		{
			return;
		}
		foreach (ElementId id in ids)
		{
			Element element = null;
			try
			{
				element = doc.GetElement(id);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			ProtectedElementInfo info = BuildProtectedElementInfo(element);
			if (info != null && (!string.Equals(info.Kind, "Family", StringComparison.OrdinalIgnoreCase) || !canEditFamilies) && (string.Equals(info.Kind, "Family", StringComparison.OrdinalIgnoreCase) || !canAddDeleteTypes))
			{
				ProtectedElementInfo previousInfo = null;
				previousIndex?.TryGetValue(GetElementIdInteger(id), out previousInfo);
				bool previousInfoAvailable = previousInfo != null;
				bool sameProtectedInfo = previousInfoAvailable && ProtectedElementInfoEquals(previousInfo, info);
				if (ShouldRecordProtectedChange(action, previousInfoAvailable, sameProtectedInfo))
				{
					events.Add(new ProtectedChangeEvent
					{
						Action = action,
						Kind = info.Kind,
						Name = info.Name,
						OriginalName = ((previousInfo == null) ? string.Empty : previousInfo.Name),
						OriginalElementName = ((previousInfo == null) ? string.Empty : previousInfo.ElementName),
						CategoryName = info.CategoryName,
						ElementIdText = GetElementIdText(id),
						State = (string.Equals(action, "Added", StringComparison.OrdinalIgnoreCase) ? "UnauthorizedAddedProtectedContent" : "UnauthorizedModifiedProtectedContent"),
						RecoveryStatus = "LoggedOnlyNativeCommandMustBeBlockedBeforeExecution",
						RequiredAction = (string.Equals(action, "Added", StringComparison.OrdinalIgnoreCase) ? "ReviewNativeCommandGuardCoverage" : "ReviewNativeCommandGuardCoverage"),
						DetectedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
					});
				}
			}
		}
	}

	private static void CollectDeleted(List<ProtectedChangeEvent> events, Dictionary<int, ProtectedElementInfo> previousIndex, ICollection<ElementId> ids, bool canEditFamilies, bool canAddDeleteTypes)
	{
		if (previousIndex == null || ids == null)
		{
			return;
		}
		foreach (ElementId id in ids)
		{
			int key = GetElementIdInteger(id);
			ProtectedElementInfo info = null;
			if (previousIndex.TryGetValue(key, out info) && (!string.Equals(info.Kind, "Family", StringComparison.OrdinalIgnoreCase) || !canEditFamilies) && (string.Equals(info.Kind, "Family", StringComparison.OrdinalIgnoreCase) || !canAddDeleteTypes))
			{
				events.Add(new ProtectedChangeEvent
				{
					Action = "Deleted",
					Kind = info.Kind,
					Name = info.Name,
					OriginalName = info.Name,
					OriginalElementName = info.ElementName,
					CategoryName = info.CategoryName,
					ElementIdText = GetElementIdText(id),
					State = "UnauthorizedDeletedProtectedContent",
					RecoveryStatus = "NeedsAdminRestoreFromStandard",
					RequiredAction = "CurrentModelCheckRequired",
					DetectedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
				});
			}
		}
	}

	private static void CollectDeletedFallback(List<ProtectedChangeEvent> events, ICollection<ElementId> ids, bool canEditFamilies, bool canAddDeleteTypes)
	{
		if (events == null || ids == null || (canEditFamilies && canAddDeleteTypes))
		{
			return;
		}
		foreach (ElementId id in ids)
		{
			events.Add(new ProtectedChangeEvent
			{
				Action = "Deleted",
				Kind = "Protected Family/Type",
				Name = "(deleted element)",
				OriginalName = "(deleted element)",
				OriginalElementName = string.Empty,
				CategoryName = string.Empty,
				ElementIdText = GetElementIdText(id),
				State = "UnauthorizedDeletedProtectedContent",
				RecoveryStatus = "BlockedBeforeCommit",
				RequiredAction = "UseFamilyBrowserOrAskAdministrator",
				DetectedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
			});
		}
	}

	private static bool ShouldRecordProtectedChange(string action, bool previousInfoAvailable, bool sameProtectedInfo)
	{
		return !string.Equals(action, "Modified", StringComparison.OrdinalIgnoreCase) || !previousInfoAvailable || !sameProtectedInfo;
	}

	private static void CollectAddedOrModifiedWithoutBaseline(List<ProtectedChangeEvent> events, Document doc, ICollection<ElementId> ids, string action, bool canEditFamilies, bool canAddDeleteTypes)
	{
		if (events == null || doc == null || ids == null)
		{
			return;
		}
		foreach (ElementId id in ids)
		{
			Element element = null;
			try
			{
				element = doc.GetElement(id);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			ProtectedElementInfo info = BuildProtectedElementInfo(element);
			if (info == null || !ShouldBlockProtectedInfo(info, canEditFamilies, canAddDeleteTypes))
			{
				continue;
			}
			events.Add(new ProtectedChangeEvent
			{
				Action = action,
				Kind = info.Kind,
				Name = info.Name,
				OriginalName = string.Empty,
				OriginalElementName = string.Empty,
				CategoryName = info.CategoryName,
				ElementIdText = GetElementIdText(id),
				State = (string.Equals(action, "Added", StringComparison.OrdinalIgnoreCase) ? "UnauthorizedAddedProtectedContent" : "UnauthorizedModifiedProtectedContent"),
				RecoveryStatus = "BlockedBeforeCommit",
				RequiredAction = "UseFamilyBrowserOrAskAdministrator",
				DetectedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
			});
		}
	}

	private static bool ShouldBlockProtectedInfo(ProtectedElementInfo info, bool canEditFamilies, bool canAddDeleteTypes)
	{
		if (info == null)
		{
			return false;
		}
		if (string.Equals(info.Kind, "Family", StringComparison.OrdinalIgnoreCase))
		{
			return !canEditFamilies;
		}
		return !canAddDeleteTypes;
	}

	private static void EnsureProtectedElementIndexForGuard(Document doc)
	{
		if (doc == null || doc.IsFamilyDocument)
		{
			return;
		}
		try
		{
			FamilyBrowserStandardPolicy policy = LoadPolicy();
			FamilyBrowserProjectPolicyContext context = BuildProjectContext(doc, policy, includeCentralPath: false);
			string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
			bool canEditFamilies;
			bool canAddDeleteTypes;
			ResolveNativeGuardPermissions(policy, currentUser, context, out canEditFamilies, out canAddDeleteTypes);
			if (canEditFamilies && canAddDeleteTypes)
			{
				return;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
			return;
		}
		EnsureProtectedElementIndexBaseline(doc);
	}

	private static void EnsureProtectedElementIndexBaseline(Document doc)
	{
		if (doc == null || doc.IsFamilyDocument)
		{
			return;
		}
		string documentKey = BuildDocumentKey(doc);
		if (string.IsNullOrWhiteSpace(documentKey))
		{
			return;
		}
		int documentToken = RuntimeHelpers.GetHashCode(doc);
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			int completeDocumentToken;
			if (ProtectedElementIndexes.ContainsKey(documentKey) && CompleteProtectedElementIndexDocumentTokens.TryGetValue(documentKey, out completeDocumentToken) && completeDocumentToken == documentToken)
			{
				return;
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		Stopwatch sw = Stopwatch.StartNew();
		RefreshProtectedElementIndex(doc);
		WriteGuardTiming("EnsureProtectedElementIndexBaseline", sw, "Document=" + SafeDocumentTitle(doc));
	}

	private static Dictionary<int, ProtectedElementInfo> GetProtectedElementIndexSnapshot(Document doc)
	{
		if (doc == null)
		{
			return null;
		}
		string documentKey = BuildDocumentKey(doc);
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			Dictionary<int, ProtectedElementInfo> index = null;
			if (!ProtectedElementIndexes.TryGetValue(documentKey, out index) || index == null)
			{
				return null;
			}
			return index.ToDictionary([SpecialName] (KeyValuePair<int, ProtectedElementInfo> x) => x.Key, [SpecialName] (KeyValuePair<int, ProtectedElementInfo> x) => x.Value);
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	private static void HandleProtectedContentUpdaterExecute(UpdaterData data)
	{
		if (data == null || IsTrustedOperationActive())
		{
			return;
		}
		Document doc = null;
		try
		{
			doc = data.GetDocument();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			doc = LastActiveDocument;
			ProjectData.ClearProjectError();
		}
		if (doc == null || doc.IsFamilyDocument)
		{
			return;
		}
		LastActiveDocument = doc;
		SyncCurrentUserIdentity(doc);
		Stopwatch sw = Stopwatch.StartNew();
		FamilyBrowserStandardPolicy policy = LoadPolicy();
		FamilyBrowserProjectPolicyContext context = BuildProjectContext(doc, policy);
		FamilyBrowserFileGuardTarget matchingTarget = FindMatchingNativeGuardTarget(policy, context);
		if (matchingTarget == null)
		{
			WriteGuardTiming("ProtectedContentUpdater.not-targeted", sw, "Document=" + SafeDocumentTitle(doc));
			return;
		}
		bool blockNestedOnlyPlacement = ShouldBlockNestedOnlyStandalonePlacement(matchingTarget, IsAdminModeEnabledForNativeGuard());
		FamilyBrowserNestedOnlyPlacementRuntimeService.InvalidateChangedFamilies(doc, data.GetAddedElementIds(), data.GetModifiedElementIds(), data.GetDeletedElementIds());
		string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		string role = ResolveNativeGuardRoleLabel(policy, currentUser, context);
		bool canEditFamilies;
		bool canAddDeleteTypes;
		ResolveNativeGuardPermissions(policy, currentUser, context, out canEditFamilies, out canAddDeleteTypes);
		WriteGuardTiming("ProtectedContentUpdater.permission", sw, "CanEditFamilies=" + canEditFamilies.ToString() + Environment.NewLine + "CanAddDeleteTypes=" + canAddDeleteTypes.ToString() + Environment.NewLine + "BlockNestedOnlyPlacement=" + blockNestedOnlyPlacement.ToString() + Environment.NewLine + "Document=" + SafeDocumentTitle(doc));
		if (canEditFamilies && canAddDeleteTypes && !blockNestedOnlyPlacement)
		{
			return;
		}
		Dictionary<int, ProtectedElementInfo> previousIndex = GetProtectedElementIndexSnapshot(doc);
		List<ProtectedChangeEvent> events = new List<ProtectedChangeEvent>();
		if (previousIndex == null)
		{
			CollectAddedOrModifiedWithoutBaseline(events, doc, data.GetAddedElementIds(), "Added", canEditFamilies, canAddDeleteTypes);
			CollectAddedOrModifiedWithoutBaseline(events, doc, data.GetModifiedElementIds(), "Modified", canEditFamilies, canAddDeleteTypes);
			CollectDeletedFallback(events, data.GetDeletedElementIds(), canEditFamilies, canAddDeleteTypes);
		}
		else
		{
			CollectAddedOrModified(events, doc, previousIndex, data.GetAddedElementIds(), "Added", canEditFamilies, canAddDeleteTypes);
			CollectAddedOrModified(events, doc, previousIndex, data.GetModifiedElementIds(), "Modified", canEditFamilies, canAddDeleteTypes);
			CollectDeleted(events, previousIndex, data.GetDeletedElementIds(), canEditFamilies, canAddDeleteTypes);
		}
		if (blockNestedOnlyPlacement)
		{
			CollectNestedOnlyStandalonePlacements(events, doc, data.GetAddedElementIds(), policy, FamilyBrowserFileGuardDisciplineService.ResolveAssignedDiscipline(policy, matchingTarget, allowLegacyFallback: false));
		}
		if (events.Count == 0)
		{
			UpdateProtectedElementIndexFromChanges(doc, data.GetAddedElementIds(), data.GetModifiedElementIds(), data.GetDeletedElementIds());
			return;
		}
		foreach (ProtectedChangeEvent item in events)
		{
			item.RecoveryStatus = "BlockedBeforeCommit";
			item.RequiredAction = "UseFamilyBrowserOrAskAdministrator";
		}
		RememberBlockedNativeAction(events.Any([SpecialName] (ProtectedChangeEvent item) => (item?.Kind ?? string.Empty).StartsWith("NestedOnlyFamilyPlacement", StringComparison.OrdinalIgnoreCase)) ? "Nested-only family standalone placement" : "Protected family/type change");
		RestoreProtectedModifiedNamesBeforeCommit(doc, events);
		PostProtectedNativeChangeFailure(doc, events);
		WriteAudit(doc, currentUser, role, events);
	}

	private static bool ShouldBlockNestedOnlyStandalonePlacement(FamilyBrowserFileGuardTarget target, bool adminModeEnabled)
	{
		return target != null && target.Enabled && target.BlockNestedOnlyStandalonePlacement && !adminModeEnabled;
	}

	private static bool ShouldBlockNestedOnlyPlacementMatch(FamilyBrowserNestedOnlyPlacementMatchResult match)
	{
		return match != null && match.State == FamilyBrowserNestedOnlyPlacementMatchState.ExactMatch;
	}

	private static void ScheduleNestedOnlyFingerprintRefreshIfRequired(Document document, FamilyBrowserStandardPolicy policy)
	{
		if (document == null || document.IsFamilyDocument || policy == null)
		{
			return;
		}
		try
		{
			FamilyBrowserProjectPolicyContext context = BuildProjectContext(document, policy, includeCentralPath: true);
			FamilyBrowserFileGuardTarget target = FindMatchingNativeGuardTarget(policy, context);
			if (target != null && target.Enabled && target.BlockNestedOnlyStandalonePlacement)
			{
				string discipline = FamilyBrowserFileGuardDisciplineService.ResolveAssignedDiscipline(policy, target, allowLegacyFallback: false);
				if (!string.IsNullOrWhiteSpace(discipline))
				{
					FamilyBrowserNestedOnlyPlacementRuntimeService.ScheduleRefresh(document, policy, discipline);
				}
			}
		}
		catch (Exception ex)
		{
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Nested-only fingerprint refresh scheduling failed", ex, "Document=" + SafeDocumentTitle(document));
		}
	}

	private static void CollectNestedOnlyStandalonePlacements(List<ProtectedChangeEvent> events, Document doc, ICollection<ElementId> addedIds, FamilyBrowserStandardPolicy policy, string discipline)
	{
		if (events == null || doc == null || doc.IsFamilyDocument || addedIds == null || addedIds.Count == 0)
		{
			return;
		}
		foreach (ElementId id in addedIds)
		{
			FamilyInstance instance = null;
			try
			{
				instance = doc.GetElement(id) as FamilyInstance;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			if (instance == null)
			{
				continue;
			}
			try
			{
				if (instance.SuperComponent != null)
				{
					continue;
				}
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				ProjectData.ClearProjectError();
				continue;
			}
			Family family = null;
			try
			{
				family = instance.Symbol?.Family;
			}
			catch (Exception projectError3)
			{
				ProjectData.SetProjectError(projectError3);
				ProjectData.ClearProjectError();
			}
			if (family == null)
			{
				continue;
			}
			string familyName = family.Name ?? string.Empty;
			string categoryName = FamilyBrowserFamilyClassificationService.ResolveCategoryName(family);
			string categoryId = FamilyBrowserFamilyClassificationService.ResolveCategoryId(family);
			FamilyBrowserNestedOnlyPlacementMatchResult match = FamilyBrowserNestedOnlyPlacementRuntimeService.EvaluatePlacement(doc, family, policy, discipline);
			if (!ShouldBlockNestedOnlyPlacementMatch(match))
			{
				continue;
			}
			FamilyBrowserNestedOnlyPlacementEntry entry = match.Entry;
			string elementIdText = GetElementIdText(id);
			if (events.Any([SpecialName] (ProtectedChangeEvent item) => item != null && (item.Kind ?? string.Empty).StartsWith("NestedOnlyFamilyPlacement", StringComparison.OrdinalIgnoreCase) && string.Equals(item.ElementIdText, elementIdText, StringComparison.Ordinal)))
			{
				continue;
			}
			events.Add(new ProtectedChangeEvent
			{
				Action = "Added",
				Kind = "NestedOnlyFamilyPlacement",
				Name = familyName,
				CategoryName = categoryName,
				ElementIdText = elementIdText,
				State = "UnauthorizedNestedOnlyStandalonePlacement",
				RecoveryStatus = "BlockedBeforeCommit",
				RequiredAction = "PlaceParentFamilyOrAskAdministrator",
				PolicyReason = "FamilyIsNestedOnlyAndFingerprintMatchesStandard",
				ParentFamilyNames = string.Join(", ", entry?.ParentFamilyNames ?? new List<string>()),
				DetectedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
			});
		}
	}

	private static string ResolveAssignedFileGuardDiscipline(Document document, FamilyBrowserStandardPolicy policy)
	{
		if (document == null || policy == null)
		{
			return string.Empty;
		}
		FamilyBrowserProjectPolicyContext context = BuildProjectContext(document, policy, includeCentralPath: true);
		FamilyBrowserFileGuardTarget target = FindMatchingNativeGuardTarget(policy, context);
		return FamilyBrowserFileGuardDisciplineService.ResolveAssignedDiscipline(policy, target, allowLegacyFallback: false);
	}

	private static FamilyBrowserNestedOnlyPlacementCatalog ResolveNestedOnlyPlacementCatalog(FamilyBrowserStandardPolicy policy)
	{
		string snapshotPath = ResolveNestedOnlyPlacementSnapshotPath(policy);
		if (string.IsNullOrWhiteSpace(snapshotPath))
		{
			return null;
		}
		string cacheKey = FamilyBrowserNestedOnlyPlacementCatalogStore.GetSidecarPath(snapshotPath);
		DateTime now = DateTime.UtcNow;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			CachedNestedOnlyPlacementCatalog cached;
			if (NestedOnlyPlacementCatalogCache.TryGetValue(cacheKey, out cached) && cached != null && (now - cached.CachedUtc).TotalSeconds < NestedOnlyPlacementCatalogCacheSeconds)
			{
				return cached.Catalog;
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		FamilyBrowserNestedOnlyPlacementCatalog catalog = null;
		try
		{
			catalog = FamilyBrowserNestedOnlyPlacementCatalogStore.TryLoadForSnapshot(snapshotPath);
		}
		catch (Exception ex)
		{
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Nested-only placement catalog load failed", ex, "SnapshotPath=" + snapshotPath);
		}
		object syncRoot2 = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot2);
		bool lockTaken2 = false;
		try
		{
			Monitor.Enter(syncRoot2, ref lockTaken2);
			NestedOnlyPlacementCatalogCache[cacheKey] = new CachedNestedOnlyPlacementCatalog
			{
				Catalog = catalog,
				CachedUtc = now
			};
			if (NestedOnlyPlacementCatalogCache.Count > 64)
			{
				NestedOnlyPlacementCatalogCache.Clear();
			}
		}
		finally
		{
			if (lockTaken2)
			{
				Monitor.Exit(syncRoot2);
			}
		}
		return catalog;
	}

	private static string ResolveNestedOnlyPlacementSnapshotPath(FamilyBrowserStandardPolicy policy)
	{
		if (policy == null)
		{
			return string.Empty;
		}
		try
		{
			string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
			FamilyBrowserStandardLibrarySlot slot = FamilyBrowserStandardPolicyStore.GetEffectiveSlot(policy);
			if (slot == null)
			{
				return string.Empty;
			}
			StandardLibraryRegistrationRecord registration = null;
			string registrationPath = FamilyBrowserStandardPolicyStore.ResolveSlotRegistrationPath(workspaceRoot, slot);
			if (!string.IsNullOrWhiteSpace(registrationPath) && File.Exists(registrationPath))
			{
				registration = DataContractJsonFileStore.Load<StandardLibraryRegistrationRecord>(registrationPath);
			}
			string snapshotPath = FamilyBrowserStandardPolicyStore.ResolveSlotSnapshotPath(workspaceRoot, slot, registration);
			if (string.IsNullOrWhiteSpace(snapshotPath) && !string.IsNullOrWhiteSpace(slot.SnapshotPath) && File.Exists(slot.SnapshotPath))
			{
				snapshotPath = slot.SnapshotPath;
			}
			return snapshotPath ?? string.Empty;
		}
		catch (Exception ex)
		{
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Nested-only placement snapshot resolution failed", ex, string.Empty);
			return string.Empty;
		}
	}

	private static void RestoreProtectedModifiedNamesBeforeCommit(Document doc, List<ProtectedChangeEvent> events)
	{
		if (doc == null || events == null || events.Count == 0)
		{
			return;
		}
		foreach (ProtectedChangeEvent item in events)
		{
			if (item == null || !string.Equals(item.Action, "Modified", StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(item.OriginalElementName))
			{
				continue;
			}
			ElementId elementId = ParseElementIdText(item.ElementIdText);
			if ((object)elementId == null)
			{
				continue;
			}
			Element element = null;
			try
			{
				element = doc.GetElement(elementId);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				element = null;
				ProjectData.ClearProjectError();
			}
			if (element == null)
			{
				continue;
			}
			try
			{
				if (TryRestoreElementName(element, item.OriginalElementName))
				{
					item.RecoveryStatus = "RestoredBeforeCommit";
					item.RequiredAction = "BlockedByNativeGuard";
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Protected native change pre-commit restore failed", ex2, "ElementId=" + item.ElementIdText + Environment.NewLine + "Kind=" + item.Kind + Environment.NewLine + "Name=" + item.Name + Environment.NewLine + "OriginalElementName=" + item.OriginalElementName);
				ProjectData.ClearProjectError();
			}
		}
	}

	private static bool TryRestoreElementName(Element element, string originalElementName)
	{
		if (element == null || string.IsNullOrWhiteSpace(originalElementName))
		{
			return false;
		}
		if (element is Family family)
		{
			if (string.Equals(family.Name, originalElementName, StringComparison.Ordinal))
			{
				return false;
			}
			family.Name = originalElementName;
			return true;
		}
		if (element is FamilySymbol symbol)
		{
			if (string.Equals(symbol.Name, originalElementName, StringComparison.Ordinal))
			{
				return false;
			}
			symbol.Name = originalElementName;
			return true;
		}
		if (element is ElementType elementType)
		{
			if (string.Equals(elementType.Name, originalElementName, StringComparison.Ordinal))
			{
				return false;
			}
			elementType.Name = originalElementName;
			return true;
		}
		return false;
	}

	private static void PostProtectedNativeChangeFailure(Document doc, List<ProtectedChangeEvent> events)
	{
		if (doc == null || events == null || events.Count == 0)
		{
			return;
		}
		try
		{
			FailureMessage message = new FailureMessage(ProtectedChangeFailureId);
			List<ElementId> failingIds = (from x in events
				select ParseElementIdText(x.ElementIdText) into x
				where (object)x != null && RevitElementIdCompat.CompatIntegerValue(x) > 0
				select x).ToList();
			if (failingIds.Count == 1)
			{
				message.SetFailingElement(failingIds[0]);
			}
			else if (failingIds.Count > 1)
			{
				message.SetFailingElements(failingIds);
			}
			doc.PostFailure(message);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Protected native change failure post failed", ex2, "Document=" + ((doc == null) ? string.Empty : doc.Title));
			ProjectData.ClearProjectError();
		}
	}

	private static ElementId ParseElementIdText(string value)
	{
		if (int.TryParse(value ?? string.Empty, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) && parsed > 0)
		{
			return new ElementId(parsed);
		}
		return null;
	}

	private static void RefreshProtectedElementIndex(Document doc)
	{
		if (doc == null || doc.IsFamilyDocument)
		{
			return;
		}
		Dictionary<int, ProtectedElementInfo> index = new Dictionary<int, ProtectedElementInfo>();
		foreach (Family family in new FilteredElementCollector(doc).OfClass(typeof(Family)).Cast<Family>())
		{
			ProtectedElementInfo info = BuildProtectedElementInfo(family);
			if (info != null)
			{
				index[GetElementIdInteger(family.Id)] = info;
			}
		}
		foreach (ElementType elementType in new FilteredElementCollector(doc).WhereElementIsElementType().Cast<ElementType>())
		{
			ProtectedElementInfo info2 = BuildProtectedElementInfo(elementType);
			if (info2 != null)
			{
				index[GetElementIdInteger(elementType.Id)] = info2;
			}
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			string documentKey = BuildDocumentKey(doc);
			ProtectedElementIndexes[documentKey] = index;
			CompleteProtectedElementIndexDocumentTokens[documentKey] = RuntimeHelpers.GetHashCode(doc);
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	private static void UpdateProtectedElementIndexFromChanges(Document doc, ICollection<ElementId> addedIds, ICollection<ElementId> modifiedIds, ICollection<ElementId> deletedIds)
	{
		if (doc == null || doc.IsFamilyDocument)
		{
			return;
		}
		string documentKey = BuildDocumentKey(doc);
		if (string.IsNullOrWhiteSpace(documentKey))
		{
			return;
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			Dictionary<int, ProtectedElementInfo> index = null;
			if (!ProtectedElementIndexes.TryGetValue(documentKey, out index) || index == null)
			{
				index = new Dictionary<int, ProtectedElementInfo>();
				ProtectedElementIndexes[documentKey] = index;
			}
			RemoveDeletedIndexEntries(index, deletedIds);
			AddOrUpdateIndexEntries(doc, index, addedIds);
			AddOrUpdateIndexEntries(doc, index, modifiedIds);
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	private static void RemoveDeletedIndexEntries(IDictionary<int, ProtectedElementInfo> index, ICollection<ElementId> ids)
	{
		if (index == null || ids == null)
		{
			return;
		}
		foreach (ElementId id in ids)
		{
			index.Remove(GetElementIdInteger(id));
		}
	}

	private static void AddOrUpdateIndexEntries(Document doc, IDictionary<int, ProtectedElementInfo> index, ICollection<ElementId> ids)
	{
		if (doc == null || index == null || ids == null)
		{
			return;
		}
		foreach (ElementId id in ids)
		{
			int key = GetElementIdInteger(id);
			if (key <= 0)
			{
				continue;
			}
			Element element = null;
			try
			{
				element = doc.GetElement(id);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				element = null;
				ProjectData.ClearProjectError();
			}
			if (element == null)
			{
				index.Remove(key);
				continue;
			}
			ProtectedElementInfo info = BuildProtectedElementInfo(element);
			if (info == null)
			{
				index.Remove(key);
			}
			else
			{
				index[key] = info;
			}
		}
	}

	private static ProtectedElementInfo BuildProtectedElementInfo(Element element)
	{
		if (element == null)
		{
			return null;
		}
		if (element is Family family)
		{
			return new ProtectedElementInfo
			{
				Kind = "Family",
				Name = (family.Name ?? string.Empty),
				ElementName = (family.Name ?? string.Empty),
				CategoryName = ResolveFamilyCategoryName(family)
			};
		}
		if (element is FamilySymbol symbol)
		{
			return new ProtectedElementInfo
			{
				Kind = "Loadable Family Type",
				Name = ResolveFamilySymbolName(symbol),
				ElementName = (symbol.Name ?? string.Empty),
				CategoryName = ResolveElementCategoryName(symbol)
			};
		}
		if (element is ElementType elementType)
		{
			return new ProtectedElementInfo
			{
				Kind = "Element Type",
				Name = (elementType.Name ?? string.Empty),
				ElementName = (elementType.Name ?? string.Empty),
				CategoryName = ResolveElementCategoryName(elementType)
			};
		}
		return null;
	}

	private static bool ProtectedElementInfoEquals(ProtectedElementInfo left, ProtectedElementInfo right)
	{
		if (left == null || right == null)
		{
			return false;
		}
		return string.Equals(left.Kind, right.Kind, StringComparison.OrdinalIgnoreCase) && string.Equals(left.Name, right.Name, StringComparison.Ordinal) && string.Equals(left.ElementName, right.ElementName, StringComparison.Ordinal) && string.Equals(left.CategoryName, right.CategoryName, StringComparison.Ordinal);
	}

	private static string ResolveFamilySymbolName(FamilySymbol symbol)
	{
		string familyName = string.Empty;
		try
		{
			familyName = symbol.FamilyName ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		string typeName = symbol.Name ?? string.Empty;
		if (string.IsNullOrWhiteSpace(familyName))
		{
			return typeName;
		}
		if (string.IsNullOrWhiteSpace(typeName))
		{
			return familyName;
		}
		return familyName + " : " + typeName;
	}

	private static string ResolveFamilyCategoryName(Family family)
	{
		string ResolveFamilyCategoryName;
		try
		{
			ResolveFamilyCategoryName = family.FamilyCategory?.Name ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveFamilyCategoryName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveFamilyCategoryName;
	}

	private static string ResolveElementCategoryName(Element element)
	{
		string ResolveElementCategoryName;
		try
		{
			ResolveElementCategoryName = element.Category?.Name ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveElementCategoryName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveElementCategoryName;
	}

	private static void ResolveNativeGuardPermissions(FamilyBrowserStandardPolicy policy, string currentUser, FamilyBrowserProjectPolicyContext context, out bool canEditFamilies, out bool canAddDeleteTypes)
	{
		string cacheKey = BuildNativeGuardDecisionCacheKey(policy, currentUser, context);
		DateTime now = DateTime.UtcNow;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			CachedNativeGuardDecision cached = null;
			if (NativeGuardDecisionCache.TryGetValue(cacheKey, out cached) && cached != null && (now - cached.CachedUtc).TotalSeconds < NativeGuardDecisionCacheSeconds)
			{
				canEditFamilies = cached.CanEditFamilies;
				canAddDeleteTypes = cached.CanAddDeleteTypes;
				return;
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		canEditFamilies = CanNativeGuardPermission(policy, currentUser, "EditFamilies", context);
		canAddDeleteTypes = CanNativeGuardPermission(policy, currentUser, "AddDeleteTypes", context);
		object syncRoot2 = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot2);
		bool lockTaken2 = false;
		try
		{
			Monitor.Enter(syncRoot2, ref lockTaken2);
			NativeGuardDecisionCache[cacheKey] = new CachedNativeGuardDecision
			{
				CanEditFamilies = canEditFamilies,
				CanAddDeleteTypes = canAddDeleteTypes,
				CachedUtc = now
			};
			if (NativeGuardDecisionCache.Count > 128)
			{
				NativeGuardDecisionCache.Clear();
			}
		}
		finally
		{
			if (lockTaken2)
			{
				Monitor.Exit(syncRoot2);
			}
		}
	}

	private static string BuildNativeGuardDecisionCacheKey(FamilyBrowserStandardPolicy policy, string currentUser, FamilyBrowserProjectPolicyContext context)
	{
		return string.Join("|", new string[8]
		{
			currentUser ?? string.Empty,
			IsAdminModeEnabledForNativeGuard().ToString(),
			BuildPolicyStamp(policy),
			(context == null) ? string.Empty : (context.ModelPath ?? string.Empty),
			(context == null) ? string.Empty : (context.CentralPath ?? string.Empty),
			(context == null) ? string.Empty : (context.ProjectTitle ?? string.Empty),
			(context == null) ? string.Empty : (context.StandardTarget ?? string.Empty),
			HasEnabledFileGuardTargets(policy).ToString()
		});
	}

	private static string BuildPolicyStamp(FamilyBrowserStandardPolicy policy)
	{
		if (policy == null)
		{
			return "(no-policy)";
		}
		FamilyBrowserFileGuardPolicy fileGuard = policy.FileGuard;
		int targetCount = 0;
		int familyBlockCount = 0;
		int typeBlockCount = 0;
		int nestedOnlyBlockCount = 0;
		if (fileGuard != null && fileGuard.Targets != null)
		{
			foreach (FamilyBrowserFileGuardTarget target in fileGuard.Targets)
			{
				if (target != null && target.Enabled)
				{
					targetCount = checked(targetCount + 1);
					familyBlockCount = checked(familyBlockCount + (target.BlockFamilyLoadAndEdit ? 1 : 0));
					typeBlockCount = checked(typeBlockCount + (target.BlockTypeChanges ? 1 : 0));
					nestedOnlyBlockCount = checked(nestedOnlyBlockCount + (target.BlockNestedOnlyStandalonePlacement ? 1 : 0));
				}
			}
		}
		return string.Join("|", new string[8]
		{
			policy.LastUpdatedUtc ?? string.Empty,
			(fileGuard == null) ? string.Empty : (fileGuard.LastUpdatedUtc ?? string.Empty),
			(fileGuard != null && fileGuard.Enabled).ToString(),
			targetCount.ToString(CultureInfo.InvariantCulture),
			(fileGuard == null) ? string.Empty : (fileGuard.RootFolder ?? string.Empty),
			familyBlockCount.ToString(CultureInfo.InvariantCulture),
			typeBlockCount.ToString(CultureInfo.InvariantCulture),
			nestedOnlyBlockCount.ToString(CultureInfo.InvariantCulture)
		});
	}

	private static bool HasEnabledFileGuardTargets(FamilyBrowserStandardPolicy policy)
	{
		FamilyBrowserFileGuardPolicy fileGuard = policy?.FileGuard;
		if (fileGuard == null || !fileGuard.Enabled || fileGuard.Targets == null)
		{
			return false;
		}
		foreach (FamilyBrowserFileGuardTarget target in fileGuard.Targets)
		{
			if (target != null && target.Enabled)
			{
				return true;
			}
		}
		return false;
	}

	private static bool IsFileGuardTargetedDocument(FamilyBrowserStandardPolicy policy, FamilyBrowserProjectPolicyContext context)
	{
		return FindMatchingNativeGuardTarget(policy, context) != null;
	}

	private static FamilyBrowserFileGuardTarget FindMatchingNativeGuardTarget(FamilyBrowserStandardPolicy policy, FamilyBrowserProjectPolicyContext context)
	{
		return FamilyBrowserFileGuardPathMatcher.FindMatchingTarget(policy?.FileGuard, context);
	}

	private static List<string> BuildNativeGuardContextPathCandidates(FamilyBrowserProjectPolicyContext context)
	{
		List<string> result = new List<string>();
		AddNativeGuardPathCandidate(result, context?.CentralPath);
		AddNativeGuardPathCandidate(result, context?.ModelPath);
		return result;
	}

	private static List<string> BuildNativeGuardTargetPathCandidates(FamilyBrowserFileGuardPolicy fileGuard, FamilyBrowserFileGuardTarget target)
	{
		List<string> result = new List<string>();
		AddNativeGuardPathCandidate(result, target?.CentralPath);
		if (fileGuard != null && target != null && !string.IsNullOrWhiteSpace(fileGuard.RootFolder) && !string.IsNullOrWhiteSpace(target.RelativePath))
		{
			try
			{
				AddNativeGuardPathCandidate(result, Path.Combine(fileGuard.RootFolder, target.RelativePath));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		return result;
	}

	private static HashSet<string> BuildNativeGuardContextNameCandidates(FamilyBrowserProjectPolicyContext context)
	{
		HashSet<string> result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		AddNativeGuardNameCandidate(result, context?.CentralPath);
		AddNativeGuardNameCandidate(result, context?.ModelPath);
		AddNativeGuardNameCandidate(result, context?.ProjectTitle);
		return result;
	}

	private static List<string> BuildNativeGuardTargetNameCandidates(FamilyBrowserFileGuardTarget target)
	{
		List<string> result = new List<string>();
		if (target == null)
		{
			return result;
		}
		result.Add(target.FileName ?? string.Empty);
		result.Add(target.CentralPath ?? string.Empty);
		result.Add(target.RelativePath ?? string.Empty);
		return result;
	}

	private static void AddNativeGuardPathCandidate(List<string> values, string value)
	{
		if (values != null && !string.IsNullOrWhiteSpace(value))
		{
			values.Add(value.Trim());
		}
	}

	private static void AddNativeGuardNameCandidate(HashSet<string> values, string value)
	{
		if (values == null)
		{
			return;
		}
		string normalized = NormalizeNativeGuardDetachedFileBase(value);
		if (!string.IsNullOrWhiteSpace(normalized))
		{
			values.Add(normalized);
		}
	}

	private static string NormalizeNativeGuardDetachedFileBase(string value)
	{
		string text = (value ?? string.Empty).Trim();
		if (text.Length == 0)
		{
			return string.Empty;
		}
		try
		{
			text = Path.GetFileNameWithoutExtension(text);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			if (text.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
			{
				text = text.Substring(0, text.Length - 4);
			}
			ProjectData.ClearProjectError();
		}
		text = text.Trim();
		string[] suffixes = new string[12]
		{
			"_detached", "-detached", ".detached", " detached", "(detached)", " - detached", " _ detached", "_detached copy", "_detached_copy", "-detached copy",
			"-detached-copy", " detached copy"
		};
		bool changed = true;
		while (changed && text.Length > 0)
		{
			changed = false;
			foreach (string suffix in suffixes)
			{
				if (text.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
				{
					text = text.Substring(0, text.Length - suffix.Length).Trim();
					changed = true;
					break;
				}
			}
		}
		return text.ToLowerInvariant();
	}

	private static string SafeDocumentTitle(Document doc)
	{
		if (doc == null)
		{
			return "(no-document)";
		}
		try
		{
			return doc.Title ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return string.Empty;
	}

	private static void WriteGuardTiming(string label, Stopwatch stopwatch, string detail = "")
	{
		if (stopwatch == null)
		{
			return;
		}
		long elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
		if (elapsedMilliseconds < SlowGuardTimingThresholdMilliseconds)
		{
			return;
		}
		try
		{
			FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "NativeGuardTiming-" + (label ?? "unknown"), new TimeoutException("Family Browser native guard path took " + elapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + " ms."), "ElapsedMilliseconds=" + elapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + Environment.NewLine + (detail ?? string.Empty));
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static FamilyBrowserStandardPolicy LoadPolicy()
	{
		Stopwatch sw = Stopwatch.StartNew();
		string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		DateTime now = DateTime.UtcNow;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (CachedPolicy != null && string.Equals(CachedPolicyUser, currentUser, StringComparison.OrdinalIgnoreCase) && (now - CachedPolicyLoadedUtc).TotalSeconds < PolicyCacheSeconds)
			{
				return CachedPolicy;
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		FamilyBrowserStandardPolicy policy = FamilyBrowserStandardPolicyStore.LoadOrCreate(HostWorkspacePathResolver.ResolveRoot(), currentUser);
		string policyStamp = BuildPolicyStamp(policy);
		object syncRoot2 = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot2);
		bool lockTaken2 = false;
		try
		{
			Monitor.Enter(syncRoot2, ref lockTaken2);
			CachedPolicy = policy;
			CachedPolicyUser = currentUser;
			CachedPolicyLoadedUtc = now;
			CachedPolicyStamp = policyStamp;
			NativeGuardDecisionCache.Clear();
			NestedOnlyPlacementCatalogCache.Clear();
		}
		finally
		{
			if (lockTaken2)
			{
				Monitor.Exit(syncRoot2);
			}
		}
		WriteGuardTiming("LoadPolicy", sw, "User=" + currentUser + Environment.NewLine + "PolicyStamp=" + policyStamp);
		return policy;
	}

	private static void SeedPolicyCache(FamilyBrowserStandardPolicy policy)
	{
		if (policy == null)
		{
			return;
		}
		string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			CachedPolicy = policy;
			CachedPolicyUser = currentUser;
			CachedPolicyLoadedUtc = DateTime.UtcNow;
			CachedPolicyStamp = BuildPolicyStamp(policy);
			NativeGuardDecisionCache.Clear();
			NestedOnlyPlacementCatalogCache.Clear();
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
	}

	private static FamilyBrowserProjectPolicyContext BuildProjectContext(Document doc, FamilyBrowserStandardPolicy policy, bool includeCentralPath = true)
	{
		FamilyBrowserProjectPolicyContext context = new FamilyBrowserProjectPolicyContext();
		if (doc != null)
		{
			context.ProjectTitle = doc.Title ?? string.Empty;
			context.ModelPath = doc.PathName ?? string.Empty;
			try
			{
				context.IsWorkshared = doc.IsWorkshared;
				if (doc.IsWorkshared)
				{
					context.CentralPath = ResolveCentralPath(doc, includeCentralPath);
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		try
		{
			FamilyBrowserStandardLibrarySlot slot = FamilyBrowserStandardPolicyStore.GetEffectiveSlot(policy);
			context.StandardTarget = ((slot == null) ? string.Empty : slot.Discipline);
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		return context;
	}

	private static string ResolveCentralPath(Document doc, bool allowResolve)
	{
		if (doc == null)
		{
			return string.Empty;
		}
		string documentKey = BuildDocumentKey(doc);
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (CentralPathCache.TryGetValue(documentKey, out string cachedPath))
			{
				return cachedPath ?? string.Empty;
			}
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		if (!allowResolve)
		{
			return string.Empty;
		}
		string resolvedPath = string.Empty;
		try
		{
			if (doc.IsWorkshared)
			{
				ModelPath centralPath = doc.GetWorksharingCentralModelPath();
				if (centralPath != null)
				{
					resolvedPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(centralPath) ?? string.Empty;
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		if (!string.IsNullOrWhiteSpace(resolvedPath))
		{
			object syncRoot2 = SyncRoot;
			ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot2);
			bool lockTaken2 = false;
			try
			{
				Monitor.Enter(syncRoot2, ref lockTaken2);
				CentralPathCache[documentKey] = resolvedPath;
			}
			finally
			{
				if (lockTaken2)
				{
					Monitor.Exit(syncRoot2);
				}
			}
		}
		return resolvedPath;
	}

	private static void SyncCurrentUserIdentity(Document doc)
	{
		try
		{
			if (doc != null && doc.Application != null)
			{
				string userName = (doc.Application.Username ?? string.Empty).Trim();
				if (!string.IsNullOrWhiteSpace(userName))
				{
					FamilyBrowserSecurityPolicyService.SetCurrentUserIdentityOverride(userName);
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static string BuildDocumentKey(Document doc)
	{
		if (doc == null)
		{
			return "(no-document)";
		}
		string path = (doc.PathName ?? string.Empty).Trim();
		if (!string.IsNullOrWhiteSpace(path))
		{
			return path;
		}
		return (doc.Title ?? string.Empty) + "|" + doc.GetHashCode().ToString(CultureInfo.InvariantCulture);
	}

	private static int GetElementIdInteger(ElementId id)
	{
		if ((object)id == null)
		{
			return -1;
		}
		return RevitElementIdCompat.CompatIntegerValue(id);
	}

	private static string GetElementIdText(ElementId id)
	{
		if ((object)id == null)
		{
			return string.Empty;
		}
		return RevitElementIdCompat.CompatIntegerValue(id).ToString(CultureInfo.InvariantCulture);
	}

	private static bool SameText(string leftValue, string rightValue)
	{
		return string.Equals(leftValue ?? string.Empty, rightValue ?? string.Empty, StringComparison.Ordinal);
	}

	private static void WriteBlockedCommandAudit(Document doc, string currentUser, string role, ProtectedNativeCommandDefinition definition, string extraDetail = null)
	{
		string detail = "BlockedCommand=" + definition.Key + Environment.NewLine + "DisplayName=" + NativeCommandDisplayName(definition) + Environment.NewLine + "RequiredPermission=" + definition.RequiredPermission + Environment.NewLine + "User=" + (currentUser ?? string.Empty) + Environment.NewLine + "Role=" + (role ?? string.Empty) + Environment.NewLine + "Document=" + ((doc == null) ? string.Empty : doc.Title);
		if (!string.IsNullOrWhiteSpace(extraDetail))
		{
			detail = detail + Environment.NewLine + extraDetail;
		}
		FamilyBrowserErrorHelp.WriteLog(HostWorkspacePathResolver.ResolveRoot(), "Native command blocked", new UnauthorizedAccessException("Family Browser policy blocked a native Revit command."), detail);
	}

	private static void WriteAudit(Document doc, string currentUser, string role, List<ProtectedChangeEvent> events)
	{
		try
		{
			string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
			if (!FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot))
			{
				return;
			}
			string text = Path.Combine(ProjectSnapshotStore.GetProjectHistoryFolder(workspaceRoot, doc), "SecurityAudit");
			Directory.CreateDirectory(text);
			string fileName = DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + "-native-change.log";
			string auditPath = Path.Combine(text, fileName);
			List<string> lines = new List<string>
			{
				"Timestamp: " + DateTime.Now.ToString("O", CultureInfo.InvariantCulture),
				"User: " + (currentUser ?? string.Empty),
				"Role: " + (role ?? string.Empty),
				"Document: " + ((doc == null) ? string.Empty : doc.Title),
				"DocumentPath: " + ((doc == null) ? string.Empty : doc.PathName),
				"Message: Unauthorized or unapproved protected family/type change was detected after a Revit document change.",
				"Result: Native guard records the attempt; IUpdater blocks protected edits before commit when available.",
				"DeletedResult: Protected deletions are blocked before commit when the updater trigger is available.",
				"DeletedRequiredAction: UseFamilyBrowserOrAskAdministrator",
				string.Empty
			};
			foreach (ProtectedChangeEvent item in events)
			{
				lines.Add(item.Action + " | " + item.Kind + " | " + item.Name + " | " + item.CategoryName + " | " + item.ElementIdText + " | " + item.State + " | " + item.RecoveryStatus + " | " + item.RequiredAction + " | " + item.DetectedAtUtc);
			}
			File.WriteAllLines(auditPath, lines);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void MarkDeletedProtectedContentDirty(Document doc, string currentUser, List<ProtectedChangeEvent> events)
	{
		if (doc == null || events == null)
		{
			return;
		}
		List<ProtectedChangeEvent> deletedEvents = events.Where([SpecialName] (ProtectedChangeEvent x) => string.Equals(x.Action, "Deleted", StringComparison.OrdinalIgnoreCase)).ToList();
		if (deletedEvents.Count != 0)
		{
			List<ProjectTrackingDirtyItem> dirtyItems = deletedEvents.Select([SpecialName] (ProtectedChangeEvent x) => new ProjectTrackingDirtyItem
			{
				Action = x.Action,
				Kind = x.Kind,
				Name = x.Name,
				CategoryName = x.CategoryName,
				ElementIdText = x.ElementIdText,
				State = x.State,
				RecoveryStatus = x.RecoveryStatus,
				RequiredAction = x.RequiredAction
			}).ToList();
			ProjectTrackingStoreService.MarkCurrentModelCheckRequired(doc, currentUser, "UnauthorizedDeletedProtectedContent", dirtyItems);
		}
	}

	private static void ShowUnauthorizedChangeWarning(string currentUser, List<ProtectedChangeEvent> events)
	{
	}

	private static void ShowRevertedChangeWarning(object items)
	{
	}
}
