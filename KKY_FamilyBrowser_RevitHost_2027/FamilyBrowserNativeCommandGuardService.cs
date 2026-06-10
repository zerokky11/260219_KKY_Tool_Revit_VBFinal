using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserNativeCommandGuardService
{
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
			//IL_000d: Unknown result type (might be due to invalid IL or missing references)
			//IL_0017: Expected O, but got Unknown
			UpdaterIdValue = new UpdaterId(addInId, ProtectedChangeUpdaterGuid);
		}

		public void Execute(UpdaterData data)
		{
			HandleProtectedContentUpdaterExecute(data);
		}

		public UpdaterId GetUpdaterId()
		{
			return UpdaterIdValue;
		}

		public string GetUpdaterName()
		{
			return "KKY Family Browser protected family/type guard";
		}

		public string GetAdditionalInformation()
		{
			return "Blocks protected family and type changes before transaction commit for non-admin users.";
		}

		public ChangePriority GetChangePriority()
		{
			return (ChangePriority)9;
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
			DetectedAtUtc = string.Empty;
		}
	}

	private static readonly object SyncRoot = RuntimeHelpers.GetObjectValue(new object());

	private static readonly List<BoundNativeCommand> BoundCommands = new List<BoundNativeCommand>();

	private static readonly Dictionary<string, Dictionary<int, ProtectedElementInfo>> ProtectedElementIndexes = new Dictionary<string, Dictionary<int, ProtectedElementInfo>>(StringComparer.OrdinalIgnoreCase);

	private static UIControlledApplication ControlledApplication;

	private static Document LastActiveDocument;

	private static DateTime LastWarningUtc = DateTime.MinValue;

	private static DateTime LastBlockedNativeActionUtc = DateTime.MinValue;

	private static string LastBlockedNativeActionName = string.Empty;

	private static bool GuardInitialized = false;

	private static int TrustedOperationDepth = 0;

	private static DateTime TrustedOperationUntilUtc = DateTime.MinValue;

	private const double PolicyCacheSeconds = 10.0;

	private const double BlockedUndoCooldownSeconds = 10.0;

	private static FamilyBrowserStandardPolicy CachedPolicy;

	private static string CachedPolicyUser = string.Empty;

	private static DateTime CachedPolicyLoadedUtc = DateTime.MinValue;

	private static bool DatabaseEventsAttached = false;

	private static ProtectedContentChangeUpdater ProtectedChangeUpdaterInstance;

	private static DateTime LastRibbonAvailabilityUpdateUtc = DateTime.MinValue;

	private static bool LastRibbonLoadFamilyAllowed = true;

	private static bool LastRibbonLoadFamilyKnown = false;

	private static bool AdminModeStateKnown = false;

	private static bool AdminModeEnabledForNativeGuard = false;

	private static readonly Guid ProtectedChangeUpdaterGuid = new Guid("49DDE62E-26D7-4A79-9D6D-CED0A197A23D");

	private static readonly FailureDefinitionId ProtectedChangeFailureId = new FailureDefinitionId(new Guid("11DB20D4-2631-43B3-B9D4-601E24C70C67"));

	private static bool ProtectedChangeFailureRegistered = false;

	private static readonly BuiltInParameter[] ProtectedNameChangeParameters;

	private static readonly ProtectedNativeCommandDefinition FamilyLoadingEventDefinition;

	private static readonly ProtectedNativeCommandDefinition FamilyDocumentSaveEventDefinition;

	private static readonly ProtectedNativeCommandDefinition FamilyDocumentSaveAsEventDefinition;

	private static readonly HashSet<string> SystemTypeNames;

	static FamilyBrowserNativeCommandGuardService()
	{
		//IL_00b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bb: Expected O, but got Unknown
		BuiltInParameter[] array = new BuiltInParameter[4];
		RuntimeHelpers.InitializeArray(array, (RuntimeFieldHandle)/*OpCode not supported: LdMemberToken*/);
		ProtectedNameChangeParameters = (BuiltInParameter[])(object)array;
		FamilyLoadingEventDefinition = new ProtectedNativeCommandDefinition
		{
			Key = "native-family-loading-event",
			DisplayNameEn = "Revit family load",
			DisplayNameKo = "Revit 패밀리 로드",
			RequiredPermission = "EditFamilies"
		};
		FamilyDocumentSaveEventDefinition = new ProtectedNativeCommandDefinition
		{
			Key = "native-family-document-save-event",
			DisplayNameEn = "Family document save",
			DisplayNameKo = "패밀리 문서 저장",
			RequiredPermission = "EditFamilies",
			FamilyDocumentOnly = true
		};
		FamilyDocumentSaveAsEventDefinition = new ProtectedNativeCommandDefinition
		{
			Key = "native-family-document-save-as-event",
			DisplayNameEn = "Family document save as",
			DisplayNameKo = "패밀리 문서 다른 이름으로 저장",
			RequiredPermission = "EditFamilies",
			FamilyDocumentOnly = true
		};
		SystemTypeNames = new HashSet<string>(new string[20]
		{
			"WallType", "FloorType", "RoofType", "CeilingType", "StairsType", "RailingType", "DuctType", "PipeType", "FlexDuctType", "FlexPipeType",
			"DuctSystemType", "PipingSystemType", "MechanicalSystemType", "ElectricalSystemType", "CableTrayType", "ConduitType", "WireType", "DuctInsulationType", "PipeInsulationType", "DuctLiningType"
		}, StringComparer.OrdinalIgnoreCase);
	}

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
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (AdminModeStateKnown)
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
		bool LoadAdminModeEnabledSetting;
		try
		{
			LoadAdminModeEnabledSetting = FamilyBrowserUserSettingsStore.LoadAdminModeEnabled();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			LoadAdminModeEnabledSetting = false;
			ProjectData.ClearProjectError();
		}
		return LoadAdminModeEnabledSetting;
	}

	private static bool CanNativeGuardPermission(FamilyBrowserStandardPolicy policy, string currentUser, string permission, FamilyBrowserProjectPolicyContext context)
	{
		return FamilyBrowserSecurityPolicyService.CanNativeGuard(policy, currentUser, permission, context, IsAdminModeEnabledForNativeGuard());
	}

	private static string ResolveNativeGuardRoleLabel(FamilyBrowserStandardPolicy policy, string currentUser, FamilyBrowserProjectPolicyContext context)
	{
		string role = FamilyBrowserSecurityPolicyService.ResolveRole(policy, currentUser, context);
		if (string.Equals(role, "Admin", StringComparison.OrdinalIgnoreCase) && !IsAdminModeEnabledForNativeGuard())
		{
			return UiText("Admin profile (Admin Mode Off)", "관리자 프로필(관리자 모드 OFF)");
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
			AttachDatabasePreEvents(application);
			RegisterProtectedChangeUpdater(application);
			ControlledApplication = application;
			GuardInitialized = true;
			UpdateProtectedRibbonAvailability(force: true);
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
			DetachDatabasePreEvents();
			UnregisterProtectedChangeUpdater();
			BoundCommands.Clear();
			ProtectedElementIndexes.Clear();
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
			DatabaseEventsAttached = false;
			ProtectedChangeUpdaterInstance = null;
			LastRibbonAvailabilityUpdateUtc = DateTime.MinValue;
			LastRibbonLoadFamilyAllowed = true;
			LastRibbonLoadFamilyKnown = false;
			AdminModeStateKnown = false;
			AdminModeEnabledForNativeGuard = false;
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
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			CachedPolicy = null;
			CachedPolicyUser = string.Empty;
			CachedPolicyLoadedUtc = DateTime.MinValue;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		UpdateProtectedRibbonAvailability(force: true);
	}

	public static void NotifyAdminModeChanged(bool enabled)
	{
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			AdminModeEnabledForNativeGuard = enabled;
			AdminModeStateKnown = true;
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
		if (!enabled && LastActiveDocument != null)
		{
			try
			{
				EnsureProtectedElementIndexForGuard(LastActiveDocument);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		UpdateProtectedRibbonAvailability(force: true);
	}

	public static void NotifyActiveDocumentChanged(Document document)
	{
		LastActiveDocument = document;
		if (document != null)
		{
			SyncCurrentUserIdentity(document);
			EnsureProtectedElementIndexForGuard(document);
			UpdateProtectedRibbonAvailability();
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
		if (IsTrustedOperationActive())
		{
			UpdateProtectedElementIndexFromChanges(doc, e.GetAddedElementIds(), e.GetModifiedElementIds(), e.GetDeletedElementIds());
			return;
		}
		FamilyBrowserStandardPolicy policy = LoadPolicy();
		FamilyBrowserProjectPolicyContext context = BuildProjectContext(doc, policy);
		string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		string role = ResolveNativeGuardRoleLabel(policy, currentUser, context);
		bool canEditFamilies = CanNativeGuardPermission(policy, currentUser, "EditFamilies", context);
		bool canAddDeleteTypes = CanNativeGuardPermission(policy, currentUser, "AddDeleteTypes", context);
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
		if (previousIndex == null)
		{
			RefreshProtectedElementIndex(doc);
			return;
		}
		List<ProtectedChangeEvent> events = new List<ProtectedChangeEvent>();
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
				e.SetProcessingResult((FailureProcessingResult)2);
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
		if (failureId == null)
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
		if (failureId == null)
		{
			return string.Empty;
		}
		string[] array = new string[2] { "Guid", "GUID" };
		foreach (string propertyName in array)
		{
			try
			{
				PropertyInfo propertyInfo = ((object)failureId).GetType().GetProperty(propertyName);
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
		return ((object)failureId).ToString();
	}

	private static void HandleFamilyLoadingIntoDocument(object sender, FamilyLoadingIntoDocumentEventArgs e)
	{
		//IL_0107: Unknown result type (might be due to invalid IL or missing references)
		if (e == null || IsTrustedOperationActive())
		{
			return;
		}
		Document doc = ((RevitAPIPreDocEventArgs)e).Document;
		if (doc != null)
		{
			LastActiveDocument = doc;
			SyncCurrentUserIdentity(doc);
		}
		NativeCommandPermissionDecision decision = EvaluateNativeCommandPermission(FamilyLoadingEventDefinition, doc);
		if (!decision.Allowed)
		{
			RememberBlockedNativeAction(FamilyLoadingEventDefinition);
			if (TryCancelPreEvent(((RevitAPIEventArgs)e).Cancellable, [SpecialName] () =>
			{
				((RevitAPIPreEventArgs)e).Cancel();
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
		HandleFamilyDocumentSaveEvent((RevitAPIPreDocEventArgs)(object)e, FamilyDocumentSaveEventDefinition);
	}

	private static void HandleDocumentSavingAs(object sender, DocumentSavingAsEventArgs e)
	{
		HandleFamilyDocumentSaveEvent((RevitAPIPreDocEventArgs)(object)e, FamilyDocumentSaveAsEventDefinition);
	}

	private static void HandleFamilyDocumentSaveEvent(RevitAPIPreDocEventArgs e, ProtectedNativeCommandDefinition definition)
	{
		//IL_00b5: Unknown result type (might be due to invalid IL or missing references)
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
			if (TryCancelPreEvent(((RevitAPIEventArgs)e).Cancellable, [SpecialName] () =>
			{
				((RevitAPIPreEventArgs)e).Cancel();
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
		return UiText("The previous Revit command was already blocked before it was committed.", "직전 Revit 명령은 커밋되기 전에 이미 차단되었습니다.") + "\r\n\r\n" + UiText("Action: ", "작업: ") + actionName + "\r\n" + UiText("There is no automatic restore transaction to undo. Continue working through Family Browser.", "되돌릴 자동 복구 트랜잭션은 없습니다. Family Browser를 통해 작업을 계속 진행하세요.");
	}

	private static void RegisterProtectedChangeUpdater(UIControlledApplication application)
	{
		if (application == null || application.ActiveAddInId == null)
		{
			return;
		}
		try
		{
			EnsureProtectedChangeFailureDefinition();
			ProtectedContentChangeUpdater updater = new ProtectedContentChangeUpdater(application.ActiveAddInId);
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
			UpdaterRegistry.RegisterUpdater((IUpdater)(object)updater);
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

	private static void EnsureProtectedChangeFailureDefinition()
	{
		if (!ProtectedChangeFailureRegistered)
		{
			try
			{
				FailureDefinition.CreateFailureDefinition(ProtectedChangeFailureId, (FailureSeverity)2, UiText("KKY Family Browser blocked a protected family or type change. Use Family Browser or ask an administrator.", "KKY Family Browser가 보호된 패밀리 또는 타입 변경을 차단했습니다. Family Browser를 사용하거나 관리자에게 문의하세요."));
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
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		//IL_002d: Expected O, but got Unknown
		//IL_0038: Unknown result type (might be due to invalid IL or missing references)
		//IL_0042: Expected O, but got Unknown
		if (updaterId != null)
		{
			RegisterProtectedChangeTriggers(updaterId, (ElementFilter)new ElementClassFilter(typeof(Family)));
			RegisterProtectedChangeTriggers(updaterId, (ElementFilter)new ElementClassFilter(typeof(FamilySymbol)));
			RegisterProtectedChangeTriggers(updaterId, (ElementFilter)new ElementClassFilter(typeof(ElementType)));
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
		//IL_002a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0034: Expected O, but got Unknown
		List<ElementId> ids = new List<ElementId>();
		HashSet<int> seen = new HashSet<int>();
		BuiltInParameter[] protectedNameChangeParameters = ProtectedNameChangeParameters;
		checked
		{
			for (int i = 0; i < protectedNameChangeParameters.Length; i++)
			{
				int value = (int)unchecked((long)protectedNameChangeParameters[i]);
				if (seen.Add(value))
				{
					ids.Add(new ElementId(unchecked((long)value)));
				}
			}
			return ids;
		}
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
						Definition = definition
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
		//IL_003b: Unknown result type (might be due to invalid IL or missing references)
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
				if (Enum.TryParse<PostableCommand>(commandName, true, out PostableCommand parsed))
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
			ProtectedNativeCommandDefinition definition = FindBoundDefinition("native-load-family");
			if (definition != null)
			{
				bool allowed = EvaluateNativeCommandPermission(definition, LastActiveDocument).Allowed;
				if (force || !LastRibbonLoadFamilyKnown || allowed != LastRibbonLoadFamilyAllowed)
				{
					LastRibbonLoadFamilyAllowed = allowed;
					LastRibbonLoadFamilyKnown = true;
					ApplyLoadFamilyRibbonEnabled(allowed);
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

	private static void ApplyLoadFamilyRibbonEnabled(bool enabled)
	{
		try
		{
			object ribbon = RuntimeHelpers.GetObjectValue(ResolveAutodeskWindowsRibbon());
			if (ribbon != null)
			{
				HashSet<int> visited = new HashSet<int>();
				ApplyLoadFamilyRibbonEnabledRecursive(RuntimeHelpers.GetObjectValue(ribbon), enabled, visited, 0);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
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

	private static void ApplyLoadFamilyRibbonEnabledRecursive(object target, bool enabled, HashSet<int> visited, int depth)
	{
		if (target == null || depth > 10)
		{
			return;
		}
		int key = RuntimeHelpers.GetHashCode(RuntimeHelpers.GetObjectValue(target));
		if (visited.Contains(key))
		{
			return;
		}
		visited.Add(key);
		if (IsLoadFamilyRibbonControl(RuntimeHelpers.GetObjectValue(target)))
		{
			SetBooleanProperty(RuntimeHelpers.GetObjectValue(target), "IsEnabled", enabled);
			SetBooleanProperty(RuntimeHelpers.GetObjectValue(target), "Enabled", enabled);
		}
		foreach (object item in EnumerateRibbonChildren(RuntimeHelpers.GetObjectValue(target)))
		{
			ApplyLoadFamilyRibbonEnabledRecursive(RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(item)), enabled, visited, checked(depth + 1));
		}
	}

	private static List<object> EnumerateRibbonChildren(object target)
	{
		List<object> result = new List<object>();
		if (target == null)
		{
			return result;
		}
		string[] array = new string[10] { "Tabs", "Panels", "Items", "LargeItems", "MediumItems", "SmallItems", "Children", "Controls", "Source", "Content" };
		foreach (string propertyName in array)
		{
			object value = RuntimeHelpers.GetObjectValue(GetPropertyValue(RuntimeHelpers.GetObjectValue(target), propertyName));
			if (value == null || value is string)
			{
				continue;
			}
			if (!(value is IEnumerable enumerable))
			{
				result.Add(RuntimeHelpers.GetObjectValue(value));
				continue;
			}
			try
			{
				foreach (object item2 in enumerable)
				{
					object item = RuntimeHelpers.GetObjectValue(item2);
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
		return new string(new char[6] { '패', '밀', '리', '리', '로', '드' });
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
		AddInCommandBinding binding = (AddInCommandBinding)((sender is AddInCommandBinding) ? sender : null);
		if (binding == null)
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
			NativeCommandPermissionDecision decision = EvaluateNativeCommandPermission(definition);
			e.CanExecute = decision.Allowed;
			if (string.Equals(definition.Key, "native-load-family", StringComparison.OrdinalIgnoreCase))
			{
				UpdateProtectedRibbonAvailability();
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void HandleBeforeExecuted(object sender, BeforeExecutedEventArgs e)
	{
		//IL_0043: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c2: Unknown result type (might be due to invalid IL or missing references)
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
					((RevitEventArgs)e).Cancel = true;
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
				((RevitEventArgs)e).Cancel = true;
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
		FamilyBrowserStandardPolicy policy = LoadPolicy();
		FamilyBrowserProjectPolicyContext context = BuildProjectContext(doc, policy);
		string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		string role = ResolveNativeGuardRoleLabel(policy, currentUser, context);
		bool allowed = CanNativeGuardPermission(policy, currentUser, definition.RequiredPermission, context);
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
				if (!string.Equals(action, "Modified", StringComparison.OrdinalIgnoreCase) || (previousInfo != null && !ProtectedElementInfoEquals(previousInfo, info)))
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

	private static void EnsureProtectedElementIndexForGuard(Document doc)
	{
		if (doc == null || doc.IsFamilyDocument)
		{
			return;
		}
		try
		{
			FamilyBrowserStandardPolicy policy = LoadPolicy();
			FamilyBrowserProjectPolicyContext context = BuildProjectContext(doc, policy);
			string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
			bool num = CanNativeGuardPermission(policy, currentUser, "EditFamilies", context);
			bool canAddDeleteTypes = CanNativeGuardPermission(policy, currentUser, "AddDeleteTypes", context);
			if (num && canAddDeleteTypes)
			{
				return;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		string documentKey = BuildDocumentKey(doc);
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (ProtectedElementIndexes.ContainsKey(documentKey))
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
		RefreshProtectedElementIndex(doc);
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
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (ProtectedElementIndexes.ContainsKey(documentKey))
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
		RefreshProtectedElementIndex(doc);
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
		FamilyBrowserStandardPolicy policy = LoadPolicy();
		FamilyBrowserProjectPolicyContext context = BuildProjectContext(doc, policy);
		string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		string role = ResolveNativeGuardRoleLabel(policy, currentUser, context);
		bool canEditFamilies = CanNativeGuardPermission(policy, currentUser, "EditFamilies", context);
		bool canAddDeleteTypes = CanNativeGuardPermission(policy, currentUser, "AddDeleteTypes", context);
		if (canEditFamilies && canAddDeleteTypes)
		{
			UpdateProtectedElementIndexFromChanges(doc, data.GetAddedElementIds(), data.GetModifiedElementIds(), data.GetDeletedElementIds());
			return;
		}
		Dictionary<int, ProtectedElementInfo> previousIndex = GetProtectedElementIndexSnapshot(doc);
		if (previousIndex == null)
		{
			RefreshProtectedElementIndex(doc);
			return;
		}
		List<ProtectedChangeEvent> events = new List<ProtectedChangeEvent>();
		CollectAddedOrModified(events, doc, previousIndex, data.GetAddedElementIds(), "Added", canEditFamilies, canAddDeleteTypes);
		CollectAddedOrModified(events, doc, previousIndex, data.GetModifiedElementIds(), "Modified", canEditFamilies, canAddDeleteTypes);
		CollectDeleted(events, previousIndex, data.GetDeletedElementIds(), canEditFamilies, canAddDeleteTypes);
		if (events.Count == 0 && data.GetDeletedElementIds() != null && data.GetDeletedElementIds().Count > 0 && previousIndex == null)
		{
			CollectDeletedFallback(events, data.GetDeletedElementIds(), canEditFamilies, canAddDeleteTypes);
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
		RememberBlockedNativeAction("Protected family/type change");
		RestoreProtectedModifiedNamesBeforeCommit(doc, events);
		PostProtectedNativeChangeFailure(doc, events);
		WriteAudit(doc, currentUser, role, events);
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
			if (elementId == null)
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
		Family family = (Family)(object)((element is Family) ? element : null);
		if (family != null)
		{
			if (string.Equals(((Element)family).Name, originalElementName, StringComparison.Ordinal))
			{
				return false;
			}
			((Element)family).Name = originalElementName;
			return true;
		}
		FamilySymbol symbol = (FamilySymbol)(object)((element is FamilySymbol) ? element : null);
		if (symbol != null)
		{
			if (string.Equals(((Element)symbol).Name, originalElementName, StringComparison.Ordinal))
			{
				return false;
			}
			((ElementType)symbol).Name = originalElementName;
			return true;
		}
		ElementType elementType = (ElementType)(object)((element is ElementType) ? element : null);
		if (elementType != null)
		{
			if (string.Equals(((Element)elementType).Name, originalElementName, StringComparison.Ordinal))
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
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Expected O, but got Unknown
		if (doc == null || events == null || events.Count == 0)
		{
			return;
		}
		try
		{
			FailureMessage message = new FailureMessage(ProtectedChangeFailureId);
			List<ElementId> failingIds = (from x in events
				select ParseElementIdText(x.ElementIdText) into x
				where x != null && RevitElementIdCompat.CompatIntegerValue(x) > 0
				select x).ToList();
			if (failingIds.Count == 1)
			{
				message.SetFailingElement(failingIds[0]);
			}
			else if (failingIds.Count > 1)
			{
				message.SetFailingElements((ICollection<ElementId>)failingIds);
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
		//IL_001f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0025: Expected O, but got Unknown
		if (int.TryParse(value ?? string.Empty, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) && parsed > 0)
		{
			return new ElementId((long)parsed);
		}
		return null;
	}

	private static void RefreshProtectedElementIndex(Document doc)
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_0072: Unknown result type (might be due to invalid IL or missing references)
		if (doc == null || doc.IsFamilyDocument)
		{
			return;
		}
		Dictionary<int, ProtectedElementInfo> index = new Dictionary<int, ProtectedElementInfo>();
		foreach (Family family in ((IEnumerable)new FilteredElementCollector(doc).OfClass(typeof(Family))).Cast<Family>())
		{
			ProtectedElementInfo info = BuildProtectedElementInfo((Element)(object)family);
			if (info != null)
			{
				index[GetElementIdInteger(((Element)family).Id)] = info;
			}
		}
		foreach (ElementType elementType in ((IEnumerable)new FilteredElementCollector(doc).WhereElementIsElementType()).Cast<ElementType>())
		{
			ProtectedElementInfo info2 = BuildProtectedElementInfo((Element)(object)elementType);
			if (info2 != null)
			{
				index[GetElementIdInteger(((Element)elementType).Id)] = info2;
			}
		}
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			ProtectedElementIndexes[BuildDocumentKey(doc)] = index;
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
		Family family = (Family)(object)((element is Family) ? element : null);
		if (family != null)
		{
			return new ProtectedElementInfo
			{
				Kind = "Family",
				Name = (((Element)family).Name ?? string.Empty),
				ElementName = (((Element)family).Name ?? string.Empty),
				CategoryName = ResolveFamilyCategoryName(family)
			};
		}
		FamilySymbol symbol = (FamilySymbol)(object)((element is FamilySymbol) ? element : null);
		if (symbol != null)
		{
			return new ProtectedElementInfo
			{
				Kind = "Loadable Family Type",
				Name = ResolveFamilySymbolName(symbol),
				ElementName = (((Element)symbol).Name ?? string.Empty),
				CategoryName = ResolveElementCategoryName((Element)(object)symbol)
			};
		}
		ElementType elementType = (ElementType)(object)((element is ElementType) ? element : null);
		if (elementType != null && SystemTypeNames.Contains(((object)elementType).GetType().Name))
		{
			return new ProtectedElementInfo
			{
				Kind = "System Type",
				Name = (((Element)elementType).Name ?? string.Empty),
				ElementName = (((Element)elementType).Name ?? string.Empty),
				CategoryName = ResolveElementCategoryName((Element)(object)elementType)
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
			familyName = ((ElementType)symbol).FamilyName ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		string typeName = ((Element)symbol).Name ?? string.Empty;
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
			Category familyCategory = family.FamilyCategory;
			ResolveFamilyCategoryName = ((familyCategory != null) ? familyCategory.Name : null) ?? string.Empty;
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
			Category category = element.Category;
			ResolveElementCategoryName = ((category != null) ? category.Name : null) ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveElementCategoryName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveElementCategoryName;
	}

	private static FamilyBrowserStandardPolicy LoadPolicy()
	{
		string currentUser = FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity();
		DateTime now = DateTime.UtcNow;
		object syncRoot = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (CachedPolicy != null && string.Equals(CachedPolicyUser, currentUser, StringComparison.OrdinalIgnoreCase) && (now - CachedPolicyLoadedUtc).TotalSeconds < 10.0)
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
		object syncRoot2 = SyncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot2);
		bool lockTaken2 = false;
		try
		{
			Monitor.Enter(syncRoot2, ref lockTaken2);
			CachedPolicy = policy;
			CachedPolicyUser = currentUser;
			CachedPolicyLoadedUtc = now;
		}
		finally
		{
			if (lockTaken2)
			{
				Monitor.Exit(syncRoot2);
			}
		}
		return policy;
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
				if (includeCentralPath && doc.IsWorkshared)
				{
					ModelPath centralPath = doc.GetWorksharingCentralModelPath();
					if (centralPath != null)
					{
						context.CentralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(centralPath);
					}
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
		if (id == null)
		{
			return -1;
		}
		return RevitElementIdCompat.CompatIntegerValue(id);
	}

	private static string GetElementIdText(ElementId id)
	{
		if (id == null)
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
