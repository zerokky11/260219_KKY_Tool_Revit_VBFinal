using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic.CompilerServices;

public sealed class FamilyBrowserRevitBridgeExternalEventHandler : IExternalEventHandler
{
	private sealed class PendingBridgeRequest
	{
		public FamilyBrowserBridgeRequest Request { get; }

		public ManualResetEvent Completed { get; }

		public FamilyBrowserBridgeResponse Response { get; set; }

		public PendingBridgeRequest(FamilyBrowserBridgeRequest request)
		{
			Completed = new ManualResetEvent(initialState: false);
			Request = request;
		}
	}

	private readonly object _syncRoot;

	private PendingBridgeRequest _pending;

	public FamilyBrowserRevitBridgeExternalEventHandler()
	{
		_syncRoot = RuntimeHelpers.GetObjectValue(new object());
	}

	public unsafe FamilyBrowserBridgeResponse ExecuteRequest(FamilyBrowserBridgeRequest request, ExternalEvent externalEvent, int timeoutMilliseconds)
	{
		//IL_005d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0062: Unknown result type (might be due to invalid IL or missing references)
		//IL_0063: Unknown result type (might be due to invalid IL or missing references)
		if (request == null)
		{
			throw new ArgumentNullException("request");
		}
		PendingBridgeRequest pending = new PendingBridgeRequest(request);
		object syncRoot = _syncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (_pending != null)
			{
				return CreateResponse(request, success: false, "Revit bridge is busy. Try again after the current command finishes.");
			}
			_pending = pending;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		ExternalEventRequest raiseResult = externalEvent.Raise();
		if ((int)raiseResult != 0)
		{
			ClearPending(pending);
			return CreateResponse(request, success: false, "Revit is not ready to process the bridge command: " + ((Enum)(*(ExternalEventRequest*)(&raiseResult))/*cast due to .constrained prefix*/).ToString());
		}
		if (!pending.Completed.WaitOne(Math.Max(1000, timeoutMilliseconds)))
		{
			ClearPending(pending);
			return CreateResponse(request, success: false, "Revit did not respond before the bridge timeout.");
		}
		if (pending.Response == null)
		{
			return CreateResponse(request, success: false, "Revit completed the bridge event without a response.");
		}
		return pending.Response;
	}

	public void Execute(UIApplication app)
	{
		FamilyBrowserRevitVersionContext.SetCurrentVersion(app);
		PendingBridgeRequest pending = null;
		object syncRoot = _syncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			pending = _pending;
			_pending = null;
		}
		finally
		{
			if (lockTaken)
			{
				Monitor.Exit(syncRoot);
			}
		}
		if (pending == null)
		{
			return;
		}
		try
		{
			pending.Response = HandleRequest(app, pending.Request);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			pending.Response = CreateResponse(pending.Request, success: false, ex2.Message, app);
			ProjectData.ClearProjectError();
		}
		finally
		{
			pending.Completed.Set();
		}
	}

	public string GetName()
	{
		return "KKY Family Browser Desktop Bridge";
	}

	private FamilyBrowserBridgeResponse HandleRequest(UIApplication app, FamilyBrowserBridgeRequest request)
	{
		string left = Normalize(request.Command);
		if (Operators.CompareString(left, Normalize("Ping"), TextCompare: false) == 0 || Operators.CompareString(left, Normalize("ListDocuments"), TextCompare: false) == 0)
		{
			return CreateResponse(request, success: true, "Connected to Revit.", app);
		}
		if (Operators.CompareString(left, Normalize("ActivateDocument"), TextCompare: false) == 0)
		{
			return ActivateDocument(app, request);
		}
		if (Operators.CompareString(left, Normalize("CheckCurrentModel"), TextCompare: false) == 0)
		{
			return RunProjectComparison(app, request);
		}
		if (Operators.CompareString(left, Normalize("RunSystemPreflight"), TextCompare: false) == 0)
		{
			return RunSystemPreflight(app, request);
		}
		if (Operators.CompareString(left, Normalize("ApplyStandardFamilies"), TextCompare: false) == 0 || Operators.CompareString(left, Normalize("ApplySystemTypes"), TextCompare: false) == 0)
		{
			return CreateResponse(request, success: false, "This desktop bridge command is reserved, but apply operations still require the Revit add-in dashboard in this build.", app);
		}
		return CreateResponse(request, success: false, "Unknown bridge command: " + (request.Command ?? string.Empty), app);
	}

	private FamilyBrowserBridgeResponse ActivateDocument(UIApplication app, FamilyBrowserBridgeRequest request)
	{
		string target = (request.TargetDocumentId ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(target))
		{
			return CreateResponse(request, success: false, "A target document path is required.", app);
		}
		string documentPath = ResolveDocumentPath(app, target);
		if (string.IsNullOrWhiteSpace(documentPath) || !File.Exists(documentPath))
		{
			return CreateResponse(request, success: false, "Only saved RVT documents can be activated from the desktop bridge right now.", app);
		}
		app.OpenAndActivateDocument(documentPath);
		return CreateResponse(request, success: true, "Activated document: " + Path.GetFileName(documentPath), app);
	}

	private FamilyBrowserBridgeResponse RunProjectComparison(UIApplication app, FamilyBrowserBridgeRequest request)
	{
		Document document = ResolveActiveDocument(app);
		if (document == null)
		{
			return CreateResponse(request, success: false, "Open a project document before checking the current model.", app);
		}
		string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
		if (!EnsureManagedDataRootForBridge(app, workspaceRoot, request))
		{
			return CreateResponse(request, success: false, T("No managed shared folder is connected. Refresh the homepage path in Family Browser first.", "공용 관리 폴더가 연결되지 않았습니다. 먼저 Family Browser에서 홈페이지 경로를 다시 확인하세요."), app);
		}
		StandardLibraryRegistrationRecord registration = LoadActiveRegistration(workspaceRoot);
		if (string.IsNullOrWhiteSpace(registration.LastSnapshotPath) || !File.Exists(registration.LastSnapshotPath))
		{
			return CreateResponse(request, success: false, "The registered standard snapshot could not be found. Re-register the standard RVT.", app);
		}
		StandardLibrarySnapshot standardSnapshot = DataContractJsonFileStore.Load<StandardLibrarySnapshot>(registration.LastSnapshotPath);
		ProjectContentSnapshot projectSnapshot = ProjectSnapshotCaptureService.Capture(document);
		string projectSnapshotPath = ProjectSnapshotStore.Save(workspaceRoot, projectSnapshot, document);
		ProjectTrackingCatalog trackingCatalog = ProjectTrackingStoreService.Load(document);
		ProjectStandardComparisonReport report = ProjectStandardComparisonService.BuildReport(registration, registration.LastSnapshotPath, standardSnapshot, projectSnapshotPath, projectSnapshot, trackingCatalog);
		ProjectStandardComparisonStore.Save(workspaceRoot, report);
		checked
		{
			int loadableActionCount = report.Summary.LoadableLoadAvailableCount + report.Summary.LoadableDifferentCount;
			int systemActionCount = report.Summary.SystemLoadAvailableCount + report.Summary.SystemDifferentCount;
			return CreateResponse(request, success: true, T("Project check completed.", "현재 모델 검사가 완료되었습니다.") + " " + T("Loadable latest", "로더블 기준 일치") + ": " + report.Summary.LoadableLatestCount + ", " + T("Loadable action", "로더블 조치") + ": " + loadableActionCount + ", " + T("System latest", "시스템 기준 일치") + ": " + report.Summary.SystemLatestCount + ", " + T("System action", "시스템 조치") + ": " + systemActionCount + ". " + T("Comparison report was saved for admin diagnostics.", "비교 보고서는 관리자 진단용으로 저장되었습니다."), app);
		}
	}

	private FamilyBrowserBridgeResponse RunSystemPreflight(UIApplication app, FamilyBrowserBridgeRequest request)
	{
		Document document = ResolveActiveDocument(app);
		if (document == null)
		{
			return CreateResponse(request, success: false, "Open a project document before running system type review.", app);
		}
		string workspaceRoot = HostWorkspacePathResolver.ResolveRoot();
		if (!EnsureManagedDataRootForBridge(app, workspaceRoot, request))
		{
			return CreateResponse(request, success: false, T("No managed shared folder is connected. Refresh the homepage path in Family Browser first.", "공용 관리 폴더가 연결되지 않았습니다. 먼저 Family Browser에서 홈페이지 경로를 다시 확인하세요."), app);
		}
		StandardLibraryRegistrationRecord registration = LoadActiveRegistration(workspaceRoot);
		Document standardDoc = null;
		bool shouldCloseStandardDoc = false;
		try
		{
			standardDoc = StandardLibraryDocumentResolver.OpenRegisteredDocument(app.Application, registration, ref shouldCloseStandardDoc);
			SystemTypePreflightReport report = SystemTypePreflightBuilderService.BuildReport(registration, standardDoc, document);
			SystemTypePreflightStore.Save(workspaceRoot, report);
			int reviewCount = checked(report.Summary.ApprovalRequiredCount + report.Summary.DependencyManualReviewCount);
			return CreateResponse(request, success: true, T("System type review completed.", "시스템 타입 검토가 완료되었습니다.") + " " + T("No change", "변경 없음") + ": " + report.Summary.NoChangeCount + ", " + T("Ready", "적용 가능") + ": " + report.Summary.ReadyCount + ", " + T("Review", "검토") + ": " + reviewCount + ", " + T("Blocked", "차단") + ": " + report.Summary.BlockedCount + ", " + T("Dependency reload", "의존 패밀리 재로드") + ": " + report.Summary.DependencyReloadCount + ". " + T("Review report was saved for admin diagnostics.", "검토 보고서는 관리자 진단용으로 저장되었습니다."), app);
		}
		finally
		{
			if (shouldCloseStandardDoc && standardDoc != null)
			{
				try
				{
					standardDoc.Close(false);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
	}

	private static StandardLibraryRegistrationRecord LoadActiveRegistration(string workspaceRoot)
	{
		string path = Path.Combine(FamilyBrowserStandardPolicyStore.GetRegistryFolder(workspaceRoot), "active-standard-library.json");
		if (!File.Exists(path))
		{
			throw new InvalidOperationException(T("No registered standard RVT was found. Register the standard RVT first.", "등록된 표준 RVT가 없습니다. 먼저 표준 RVT를 등록하세요."));
		}
		return DataContractJsonFileStore.Load<StandardLibraryRegistrationRecord>(path);
	}

	private static bool EnsureManagedDataRootForBridge(UIApplication app, string workspaceRoot, FamilyBrowserBridgeRequest request)
	{
		try
		{
			FamilyBrowserDeploymentBootstrapService.TryApply(workspaceRoot, (request == null || string.IsNullOrWhiteSpace(request.CreatedBy)) ? FamilyBrowserSecurityPolicyService.ResolveCurrentUserIdentity() : request.CreatedBy);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return FamilyBrowserStandardPolicyStore.IsManagedDataRootAvailable(workspaceRoot);
	}

	private static string T(string englishText, string koreanText)
	{
		return FamilyBrowserLanguageService.Text(englishText, koreanText);
	}

	private static Document ResolveActiveDocument(UIApplication app)
	{
		if (app == null || app.ActiveUIDocument == null)
		{
			return null;
		}
		return app.ActiveUIDocument.Document;
	}

	private static string ResolveDocumentPath(UIApplication app, string targetDocumentId)
	{
		//IL_002d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0033: Expected O, but got Unknown
		if (app == null || string.IsNullOrWhiteSpace(targetDocumentId))
		{
			return string.Empty;
		}
		foreach (Document document2 in app.Application.Documents)
		{
			Document document = document2;
			if (string.Equals(BuildDocumentId(document), targetDocumentId, StringComparison.OrdinalIgnoreCase) || string.Equals(document.PathName, targetDocumentId, StringComparison.OrdinalIgnoreCase))
			{
				return document.PathName;
			}
		}
		return targetDocumentId;
	}

	private static FamilyBrowserBridgeResponse CreateResponse(FamilyBrowserBridgeRequest request, bool success, string message, UIApplication app = null)
	{
		FamilyBrowserBridgeResponse response = new FamilyBrowserBridgeResponse
		{
			RequestId = ((request == null) ? string.Empty : request.RequestId),
			Success = success,
			Message = (message ?? string.Empty)
		};
		if (app != null)
		{
			response.RevitVersion = ResolveRevitVersion(app);
			response.ActiveDocumentTitle = ResolveActiveDocumentTitle(app);
			response.Documents = ListDocuments(app);
		}
		return response;
	}

	private static List<FamilyBrowserBridgeDocumentInfo> ListDocuments(UIApplication app)
	{
		//IL_0030: Unknown result type (might be due to invalid IL or missing references)
		//IL_0037: Expected O, but got Unknown
		List<FamilyBrowserBridgeDocumentInfo> documents = new List<FamilyBrowserBridgeDocumentInfo>();
		if (app == null)
		{
			return documents;
		}
		Document activeDocument = ResolveActiveDocument(app);
		foreach (Document document2 in app.Application.Documents)
		{
			Document document = document2;
			documents.Add(new FamilyBrowserBridgeDocumentInfo
			{
				DocumentId = BuildDocumentId(document),
				Title = document.Title,
				Path = (document.PathName ?? string.Empty),
				CentralPath = ResolveCentralPath(document),
				IsActive = object.ReferenceEquals(document, activeDocument),
				IsWorkshared = document.IsWorkshared
			});
		}
		return documents.OrderByDescending([SpecialName] (FamilyBrowserBridgeDocumentInfo x) => x.IsActive).ThenBy<FamilyBrowserBridgeDocumentInfo, string>([SpecialName] (FamilyBrowserBridgeDocumentInfo x) => x.Title, StringComparer.OrdinalIgnoreCase).ToList();
	}

	private static string BuildDocumentId(Document document)
	{
		if (document == null)
		{
			return string.Empty;
		}
		if (!string.IsNullOrWhiteSpace(document.PathName))
		{
			return document.PathName;
		}
		return document.Title + "|" + document.GetHashCode().ToString(CultureInfo.InvariantCulture);
	}

	private static string ResolveCentralPath(Document document)
	{
		string ResolveCentralPath;
		if (document == null || !document.IsWorkshared)
		{
			ResolveCentralPath = string.Empty;
		}
		else
		{
			try
			{
				ModelPath modelPath = document.GetWorksharingCentralModelPath();
				ResolveCentralPath = ((modelPath != null) ? ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath) : string.Empty);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ResolveCentralPath = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return ResolveCentralPath;
	}

	private static string ResolveActiveDocumentTitle(UIApplication app)
	{
		Document document = ResolveActiveDocument(app);
		if (document == null)
		{
			return string.Empty;
		}
		return document.Title;
	}

	private static string ResolveRevitVersion(UIApplication app)
	{
		if (app == null || app.Application == null)
		{
			return string.Empty;
		}
		return app.Application.VersionNumber;
	}

	private void ClearPending(PendingBridgeRequest pending)
	{
		object syncRoot = _syncRoot;
		ObjectFlowControl.CheckForSyncLockOnValueType(syncRoot);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(syncRoot, ref lockTaken);
			if (object.ReferenceEquals(_pending, pending))
			{
				_pending = null;
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

	private static string Normalize(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().ToLowerInvariant();
	}
}
