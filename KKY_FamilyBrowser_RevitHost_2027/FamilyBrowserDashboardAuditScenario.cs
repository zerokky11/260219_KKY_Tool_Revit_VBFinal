using System;
using System.Collections.Generic;
using System.Globalization;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

public sealed class FamilyBrowserDashboardAuditScenario
{
	public string Name { get; set; }

	public string WorkspaceRoot { get; set; }

	public string LanguageCode { get; set; }

	public string InitialLanguageCode { get; set; }

	public string ThemeCode { get; set; }

	public string MeasurementUnitCode { get; set; }

	public string ActiveTab { get; set; }

	public string BrowseDisciplineKey { get; set; }

	public string PolicyActiveDisciplineKey { get; set; }

	public bool AdminMode { get; set; }

	public bool AdminProfile { get; set; }

	public bool FileGuardProtected { get; set; }

	public bool StandardRvtRegistered { get; set; }

	public bool StandardListRegistered { get; set; }

	public bool StandardRvtChanged { get; set; }

	public bool StandardRvtUnavailable { get; set; }

	public bool ProjectCatalogBaselineMissing { get; set; }

	public bool ProjectCatalogChanged { get; set; }

	public bool ProjectCatalogUntracked { get; set; }

	public int TrackingPendingCount { get; set; }

	public bool IncludeRows { get; set; }

	public bool IncludePendingRows { get; set; }

	public bool IncludeRequests { get; set; }

	public bool IncludeUnregistered { get; set; }

	public bool IncludeReadinessWarning { get; set; }

	public bool ManagedFolderUnavailable { get; set; }

	public bool ManagedFolderTestOverride { get; set; }

	public bool HomepageManagedFolderAvailable { get; set; }

	public bool CompareDetailedSystemTypeComponents { get; set; }

	public bool TrackProjectElementChanges { get; set; }

	public int SyntheticFamilyCount { get; set; }

	public int SyntheticSystemCount { get; set; }

	public string UserIdentity { get; set; }

	public string ProjectTitle { get; set; }

	public string ProjectPath { get; set; }

	public string CentralPath { get; set; }

	public FamilyBrowserDashboardAuditScenario()
	{
		Name = "audit-default";
		WorkspaceRoot = string.Empty;
		LanguageCode = "ko";
		InitialLanguageCode = string.Empty;
		ThemeCode = "light";
		MeasurementUnitCode = "mm";
		ActiveTab = "home";
		BrowseDisciplineKey = "Mechanical";
		PolicyActiveDisciplineKey = string.Empty;
		AdminMode = true;
		AdminProfile = false;
		FileGuardProtected = false;
		StandardRvtRegistered = true;
		StandardListRegistered = true;
		StandardRvtChanged = false;
		StandardRvtUnavailable = false;
		ProjectCatalogBaselineMissing = false;
		ProjectCatalogChanged = false;
		ProjectCatalogUntracked = false;
		TrackingPendingCount = 0;
		IncludeRows = true;
		IncludePendingRows = false;
		IncludeRequests = true;
		IncludeUnregistered = true;
		IncludeReadinessWarning = false;
		ManagedFolderUnavailable = false;
		ManagedFolderTestOverride = false;
		HomepageManagedFolderAvailable = false;
		CompareDetailedSystemTypeComponents = true;
		TrackProjectElementChanges = false;
		SyntheticFamilyCount = 0;
		SyntheticSystemCount = 0;
		UserIdentity = "KKY_UI_AUDIT_ADMIN";
		ProjectTitle = "UI Audit Project";
		ProjectPath = "C:\\KKY Audit\\Local\\UI_Audit_Local.rvt";
		CentralPath = "C:\\KKY Audit\\Central\\UI_Audit_Central.rvt";
	}
}

public partial class FamilyBrowserDashboardHtmlForm
{
	private string _auditMeasurementDisplayUnit;

	private bool? _auditDetailedSystemTypeComparisonEnabled;

	private int _auditTrackingPendingCount = -1;

	public static string BuildDashboardHtmlForAudit(FamilyBrowserDashboardAuditScenario scenario)
	{
		using (FamilyBrowserDashboardHtmlForm form = new FamilyBrowserDashboardHtmlForm(scenario ?? new FamilyBrowserDashboardAuditScenario()))
		{
			return form.BuildDashboardHtml();
		}
	}

	public static string BuildStartupShellHtmlForAudit(string languageCode)
	{
		return BuildStartupShellHtml(!string.Equals(languageCode, "en", StringComparison.OrdinalIgnoreCase), FamilyBrowserUiTheme.Light);
	}

	private FamilyBrowserDashboardHtmlForm(FamilyBrowserDashboardAuditScenario scenario)
	{
		_uiApplication = null;
		_workspaceRoot = ResolveAuditWorkspaceRoot(scenario);
		_browser = null;
		_startupOverlay = null;
		_modelessActionDispatcher = null;
		_requestRefreshTimer = null;
		_homepageSecurityRefreshTimer = null;
		_startupOverlayReadyTimer = null;
		_projectContentMutationSerials = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
		_projectScanMutationBaselines = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
		InitializeAuditScenarioState(scenario ?? new FamilyBrowserDashboardAuditScenario());
	}

	private static string ResolveAuditWorkspaceRoot(FamilyBrowserDashboardAuditScenario scenario)
	{
		if (scenario != null && !string.IsNullOrWhiteSpace(scenario.WorkspaceRoot))
		{
			return scenario.WorkspaceRoot;
		}
		try
		{
			return HostWorkspacePathResolver.ResolveRoot();
		}
		catch
		{
			return AppDomain.CurrentDomain.BaseDirectory;
		}
	}

	private void InitializeAuditScenarioState(FamilyBrowserDashboardAuditScenario scenario)
	{
		string targetLanguageCode = string.Equals(scenario.LanguageCode, "en", StringComparison.OrdinalIgnoreCase) ? "en" : "ko";
		_auditMeasurementDisplayUnit = FamilyBrowserMeasurementUnitPreferenceService.Normalize(scenario.MeasurementUnitCode);
		_auditDetailedSystemTypeComparisonEnabled = scenario.CompareDetailedSystemTypeComponents;
		_auditTrackingPendingCount = Math.Max(0, scenario.TrackingPendingCount);
		_languageCode = string.IsNullOrWhiteSpace(scenario.InitialLanguageCode) ? targetLanguageCode : (string.Equals(scenario.InitialLanguageCode, "en", StringComparison.OrdinalIgnoreCase) ? "en" : "ko");
		_uiTheme = FamilyBrowserUiThemeService.Parse(scenario.ThemeCode);
		_adminModeEnabled = scenario.AdminMode;
		_standardPolicy = BuildAuditPolicy(scenario);
		bool isAdminProfile = scenario.AdminMode || scenario.AdminProfile;
		FamilyBrowserSecurityPolicyService.SetCurrentUserIdentityOverride(string.IsNullOrWhiteSpace(scenario.UserIdentity) ? (isAdminProfile ? "KKY_UI_AUDIT_ADMIN" : "KKY_UI_AUDIT_MODELER") : scenario.UserIdentity);
		_registration = scenario.StandardRvtRegistered ? BuildAuditRegistration() : null;
		_lastComparisonPath = string.Empty;
		_lastPreflightPath = string.Empty;
		_projectTitle = string.IsNullOrWhiteSpace(scenario.ProjectTitle) ? (string.IsNullOrWhiteSpace(scenario.Name) ? "UI Audit Project" : scenario.Name) : scenario.ProjectTitle;
		_projectPath = string.IsNullOrWhiteSpace(scenario.ProjectPath) ? "C:\\KKY Audit\\Local\\UI_Audit_Local.rvt" : scenario.ProjectPath;
		_standardText = scenario.StandardRvtRegistered ? T("Standard: Mechanical / registered", "표준: 설비 / 등록 1/6") : T("Standard: not registered", "표준: 미등록");
		_loadText = T("Load Available: ", "로드 가능: ") + (scenario.IncludeRows ? "1" : "0");
		_updateText = T("Update Available: ", "업데이트 가능: ") + (scenario.IncludeRows ? "1" : "0");
		_permissionText = scenario.AdminMode ? T("Permission: Admin", "권한: 관리자") : (isAdminProfile ? T("Permission: Admin profile (Admin Mode Off)", "권한: 관리자 프로필 (관리자 모드 OFF)") : T("Permission: Modeler", "권한: 모델러"));
		_trackingText = scenario.IncludeReadinessWarning ? T("Tracking: review needed", "추적: 검토 필요") : T("Tracking: ready", "추적: 준비됨");
		_trackingTone = scenario.IncludeReadinessWarning ? "warn" : "good";
		_statusMessage = "UI audit scenario";
		_systemSummary = "UI audit system summary";
		_modelStateText = "Audit model";
		_centralPathText = string.IsNullOrWhiteSpace(scenario.CentralPath) ? "-" : scenario.CentralPath;
		_standardPathText = scenario.StandardRvtRegistered ? "audit://standard.rvt" : "-";
		_standardSnapshotText = scenario.StandardRvtRegistered ? "audit snapshot" : "-";
		_healthText = scenario.IncludeReadinessWarning ? T("Warnings", "주의") : T("Ready", "준비");
		_projectOnlyText = T("Review needed: ", "검토 필요: ") + (scenario.IncludeUnregistered ? "2" : "0");
		_userRoleText = scenario.AdminMode ? "Admin" : (isAdminProfile ? T("Admin profile (Admin Mode Off)", "관리자 프로필 (관리자 모드 OFF)") : "Modeler");
		_nextWorkflowText = "UI audit next workflow";
		_projectScanText = "UI audit project scan";
		_loadableRows = ExpandAuditFamilyRows(BuildAuditFamilyRows(scenario), scenario.SyntheticFamilyCount);
		_systemRows = ExpandAuditSystemRows(BuildAuditSystemRows(scenario), scenario.SyntheticSystemCount);
		_unregisteredFamilyRows = scenario.IncludeUnregistered ? BuildAuditUnregisteredRows("Family") : new List<UnregisteredProjectItemRow>();
		_unregisteredSystemRows = scenario.IncludeUnregistered ? BuildAuditUnregisteredRows("System") : new List<UnregisteredProjectItemRow>();
		_unregisteredListStatusText = scenario.IncludeUnregistered ? T("Audit unregistered items are available.", "감사용 미등록 항목이 있습니다.") : string.Empty;
		_requestRecords = scenario.IncludeRequests ? BuildAuditRequests() : new List<FamilyBrowserRequestRecord>();
		_requestRecordsLoaded = scenario.IncludeRequests;
		_requestItems = new List<string>();
		_auditItems = new List<string>();
		_readinessReport = BuildAuditReadinessReport(scenario);
		_lastOperationalReadinessReport = _readinessReport;
		_deploymentBootstrapResult = null;
		_auditManagedFolderTestOverride = scenario.ManagedFolderTestOverride;
		_auditTestManagedFolderRoot = "\\\\audit-test\\KKY\\FamilyBrowser";
		_homepageManagedFolderProbeResult = scenario.HomepageManagedFolderAvailable ? new FamilyBrowserHomepageManagedFolderProbeResult
		{
			Available = true,
			Source = "audit://homepage-bootstrap",
			BootstrapUrl = "audit://homepage-bootstrap",
			ManagedPolicyPath = "\\\\audit-homepage\\KKY\\FamilyBrowser\\Config\\standard-policy.json",
			ManagedRootPath = "\\\\audit-homepage\\KKY\\FamilyBrowser",
			CheckedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
		} : null;
		_managedFolderUnavailable = scenario.ManagedFolderUnavailable;
		_managedFolderOnboardingIssue = scenario.ManagedFolderUnavailable ? "Audit management folder is unavailable." : string.Empty;
		_lastSystemPreflightReport = null;
		_activeStandardListCatalog = null;
		// Missing registration and missing scan are separate setup states.
		_activeStandardScanNeeded = false;
		_modelerStandardListFilterMissing = scenario.StandardRvtRegistered && !scenario.StandardListRegistered;
		if (_registration != null)
		{
			FamilyBrowserStandardRevisionState revisionState = new FamilyBrowserStandardRevisionState
			{
				SourceId = _registration.SourceId,
				StandardRvtPath = _registration.ResolvedPath,
				StateCode = scenario.StandardRvtChanged ? "Changed" : (scenario.StandardRvtUnavailable ? "Unavailable" : "Current"),
				CheckedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
				BaselineAtUtc = DateTime.UtcNow.AddDays(-1.0).ToString("O", CultureInfo.InvariantCulture),
				RecordedLastWriteUtc = DateTime.UtcNow.AddDays(-1.0).ToString("O", CultureInfo.InvariantCulture),
				CurrentLastWriteUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
				RecordedLength = 1024L,
				CurrentLength = scenario.StandardRvtChanged ? 2048L : 1024L,
				Changed = scenario.StandardRvtChanged,
				Unavailable = scenario.StandardRvtUnavailable,
				Reason = scenario.StandardRvtChanged ? "Audit source changed after the accepted scan." : (scenario.StandardRvtUnavailable ? "Audit source is unavailable." : "Audit source matches the accepted scan.")
			};
			StoreStandardRevisionState(_registration, scenario.BrowseDisciplineKey, revisionState);
			_activeStandardRevisionState = revisionState;
			_activeStandardRevisionBlocked = revisionState.BlocksStandardUse;
			if (_activeStandardRevisionBlocked)
			{
				_activeStandardScanNeeded = true;
				_loadableRows.Clear();
				_systemRows.Clear();
				_trackingText = T("Standard source: rescan required", "표준 원본: 재스캔 필요");
				_trackingTone = "bad";
				_statusMessage = BuildStandardRevisionBlockingMessage(revisionState);
			}
		}
		_activeTab = NormalizeDashboardTabName(scenario.ActiveTab);
		_browseDisciplineKey = string.IsNullOrWhiteSpace(scenario.BrowseDisciplineKey) ? "Mechanical" : scenario.BrowseDisciplineKey;
		_lastDashboardUiStateJson = string.Empty;
		_fastComparisonCacheKey = string.Empty;
		_fastComparisonCachedProjectScanText = string.Empty;
		_modelerAllSlotsCacheKey = string.Empty;
		_modelerAllSlotsCachedUnregisteredListStatusText = string.Empty;
		_permissionDiagnosticCacheKey = string.Empty;
		_dashboardPermissionCacheKey = string.Empty;
		_longOperationTitle = string.Empty;
		_longOperationMessage = string.Empty;
		_longOperationTotal = 100;
		_longOperationLastPaintUtc = DateTime.MinValue;
		_startupOverlayReadyPollStartedUtc = DateTime.MinValue;
		_startupOverlayProgressLastPaintUtc = DateTime.MinValue;
		InitializeProjectCatalogAuditState(scenario);
		if (!string.Equals(_languageCode, targetLanguageCode, StringComparison.Ordinal))
		{
			_languageCode = targetLanguageCode;
			RefreshLanguageSensitiveDisplayState();
			_statusMessage = T("Language changed.", "언어를 변경했습니다.");
		}
		AppendAuditPendingRows(scenario);
		_systemComparisonRows = new List<SystemRow>(_systemRows);
	}

	private void AppendAuditPendingRows(FamilyBrowserDashboardAuditScenario scenario)
	{
		if (scenario == null || !scenario.IncludePendingRows || !scenario.StandardRvtRegistered || !scenario.StandardListRegistered)
		{
			return;
		}
		_loadableRows.Add(new BrowserRow
		{
			Status = T("Loaded - Save/Sync Pending", "로드됨 · 저장/동기화 대기"),
			RawStatus = "FamilyPendingSaveOrSync",
			DisciplineKey = "Mechanical",
			DisciplineLabel = T("Mechanical", "설비"),
			Name = "AUDIT_PENDING_FAMILY",
			Category = "Mechanical Equipment",
			CategoryGroup = "Model",
			Action = T("Save or synchronize to confirm", "저장 또는 동기화 후 확정"),
			Notes = T("Loaded in the current Revit session. Closing without saving discards this temporary state.", "현재 Revit 세션에 로드되었습니다. 저장하지 않고 닫으면 이 임시 상태는 폐기됩니다."),
			ApprovedRev = "A-PENDING",
			LoadedRev = T("Project status: Loaded - save/sync pending", "프로젝트 상태: 로드됨 · 저장/동기화 대기"),
			ChangeSummary = T("Temporary until a successful save or synchronization.", "저장 또는 동기화가 성공할 때까지 임시 상태입니다.")
		});
		_systemRows.Add(new SystemRow
		{
			Status = T("Applied - Save/Sync Pending", "적용됨 · 저장/동기화 대기"),
			RawStatus = "SystemPendingSaveOrSync",
			DisciplineKey = "Mechanical",
			DisciplineLabel = T("Mechanical", "설비"),
			Name = "AUDIT_PENDING_SYSTEM",
			Category = "Duct System",
			SystemFamilyKind = "DuctType",
			Action = T("Save or synchronize to confirm", "저장 또는 동기화 후 확정"),
			Notes = T("Applied in the current Revit session. Closing without saving discards this temporary state.", "현재 Revit 세션에 적용되었습니다. 저장하지 않고 닫으면 이 임시 상태는 폐기됩니다."),
			DifferenceSummaryTable = string.Empty
		});
	}

	private static FamilyBrowserStandardPolicy BuildAuditPolicy(FamilyBrowserDashboardAuditScenario scenario)
	{
		FamilyBrowserStandardPolicy policy = new FamilyBrowserStandardPolicy();
		policy.LastUpdatedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		policy.LastUpdatedBy = "ui-audit";
		policy.ActiveDiscipline = !string.IsNullOrWhiteSpace(scenario.PolicyActiveDisciplineKey)
			? scenario.PolicyActiveDisciplineKey
			: (string.IsNullOrWhiteSpace(scenario.BrowseDisciplineKey) ? "Mechanical" : scenario.BrowseDisciplineKey);
		policy.CompareDetailedSystemTypeComponents = scenario.CompareDetailedSystemTypeComponents;
		policy.TrackProjectElementChanges = scenario.TrackProjectElementChanges;
		policy.Security.AdminUsers.Clear();
		if (scenario.AdminMode || scenario.AdminProfile)
		{
			policy.Security.AdminUsers.Add(string.IsNullOrWhiteSpace(scenario.UserIdentity) ? "KKY_UI_AUDIT_ADMIN" : scenario.UserIdentity);
		}
		policy.Security.AllowUnlistedUsersAsModelers = true;
		policy.Security.AllowModelersToLoadFamilies = true;
		policy.Security.AllowModelersToApplySystemTypes = true;
		policy.Security.AllowModelersToSubmitRequests = true;
		if (scenario.FileGuardProtected)
		{
			policy.FileGuard.Enabled = true;
			policy.FileGuard.RootFolder = Path.GetDirectoryName(string.IsNullOrWhiteSpace(scenario.ProjectPath) ? "C:\\KKY Audit\\Local\\UI_Audit_Local.rvt" : scenario.ProjectPath) ?? string.Empty;
			policy.FileGuard.Targets.Clear();
			policy.FileGuard.Targets.Add(new FamilyBrowserFileGuardTarget
			{
				Enabled = true,
				FileName = Path.GetFileName(string.IsNullOrWhiteSpace(scenario.ProjectPath) ? "UI_Audit_Local.rvt" : scenario.ProjectPath),
				Discipline = policy.ActiveDiscipline,
				BlockFamilyLoadAndEdit = true,
				BlockTypeChanges = true,
				LastUpdatedUtc = policy.LastUpdatedUtc,
				LastUpdatedBy = "ui-audit"
			});
		}
		policy.RequestStore.Mode = "Network";
		policy.RequestStore.Path = Path.Combine(Path.GetTempPath(), "KKY_FamilyBrowser_UiAudit_RequestStore");
		foreach (FamilyBrowserStandardLibrarySlot slot in policy.DisciplineLibraries)
		{
			if (slot == null)
			{
				continue;
			}
			slot.Enabled = true;
			if (scenario.StandardRvtRegistered)
			{
				slot.StandardRvtPath = "audit://standard-" + slot.Discipline + ".rvt";
				slot.RegistrationPath = "audit://registration-" + slot.Discipline + ".json";
				slot.SnapshotPath = "audit://snapshot-" + slot.Discipline + ".json";
				slot.SourceId = "audit-" + slot.Discipline;
				slot.LastSnapshotAtUtc = policy.LastUpdatedUtc;
			}
			if (scenario.StandardListRegistered)
			{
				slot.StandardListPath = "audit://standard-list-" + slot.Discipline + ".xlsx";
				slot.StandardListSheetName = "Standards";
			}
		}
		return policy;
	}

	private static StandardLibraryRegistrationRecord BuildAuditRegistration()
	{
		return new StandardLibraryRegistrationRecord
		{
			SourceId = "audit-standard",
			DisplayName = "Audit Standard RVT",
			SourceKind = "File",
			Locator = "audit://standard.rvt",
			ResolvedPath = "audit://standard.rvt",
			SnapshotMode = "Precise",
			RegisteredAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			RegisteredBy = "ui-audit",
			LastSnapshotAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			LastSnapshotPath = "audit://snapshot.json",
			RevitVersion = "Audit"
		};
	}

	private List<BrowserRow> BuildAuditFamilyRows(FamilyBrowserDashboardAuditScenario scenario)
	{
		if (!scenario.StandardRvtRegistered || !scenario.StandardListRegistered || !scenario.IncludeRows)
		{
			return new List<BrowserRow>();
		}
		string previewImagePath = EnsureAuditPreviewImage(_workspaceRoot);
		return new List<BrowserRow>
		{
			new BrowserRow
			{
				Status = T("Load Available", "로드 가능"),
				RawStatus = "LoadAvailable",
				DisciplineKey = "Mechanical",
				DisciplineLabel = T("Mechanical", "설비"),
				Name = "AUDIT_SUPPLY_DIFFUSER",
				Category = "Mechanical Equipment",
				CategoryGroup = "Model",
				Action = T("Load", "로드"),
				Notes = T("Audit loadable family row.", "감사용 로드 가능 패밀리 행입니다."),
				ApprovedRev = "A-001",
				LoadedRev = "-",
				ChangeSummary = T("Missing in project", "프로젝트에 없음"),
				DifferenceSummaryTable = string.Empty,
				TypeSummary = "600x600\n1200x300",
				ParameterSummary = T("Family Parameters\n- Manufacturer: KKY\nInstance Parameters\n- Mark: SUP-01 [instance]\nType Parameters: 600x600\n- Width: 600 [shared]\n- Height: 600 [shared]\n- Airflow: Formula: if(Width > 1000, Width * Height / 1000000, Width * Height / 2000000) + FlowOffset * DiversityFactor\nLookup CSV\n- Audit_SizeTable: 12 rows x 5 columns", "패밀리 파라미터\n- 제조사: KKY\n인스턴스 파라미터\n- 마크: SUP-01 [인스턴스]\n타입 파라미터: 600x600\n- 폭: 600 [공유]\n- 높이: 600 [공유]\n- 풍량: 수식: if(Width > 1000, Width * Height / 1000000, Width * Height / 2000000) + FlowOffset * DiversityFactor\nCSV 테이블\n- Audit_SizeTable: 12행 x 5열"),
				TypeParameterSummary = "@type\t600x600\n@row\t600x600\tShared\tType\tWidth\t600\t-\n@row\t600x600\tShared\tType\tHeight\t600\t-\n@row\t600x600\tShared\tType\tAirflow\t450 CFM\tif(Width > 1000, Width * Height / 1000000, Width * Height / 2000000) + FlowOffset * DiversityFactor\n@row\t600x600\tFamily\tType\tIsEnabled\tYes\t-\n@type\t1200x300\n@row\t1200x300\tShared\tType\tWidth\t1200\t-\n@row\t1200x300\tShared\tType\tHeight\t300\t-\n@row\t1200x300\tShared\tType\tAirflow\t900 CFM\tif(Width > 1000, Width * Height / 1000000, Width * Height / 2000000) + FlowOffset * DiversityFactor\n@row\t1200x300\tFamily\tType\tIsEnabled\tNo\t-",
				NestedSummary = "AUDIT_SUPPLY_DIFFUSER_FACE\nAir Terminals\tAUDIT_SUPPLY_DIFFUSER_FACE\nAUDIT_FLOW_BOX\nMechanical Equipment\tAUDIT_FLOW_BOX",
				PreviewImagePath = previewImagePath,
				PreviewDiagnostic = "Audit 3D preview fixture"
			},
			new BrowserRow
			{
				Status = T("Update Available", "업데이트 가능"),
				RawStatus = "UpdateAvailable",
				DisciplineKey = "Electrical",
				DisciplineLabel = T("Electrical", "전기"),
				Name = "AUDIT_PANEL_TAG",
				Category = "Generic Annotation",
				CategoryGroup = "Annotation",
				Action = T("Review", "검토"),
				Notes = T("Audit annotation family row.", "감사용 주석 패밀리 행입니다."),
				ApprovedRev = "A-002",
				LoadedRev = "A-001",
				ChangeSummary = T("Fingerprint differs", "Fingerprint 다름"),
				DifferenceSummaryTable = T("Type Count\t2 types\t1 type\tType count differs\nParameter\tWidth: 600\tWidth: 750\tValue differs\nLookup CSV\tCSV: yes / table=Audit_SizeTable / 12 rows x 5 columns\tCSV: yes / table=Audit_SizeTable / 10 rows x 5 columns\tLookup CSV row/column count differs.", "타입 수\t2개\t1개\t타입 수 다름\n파라미터\tWidth: 600\tWidth: 750\t값 다름\nCSV 테이블\tCSV 있음 / 테이블=Audit_SizeTable / 12행 x 5열\tCSV 있음 / 테이블=Audit_SizeTable / 10행 x 5열\tCSV 테이블 행/열 개수가 다릅니다."),
				TypeSummary = "Default",
				ParameterSummary = "Label=Panel"
			},
			new BrowserRow
			{
				Status = T("Different From Standard", "표준과 다름"),
				RawStatus = "DifferentFromStandard",
				DisciplineKey = "Mechanical",
				DisciplineLabel = T("Mechanical", "설비"),
				Name = "AUDIT_COMPOSITE_PARENT",
				Category = "Mechanical Equipment",
				CategoryGroup = "Model",
				Action = T("Review", "검토"),
				Notes = T("Nested dependency requires review: AUDIT_NESTED_FLOW_BOX.", "하위 패밀리 차이로 검토 필요: AUDIT_NESTED_FLOW_BOX."),
				ApprovedRev = "A-010",
				LoadedRev = "A-010",
				ChangeSummary = T("Nested family differs: AUDIT_NESTED_FLOW_BOX - Parameter differs.", "하위 패밀리 차이: AUDIT_NESTED_FLOW_BOX - 파라미터가 다릅니다."),
				DifferenceSummaryTable = T("Nested family\tNested Family: AUDIT_NESTED_FLOW_BOX | Fingerprint: 0A1B2C3D4E5F\tNested Family: AUDIT_NESTED_FLOW_BOX | Fingerprint: F5E4D3C2B1A0\tNested family fingerprint differs\nparameters/formulas\tWidth=600\tWidth=500\tAUDIT_NESTED_FLOW_BOX / parameters/formulas: Parameter value differs.", "하위 패밀리\t하위 패밀리: AUDIT_NESTED_FLOW_BOX | Fingerprint: 0A1B2C3D4E5F\t하위 패밀리: AUDIT_NESTED_FLOW_BOX | Fingerprint: F5E4D3C2B1A0\t하위 패밀리 Fingerprint 차이\n파라미터/공식\tWidth=600\tWidth=500\tAUDIT_NESTED_FLOW_BOX / 파라미터/공식: 파라미터 값이 다릅니다."),
				TypeSummary = "Standard",
				NestedSummary = "Mechanical Equipment\tAUDIT_NESTED_FLOW_BOX",
				HasNestedLoadableDifference = true
			},
			new BrowserRow
			{
				Status = T("Different From Standard", "표준과 다름"),
				RawStatus = "DifferentFromStandard",
				DisciplineKey = "Mechanical",
				DisciplineLabel = T("Mechanical", "설비"),
				Name = "AUDIT_NESTED_FLOW_BOX",
				Category = "Mechanical Equipment",
				CategoryGroup = "Model",
				Action = T("Review parent family", "상위 패밀리 검토"),
				Notes = T("Nested family used by parent families: AUDIT_COMPOSITE_PARENT. Nested helper rows are not loaded independently.", "이 하위 패밀리를 사용하는 상위 패밀리: AUDIT_COMPOSITE_PARENT. 하위/보조 패밀리 행은 단독으로 로드하지 않습니다."),
				ApprovedRev = "A-010",
				LoadedRev = "A-009",
				ChangeSummary = T("Parameter differs: Width.", "파라미터 차이: Width."),
				DifferenceSummaryTable = T("Parameter\tWidth=600\tWidth=500\tNested family parameter differs", "파라미터\tWidth=600\tWidth=500\t하위 패밀리 파라미터가 다릅니다."),
				TypeSummary = "Standard",
				IsNestedLoadableChild = true,
				HasNestedLoadableDifference = true
			},
			new BrowserRow
			{
				Status = T("Nested Family Missing", "하위 패밀리 누락"),
				RawStatus = "NestedMissingFromParent",
				DisciplineKey = "Mechanical",
				DisciplineLabel = T("Mechanical", "설비"),
				Name = "AUDIT_NESTED_MISSING_CHILD",
				Category = "Mechanical Equipment",
				CategoryGroup = "Model",
				Action = T("Update parent family", "상위 패밀리 업데이트 필요"),
				Notes = T("Missing from parent family: AUDIT_MISSING_COMPOSITE_PARENT", "상위 패밀리에서 누락: AUDIT_MISSING_COMPOSITE_PARENT"),
				ApprovedRev = "A-011",
				LoadedRev = "-",
				ChangeSummary = T("Nested family missing from parent family: AUDIT_NESTED_MISSING_CHILD.", "상위 패밀리에서 누락된 하위 패밀리: AUDIT_NESTED_MISSING_CHILD."),
				DifferenceSummaryTable = T("Nested family\tAUDIT_NESTED_MISSING_CHILD\t-\tMissing from parent family AUDIT_MISSING_COMPOSITE_PARENT", "하위 패밀리\tAUDIT_NESTED_MISSING_CHILD\t-\t상위 패밀리 AUDIT_MISSING_COMPOSITE_PARENT에서 누락"),
				TypeSummary = "Standard",
				IsNestedLoadableChild = true,
				HasNestedLoadableDifference = true
			},
			new BrowserRow
			{
				Status = T("Matches Standard", "표준과 일치"),
				RawStatus = "LoadedLatest",
				DisciplineKey = "Mechanical",
				DisciplineLabel = T("Mechanical", "설비"),
				Name = "AUDIT_MATCHING_NESTED_CHILD",
				Category = "Mechanical Equipment",
				CategoryGroup = "Model",
				Action = T("No action", "작업 없음"),
				Notes = T("Matching nested helper stays hidden.", "일치하는 하위/보조 패밀리는 숨깁니다."),
				IsNestedLoadableChild = true
			},
			new BrowserRow
			{
				Status = T("Matches Standard", "표준과 일치"),
				RawStatus = "LoadedLatest",
				DisciplineKey = "Mechanical",
				DisciplineLabel = T("Mechanical", "설비"),
				Name = "AUDIT_RETURN_GRILLE",
				Category = "Air Terminals",
				CategoryGroup = "Model",
				Action = T("No action", "작업 없음"),
				Notes = T("Audit matched family row.", "감사용 표준 일치 패밀리 행입니다."),
				ApprovedRev = "A-003",
				LoadedRev = "A-003",
				ChangeSummary = T("No difference", "차이 없음")
			},
			new BrowserRow
			{
				Status = T("Loaded - Check Needed", "로드됨(검사 필요)"),
				RawStatus = "LoadedNameMatch",
				DisciplineKey = "discipline-architecture",
				DisciplineLabel = T("Architectural", "건축"),
				Name = "AUDIT_DOOR_TAG",
				Category = "Generic Annotation",
				CategoryGroup = "Annotation",
				Action = T("Review", "검토"),
				Notes = T("Audit loaded-name-match row.", "감사용 로드됨 검사 필요 행입니다."),
				ApprovedRev = "A-004",
				LoadedRev = "-",
				ChangeSummary = T("Name match only", "이름만 일치")
			},
			new BrowserRow
			{
				Status = T("Needs Review", "검토 대상"),
				RawStatus = "ManualReview",
				DisciplineKey = "Structure",
				DisciplineLabel = T("Structure", "구조"),
				Name = "AUDIT_REVIEW_FAMILY",
				Category = "Structural Framing",
				CategoryGroup = "Model",
				Action = T("Review", "검토"),
				Notes = T("Audit needs-review family row.", "감사용 검토 대상 패밀리 행입니다."),
				ApprovedRev = "A-005",
				LoadedRev = "A-000",
				ChangeSummary = T("Manual review required", "수동 검토 필요")
			}
		};
	}

	private static string EnsureAuditPreviewImage(string workspaceRoot)
	{
		try
		{
			string root = string.IsNullOrWhiteSpace(workspaceRoot) ? Path.GetTempPath() : workspaceRoot;
			string folder = Path.Combine(root, "artifacts", "family-browser-ui-audit", "fixtures");
			Directory.CreateDirectory(folder);
			string path = Path.Combine(folder, "audit-family-preview.png");
			if (File.Exists(path) && new FileInfo(path).Length > 0)
			{
				return path;
			}
			using (Bitmap bitmap = new Bitmap(280, 180))
			using (Graphics graphics = Graphics.FromImage(bitmap))
			using (Font titleFont = new Font("Segoe UI", 16f, FontStyle.Bold, GraphicsUnit.Point))
			using (Font labelFont = new Font("Segoe UI", 10f, FontStyle.Regular, GraphicsUnit.Point))
			using (Brush background = new SolidBrush(Color.FromArgb(244, 248, 246)))
			using (Brush primary = new SolidBrush(Color.FromArgb(24, 115, 79)))
			using (Brush dark = new SolidBrush(Color.FromArgb(16, 45, 36)))
			using (Pen line = new Pen(Color.FromArgb(34, 180, 127), 4f))
			{
				graphics.Clear(Color.White);
				graphics.FillRectangle(background, 0, 0, bitmap.Width, bitmap.Height);
				graphics.DrawRectangle(line, 18, 22, 244, 118);
				graphics.FillEllipse(primary, 64, 56, 44, 44);
				graphics.FillRectangle(primary, 126, 62, 96, 28);
				graphics.DrawString("AUDIT 3D", titleFont, dark, 58, 112);
				graphics.DrawString("Family preview fixture", labelFont, dark, 64, 142);
				bitmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
			}
			return path;
		}
		catch
		{
			return string.Empty;
		}
	}

	private List<BrowserRow> ExpandAuditFamilyRows(List<BrowserRow> rows, int targetCount)
	{
		rows = rows ?? new List<BrowserRow>();
		if (targetCount <= 0)
		{
			return rows;
		}
		int visibleCount = CountVisibleAuditFamilyRows(rows);
		while (visibleCount > targetCount)
		{
			for (int index = rows.Count - 1; index >= 0; index--)
			{
				if (!IsVisibleAuditFamilyRow(rows[index]))
				{
					continue;
				}
				rows.RemoveAt(index);
				visibleCount--;
				break;
			}
		}
		while (visibleCount < targetCount)
		{
			int number = rows.Count + 1;
			rows.Add(new BrowserRow
			{
				Status = T("Load Available", "로드 가능"),
				RawStatus = "LoadAvailable",
				DisciplineKey = number % 2 == 0 ? "Mechanical" : "Electrical",
				DisciplineLabel = number % 2 == 0 ? T("Mechanical", "설비") : T("Electrical", "전기"),
				Name = "AUDIT_FAMILY_" + number.ToString("D4", CultureInfo.InvariantCulture),
				Category = number % 2 == 0 ? "Mechanical Equipment" : "Electrical Fixtures",
				CategoryGroup = "Model",
				Action = T("Load", "로드"),
				Notes = T("Synthetic performance row.", "성능 검증용 행입니다."),
				ApprovedRev = "P-" + number.ToString("D4", CultureInfo.InvariantCulture),
				LoadedRev = "-",
				ChangeSummary = T("Missing in project", "프로젝트에 없음")
			});
			visibleCount++;
		}
		return rows;
	}

	private static int CountVisibleAuditFamilyRows(List<BrowserRow> rows)
	{
		int count = 0;
		for (int index = 0; index < rows.Count; index++)
		{
			if (IsVisibleAuditFamilyRow(rows[index]))
			{
				count++;
			}
		}
		return count;
	}

	private static bool IsVisibleAuditFamilyRow(BrowserRow row)
	{
		return row != null && (!row.IsNestedLoadableChild || row.HasNestedLoadableDifference);
	}

	private List<SystemRow> ExpandAuditSystemRows(List<SystemRow> rows, int targetCount)
	{
		rows = rows ?? new List<SystemRow>();
		if (targetCount <= 0)
		{
			return rows;
		}
		while (rows.Count > targetCount)
		{
			rows.RemoveAt(rows.Count - 1);
		}
		while (rows.Count < targetCount)
		{
			int number = rows.Count + 1;
			rows.Add(new SystemRow
			{
				Status = T("Apply Available", "적용 가능"),
				RawStatus = "LoadAvailable",
				DisciplineKey = number % 2 == 0 ? "Mechanical" : "FireProtection",
				DisciplineLabel = number % 2 == 0 ? T("Mechanical", "설비") : T("Fire Protection", "소방"),
				Name = "AUDIT_SYSTEM_" + number.ToString("D4", CultureInfo.InvariantCulture),
				Category = number % 2 == 0 ? "Duct System" : "Piping System",
				SystemFamilyKind = number % 2 == 0 ? "Duct" : "Pipe",
				Action = T("Apply", "적용"),
				Notes = T("Synthetic performance row.", "성능 검증용 행입니다.")
			});
		}
		return rows;
	}

	private List<SystemRow> BuildAuditSystemRows(FamilyBrowserDashboardAuditScenario scenario)
	{
		if (!scenario.StandardRvtRegistered || !scenario.StandardListRegistered || !scenario.IncludeRows)
		{
			return new List<SystemRow>();
		}
		return new List<SystemRow>
		{
			new SystemRow
			{
				Status = T("Apply Available", "적용 가능"),
				RawStatus = "LoadAvailable",
				DisciplineKey = "Mechanical",
				DisciplineLabel = T("Mechanical", "설비"),
				Name = "AUDIT_SUPPLY_AIR",
				Category = "Duct System",
				SystemFamilyKind = "Duct",
				Action = T("Apply", "적용"),
				Notes = T("Audit system type row.", "감사용 시스템 타입 행입니다."),
				ParameterSummary = "@system-detail-v1\n@section\tidentity\n@row\tidentity\tClass\tDuctType\n@row\tidentity\tCategory\tDuct System\n@row\tidentity\tType\tAUDIT_SUPPLY_AIR\n@row\tidentity\tSegment\t-\n@row\tidentity\tMaterial\t-\n@section\trouting\n@route\tSegments\t0\tAUDIT_DUCT_SEGMENT\tDuctSegment\tDucts\t12\tsize 1min=0.3280839895013123 max=0.984251968503937; size 2 min=-1E+30 max=1E+30\n@route\tElbows\t0\tAUDIT_DUCT_ELBOW / Standard\tFamilySymbol\tDuct Fittings\t-\t\n@row\trouting\tSegments[0]\tAUDIT_DUCT_SEGMENT | DuctSegment / Ducts | sizes=12 | criteria=size 1min=0.3280839895013123 max=0.984251968503937; size 2 min=-1E+30 max=1E+30\n@section\tsegments\n@row\tsegments\tAUDIT_DUCT_SEGMENT\tDuctSegment / Ducts | sizes=12\n@section\tdependencies\n@row\tdependencies\tElbows\tAUDIT_DUCT_ELBOW / Standard\n@section\tlayers\n@layer\t1\tFinish 1 [4]\tAUDIT_BRICK\t0.3280839895013123\t100 mm\tfalse\tfalse\tfalse\n@row\tlayers\t#1\tFinish 1 [4] / AUDIT_BRICK / 100 mm\n@layer\t2\tStructure [1]\tAUDIT_CONCRETE\t0.6561679790026246\t200 mm\ttrue\ttrue\tfalse\n@row\tlayers\t#2\tStructure [1] / AUDIT_CONCRETE / 200 mm\n@layer\t3\tFinish 2 [5]\tAUDIT_GYPSUM\t0.04921259842519685\t15 mm\tfalse\tfalse\ttrue\n@row\tlayers\t#3\tFinish 2 [5] / AUDIT_GYPSUM / 15 mm",
				LayerSummary = "Finish 1 [4] / AUDIT_BRICK / 100 mm"
			},
			new SystemRow
			{
				Status = T("Different From Standard", "표준과 다름"),
				RawStatus = "DifferentFromStandard",
				DisciplineKey = "Mechanical",
				DisciplineLabel = T("Mechanical", "설비"),
				Name = "AUDIT_GUARDRAIL",
				Category = "Railings",
				SystemFamilyKind = "RailingType",
				Action = T("Review", "검토"),
				Notes = T("Audit system type row.", "감사용 시스템 타입 행입니다."),
				ParameterSummary = "@system-detail-v1\n@section\tidentity\n@row\tidentity\tClass\tRailingType\n@row\tidentity\tCategory\tRailings\n@row\tidentity\tType\tAUDIT_GUARDRAIL\n@section\tcomponents\n@component\tcomponents\tTopRailType\tElementReference\tAUDIT_TOP_RAIL\tAUDIT_TOP_RAIL\tTopRailType / Railings\tRailingType/TopRailType\n@component\tcomponents\tPrimaryHandrailType\tElementReference\tAUDIT_HANDRAIL\tAUDIT_HANDRAIL\tHandRailType / Railings\tRailingType/PrimaryHandrailType\n@component\tcomponents\tBalusterPlacement.PrimaryPattern[0]\tElementReference\tAUDIT_BALUSTER|Standard\tAUDIT_BALUSTER / Standard\tFamilySymbol / Balusters\tRailingType/BalusterPlacement/PrimaryPattern[0]\n@component\tcomponents\tRail Height\tLength\t3\t914.4 mm\tRailingType / Railings\tRailingType/Height\n@row\tcomponents\tTopRailType\tAUDIT_TOP_RAIL | TopRailType / Railings\n@row\tcomponents\tPrimaryHandrailType\tAUDIT_HANDRAIL | HandRailType / Railings\n@row\tcomponents\tBalusterPlacement.PrimaryPattern[0]\tAUDIT_BALUSTER / Standard | FamilySymbol / Balusters\n@section\tcomponent-differences\n@component-diff\tcomponent-differences\tBalusterPlacement.PrimaryPattern[0]\tElementReference\tAUDIT_BALUSTER|Standard\tAUDIT_BALUSTER / Standard\tElementReference\tAUDIT_BALUSTER|Light\tAUDIT_BALUSTER / Light\tRailingType/BalusterPlacement/PrimaryPattern[0]\n@component-diff\tcomponent-differences\tBaluster Offset\tLength\t0.25\t76.2 mm\tLength\t0.5\t152.4 mm\tRailingType/BalusterOffset\n@row\tcomponent-differences\tBalusterPlacement.PrimaryPattern[0]\tStandard: AUDIT_BALUSTER / Standard | Project: AUDIT_BALUSTER / Light | referenced type differs",
				LayerSummary = "Top rail / handrail / baluster"
			},
			new SystemRow
			{
				Status = T("Review", "검토"),
				RawStatus = "DifferentFromStandard",
				DisciplineKey = "Mechanical",
				DisciplineLabel = T("Mechanical", "설비"),
				Name = "AUDIT_CURTAIN_WALL",
				Category = "Walls",
				SystemFamilyKind = "WallType",
				Action = T("Review", "검토"),
				Notes = T("Audit system type row.", "감사용 시스템 타입 행입니다."),
				ParameterSummary = "@system-detail-v1\n@section\tidentity\n@row\tidentity\tClass\tWallType\n@row\tidentity\tCategory\tWalls\n@row\tidentity\tType\tAUDIT_CURTAIN_WALL\n@section\tcurtain-components\n@component\tcurtain-components\tDefault Curtain Panel\tElementReference\tAUDIT_SYSTEM_PANEL|Glazed\tAUDIT_SYSTEM_PANEL / Glazed\tFamilySymbol / Curtain Panels\tCurtainWall/DefaultPanel\n@component\tcurtain-components\tDependent Panel Type\tElementReference\tAUDIT_PANEL_SUPPORT|Standard\tAUDIT_PANEL_SUPPORT / Standard\tFamilySymbol / Generic Models\tCurtainWall/DefaultPanel/Support\n@component\tcurtain-components\tPanel Width\tLength\t4\t1219.2 mm\tPanelType / Curtain Panels\tCurtainWall/DefaultPanel/TypeParameters/Width\n@row\tcurtain-components\tDefault Curtain Panel\tAUDIT_SYSTEM_PANEL / Glazed | FamilySymbol / Curtain Panels\n@row\tcurtain-components\tDependent Panel Type\tAUDIT_PANEL_SUPPORT / Standard | FamilySymbol / Generic Models\n@section\tcurtain-component-differences\n@component-diff\tcurtain-component-differences\tDefault Curtain Panel\tElementReference\tAUDIT_SYSTEM_PANEL|Glazed\tAUDIT_SYSTEM_PANEL / Glazed\tElementReference\tAUDIT_SYSTEM_PANEL|Solid\tAUDIT_SYSTEM_PANEL / Solid\tCurtainWall/DefaultPanel\n@component-diff\tcurtain-component-differences\tPanel Width\tLength\t4\t1219.2 mm\tLength\t5\t1524 mm\tCurtainWall/DefaultPanel/TypeParameters/Width\n@row\tcurtain-component-differences\tDefault Curtain Panel\tStandard: AUDIT_SYSTEM_PANEL / Glazed | Current: AUDIT_SYSTEM_PANEL / Solid | referenced type differs",
				LayerSummary = "Curtain panel dependencies"
			},
			new SystemRow
			{
				Status = T("Review", "검토"),
				RawStatus = "DifferentFromStandard",
				DisciplineKey = "Mechanical",
				DisciplineLabel = T("Mechanical", "설비"),
				Name = "AUDIT_SYSTEM_PANEL_TYPE",
				Category = "Curtain Panels",
				SystemFamilyKind = "PanelType",
				Action = T("Review", "검토"),
				Notes = T("Audit system type row.", "감사용 시스템 타입 행입니다."),
				ParameterSummary = "@system-detail-v1\n@section\tidentity\n@row\tidentity\tClass\tPanelType\n@row\tidentity\tCategory\tCurtain Panels\n@row\tidentity\tType\tAUDIT_SYSTEM_PANEL_TYPE\n@section\tcurtain-components\n@component\tcurtain-components\tCurtain Panel Type\tElementReference\tAUDIT_SYSTEM_PANEL_TYPE|Glazed\tAUDIT_SYSTEM_PANEL_TYPE / Glazed\tPanelType / Curtain Panels\tPanelType\n@component\tcurtain-components\tDependent Panel Family\tElementReference\tAUDIT_PANEL_INSERT|Standard\tAUDIT_PANEL_INSERT / Standard\tFamilySymbol / Curtain Panels\tPanelType/Insert\n@component\tcurtain-components\tPanel Thickness\tLength\t0.25\t76.2 mm\tPanelType / Curtain Panels\tPanelType/TypeParameters/Thickness\n@row\tcurtain-components\tCurtain Panel Type\tAUDIT_SYSTEM_PANEL_TYPE / Glazed | PanelType / Curtain Panels\n@row\tcurtain-components\tDependent Panel Family\tAUDIT_PANEL_INSERT / Standard | FamilySymbol / Curtain Panels\n@section\tcurtain-component-differences\n@component-diff\tcurtain-component-differences\tDependent Panel Family\tElementReference\tAUDIT_PANEL_INSERT|Standard\tAUDIT_PANEL_INSERT / Standard\tElementReference\tAUDIT_PANEL_INSERT|Light\tAUDIT_PANEL_INSERT / Light\tPanelType/Insert\n@component-diff\tcurtain-component-differences\tPanel Thickness\tLength\t0.25\t76.2 mm\tLength\t0.5\t152.4 mm\tPanelType/TypeParameters/Thickness\n@row\tcurtain-component-differences\tDependent Panel Family\tStandard: AUDIT_PANEL_INSERT / Standard | Current: AUDIT_PANEL_INSERT / Light | referenced type differs",
				LayerSummary = "Direct curtain panel dependencies"
			},
			new SystemRow
			{
				Status = T("Review", "검토"),
				RawStatus = "DifferentFromStandard",
				DisciplineKey = "FireProtection",
				DisciplineLabel = T("Fire Protection", "소방"),
				Name = "AUDIT_SPRINKLER_WET",
				Category = "Piping System",
				SystemFamilyKind = "Pipe",
				Action = T("Review", "검토"),
				Notes = T("Audit review system row.", "감사용 검토 시스템 행입니다."),
				LayerSummary = "Wet / Standard"
			},
			new SystemRow
			{
				Status = T("Matches Standard", "표준과 일치"),
				RawStatus = "LoadedLatest",
				DisciplineKey = "Mechanical",
				DisciplineLabel = T("Mechanical", "설비"),
				Name = "AUDIT_RETURN_AIR",
				Category = "Duct System",
				SystemFamilyKind = "Duct",
				Action = T("No action", "작업 없음"),
				Notes = T("Audit matched system row.", "감사용 표준 일치 시스템 행입니다."),
				LayerSummary = "Return Air / Standard"
			},
			new SystemRow
			{
				Status = T("Duplicate Risk", "중복 위험"),
				RawStatus = "DuplicateRisk",
				DisciplineKey = "Electrical",
				DisciplineLabel = T("Electrical", "전기"),
				Name = "AUDIT_POWER_SYSTEM",
				Category = "Electrical System",
				SystemFamilyKind = "Electrical",
				Action = T("Review", "검토"),
				Notes = T("Audit duplicate-risk system row.", "감사용 중복 위험 시스템 행입니다."),
				LayerSummary = "Power / Duplicate risk"
			},
			new SystemRow
			{
				Status = T("Needs Review", "검토 대상"),
				RawStatus = "ManualReview",
				DisciplineKey = "Structure",
				DisciplineLabel = T("Structure", "구조"),
				Name = "AUDIT_REVIEW_SYSTEM",
				Category = "Analytical System",
				SystemFamilyKind = "Other",
				Action = T("Review", "검토"),
				Notes = T("Audit needs-review system row.", "감사용 검토 대상 시스템 행입니다."),
				LayerSummary = "Review required"
			}
		};
	}

	private static List<UnregisteredProjectItemRow> BuildAuditUnregisteredRows(string kind)
	{
		return new List<UnregisteredProjectItemRow>
		{
			new UnregisteredProjectItemRow
			{
				ItemKind = kind,
				Name = kind + "_AUDIT_PROJECT_ONLY",
				CategoryName = kind == "System" ? "Piping System" : "Mechanical Equipment",
				TypeClassName = kind == "System" ? "Pipe" : string.Empty,
				Notes = "Audit project-only item",
				Source = "UI Audit"
			}
		};
	}

	private static List<FamilyBrowserRequestRecord> BuildAuditRequests()
	{
		string now = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		return new List<FamilyBrowserRequestRecord>
		{
			new FamilyBrowserRequestRecord
			{
				RequestId = "AUDIT-REQ-001",
				RequestKind = "FAMILY",
				Status = "Submitted",
				CreatedAtUtc = now,
				UpdatedAtUtc = now,
				CreatedBy = "KKY_UI_AUDIT_MODELER",
				LastUpdatedBy = "KKY_UI_AUDIT_MODELER",
				ProjectTitle = "UI Audit Project",
				ItemName = "AUDIT_SUPPLY_DIFFUSER",
				CategoryName = "Mechanical Equipment",
				Discipline = "Mechanical",
				Reason = "Audit request row"
			},
			new FamilyBrowserRequestRecord
			{
				RequestId = "AUDIT-REQ-002",
				RequestKind = "SYSTEM",
				Status = "Reviewing",
				CreatedAtUtc = now,
				UpdatedAtUtc = now,
				CreatedBy = "KKY_UI_AUDIT_MODELER",
				LastUpdatedBy = "KKY_UI_AUDIT_ADMIN",
				ProjectTitle = "UI Audit Project",
				ItemName = "AUDIT_SUPPLY_AIR",
				CategoryName = "Duct System",
				Discipline = "Mechanical",
				Reason = "Audit in-progress row"
			}
		};
	}

	private FamilyBrowserOperationalReadinessReport BuildAuditReadinessReport(FamilyBrowserDashboardAuditScenario scenario)
	{
		FamilyBrowserOperationalReadinessReport report = new FamilyBrowserOperationalReadinessReport
		{
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			PolicyPath = "audit://policy.json",
			StandardTarget = "Mechanical"
		};
		if (!scenario.StandardRvtRegistered)
		{
			report.BlockingCount = 1;
			report.Items.Add(new FamilyBrowserOperationalReadinessItem
			{
				Area = "Standard RVT",
				Status = "Blocked",
				Message = "Audit scenario has no registered standard RVT.",
				Action = "Register Standard RVT"
			});
			return report;
		}
		if (!scenario.StandardListRegistered || scenario.IncludeReadinessWarning)
		{
			report.WarningCount = 1;
			report.Items.Add(new FamilyBrowserOperationalReadinessItem
			{
				Area = "Standard List",
				Status = "Warning",
				Message = "Audit scenario standard list needs review.",
				Action = "Register Standard Families"
			});
			return report;
		}
		report.ReadyCount = 3;
		report.Items.Add(new FamilyBrowserOperationalReadinessItem
		{
			Area = "Policy",
			Status = "Ready",
			Message = "Audit policy is ready.",
			Action = string.Empty
		});
		return report;
	}
}
