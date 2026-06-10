using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Microsoft.VisualBasic.CompilerServices;

public sealed class SystemTypePreflightBuilderService
{
	private sealed class ReferenceMappingCheck
	{
		public Element SourceReference { get; set; }

		public Element TargetReference { get; set; }

		public string Blocker { get; set; }

		public ReferenceMappingCheck()
		{
			Blocker = string.Empty;
		}
	}

	private sealed class EffectiveRoutingRule
	{
		public int SourceIndex { get; set; }

		public RoutingPreferenceRule Rule { get; set; }
	}

	private sealed class NonFamilyRoutingPartCheck
	{
		public string Action { get; set; }

		public string Status { get; set; }

		public string Notes { get; set; }

		public bool IsBlocking { get; set; }

		public NonFamilyRoutingPartCheck()
		{
			Action = string.Empty;
			Status = string.Empty;
			Notes = string.Empty;
		}

		public static NonFamilyRoutingPartCheck Ready(string actionName, string statusName, string notes)
		{
			return new NonFamilyRoutingPartCheck
			{
				Action = (actionName ?? string.Empty),
				Status = (statusName ?? string.Empty),
				Notes = (notes ?? string.Empty),
				IsBlocking = false
			};
		}

		public static NonFamilyRoutingPartCheck Block(string actionName, string statusName, string notes)
		{
			return new NonFamilyRoutingPartCheck
			{
				Action = (actionName ?? string.Empty),
				Status = (statusName ?? string.Empty),
				Notes = (notes ?? string.Empty),
				IsBlocking = true
			};
		}
	}

	private SystemTypePreflightBuilderService()
	{
	}

	public static SystemTypePreflightReport BuildReport(StandardLibraryRegistrationRecord registration, Document standardDocument, Document projectDocument, bool includeApplyDryRun = true, Action<int, int, string> progress = null)
	{
		if (registration == null)
		{
			throw new ArgumentNullException("registration");
		}
		if (standardDocument == null)
		{
			throw new ArgumentNullException("standardDocument");
		}
		if (projectDocument == null)
		{
			throw new ArgumentNullException("projectDocument");
		}
		Dictionary<string, string> standardLoadableContentFingerprintCache = new Dictionary<string, string>(StringComparer.Ordinal);
		Dictionary<string, string> projectLoadableContentFingerprintCache = new Dictionary<string, string>(StringComparer.Ordinal);
		ReportProgress(progress, 2, 100, T("Reading standard system type definitions...", "표준 시스템 타입 정의 읽는 중..."));
		checked
		{
			SystemTypeCatalogSnapshot standardCatalog = SystemTypeSemanticCaptureService.Capture(standardDocument, registration.SourceId, standardLoadableContentFingerprintCache, includeDeepLoadableContent: true, [SpecialName] (int current, int total, string message) =>
			{
				ReportProgress(progress, 2 + (int)Math.Round((double)current / (double)Math.Max(1, total) * 24.0), 100, message);
			});
			ReportProgress(progress, 28, 100, T("Reading project system type definitions...", "프로젝트 시스템 타입 정의 읽는 중..."));
			SystemTypeCatalogSnapshot projectCatalog = SystemTypeSemanticCaptureService.Capture(projectDocument, "project|" + projectDocument.Title, projectLoadableContentFingerprintCache, includeDeepLoadableContent: true, [SpecialName] (int current, int total, string message) =>
			{
				ReportProgress(progress, 28 + (int)Math.Round((double)current / (double)Math.Max(1, total) * 18.0), 100, message);
			});
			ReportProgress(progress, 48, 100, T("Reading project routing dependency families...", "프로젝트 라우팅 의존 패밀리 읽는 중..."));
			ProjectContentSnapshot projectLoadableSnapshot = ProjectSnapshotCaptureService.Capture(projectDocument, projectLoadableContentFingerprintCache, includeDeepLoadableContent: false, [SpecialName] (int current, int total, string message) =>
			{
				ReportProgress(progress, 48 + (int)Math.Round((double)current / (double)Math.Max(1, total) * 22.0), 100, message);
			});
			ReportProgress(progress, 72, 100, T("Building routing dependency catalog...", "라우팅 의존성 카탈로그 작성 중..."));
			RoutingFamilyCatalogSnapshot routingCatalog = RoutingFamilyCatalogBuilder.Build(registration.SourceId, projectLoadableSnapshot);
			ReportProgress(progress, 78, 100, T("Building system type sync plan...", "시스템 타입 동기화 계획 작성 중..."));
			SystemTypeSyncPlan syncPlan = SystemTypeSyncService.BuildPlan(standardCatalog, projectCatalog);
			List<SystemTypePreflightDiagnostic> preflightDiagnostics = new List<SystemTypePreflightDiagnostic>();
			ReportProgress(progress, 84, 100, T("Verifying same-name routing preferences...", "같은 이름 라우팅 환경설정 검증 중..."));
			NormalizeSameNameRoutingActions(syncPlan, standardDocument, projectDocument, preflightDiagnostics);
			ReportProgress(progress, 88, 100, T("Building routing dependency plan...", "라우팅 의존성 계획 작성 중..."));
			RoutingDependencyPreflightPlan dependencyPlan = RoutingDependencyPreflightService.BuildPlan(standardCatalog, routingCatalog);
			ReportProgress(progress, 92, 100, T("Building system type execution plan...", "시스템 타입 실행 계획 작성 중..."));
			SystemSyncExecutionPlan executionPlan = SystemSyncExecutionPlannerService.BuildPlan(syncPlan, dependencyPlan);
			AppendNonFamilyRoutingPartBlocks(executionPlan, standardDocument, projectDocument);
			SystemTypePreflightReport report = new SystemTypePreflightReport
			{
				GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
				StandardDisplayName = registration.DisplayName,
				ProjectDocumentTitle = projectDocument.Title,
				ProjectDocumentPath = ProjectSnapshotStore.ResolveProjectIdentityPath(projectDocument),
				StandardCatalog = standardCatalog,
				ProjectCatalog = projectCatalog,
				ProjectRoutingFamilies = routingCatalog,
				SyncPlan = syncPlan,
				DependencyPlan = dependencyPlan,
				ExecutionPlan = executionPlan,
				Diagnostics = preflightDiagnostics,
				Summary = BuildSummary(executionPlan, dependencyPlan)
			};
			if (includeApplyDryRun)
			{
				ReportProgress(progress, 96, 100, T("Running apply dry-run checks...", "적용 사전 점검 실행 중..."));
				AppendApplyDryRunBlocks(report, standardDocument, projectDocument);
			}
			report.Summary = BuildSummary(report.ExecutionPlan, report.DependencyPlan);
			ReportProgress(progress, 100, 100, T("System type review completed.", "시스템 타입 검토가 완료되었습니다."));
			return report;
		}
	}

	public static SystemTypePreflightReport BuildSelectedReport(StandardLibraryRegistrationRecord registration, Document standardDocument, Document projectDocument, string selectedCategoryName, string selectedSystemTypeName, string selectedSystemFamilyKind, bool includeApplyDryRun = false, Action<int, int, string> progress = null)
	{
		if (registration == null)
		{
			throw new ArgumentNullException("registration");
		}
		if (standardDocument == null)
		{
			throw new ArgumentNullException("standardDocument");
		}
		if (projectDocument == null)
		{
			throw new ArgumentNullException("projectDocument");
		}
		if (string.IsNullOrWhiteSpace(selectedSystemTypeName) || string.IsNullOrWhiteSpace(selectedSystemFamilyKind))
		{
			return BuildReport(registration, standardDocument, projectDocument, includeApplyDryRun, progress);
		}
		Dictionary<string, string> standardLoadableContentFingerprintCache = new Dictionary<string, string>(StringComparer.Ordinal);
		Dictionary<string, string> projectLoadableContentFingerprintCache = new Dictionary<string, string>(StringComparer.Ordinal);
		ReportProgress(progress, 4, 100, T("Reading selected standard system type...", "선택한 표준 시스템 타입 읽는 중..."));
		checked
		{
			SystemTypeCatalogSnapshot standardCatalog = SystemTypeSemanticCaptureService.CaptureSelected(standardDocument, registration.SourceId, selectedSystemFamilyKind, selectedCategoryName, selectedSystemTypeName, standardLoadableContentFingerprintCache, includeDeepLoadableContent: false, [SpecialName] (int current, int total, string message) =>
			{
				ReportProgress(progress, 4 + (int)Math.Round((double)current / (double)Math.Max(1, total) * 18.0), 100, message);
			});
			ReportProgress(progress, 26, 100, T("Reading selected project system type...", "선택한 프로젝트 시스템 타입 읽는 중..."));
			SystemTypeCatalogSnapshot projectCatalog = SystemTypeSemanticCaptureService.CaptureSelected(projectDocument, "project|" + projectDocument.Title, selectedSystemFamilyKind, selectedCategoryName, selectedSystemTypeName, projectLoadableContentFingerprintCache, includeDeepLoadableContent: false, [SpecialName] (int current, int total, string message) =>
			{
				ReportProgress(progress, 26 + (int)Math.Round((double)current / (double)Math.Max(1, total) * 14.0), 100, message);
			});
			ReportProgress(progress, 44, 100, T("Collecting selected system type dependency families...", "선택한 시스템 타입 의존 패밀리 수집 중..."));
			IEnumerable<string> dependencyFamilyNames = CollectDependencyFamilyNames(standardCatalog);
			ReportProgress(progress, 50, 100, T("Reading selected project dependency families...", "선택한 프로젝트 의존 패밀리 읽는 중..."));
			ProjectContentSnapshot projectLoadableSnapshot = ProjectSnapshotCaptureService.CaptureLoadableFamiliesByNames(projectDocument, dependencyFamilyNames, projectLoadableContentFingerprintCache, includeDeepLoadableContent: false, includeInstanceCounts: true, [SpecialName] (int current, int total, string message) =>
			{
				ReportProgress(progress, 50 + (int)Math.Round((double)current / (double)Math.Max(1, total) * 18.0), 100, message);
			});
			ReportProgress(progress, 70, 100, T("Building selected routing dependency catalog...", "선택 항목 라우팅 의존성 카탈로그 작성 중..."));
			RoutingFamilyCatalogSnapshot routingCatalog = RoutingFamilyCatalogBuilder.Build(registration.SourceId, projectLoadableSnapshot);
			ReportProgress(progress, 76, 100, T("Building selected system type sync plan...", "선택 항목 시스템 타입 동기화 계획 작성 중..."));
			SystemTypeSyncPlan syncPlan = SystemTypeSyncService.BuildPlan(standardCatalog, projectCatalog);
			List<SystemTypePreflightDiagnostic> preflightDiagnostics = new List<SystemTypePreflightDiagnostic>();
			ReportProgress(progress, 84, 100, T("Verifying selected routing preferences...", "선택 항목 라우팅 환경설정 검증 중..."));
			NormalizeSameNameRoutingActions(syncPlan, standardDocument, projectDocument, preflightDiagnostics);
			ReportProgress(progress, 88, 100, T("Building selected dependency plan...", "선택 항목 의존성 계획 작성 중..."));
			RoutingDependencyPreflightPlan dependencyPlan = RoutingDependencyPreflightService.BuildPlan(standardCatalog, routingCatalog);
			ReportProgress(progress, 92, 100, T("Building selected execution plan...", "선택 항목 실행 계획 작성 중..."));
			SystemSyncExecutionPlan executionPlan = SystemSyncExecutionPlannerService.BuildPlan(syncPlan, dependencyPlan);
			AppendNonFamilyRoutingPartBlocks(executionPlan, standardDocument, projectDocument);
			SystemTypePreflightReport report = new SystemTypePreflightReport
			{
				GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
				StandardDisplayName = registration.DisplayName,
				ProjectDocumentTitle = projectDocument.Title,
				ProjectDocumentPath = ProjectSnapshotStore.ResolveProjectIdentityPath(projectDocument),
				StandardCatalog = standardCatalog,
				ProjectCatalog = projectCatalog,
				ProjectRoutingFamilies = routingCatalog,
				SyncPlan = syncPlan,
				DependencyPlan = dependencyPlan,
				ExecutionPlan = executionPlan,
				Diagnostics = preflightDiagnostics,
				Summary = BuildSummary(executionPlan, dependencyPlan)
			};
			if (includeApplyDryRun)
			{
				ReportProgress(progress, 96, 100, T("Running selected apply dry-run checks...", "선택 항목 적용 사전 점검 실행 중..."));
				AppendApplyDryRunBlocks(report, standardDocument, projectDocument);
			}
			report.Summary = BuildSummary(report.ExecutionPlan, report.DependencyPlan);
			ReportProgress(progress, 100, 100, T("Selected system type review completed.", "선택한 시스템 타입 검토가 완료되었습니다."));
			return report;
		}
	}

	private static void ReportProgress(Action<int, int, string> progress, int current, int total, string message)
	{
		if (progress != null)
		{
			try
			{
				int safeTotal = Math.Max(1, total);
				progress(Math.Max(0, Math.Min(current, safeTotal)), safeTotal, message ?? string.Empty);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
	}

	private static string T(string englishText, string koreanText)
	{
		return FamilyBrowserLanguageService.Text(englishText, koreanText);
	}

	public static void NormalizeSameNameRoutingActions(SystemTypePreflightReport report, Document standardDocument, Document projectDocument)
	{
		if (report != null)
		{
			if (report.Diagnostics == null)
			{
				report.Diagnostics = new List<SystemTypePreflightDiagnostic>();
			}
			NormalizeSameNameRoutingActions(report.SyncPlan, standardDocument, projectDocument, report.Diagnostics);
			report.ExecutionPlan = SystemSyncExecutionPlannerService.BuildPlan(report.SyncPlan, report.DependencyPlan);
			AppendNonFamilyRoutingPartBlocks(report.ExecutionPlan, standardDocument, projectDocument);
			report.Summary = BuildSummary(report.ExecutionPlan, report.DependencyPlan);
		}
	}

	public unsafe static int EnsureRoutingFamilyDependencyItems(SystemTypePreflightReport report, Document standardDocument, Document projectDocument)
	{
		//IL_01d7: Unknown result type (might be due to invalid IL or missing references)
		//IL_01dc: Unknown result type (might be due to invalid IL or missing references)
		//IL_01de: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e1: Invalid comparison between Unknown and I4
		//IL_01e9: Unknown result type (might be due to invalid IL or missing references)
		//IL_024e: Unknown result type (might be due to invalid IL or missing references)
		if (report == null || standardDocument == null || projectDocument == null)
		{
			return 0;
		}
		if (report.SyncPlan == null)
		{
			report.SyncPlan = new SystemTypeSyncPlan();
		}
		if (report.DependencyPlan == null)
		{
			report.DependencyPlan = new RoutingDependencyPreflightPlan();
		}
		if (report.DependencyPlan.Items == null)
		{
			report.DependencyPlan.Items = new List<RoutingDependencyPreflightItem>();
		}
		if (report.Diagnostics == null)
		{
			report.Diagnostics = new List<SystemTypePreflightDiagnostic>();
		}
		HashSet<string> existingKeys = new HashSet<string>(from x in report.DependencyPlan.Items
			where x != null
			select BuildRoutingDependencyPreflightKey(x.SystemFamilyKind, x.SystemTypeName, x.SourceLibraryFamilyId, x.SourceFamilyName, x.SourceTypeName), StringComparer.Ordinal);
		int addedCount = 0;
		foreach (SystemTypeSyncPlanItem item in report.SyncPlan.Items ?? new List<SystemTypeSyncPlanItem>())
		{
			if (item == null || !SystemTypeSupportPolicyService.RequiresDependencyRefresh(item.SystemFamilyKind))
			{
				continue;
			}
			switch (Normalize(item.Action))
			{
			case "createmissingtype":
			case "overwritedestination":
			case "consolidateduplicatesuffixtypes":
			{
				ElementType sourceType = FindSystemType(standardDocument, item.SystemFamilyKind, item.CategoryName, item.SourceTypeName);
				if (sourceType == null)
				{
					AddRoutingDependencyDiagnostic(report.Diagnostics, item, "CannotVerifyRouting", "The source system type was not found while ensuring routing dependency items.");
					break;
				}
				RoutingPreferenceManager manager = TryGetRoutingPreferenceManager(sourceType);
				if (manager == null)
				{
					AddRoutingDependencyDiagnostic(report.Diagnostics, item, "CannotVerifyRouting", "The source system type has no readable routing preference manager.");
					break;
				}
				foreach (RoutingPreferenceRuleGroupType group in Enum.GetValues(typeof(RoutingPreferenceRuleGroupType)).Cast<RoutingPreferenceRuleGroupType>())
				{
					if ((int)group == -1)
					{
						continue;
					}
					int ruleCount;
					try
					{
						ruleCount = manager.GetNumberOfRules(group);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						AddRoutingDependencyDiagnostic(report.Diagnostics, item, "CannotVerifyRouting", "Could not read routing rule count. group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / " + ex2.Message);
						ProjectData.ClearProjectError();
						continue;
					}
					int num = checked(ruleCount - 1);
					for (int index = 0; index <= num; index = checked(index + 1))
					{
						RoutingPreferenceRule rule = null;
						try
						{
							rule = manager.GetRule(group, index);
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							AddRoutingDependencyDiagnostic(report.Diagnostics, item, "CannotVerifyRouting", "Could not read routing rule. group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " index=" + index.ToString(CultureInfo.InvariantCulture) + " / " + ex4.Message);
							ProjectData.ClearProjectError();
							continue;
						}
						if (!RoutingRuleHasMappablePart(rule))
						{
							continue;
						}
						Element element = standardDocument.GetElement(rule.MEPPartId);
						FamilySymbol sourceSymbol = (FamilySymbol)(object)((element is FamilySymbol) ? element : null);
						if (sourceSymbol != null && sourceSymbol.Family != null)
						{
							string sourceFamilyName = (((Element)sourceSymbol.Family).Name ?? string.Empty).Trim();
							string sourceTypeName = ResolveElementName((Element)(object)sourceSymbol);
							string sourceFamilyKey = RoutingFamilyCatalogBuilder.BuildFamilyKey(ResolveCategoryName((Element)(object)sourceSymbol), sourceFamilyName);
							string dependencyKey = BuildRoutingDependencyPreflightKey(item.SystemFamilyKind, item.SourceTypeName, sourceFamilyKey, sourceFamilyName, sourceTypeName);
							if (!existingKeys.Contains(dependencyKey))
							{
								FamilySymbol targetSymbol = FindTargetFamilySymbol(projectDocument, sourceSymbol);
								string dependencyAction = ((targetSymbol == null) ? "LoadMissingDependencyFamily" : "ReuseLoadedDependency");
								string dependencyReason = ((targetSymbol == null) ? "Routing preference family/type is required by the selected standard system type and is not loaded in the current project." : "Routing preference family/type is already loaded in the current project.");
								report.DependencyPlan.Items.Add(new RoutingDependencyPreflightItem
								{
									SystemFamilyKind = item.SystemFamilyKind,
									SystemTypeName = item.SourceTypeName,
									DependencyRole = ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString(),
									SourceLibraryFamilyId = sourceFamilyKey,
									SourceFamilyName = sourceFamilyName,
									SourceTypeName = sourceTypeName,
									SourceTypeFingerprint = SystemTypeFingerprintService.ComputeSimpleTypeFingerprint(sourceFamilyKey, sourceTypeName),
									Action = dependencyAction,
									Reason = dependencyReason
								});
								existingKeys.Add(dependencyKey);
								addedCount = checked(addedCount + 1);
								report.Diagnostics.Add(new SystemTypePreflightDiagnostic
								{
									Stage = "EnsureRoutingFamilyDependencyItems",
									PlannedAction = (item.Action ?? string.Empty),
									NormalizedAction = dependencyAction,
									SystemTypeName = (item.SourceTypeName ?? string.Empty),
									SystemFamilyKind = (item.SystemFamilyKind ?? string.Empty),
									Reason = ((targetSymbol == null) ? "RoutingDependencyMissing" : "RoutingDependencyVerifiedLoaded"),
									Details = "group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " index=" + index.ToString(CultureInfo.InvariantCulture) + " family=" + sourceFamilyName + " type=" + sourceTypeName
								});
							}
						}
					}
				}
				break;
			}
			}
		}
		if (addedCount > 0)
		{
			report.ExecutionPlan = SystemSyncExecutionPlannerService.BuildPlan(report.SyncPlan, report.DependencyPlan);
			AppendNonFamilyRoutingPartBlocks(report.ExecutionPlan, standardDocument, projectDocument);
			report.Summary = BuildSummary(report.ExecutionPlan, report.DependencyPlan);
		}
		return addedCount;
	}

	private static void AddRoutingDependencyDiagnostic(ICollection<SystemTypePreflightDiagnostic> diagnostics, SystemTypeSyncPlanItem item, string reason, string details)
	{
		if (diagnostics != null && item != null)
		{
			diagnostics.Add(new SystemTypePreflightDiagnostic
			{
				Stage = "EnsureRoutingFamilyDependencyItems",
				PlannedAction = (item.Action ?? string.Empty),
				NormalizedAction = (item.Action ?? string.Empty),
				SystemTypeName = (item.SourceTypeName ?? string.Empty),
				SystemFamilyKind = (item.SystemFamilyKind ?? string.Empty),
				Reason = (reason ?? string.Empty),
				Details = (details ?? string.Empty)
			});
		}
	}

	private static string BuildRoutingDependencyPreflightKey(string systemFamilyKind, string systemTypeName, string sourceLibraryFamilyId, string sourceFamilyName, string sourceTypeName)
	{
		return Normalize(systemFamilyKind) + "|" + Normalize(systemTypeName) + "|" + Normalize(sourceLibraryFamilyId) + "|" + Normalize(sourceFamilyName) + "|" + Normalize(sourceTypeName);
	}

	private static void NormalizeSameNameRoutingActions(SystemTypeSyncPlan syncPlan, Document standardDocument, Document projectDocument, ICollection<SystemTypePreflightDiagnostic> diagnostics)
	{
		if (syncPlan == null || syncPlan.Items == null || standardDocument == null || projectDocument == null)
		{
			return;
		}
		foreach (SystemTypeSyncPlanItem item in syncPlan.Items)
		{
			if (item == null || !SystemTypeSupportPolicyService.RequiresDependencyRefresh(item.SystemFamilyKind))
			{
				continue;
			}
			string normalizedAction = Normalize(item.Action);
			if (!string.Equals(normalizedAction, "keepdestination", StringComparison.Ordinal) && !string.Equals(normalizedAction, "consolidateduplicatesuffixtypes", StringComparison.Ordinal))
			{
				continue;
			}
			string plannedAction = CanonicalSystemTypeActionName(item.Action);
			string normalizedActionName = plannedAction;
			ElementType sourceType = FindSystemType(standardDocument, item.SystemFamilyKind, item.CategoryName, item.SourceTypeName);
			ElementType targetType = FindSystemType(projectDocument, item.SystemFamilyKind, item.CategoryName, item.SourceTypeName);
			string diagnosticReason = string.Empty;
			string diagnosticDetails = string.Empty;
			if (TargetSystemTypeMatchesStandardRoutingForPreflight(projectDocument, standardDocument, sourceType, targetType, ref diagnosticReason, ref diagnosticDetails))
			{
				normalizedActionName = ((string.Equals(normalizedAction, "keepdestination", StringComparison.Ordinal) && item.RelatedDuplicateNames != null && item.RelatedDuplicateNames.Count > 0) ? "ConsolidateDuplicateSuffixTypes" : plannedAction);
				if (!string.Equals(item.Action, normalizedActionName, StringComparison.Ordinal))
				{
					item.Action = normalizedActionName;
				}
				AddPreflightNormalizationDiagnostic(diagnostics, plannedAction, normalizedActionName, item, diagnosticReason, diagnosticDetails);
			}
			else
			{
				item.Action = "OverwriteDestination";
				normalizedActionName = "OverwriteDestination";
				item.Reason = "Same-name routing preference cannot be proven to match the registered standard, so this type must be overwritten before apply. " + diagnosticDetails;
				item.DiffSummary.Add("Routing preference live comparison failed: " + diagnosticDetails);
				AddPreflightNormalizationDiagnostic(diagnostics, plannedAction, normalizedActionName, item, diagnosticReason, diagnosticDetails);
			}
		}
	}

	private unsafe static bool TargetSystemTypeMatchesStandardRoutingForPreflight(Document targetDocument, Document standardDocument, ElementType sourceType, ElementType targetType, ref string diagnosticReason, ref string diagnosticDetails)
	{
		//IL_00a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ad: Invalid comparison between Unknown and I4
		//IL_00b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bd: Unknown result type (might be due to invalid IL or missing references)
		diagnosticReason = "CannotVerifyRouting";
		diagnosticDetails = string.Empty;
		bool TargetSystemTypeMatchesStandardRoutingForPreflight;
		if (targetDocument == null || standardDocument == null)
		{
			diagnosticDetails = "Document context is unavailable for routing comparison.";
			TargetSystemTypeMatchesStandardRoutingForPreflight = false;
		}
		else if (sourceType == null)
		{
			diagnosticDetails = "The standard system type could not be resolved for live routing comparison.";
			TargetSystemTypeMatchesStandardRoutingForPreflight = false;
		}
		else if (targetType == null)
		{
			diagnosticDetails = "The target system type does not exist for live routing comparison.";
			TargetSystemTypeMatchesStandardRoutingForPreflight = false;
		}
		else
		{
			try
			{
				RoutingPreferenceManager sourceManager = TryGetRoutingPreferenceManager(sourceType);
				RoutingPreferenceManager targetManager = TryGetRoutingPreferenceManager(targetType);
				if (sourceManager == null && targetManager == null)
				{
					diagnosticDetails = "Routing preference manager is unavailable, so the target cannot be proven to match the standard.";
					TargetSystemTypeMatchesStandardRoutingForPreflight = false;
				}
				else if (sourceManager == null || targetManager == null)
				{
					diagnosticDetails = "Routing preference manager availability differs between standard and target.";
					TargetSystemTypeMatchesStandardRoutingForPreflight = false;
				}
				else
				{
					foreach (RoutingPreferenceRuleGroupType group in Enum.GetValues(typeof(RoutingPreferenceRuleGroupType)).Cast<RoutingPreferenceRuleGroupType>())
					{
						if ((int)group == -1)
						{
							continue;
						}
						List<EffectiveRoutingRule> sourceRules = BuildEffectiveRoutingRules(sourceManager, group);
						List<EffectiveRoutingRule> targetRules = BuildEffectiveRoutingRules(targetManager, group);
						if (sourceRules.Count == targetRules.Count)
						{
							int num = checked(sourceRules.Count - 1);
							int index = 0;
							while (index <= num)
							{
								RoutingPreferenceRule sourceRule = sourceRules[index].Rule;
								RoutingPreferenceRule targetRule = targetRules[index].Rule;
								if (sourceRule == null || targetRule == null)
								{
									diagnosticReason = "CannotVerifyRouting";
									diagnosticDetails = "A routing rule could not be read. group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " index=" + index.ToString(CultureInfo.InvariantCulture) + " sourceIndex=" + sourceRules[index].SourceIndex.ToString(CultureInfo.InvariantCulture) + " targetIndex=" + targetRules[index].SourceIndex.ToString(CultureInfo.InvariantCulture);
									TargetSystemTypeMatchesStandardRoutingForPreflight = false;
								}
								else
								{
									Element sourcePart = ((sourceRule.MEPPartId == null || sourceRule.MEPPartId == ElementId.InvalidElementId) ? null : standardDocument.GetElement(sourceRule.MEPPartId));
									if (sourceRule.MEPPartId != null && sourceRule.MEPPartId != ElementId.InvalidElementId && sourcePart == null)
									{
										diagnosticReason = "CannotVerifyRouting";
										diagnosticDetails = "A standard routing rule part could not be resolved. group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " index=" + index.ToString(CultureInfo.InvariantCulture) + " sourceIndex=" + sourceRules[index].SourceIndex.ToString(CultureInfo.InvariantCulture) + " sourcePartId=" + FormatElementId(sourceRule.MEPPartId);
										TargetSystemTypeMatchesStandardRoutingForPreflight = false;
									}
									else
									{
										ElementId expectedTargetPartId = ResolveExpectedRoutingPartIdForPreflight(targetDocument, standardDocument, sourceRule.MEPPartId);
										if (sourcePart != null && expectedTargetPartId == ElementId.InvalidElementId)
										{
											diagnosticReason = "CannotVerifyRouting";
											diagnosticDetails = "A routing rule mapped part could not be resolved in the target. group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " index=" + index.ToString(CultureInfo.InvariantCulture) + " sourceIndex=" + sourceRules[index].SourceIndex.ToString(CultureInfo.InvariantCulture) + " sourcePart=" + ((object)sourcePart).GetType().Name + ":" + ResolveElementName(sourcePart);
											TargetSystemTypeMatchesStandardRoutingForPreflight = false;
										}
										else if (!ElementIdsEqual(expectedTargetPartId, targetRule.MEPPartId))
										{
											diagnosticReason = "RoutingPreferenceMismatch";
											diagnosticDetails = "A routing rule mapped part differs. group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " index=" + index.ToString(CultureInfo.InvariantCulture) + " sourceIndex=" + sourceRules[index].SourceIndex.ToString(CultureInfo.InvariantCulture) + " targetIndex=" + targetRules[index].SourceIndex.ToString(CultureInfo.InvariantCulture) + " expected=" + FormatElementId(expectedTargetPartId) + " actual=" + FormatElementId(targetRule.MEPPartId);
											TargetSystemTypeMatchesStandardRoutingForPreflight = false;
										}
										else
										{
											Element targetPart = ((targetRule.MEPPartId == null || targetRule.MEPPartId == ElementId.InvalidElementId) ? null : targetDocument.GetElement(targetRule.MEPPartId));
											if (sourcePart != null && !RoutingPartDefinitionMatchesForPreflight(standardDocument, sourcePart, targetDocument, targetPart))
											{
												diagnosticReason = "RoutingPreferenceMismatch";
												diagnosticDetails = "A routing part definition differs. group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " index=" + index.ToString(CultureInfo.InvariantCulture) + " sourceIndex=" + sourceRules[index].SourceIndex.ToString(CultureInfo.InvariantCulture) + " part=" + ((object)sourcePart).GetType().Name + ":" + ResolveElementName(sourcePart);
												TargetSystemTypeMatchesStandardRoutingForPreflight = false;
											}
											else
											{
												if (string.Equals(sourceRule.Description ?? string.Empty, targetRule.Description ?? string.Empty, StringComparison.Ordinal) && ResolveCriterionCount(sourceRule) == ResolveCriterionCount(targetRule) && string.Equals(BuildRoutingCriteriaSignature(sourceRule), BuildRoutingCriteriaSignature(targetRule), StringComparison.Ordinal))
												{
													index = checked(index + 1);
													continue;
												}
												diagnosticReason = "RoutingPreferenceMismatch";
												diagnosticDetails = "Routing rule description or criteria differ. group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " index=" + index.ToString(CultureInfo.InvariantCulture) + " sourceIndex=" + sourceRules[index].SourceIndex.ToString(CultureInfo.InvariantCulture) + " targetIndex=" + targetRules[index].SourceIndex.ToString(CultureInfo.InvariantCulture);
												TargetSystemTypeMatchesStandardRoutingForPreflight = false;
											}
										}
									}
								}
								goto end_IL_004a;
							}
							continue;
						}
						diagnosticReason = "RoutingPreferenceMismatch";
						diagnosticDetails = "Effective routing rule count differs. group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " source=" + sourceRules.Count.ToString(CultureInfo.InvariantCulture) + " target=" + targetRules.Count.ToString(CultureInfo.InvariantCulture);
						TargetSystemTypeMatchesStandardRoutingForPreflight = false;
						goto end_IL_004a;
					}
					diagnosticReason = "RoutingVerifiedEqual";
					diagnosticDetails = "Routing preference comparison matched the standard.";
					TargetSystemTypeMatchesStandardRoutingForPreflight = true;
				}
				end_IL_004a:;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				diagnosticReason = "CannotVerifyRouting";
				diagnosticDetails = "Live routing comparison failed: " + ex2.Message;
				TargetSystemTypeMatchesStandardRoutingForPreflight = false;
				ProjectData.ClearProjectError();
			}
		}
		return TargetSystemTypeMatchesStandardRoutingForPreflight;
	}

	private static ElementId ResolveExpectedRoutingPartIdForPreflight(Document targetDocument, Document standardDocument, ElementId sourcePartId)
	{
		if (sourcePartId == null || sourcePartId == ElementId.InvalidElementId)
		{
			return ElementId.InvalidElementId;
		}
		Element sourcePart = standardDocument.GetElement(sourcePartId);
		if (sourcePart == null)
		{
			return ElementId.InvalidElementId;
		}
		FamilySymbol sourceSymbol = (FamilySymbol)(object)((sourcePart is FamilySymbol) ? sourcePart : null);
		if (sourceSymbol != null)
		{
			FamilySymbol targetSymbol = FindTargetFamilySymbol(targetDocument, sourceSymbol);
			return (targetSymbol == null) ? ElementId.InvalidElementId : ((Element)targetSymbol).Id;
		}
		PipeSegment sourceSegment = (PipeSegment)(object)((sourcePart is PipeSegment) ? sourcePart : null);
		if (sourceSegment != null)
		{
			ReferenceMappingCheck materialCheck = ResolvePotentialMappedReference(targetDocument, standardDocument, ((Segment)sourceSegment).MaterialId, "pipe segment material");
			ReferenceMappingCheck scheduleCheck = ResolvePotentialMappedReference(targetDocument, standardDocument, sourceSegment.ScheduleTypeId, "pipe segment schedule type");
			if (materialCheck.TargetReference != null && scheduleCheck.TargetReference != null)
			{
				PipeSegment byCombination = FindPipeSegmentByMaterialAndSchedule(targetDocument, materialCheck.TargetReference.Id, scheduleCheck.TargetReference.Id);
				if (byCombination != null)
				{
					return ((Element)byCombination).Id;
				}
			}
		}
		Element targetPart = ResolveMatchingTargetElement(targetDocument, sourcePart);
		return (targetPart == null) ? ElementId.InvalidElementId : targetPart.Id;
	}

	private static bool RoutingPartDefinitionMatchesForPreflight(Document sourceDocument, Element sourcePart, Document targetDocument, Element targetPart)
	{
		if (sourcePart == null || targetPart == null)
		{
			return false;
		}
		if (sourcePart is FamilySymbol)
		{
			return targetPart is FamilySymbol;
		}
		return RoutingPartDefinitionsMatch(sourceDocument, sourcePart, targetDocument, targetPart);
	}

	private static FamilySymbol FindTargetFamilySymbol(Document targetDocument, FamilySymbol sourceSymbol)
	{
		//IL_004a: Unknown result type (might be due to invalid IL or missing references)
		if (targetDocument == null || sourceSymbol == null)
		{
			return null;
		}
		Family family = sourceSymbol.Family;
		string value = ((family != null) ? ((Element)family).Name : null) ?? string.Empty;
		string value2 = ResolveElementName((Element)(object)sourceSymbol);
		string right = ResolveCategoryName((Element)(object)sourceSymbol);
		return ((IEnumerable)new FilteredElementCollector(targetDocument).OfClass(typeof(FamilySymbol))).Cast<FamilySymbol>().FirstOrDefault([SpecialName] (FamilySymbol x) =>
		{
			if (x != null)
			{
				Family family2 = x.Family;
				if (string.Equals(Normalize(((family2 != null) ? ((Element)family2).Name : null) ?? string.Empty), Normalize(value), StringComparison.Ordinal) && string.Equals(Normalize(ResolveElementName((Element)(object)x)), Normalize(value2), StringComparison.Ordinal))
				{
					return CategoryNamesMatch(ResolveCategoryName((Element)(object)x), right);
				}
			}
			return false;
		});
	}

	private static List<EffectiveRoutingRule> BuildEffectiveRoutingRules(RoutingPreferenceManager manager, RoutingPreferenceRuleGroupType group)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		List<EffectiveRoutingRule> result = new List<EffectiveRoutingRule>();
		if (manager == null)
		{
			return result;
		}
		checked
		{
			int num = manager.GetNumberOfRules(group) - 1;
			for (int index = 0; index <= num; index++)
			{
				RoutingPreferenceRule rule = manager.GetRule(group, index);
				if (RoutingRuleHasMappablePart(rule))
				{
					result.Add(new EffectiveRoutingRule
					{
						SourceIndex = index,
						Rule = rule
					});
				}
			}
			return result;
		}
	}

	private static bool RoutingRuleHasMappablePart(RoutingPreferenceRule rule)
	{
		if (rule == null || rule.MEPPartId == null)
		{
			return false;
		}
		return rule.MEPPartId != ElementId.InvalidElementId && RevitElementIdCompat.CompatIntegerValue(rule.MEPPartId) > 0;
	}

	private static int ResolveCriterionCount(RoutingPreferenceRule rule)
	{
		int ResolveCriterionCount;
		if (rule == null)
		{
			ResolveCriterionCount = -1;
		}
		else
		{
			try
			{
				ResolveCriterionCount = rule.NumberOfCriteria;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ResolveCriterionCount = -1;
				ProjectData.ClearProjectError();
			}
		}
		return ResolveCriterionCount;
	}

	private static string BuildRoutingCriteriaSignature(RoutingPreferenceRule rule)
	{
		if (rule == null)
		{
			return string.Empty;
		}
		int count = ResolveCriterionCount(rule);
		if (count <= 0)
		{
			return string.Empty;
		}
		List<string> criteria = new List<string>();
		checked
		{
			int num = count - 1;
			for (int index = 0; index <= num; index++)
			{
				try
				{
					RoutingCriterionBase criterion = rule.GetCriterion(index);
					if (criterion == null)
					{
						criteria.Add("null");
						continue;
					}
					List<string> parts = new List<string> { ((object)criterion).GetType().Name };
					object minimumValue = RuntimeHelpers.GetObjectValue(TryGetPropertyValue(criterion, "MinimumSize"));
					if (minimumValue != null)
					{
						parts.Add("min=" + Convert.ToString(RuntimeHelpers.GetObjectValue(minimumValue), CultureInfo.InvariantCulture));
					}
					object maximumValue = RuntimeHelpers.GetObjectValue(TryGetPropertyValue(criterion, "MaximumSize"));
					if (maximumValue != null)
					{
						parts.Add("max=" + Convert.ToString(RuntimeHelpers.GetObjectValue(maximumValue), CultureInfo.InvariantCulture));
					}
					criteria.Add(string.Join(":", parts));
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					criteria.Add("read-error");
					ProjectData.ClearProjectError();
				}
			}
			return string.Join("&", criteria);
		}
	}

	private static object TryGetPropertyValue(object instance, string propertyName)
	{
		object TryGetPropertyValue;
		if (instance == null)
		{
			TryGetPropertyValue = null;
		}
		else
		{
			PropertyInfo propertyInfo = instance.GetType().GetProperty(propertyName);
			if ((object)propertyInfo == null)
			{
				TryGetPropertyValue = null;
			}
			else
			{
				try
				{
					TryGetPropertyValue = propertyInfo.GetValue(RuntimeHelpers.GetObjectValue(instance), null);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					TryGetPropertyValue = null;
					ProjectData.ClearProjectError();
				}
			}
		}
		return TryGetPropertyValue;
	}

	private static string FormatElementId(ElementId id)
	{
		if (id == null)
		{
			return "(null)";
		}
		return RevitElementIdCompat.CompatIntegerValue(id).ToString(CultureInfo.InvariantCulture);
	}

	private static void AddPreflightNormalizationDiagnostic(ICollection<SystemTypePreflightDiagnostic> diagnostics, string plannedAction, string normalizedAction, SystemTypeSyncPlanItem item, string reason, string details)
	{
		if (diagnostics != null && item != null)
		{
			SystemTypePreflightDiagnostic systemTypePreflightDiagnostic = new SystemTypePreflightDiagnostic
			{
				Stage = "NormalizeSameNameRoutingActions",
				PlannedAction = (plannedAction ?? string.Empty),
				NormalizedAction = (normalizedAction ?? string.Empty),
				SystemTypeName = (item.SourceTypeName ?? string.Empty),
				SystemFamilyKind = (item.SystemFamilyKind ?? string.Empty),
				Reason = (reason ?? string.Empty),
				Details = (details ?? string.Empty)
			};
			if (!diagnostics.Any([SpecialName] (SystemTypePreflightDiagnostic x) => x != null && string.Equals(x.Stage, systemTypePreflightDiagnostic.Stage, StringComparison.Ordinal) && string.Equals(x.PlannedAction, systemTypePreflightDiagnostic.PlannedAction, StringComparison.Ordinal) && string.Equals(x.NormalizedAction, systemTypePreflightDiagnostic.NormalizedAction, StringComparison.Ordinal) && string.Equals(x.SystemTypeName, systemTypePreflightDiagnostic.SystemTypeName, StringComparison.Ordinal) && string.Equals(x.SystemFamilyKind, systemTypePreflightDiagnostic.SystemFamilyKind, StringComparison.Ordinal) && string.Equals(x.Reason, systemTypePreflightDiagnostic.Reason, StringComparison.Ordinal) && string.Equals(x.Details, systemTypePreflightDiagnostic.Details, StringComparison.Ordinal)))
			{
				diagnostics.Add(systemTypePreflightDiagnostic);
			}
		}
	}

	private static string CanonicalSystemTypeActionName(string action)
	{
		return Normalize(action) switch
		{
			"keepdestination" => "KeepDestination", 
			"createmissingtype" => "CreateMissingType", 
			"skipmissingtype" => "SkipMissingType", 
			"overwritedestination" => "OverwriteDestination", 
			"consolidateduplicatesuffixtypes" => "ConsolidateDuplicateSuffixTypes", 
			"manualreview" => "ManualReview", 
			_ => action ?? string.Empty, 
		};
	}

	private static IEnumerable<string> CollectDependencyFamilyNames(SystemTypeCatalogSnapshot catalog)
	{
		return (from x in (catalog?.Types ?? new List<SystemTypeSemanticSnapshot>()).SelectMany([SpecialName] (SystemTypeSemanticSnapshot x) => x?.RoutingDependencies ?? new List<RoutingDependencySnapshot>())
			where x != null && !string.IsNullOrWhiteSpace(x.FamilyName)
			select x.FamilyName).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
	}

	private static void AppendApplyDryRunBlocks(SystemTypePreflightReport report, Document standardDocument, Document projectDocument)
	{
		//IL_015e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0164: Invalid comparison between Unknown and I4
		//IL_0167: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ea: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f0: Expected O, but got Unknown
		//IL_00f1: Unknown result type (might be due to invalid IL or missing references)
		if (report == null || standardDocument == null || projectDocument == null)
		{
			return;
		}
		List<SystemSyncExecutionItem> actionableItems = (from x in report.ExecutionPlan?.Items ?? new List<SystemSyncExecutionItem>()
			where x != null
			where !string.Equals(x.ExecutionStatus, "Blocked", StringComparison.OrdinalIgnoreCase)
			where SystemTypeSupportPolicyService.CanApply(x.SystemFamilyKind)
			where !IsNoOpExecutionItem(x)
			select x).ToList();
		if (actionableItems.Count == 0)
		{
			return;
		}
		SystemTypeApplyExecutionReport dryRunReport = null;
		TransactionGroup group = new TransactionGroup(projectDocument, "KKY Family Browser System Type Apply Dry Run");
		try
		{
			try
			{
				group.Start();
				dryRunReport = SystemTypeApplyExecutionService.Execute(projectDocument, standardDocument, report, string.Empty);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				foreach (SystemSyncExecutionItem item in actionableItems)
				{
					AppendBlockingStep(item, "ValidateApplyDryRun", "System type apply dry-run could not be completed. Resolve this before applying standard system types. Reason: " + ex2.Message);
				}
				ProjectData.ClearProjectError();
			}
			finally
			{
				try
				{
					if ((int)group.GetStatus() == 1)
					{
						group.RollBack();
					}
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
		finally
		{
			((IDisposable)group)?.Dispose();
		}
		if (dryRunReport == null)
		{
			return;
		}
		List<SystemTypeApplyExecutionItem> failedItems = (from x in dryRunReport.Items ?? new List<SystemTypeApplyExecutionItem>()
			where x != null
			where string.Equals(x.Outcome, "Failed", StringComparison.OrdinalIgnoreCase)
			select x).ToList();
		foreach (SystemTypeApplyExecutionItem failed in failedItems)
		{
			SystemSyncExecutionItem executionItem = FindExecutionItem(report.ExecutionPlan, failed.SystemFamilyKind, failed.CategoryName, failed.SystemTypeName);
			if (executionItem != null)
			{
				AppendBlockingStep(executionItem, "ValidateApplyDryRun", "System type apply dry-run failed before real execution. This prevents partial routing preferences or duplicate system content. Reason: " + failed.Details);
			}
		}
	}

	private static bool IsNoOpExecutionItem(SystemSyncExecutionItem item)
	{
		if (item == null)
		{
			return true;
		}
		if (!string.Equals(item.ExecutionStatus, "NoChange", StringComparison.OrdinalIgnoreCase))
		{
			return false;
		}
		return item.Steps.All([SpecialName] (SystemSyncExecutionStep x) => string.Equals(x.Action, "KeepExistingSystemType", StringComparison.OrdinalIgnoreCase) || string.Equals(x.Action, "ValidateDependency", StringComparison.OrdinalIgnoreCase));
	}

	private static SystemSyncExecutionItem FindExecutionItem(SystemSyncExecutionPlan plan, string systemFamilyKind, string categoryName, string typeName)
	{
		return (plan?.Items ?? new List<SystemSyncExecutionItem>()).FirstOrDefault([SpecialName] (SystemSyncExecutionItem x) => x != null && string.Equals(Normalize(x.SystemFamilyKind), Normalize(systemFamilyKind), StringComparison.Ordinal) && CategoryNamesMatch(x.CategoryName, categoryName) && string.Equals(Normalize(x.SystemTypeName), Normalize(typeName), StringComparison.Ordinal));
	}

	private static void AppendBlockingStep(SystemSyncExecutionItem item, string actionName, string notes)
	{
		if (item != null)
		{
			item.Steps.Add(new SystemSyncExecutionStep
			{
				SequenceNo = checked(item.Steps.Count + 1),
				Action = actionName,
				Status = "ManualReview",
				TargetKind = item.SystemFamilyKind,
				TargetName = item.SystemTypeName,
				Notes = notes
			});
			item.HasManualReview = true;
			item.RequiresApproval = false;
			item.ExecutionStatus = "Blocked";
			item.Summary = "System type apply readiness failed. Resolve the listed issue before applying this standard type.";
			if (!item.BlockingReasons.Contains(actionName + ": " + notes))
			{
				item.BlockingReasons.Add(actionName + ": " + notes);
			}
		}
	}

	private static void AppendNonFamilyRoutingPartBlocks(SystemSyncExecutionPlan executionPlan, Document standardDocument, Document projectDocument)
	{
		if (executionPlan == null || standardDocument == null || projectDocument == null)
		{
			return;
		}
		checked
		{
			foreach (SystemSyncExecutionItem executionItem in executionPlan.Items ?? new List<SystemSyncExecutionItem>())
			{
				if (executionItem == null || !SystemTypeSupportPolicyService.CanApply(executionItem.SystemFamilyKind))
				{
					continue;
				}
				ElementType sourceType = FindSystemType(standardDocument, executionItem.SystemFamilyKind, executionItem.CategoryName, executionItem.SystemTypeName);
				if (sourceType == null)
				{
					continue;
				}
				List<NonFamilyRoutingPartCheck> checks = FindNonFamilyRoutingPartChecks(standardDocument, projectDocument, sourceType);
				foreach (NonFamilyRoutingPartCheck check in (from x in checks.Where([SpecialName] (NonFamilyRoutingPartCheck x) => x != null).GroupBy([SpecialName] (NonFamilyRoutingPartCheck x) => x.Action + "|" + x.Status + "|" + x.Notes, StringComparer.OrdinalIgnoreCase)
					select x.First() into x
					orderby x.IsBlocking
					select x).ThenBy([SpecialName] (NonFamilyRoutingPartCheck x) => x.Notes, StringComparer.OrdinalIgnoreCase))
				{
					executionItem.Steps.Add(new SystemSyncExecutionStep
					{
						SequenceNo = executionItem.Steps.Count + 1,
						Action = check.Action,
						Status = check.Status,
						TargetKind = executionItem.SystemFamilyKind,
						TargetName = executionItem.SystemTypeName,
						Notes = check.Notes
					});
					if (check.IsBlocking)
					{
						executionItem.BlockingReasons.Add(check.Action + ": " + check.Notes);
						executionItem.HasManualReview = true;
						executionItem.ExecutionStatus = "Blocked";
						executionItem.Summary = "Non-family system references must match the standard before this system type can be synchronized.";
					}
				}
				List<string> issues = FindNonFamilyReferenceIssues(standardDocument, projectDocument, sourceType);
				foreach (string issue in issues.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase))
				{
					executionItem.Steps.Add(new SystemSyncExecutionStep
					{
						SequenceNo = executionItem.Steps.Count + 1,
						Action = "ReviewRoutingPart",
						Status = "ManualReview",
						TargetKind = executionItem.SystemFamilyKind,
						TargetName = executionItem.SystemTypeName,
						Notes = issue
					});
					executionItem.BlockingReasons.Add("ReviewRoutingPart: " + issue);
					executionItem.HasManualReview = true;
					executionItem.ExecutionStatus = "Blocked";
					executionItem.Summary = "Non-family system references must match the standard before this system type can be synchronized.";
				}
			}
		}
	}

	private static List<NonFamilyRoutingPartCheck> FindNonFamilyRoutingPartChecks(Document standardDocument, Document projectDocument, ElementType sourceType)
	{
		//IL_0038: Unknown result type (might be due to invalid IL or missing references)
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0042: Invalid comparison between Unknown and I4
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_006f: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bc: Unknown result type (might be due to invalid IL or missing references)
		List<NonFamilyRoutingPartCheck> checks = new List<NonFamilyRoutingPartCheck>();
		RoutingPreferenceManager manager = TryGetRoutingPreferenceManager(sourceType);
		if (manager == null)
		{
			return checks;
		}
		foreach (RoutingPreferenceRuleGroupType group in Enum.GetValues(typeof(RoutingPreferenceRuleGroupType)).Cast<RoutingPreferenceRuleGroupType>())
		{
			if ((int)group == -1)
			{
				continue;
			}
			int ruleCount;
			try
			{
				ruleCount = manager.GetNumberOfRules(group);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
				continue;
			}
			checked
			{
				int num = ruleCount - 1;
				for (int index = 0; index <= num; index++)
				{
					RoutingPreferenceRule rule = null;
					try
					{
						rule = manager.GetRule(group, index);
					}
					catch (Exception projectError2)
					{
						ProjectData.SetProjectError(projectError2);
						ProjectData.ClearProjectError();
						continue;
					}
					Element sourcePart = ((((rule != null) ? rule.MEPPartId : null) == null) ? null : standardDocument.GetElement(rule.MEPPartId));
					if (sourcePart != null && !(sourcePart is FamilySymbol))
					{
						checks.Add(BuildNonFamilyRoutingPartCheck(standardDocument, projectDocument, sourcePart, group, index));
					}
				}
			}
		}
		return checks;
	}

	private unsafe static NonFamilyRoutingPartCheck BuildNonFamilyRoutingPartCheck(Document standardDocument, Document projectDocument, Element sourcePart, RoutingPreferenceRuleGroupType group, int ruleIndex)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		PipeSegment sourceSegment = (PipeSegment)(object)((sourcePart is PipeSegment) ? sourcePart : null);
		if (sourceSegment != null)
		{
			return BuildPipeSegmentRoutingPartCheck(standardDocument, projectDocument, sourceSegment, group, ruleIndex);
		}
		Element targetPart = ResolveMatchingTargetElement(projectDocument, sourcePart);
		if (targetPart == null)
		{
			return NonFamilyRoutingPartCheck.Block("ReviewRoutingPart", "ManualReview", "A non-family routing preference part is missing and does not yet have an explicit safe ensure path. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + ruleIndex.ToString(CultureInfo.InvariantCulture) + " / Part: " + ((object)sourcePart).GetType().Name + " : " + ResolveElementName(sourcePart));
		}
		if (!RoutingPartDefinitionsMatch(standardDocument, sourcePart, projectDocument, targetPart))
		{
			return NonFamilyRoutingPartCheck.Block("ReviewRoutingPart", "ManualReview", "Routing part differs from the standard RVT. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + ruleIndex.ToString(CultureInfo.InvariantCulture) + " / Part: " + ((object)sourcePart).GetType().Name + " : " + ResolveElementName(sourcePart));
		}
		return NonFamilyRoutingPartCheck.Ready("ReuseRoutingPart", "Ready", "Routing part can be reused. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + ruleIndex.ToString(CultureInfo.InvariantCulture) + " / Part: " + ((object)sourcePart).GetType().Name + " : " + ResolveElementName(sourcePart));
	}

	private unsafe static NonFamilyRoutingPartCheck BuildPipeSegmentRoutingPartCheck(Document standardDocument, Document projectDocument, PipeSegment sourceSegment, RoutingPreferenceRuleGroupType group, int ruleIndex)
	{
		if (string.IsNullOrWhiteSpace(BuildPipeSegmentSizeSignature((Segment)(object)sourceSegment)))
		{
			return NonFamilyRoutingPartCheck.Block("ReviewPipeSegment", "ManualReview", "The standard pipe segment size table could not be read. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + ruleIndex.ToString(CultureInfo.InvariantCulture) + " / Segment: " + ResolveElementName((Element)(object)sourceSegment));
		}
		Element targetByName = ResolveMatchingTargetElement(projectDocument, (Element)(object)sourceSegment);
		if (targetByName != null && RoutingPartDefinitionsMatch(standardDocument, (Element)(object)sourceSegment, projectDocument, targetByName))
		{
			return NonFamilyRoutingPartCheck.Ready("ReusePipeSegment", "Ready", "Pipe segment can be reused. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + ruleIndex.ToString(CultureInfo.InvariantCulture) + " / Segment: " + ResolveElementName((Element)(object)sourceSegment));
		}
		ReferenceMappingCheck materialCheck = ResolvePotentialMappedReference(projectDocument, standardDocument, ((Segment)sourceSegment).MaterialId, "pipe segment material");
		if (materialCheck.SourceReference == null)
		{
			return NonFamilyRoutingPartCheck.Block("ReviewPipeSegmentMaterial", "ManualReview", (materialCheck.Blocker ?? "The standard pipe segment material could not be resolved.") + " Segment: " + ResolveElementName((Element)(object)sourceSegment));
		}
		ReferenceMappingCheck scheduleCheck = ResolvePotentialMappedReference(projectDocument, standardDocument, sourceSegment.ScheduleTypeId, "pipe segment schedule type");
		if (scheduleCheck.SourceReference == null)
		{
			return NonFamilyRoutingPartCheck.Block("ReviewPipeSegmentSchedule", "ManualReview", (scheduleCheck.Blocker ?? "The standard pipe segment schedule type could not be resolved.") + " Segment: " + ResolveElementName((Element)(object)sourceSegment));
		}
		if (materialCheck.TargetReference != null && scheduleCheck.TargetReference != null)
		{
			PipeSegment existingByCombination = FindPipeSegmentByMaterialAndSchedule(projectDocument, materialCheck.TargetReference.Id, scheduleCheck.TargetReference.Id);
			if (existingByCombination != null)
			{
				if (RoutingPartDefinitionsMatch(standardDocument, (Element)(object)sourceSegment, projectDocument, (Element)(object)existingByCombination))
				{
					return NonFamilyRoutingPartCheck.Ready("ReusePipeSegment", "Ready", "Pipe segment can reuse an existing project segment with matching material, schedule type, and size table. Standard segment: " + ResolveElementName((Element)(object)sourceSegment) + " / Project segment: " + ResolveElementName((Element)(object)existingByCombination));
				}
				return NonFamilyRoutingPartCheck.Ready("SyncPipeSegment", "Ready", "A project pipe segment already uses the standard material and schedule type, but its size table or definition differs. Admin-authoritative apply will synchronize this segment instead of creating a duplicate. Standard segment: " + ResolveElementName((Element)(object)sourceSegment) + " / Project segment: " + ResolveElementName((Element)(object)existingByCombination));
			}
		}
		if (targetByName != null)
		{
			return NonFamilyRoutingPartCheck.Ready("ResolveStalePipeSegmentName", "Ready", "A same-name pipe segment exists with a non-standard material or schedule definition. Admin-authoritative apply will route to the correct material/schedule segment and report the stale segment for cleanup. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + ruleIndex.ToString(CultureInfo.InvariantCulture) + " / Segment: " + ResolveElementName((Element)(object)sourceSegment));
		}
		return NonFamilyRoutingPartCheck.Ready("CreatePipeSegment", "Ready", "Pipe segment is missing and is planned for creation during the system type transaction. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + ruleIndex.ToString(CultureInfo.InvariantCulture) + " / Segment: " + ResolveElementName((Element)(object)sourceSegment));
	}

	private static List<string> FindNonFamilyReferenceIssues(Document standardDocument, Document projectDocument, ElementType sourceType)
	{
		//IL_0035: Unknown result type (might be due to invalid IL or missing references)
		//IL_003b: Invalid comparison between Unknown and I4
		List<string> issues = new List<string>();
		if (sourceType == null)
		{
			return issues;
		}
		foreach (Parameter parameter in ((IEnumerable)((Element)sourceType).Parameters).Cast<Parameter>())
		{
			if (parameter == null || (int)parameter.StorageType != 4 || !parameter.HasValue)
			{
				continue;
			}
			ElementId sourceId = null;
			try
			{
				sourceId = parameter.AsElementId();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
				continue;
			}
			if (sourceId == null || sourceId == ElementId.InvalidElementId || RevitElementIdCompat.CompatIntegerValue(sourceId) < 0)
			{
				continue;
			}
			Element sourceReference = standardDocument.GetElement(sourceId);
			if (sourceReference != null && !(sourceReference is FamilySymbol) && !(sourceReference is Material) && !(sourceReference is PipeScheduleType) && !(sourceReference is PipeSegment) && !IsAuthoritativeRoutingDependencyReference(sourceReference))
			{
				Element targetReference = ResolveMatchingTargetElement(projectDocument, sourceReference);
				if (targetReference != null && !RoutingPartSignatureService.Matches(standardDocument, sourceReference, projectDocument, targetReference))
				{
					issues.Add("Referenced system definition differs from the standard RVT. Parameter: " + ResolveParameterName(parameter) + " / Reference: " + ((object)sourceReference).GetType().Name + " : " + ResolveElementName(sourceReference));
				}
			}
		}
		return issues.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
	}

	private static bool IsAuthoritativeRoutingDependencyReference(Element element)
	{
		if (element == null)
		{
			return false;
		}
		return element is Material || element is PipeScheduleType || element is PipeSegment || RuntimeTypeMatches(element, "Autodesk.Revit.DB.Material", "Autodesk.Revit.DB.Plumbing.PipeScheduleType", "Autodesk.Revit.DB.Plumbing.PipeSegment");
	}

	private static bool RuntimeTypeMatches(Element element, params string[] typeFullNames)
	{
		if (element == null || typeFullNames == null || typeFullNames.Length == 0)
		{
			return false;
		}
		HashSet<string> requested = new HashSet<string>(typeFullNames, StringComparer.Ordinal);
		Type currentType = ((object)element).GetType();
		while ((object)currentType != null)
		{
			if (requested.Contains(currentType.FullName ?? string.Empty))
			{
				return true;
			}
			currentType = currentType.BaseType;
		}
		return false;
	}

	private static ElementType FindSystemType(Document doc, string systemFamilyKind, string categoryName, string typeName)
	{
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		if (doc == null)
		{
			return null;
		}
		return ((IEnumerable)new FilteredElementCollector(doc).WhereElementIsElementType()).Cast<ElementType>().FirstOrDefault([SpecialName] (ElementType x) => x != null && !(x is FamilySymbol) && string.Equals(((object)x).GetType().Name, systemFamilyKind, StringComparison.OrdinalIgnoreCase) && CategoryNamesMatch(ResolveCategoryName((Element)(object)x), categoryName) && string.Equals(Normalize(ResolveElementName((Element)(object)x)), Normalize(typeName), StringComparison.Ordinal));
	}

	private static Element ResolveMatchingTargetElement(Document targetDocument, Element sourceElement)
	{
		//IL_003a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0063: Unknown result type (might be due to invalid IL or missing references)
		if (targetDocument == null || sourceElement == null)
		{
			return null;
		}
		string name = ((object)sourceElement).GetType().Name;
		string sourceName = ResolveElementName(sourceElement);
		string sourceCategory = ResolveCategoryName(sourceElement);
		Element typeMatch = ((IEnumerable)new FilteredElementCollector(targetDocument).WhereElementIsElementType()).Cast<Element>().FirstOrDefault([SpecialName] (Element x) => ElementIdentityMatches(x, name, sourceName, sourceCategory));
		if (typeMatch != null)
		{
			return typeMatch;
		}
		return ((IEnumerable)new FilteredElementCollector(targetDocument).WhereElementIsNotElementType()).Cast<Element>().FirstOrDefault([SpecialName] (Element x) => ElementIdentityMatches(x, name, sourceName, sourceCategory));
	}

	private static bool ElementIdentityMatches(Element element, string sourceClassName, string sourceName, string sourceCategory)
	{
		if (element == null)
		{
			return false;
		}
		return string.Equals(((object)element).GetType().Name, sourceClassName, StringComparison.OrdinalIgnoreCase) && string.Equals(Normalize(ResolveElementName(element)), Normalize(sourceName), StringComparison.Ordinal) && CategoryNamesMatch(ResolveCategoryName(element), sourceCategory);
	}

	private static ReferenceMappingCheck ResolveExistingMappedReference(Document targetDocument, Document standardDocument, ElementId sourceId, string referenceLabel)
	{
		ReferenceMappingCheck result = new ReferenceMappingCheck();
		if (sourceId == null || sourceId == ElementId.InvalidElementId || RevitElementIdCompat.CompatIntegerValue(sourceId) < 0)
		{
			result.Blocker = "The standard " + referenceLabel + " reference is not valid.";
			return result;
		}
		Element sourceReference = standardDocument.GetElement(sourceId);
		if (sourceReference == null)
		{
			result.Blocker = "The standard " + referenceLabel + " reference could not be resolved.";
			return result;
		}
		result.SourceReference = sourceReference;
		Element targetReference = ResolveMatchingTargetElement(targetDocument, sourceReference);
		if (targetReference == null)
		{
			return result;
		}
		if (!RoutingPartDefinitionsMatch(standardDocument, sourceReference, targetDocument, targetReference))
		{
			result.Blocker = "A " + referenceLabel + " with the same name exists in the current project, but its definition differs from the standard RVT. Reference: " + ((object)sourceReference).GetType().Name + " : " + ResolveElementName(sourceReference);
			return result;
		}
		result.TargetReference = targetReference;
		return result;
	}

	private static ReferenceMappingCheck ResolvePotentialMappedReference(Document targetDocument, Document standardDocument, ElementId sourceId, string referenceLabel)
	{
		ReferenceMappingCheck result = new ReferenceMappingCheck();
		if (sourceId == null || sourceId == ElementId.InvalidElementId || RevitElementIdCompat.CompatIntegerValue(sourceId) < 0)
		{
			result.Blocker = "The standard " + referenceLabel + " reference is not valid.";
			return result;
		}
		Element sourceReference = standardDocument.GetElement(sourceId);
		if (sourceReference == null)
		{
			result.Blocker = "The standard " + referenceLabel + " reference could not be resolved.";
			return result;
		}
		result.SourceReference = sourceReference;
		result.TargetReference = ResolveMatchingTargetElement(targetDocument, sourceReference);
		return result;
	}

	private static bool RoutingPartDefinitionsMatch(Document sourceDocument, Element sourcePart, Document targetDocument, Element targetPart)
	{
		if (sourcePart == null || targetPart == null)
		{
			return false;
		}
		if (RoutingPartSignatureService.Matches(sourceDocument, sourcePart, targetDocument, targetPart))
		{
			return true;
		}
		PipeSegment sourceSegment = (PipeSegment)(object)((sourcePart is PipeSegment) ? sourcePart : null);
		PipeSegment targetSegment = (PipeSegment)(object)((targetPart is PipeSegment) ? targetPart : null);
		if (sourceSegment == null || targetSegment == null)
		{
			return false;
		}
		if (!string.Equals(BuildPipeSegmentSizeSignature((Segment)(object)sourceSegment), BuildPipeSegmentSizeSignature((Segment)(object)targetSegment), StringComparison.Ordinal))
		{
			return false;
		}
		ReferenceMappingCheck materialCheck = ResolveExistingMappedReference(targetDocument, sourceDocument, ((Segment)sourceSegment).MaterialId, "pipe segment material");
		if (materialCheck.TargetReference == null || !string.IsNullOrWhiteSpace(materialCheck.Blocker))
		{
			return false;
		}
		ReferenceMappingCheck scheduleCheck = ResolveExistingMappedReference(targetDocument, sourceDocument, sourceSegment.ScheduleTypeId, "pipe segment schedule type");
		if (scheduleCheck.TargetReference == null || !string.IsNullOrWhiteSpace(scheduleCheck.Blocker))
		{
			return false;
		}
		return ElementIdsEqual(materialCheck.TargetReference.Id, ((Segment)targetSegment).MaterialId) && ElementIdsEqual(scheduleCheck.TargetReference.Id, targetSegment.ScheduleTypeId);
	}

	private static PipeSegment FindPipeSegmentByMaterialAndSchedule(Document targetDocument, ElementId materialId, ElementId scheduleTypeId)
	{
		//IL_002c: Unknown result type (might be due to invalid IL or missing references)
		if (targetDocument == null || materialId == null || scheduleTypeId == null)
		{
			return null;
		}
		return ((IEnumerable)new FilteredElementCollector(targetDocument).OfClass(typeof(PipeSegment))).Cast<PipeSegment>().FirstOrDefault([SpecialName] (PipeSegment x) => x != null && ElementIdsEqual(((Segment)x).MaterialId, materialId) && ElementIdsEqual(x.ScheduleTypeId, scheduleTypeId));
	}

	private static string BuildPipeSegmentSizeSignature(Segment segment)
	{
		string BuildPipeSegmentSizeSignature;
		if (segment == null)
		{
			BuildPipeSegmentSizeSignature = string.Empty;
		}
		else
		{
			List<string> parts = new List<string>();
			try
			{
				foreach (MEPSize size in segment.GetSizes())
				{
					if (size != null)
					{
						parts.Add(size.NominalDiameter.ToString("G17", CultureInfo.InvariantCulture) + ":" + size.InnerDiameter.ToString("G17", CultureInfo.InvariantCulture) + ":" + size.OuterDiameter.ToString("G17", CultureInfo.InvariantCulture) + ":" + size.UsedInSizeLists.ToString(CultureInfo.InvariantCulture) + ":" + size.UsedInSizing.ToString(CultureInfo.InvariantCulture));
					}
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				BuildPipeSegmentSizeSignature = string.Empty;
				ProjectData.ClearProjectError();
				goto IL_014b;
			}
			BuildPipeSegmentSizeSignature = string.Join("|", parts.OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal));
		}
		goto IL_014b;
		IL_014b:
		return BuildPipeSegmentSizeSignature;
	}

	private static bool ElementIdsEqual(ElementId left, ElementId right)
	{
		if (left == null || right == null)
		{
			return left == null && right == null;
		}
		return RevitElementIdCompat.CompatIntegerValue(left) == RevitElementIdCompat.CompatIntegerValue(right);
	}

	private static RoutingPreferenceManager TryGetRoutingPreferenceManager(ElementType elementType)
	{
		RoutingPreferenceManager TryGetRoutingPreferenceManager;
		if (elementType == null)
		{
			TryGetRoutingPreferenceManager = null;
		}
		else
		{
			PropertyInfo propertyInfo = ((object)elementType).GetType().GetProperty("RoutingPreferenceManager");
			if ((object)propertyInfo == null)
			{
				TryGetRoutingPreferenceManager = null;
			}
			else
			{
				try
				{
					object value = propertyInfo.GetValue(elementType, null);
					TryGetRoutingPreferenceManager = (RoutingPreferenceManager)((value is RoutingPreferenceManager) ? value : null);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					TryGetRoutingPreferenceManager = null;
					ProjectData.ClearProjectError();
				}
			}
		}
		return TryGetRoutingPreferenceManager;
	}

	private static string ResolveElementName(Element element)
	{
		string ResolveElementName;
		if (element == null)
		{
			ResolveElementName = string.Empty;
		}
		else
		{
			try
			{
				ResolveElementName = element.Name ?? string.Empty;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ResolveElementName = string.Empty;
				ProjectData.ClearProjectError();
			}
		}
		return ResolveElementName;
	}

	private static string ResolveParameterName(Parameter parameter)
	{
		string ResolveParameterName;
		try
		{
			object obj;
			if (parameter == null)
			{
				obj = null;
			}
			else
			{
				Definition definition = parameter.Definition;
				obj = ((definition != null) ? definition.Name : null);
			}
			if (obj == null)
			{
				obj = string.Empty;
			}
			ResolveParameterName = (string)obj;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveParameterName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveParameterName;
	}

	private static string ResolveCategoryName(Element element)
	{
		string ResolveCategoryName;
		try
		{
			Category category = element.Category;
			ResolveCategoryName = ((category != null) ? category.Name : null) ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveCategoryName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveCategoryName;
	}

	private static bool CategoryNamesMatch(string left, string right)
	{
		if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
		{
			return string.IsNullOrWhiteSpace(left) && string.IsNullOrWhiteSpace(right);
		}
		return string.Equals(Normalize(left), Normalize(right), StringComparison.Ordinal);
	}

	private static string Normalize(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().ToLowerInvariant();
	}

	public static SystemTypePreflightSummary BuildSummary(SystemSyncExecutionPlan executionPlan, RoutingDependencyPreflightPlan dependencyPlan)
	{
		SystemTypePreflightSummary summary = new SystemTypePreflightSummary();
		HashSet<string> activeDependencyKeys = new HashSet<string>(StringComparer.Ordinal);
		List<SystemSyncExecutionItem> executionItems = executionPlan?.Items ?? new List<SystemSyncExecutionItem>();
		checked
		{
			foreach (SystemSyncExecutionItem item in executionItems)
			{
				switch (item.ExecutionStatus)
				{
				case "NoChange":
					summary.NoChangeCount++;
					break;
				case "Ready":
					summary.ReadyCount++;
					break;
				case "ApprovalRequired":
					summary.ApprovalRequiredCount++;
					break;
				case "Blocked":
					summary.BlockedCount++;
					break;
				}
				if (!string.Equals(Normalize(item.SyncAction), "skipmissingtype", StringComparison.Ordinal) && !string.Equals(Normalize(item.ExecutionStatus), "skippedmissing", StringComparison.Ordinal))
				{
					activeDependencyKeys.Add(BuildSystemPreflightKey(item.SystemFamilyKind, item.SystemTypeName));
				}
				if (item.RequiresLoadableFoundation)
				{
					summary.LoadableFoundationBlockedCount++;
				}
			}
			foreach (RoutingDependencyPreflightItem item2 in dependencyPlan?.Items ?? new List<RoutingDependencyPreflightItem>())
			{
				if (executionItems.Count <= 0 || activeDependencyKeys.Contains(BuildSystemPreflightKey(item2.SystemFamilyKind, item2.SystemTypeName)))
				{
					switch (item2.Action)
					{
					case "LoadMissingDependencyFamily":
						summary.MissingDependencyFamilyCount++;
						break;
					case "ReloadFamilyOverwrite":
					case "ReuseAndCleanupDuplicateTypes":
					case "PromoteOrRenameDependencyType":
						summary.DependencyReloadCount++;
						break;
					case "ManualReviewNameOnlyMatch":
						summary.DependencyManualReviewCount++;
						break;
					}
				}
			}
			return summary;
		}
	}

	private static string BuildSystemPreflightKey(string systemFamilyKind, string typeName)
	{
		return Normalize(systemFamilyKind) + "|" + Normalize(typeName);
	}
}
