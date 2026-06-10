using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.Exceptions;
using Microsoft.VisualBasic.CompilerServices;

public sealed class SystemTypeApplyExecutionService
{
	private sealed class AllowedLoadedFamilyIdentity
	{
		public string Name { get; set; }

		public string CategoryName { get; set; }

		public AllowedLoadedFamilyIdentity()
		{
			Name = string.Empty;
			CategoryName = string.Empty;
		}
	}

	private sealed class DuplicateCleanupResult
	{
		public List<string> AttemptedNames { get; set; }

		public List<string> DeletedNames { get; set; }

		public List<string> FailedNames { get; set; }

		public string ExceptionMessage { get; set; }

		public bool AllDeleted
		{
			get
			{
				if (FailedNames.Count == 0)
				{
					return AttemptedNames.Count == DeletedNames.Count;
				}
				return false;
			}
		}

		public DuplicateCleanupResult()
		{
			AttemptedNames = new List<string>();
			DeletedNames = new List<string>();
			FailedNames = new List<string>();
			ExceptionMessage = string.Empty;
		}
	}

	private sealed class RoutingRuleComparisonItem
	{
		public int SourceIndex { get; set; }

		public RoutingPreferenceRule Rule { get; set; }

		public ElementId ExpectedTargetPartId { get; set; }

		public RoutingRuleComparisonItem()
		{
			ExpectedTargetPartId = ElementId.InvalidElementId;
		}
	}

	private sealed class MaterialEnsureResult
	{
		public ElementId MaterialId { get; set; }

		public bool ExactSignatureMatch { get; set; }

		public string Warning { get; set; }

		public MaterialEnsureResult()
		{
			MaterialId = ElementId.InvalidElementId;
			Warning = string.Empty;
		}
	}

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

	private sealed class NonFamilyRoutingDependencyCheck
	{
		public string Action { get; set; }

		public string Notes { get; set; }

		public bool IsBlocking { get; set; }

		public NonFamilyRoutingDependencyCheck()
		{
			Action = string.Empty;
			Notes = string.Empty;
		}

		public static NonFamilyRoutingDependencyCheck Ready(string actionName, string notes)
		{
			return new NonFamilyRoutingDependencyCheck
			{
				Action = (actionName ?? string.Empty),
				Notes = (notes ?? string.Empty),
				IsBlocking = false
			};
		}

		public static NonFamilyRoutingDependencyCheck Block(string actionName, string notes)
		{
			return new NonFamilyRoutingDependencyCheck
			{
				Action = (actionName ?? string.Empty),
				Notes = (notes ?? string.Empty),
				IsBlocking = true
			};
		}
	}

	private sealed class ApplySummarySnapshot
	{
		public int CreatedCount { get; set; }

		public int OverwrittenCount { get; set; }

		public int ConsolidatedCount { get; set; }

		public int DependencyLoadedCount { get; set; }

		public int RetypedElementCount { get; set; }

		public int DeletedObsoleteTypeCount { get; set; }

		public int TrackingRefreshedCount { get; set; }

		public int BlockedCount { get; set; }

		public int SkippedCount { get; set; }

		public int FailedCount { get; set; }
	}

	private sealed class CriticalSystemTypeReferenceException : Exception
	{
		public CriticalSystemTypeReferenceException(string message)
			: base(message)
		{
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__112_002D0
	{
		public string _0024VB_0024Local_normalizedName;

		public Func<Family, bool> _0024I0;

		public _Closure_0024__112_002D0(_Closure_0024__112_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_normalizedName = arg0._0024VB_0024Local_normalizedName;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(Family x)
		{
			if (x != null)
			{
				return string.Equals(Normalize(((Element)x).Name), _0024VB_0024Local_normalizedName, StringComparison.Ordinal);
			}
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__119_002D0
	{
		public ISet<int> _0024VB_0024Local_familyStateBefore;

		public string _0024VB_0024Local_expectedName;

		public _Closure_0024__119_002D0(_Closure_0024__119_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_familyStateBefore = arg0._0024VB_0024Local_familyStateBefore;
				_0024VB_0024Local_expectedName = arg0._0024VB_0024Local_expectedName;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__1(Family x)
		{
			if (x != null && ((Element)x).Id != null)
			{
				return !_0024VB_0024Local_familyStateBefore.Contains(RevitElementIdCompat.CompatIntegerValue(((Element)x).Id));
			}
			return false;
		}

		[SpecialName]
		internal bool _Lambda_0024__3(string x)
		{
			return !string.Equals(Normalize(x), Normalize(_0024VB_0024Local_expectedName), StringComparison.Ordinal);
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__120_002D0
	{
		public ISet<int> _0024VB_0024Local_familyStateBefore;

		public string _0024VB_0024Local_expectedName;

		public _Closure_0024__120_002D0(_Closure_0024__120_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_familyStateBefore = arg0._0024VB_0024Local_familyStateBefore;
				_0024VB_0024Local_expectedName = arg0._0024VB_0024Local_expectedName;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__1(Family x)
		{
			if (x != null && ((Element)x).Id != null)
			{
				return !_0024VB_0024Local_familyStateBefore.Contains(RevitElementIdCompat.CompatIntegerValue(((Element)x).Id));
			}
			return false;
		}

		[SpecialName]
		internal bool _Lambda_0024__3(string x)
		{
			return !string.Equals(Normalize(x), Normalize(_0024VB_0024Local_expectedName), StringComparison.Ordinal);
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__146_002D0
	{
		public string _0024VB_0024Local_candidate;

		public _Closure_0024__146_002D0(_Closure_0024__146_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_candidate = arg0._0024VB_0024Local_candidate;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(ElementType x)
		{
			return string.Equals(((Element)x).Name, _0024VB_0024Local_candidate, StringComparison.OrdinalIgnoreCase);
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__51_002D0
	{
		public Segment _0024VB_0024Local_sourceSegment;

		public _Closure_0024__51_002D0(_Closure_0024__51_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_sourceSegment = arg0._0024VB_0024Local_sourceSegment;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(MEPSize x)
		{
			if (x != null)
			{
				return !SegmentContainsNominalDiameter(_0024VB_0024Local_sourceSegment, x.NominalDiameter);
			}
			return false;
		}
	}

	private const double PipeSegmentSizeTolerance = 1E-09;

	private SystemTypeApplyExecutionService()
	{
	}

	public static SystemTypeApplyExecutionReport Execute(Document targetDocument, Document standardDocument, SystemTypePreflightReport preflightReport, string preflightPath, Action<int, int, string> progress = null)
	{
		//IL_084f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0856: Expected O, but got Unknown
		//IL_0858: Unknown result type (might be due to invalid IL or missing references)
		//IL_0876: Unknown result type (might be due to invalid IL or missing references)
		if (targetDocument == null)
		{
			throw new ArgumentNullException("targetDocument");
		}
		if (standardDocument == null)
		{
			throw new ArgumentNullException("standardDocument");
		}
		if (preflightReport == null)
		{
			throw new ArgumentNullException("preflightReport");
		}
		SystemTypeApplyExecutionReport report = new SystemTypeApplyExecutionReport
		{
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			ProjectDocumentTitle = (preflightReport.ProjectDocumentTitle ?? targetDocument.Title),
			ProjectDocumentPath = (string.IsNullOrWhiteSpace(preflightReport.ProjectDocumentPath) ? ProjectSnapshotStore.ResolveProjectIdentityPath(targetDocument) : preflightReport.ProjectDocumentPath),
			StandardDisplayName = preflightReport.StandardDisplayName,
			PreflightPath = (preflightPath ?? string.Empty)
		};
		List<RoutingDependencyPreflightItem> dependencyPlanItems = CollectDependencyItems(preflightReport.DependencyPlan);
		Dictionary<string, Family> standardFamilyMap = BuildStandardFamilyMap(standardDocument, dependencyPlanItems);
		Dictionary<string, ElementType> standardSystemMap = BuildSystemTypeMap(standardDocument, CollectSystemTypeKeys(preflightReport.SyncPlan));
		Dictionary<int, List<ElementId>> targetUsageMap = null;
		HashSet<string> refreshedDependencyFamilies = new HashSet<string>(StringComparer.Ordinal);
		List<SystemSyncExecutionItem> executionItems = (preflightReport.ExecutionPlan?.Items ?? new List<SystemSyncExecutionItem>()).OrderBy([SpecialName] (SystemSyncExecutionItem x) => SystemTypeIdentityService.BuildKey(x.SystemFamilyKind, x.CategoryName, x.SystemTypeName), StringComparer.Ordinal).ToList();
		int total = Math.Max(1, executionItems.Count);
		ReportProgress(progress, 0, total, T("Preparing system type apply execution...", "시스템 타입 적용 실행 준비 중..."));
		checked
		{
			using (FamilyBrowserNativeCommandGuardService.BeginTrustedOperation("Family Browser system type apply"))
			{
				int num = executionItems.Count - 1;
				for (int itemIndex = 0; itemIndex <= num; itemIndex++)
				{
					SystemSyncExecutionItem executionPlanItem = executionItems[itemIndex];
					ReportProgress(progress, itemIndex + 1, total, T("Applying system type ", "시스템 타입 적용 중 ") + (itemIndex + 1).ToString(CultureInfo.InvariantCulture) + "/" + total.ToString(CultureInfo.InvariantCulture) + ": " + (executionPlanItem.SystemFamilyKind ?? string.Empty) + " / " + (executionPlanItem.SystemTypeName ?? string.Empty));
					string identityKey = SystemTypeIdentityService.BuildKey(executionPlanItem.SystemFamilyKind, executionPlanItem.CategoryName, executionPlanItem.SystemTypeName);
					SystemTypeSyncPlanItem syncItem = FindSyncPlanItem(preflightReport.SyncPlan, executionPlanItem.SystemFamilyKind, executionPlanItem.CategoryName, executionPlanItem.SystemTypeName);
					IEnumerable<RoutingDependencyPreflightItem> dependencyItems = FindDependencyItems(preflightReport.DependencyPlan, executionPlanItem.SystemFamilyKind, executionPlanItem.SystemTypeName);
					SystemTypeApplyExecutionItem systemTypeApplyExecutionItem = new SystemTypeApplyExecutionItem();
					systemTypeApplyExecutionItem.IdentityKey = identityKey;
					systemTypeApplyExecutionItem.SystemFamilyKind = executionPlanItem.SystemFamilyKind;
					systemTypeApplyExecutionItem.CategoryName = executionPlanItem.CategoryName;
					systemTypeApplyExecutionItem.SystemTypeName = executionPlanItem.SystemTypeName;
					systemTypeApplyExecutionItem.SyncAction = syncItem?.Action ?? executionPlanItem.SyncAction;
					systemTypeApplyExecutionItem.PreflightStatus = executionPlanItem.ExecutionStatus;
					systemTypeApplyExecutionItem.Outcome = "Skipped";
					systemTypeApplyExecutionItem.Details = executionPlanItem.Summary;
					SystemTypeApplyExecutionItem resultItem = systemTypeApplyExecutionItem;
					if (!SystemTypeSupportPolicyService.CanApply(executionPlanItem.SystemFamilyKind))
					{
						resultItem.Outcome = "SkippedReviewOnlyKind";
						resultItem.Details = T("This system family kind is configured as review-only: ", "이 시스템 패밀리 종류는 검토 전용으로 설정되어 있습니다: ") + executionPlanItem.SystemFamilyKind;
						report.Summary.SkippedCount++;
						report.Items.Add(resultItem);
						continue;
					}
					if (string.Equals(Normalize(executionPlanItem.ExecutionStatus), "blocked", StringComparison.Ordinal))
					{
						resultItem.Outcome = "Blocked";
						resultItem.Details = BuildBlockedReason(executionPlanItem);
						report.Summary.BlockedCount++;
						report.Items.Add(resultItem);
						continue;
					}
					if (syncItem == null)
					{
						resultItem.Outcome = "Failed";
						resultItem.Details = T("The system sync plan item could not be resolved.", "시스템 동기화 계획 항목을 확인하지 못했습니다.");
						report.Summary.FailedCount++;
						report.Items.Add(resultItem);
						continue;
					}
					RoutingDependencyPreflightItem unsupportedDependency = dependencyItems.FirstOrDefault([SpecialName] (RoutingDependencyPreflightItem x) => !IsSupportedDependencyAction(x.Action));
					if (unsupportedDependency != null)
					{
						resultItem.Outcome = "Blocked";
						resultItem.Details = unsupportedDependency.Reason;
						report.Summary.BlockedCount++;
						report.Items.Add(resultItem);
						continue;
					}
					try
					{
						string plannedAction = Normalize(syncItem.Action ?? executionPlanItem.SyncAction);
						string normalizedAction = NormalizeSelectedSystemTypeApplyAction(targetDocument, standardDocument, standardSystemMap, syncItem, executionPlanItem, resultItem);
						string canonicalAction = CanonicalSystemTypeActionName(normalizedAction);
						if (!string.Equals(Normalize(syncItem.Action), normalizedAction, StringComparison.Ordinal))
						{
							AddSystemApplyLog(resultItem, "SystemApply.ActionNormalized", "planned=" + (syncItem.Action ?? string.Empty) + " normalized=" + canonicalAction);
							syncItem.Action = canonicalAction;
							executionPlanItem.SyncAction = canonicalAction;
							resultItem.SyncAction = canonicalAction;
						}
						else
						{
							resultItem.SyncAction = canonicalAction;
						}
						AddSystemApplyLog(resultItem, "SystemApply.Item", "planned action=" + plannedAction + " status=" + (executionPlanItem.ExecutionStatus ?? string.Empty) + " dependencyActions=" + CountExecutableDependencyActions(dependencyItems).ToString(CultureInfo.InvariantCulture));
						AddSystemApplyLog(resultItem, "SystemApply.Start", "target=" + targetDocument.Title + " standard=" + standardDocument.Title + " kind=" + executionPlanItem.SystemFamilyKind + " category=" + executionPlanItem.CategoryName + " type=" + executionPlanItem.SystemTypeName + " planned=" + plannedAction + " normalized=" + normalizedAction + " targetExisting=" + ResolveTargetSystemTypeState(targetDocument, syncItem));
						List<NonFamilyRoutingDependencyCheck> nonFamilyPlan = ValidateNonFamilyRoutingDependencies(targetDocument, standardDocument, standardSystemMap, syncItem, executionPlanItem);
						List<string> nonFamilyBlockers = (from x in nonFamilyPlan
							where x?.IsBlocking ?? false
							select x.Notes into x
							where !string.IsNullOrWhiteSpace(x)
							select x).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
						if (nonFamilyBlockers.Count > 0)
						{
							resultItem.Outcome = "Blocked";
							resultItem.Details = T("System type apply was stopped before loading dependency families because non-family routing dependencies are not safe to map. ", "비패밀리 라우팅 의존 항목을 안전하게 매핑할 수 없어 의존 패밀리 로드 전에 시스템 타입 적용을 중단했습니다. ") + string.Join(" / ", nonFamilyBlockers);
							resultItem.Messages.AddRange(from x in nonFamilyPlan
								select x.Notes into x
								where !string.IsNullOrWhiteSpace(x)
								select x);
							report.Summary.BlockedCount++;
							report.Items.Add(resultItem);
							continue;
						}
						foreach (NonFamilyRoutingDependencyCheck planItem in nonFamilyPlan.Where([SpecialName] (NonFamilyRoutingDependencyCheck x) => x != null && !string.IsNullOrWhiteSpace(x.Notes)))
						{
							resultItem.Messages.Add(planItem.Action + ": " + planItem.Notes);
						}
						if (ShouldUseAtomicSystemTypeApplyGroup(syncItem, executionPlanItem))
						{
							ApplySummarySnapshot summarySnapshot = CaptureApplySummarySnapshot(report.Summary);
							HashSet<string> refreshedSnapshot = new HashSet<string>(refreshedDependencyFamilies, StringComparer.Ordinal);
							int dependencyActionCountBefore = resultItem.DependencyActions.Count;
							TransactionGroup itemGroup = new TransactionGroup(targetDocument, "KKY Family Browser Apply System Type");
							try
							{
								itemGroup.Start();
								try
								{
									ExecutePlannedSystemTypeAction(targetDocument, standardDocument, standardFamilyMap, standardSystemMap, ref targetUsageMap, syncItem, executionPlanItem, dependencyItems, refreshedDependencyFamilies, report, resultItem);
									itemGroup.Assimilate();
									resultItem.Messages.Add(T("Atomic system type apply group completed.", "시스템 타입 원자 적용 그룹이 완료되었습니다."));
								}
								catch (Exception projectError)
								{
									ProjectData.SetProjectError(projectError);
									TryRollback(itemGroup);
									RestoreApplySummarySnapshot(report.Summary, summarySnapshot);
									RestoreRefreshedDependencyFamilies(refreshedDependencyFamilies, refreshedSnapshot);
									TrimDependencyActions(resultItem, dependencyActionCountBefore);
									resultItem.Messages.Add(T("Atomic system type apply group was rolled back. Dependency fitting families, non-family routing parts, and the target system type should not remain as a fitting-only partial result.", "시스템 타입 원자 적용 그룹이 롤백되었습니다. 의존 피팅 패밀리, 비패밀리 라우팅 부품, 대상 시스템 타입이 피팅만 적용된 부분 결과로 남지 않아야 합니다."));
									throw;
								}
							}
							finally
							{
								((IDisposable)itemGroup)?.Dispose();
							}
						}
						else
						{
							ExecutePlannedSystemTypeAction(targetDocument, standardDocument, standardFamilyMap, standardSystemMap, ref targetUsageMap, syncItem, executionPlanItem, dependencyItems, refreshedDependencyFamilies, report, resultItem);
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						resultItem.Outcome = "Failed";
						resultItem.Details = ex2.Message;
						resultItem.Messages.Add(ex2.ToString());
						report.Summary.FailedCount++;
						ProjectData.ClearProjectError();
					}
					report.Items.Add(resultItem);
				}
			}
			ReportProgress(progress, total, total, T("System type apply execution finished.", "시스템 타입 적용 실행이 완료되었습니다."));
			return report;
		}
	}

	private static string T(string englishText, string koreanText)
	{
		return FamilyBrowserLanguageService.Text(englishText, koreanText);
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

	private static void ExecutePlannedSystemTypeAction(Document targetDocument, Document standardDocument, IDictionary<string, Family> standardFamilyMap, IDictionary<string, ElementType> standardSystemMap, ref Dictionary<int, List<ElementId>> targetUsageMap, SystemTypeSyncPlanItem syncItem, SystemSyncExecutionItem executionPlanItem, IEnumerable<RoutingDependencyPreflightItem> dependencyItems, ISet<string> refreshedDependencyFamilies, SystemTypeApplyExecutionReport report, SystemTypeApplyExecutionItem resultItem)
	{
		//IL_0128: Unknown result type (might be due to invalid IL or missing references)
		//IL_012e: Expected O, but got Unknown
		//IL_012f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0142: Unknown result type (might be due to invalid IL or missing references)
		string normalizedAction = Normalize(syncItem.Action);
		bool willMutateSystemType = IsSystemTypeMutationAction(normalizedAction);
		checked
		{
			if (string.Equals(normalizedAction, "skipmissingtype", StringComparison.Ordinal))
			{
				resultItem.Outcome = "SkippedMissing";
				resultItem.Details = T("The registered standard system type is not loaded in the current project, so it is left uncreated.", "등록된 표준 시스템 타입이 현재 프로젝트에 로드되어 있지 않아 새로 만들지 않고 건너뜁니다.");
				AddSystemApplyLog(resultItem, "SystemApply.SkipMissingType", resultItem.Details);
				report.Summary.SkippedCount++;
				return;
			}
			if (SystemTypeSupportPolicyService.RequiresDependencyRefresh(executionPlanItem.SystemFamilyKind) && !willMutateSystemType)
			{
				if (HasDependencyFamilyMutationAction(dependencyItems))
				{
					resultItem.Outcome = "FailedPlannerNoMutation";
					resultItem.Details = T("The selected Pipe Type apply planned dependency refresh but no create/overwrite/consolidate system type action. This would leave dependencies without applying the Pipe Type. Re-run Current Model Check or force an overwrite plan.", "선택한 배관 타입 적용 계획에 의존 항목 갱신은 있지만 시스템 타입 생성/덮어쓰기/정리 작업이 없습니다. 이 상태는 배관 타입 적용 없이 의존 항목만 남길 수 있습니다. 현재 모델 검사를 다시 실행하거나 덮어쓰기 계획을 강제로 생성하세요.");
					AddSystemApplyLog(resultItem, "SystemApply.DependencyEnsure.SkippedPlannerNoMutation", resultItem.Details);
					report.Summary.FailedCount++;
					return;
				}
				AddSystemApplyLog(resultItem, "SystemApply.DependencyEnsure.Skipped", T("No create/overwrite action is required for this selected system type.", "선택한 시스템 타입에는 생성/덮어쓰기 작업이 필요하지 않습니다."));
			}
			if (SystemTypeSupportPolicyService.RequiresDependencyRefresh(executionPlanItem.SystemFamilyKind) && willMutateSystemType)
			{
				AddSystemApplyLog(resultItem, "SystemApply.DependencyEnsure.Start", "action=" + CanonicalSystemTypeActionName(normalizedAction));
				Transaction tx = new Transaction(targetDocument, "KKY Family Browser Ensure Routing Dependencies");
				try
				{
					tx.Start();
					try
					{
						EnsureNonFamilyRoutingDependencies(targetDocument, standardDocument, standardSystemMap, syncItem, resultItem);
						tx.Commit();
						AddSystemApplyLog(resultItem, "SystemApply.DependencyEnsure.Success", T("Non-family routing dependencies are ready.", "비패밀리 라우팅 의존 항목이 준비되었습니다."));
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						TryRollback(tx);
						throw;
					}
				}
				finally
				{
					((IDisposable)tx)?.Dispose();
				}
				AddSystemApplyLog(resultItem, "SystemApply.FamilyDependencyLoad.Start", "dependencyActions=" + CountExecutableDependencyActions(dependencyItems).ToString(CultureInfo.InvariantCulture));
				ApplyDependencyFamilies(targetDocument, standardDocument, standardFamilyMap, dependencyItems, refreshedDependencyFamilies, report, resultItem);
				RegenerateAfterFamilyLoad(targetDocument, resultItem, "SystemApply.FamilyDependencyLoad.Regenerated");
				AddSystemApplyLog(resultItem, "SystemApply.FamilyDependencyLoad.Success", T("Routing dependency families are ready.", "라우팅 의존 패밀리가 준비되었습니다."));
				AddSystemApplyLog(resultItem, "SystemApply.RoutingFamilySymbolEnsure.Start", T("Verifying loaded routing FamilySymbol references before the type transaction.", "타입 트랜잭션 전에 로드된 라우팅 FamilySymbol 참조를 확인합니다."));
				EnsureRoutingFamilySymbolDependencies(targetDocument, standardDocument, standardSystemMap, syncItem, dependencyItems, resultItem);
				AddSystemApplyLog(resultItem, "SystemApply.RoutingFamilySymbolEnsure.Success", T("Routing family symbols are ready before the type transaction.", "타입 트랜잭션 전에 라우팅 패밀리 심볼이 준비되었습니다."));
			}
			switch (normalizedAction)
			{
			case "createmissingtype":
				AddSystemApplyLog(resultItem, "SystemApply.TypeCreateOrOverwrite.Start", "action=CreateMissingType");
				ExecuteCreateMissingType(targetDocument, standardDocument, standardSystemMap, targetUsageMap, syncItem, dependencyItems, resultItem, report);
				AddSystemApplyLog(resultItem, "SystemApply.TypeCreateOrOverwrite.Success", "action=CreateMissingType outcome=" + resultItem.Outcome);
				break;
			case "overwritedestination":
				EnsureTargetUsageMap(targetDocument, ref targetUsageMap);
				AddSystemApplyLog(resultItem, "SystemApply.TypeCreateOrOverwrite.Start", "action=OverwriteDestination");
				ExecuteOverwriteDestination(targetDocument, standardDocument, standardSystemMap, targetUsageMap, syncItem, dependencyItems, resultItem, report);
				AddSystemApplyLog(resultItem, "SystemApply.TypeCreateOrOverwrite.Success", "action=OverwriteDestination outcome=" + resultItem.Outcome);
				break;
			case "consolidateduplicatesuffixtypes":
				EnsureTargetUsageMap(targetDocument, ref targetUsageMap);
				ExecuteConsolidateDuplicateTypes(targetDocument, targetUsageMap, syncItem, resultItem, report);
				break;
			case "keepdestination":
				if (syncItem.RelatedDuplicateNames.Count > 0)
				{
					EnsureTargetUsageMap(targetDocument, ref targetUsageMap);
					ExecuteConsolidateDuplicateTypes(targetDocument, targetUsageMap, syncItem, resultItem, report);
				}
				else
				{
					resultItem.Outcome = "NoChange";
					resultItem.Details = T("The canonical type already matches the standard source.", "표준 타입이 이미 표준 원본과 일치합니다.");
					report.Summary.SkippedCount++;
				}
				break;
			default:
				resultItem.Outcome = "SkippedUnsupportedAction";
				resultItem.Details = T("The planned action is not yet executable.", "계획된 작업은 아직 실행할 수 없습니다.");
				report.Summary.SkippedCount++;
				break;
			}
		}
	}

	private static bool ShouldUseAtomicSystemTypeApplyGroup(SystemTypeSyncPlanItem syncItem, SystemSyncExecutionItem executionPlanItem)
	{
		if (syncItem == null || executionPlanItem == null)
		{
			return false;
		}
		if (!SystemTypeSupportPolicyService.RequiresDependencyRefresh(executionPlanItem.SystemFamilyKind))
		{
			return false;
		}
		return IsSystemTypeRoutingRebuildAction(syncItem.Action);
	}

	private static string NormalizeSelectedSystemTypeApplyAction(Document targetDocument, Document standardDocument, IDictionary<string, ElementType> standardSystemMap, SystemTypeSyncPlanItem syncItem, SystemSyncExecutionItem executionPlanItem, SystemTypeApplyExecutionItem resultItem)
	{
		string action = Normalize(syncItem?.Action ?? executionPlanItem?.SyncAction);
		if (syncItem == null || executionPlanItem == null)
		{
			return action;
		}
		if (!SystemTypeSupportPolicyService.RequiresDependencyRefresh(executionPlanItem.SystemFamilyKind))
		{
			if (string.Equals(action, "keepdestination", StringComparison.Ordinal) && syncItem.RelatedDuplicateNames != null && syncItem.RelatedDuplicateNames.Count > 0)
			{
				return "consolidateduplicatesuffixtypes";
			}
			return action;
		}
		if (string.Equals(action, "keepdestination", StringComparison.Ordinal))
		{
			if (!TargetSystemTypeMatchesStandardRouting(targetDocument, standardDocument, standardSystemMap, syncItem, resultItem))
			{
				return "overwritedestination";
			}
			if (syncItem.RelatedDuplicateNames != null && syncItem.RelatedDuplicateNames.Count > 0)
			{
				return "consolidateduplicatesuffixtypes";
			}
		}
		return action;
	}

	private unsafe static bool TargetSystemTypeMatchesStandardRouting(Document targetDocument, Document standardDocument, IDictionary<string, ElementType> standardSystemMap, SystemTypeSyncPlanItem syncItem, SystemTypeApplyExecutionItem resultItem)
	{
		//IL_0119: Unknown result type (might be due to invalid IL or missing references)
		//IL_011e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0120: Unknown result type (might be due to invalid IL or missing references)
		//IL_0123: Invalid comparison between Unknown and I4
		//IL_012c: Unknown result type (might be due to invalid IL or missing references)
		//IL_013e: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e8: Unknown result type (might be due to invalid IL or missing references)
		bool TargetSystemTypeMatchesStandardRouting;
		if (targetDocument == null || standardDocument == null || syncItem == null)
		{
			TargetSystemTypeMatchesStandardRouting = false;
		}
		else
		{
			try
			{
				ElementType sourceType = ResolveStandardSystemType(standardSystemMap, syncItem);
				if (sourceType == null)
				{
					AddSystemApplyLog(resultItem, "SystemApply.KeepDestinationRejected", T("The standard system type could not be resolved for live routing comparison.", "실시간 라우팅 비교를 위한 표준 시스템 타입을 확인하지 못했습니다."));
					TargetSystemTypeMatchesStandardRouting = false;
				}
				else
				{
					ElementType targetType = null;
					HashSet<string> requestedKeys = new HashSet<string>(StringComparer.Ordinal) { BuildIdentityKey(syncItem) };
					BuildSystemTypeMap(targetDocument, requestedKeys).TryGetValue(BuildIdentityKey(syncItem), out targetType);
					if (targetType == null)
					{
						AddSystemApplyLog(resultItem, "SystemApply.KeepDestinationRejected", T("The target system type does not exist for live routing comparison.", "실시간 라우팅 비교를 위한 대상 시스템 타입이 없습니다."));
						TargetSystemTypeMatchesStandardRouting = false;
					}
					else
					{
						RoutingPreferenceManager sourceManager = TryGetRoutingPreferenceManager(sourceType);
						RoutingPreferenceManager targetManager = TryGetRoutingPreferenceManager(targetType);
						if (sourceManager == null && targetManager == null)
						{
							AddSystemApplyLog(resultItem, "SystemApply.KeepDestinationRejected", T("Routing preference manager is unavailable, so the target cannot be proven to match the standard.", "라우팅 환경설정 관리자를 사용할 수 없어 대상이 표준과 일치한다고 증명할 수 없습니다."));
							TargetSystemTypeMatchesStandardRouting = false;
						}
						else if (sourceManager == null || targetManager == null)
						{
							AddSystemApplyLog(resultItem, "SystemApply.KeepDestinationRejected", T("Routing preference manager availability differs between standard and target.", "표준과 대상의 라우팅 환경설정 관리자 사용 가능 여부가 다릅니다."));
							TargetSystemTypeMatchesStandardRouting = false;
						}
						else
						{
							foreach (RoutingPreferenceRuleGroupType group in Enum.GetValues(typeof(RoutingPreferenceRuleGroupType)).Cast<RoutingPreferenceRuleGroupType>())
							{
								if ((int)group == -1)
								{
									continue;
								}
								List<RoutingRuleComparisonItem> expectedRules = BuildExpectedRoutingRulesForComparison(targetDocument, standardDocument, sourceManager, group, resultItem, "SystemApply.KeepDestinationRoutingRuleSkipped");
								int targetRuleCount = targetManager.GetNumberOfRules(group);
								if (expectedRules.Count == targetRuleCount)
								{
									int num = checked(expectedRules.Count - 1);
									int index = 0;
									while (index <= num)
									{
										RoutingRuleComparisonItem expectedRule = expectedRules[index];
										RoutingPreferenceRule sourceRule = expectedRule.Rule;
										RoutingPreferenceRule targetRule = targetManager.GetRule(group, index);
										if (sourceRule == null || targetRule == null)
										{
											AddSystemApplyLog(resultItem, "SystemApply.KeepDestinationRejected", "A routing rule could not be read. group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " index=" + index.ToString(CultureInfo.InvariantCulture) + " sourceIndex=" + expectedRule.SourceIndex.ToString(CultureInfo.InvariantCulture));
											TargetSystemTypeMatchesStandardRouting = false;
										}
										else
										{
											ElementId expectedTargetPartId = expectedRule.ExpectedTargetPartId;
											if (!ElementIdsEqual(expectedTargetPartId, targetRule.MEPPartId))
											{
												AddSystemApplyLog(resultItem, "SystemApply.KeepDestinationRejected", "A routing rule mapped part differs. group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " index=" + index.ToString(CultureInfo.InvariantCulture) + " sourceIndex=" + expectedRule.SourceIndex.ToString(CultureInfo.InvariantCulture) + " expected=" + FormatElementId(expectedTargetPartId) + " actual=" + FormatElementId(targetRule.MEPPartId));
												TargetSystemTypeMatchesStandardRouting = false;
											}
											else
											{
												Element sourcePart = ((sourceRule.MEPPartId == null || sourceRule.MEPPartId == ElementId.InvalidElementId) ? null : standardDocument.GetElement(sourceRule.MEPPartId));
												Element targetPart = ((targetRule.MEPPartId == null || targetRule.MEPPartId == ElementId.InvalidElementId) ? null : targetDocument.GetElement(targetRule.MEPPartId));
												if (sourcePart != null && !RoutingPartDefinitionMatchesForPostCheck(standardDocument, sourcePart, targetDocument, targetPart))
												{
													AddSystemApplyLog(resultItem, "SystemApply.KeepDestinationRejected", "A routing part definition differs. group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " index=" + index.ToString(CultureInfo.InvariantCulture) + " sourceIndex=" + expectedRule.SourceIndex.ToString(CultureInfo.InvariantCulture) + " part=" + ((object)sourcePart).GetType().Name + ":" + ResolveElementName(sourcePart));
													TargetSystemTypeMatchesStandardRouting = false;
												}
												else
												{
													if (string.Equals(sourceRule.Description ?? string.Empty, targetRule.Description ?? string.Empty, StringComparison.Ordinal) && ResolveCriterionCount(sourceRule) == ResolveCriterionCount(targetRule))
													{
														index = checked(index + 1);
														continue;
													}
													AddSystemApplyLog(resultItem, "SystemApply.KeepDestinationRejected", "Routing rule description or criterion count differs. group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " index=" + index.ToString(CultureInfo.InvariantCulture) + " sourceIndex=" + expectedRule.SourceIndex.ToString(CultureInfo.InvariantCulture));
													TargetSystemTypeMatchesStandardRouting = false;
												}
											}
										}
										goto end_IL_0011;
									}
									continue;
								}
								AddSystemApplyLog(resultItem, "SystemApply.KeepDestinationRejected", "Effective routing rule count differs. group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " source=" + expectedRules.Count.ToString(CultureInfo.InvariantCulture) + " target=" + targetRuleCount.ToString(CultureInfo.InvariantCulture));
								TargetSystemTypeMatchesStandardRouting = false;
								goto end_IL_0011;
							}
							AddSystemApplyLog(resultItem, "SystemApply.KeepDestinationAccepted", T("Live routing preference comparison matched the standard.", "실시간 라우팅 환경설정 비교 결과가 표준과 일치합니다."));
							TargetSystemTypeMatchesStandardRouting = true;
						}
					}
				}
				end_IL_0011:;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				AddSystemApplyLog(resultItem, "SystemApply.KeepDestinationRejected", T("Live routing comparison failed: ", "실시간 라우팅 비교 실패: ") + ResolveExceptionMessage(ex2));
				TargetSystemTypeMatchesStandardRouting = false;
				ProjectData.ClearProjectError();
			}
		}
		return TargetSystemTypeMatchesStandardRouting;
	}

	private static bool IsSystemTypeRoutingRebuildAction(string action)
	{
		string left = Normalize(action);
		if (Operators.CompareString(left, "createmissingtype", TextCompare: false) == 0 || Operators.CompareString(left, "overwritedestination", TextCompare: false) == 0)
		{
			return true;
		}
		return false;
	}

	private static bool IsSystemTypeMutationAction(string action)
	{
		switch (Normalize(action))
		{
		case "createmissingtype":
		case "overwritedestination":
		case "consolidateduplicatesuffixtypes":
			return true;
		default:
			return false;
		}
	}

	private static bool HasDependencyFamilyMutationAction(IEnumerable<RoutingDependencyPreflightItem> dependencyItems)
	{
		return CountExecutableDependencyActions(dependencyItems) > 0;
	}

	private static int CountExecutableDependencyActions(IEnumerable<RoutingDependencyPreflightItem> dependencyItems)
	{
		return (dependencyItems ?? new List<RoutingDependencyPreflightItem>()).Count([SpecialName] (RoutingDependencyPreflightItem x) =>
		{
			switch (Normalize(x?.Action))
			{
			case "loadmissingdependencyfamily":
			case "reloadfamilyoverwrite":
			case "promoteorrenamedependencytype":
				return true;
			default:
				return false;
			}
		});
	}

	private static string CanonicalSystemTypeActionName(string action)
	{
		return Normalize(action) switch
		{
			"createmissingtype" => "CreateMissingType", 
			"skipmissingtype" => "SkipMissingType", 
			"overwritedestination" => "OverwriteDestination", 
			"consolidateduplicatesuffixtypes" => "ConsolidateDuplicateSuffixTypes", 
			"keepdestination" => "KeepDestination", 
			_ => action ?? string.Empty, 
		};
	}

	private static string ResolveTargetSystemTypeState(Document targetDocument, SystemTypeSyncPlanItem syncItem)
	{
		string ResolveTargetSystemTypeState;
		if (targetDocument == null || syncItem == null)
		{
			ResolveTargetSystemTypeState = "unavailable";
		}
		else
		{
			try
			{
				ElementType targetType = null;
				HashSet<string> requestedKeys = new HashSet<string>(StringComparer.Ordinal) { BuildIdentityKey(syncItem) };
				BuildSystemTypeMap(targetDocument, requestedKeys).TryGetValue(BuildIdentityKey(syncItem), out targetType);
				ResolveTargetSystemTypeState = ((targetType != null) ? (((object)targetType).GetType().Name + ":" + ResolveElementName((Element)(object)targetType) + ":" + FormatElementId(((Element)targetType).Id)) : "missing");
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ResolveTargetSystemTypeState = "error:" + ResolveExceptionMessage(ex2);
				ProjectData.ClearProjectError();
			}
		}
		return ResolveTargetSystemTypeState;
	}

	private static void ApplyDependencyFamilies(Document targetDocument, Document standardDocument, IDictionary<string, Family> standardFamilyMap, IEnumerable<RoutingDependencyPreflightItem> dependencyItems, ISet<string> refreshedDependencyFamilies, SystemTypeApplyExecutionReport report, SystemTypeApplyExecutionItem resultItem)
	{
		LoadableFamilyLoadOptions loadOptions = new LoadableFamilyLoadOptions(overwriteParameterValues: true);
		List<RoutingDependencyPreflightItem> dependencyItemList = (dependencyItems ?? new List<RoutingDependencyPreflightItem>()).Where([SpecialName] (RoutingDependencyPreflightItem x) => x != null).ToList();
		foreach (RoutingDependencyPreflightItem dependencyItem in dependencyItemList.OrderBy([SpecialName] (RoutingDependencyPreflightItem x) => Normalize(x.SourceFamilyName), StringComparer.Ordinal).ThenBy([SpecialName] (RoutingDependencyPreflightItem x) => Normalize(x.SourceTypeName), StringComparer.Ordinal))
		{
			switch (Normalize(dependencyItem.Action))
			{
			case "reuseloadeddependency":
				resultItem.DependencyActions.Add(T("Reuse dependency: ", "의존 항목 재사용: ") + dependencyItem.SourceFamilyName + " : " + dependencyItem.SourceTypeName);
				break;
			case "reuseandcleanupduplicatetypes":
				resultItem.DependencyActions.Add(T("Reuse canonical dependency and keep duplicate type cleanup for review: ", "표준 의존 항목을 재사용하고 중복 타입 정리는 검토로 남김: ") + dependencyItem.SourceFamilyName + " : " + dependencyItem.SourceTypeName);
				break;
			case "loadmissingdependencyfamily":
			case "reloadfamilyoverwrite":
			case "promoteorrenamedependencytype":
				LoadStandardDependencyFamily(targetDocument, standardDocument, standardFamilyMap, dependencyItem, dependencyItemList, loadOptions, refreshedDependencyFamilies, report, resultItem);
				break;
			case "manualreviewnameonlymatch":
				throw new InvalidOperationException(T("A dependency family with the same name exists, but its canonical category identity differs from the standard RVT. Review it before applying the system type: ", "같은 이름의 의존 패밀리가 있지만 표준 RVT의 카테고리 기준과 다릅니다. 시스템 타입 적용 전에 검토하세요: ") + dependencyItem.SourceFamilyName);
			default:
				throw new InvalidOperationException(T("Unsupported dependency action was found during system type apply: ", "시스템 타입 적용 중 지원하지 않는 의존 작업이 발견되었습니다: ") + dependencyItem.Action);
			}
		}
	}

	private unsafe static void EnsureNonFamilyRoutingDependencies(Document targetDocument, Document standardDocument, IDictionary<string, ElementType> standardSystemMap, SystemTypeSyncPlanItem syncItem, SystemTypeApplyExecutionItem resultItem)
	{
		//IL_0099: Unknown result type (might be due to invalid IL or missing references)
		//IL_009e: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a3: Invalid comparison between Unknown and I4
		//IL_00aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d6: Unknown result type (might be due to invalid IL or missing references)
		if (targetDocument == null || standardDocument == null || syncItem == null)
		{
			return;
		}
		string left = Normalize(syncItem.Action);
		if (Operators.CompareString(left, "createmissingtype", TextCompare: false) != 0 && Operators.CompareString(left, "overwritedestination", TextCompare: false) != 0)
		{
			return;
		}
		RoutingPreferenceManager manager = TryGetRoutingPreferenceManager(ResolveStandardSystemType(standardSystemMap, syncItem) ?? throw new InvalidOperationException(T("The source system type was not found in the registered standard RVT before non-family routing dependency ensure: ", "비패밀리 라우팅 의존성 보정 전에 등록된 표준 RVT에서 원본 시스템 타입을 찾지 못했습니다: ") + syncItem.SourceTypeName));
		if (manager == null)
		{
			return;
		}
		HashSet<int> ensuredSegments = new HashSet<int>();
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
			int num = checked(ruleCount - 1);
			for (int index = 0; index <= num; index = checked(index + 1))
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
				Element obj = ((((rule != null) ? rule.MEPPartId : null) == null || rule.MEPPartId == ElementId.InvalidElementId) ? null : standardDocument.GetElement(rule.MEPPartId));
				PipeSegment sourceSegment = (PipeSegment)(object)((obj is PipeSegment) ? obj : null);
				if (sourceSegment != null && ((Element)sourceSegment).Id != null && !ensuredSegments.Contains(RevitElementIdCompat.CompatIntegerValue(((Element)sourceSegment).Id)))
				{
					ensuredSegments.Add(RevitElementIdCompat.CompatIntegerValue(((Element)sourceSegment).Id));
					EnsurePipeSegmentAuthoritative(targetDocument, standardDocument, sourceSegment, SystemTypeApplyAuthorityMode.AdminAuthoritative, resultItem);
					AddResultMessage(resultItem, T("Non-family routing dependency ensured before fitting family load. Rule group: ", "피팅 패밀리 로드 전에 비패밀리 라우팅 의존성을 보정했습니다. 규칙 그룹: ") + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + T(" / Segment: ", " / 세그먼트: ") + ResolveElementName((Element)(object)sourceSegment));
				}
			}
		}
	}

	private unsafe static void EnsureRoutingFamilySymbolDependencies(Document targetDocument, Document standardDocument, IDictionary<string, ElementType> standardSystemMap, SystemTypeSyncPlanItem syncItem, IEnumerable<RoutingDependencyPreflightItem> dependencyItems, SystemTypeApplyExecutionItem resultItem)
	{
		//IL_00b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bb: Invalid comparison between Unknown and I4
		//IL_00c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ee: Unknown result type (might be due to invalid IL or missing references)
		if (targetDocument == null || standardDocument == null || syncItem == null)
		{
			return;
		}
		string left = Normalize(syncItem.Action);
		if (Operators.CompareString(left, "createmissingtype", TextCompare: false) != 0 && Operators.CompareString(left, "overwritedestination", TextCompare: false) != 0)
		{
			return;
		}
		RoutingPreferenceManager manager = TryGetRoutingPreferenceManager(ResolveStandardSystemType(standardSystemMap, syncItem) ?? throw new InvalidOperationException(T("The source system type was not found in the registered standard RVT before routing family ensure: ", "라우팅 패밀리 보정 전에 등록된 표준 RVT에서 원본 시스템 타입을 찾지 못했습니다: ") + syncItem.SourceTypeName));
		if (manager == null)
		{
			AddSystemApplyLog(resultItem, "SystemApply.RoutingFamilySymbolEnsure.Skipped", "The source system type has no readable routing preference manager.");
			return;
		}
		HashSet<string> ensuredSymbols = new HashSet<string>(StringComparer.Ordinal);
		int inspectedCount = 0;
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
					Element obj = ((((rule != null) ? rule.MEPPartId : null) == null || rule.MEPPartId == ElementId.InvalidElementId) ? null : standardDocument.GetElement(rule.MEPPartId));
					FamilySymbol sourceSymbol = (FamilySymbol)(object)((obj is FamilySymbol) ? obj : null);
					if (sourceSymbol != null)
					{
						string symbolKey = BuildRoutingFamilySymbolKey(sourceSymbol);
						if (!ensuredSymbols.Contains(symbolKey))
						{
							inspectedCount++;
							ensuredSymbols.Add(symbolKey);
							string[] obj2 = new string[6] { "family=", null, null, null, null, null };
							Family family = sourceSymbol.Family;
							obj2[1] = ((family != null) ? ((Element)family).Name : null) ?? string.Empty;
							obj2[2] = " type=";
							obj2[3] = ResolveElementName((Element)(object)sourceSymbol);
							obj2[4] = " group=";
							obj2[5] = ((Enum)(*unchecked((RoutingPreferenceRuleGroupType*)(&group)))/*cast due to .constrained prefix*/).ToString();
							AddSystemApplyLog(resultItem, "SystemApply.RoutingFamilySymbolEnsure.Item", string.Concat(obj2));
							EnsureRoutingFamilySymbolLoaded(targetDocument, standardDocument, sourceSymbol, dependencyItems, resultItem);
						}
					}
				}
			}
		}
		AddSystemApplyLog(resultItem, "SystemApply.RoutingFamilySymbolEnsure.Checked", "symbols=" + inspectedCount.ToString(CultureInfo.InvariantCulture));
	}

	private static void LoadStandardDependencyFamily(Document targetDocument, Document standardDocument, IDictionary<string, Family> standardFamilyMap, RoutingDependencyPreflightItem dependencyItem, IEnumerable<RoutingDependencyPreflightItem> selectedDependencyItems, LoadableFamilyLoadOptions loadOptions, ISet<string> refreshedDependencyFamilies, SystemTypeApplyExecutionReport report, SystemTypeApplyExecutionItem resultItem)
	{
		string dependencyKey = BuildDependencyFamilyKey(dependencyItem);
		if (refreshedDependencyFamilies.Contains(dependencyKey))
		{
			resultItem.DependencyActions.Add(T("Dependency already refreshed: ", "이미 갱신된 의존 항목: ") + dependencyItem.SourceFamilyName);
			return;
		}
		Family standardFamily = ResolveStandardDependencyFamily(standardFamilyMap, dependencyItem);
		if (standardFamily == null)
		{
			throw new InvalidOperationException(T("The required dependency family was not found in the registered standard RVT: ", "필요한 의존 패밀리를 등록된 표준 RVT에서 찾지 못했습니다: ") + dependencyItem.SourceFamilyName);
		}
		if (standardFamily.IsInPlace)
		{
			throw new InvalidOperationException(T("In-place dependency families cannot be loaded for system type apply: ", "내부 의존 패밀리는 시스템 타입 적용을 위해 로드할 수 없습니다: ") + dependencyItem.SourceFamilyName);
		}
		if (!standardFamily.IsEditable)
		{
			throw new InvalidOperationException(T("The dependency family exists in the standard RVT but is not editable through the Revit API: ", "의존 패밀리가 표준 RVT에 있지만 Revit API로 편집할 수 없습니다: ") + dependencyItem.SourceFamilyName);
		}
		EnsureDependencyFamilyCanOverwrite(targetDocument, standardFamily, dependencyItem);
		Document familyDoc = null;
		checked
		{
			try
			{
				ISet<int> familyStateBefore = CaptureFamilyNameState(targetDocument);
				familyDoc = standardDocument.EditFamily(standardFamily);
				List<AllowedLoadedFamilyIdentity> allowedLoadedFamilies = BuildAllowedLoadedFamilyIdentities(standardFamily, familyDoc, standardDocument, selectedDependencyItems, new List<string> { dependencyItem.SourceFamilyName });
				Family loadedFamily = familyDoc.LoadFamily(targetDocument, (IFamilyLoadOptions)(object)loadOptions);
				if (loadedFamily == null)
				{
					throw new InvalidOperationException(T("Revit returned no family reference after dependency load: ", "의존 패밀리 로드 후 Revit이 패밀리 참조를 반환하지 않았습니다: ") + dependencyItem.SourceFamilyName);
				}
				RegenerateAfterFamilyLoad(targetDocument, resultItem, "DependencyFamilyLoad.Regenerated", dependencyItem.SourceFamilyName);
				GuardDependencyLoadDidNotCreateDuplicateFamilies(targetDocument, familyStateBefore, standardFamily, dependencyItem, loadedFamily, allowedLoadedFamilies, resultItem);
				refreshedDependencyFamilies.Add(dependencyKey);
				report.Summary.DependencyLoadedCount++;
				string verb = (string.Equals(Normalize(dependencyItem.Action), "loadmissingdependencyfamily", StringComparison.Ordinal) ? T("Load missing dependency from standard RVT: ", "표준 RVT에서 누락 의존 항목 로드: ") : T("Refresh dependency from standard RVT: ", "표준 RVT에서 의존 항목 갱신: "));
				resultItem.DependencyActions.Add(verb + dependencyItem.SourceFamilyName + " : " + dependencyItem.SourceTypeName);
			}
			finally
			{
				if (familyDoc != null)
				{
					try
					{
						familyDoc.Close(false);
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						ProjectData.ClearProjectError();
					}
				}
			}
		}
	}

	private static void ExecuteCreateMissingType(Document targetDocument, Document standardDocument, IDictionary<string, ElementType> standardSystemMap, IDictionary<int, List<ElementId>> targetUsageMap, SystemTypeSyncPlanItem syncItem, IEnumerable<RoutingDependencyPreflightItem> dependencyItems, SystemTypeApplyExecutionItem resultItem, SystemTypeApplyExecutionReport report)
	{
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		//IL_002f: Expected O, but got Unknown
		//IL_0030: Unknown result type (might be due to invalid IL or missing references)
		//IL_009e: Unknown result type (might be due to invalid IL or missing references)
		ElementType sourceType = ResolveStandardSystemType(standardSystemMap, syncItem);
		if (sourceType == null)
		{
			throw new InvalidOperationException(T("The source system type was not found in the registered standard RVT.", "등록된 표준 RVT에서 원본 시스템 타입을 찾지 못했습니다."));
		}
		ElementType copiedType = null;
		Transaction tx = new Transaction(targetDocument, "KKY Family Browser Create System Type");
		try
		{
			tx.Start();
			try
			{
				if (CanRebuildTypeInDestination(sourceType))
				{
					copiedType = CreateDestinationTypeFromExisting(targetDocument, syncItem);
					if (copiedType != null)
					{
						ApplyStandardSystemTypeDefinition(targetDocument, standardDocument, sourceType, copiedType, dependencyItems, resultItem);
					}
					else
					{
						copiedType = CopyCanonicalType(targetDocument, standardDocument, syncItem, sourceType, dependencyItems);
						if (copiedType != null)
						{
							ApplyStandardSystemTypeDefinition(targetDocument, standardDocument, sourceType, copiedType, dependencyItems, resultItem);
						}
					}
				}
				else
				{
					copiedType = CopyCanonicalType(targetDocument, standardDocument, syncItem, sourceType, dependencyItems);
				}
				if (copiedType == null)
				{
					throw new InvalidOperationException(T("The canonical system type was not copied into the target project.", "표준 시스템 타입이 대상 프로젝트로 복사되지 않았습니다."));
				}
				tx.Commit();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				TryRollback(tx);
				throw;
			}
		}
		finally
		{
			((IDisposable)tx)?.Dispose();
		}
		if (copiedType == null)
		{
			throw new InvalidOperationException(T("The canonical system type was not copied into the target project.", "표준 시스템 타입이 대상 프로젝트로 복사되지 않았습니다."));
		}
		resultItem.Outcome = "Created";
		resultItem.AppliedTypeName = ((Element)copiedType).Name;
		resultItem.Details = T("Canonical system type copied from the registered standard RVT.", "등록된 표준 RVT에서 표준 시스템 타입을 복사했습니다.");
		checked
		{
			report.Summary.CreatedCount++;
		}
	}

	private static void ExecuteOverwriteDestination(Document targetDocument, Document standardDocument, IDictionary<string, ElementType> standardSystemMap, IDictionary<int, List<ElementId>> targetUsageMap, SystemTypeSyncPlanItem syncItem, IEnumerable<RoutingDependencyPreflightItem> dependencyItems, SystemTypeApplyExecutionItem resultItem, SystemTypeApplyExecutionReport report)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003b: Expected O, but got Unknown
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0115: Unknown result type (might be due to invalid IL or missing references)
		ElementType sourceType = ResolveStandardSystemType(standardSystemMap, syncItem);
		if (sourceType == null)
		{
			throw new InvalidOperationException(T("The source system type was not found in the registered standard RVT.", "등록된 표준 RVT에서 원본 시스템 타입을 찾지 못했습니다."));
		}
		ElementType copiedType = null;
		string backupName = string.Empty;
		int retypedCount = 0;
		int deletedCount = 0;
		Transaction tx = new Transaction(targetDocument, "KKY Family Browser Overwrite System Type");
		try
		{
			tx.Start();
			try
			{
				Dictionary<string, ElementType> dictionary = BuildSystemTypeMap(targetDocument);
				string identityKey = BuildIdentityKey(syncItem);
				ElementType existingType = null;
				if (!dictionary.TryGetValue(identityKey, out existingType) || existingType == null)
				{
					throw new InvalidOperationException(T("The destination system type was not found for overwrite: ", "덮어쓸 대상 시스템 타입을 찾지 못했습니다: ") + syncItem.SourceTypeName);
				}
				if (CanRebuildTypeInDestination(sourceType))
				{
					copiedType = existingType;
					ApplyStandardSystemTypeDefinition(targetDocument, standardDocument, sourceType, copiedType, dependencyItems, resultItem);
				}
				else
				{
					backupName = (existingType.Name = GenerateTemporaryTypeName(targetDocument, syncItem.SourceTypeName));
					copiedType = CopyCanonicalType(targetDocument, standardDocument, syncItem, sourceType, dependencyItems);
					if (copiedType == null)
					{
						throw new InvalidOperationException(T("The canonical system type copy did not return a destination type.", "표준 시스템 타입 복사 후 대상 타입이 반환되지 않았습니다."));
					}
				}
				List<ElementType> obsoleteTypes = new List<ElementType> { existingType };
				if (copiedType == existingType)
				{
					obsoleteTypes.Clear();
				}
				obsoleteTypes.AddRange(ResolveDuplicateTypes(targetDocument, syncItem));
				ConsolidateObsoleteTypes(targetDocument, targetUsageMap, obsoleteTypes, copiedType, ref retypedCount, ref deletedCount);
				tx.Commit();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				TryRollback(tx);
				throw;
			}
		}
		finally
		{
			((IDisposable)tx)?.Dispose();
		}
		resultItem.Outcome = "Overwritten";
		resultItem.AppliedTypeName = ((copiedType != null) ? ((Element)copiedType).Name : null) ?? syncItem.SourceTypeName;
		resultItem.BackupTypeName = backupName;
		resultItem.RetypedElementCount = retypedCount;
		resultItem.DeletedObsoleteTypeCount = deletedCount;
		resultItem.Details = T("Canonical system type copied, existing usages remapped, and obsolete types removed.", "표준 시스템 타입을 복사하고 기존 사용 항목을 다시 연결한 뒤 오래된 타입을 제거했습니다.");
		checked
		{
			report.Summary.OverwrittenCount++;
			report.Summary.RetypedElementCount += retypedCount;
			report.Summary.DeletedObsoleteTypeCount += deletedCount;
		}
	}

	private static void ExecuteConsolidateDuplicateTypes(Document targetDocument, IDictionary<int, List<ElementId>> targetUsageMap, SystemTypeSyncPlanItem syncItem, SystemTypeApplyExecutionItem resultItem, SystemTypeApplyExecutionReport report)
	{
		//IL_000a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0010: Expected O, but got Unknown
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Expected O, but got Unknown
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_0075: Unknown result type (might be due to invalid IL or missing references)
		//IL_0097: Unknown result type (might be due to invalid IL or missing references)
		int retypedCount = 0;
		int deletedCount = 0;
		TransactionGroup group = new TransactionGroup(targetDocument, "KKY Family Browser Consolidate System Types");
		try
		{
			group.Start();
			try
			{
				Transaction tx = new Transaction(targetDocument, "KKY Family Browser Consolidate System Types");
				try
				{
					tx.Start();
					try
					{
						Dictionary<string, ElementType> dictionary = BuildSystemTypeMap(targetDocument);
						ElementType canonicalType = null;
						if (!dictionary.TryGetValue(BuildIdentityKey(syncItem), out canonicalType) || canonicalType == null)
						{
							throw new InvalidOperationException(T("The canonical destination type was not found for duplicate consolidation.", "중복 타입 정리를 위한 표준 대상 타입을 찾지 못했습니다."));
						}
						IEnumerable<ElementType> obsoleteTypes = ResolveDuplicateTypes(targetDocument, syncItem);
						ConsolidateObsoleteTypes(targetDocument, targetUsageMap, obsoleteTypes, canonicalType, ref retypedCount, ref deletedCount);
						tx.Commit();
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						TryRollback(tx);
						throw;
					}
				}
				finally
				{
					((IDisposable)tx)?.Dispose();
				}
				group.Assimilate();
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				TryRollback(group);
				throw;
			}
		}
		finally
		{
			((IDisposable)group)?.Dispose();
		}
		resultItem.Outcome = "Consolidated";
		resultItem.AppliedTypeName = syncItem.SourceTypeName;
		resultItem.RetypedElementCount = retypedCount;
		resultItem.DeletedObsoleteTypeCount = deletedCount;
		resultItem.Details = T("Duplicate-suffix system types were remapped to the canonical type and deleted.", "중복 접미사 시스템 타입을 표준 타입으로 다시 연결하고 삭제했습니다.");
		checked
		{
			report.Summary.ConsolidatedCount++;
			report.Summary.RetypedElementCount += retypedCount;
			report.Summary.DeletedObsoleteTypeCount += deletedCount;
		}
	}

	private static ElementType CopyCanonicalType(Document targetDocument, Document standardDocument, SystemTypeSyncPlanItem syncItem, ElementType sourceType, IEnumerable<RoutingDependencyPreflightItem> dependencyItems)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		ISet<int> familyStateBefore = CaptureFamilyNameState(targetDocument);
		List<Element> copiedElements = new List<Element>();
		CopyPasteOptions options = new CopyPasteOptions();
		try
		{
			options.SetDuplicateTypeNamesHandler((IDuplicateTypeNamesHandler)(object)new CopyPasteUseDestinationTypesHandler());
			ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(standardDocument, (ICollection<ElementId>)new List<ElementId> { ((Element)sourceType).Id }, targetDocument, Transform.Identity, options);
			if (copiedIds != null)
			{
				foreach (ElementId copiedId in copiedIds)
				{
					Element copied = targetDocument.GetElement(copiedId);
					if (copied != null)
					{
						copiedElements.Add(copied);
					}
				}
			}
		}
		finally
		{
			((IDisposable)options)?.Dispose();
		}
		targetDocument.Regenerate();
		GuardAgainstCopiedFamilies(targetDocument, familyStateBefore, dependencyItems);
		Dictionary<string, ElementType> dictionary = BuildSystemTypeMap(targetDocument);
		ElementType copiedType = null;
		dictionary.TryGetValue(BuildIdentityKey(syncItem), out copiedType);
		if (copiedType == null)
		{
			throw new CriticalSystemTypeReferenceException(T("System type copy did not create the exact canonical identity. ", "시스템 타입 복사 후 정확한 표준 식별자가 생성되지 않았습니다. ") + T("The operation was stopped to avoid leaving duplicate or numbered system types. ", "중복 또는 번호가 붙은 시스템 타입이 남지 않도록 작업을 중단했습니다. ") + T("Standard: ", "표준: ") + ((object)sourceType).GetType().Name + " : " + ResolveElementName((Element)(object)sourceType) + T(" / Copied: ", " / 복사됨: ") + BuildCopiedElementSummary(copiedElements));
		}
		if (!ElementIdentityMatches((Element)(object)copiedType, ((object)sourceType).GetType().Name, ResolveElementName((Element)(object)sourceType), ResolveCategoryName((Element)(object)sourceType)))
		{
			throw new CriticalSystemTypeReferenceException(T("System type copy created a different identity than the registered standard. ", "시스템 타입 복사 결과가 등록된 표준과 다른 식별자로 생성되었습니다. ") + T("The operation was stopped to avoid duplicate or suffixed system types. ", "중복 또는 접미사가 붙은 시스템 타입을 피하기 위해 작업을 중단했습니다. ") + T("Standard: ", "표준: ") + ((object)sourceType).GetType().Name + " : " + ResolveElementName((Element)(object)sourceType) + T(" / Created: ", " / 생성됨: ") + ((object)copiedType).GetType().Name + " : " + ResolveElementName((Element)(object)copiedType));
		}
		return copiedType;
	}

	private static bool CanRebuildTypeInDestination(ElementType sourceType)
	{
		return TryGetRoutingPreferenceManager(sourceType) != null;
	}

	private static ElementType CreateDestinationTypeFromExisting(Document targetDocument, SystemTypeSyncPlanItem syncItem)
	{
		Dictionary<string, ElementType> currentTargetMap = BuildSystemTypeMap(targetDocument);
		ElementType existingType = null;
		if (currentTargetMap.TryGetValue(BuildIdentityKey(syncItem), out existingType) && existingType != null)
		{
			return existingType;
		}
		ElementType template = currentTargetMap.Values.Where([SpecialName] (ElementType x) => x != null).FirstOrDefault([SpecialName] (ElementType x) => string.Equals(((object)x).GetType().Name, syncItem.SystemFamilyKind, StringComparison.OrdinalIgnoreCase) && CategoryNamesMatch(ResolveCategoryName((Element)(object)x), syncItem.CategoryName));
		if (template == null)
		{
			return null;
		}
		ElementType createdType = template.Duplicate(syncItem.SourceTypeName);
		if (createdType == null)
		{
			throw new InvalidOperationException(T("Revit did not create a destination system type from the existing template: ", "Revit이 기존 템플릿에서 대상 시스템 타입을 만들지 못했습니다: ") + syncItem.SourceTypeName);
		}
		if (!ElementIdentityMatches((Element)(object)createdType, syncItem.SystemFamilyKind, syncItem.SourceTypeName, syncItem.CategoryName))
		{
			throw new CriticalSystemTypeReferenceException(T("Revit created a system type with a different identity than requested. ", "Revit이 요청한 것과 다른 식별자의 시스템 타입을 생성했습니다. ") + T("The operation was stopped to avoid duplicate or suffixed system types. ", "중복 또는 접미사가 붙은 시스템 타입을 피하기 위해 작업을 중단했습니다. ") + T("Requested: ", "요청: ") + syncItem.SystemFamilyKind + " : " + syncItem.SourceTypeName + T(" / Created: ", " / 생성됨: ") + ((object)createdType).GetType().Name + " : " + ResolveElementName((Element)(object)createdType));
		}
		return createdType;
	}

	private static void ApplyStandardSystemTypeDefinition(Document targetDocument, Document standardDocument, ElementType sourceType, ElementType targetType, IEnumerable<RoutingDependencyPreflightItem> dependencyItems, SystemTypeApplyExecutionItem resultItem)
	{
		if (sourceType == null || targetType == null)
		{
			throw new InvalidOperationException(T("System type source or destination could not be resolved.", "시스템 타입 원본 또는 대상을 확인하지 못했습니다."));
		}
		CopyWritableTypeParameters(targetDocument, standardDocument, sourceType, targetType);
		ApplyRoutingPreferenceRules(targetDocument, standardDocument, sourceType, targetType, dependencyItems, resultItem);
		targetDocument.Regenerate();
		PostCheckSystemTypeAgainstStandard(targetDocument, standardDocument, sourceType, targetType, resultItem);
	}

	private static void CopyWritableTypeParameters(Document targetDocument, Document standardDocument, ElementType sourceType, ElementType targetType)
	{
		foreach (Parameter sourceParameter in ((IEnumerable)((Element)sourceType).Parameters).Cast<Parameter>())
		{
			if (sourceParameter != null && sourceParameter.Definition != null && sourceParameter.HasValue && !ShouldSkipTypeParameter(sourceParameter))
			{
				Parameter targetParameter = ResolveWritableParameter(targetType, sourceParameter);
				if (targetParameter != null)
				{
					TrySetParameterValue(targetDocument, standardDocument, targetParameter, sourceParameter);
				}
			}
		}
	}

	private static bool ShouldSkipTypeParameter(Parameter parameter)
	{
		try
		{
			if (parameter.Id != null && RevitElementIdCompat.CompatIntegerValue(parameter.Id) == -1002001)
			{
				return true;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		Definition definition = parameter.Definition;
		string name = ((definition != null) ? definition.Name : null) ?? string.Empty;
		return string.Equals(Normalize(name), "type name", StringComparison.Ordinal) || string.Equals(Normalize(name), "family name", StringComparison.Ordinal);
	}

	private static Parameter ResolveWritableParameter(ElementType targetType, Parameter sourceParameter)
	{
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
		//IL_0046: Unknown result type (might be due to invalid IL or missing references)
		Definition definition = sourceParameter.Definition;
		string name = ((definition != null) ? definition.Name : null) ?? string.Empty;
		if (string.IsNullOrWhiteSpace(name))
		{
			return null;
		}
		Parameter targetParameter = ((Element)targetType).LookupParameter(name);
		if (targetParameter == null || targetParameter.IsReadOnly)
		{
			return null;
		}
		if (targetParameter.StorageType != sourceParameter.StorageType)
		{
			return null;
		}
		return targetParameter;
	}

	private static void TrySetParameterValue(Document targetDocument, Document standardDocument, Parameter targetParameter, Parameter sourceParameter)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_0009: Unknown result type (might be due to invalid IL or missing references)
		//IL_001f: Expected I4, but got Unknown
		try
		{
			StorageType storageType = sourceParameter.StorageType;
			switch (storageType - 1)
			{
			case 2:
				targetParameter.Set(sourceParameter.AsString());
				break;
			case 1:
				targetParameter.Set(sourceParameter.AsDouble());
				break;
			case 0:
				targetParameter.Set(sourceParameter.AsInteger());
				break;
			case 3:
			{
				ElementId sourceId = sourceParameter.AsElementId();
				if (sourceId == null || sourceId == ElementId.InvalidElementId)
				{
					targetParameter.Set(ElementId.InvalidElementId);
					break;
				}
				if (RevitElementIdCompat.CompatIntegerValue(sourceId) < 0)
				{
					targetParameter.Set(sourceId);
					break;
				}
				Element sourceReference = standardDocument.GetElement(sourceId);
				if (sourceReference == null)
				{
					break;
				}
				if (sourceReference is Material || sourceReference is PipeScheduleType || sourceReference is PipeSegment)
				{
					ElementId mappedId = MapReferenceElementId(targetDocument, standardDocument, sourceId);
					if (mappedId == null || mappedId == ElementId.InvalidElementId)
					{
						throw new CriticalSystemTypeReferenceException("A referenced system definition could not be mapped into the current project. The system type apply was stopped to avoid a partial standard type. Parameter: " + ResolveParameterName(sourceParameter) + " / Reference: " + ((object)sourceReference).GetType().Name + " : " + ResolveElementName(sourceReference));
					}
					targetParameter.Set(mappedId);
					break;
				}
				Element targetReference = ResolveMatchingTargetElement(targetDocument, sourceReference);
				if (targetReference == null && !(sourceReference is FamilySymbol))
				{
					targetReference = CopyNonFamilyElementToTarget(targetDocument, standardDocument, sourceReference);
				}
				if (targetReference == null)
				{
					throw new CriticalSystemTypeReferenceException("A referenced system definition could not be mapped into the current project. The system type apply was stopped to avoid a partial standard type. Parameter: " + ResolveParameterName(sourceParameter) + " / Reference: " + ((object)sourceReference).GetType().Name + " : " + ResolveElementName(sourceReference));
				}
				EnsureMappedElementIdReferenceMatches(targetDocument, standardDocument, sourceReference, targetReference, sourceParameter);
				targetParameter.Set(targetReference.Id);
				break;
			}
			}
		}
		catch (CriticalSystemTypeReferenceException ex)
		{
			ProjectData.SetProjectError(ex);
			CriticalSystemTypeReferenceException ex2 = ex;
			throw;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			throw new InvalidOperationException("A writable system type parameter could not be copied from the standard RVT. The system type apply was stopped to avoid a partial standard type. Parameter: " + ResolveParameterName(sourceParameter), ex4);
		}
	}

	private static void EnsureMappedElementIdReferenceMatches(Document targetDocument, Document standardDocument, Element sourceReference, Element targetReference, Parameter sourceParameter)
	{
		if (sourceReference != null && targetReference != null && !(sourceReference is FamilySymbol) && !(sourceReference is Material) && !(sourceReference is PipeScheduleType) && !RoutingPartDefinitionsMatch(standardDocument, sourceReference, targetDocument, targetReference))
		{
			throw new CriticalSystemTypeReferenceException("A referenced system definition with the same name exists in the current project, but its definition differs from the standard RVT. The system type apply was stopped before using a non-standard reference. Parameter: " + ResolveParameterName(sourceParameter) + " / Reference: " + ((object)sourceReference).GetType().Name + " : " + ResolveElementName(sourceReference));
		}
	}

	private unsafe static void ApplyRoutingPreferenceRules(Document targetDocument, Document standardDocument, ElementType sourceType, ElementType targetType, IEnumerable<RoutingDependencyPreflightItem> dependencyItems, SystemTypeApplyExecutionItem resultItem)
	{
		//IL_009d: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a5: Invalid comparison between Unknown and I4
		//IL_00ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fb: Unknown result type (might be due to invalid IL or missing references)
		//IL_013e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0219: Unknown result type (might be due to invalid IL or missing references)
		//IL_0277: Unknown result type (might be due to invalid IL or missing references)
		//IL_027e: Expected O, but got Unknown
		//IL_0282: Unknown result type (might be due to invalid IL or missing references)
		//IL_0289: Unknown result type (might be due to invalid IL or missing references)
		RoutingPreferenceManager sourceManager = TryGetRoutingPreferenceManager(sourceType);
		RoutingPreferenceManager targetManager = TryGetRoutingPreferenceManager(targetType);
		if (sourceManager == null && targetManager == null)
		{
			return;
		}
		if (sourceManager == null || targetManager == null)
		{
			throw new InvalidOperationException("Routing preference manager availability does not match between the standard RVT and the current project. The system type apply was stopped to avoid creating a partial system type. Standard type: " + ((object)sourceType).GetType().Name + " : " + ResolveElementName((Element)(object)sourceType) + " / Target type: " + ((object)targetType).GetType().Name + " : " + ResolveElementName((Element)(object)targetType));
		}
		foreach (RoutingPreferenceRuleGroupType group in Enum.GetValues(typeof(RoutingPreferenceRuleGroupType)).Cast<RoutingPreferenceRuleGroupType>())
		{
			if ((int)group == -1)
			{
				continue;
			}
			int targetRuleCount;
			try
			{
				targetRuleCount = targetManager.GetNumberOfRules(group);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				throw new InvalidOperationException("Existing routing preference rules could not be inspected before replacement. The system type apply was stopped to avoid leaving mixed routing rules. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString(), ex2);
			}
			int num;
			checked
			{
				for (int index = targetRuleCount - 1; index >= 0; index += -1)
				{
					targetManager.RemoveRule(group, index);
				}
				int sourceRuleCount;
				try
				{
					sourceRuleCount = sourceManager.GetNumberOfRules(group);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					throw new InvalidOperationException("Standard routing preference rules could not be inspected. The system type apply was stopped before rebuilding the target type. Rule group: " + ((Enum)(*unchecked((RoutingPreferenceRuleGroupType*)(&group)))/*cast due to .constrained prefix*/).ToString(), ex4);
				}
				num = sourceRuleCount - 1;
			}
			for (int i = 0; i <= num; i = checked(i + 1))
			{
				RoutingPreferenceRule sourceRule = null;
				try
				{
					sourceRule = sourceManager.GetRule(group, i);
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					throw new InvalidOperationException("A standard routing preference rule could not be read. The system type apply was stopped before rebuilding the target type. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + i.ToString(CultureInfo.InvariantCulture), ex6);
				}
				if (sourceRule == null)
				{
					throw new InvalidOperationException("A standard routing preference rule returned no data. The system type apply was stopped before rebuilding the target type. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + i.ToString(CultureInfo.InvariantCulture));
				}
				if (!RoutingRuleHasMappablePart(sourceRule))
				{
					AddSystemApplyLog(resultItem, "SystemApply.RoutingRuleSkipped", "group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " index=" + i.ToString(CultureInfo.InvariantCulture) + " reason=InvalidMEPPartId");
					continue;
				}
				ElementId mappedPartId = MapRoutingPartId(targetDocument, standardDocument, sourceRule.MEPPartId, group, dependencyItems, resultItem);
				if (mappedPartId == null || mappedPartId == ElementId.InvalidElementId)
				{
					throw new InvalidOperationException("A standard routing preference rule could not be mapped into the current project. The system type apply was stopped before rebuilding the target type. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + i.ToString(CultureInfo.InvariantCulture));
				}
				RoutingPreferenceRule clonedRule = new RoutingPreferenceRule(mappedPartId, sourceRule.Description ?? string.Empty);
				CopyRoutingCriteria(sourceRule, clonedRule, group);
				targetManager.AddRule(group, clonedRule);
			}
		}
	}

	private unsafe static void PostCheckSystemTypeAgainstStandard(Document targetDocument, Document standardDocument, ElementType sourceType, ElementType targetType, SystemTypeApplyExecutionItem resultItem)
	{
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0064: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Unknown result type (might be due to invalid IL or missing references)
		//IL_0067: Invalid comparison between Unknown and I4
		//IL_006f: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Unknown result type (might be due to invalid IL or missing references)
		//IL_011a: Unknown result type (might be due to invalid IL or missing references)
		RoutingPreferenceManager sourceManager = TryGetRoutingPreferenceManager(sourceType);
		RoutingPreferenceManager targetManager = TryGetRoutingPreferenceManager(targetType);
		if (sourceManager == null && targetManager == null)
		{
			return;
		}
		if (sourceManager == null || targetManager == null)
		{
			throw new InvalidOperationException("Post-check failed because routing preference manager availability differs after apply. Standard type: " + ResolveElementName((Element)(object)sourceType) + " / Target type: " + ResolveElementName((Element)(object)targetType));
		}
		foreach (RoutingPreferenceRuleGroupType group in Enum.GetValues(typeof(RoutingPreferenceRuleGroupType)).Cast<RoutingPreferenceRuleGroupType>())
		{
			if ((int)group == -1)
			{
				continue;
			}
			List<RoutingRuleComparisonItem> expectedRules = BuildExpectedRoutingRulesForComparison(targetDocument, standardDocument, sourceManager, group, resultItem, "SystemApply.PostCheckRoutingRuleSkipped");
			int targetRuleCount = targetManager.GetNumberOfRules(group);
			if (expectedRules.Count != targetRuleCount)
			{
				throw new InvalidOperationException("Post-check failed because effective routing preference rule count differs from the registered standard. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Standard effective count: " + expectedRules.Count.ToString(CultureInfo.InvariantCulture) + " / Target count: " + targetRuleCount.ToString(CultureInfo.InvariantCulture));
			}
			int num = checked(expectedRules.Count - 1);
			for (int index = 0; index <= num; index = checked(index + 1))
			{
				RoutingRuleComparisonItem expectedRule = expectedRules[index];
				RoutingPreferenceRule sourceRule = expectedRule.Rule;
				RoutingPreferenceRule targetRule = targetManager.GetRule(group, index);
				if (sourceRule == null || targetRule == null)
				{
					throw new InvalidOperationException("Post-check failed because a routing preference rule could not be read after apply. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + index.ToString(CultureInfo.InvariantCulture) + " / Standard source index: " + expectedRule.SourceIndex.ToString(CultureInfo.InvariantCulture));
				}
				ElementId expectedTargetPartId = expectedRule.ExpectedTargetPartId;
				if (!ElementIdsEqual(expectedTargetPartId, targetRule.MEPPartId))
				{
					throw new InvalidOperationException("Post-check failed because routing preference rule order or mapped part differs from the registered standard. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + index.ToString(CultureInfo.InvariantCulture) + " / Standard source index: " + expectedRule.SourceIndex.ToString(CultureInfo.InvariantCulture) + " / Expected target part id: " + FormatElementId(expectedTargetPartId) + " / Actual target part id: " + FormatElementId(targetRule.MEPPartId));
				}
				Element sourcePart = ((sourceRule.MEPPartId == null || sourceRule.MEPPartId == ElementId.InvalidElementId) ? null : standardDocument.GetElement(sourceRule.MEPPartId));
				Element targetPart = ((targetRule.MEPPartId == null || targetRule.MEPPartId == ElementId.InvalidElementId) ? null : targetDocument.GetElement(targetRule.MEPPartId));
				if (sourcePart != null && !RoutingPartDefinitionMatchesForPostCheck(standardDocument, sourcePart, targetDocument, targetPart))
				{
					throw new InvalidOperationException(T("Post-check failed because a routing preference part definition differs from the registered standard. ", "사후 검증 실패: 라우팅 환경설정 부품 정의가 등록된 표준과 다릅니다. ") + T("Rule group: ", "규칙 그룹: ") + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + T(" / Rule index: ", " / 규칙 인덱스: ") + index.ToString(CultureInfo.InvariantCulture) + T(" / Standard source index: ", " / 표준 원본 인덱스: ") + expectedRule.SourceIndex.ToString(CultureInfo.InvariantCulture) + T(" / Part: ", " / 부품: ") + ((object)sourcePart).GetType().Name + " : " + ResolveElementName(sourcePart));
				}
				if (!string.Equals(sourceRule.Description ?? string.Empty, targetRule.Description ?? string.Empty, StringComparison.Ordinal))
				{
					throw new InvalidOperationException("Post-check failed because a routing preference rule description differs from the registered standard. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + index.ToString(CultureInfo.InvariantCulture) + " / Standard source index: " + expectedRule.SourceIndex.ToString(CultureInfo.InvariantCulture));
				}
				if (ResolveCriterionCount(sourceRule) != ResolveCriterionCount(targetRule))
				{
					throw new InvalidOperationException("Post-check failed because routing preference criterion count differs from the registered standard. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + index.ToString(CultureInfo.InvariantCulture));
				}
			}
		}
		AddResultMessage(resultItem, T("Post-check passed: routing preference rules and mapped parts match the registered standard.", "사후 검증 통과: 라우팅 환경설정 규칙과 매핑된 부품이 등록된 표준과 일치합니다."));
	}

	private unsafe static List<RoutingRuleComparisonItem> BuildExpectedRoutingRulesForComparison(Document targetDocument, Document standardDocument, RoutingPreferenceManager sourceManager, RoutingPreferenceRuleGroupType group, SystemTypeApplyExecutionItem resultItem, string skippedStage)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0022: Unknown result type (might be due to invalid IL or missing references)
		List<RoutingRuleComparisonItem> result = new List<RoutingRuleComparisonItem>();
		if (sourceManager == null)
		{
			return result;
		}
		int num = checked(sourceManager.GetNumberOfRules(group) - 1);
		for (int sourceIndex = 0; sourceIndex <= num; sourceIndex = checked(sourceIndex + 1))
		{
			RoutingPreferenceRule sourceRule = sourceManager.GetRule(group, sourceIndex);
			if (sourceRule == null)
			{
				throw new InvalidOperationException("A standard routing preference rule returned no data. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + sourceIndex.ToString(CultureInfo.InvariantCulture));
			}
			if (!RoutingRuleHasMappablePart(sourceRule))
			{
				AddSystemApplyLog(resultItem, skippedStage, "group=" + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " sourceIndex=" + sourceIndex.ToString(CultureInfo.InvariantCulture) + " reason=InvalidMEPPartId");
				continue;
			}
			ElementId expectedTargetPartId = ResolveExpectedRoutingPartIdForPostCheck(targetDocument, standardDocument, sourceRule.MEPPartId);
			if (expectedTargetPartId == null || expectedTargetPartId == ElementId.InvalidElementId)
			{
				throw new InvalidOperationException("A standard routing preference rule has a valid source part, but it could not be mapped into the current project. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + sourceIndex.ToString(CultureInfo.InvariantCulture) + " / Source part id: " + FormatElementId(sourceRule.MEPPartId));
			}
			result.Add(new RoutingRuleComparisonItem
			{
				SourceIndex = sourceIndex,
				Rule = sourceRule,
				ExpectedTargetPartId = expectedTargetPartId
			});
		}
		return result;
	}

	private static bool RoutingRuleHasMappablePart(RoutingPreferenceRule rule)
	{
		if (rule == null || rule.MEPPartId == null)
		{
			return false;
		}
		return rule.MEPPartId != ElementId.InvalidElementId && RevitElementIdCompat.CompatIntegerValue(rule.MEPPartId) > 0;
	}

	private static bool RoutingPartDefinitionMatchesForPostCheck(Document sourceDocument, Element sourcePart, Document targetDocument, Element targetPart)
	{
		if (sourcePart == null || targetPart == null)
		{
			return false;
		}
		if (sourcePart is FamilySymbol)
		{
			return targetPart is FamilySymbol;
		}
		PipeSegment sourceSegment = (PipeSegment)(object)((sourcePart is PipeSegment) ? sourcePart : null);
		if (sourceSegment != null)
		{
			PipeSegment targetSegment = (PipeSegment)(object)((targetPart is PipeSegment) ? targetPart : null);
			if (targetSegment == null)
			{
				return false;
			}
			ReferenceMappingCheck materialCheck = ResolvePotentialMappedReference(targetDocument, sourceDocument, ((Segment)sourceSegment).MaterialId, "pipe segment material");
			ReferenceMappingCheck scheduleCheck = ResolvePotentialMappedReference(targetDocument, sourceDocument, sourceSegment.ScheduleTypeId, "pipe segment schedule type");
			if (materialCheck.TargetReference == null || scheduleCheck.TargetReference == null)
			{
				return false;
			}
			return PipeSegmentDefinitionMatchesForRouting(sourceDocument, sourceSegment, targetDocument, targetSegment, materialCheck.TargetReference.Id, scheduleCheck.TargetReference.Id);
		}
		return RoutingPartDefinitionsMatch(sourceDocument, sourcePart, targetDocument, targetPart);
	}

	private static ElementId ResolveExpectedRoutingPartIdForPostCheck(Document targetDocument, Document standardDocument, ElementId sourcePartId)
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

	private static string FormatElementId(ElementId id)
	{
		if (id == null)
		{
			return "(null)";
		}
		return RevitElementIdCompat.CompatIntegerValue(id).ToString(CultureInfo.InvariantCulture);
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

	private unsafe static ElementId MapRoutingPartId(Document targetDocument, Document standardDocument, ElementId sourcePartId, RoutingPreferenceRuleGroupType ruleGroup, IEnumerable<RoutingDependencyPreflightItem> dependencyItems, SystemTypeApplyExecutionItem resultItem)
	{
		//IL_0069: Unknown result type (might be due to invalid IL or missing references)
		//IL_008b: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b0: Expected O, but got Unknown
		//IL_017d: Unknown result type (might be due to invalid IL or missing references)
		if (sourcePartId == null || sourcePartId == ElementId.InvalidElementId)
		{
			return ElementId.InvalidElementId;
		}
		Element sourcePart = standardDocument.GetElement(sourcePartId);
		if (sourcePart == null)
		{
			throw new InvalidOperationException(T("A routing preference part could not be resolved in the standard RVT: ", "표준 RVT에서 라우팅 환경설정 부품을 확인하지 못했습니다: ") + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&ruleGroup))/*cast due to .constrained prefix*/).ToString());
		}
		PipeSegment sourcePipeSegment = (PipeSegment)(object)((sourcePart is PipeSegment) ? sourcePart : null);
		if (sourcePipeSegment != null)
		{
			PipeSegment targetPipeSegment = EnsurePipeSegmentAuthoritative(targetDocument, standardDocument, sourcePipeSegment, SystemTypeApplyAuthorityMode.AdminAuthoritative, resultItem);
			EnsureRoutingPartDefinitionMatches(standardDocument, sourcePart, targetDocument, (Element)(object)targetPipeSegment, ruleGroup);
			return ((Element)targetPipeSegment).Id;
		}
		Element targetPart = ResolveMatchingTargetElement(targetDocument, sourcePart);
		if (targetPart != null)
		{
			EnsureRoutingPartDefinitionMatches(standardDocument, sourcePart, targetDocument, targetPart, ruleGroup);
			return targetPart.Id;
		}
		if (sourcePart is FamilySymbol)
		{
			FamilySymbol sourceSymbol = (FamilySymbol)sourcePart;
			if (targetDocument.IsModifiable)
			{
				string[] obj = new string[5]
				{
					T("The standard routing dependency family/type was not loaded before the system type transaction. ", "시스템 타입 트랜잭션 전에 표준 라우팅 의존 패밀리/타입이 로드되지 않았습니다. "),
					T("Family loading must occur before Transaction B. Family: ", "패밀리 로드는 Transaction B 전에 완료되어야 합니다. 패밀리: "),
					null,
					null,
					null
				};
				Family family = sourceSymbol.Family;
				obj[2] = ((family != null) ? ((Element)family).Name : null) ?? string.Empty;
				obj[3] = T(" / Type: ", " / 타입: ");
				obj[4] = ResolveElementName((Element)(object)sourceSymbol);
				throw new InvalidOperationException(string.Concat(obj));
			}
			return ((Element)EnsureRoutingFamilySymbolLoaded(targetDocument, standardDocument, sourceSymbol, dependencyItems, resultItem)).Id;
		}
		targetPart = CopyNonFamilyElementToTarget(targetDocument, standardDocument, sourcePart);
		if (targetPart == null)
		{
			throw new InvalidOperationException(T("A routing preference part could not be copied or matched in the target project: ", "대상 프로젝트에서 라우팅 환경설정 부품을 복사하거나 매칭하지 못했습니다: ") + ((object)sourcePart).GetType().Name + " : " + ResolveElementName(sourcePart));
		}
		EnsureRoutingPartDefinitionMatches(standardDocument, sourcePart, targetDocument, targetPart, ruleGroup);
		return targetPart.Id;
	}

	private unsafe static void EnsureRoutingPartDefinitionMatches(Document sourceDocument, Element sourcePart, Document targetDocument, Element targetPart, RoutingPreferenceRuleGroupType ruleGroup)
	{
		if (sourcePart == null || targetPart == null || sourcePart is FamilySymbol)
		{
			return;
		}
		PipeSegment sourceSegment = (PipeSegment)(object)((sourcePart is PipeSegment) ? sourcePart : null);
		if (sourceSegment != null)
		{
			PipeSegment targetSegment = (PipeSegment)(object)((targetPart is PipeSegment) ? targetPart : null);
			Element element = sourceDocument.GetElement(((Segment)sourceSegment).MaterialId);
			Material sourceMaterial = (Material)(object)((element is Material) ? element : null);
			Element element2 = sourceDocument.GetElement(sourceSegment.ScheduleTypeId);
			Element sourceSchedule = ((element2 is PipeScheduleType) ? element2 : null);
			ElementId targetMaterialId = EnsureMaterialForPipeSegment(sourceMaterial, sourceDocument, targetDocument, null).MaterialId;
			ElementId targetScheduleId = EnsurePipeScheduleTypeForPipeSegment((PipeScheduleType)(object)sourceSchedule, sourceDocument, targetDocument, null);
			if (!PipeSegmentDefinitionMatchesForRouting(sourceDocument, sourceSegment, targetDocument, targetSegment, targetMaterialId, targetScheduleId))
			{
				throw new InvalidOperationException(T("A routing preference pipe segment exists in the current project, but its material, schedule, size table, or managed routing values differ from the standard RVT. ", "현재 프로젝트에 라우팅 환경설정 배관 세그먼트가 있지만 재료, 스케줄, 사이즈 테이블 또는 관리 라우팅 값이 표준 RVT와 다릅니다. ") + T("Rule group: ", "규칙 그룹: ") + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&ruleGroup))/*cast due to .constrained prefix*/).ToString() + T(" / Segment: ", " / 세그먼트: ") + ResolveElementName((Element)(object)sourceSegment));
			}
		}
		else if (!RoutingPartDefinitionsMatch(sourceDocument, sourcePart, targetDocument, targetPart))
		{
			throw new InvalidOperationException(T("A routing preference part with the same name exists in the current project, but its definition differs from the standard RVT. ", "현재 프로젝트에 같은 이름의 라우팅 환경설정 부품이 있지만 정의가 표준 RVT와 다릅니다. ") + T("Review the segment or routing part before applying this system type. ", "이 시스템 타입을 적용하기 전에 세그먼트 또는 라우팅 부품을 검토하세요. ") + T("Rule group: ", "규칙 그룹: ") + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&ruleGroup))/*cast due to .constrained prefix*/).ToString() + T(" / Part: ", " / 부품: ") + ((object)sourcePart).GetType().Name + " : " + ResolveElementName(sourcePart));
		}
	}

	private static Element ResolveMatchingTargetElement(Document targetDocument, Element sourceElement)
	{
		//IL_0053: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_001f: Expected O, but got Unknown
		//IL_007c: Unknown result type (might be due to invalid IL or missing references)
		if (sourceElement == null)
		{
			return null;
		}
		if (sourceElement is FamilySymbol)
		{
			FamilySymbol sourceSymbol = (FamilySymbol)sourceElement;
			return (Element)(object)FindTargetFamilySymbol(targetDocument, sourceSymbol);
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

	private static FamilySymbol FindTargetFamilySymbol(Document targetDocument, FamilySymbol sourceSymbol)
	{
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
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

	private static string BuildRoutingFamilySymbolKey(FamilySymbol sourceSymbol)
	{
		if (sourceSymbol == null)
		{
			return string.Empty;
		}
		string[] array = new string[5];
		Family family = sourceSymbol.Family;
		array[0] = Normalize(((family != null) ? ((Element)family).Name : null) ?? string.Empty);
		array[1] = "|";
		array[2] = Normalize(ResolveElementName((Element)(object)sourceSymbol));
		array[3] = "|";
		array[4] = Normalize(ResolveCategoryName((Element)(object)sourceSymbol));
		return string.Concat(array);
	}

	private static bool ElementIdentityMatches(Element element, string sourceClassName, string sourceName, string sourceCategory)
	{
		if (element == null)
		{
			return false;
		}
		return string.Equals(((object)element).GetType().Name, sourceClassName, StringComparison.OrdinalIgnoreCase) && string.Equals(Normalize(ResolveElementName(element)), Normalize(sourceName), StringComparison.Ordinal) && CategoryNamesMatch(ResolveCategoryName(element), sourceCategory);
	}

	private static Element CopyNonFamilyElementToTarget(Document targetDocument, Document standardDocument, Element sourceElement)
	{
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Expected O, but got Unknown
		Element createdPipeSegment = TryCreatePipeSegmentInTarget(targetDocument, standardDocument, sourceElement);
		if (createdPipeSegment != null)
		{
			return createdPipeSegment;
		}
		ISet<int> familyStateBefore = CaptureFamilyNameState(targetDocument);
		Element firstCopiedElement = null;
		CopyPasteOptions options = new CopyPasteOptions();
		try
		{
			options.SetDuplicateTypeNamesHandler((IDuplicateTypeNamesHandler)(object)new CopyPasteUseDestinationTypesHandler());
			ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(standardDocument, (ICollection<ElementId>)new List<ElementId> { sourceElement.Id }, targetDocument, Transform.Identity, options);
			if (copiedIds != null)
			{
				foreach (ElementId copiedId in copiedIds)
				{
					Element copied = targetDocument.GetElement(copiedId);
					if (copied != null && firstCopiedElement == null)
					{
						firstCopiedElement = copied;
					}
				}
			}
		}
		finally
		{
			((IDisposable)options)?.Dispose();
		}
		targetDocument.Regenerate();
		GuardAgainstCopiedFamilies(targetDocument, familyStateBefore);
		if (firstCopiedElement != null && !ElementIdentityMatches(firstCopiedElement, ((object)sourceElement).GetType().Name, ResolveElementName(sourceElement), ResolveCategoryName(sourceElement)))
		{
			throw new CriticalSystemTypeReferenceException(T("A routing preference part was copied with a different identity than the standard RVT. ", "라우팅 환경설정 부품이 표준 RVT와 다른 식별자로 복사되었습니다. ") + T("This can create duplicate or suffixed routing parts, so the system type apply was stopped. ", "중복 또는 접미사 라우팅 부품이 생길 수 있어 시스템 타입 적용을 중단했습니다. ") + T("Standard: ", "표준: ") + ((object)sourceElement).GetType().Name + " : " + ResolveElementName(sourceElement) + T(" / Created: ", " / 생성됨: ") + ((object)firstCopiedElement).GetType().Name + " : " + ResolveElementName(firstCopiedElement));
		}
		return firstCopiedElement ?? ResolveMatchingTargetElement(targetDocument, sourceElement);
	}

	private static FamilySymbol EnsureRoutingFamilySymbolLoaded(Document targetDocument, Document standardDocument, FamilySymbol sourceSymbol, IEnumerable<RoutingDependencyPreflightItem> selectedDependencyItems, SystemTypeApplyExecutionItem resultItem)
	{
		if (sourceSymbol == null)
		{
			throw new InvalidOperationException(T("A routing preference fitting symbol could not be resolved in the registered standard RVT.", "등록된 표준 RVT에서 라우팅 환경설정 피팅 심볼을 확인하지 못했습니다."));
		}
		FamilySymbol targetSymbol = FindTargetFamilySymbol(targetDocument, sourceSymbol);
		if (targetSymbol != null)
		{
			return targetSymbol;
		}
		Family sourceFamily = sourceSymbol.Family;
		string familyName = ((sourceFamily != null) ? ((Element)sourceFamily).Name : null) ?? string.Empty;
		string typeName = ResolveElementName((Element)(object)sourceSymbol);
		if (sourceFamily == null)
		{
			throw new InvalidOperationException(T("A routing preference fitting symbol has no readable family in the registered standard RVT: ", "등록된 표준 RVT의 라우팅 환경설정 피팅 심볼에서 패밀리를 읽을 수 없습니다: ") + typeName);
		}
		if (sourceFamily.IsInPlace)
		{
			throw new InvalidOperationException(T("An in-place routing preference family cannot be loaded for system type apply: ", "내부 라우팅 환경설정 패밀리는 시스템 타입 적용을 위해 로드할 수 없습니다: ") + familyName);
		}
		if (!sourceFamily.IsEditable)
		{
			throw new InvalidOperationException(T("The routing preference family exists in the standard RVT but is not editable through the Revit API: ", "라우팅 환경설정 패밀리가 표준 RVT에 있지만 Revit API로 편집할 수 없습니다: ") + familyName);
		}
		Document familyDoc = null;
		try
		{
			ISet<int> familyStateBefore = CaptureFamilyNameState(targetDocument);
			familyDoc = standardDocument.EditFamily(sourceFamily);
			List<AllowedLoadedFamilyIdentity> allowedLoadedFamilies = BuildAllowedLoadedFamilyIdentities(sourceFamily, familyDoc, standardDocument, selectedDependencyItems, new List<string> { familyName });
			Family loadedFamily = familyDoc.LoadFamily(targetDocument, (IFamilyLoadOptions)(object)new LoadableFamilyLoadOptions(overwriteParameterValues: true));
			if (loadedFamily == null)
			{
				throw new InvalidOperationException(T("Revit returned no family reference after routing dependency load: ", "라우팅 의존 패밀리 로드 후 Revit이 패밀리 참조를 반환하지 않았습니다: ") + familyName);
			}
			RegenerateAfterFamilyLoad(targetDocument, resultItem, "RoutingFamilyLoad.Regenerated", familyName);
			GuardLoadedRoutingFamilyDidNotCreateDuplicateFamilies(targetDocument, familyStateBefore, sourceFamily, loadedFamily, allowedLoadedFamilies, resultItem);
		}
		finally
		{
			if (familyDoc != null)
			{
				try
				{
					familyDoc.Close(false);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
		targetSymbol = FindTargetFamilySymbol(targetDocument, sourceSymbol);
		if (targetSymbol == null)
		{
			throw new InvalidOperationException("A routing preference fitting family was loaded, but the exact standard family/type could not be resolved in the target project. Family: " + familyName + " / Type: " + typeName + " / Category: " + ResolveCategoryName((Element)(object)sourceSymbol));
		}
		AddResultMessage(resultItem, T("Apply-time fallback loaded missing routing family symbol: ", "적용 중 대체 처리로 누락된 라우팅 패밀리 심볼을 로드했습니다: ") + familyName + " : " + typeName);
		return targetSymbol;
	}

	private static Element TryCreatePipeSegmentInTarget(Document targetDocument, Document standardDocument, Element sourceElement)
	{
		PipeSegment sourceSegment = (PipeSegment)(object)((sourceElement is PipeSegment) ? sourceElement : null);
		if (sourceSegment == null)
		{
			return null;
		}
		return (Element)(object)EnsurePipeSegmentAuthoritative(targetDocument, standardDocument, sourceSegment, SystemTypeApplyAuthorityMode.AdminAuthoritative, null);
	}

	private static PipeSegment EnsurePipeSegmentAuthoritative(Document targetDocument, Document standardDocument, PipeSegment sourceSegment, SystemTypeApplyAuthorityMode authorityMode, SystemTypeApplyExecutionItem resultItem)
	{
		//IL_0265: Unknown result type (might be due to invalid IL or missing references)
		//IL_026b: Expected O, but got Unknown
		//IL_026d: Expected O, but got Unknown
		if (sourceSegment == null)
		{
			return null;
		}
		List<MEPSize> sizeCopies = CloneSegmentSizes((Segment)(object)sourceSegment);
		if (sizeCopies.Count == 0)
		{
			throw new InvalidOperationException("The standard pipe segment has no readable size table. The system type apply was stopped because pipe segment OD/ID/ND data cannot be verified: " + ResolveElementName((Element)(object)sourceSegment));
		}
		try
		{
			Element element = standardDocument.GetElement(((Segment)sourceSegment).MaterialId);
			Material sourceMaterial = (Material)(object)((element is Material) ? element : null);
			Element element2 = standardDocument.GetElement(sourceSegment.ScheduleTypeId);
			Element sourceSchedule = ((element2 is PipeScheduleType) ? element2 : null);
			ElementId targetMaterialId = EnsureMaterialForPipeSegment(sourceMaterial, standardDocument, targetDocument, resultItem).MaterialId;
			ElementId targetScheduleTypeId = EnsurePipeScheduleTypeForPipeSegment((PipeScheduleType)(object)sourceSchedule, standardDocument, targetDocument, resultItem);
			if (targetMaterialId == null || targetMaterialId == ElementId.InvalidElementId)
			{
				throw new InvalidOperationException("The standard pipe segment material could not be mapped into the current project: " + ResolveElementName((Element)(object)sourceSegment));
			}
			if (targetScheduleTypeId == null || targetScheduleTypeId == ElementId.InvalidElementId)
			{
				throw new InvalidOperationException("The standard pipe segment schedule type could not be mapped into the current project: " + ResolveElementName((Element)(object)sourceSegment));
			}
			Element obj = ResolveMatchingTargetElement(targetDocument, (Element)(object)sourceSegment);
			PipeSegment targetByName = (PipeSegment)(object)((obj is PipeSegment) ? obj : null);
			if (targetByName != null && ElementIdsEqual(((Segment)targetByName).MaterialId, targetMaterialId) && ElementIdsEqual(targetByName.ScheduleTypeId, targetScheduleTypeId))
			{
				SynchronizePipeSegmentDefinition(targetDocument, standardDocument, sourceSegment, targetByName, authorityMode, resultItem, targetMaterialId, targetScheduleTypeId);
				return targetByName;
			}
			PipeSegment existingByMaterialAndSchedule = FindPipeSegmentByMaterialAndSchedule(targetDocument, targetMaterialId, targetScheduleTypeId);
			if (existingByMaterialAndSchedule != null)
			{
				if (targetByName != null && !ElementIdsEqual(((Element)targetByName).Id, ((Element)existingByMaterialAndSchedule).Id))
				{
					ReportStalePipeSegment(targetDocument, targetByName, sourceSegment, resultItem);
				}
				SynchronizePipeSegmentDefinition(targetDocument, standardDocument, sourceSegment, existingByMaterialAndSchedule, authorityMode, resultItem, targetMaterialId, targetScheduleTypeId);
				return existingByMaterialAndSchedule;
			}
			if (targetByName != null)
			{
				if (TryDeleteUnusedStalePipeSegment(targetDocument, targetByName, sourceSegment, resultItem))
				{
					targetByName = null;
				}
				else
				{
					ReportStalePipeSegment(targetDocument, targetByName, sourceSegment, resultItem);
				}
			}
			PipeSegment created = PipeSegment.Create(targetDocument, targetMaterialId, targetScheduleTypeId, (ICollection<MEPSize>)sizeCopies);
			if (created == null)
			{
				throw new InvalidOperationException("Revit did not create a pipe segment from the standard size table: " + ResolveElementName((Element)(object)sourceSegment));
			}
			try
			{
				((Element)created).Name = ResolveElementName((Element)(object)sourceSegment);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			CopyWritableElementParameters(targetDocument, standardDocument, (Element)(object)sourceSegment, (Element)(object)created);
			SyncPipeSegmentManagedRoutingValues(sourceSegment, created);
			targetDocument.Regenerate();
			if (!string.Equals(Normalize(ResolveElementName((Element)(object)created)), Normalize(ResolveElementName((Element)(object)sourceSegment)), StringComparison.Ordinal))
			{
				AddResultMessage(resultItem, "PipeSegment rename was not exact after create, but routing will use the created segment because material, schedule, and size table are verified. Standard name: " + ResolveElementName((Element)(object)sourceSegment) + " / Created name: " + ResolveElementName((Element)(object)created));
			}
			if (!PipeSegmentDefinitionMatchesForRouting(standardDocument, sourceSegment, targetDocument, created, targetMaterialId, targetScheduleTypeId))
			{
				throw new InvalidOperationException("The standard pipe segment was created, but its final material, schedule type, size table, or managed routing values do not match the registered standard RVT. The Pipe Type apply was rolled back to avoid a partial standard system. Segment: " + ResolveElementName((Element)(object)sourceSegment));
			}
			AddResultMessage(resultItem, T("PipeSegment created from standard: ", "표준에서 배관 세그먼트를 생성했습니다: ") + ResolveElementName((Element)(object)sourceSegment));
			return created;
		}
		catch (CriticalSystemTypeReferenceException ex)
		{
			ProjectData.SetProjectError(ex);
			CriticalSystemTypeReferenceException ex2 = ex;
			throw;
		}
		catch (DisabledDisciplineException ex3)
		{
			ProjectData.SetProjectError((Exception)ex3);
			DisabledDisciplineException ex4 = ex3;
			throw new InvalidOperationException("This project does not have Mechanical/Electrical/Piping discipline enabled, so Pipe Type creation cannot continue. Segment: " + ResolveElementName((Element)(object)sourceSegment), (Exception)(object)ex4);
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			throw new InvalidOperationException("The standard pipe segment could not be rebuilt in the current project. The system type apply was stopped before falling back to copy/paste, because pipe segment OD/ID/ND data must stay explicit. Segment: " + ResolveElementName((Element)(object)sourceSegment), ex6);
		}
		finally
		{
			foreach (MEPSize size in sizeCopies)
			{
				try
				{
					size.Dispose();
				}
				catch (Exception projectError2)
				{
					ProjectData.SetProjectError(projectError2);
					ProjectData.ClearProjectError();
				}
			}
		}
	}

	private static void SynchronizePipeSegmentDefinition(Document targetDocument, Document standardDocument, PipeSegment sourceSegment, PipeSegment targetSegment, SystemTypeApplyAuthorityMode authorityMode, SystemTypeApplyExecutionItem resultItem, ElementId expectedTargetMaterialId, ElementId expectedTargetScheduleId)
	{
		if (sourceSegment == null || targetSegment == null)
		{
			throw new InvalidOperationException(T("Pipe segment synchronization could not resolve the source or target segment.", "배관 세그먼트 동기화에서 원본 또는 대상 세그먼트를 확인하지 못했습니다."));
		}
		if (authorityMode != SystemTypeApplyAuthorityMode.AdminAuthoritative && !PipeSegmentDefinitionMatchesForRouting(standardDocument, sourceSegment, targetDocument, targetSegment, expectedTargetMaterialId, expectedTargetScheduleId))
		{
			throw new InvalidOperationException(T("A project pipe segment uses the standard material/schedule identity, but its definition differs. ", "프로젝트 배관 세그먼트가 표준 재료/스케줄 식별자를 사용하지만 정의가 다릅니다. ") + T("Modeler-safe mode blocks this change until an administrator applies the authoritative standard. Segment: ", "관리자가 권위 표준을 적용하기 전까지 모델러 안전 모드에서 이 변경을 차단합니다. 세그먼트: ") + ResolveElementName((Element)(object)sourceSegment));
		}
		SyncPipeSegmentSizeTable(targetDocument, (Segment)(object)sourceSegment, (Segment)(object)targetSegment, authorityMode, resultItem);
		SyncPipeSegmentManagedRoutingValues(sourceSegment, targetSegment);
		CopyWritableElementParameters(targetDocument, standardDocument, (Element)(object)sourceSegment, (Element)(object)targetSegment);
		TryRenamePipeSegmentToStandard(targetDocument, sourceSegment, targetSegment, resultItem);
		targetDocument.Regenerate();
		if (!PipeSegmentDefinitionMatchesForRouting(standardDocument, sourceSegment, targetDocument, targetSegment, expectedTargetMaterialId, expectedTargetScheduleId))
		{
			throw new InvalidOperationException(T("A project pipe segment was synchronized, but its final material, schedule type, or size table still differs from the standard RVT. ", "프로젝트 배관 세그먼트를 동기화했지만 최종 재료, 스케줄 타입 또는 사이즈 테이블이 여전히 표준 RVT와 다릅니다. ") + T("The Pipe Type apply was rolled back to avoid a partial standard system. ", "부분 표준 시스템이 남지 않도록 배관 타입 적용을 롤백했습니다. ") + T("Standard segment: ", "표준 세그먼트: ") + ResolveElementName((Element)(object)sourceSegment) + T(" / Project segment: ", " / 프로젝트 세그먼트: ") + ResolveElementName((Element)(object)targetSegment));
		}
		AddResultMessage(resultItem, T("PipeSegment synchronized from standard: ", "표준에서 배관 세그먼트를 동기화했습니다: ") + ResolveElementName((Element)(object)sourceSegment) + " -> " + ResolveElementName((Element)(object)targetSegment));
	}

	private static void SyncPipeSegmentSizeTable(Document targetDocument, Segment sourceSegment, Segment targetSegment, SystemTypeApplyAuthorityMode authorityMode, SystemTypeApplyExecutionItem resultItem)
	{
		_Closure_0024__51_002D0 arg = default(_Closure_0024__51_002D0);
		_Closure_0024__51_002D0 CS_0024_003C_003E8__locals8 = new _Closure_0024__51_002D0(arg);
		CS_0024_003C_003E8__locals8._0024VB_0024Local_sourceSegment = sourceSegment;
		if (PipeSegmentSizeTablesMatch(CS_0024_003C_003E8__locals8._0024VB_0024Local_sourceSegment, targetSegment))
		{
			return;
		}
		if (authorityMode != SystemTypeApplyAuthorityMode.AdminAuthoritative)
		{
			throw new InvalidOperationException(T("The target pipe segment size table differs from the registered standard RVT and cannot be changed in modeler-safe mode.", "대상 배관 세그먼트 사이즈 테이블이 등록된 표준 RVT와 다르며 모델러 안전 모드에서는 변경할 수 없습니다."));
		}
		MethodInfo addSizeMethod = FindSegmentAddSizeMethod(targetSegment);
		MethodInfo removeSizeMethod = FindSegmentRemoveSizeMethod(targetSegment);
		List<MEPSize> sourceSizes = CloneSegmentSizes(CS_0024_003C_003E8__locals8._0024VB_0024Local_sourceSegment);
		try
		{
			if ((object)addSizeMethod == null)
			{
				List<string> missingOrDifferent = FindMissingOrDifferentStandardSizes(sourceSizes, targetSegment);
				if (missingOrDifferent.Count > 0)
				{
					throw new InvalidOperationException(T("The target pipe segment is missing or differs from standard sizes, but this Revit API version does not expose AddSize(MEPSize). ", "대상 배관 세그먼트에 표준 사이즈가 없거나 다르지만 이 Revit API 버전은 AddSize(MEPSize)를 제공하지 않습니다. ") + T("The Pipe Type apply was stopped before leaving a partial segment. Missing or different nominal diameters: ", "부분 세그먼트가 남지 않도록 배관 타입 적용을 중단했습니다. 누락 또는 상이한 공칭 직경: ") + string.Join(", ", missingOrDifferent));
				}
			}
			foreach (MEPSize sourceSize in sourceSizes)
			{
				string sourceNdKey = BuildMEPSizeNominalDiameterKey(sourceSize);
				MEPSize targetSize = FindTargetSizeByNominalDiameter(targetSegment, sourceSize.NominalDiameter);
				if (targetSize == null)
				{
					AddPipeSegmentSize(targetSegment, addSizeMethod, sourceSize, sourceNdKey);
					targetDocument.Regenerate();
				}
				else if (!MEPSizeDefinitionsMatch(sourceSize, targetSize))
				{
					if ((object)removeSizeMethod == null)
					{
						throw new InvalidOperationException(T("The target pipe segment has a size with the same nominal diameter as the standard, but ID/OD/flags differ and this Revit API version does not expose RemoveSize(Double). ", "대상 배관 세그먼트에 표준과 같은 공칭 직경의 사이즈가 있지만 ID/OD/플래그가 다르고 이 Revit API 버전은 RemoveSize(Double)를 제공하지 않습니다. ") + T("The Pipe Type apply was stopped because the standard size cannot be replaced. Nominal diameter: ", "표준 사이즈를 교체할 수 없어 배관 타입 적용을 중단했습니다. 공칭 직경: ") + sourceNdKey);
					}
					ReplacePipeSegmentSize(targetDocument, CS_0024_003C_003E8__locals8._0024VB_0024Local_sourceSegment, targetSegment, addSizeMethod, removeSizeMethod, sourceSize, targetSize, resultItem);
				}
			}
			if (!PipeSegmentContainsStandardSizes(CS_0024_003C_003E8__locals8._0024VB_0024Local_sourceSegment, targetSegment))
			{
				throw new InvalidOperationException(T("The target pipe segment size table is still missing one or more standard source sizes after add/replace synchronization.", "추가/교체 동기화 후에도 대상 배관 세그먼트 사이즈 테이블에 하나 이상의 표준 원본 사이즈가 없습니다."));
			}
			List<MEPSize> extraTargetSizes = (from x in targetSegment.GetSizes()
				where x != null && !SegmentContainsNominalDiameter(CS_0024_003C_003E8__locals8._0024VB_0024Local_sourceSegment, x.NominalDiameter)
				select x).ToList();
			if (extraTargetSizes.Count > 0)
			{
				if ((object)removeSizeMethod == null)
				{
					AddResultMessage(resultItem, "Admin cleanup warning: extra pipe segment sizes remain because this Revit API version does not expose RemoveSize(Double). Segment: " + ResolveElementName((Element)(object)targetSegment) + " / Extra nominal diameter count: " + extraTargetSizes.Count.ToString(CultureInfo.InvariantCulture));
				}
				else
				{
					foreach (MEPSize targetSize2 in extraTargetSizes)
					{
						string targetNdKey = BuildMEPSizeNominalDiameterKey(targetSize2);
						Exception removeError = null;
						if (!TryRemovePipeSegmentSize(targetSegment, removeSizeMethod, targetSize2.NominalDiameter, ref removeError))
						{
							AddResultMessage(resultItem, "Admin cleanup warning: an extra pipe segment size could not be removed and will remain for strict cleanup review. Segment: " + ResolveElementName((Element)(object)targetSegment) + " / Nominal diameter: " + targetNdKey + " / Reason: " + ResolveExceptionMessage(removeError));
						}
					}
					targetDocument.Regenerate();
				}
			}
			if (!PipeSegmentContainsStandardSizes(CS_0024_003C_003E8__locals8._0024VB_0024Local_sourceSegment, targetSegment))
			{
				throw new InvalidOperationException("The target pipe segment size table lost one or more standard source sizes during extra-size cleanup.");
			}
			int remainingExtraCount = CountExtraPipeSegmentSizes(CS_0024_003C_003E8__locals8._0024VB_0024Local_sourceSegment, targetSegment);
			if (remainingExtraCount > 0)
			{
				AddResultMessage(resultItem, "Admin cleanup warning: selected Pipe Type apply kept extra pipe segment sizes after ensuring all standard sizes. Strict exact size-table cleanup should be run separately. Segment: " + ResolveElementName((Element)(object)targetSegment) + " / Extra nominal diameter count: " + remainingExtraCount.ToString(CultureInfo.InvariantCulture));
			}
		}
		catch (TargetInvocationException ex)
		{
			ProjectData.SetProjectError(ex);
			TargetInvocationException ex2 = ex;
			throw new InvalidOperationException(T("The target pipe segment size table could not be synchronized.", "대상 배관 세그먼트 사이즈 테이블을 동기화하지 못했습니다."), ex2.InnerException);
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			throw new InvalidOperationException(T("The target pipe segment size table could not be synchronized.", "대상 배관 세그먼트 사이즈 테이블을 동기화하지 못했습니다."), ex4);
		}
		finally
		{
			foreach (MEPSize sourceSize2 in sourceSizes)
			{
				try
				{
					sourceSize2.Dispose();
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
	}

	private static void AddPipeSegmentSize(Segment targetSegment, MethodInfo addSizeMethod, MEPSize sourceSize, string nominalDiameterKey)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		if ((object)addSizeMethod == null)
		{
			throw new InvalidOperationException("The target pipe segment is missing a standard size, but this Revit API version does not expose AddSize(MEPSize). Nominal diameter: " + nominalDiameterKey);
		}
		MEPSize addSize = new MEPSize(sourceSize.NominalDiameter, sourceSize.InnerDiameter, sourceSize.OuterDiameter, sourceSize.UsedInSizeLists, sourceSize.UsedInSizing);
		try
		{
			addSizeMethod.Invoke(targetSegment, new object[1] { addSize });
		}
		catch (TargetInvocationException ex)
		{
			ProjectData.SetProjectError(ex);
			TargetInvocationException ex2 = ex;
			throw new InvalidOperationException("The standard pipe segment size could not be added or replaced in the target project. Nominal diameter: " + nominalDiameterKey, ex2.InnerException);
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			throw new InvalidOperationException("The standard pipe segment size could not be added or replaced in the target project. Nominal diameter: " + nominalDiameterKey, ex4);
		}
		finally
		{
			try
			{
				addSize.Dispose();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
	}

	private static void ReplacePipeSegmentSize(Document targetDocument, Segment sourceSegment, Segment targetSegment, MethodInfo addSizeMethod, MethodInfo removeSizeMethod, MEPSize sourceSize, MEPSize targetSize, SystemTypeApplyExecutionItem resultItem)
	{
		string sourceNdKey = BuildMEPSizeNominalDiameterKey(sourceSize);
		if (sourceSize == null || targetSize == null)
		{
			throw new InvalidOperationException(T("The target pipe segment standard-size replacement could not resolve the source or target size.", "대상 배관 세그먼트 표준 사이즈 교체에서 원본 또는 대상 사이즈를 확인하지 못했습니다."));
		}
		double targetNominalDiameter = targetSize.NominalDiameter;
		MEPSize temporarySize = null;
		double? temporaryNominalDiameter = null;
		try
		{
			if (CountReadablePipeSegmentSizes(targetSegment) <= 1)
			{
				temporarySize = BuildTemporaryPipeSegmentSize(sourceSize, sourceSegment, targetSegment);
				temporaryNominalDiameter = temporarySize.NominalDiameter;
				AddPipeSegmentSize(targetSegment, addSizeMethod, temporarySize, BuildMEPSizeNominalDiameterKey(temporarySize));
				targetDocument.Regenerate();
			}
			RemovePipeSegmentSizeRequired(targetSegment, removeSizeMethod, targetNominalDiameter, "The target pipe segment standard-size replacement failed because the existing same-nominal-diameter size could not be removed. Nominal diameter: " + sourceNdKey);
			targetDocument.Regenerate();
			AddPipeSegmentSize(targetSegment, addSizeMethod, sourceSize, sourceNdKey);
			targetDocument.Regenerate();
			if (temporaryNominalDiameter.HasValue)
			{
				Exception removeError = null;
				if (!TryRemovePipeSegmentSize(targetSegment, removeSizeMethod, temporaryNominalDiameter.Value, ref removeError))
				{
					AddResultMessage(resultItem, "Admin cleanup warning: temporary pipe segment size could not be removed after all standard sizes were restored. Selected Pipe Type apply may continue because routing uses the verified standard segment; strict exact size-table cleanup is still required. Segment: " + ResolveElementName((Element)(object)targetSegment) + " / Temporary nominal diameter: " + FormatLength(temporaryNominalDiameter.Value) + " / Reason: " + ResolveExceptionMessage(removeError));
				}
				else
				{
					targetDocument.Regenerate();
				}
			}
		}
		finally
		{
			if (temporarySize != null)
			{
				try
				{
					temporarySize.Dispose();
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
	}

	private static void RemovePipeSegmentSizeRequired(Segment targetSegment, MethodInfo removeSizeMethod, double nominalDiameter, string failureMessage)
	{
		Exception removeError = null;
		if (!TryRemovePipeSegmentSize(targetSegment, removeSizeMethod, nominalDiameter, ref removeError))
		{
			throw new InvalidOperationException(failureMessage, removeError);
		}
	}

	private static bool TryRemovePipeSegmentSize(Segment targetSegment, MethodInfo removeSizeMethod, double nominalDiameter, ref Exception errorResult)
	{
		bool TryRemovePipeSegmentSize;
		if (targetSegment == null || (object)removeSizeMethod == null)
		{
			errorResult = new InvalidOperationException("RemoveSize(Double) is not available for this pipe segment.");
			TryRemovePipeSegmentSize = false;
		}
		else
		{
			try
			{
				removeSizeMethod.Invoke(targetSegment, new object[1] { nominalDiameter });
				errorResult = null;
				TryRemovePipeSegmentSize = true;
			}
			catch (TargetInvocationException ex)
			{
				ProjectData.SetProjectError(ex);
				TargetInvocationException ex2 = ex;
				errorResult = ex2.InnerException ?? ex2;
				TryRemovePipeSegmentSize = false;
				ProjectData.ClearProjectError();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				errorResult = ex4;
				TryRemovePipeSegmentSize = false;
				ProjectData.ClearProjectError();
			}
		}
		return TryRemovePipeSegmentSize;
	}

	private static int CountReadablePipeSegmentSizes(Segment segment)
	{
		int CountReadablePipeSegmentSizes;
		if (segment == null)
		{
			CountReadablePipeSegmentSizes = 0;
		}
		else
		{
			try
			{
				CountReadablePipeSegmentSizes = (from x in segment.GetSizes()
					where x != null
					select x).Count();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				CountReadablePipeSegmentSizes = 0;
				ProjectData.ClearProjectError();
			}
		}
		return CountReadablePipeSegmentSizes;
	}

	private static bool PipeSegmentHasReadableSizes(Segment segment)
	{
		return CountReadablePipeSegmentSizes(segment) > 0;
	}

	private static MEPSize BuildTemporaryPipeSegmentSize(MEPSize sourceSize, Segment sourceSegment, Segment targetSegment)
	{
		//IL_00f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fe: Expected O, but got Unknown
		if (sourceSize == null)
		{
			throw new InvalidOperationException(T("A temporary pipe segment size could not be created because the source size is unreadable.", "원본 사이즈를 읽을 수 없어 임시 배관 세그먼트 사이즈를 만들지 못했습니다."));
		}
		double sourceNominalDiameter = ((sourceSize.NominalDiameter > 1E-09) ? sourceSize.NominalDiameter : 0.01);
		double[] array = new double[12]
		{
			0.001, 0.002, 0.005, 0.01, 0.02, 0.05, 0.1, 0.2, 0.5, 1.0,
			2.0, 5.0
		};
		foreach (double offset in array)
		{
			double[] array2 = new double[2] { 1.0, -1.0 };
			foreach (double direction in array2)
			{
				double candidateNominalDiameter = sourceNominalDiameter + offset * direction;
				if (!(candidateNominalDiameter <= 1E-09) && !SegmentContainsNominalDiameter(sourceSegment, candidateNominalDiameter) && !SegmentContainsNominalDiameter(targetSegment, candidateNominalDiameter))
				{
					double candidateInnerDiameter = Math.Max(1E-09, candidateNominalDiameter * 0.8);
					double candidateOuterDiameter = Math.Max(candidateNominalDiameter * 1.1, candidateInnerDiameter + 1E-09);
					return new MEPSize(candidateNominalDiameter, candidateInnerDiameter, candidateOuterDiameter, false, false);
				}
			}
		}
		throw new InvalidOperationException("A temporary pipe segment size could not be generated because every candidate nominal diameter already exists in the source or target size table.");
	}

	private static List<string> FindMissingOrDifferentStandardSizes(IEnumerable<MEPSize> sourceSizes, Segment targetSegment)
	{
		List<string> result = new List<string>();
		foreach (MEPSize sourceSize in sourceSizes ?? new List<MEPSize>())
		{
			if (sourceSize != null)
			{
				MEPSize targetSize = FindTargetSizeByNominalDiameter(targetSegment, sourceSize.NominalDiameter);
				if (targetSize == null || !MEPSizeDefinitionsMatch(sourceSize, targetSize))
				{
					result.Add(BuildMEPSizeNominalDiameterKey(sourceSize));
				}
			}
		}
		return result.Distinct(StringComparer.Ordinal).OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal).ToList();
	}

	private static MEPSize FindTargetSizeByNominalDiameter(Segment segment, double nominalDiameter)
	{
		MEPSize FindTargetSizeByNominalDiameter;
		if (segment == null)
		{
			FindTargetSizeByNominalDiameter = null;
		}
		else
		{
			try
			{
				FindTargetSizeByNominalDiameter = (from x in segment.GetSizes()
					where x != null
					select x).FirstOrDefault([SpecialName] (MEPSize x) => LengthValuesMatch(x.NominalDiameter, nominalDiameter));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				FindTargetSizeByNominalDiameter = null;
				ProjectData.ClearProjectError();
			}
		}
		return FindTargetSizeByNominalDiameter;
	}

	private static bool MEPSizeDefinitionsMatch(MEPSize left, MEPSize right)
	{
		if (left == null || right == null)
		{
			return left == null && right == null;
		}
		return LengthValuesMatch(left.NominalDiameter, right.NominalDiameter) && LengthValuesMatch(left.InnerDiameter, right.InnerDiameter) && LengthValuesMatch(left.OuterDiameter, right.OuterDiameter) && left.UsedInSizeLists == right.UsedInSizeLists && left.UsedInSizing == right.UsedInSizing;
	}

	private static MethodInfo FindSegmentAddSizeMethod(Segment segment)
	{
		return ((object)segment)?.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public).FirstOrDefault([SpecialName] (MethodInfo x) => string.Equals(x.Name, "AddSize", StringComparison.Ordinal) && x.GetParameters().Length == 1 && typeof(MEPSize).IsAssignableFrom(x.GetParameters()[0].ParameterType));
	}

	private static MethodInfo FindSegmentRemoveSizeMethod(Segment segment)
	{
		return ((object)segment)?.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public).FirstOrDefault([SpecialName] (MethodInfo x) => string.Equals(x.Name, "RemoveSize", StringComparison.Ordinal) && x.GetParameters().Length == 1 && x.GetParameters()[0].ParameterType == typeof(double));
	}

	private static void SyncPipeSegmentManagedRoutingValues(PipeSegment sourceSegment, PipeSegment targetSegment)
	{
		TryCopyWritableProperty(sourceSegment, targetSegment, "Roughness");
	}

	private static void TryRenamePipeSegmentToStandard(Document targetDocument, PipeSegment sourceSegment, PipeSegment targetSegment, SystemTypeApplyExecutionItem resultItem)
	{
		string sourceName = ResolveElementName((Element)(object)sourceSegment);
		if (string.IsNullOrWhiteSpace(sourceName) || string.Equals(Normalize(sourceName), Normalize(ResolveElementName((Element)(object)targetSegment)), StringComparison.Ordinal))
		{
			return;
		}
		Element obj = ResolveMatchingTargetElement(targetDocument, (Element)(object)sourceSegment);
		PipeSegment sameNameSegment = (PipeSegment)(object)((obj is PipeSegment) ? obj : null);
		if (sameNameSegment != null && !ElementIdsEqual(((Element)sameNameSegment).Id, ((Element)targetSegment).Id))
		{
			AddResultMessage(resultItem, "Stale pipe segment keeps the standard name, so the material/schedule counterpart was used without renaming. Admin cleanup is required if the stale segment is still referenced. Standard name: " + sourceName + " / Used segment: " + ResolveElementName((Element)(object)targetSegment));
			return;
		}
		try
		{
			((Element)targetSegment).Name = sourceName;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			AddResultMessage(resultItem, "PipeSegment rename was skipped by Revit, but routing was reconnected to the correct material/schedule counterpart. Segment: " + ResolveElementName((Element)(object)targetSegment) + " / Standard name: " + sourceName + " / Reason: " + ex2.Message);
			ProjectData.ClearProjectError();
		}
	}

	private static bool TryDeleteUnusedStalePipeSegment(Document targetDocument, PipeSegment staleSegment, PipeSegment sourceSegment, SystemTypeApplyExecutionItem resultItem)
	{
		bool TryDeleteUnusedStalePipeSegment;
		if (staleSegment == null || targetDocument == null)
		{
			TryDeleteUnusedStalePipeSegment = false;
		}
		else if (IsPipeSegmentReferencedByRoutingPreferences(targetDocument, ((Element)staleSegment).Id))
		{
			TryDeleteUnusedStalePipeSegment = false;
		}
		else
		{
			try
			{
				string staleName = ResolveElementName((Element)(object)staleSegment);
				targetDocument.Delete(((Element)staleSegment).Id);
				targetDocument.Regenerate();
				AddResultMessage(resultItem, "Unused stale pipe segment deleted before authoritative sync: " + staleName + " / Standard segment: " + ResolveElementName((Element)(object)sourceSegment));
				TryDeleteUnusedStalePipeSegment = true;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				TryDeleteUnusedStalePipeSegment = false;
				ProjectData.ClearProjectError();
			}
		}
		return TryDeleteUnusedStalePipeSegment;
	}

	private static bool IsPipeSegmentReferencedByRoutingPreferences(Document targetDocument, ElementId segmentId)
	{
		//IL_00b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c2: Invalid comparison between Unknown and I4
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ea: Unknown result type (might be due to invalid IL or missing references)
		if (targetDocument == null || segmentId == null || segmentId == ElementId.InvalidElementId)
		{
			return false;
		}
		foreach (ElementType item in from Element x in (IEnumerable)new FilteredElementCollector(targetDocument).WhereElementIsElementType()
			select (ElementType)(object)((x is ElementType) ? x : null) into x
			where x != null
			select x)
		{
			RoutingPreferenceManager manager = TryGetRoutingPreferenceManager(item);
			if (manager == null)
			{
				continue;
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
						try
						{
							RoutingPreferenceRule rule = manager.GetRule(group, index);
							if (rule != null && ElementIdsEqual(rule.MEPPartId, segmentId))
							{
								return true;
							}
						}
						catch (Exception projectError2)
						{
							ProjectData.SetProjectError(projectError2);
							ProjectData.ClearProjectError();
						}
					}
				}
			}
		}
		return false;
	}

	private static void ReportStalePipeSegment(Document targetDocument, PipeSegment staleSegment, PipeSegment sourceSegment, SystemTypeApplyExecutionItem resultItem)
	{
		if (staleSegment != null)
		{
			string usage = (IsPipeSegmentReferencedByRoutingPreferences(targetDocument, ((Element)staleSegment).Id) ? "referenced" : "unreferenced");
			AddResultMessage(resultItem, "Stale same-name pipe segment detected and not used for this Pipe Type routing. The authoritative material/schedule counterpart will be used instead. Standard segment: " + ResolveElementName((Element)(object)sourceSegment) + " / Stale project segment: " + ResolveElementName((Element)(object)staleSegment) + " / Stale usage: " + usage);
		}
	}

	private static List<MEPSize> CloneSegmentSizes(Segment sourceSegment)
	{
		//IL_005d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0067: Expected O, but got Unknown
		List<MEPSize> sizes = new List<MEPSize>();
		if (sourceSegment == null)
		{
			return sizes;
		}
		try
		{
			foreach (MEPSize sourceSize in sourceSegment.GetSizes())
			{
				if (sourceSize == null)
				{
					throw new InvalidOperationException(T("The pipe segment size table contains an unreadable size row.", "배관 세그먼트 사이즈 테이블에 읽을 수 없는 사이즈 행이 있습니다."));
				}
				sizes.Add(new MEPSize(sourceSize.NominalDiameter, sourceSize.InnerDiameter, sourceSize.OuterDiameter, sourceSize.UsedInSizeLists, sourceSize.UsedInSizing));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			throw new InvalidOperationException("The standard pipe segment size table could not be read. The system type apply was stopped because OD/ID/ND data cannot be trusted.", ex2);
		}
		return sizes;
	}

	private static List<NonFamilyRoutingDependencyCheck> ValidateNonFamilyRoutingDependencies(Document targetDocument, Document standardDocument, IDictionary<string, ElementType> standardSystemMap, SystemTypeSyncPlanItem syncItem, SystemSyncExecutionItem executionPlanItem)
	{
		//IL_00b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bc: Invalid comparison between Unknown and I4
		//IL_00c3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0136: Unknown result type (might be due to invalid IL or missing references)
		List<NonFamilyRoutingDependencyCheck> result = new List<NonFamilyRoutingDependencyCheck>();
		if (targetDocument == null || standardDocument == null || syncItem == null || executionPlanItem == null)
		{
			return result;
		}
		string left = Normalize(syncItem.Action);
		if (Operators.CompareString(left, "createmissingtype", TextCompare: false) != 0 && Operators.CompareString(left, "overwritedestination", TextCompare: false) != 0)
		{
			return result;
		}
		ElementType sourceType = ResolveStandardSystemType(standardSystemMap, syncItem);
		if (sourceType == null)
		{
			result.Add(NonFamilyRoutingDependencyCheck.Block("ValidateRoutingPart", "The source system type was not found in the registered standard RVT before dependency family load: " + syncItem.SourceTypeName));
			return result;
		}
		RoutingPreferenceManager manager = TryGetRoutingPreferenceManager(sourceType);
		if (manager == null)
		{
			return result;
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
						result.AddRange(BuildNonFamilyRoutingPartChecks(targetDocument, standardDocument, sourcePart, group, index));
					}
				}
			}
		}
		return result;
	}

	private unsafe static IEnumerable<NonFamilyRoutingDependencyCheck> BuildNonFamilyRoutingPartChecks(Document targetDocument, Document standardDocument, Element sourcePart, RoutingPreferenceRuleGroupType group, int ruleIndex)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		List<NonFamilyRoutingDependencyCheck> result = new List<NonFamilyRoutingDependencyCheck>();
		PipeSegment sourceSegment = (PipeSegment)(object)((sourcePart is PipeSegment) ? sourcePart : null);
		if (sourceSegment != null)
		{
			result.Add(BuildPipeSegmentRoutingPartCheck(targetDocument, standardDocument, sourceSegment, group, ruleIndex));
			return result;
		}
		Element targetPart = ResolveMatchingTargetElement(targetDocument, sourcePart);
		if (targetPart == null)
		{
			result.Add(NonFamilyRoutingDependencyCheck.Block("MissingRoutingPart", "A non-family routing preference part is missing in the current project and does not yet have an explicit safe ensure path. Apply was stopped before dependency family load. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + ruleIndex.ToString(CultureInfo.InvariantCulture) + " / Part: " + ((object)sourcePart).GetType().Name + " : " + ResolveElementName(sourcePart)));
			return result;
		}
		if (!RoutingPartDefinitionsMatch(standardDocument, sourcePart, targetDocument, targetPart))
		{
			result.Add(NonFamilyRoutingDependencyCheck.Block("RoutingPartConflict", "A non-family routing preference part exists in the current project, but its definition differs from the standard RVT. Apply was stopped before dependency family load. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + ruleIndex.ToString(CultureInfo.InvariantCulture) + " / Part: " + ((object)sourcePart).GetType().Name + " : " + ResolveElementName(sourcePart)));
			return result;
		}
		result.Add(NonFamilyRoutingDependencyCheck.Ready("ReuseRoutingPart", "Routing preference part is ready to reuse. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + ruleIndex.ToString(CultureInfo.InvariantCulture) + " / Part: " + ((object)sourcePart).GetType().Name + " : " + ResolveElementName(sourcePart)));
		return result;
	}

	private unsafe static NonFamilyRoutingDependencyCheck BuildPipeSegmentRoutingPartCheck(Document targetDocument, Document standardDocument, PipeSegment sourceSegment, RoutingPreferenceRuleGroupType group, int ruleIndex)
	{
		if (!PipeSegmentHasReadableSizes((Segment)(object)sourceSegment))
		{
			return NonFamilyRoutingDependencyCheck.Block("PipeSegmentSizeTableUnreadable", "The standard pipe segment size table could not be read before dependency family load. Apply was stopped because the target segment cannot be verified. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + ruleIndex.ToString(CultureInfo.InvariantCulture) + " / Segment: " + ResolveElementName((Element)(object)sourceSegment));
		}
		Element targetByName = ResolveMatchingTargetElement(targetDocument, (Element)(object)sourceSegment);
		if (targetByName != null && RoutingPartDefinitionsMatch(standardDocument, (Element)(object)sourceSegment, targetDocument, targetByName))
		{
			return NonFamilyRoutingDependencyCheck.Ready("ReusePipeSegment", "Pipe segment is ready to reuse. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + ruleIndex.ToString(CultureInfo.InvariantCulture) + " / Segment: " + ResolveElementName((Element)(object)sourceSegment));
		}
		ReferenceMappingCheck materialCheck = ResolvePotentialMappedReference(targetDocument, standardDocument, ((Segment)sourceSegment).MaterialId, "pipe segment material");
		if (materialCheck.SourceReference == null)
		{
			return NonFamilyRoutingDependencyCheck.Block("PipeSegmentMaterialConflict", (materialCheck.Blocker ?? "The standard pipe segment material could not be resolved.") + " Apply was stopped before dependency family load. Segment: " + ResolveElementName((Element)(object)sourceSegment));
		}
		ReferenceMappingCheck scheduleCheck = ResolvePotentialMappedReference(targetDocument, standardDocument, sourceSegment.ScheduleTypeId, "pipe segment schedule type");
		if (scheduleCheck.SourceReference == null)
		{
			return NonFamilyRoutingDependencyCheck.Block("PipeSegmentScheduleConflict", (scheduleCheck.Blocker ?? "The standard pipe segment schedule type could not be resolved.") + " Apply was stopped before dependency family load. Segment: " + ResolveElementName((Element)(object)sourceSegment));
		}
		if (materialCheck.TargetReference != null && scheduleCheck.TargetReference != null)
		{
			PipeSegment existingByCombination = FindPipeSegmentByMaterialAndSchedule(targetDocument, materialCheck.TargetReference.Id, scheduleCheck.TargetReference.Id);
			if (existingByCombination != null)
			{
				if (RoutingPartDefinitionsMatch(standardDocument, (Element)(object)sourceSegment, targetDocument, (Element)(object)existingByCombination))
				{
					return NonFamilyRoutingDependencyCheck.Ready("ReusePipeSegmentByMaterialSchedule", "Pipe segment will reuse an existing project segment with matching material, schedule type, and size table. Standard segment: " + ResolveElementName((Element)(object)sourceSegment) + " / Project segment: " + ResolveElementName((Element)(object)existingByCombination));
				}
				return NonFamilyRoutingDependencyCheck.Ready("SyncPipeSegmentByMaterialSchedule", "A project pipe segment already uses the standard material and schedule type, but its size table or writable definition differs. Admin-authoritative apply will synchronize that existing segment instead of creating a duplicate. Standard segment: " + ResolveElementName((Element)(object)sourceSegment) + " / Project segment: " + ResolveElementName((Element)(object)existingByCombination));
			}
		}
		if (targetByName != null)
		{
			return NonFamilyRoutingDependencyCheck.Ready("ResolveStalePipeSegmentName", "A same-name pipe segment exists with a non-standard material or schedule definition. Admin-authoritative apply will resolve or create the correct material/schedule counterpart, reconnect routing to it, and report the stale segment for cleanup. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + ruleIndex.ToString(CultureInfo.InvariantCulture) + " / Segment: " + ResolveElementName((Element)(object)sourceSegment));
		}
		return NonFamilyRoutingDependencyCheck.Ready("CreatePipeSegment", "Pipe segment can be created during the system type transaction. Dependency families have not been loaded yet. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Rule index: " + ruleIndex.ToString(CultureInfo.InvariantCulture) + " / Segment: " + ResolveElementName((Element)(object)sourceSegment));
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
		PipeSegment sourceSegment = (PipeSegment)(object)((sourcePart is PipeSegment) ? sourcePart : null);
		PipeSegment targetSegment = (PipeSegment)(object)((targetPart is PipeSegment) ? targetPart : null);
		if (sourceSegment != null || targetSegment != null)
		{
			if (sourceSegment == null || targetSegment == null)
			{
				return false;
			}
			ReferenceMappingCheck materialCheck = ResolvePotentialMappedReference(targetDocument, sourceDocument, ((Segment)sourceSegment).MaterialId, "pipe segment material");
			ReferenceMappingCheck scheduleCheck = ResolvePotentialMappedReference(targetDocument, sourceDocument, sourceSegment.ScheduleTypeId, "pipe segment schedule type");
			if (materialCheck.TargetReference == null || scheduleCheck.TargetReference == null)
			{
				return false;
			}
			return PipeSegmentDefinitionMatchesForRouting(sourceDocument, sourceSegment, targetDocument, targetSegment, materialCheck.TargetReference.Id, scheduleCheck.TargetReference.Id);
		}
		return RoutingPartSignatureService.Matches(sourceDocument, sourcePart, targetDocument, targetPart);
	}

	private static bool PipeSegmentDefinitionMatchesForRouting(Document sourceDocument, PipeSegment sourceSegment, Document targetDocument, PipeSegment targetSegment, ElementId expectedTargetMaterialId, ElementId expectedTargetScheduleId)
	{
		if (sourceDocument == null || targetDocument == null || sourceSegment == null || targetSegment == null)
		{
			return false;
		}
		if (expectedTargetMaterialId == null || expectedTargetMaterialId == ElementId.InvalidElementId || expectedTargetScheduleId == null || expectedTargetScheduleId == ElementId.InvalidElementId)
		{
			return false;
		}
		if (!ElementIdsEqual(((Segment)targetSegment).MaterialId, expectedTargetMaterialId) || !ElementIdsEqual(targetSegment.ScheduleTypeId, expectedTargetScheduleId))
		{
			return false;
		}
		if (!PipeSegmentContainsStandardSizes((Segment)(object)sourceSegment, (Segment)(object)targetSegment))
		{
			return false;
		}
		return PipeSegmentManagedRoutingValuesMatch(sourceSegment, targetSegment);
	}

	private static bool PipeSegmentContainsStandardSizes(Segment sourceSegment, Segment targetSegment)
	{
		bool PipeSegmentContainsStandardSizes;
		if (sourceSegment == null || targetSegment == null)
		{
			PipeSegmentContainsStandardSizes = false;
		}
		else
		{
			ICollection<MEPSize> sourceSizes = null;
			try
			{
				sourceSizes = sourceSegment.GetSizes();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				PipeSegmentContainsStandardSizes = false;
				ProjectData.ClearProjectError();
				goto IL_00ac;
			}
			PipeSegmentContainsStandardSizes = sourceSizes != null && sourceSizes.Where([SpecialName] (MEPSize x) => x != null).Count() != 0 && sourceSizes.Where([SpecialName] (MEPSize x) => x != null).All([SpecialName] (MEPSize sourceSize) => PipeSegmentContainsSizeDefinition(targetSegment, sourceSize));
		}
		goto IL_00ac;
		IL_00ac:
		return PipeSegmentContainsStandardSizes;
	}

	private static bool PipeSegmentSizeTablesMatch(Segment sourceSegment, Segment targetSegment)
	{
		if (!PipeSegmentContainsStandardSizes(sourceSegment, targetSegment))
		{
			return false;
		}
		return CountExtraPipeSegmentSizes(sourceSegment, targetSegment) == 0;
	}

	private static bool PipeSegmentContainsSizeDefinition(Segment segment, MEPSize sourceSize)
	{
		bool PipeSegmentContainsSizeDefinition;
		if (segment == null || sourceSize == null)
		{
			PipeSegmentContainsSizeDefinition = false;
		}
		else
		{
			try
			{
				PipeSegmentContainsSizeDefinition = (from x in segment.GetSizes()
					where x != null
					select x).Any([SpecialName] (MEPSize targetSize) => MEPSizeDefinitionsMatch(sourceSize, targetSize));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				PipeSegmentContainsSizeDefinition = false;
				ProjectData.ClearProjectError();
			}
		}
		return PipeSegmentContainsSizeDefinition;
	}

	private static int CountExtraPipeSegmentSizes(Segment sourceSegment, Segment targetSegment)
	{
		int CountExtraPipeSegmentSizes;
		if (sourceSegment == null || targetSegment == null)
		{
			CountExtraPipeSegmentSizes = 0;
		}
		else
		{
			try
			{
				CountExtraPipeSegmentSizes = (from x in targetSegment.GetSizes()
					where x != null
					select x).Count([SpecialName] (MEPSize targetSize) => !SegmentContainsNominalDiameter(sourceSegment, targetSize.NominalDiameter));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				CountExtraPipeSegmentSizes = 0;
				ProjectData.ClearProjectError();
			}
		}
		return CountExtraPipeSegmentSizes;
	}

	private static bool PipeSegmentManagedRoutingValuesMatch(PipeSegment sourceSegment, PipeSegment targetSegment)
	{
		return ScalarPropertyValuesMatch(sourceSegment, targetSegment, "Roughness");
	}

	private static bool ScalarPropertyValuesMatch(object source, object target, string propertyName)
	{
		bool ScalarPropertyValuesMatch;
		if (source == null || target == null || string.IsNullOrWhiteSpace(propertyName))
		{
			ScalarPropertyValuesMatch = true;
		}
		else
		{
			try
			{
				PropertyInfo sourceProperty = source.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
				PropertyInfo targetProperty = target.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
				if ((object)sourceProperty == null || (object)targetProperty == null || sourceProperty.GetIndexParameters().Length > 0 || targetProperty.GetIndexParameters().Length > 0)
				{
					ScalarPropertyValuesMatch = true;
				}
				else
				{
					object sourceValue = RuntimeHelpers.GetObjectValue(sourceProperty.GetValue(RuntimeHelpers.GetObjectValue(source), null));
					object targetValue = RuntimeHelpers.GetObjectValue(targetProperty.GetValue(RuntimeHelpers.GetObjectValue(target), null));
					ScalarPropertyValuesMatch = ((sourceValue != null && targetValue != null) ? ((!(sourceValue is double) || !(targetValue is double)) ? string.Equals(Convert.ToString(RuntimeHelpers.GetObjectValue(sourceValue), CultureInfo.InvariantCulture), Convert.ToString(RuntimeHelpers.GetObjectValue(targetValue), CultureInfo.InvariantCulture), StringComparison.Ordinal) : LengthValuesMatch(Conversions.ToDouble(sourceValue), Conversions.ToDouble(targetValue))) : (sourceValue == null && targetValue == null));
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ScalarPropertyValuesMatch = true;
				ProjectData.ClearProjectError();
			}
		}
		return ScalarPropertyValuesMatch;
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

	private static string BuildMEPSizeNominalDiameterKey(MEPSize size)
	{
		if (size == null)
		{
			return string.Empty;
		}
		return FormatLength(size.NominalDiameter);
	}

	private static bool SegmentContainsNominalDiameter(Segment segment, double nominalDiameter)
	{
		bool SegmentContainsNominalDiameter;
		if (segment == null)
		{
			SegmentContainsNominalDiameter = false;
		}
		else
		{
			try
			{
				SegmentContainsNominalDiameter = (from x in segment.GetSizes()
					where x != null
					select x).Any([SpecialName] (MEPSize size) => LengthValuesMatch(size.NominalDiameter, nominalDiameter));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				SegmentContainsNominalDiameter = false;
				ProjectData.ClearProjectError();
			}
		}
		return SegmentContainsNominalDiameter;
	}

	private static bool LengthValuesMatch(double left, double right)
	{
		double scale = Math.Max(1.0, Math.Max(Math.Abs(left), Math.Abs(right)));
		return Math.Abs(left - right) <= 1E-09 * scale;
	}

	private static string FormatLength(double value)
	{
		return value.ToString("G17", CultureInfo.InvariantCulture);
	}

	private static bool ElementIdsEqual(ElementId left, ElementId right)
	{
		if (left == null || right == null)
		{
			return left == null && right == null;
		}
		return RevitElementIdCompat.CompatIntegerValue(left) == RevitElementIdCompat.CompatIntegerValue(right);
	}

	private static ElementId MapReferenceElementId(Document targetDocument, Document standardDocument, ElementId sourceId)
	{
		//IL_004a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0056: Expected O, but got Unknown
		if (sourceId == null || sourceId == ElementId.InvalidElementId)
		{
			return ElementId.InvalidElementId;
		}
		if (RevitElementIdCompat.CompatIntegerValue(sourceId) < 0)
		{
			return sourceId;
		}
		Element sourceReference = standardDocument.GetElement(sourceId);
		if (sourceReference == null)
		{
			return ElementId.InvalidElementId;
		}
		if (sourceReference is Material)
		{
			return EnsureMaterialAuthoritative((Material)sourceReference, standardDocument, targetDocument);
		}
		PipeSegment sourcePipeSegment = (PipeSegment)(object)((sourceReference is PipeSegment) ? sourceReference : null);
		if (sourcePipeSegment != null)
		{
			PipeSegment targetPipeSegment = EnsurePipeSegmentAuthoritative(targetDocument, standardDocument, sourcePipeSegment, SystemTypeApplyAuthorityMode.AdminAuthoritative, null);
			return (targetPipeSegment == null) ? ElementId.InvalidElementId : ((Element)targetPipeSegment).Id;
		}
		PipeScheduleType sourceSchedule = (PipeScheduleType)(object)((sourceReference is PipeScheduleType) ? sourceReference : null);
		if (sourceSchedule != null)
		{
			return EnsurePipeScheduleType(sourceSchedule, standardDocument, targetDocument);
		}
		Element targetReference = ResolveMatchingTargetElement(targetDocument, sourceReference);
		if (targetReference != null)
		{
			EnsureMappedElementIdReferenceMatches(targetDocument, standardDocument, sourceReference, targetReference, null);
			return targetReference.Id;
		}
		if (targetReference == null && !(sourceReference is FamilySymbol))
		{
			targetReference = CopyNonFamilyElementToTarget(targetDocument, standardDocument, sourceReference);
		}
		if (targetReference != null)
		{
			EnsureMappedElementIdReferenceMatches(targetDocument, standardDocument, sourceReference, targetReference, null);
		}
		return (targetReference == null) ? ElementId.InvalidElementId : targetReference.Id;
	}

	private static ElementId EnsureMaterialAuthoritative(Material sourceMaterial, Document sourceDocument, Document targetDocument)
	{
		return EnsureMaterialForPipeSegment(sourceMaterial, sourceDocument, targetDocument, null).MaterialId;
	}

	private static MaterialEnsureResult EnsureMaterialForPipeSegment(Material sourceMaterial, Document sourceDocument, Document targetDocument, SystemTypeApplyExecutionItem resultItem)
	{
		MaterialEnsureResult result = new MaterialEnsureResult();
		if (sourceMaterial == null)
		{
			result.MaterialId = ElementId.InvalidElementId;
			return result;
		}
		string materialName = ResolveElementName((Element)(object)sourceMaterial);
		if (string.IsNullOrWhiteSpace(materialName))
		{
			throw new InvalidOperationException(T("A standard pipe segment material has no readable name.", "표준 배관 세그먼트 재료 이름을 읽을 수 없습니다."));
		}
		Material targetMaterial = FindMaterialByName(targetDocument, materialName);
		if (targetMaterial == null)
		{
			targetMaterial = TryCopyMaterialToTarget(sourceDocument, sourceMaterial, targetDocument, resultItem);
		}
		if (targetMaterial == null)
		{
			ElementId createdId = Material.Create(targetDocument, materialName);
			Element element = targetDocument.GetElement(createdId);
			targetMaterial = (Material)(object)((element is Material) ? element : null);
			if (targetMaterial == null)
			{
				throw new InvalidOperationException(T("Revit did not create the standard material: ", "Revit이 표준 재료를 만들지 못했습니다: ") + materialName);
			}
		}
		SyncBasicMaterialProperties(sourceDocument, targetDocument, sourceMaterial, targetMaterial);
		CopyWritableMaterialParametersBestEffort(targetDocument, sourceDocument, sourceMaterial, targetMaterial, resultItem);
		targetDocument.Regenerate();
		result.MaterialId = ((Element)targetMaterial).Id;
		result.ExactSignatureMatch = MaterialDefinitionsMatch(sourceDocument, sourceMaterial, targetDocument, targetMaterial);
		if (!result.ExactSignatureMatch)
		{
			result.Warning = "Material exact signature differs after copy/sync; Pipe Type apply will continue with the resolved same-name target material and report it for admin review. Material: " + materialName;
			AddResultMessage(resultItem, result.Warning);
		}
		return result;
	}

	private static Material TryCopyMaterialToTarget(Document sourceDocument, Material sourceMaterial, Document targetDocument, SystemTypeApplyExecutionItem resultItem)
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Expected O, but got Unknown
		if (sourceDocument == null || sourceMaterial == null || targetDocument == null)
		{
			return null;
		}
		string materialName = ResolveElementName((Element)(object)sourceMaterial);
		try
		{
			CopyPasteOptions options = new CopyPasteOptions();
			try
			{
				options.SetDuplicateTypeNamesHandler((IDuplicateTypeNamesHandler)(object)new CopyPasteUseDestinationTypesHandler());
				ElementTransformUtils.CopyElements(sourceDocument, (ICollection<ElementId>)new List<ElementId> { ((Element)sourceMaterial).Id }, targetDocument, Transform.Identity, options);
			}
			finally
			{
				((IDisposable)options)?.Dispose();
			}
			targetDocument.Regenerate();
			Material copied = FindMaterialByName(targetDocument, materialName);
			if (copied != null)
			{
				return copied;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			AddResultMessage(resultItem, T("Material exact copy failed, so same-name create/reuse will be used. Material: ", "재료 정확 복사에 실패하여 같은 이름 생성/재사용으로 진행합니다. 재료: ") + materialName + T(" / Reason: ", " / 사유: ") + ex2.Message);
			ProjectData.ClearProjectError();
		}
		return null;
	}

	private static bool MaterialDefinitionsMatch(Document sourceDocument, Material sourceMaterial, Document targetDocument, Material targetMaterial)
	{
		bool MaterialDefinitionsMatch;
		try
		{
			MaterialDefinitionsMatch = RoutingPartSignatureService.Matches(sourceDocument, (Element)(object)sourceMaterial, targetDocument, (Element)(object)targetMaterial);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			MaterialDefinitionsMatch = false;
			ProjectData.ClearProjectError();
		}
		return MaterialDefinitionsMatch;
	}

	private static void CopyWritableMaterialParametersBestEffort(Document targetDocument, Document standardDocument, Material sourceMaterial, Material targetMaterial, SystemTypeApplyExecutionItem resultItem)
	{
		if (sourceMaterial == null || targetMaterial == null)
		{
			return;
		}
		foreach (Parameter sourceParameter in ((IEnumerable)((Element)sourceMaterial).Parameters).Cast<Parameter>())
		{
			if (sourceParameter == null || sourceParameter.Definition == null || !sourceParameter.HasValue || ShouldSkipTypeParameter(sourceParameter))
			{
				continue;
			}
			Parameter targetParameter = ResolveWritableElementParameter((Element)(object)targetMaterial, sourceParameter);
			if (targetParameter != null)
			{
				try
				{
					TrySetParameterValue(targetDocument, standardDocument, targetParameter, sourceParameter);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					AddResultMessage(resultItem, "Material parameter sync warning: a writable material parameter could not be copied, so Pipe Type apply will continue with the resolved material. Material: " + ResolveElementName((Element)(object)sourceMaterial) + " / Parameter: " + ResolveParameterName(sourceParameter) + " / Reason: " + ResolveExceptionMessage(ex2));
					ProjectData.ClearProjectError();
				}
			}
		}
	}

	private static ElementId EnsurePipeScheduleType(PipeScheduleType sourceSchedule, Document sourceDocument, Document targetDocument)
	{
		return EnsurePipeScheduleTypeForPipeSegment(sourceSchedule, sourceDocument, targetDocument, null);
	}

	private static ElementId EnsurePipeScheduleTypeForPipeSegment(PipeScheduleType sourceSchedule, Document sourceDocument, Document targetDocument, SystemTypeApplyExecutionItem resultItem)
	{
		if (sourceSchedule == null)
		{
			return ElementId.InvalidElementId;
		}
		string scheduleName = ResolveElementName((Element)(object)sourceSchedule);
		if (string.IsNullOrWhiteSpace(scheduleName))
		{
			throw new InvalidOperationException(T("A standard pipe segment schedule type has no readable name.", "표준 배관 세그먼트 스케줄 타입 이름을 읽을 수 없습니다."));
		}
		PipeScheduleType targetSchedule = FindPipeScheduleTypeByName(targetDocument, scheduleName);
		if (targetSchedule != null)
		{
			return ((Element)targetSchedule).Id;
		}
		ElementId createdId = CreatePipeScheduleType(targetDocument, scheduleName);
		if (createdId == null || createdId == ElementId.InvalidElementId)
		{
			throw new InvalidOperationException(T("Revit did not create the standard pipe schedule type: ", "Revit이 표준 배관 스케줄 타입을 만들지 못했습니다: ") + scheduleName);
		}
		Element element = targetDocument.GetElement(createdId);
		Element obj = ((element is PipeScheduleType) ? element : null) ?? throw new InvalidOperationException(T("The created pipe schedule type could not be resolved: ", "생성된 배관 스케줄 타입을 확인하지 못했습니다: ") + scheduleName);
		targetDocument.Regenerate();
		if (!string.Equals(Normalize(ResolveElementName(obj)), Normalize(scheduleName), StringComparison.Ordinal))
		{
			throw new InvalidOperationException("The standard pipe schedule type was created, but Revit returned a different name. The Pipe Type apply was rolled back to avoid a partial standard system. Schedule: " + scheduleName);
		}
		AddResultMessage(resultItem, T("PipeScheduleType created from standard: ", "표준에서 배관 스케줄 타입을 생성했습니다: ") + scheduleName);
		return obj.Id;
	}

	private static Material FindMaterialByName(Document targetDocument, string materialName)
	{
		//IL_0022: Unknown result type (might be due to invalid IL or missing references)
		if (targetDocument == null || string.IsNullOrWhiteSpace(materialName))
		{
			return null;
		}
		return ((IEnumerable)new FilteredElementCollector(targetDocument).OfClass(typeof(Material))).Cast<Material>().FirstOrDefault([SpecialName] (Material x) => x != null && string.Equals(Normalize(ResolveElementName((Element)(object)x)), Normalize(materialName), StringComparison.Ordinal));
	}

	private static void SyncBasicMaterialProperties(Document sourceDocument, Document targetDocument, Material sourceMaterial, Material targetMaterial)
	{
		if (sourceMaterial != null && targetMaterial != null)
		{
			TryCopyWritableMaterialProperty(sourceDocument, targetDocument, sourceMaterial, targetMaterial, "Color");
			TryCopyWritableMaterialProperty(sourceDocument, targetDocument, sourceMaterial, targetMaterial, "Transparency");
			TryCopyWritableMaterialProperty(sourceDocument, targetDocument, sourceMaterial, targetMaterial, "Smoothness");
			TryCopyWritableMaterialProperty(sourceDocument, targetDocument, sourceMaterial, targetMaterial, "Shininess");
			TryCopyWritableMaterialProperty(sourceDocument, targetDocument, sourceMaterial, targetMaterial, "SurfaceForegroundPatternId");
			TryCopyWritableMaterialProperty(sourceDocument, targetDocument, sourceMaterial, targetMaterial, "SurfaceBackgroundPatternId");
			TryCopyWritableMaterialProperty(sourceDocument, targetDocument, sourceMaterial, targetMaterial, "CutForegroundPatternId");
			TryCopyWritableMaterialProperty(sourceDocument, targetDocument, sourceMaterial, targetMaterial, "CutBackgroundPatternId");
			TryCopyWritableMaterialProperty(sourceDocument, targetDocument, sourceMaterial, targetMaterial, "SurfaceForegroundPatternColor");
			TryCopyWritableMaterialProperty(sourceDocument, targetDocument, sourceMaterial, targetMaterial, "SurfaceBackgroundPatternColor");
			TryCopyWritableMaterialProperty(sourceDocument, targetDocument, sourceMaterial, targetMaterial, "CutForegroundPatternColor");
			TryCopyWritableMaterialProperty(sourceDocument, targetDocument, sourceMaterial, targetMaterial, "CutBackgroundPatternColor");
		}
	}

	private static void TryCopyWritableMaterialProperty(Document sourceDocument, Document targetDocument, object source, object target, string propertyName)
	{
		//IL_0067: Unknown result type (might be due to invalid IL or missing references)
		//IL_0071: Expected O, but got Unknown
		if (source == null || target == null || string.IsNullOrWhiteSpace(propertyName))
		{
			return;
		}
		try
		{
			PropertyInfo sourceProperty = source.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
			PropertyInfo targetProperty = target.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
			if ((object)sourceProperty == null || (object)targetProperty == null || !targetProperty.CanWrite)
			{
				return;
			}
			object sourceValue = RuntimeHelpers.GetObjectValue(sourceProperty.GetValue(RuntimeHelpers.GetObjectValue(source), null));
			if (sourceValue is ElementId)
			{
				ElementId mappedId = MapReferenceElementId(targetDocument, sourceDocument, (ElementId)sourceValue);
				if (mappedId != null && !(mappedId == ElementId.InvalidElementId))
				{
					targetProperty.SetValue(RuntimeHelpers.GetObjectValue(target), mappedId, null);
				}
			}
			else
			{
				targetProperty.SetValue(RuntimeHelpers.GetObjectValue(target), RuntimeHelpers.GetObjectValue(sourceValue), null);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void TryCopyWritableProperty(object source, object target, string propertyName)
	{
		if (source == null || target == null || string.IsNullOrWhiteSpace(propertyName))
		{
			return;
		}
		try
		{
			PropertyInfo sourceProperty = source.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
			PropertyInfo targetProperty = target.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
			if ((object)sourceProperty != null && (object)targetProperty != null && targetProperty.CanWrite)
			{
				targetProperty.SetValue(RuntimeHelpers.GetObjectValue(target), RuntimeHelpers.GetObjectValue(sourceProperty.GetValue(RuntimeHelpers.GetObjectValue(source), null)), null);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static PipeScheduleType FindPipeScheduleTypeByName(Document targetDocument, string scheduleName)
	{
		//IL_0053: Unknown result type (might be due to invalid IL or missing references)
		if (targetDocument == null || string.IsNullOrWhiteSpace(scheduleName))
		{
			return null;
		}
		ElementId scheduleId = TryGetPipeScheduleTypeId(targetDocument, scheduleName);
		if (scheduleId != null && scheduleId != ElementId.InvalidElementId)
		{
			Element element = targetDocument.GetElement(scheduleId);
			PipeScheduleType schedule = (PipeScheduleType)(object)((element is PipeScheduleType) ? element : null);
			if (schedule != null)
			{
				return schedule;
			}
		}
		return (from Element x in (IEnumerable)new FilteredElementCollector(targetDocument).WhereElementIsElementType()
			select (PipeScheduleType)(object)((x is PipeScheduleType) ? x : null)).FirstOrDefault([SpecialName] (PipeScheduleType x) => x != null && string.Equals(Normalize(ResolveElementName((Element)(object)x)), Normalize(scheduleName), StringComparison.Ordinal));
	}

	private static ElementId TryGetPipeScheduleTypeId(Document targetDocument, string scheduleName)
	{
		ElementId TryGetPipeScheduleTypeId;
		try
		{
			MethodInfo method = typeof(PipeScheduleType).GetMethods(BindingFlags.Static | BindingFlags.Public).FirstOrDefault([SpecialName] (MethodInfo x) => string.Equals(x.Name, "GetPipeScheduleId", StringComparison.Ordinal) && x.GetParameters().Length == 2);
			if ((object)method == null)
			{
				TryGetPipeScheduleTypeId = ElementId.InvalidElementId;
			}
			else
			{
				object obj = method.Invoke(null, new object[2] { targetDocument, scheduleName });
				TryGetPipeScheduleTypeId = (ElementId)((obj is ElementId) ? obj : null);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			TryGetPipeScheduleTypeId = ElementId.InvalidElementId;
			ProjectData.ClearProjectError();
		}
		return TryGetPipeScheduleTypeId;
	}

	private static ElementId CreatePipeScheduleType(Document targetDocument, string scheduleName)
	{
		try
		{
			MethodInfo method = typeof(PipeScheduleType).GetMethods(BindingFlags.Static | BindingFlags.Public).FirstOrDefault([SpecialName] (MethodInfo x) => string.Equals(x.Name, "Create", StringComparison.Ordinal) && x.GetParameters().Length == 2);
			if ((object)method == null)
			{
				return ElementId.InvalidElementId;
			}
			object created = RuntimeHelpers.GetObjectValue(method.Invoke(null, new object[2] { targetDocument, scheduleName }));
			ElementId createdId = (ElementId)((created is ElementId) ? created : null);
			if (createdId != null)
			{
				return createdId;
			}
			Element createdElement = (Element)((created is Element) ? created : null);
			if (createdElement != null)
			{
				return createdElement.Id;
			}
		}
		catch (TargetInvocationException ex)
		{
			ProjectData.SetProjectError(ex);
			TargetInvocationException ex2 = ex;
			throw new InvalidOperationException(T("The standard pipe schedule type could not be created: ", "표준 배관 스케줄 타입을 만들지 못했습니다: ") + scheduleName, ex2.InnerException);
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			throw new InvalidOperationException(T("The standard pipe schedule type could not be created: ", "표준 배관 스케줄 타입을 만들지 못했습니다: ") + scheduleName, ex4);
		}
		return ElementId.InvalidElementId;
	}

	private static string BuildCopiedElementSummary(IEnumerable<Element> copiedElements)
	{
		List<string> values = (from x in copiedElements ?? new List<Element>()
			where x != null
			select ((object)x).GetType().Name + " : " + ResolveElementName(x)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
		if (values.Count == 0)
		{
			return "(none)";
		}
		return string.Join(", ", values);
	}

	private static void CopyWritableElementParameters(Document targetDocument, Document standardDocument, Element sourceElement, Element targetElement)
	{
		if (sourceElement == null || targetElement == null)
		{
			return;
		}
		foreach (Parameter sourceParameter in ((IEnumerable)sourceElement.Parameters).Cast<Parameter>())
		{
			if (sourceParameter != null && sourceParameter.Definition != null && sourceParameter.HasValue && !ShouldSkipTypeParameter(sourceParameter))
			{
				Parameter targetParameter = ResolveWritableElementParameter(targetElement, sourceParameter);
				if (targetParameter != null)
				{
					TrySetParameterValue(targetDocument, standardDocument, targetParameter, sourceParameter);
				}
			}
		}
	}

	private static Parameter ResolveWritableElementParameter(Element targetElement, Parameter sourceParameter)
	{
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
		//IL_0046: Unknown result type (might be due to invalid IL or missing references)
		Definition definition = sourceParameter.Definition;
		string name = ((definition != null) ? definition.Name : null) ?? string.Empty;
		if (string.IsNullOrWhiteSpace(name))
		{
			return null;
		}
		Parameter targetParameter = targetElement.LookupParameter(name);
		if (targetParameter == null || targetParameter.IsReadOnly)
		{
			return null;
		}
		if (targetParameter.StorageType != sourceParameter.StorageType)
		{
			return null;
		}
		return targetParameter;
	}

	private unsafe static void CopyRoutingCriteria(RoutingPreferenceRule sourceRule, RoutingPreferenceRule targetRule, RoutingPreferenceRuleGroupType group)
	{
		int criterionCount;
		try
		{
			criterionCount = sourceRule.NumberOfCriteria;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			throw new InvalidOperationException("Routing preference criteria count could not be read. The system type apply was stopped before rebuilding routing preferences. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString(), ex2);
		}
		int num = checked(criterionCount - 1);
		for (int index = 0; index <= num; index = checked(index + 1))
		{
			try
			{
				RoutingCriterionBase criterion = sourceRule.GetCriterion(index);
				if (criterion == null)
				{
					throw new InvalidOperationException(T("The routing preference criterion returned no data.", "라우팅 환경설정 조건이 데이터를 반환하지 않았습니다."));
				}
				targetRule.AddCriterion(criterion);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				throw new InvalidOperationException("Routing preference criteria could not be copied. The system type apply was stopped before creating a partial routing rule. Rule group: " + ((Enum)(*(RoutingPreferenceRuleGroupType*)(&group))/*cast due to .constrained prefix*/).ToString() + " / Criterion index: " + index.ToString(CultureInfo.InvariantCulture), ex4);
			}
		}
	}

	private static void EnsureDependencyFamilyCanOverwrite(Document targetDocument, Family standardFamily, RoutingDependencyPreflightItem dependencyItem)
	{
		if (targetDocument == null || standardFamily == null || dependencyItem == null)
		{
			return;
		}
		string familyName = (dependencyItem.SourceFamilyName ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(familyName))
		{
			return;
		}
		string standardCategoryRaw = ResolveFamilyCategoryName(standardFamily);
		string text = Normalize(standardCategoryRaw);
		if (string.IsNullOrWhiteSpace(text))
		{
			return;
		}
		List<VB_0024AnonymousType_1<Family, string>> categoryMismatches = (from f in FindFamiliesByExactName(targetDocument, familyName)
			select new VB_0024AnonymousType_1<Family, string>(f, ResolveFamilyCategoryName(f)) into x
			where !string.IsNullOrWhiteSpace(Normalize(x.CategoryName)) && !string.Equals(Normalize(x.CategoryName), text, StringComparison.Ordinal)
			select x).ToList();
		if (categoryMismatches.Count != 0)
		{
			List<string> projectCategories = (from x in categoryMismatches
				select x.CategoryName into x
				where !string.IsNullOrWhiteSpace(x)
				select x).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
			throw new InvalidOperationException("A dependency family with the same name exists in the current project, but its category differs from the standard RVT. Review this family before applying the system type. Family: " + familyName + " / Standard category: " + standardCategoryRaw + " / Project category: " + ((projectCategories.Count == 0) ? string.Empty : string.Join(", ", projectCategories)));
		}
	}

	private static IEnumerable<Family> FindFamiliesByExactName(Document targetDocument, string familyName)
	{
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		string b = Normalize(familyName);
		return from Family x in (IEnumerable)new FilteredElementCollector(targetDocument).OfClass(typeof(Family))
			where x != null && string.Equals(Normalize(((Element)x).Name), b, StringComparison.Ordinal)
			select x;
	}

	private static ISet<int> CaptureFamilyNameState(Document targetDocument)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		return new HashSet<int>(from Family x in (IEnumerable)new FilteredElementCollector(targetDocument).OfClass(typeof(Family))
			where x != null && ((Element)x).Id != null
			select RevitElementIdCompat.CompatIntegerValue(((Element)x).Id));
	}

	private static List<AllowedLoadedFamilyIdentity> BuildAllowedLoadedFamilyIdentities(Family sourceFamily, Document familyDocument, Document standardDocument, IEnumerable<RoutingDependencyPreflightItem> selectedDependencyItems, IEnumerable<string> explicitFamilyNames)
	{
		//IL_0178: Unknown result type (might be due to invalid IL or missing references)
		//IL_021b: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e8: Unknown result type (might be due to invalid IL or missing references)
		List<AllowedLoadedFamilyIdentity> result = new List<AllowedLoadedFamilyIdentity>();
		AddAllowedLoadedFamilyIdentity(result, sourceFamily);
		List<string> requestedNames = new List<string>();
		if (sourceFamily != null)
		{
			requestedNames.Add(((Element)sourceFamily).Name);
		}
		requestedNames.AddRange(explicitFamilyNames ?? new List<string>());
		requestedNames.AddRange(from x in selectedDependencyItems ?? new List<RoutingDependencyPreflightItem>()
			where x != null
			select x.SourceFamilyName);
		foreach (string familyName in UniqueSortedNames(requestedNames))
		{
			AddAllowedLoadedFamilyIdentityByName(result, standardDocument, familyName);
			AddAllowedLoadedFamilyName(result, familyName);
		}
		if (familyDocument == null)
		{
			return result;
		}
		List<string> nestedFamilyNames = new List<string>();
		try
		{
			foreach (Family nestedFamily in from Family x in (IEnumerable)new FilteredElementCollector(familyDocument).OfClass(typeof(Family))
				where x != null
				select x)
			{
				AddAllowedLoadedFamilyIdentity(result, nestedFamily);
				nestedFamilyNames.Add(((Element)nestedFamily).Name);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			foreach (FamilySymbol symbol in from FamilySymbol x in (IEnumerable)new FilteredElementCollector(familyDocument).OfClass(typeof(FamilySymbol))
				where x != null
				select x)
			{
				AddAllowedLoadedFamilyIdentity(result, symbol.Family);
				if (symbol.Family != null)
				{
					nestedFamilyNames.Add(((Element)symbol.Family).Name);
				}
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		try
		{
			foreach (FamilyInstance item in from FamilyInstance x in (IEnumerable)new FilteredElementCollector(familyDocument).OfClass(typeof(FamilyInstance))
				where x != null
				select x)
			{
				FamilySymbol symbol2 = item.Symbol;
				if (symbol2 != null)
				{
					AddAllowedLoadedFamilyIdentity(result, symbol2.Family);
					if (symbol2.Family != null)
					{
						nestedFamilyNames.Add(((Element)symbol2.Family).Name);
					}
				}
			}
		}
		catch (Exception projectError3)
		{
			ProjectData.SetProjectError(projectError3);
			ProjectData.ClearProjectError();
		}
		foreach (string familyName2 in UniqueSortedNames(nestedFamilyNames))
		{
			AddAllowedLoadedFamilyIdentityByName(result, standardDocument, familyName2);
		}
		return result;
	}

	private static void AddAllowedLoadedFamilyIdentityByName(IList<AllowedLoadedFamilyIdentity> identities, Document standardDocument, string familyName)
	{
		//IL_002b: Unknown result type (might be due to invalid IL or missing references)
		_Closure_0024__112_002D0 arg = default(_Closure_0024__112_002D0);
		_Closure_0024__112_002D0 CS_0024_003C_003E8__locals2 = new _Closure_0024__112_002D0(arg);
		if (identities == null || standardDocument == null || string.IsNullOrWhiteSpace(familyName))
		{
			return;
		}
		CS_0024_003C_003E8__locals2._0024VB_0024Local_normalizedName = Normalize(familyName);
		try
		{
			foreach (Family standardFamily in from Family x in (IEnumerable)new FilteredElementCollector(standardDocument).OfClass(typeof(Family))
				where x != null && string.Equals(Normalize(((Element)x).Name), CS_0024_003C_003E8__locals2._0024VB_0024Local_normalizedName, StringComparison.Ordinal)
				select x)
			{
				AddAllowedLoadedFamilyIdentity(identities, standardFamily);
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void AddAllowedLoadedFamilyName(IList<AllowedLoadedFamilyIdentity> identities, string familyName)
	{
		if (identities != null && !string.IsNullOrWhiteSpace(familyName))
		{
			string b = Normalize(familyName);
			if (!identities.Any([SpecialName] (AllowedLoadedFamilyIdentity x) => x != null && string.Equals(Normalize(x.Name), b, StringComparison.Ordinal) && string.IsNullOrWhiteSpace(Normalize(x.CategoryName))))
			{
				identities.Add(new AllowedLoadedFamilyIdentity
				{
					Name = familyName.Trim(),
					CategoryName = string.Empty
				});
			}
		}
	}

	private static void AddAllowedLoadedFamilyIdentity(IList<AllowedLoadedFamilyIdentity> identities, Family family)
	{
		if (identities == null || family == null)
		{
			return;
		}
		string familyName = (((Element)family).Name ?? string.Empty).Trim();
		if (!string.IsNullOrWhiteSpace(familyName))
		{
			string categoryName = ResolveFamilyCategoryName(family);
			string b = Normalize(familyName);
			string b2 = Normalize(categoryName);
			if (!identities.Any([SpecialName] (AllowedLoadedFamilyIdentity x) => x != null && string.Equals(Normalize(x.Name), b, StringComparison.Ordinal) && string.Equals(Normalize(x.CategoryName), b2, StringComparison.Ordinal)))
			{
				identities.Add(new AllowedLoadedFamilyIdentity
				{
					Name = familyName,
					CategoryName = categoryName
				});
			}
		}
	}

	private static bool LoadedFamilyIsAllowed(Family family, IEnumerable<AllowedLoadedFamilyIdentity> allowedFamilies)
	{
		if (family == null)
		{
			return false;
		}
		string familyName = ((Element)family).Name ?? string.Empty;
		string text = Normalize(familyName);
		if (string.IsNullOrWhiteSpace(text))
		{
			return false;
		}
		if (HasNumericDuplicateSuffixOfAllowedFamily(familyName, allowedFamilies))
		{
			return false;
		}
		string left = ResolveFamilyCategoryName(family);
		List<AllowedLoadedFamilyIdentity> matchingAllowedFamilies = (allowedFamilies ?? new List<AllowedLoadedFamilyIdentity>()).Where([SpecialName] (AllowedLoadedFamilyIdentity x) => x != null && string.Equals(text, Normalize(x.Name), StringComparison.Ordinal)).ToList();
		if (matchingAllowedFamilies.Count == 0)
		{
			return false;
		}
		List<AllowedLoadedFamilyIdentity> knownCategoryMatches = matchingAllowedFamilies.Where([SpecialName] (AllowedLoadedFamilyIdentity x) => !string.IsNullOrWhiteSpace(Normalize(x.CategoryName))).ToList();
		if (knownCategoryMatches.Count == 0)
		{
			return true;
		}
		return knownCategoryMatches.Any([SpecialName] (AllowedLoadedFamilyIdentity x) => CategoryNamesMatch(left, x.CategoryName));
	}

	private static bool HasNumericDuplicateSuffixOfAllowedFamily(string familyName, IEnumerable<AllowedLoadedFamilyIdentity> allowedFamilies)
	{
		string normalizedName = Normalize(familyName);
		string b = Normalize(RemoveDuplicateSuffix(familyName));
		if (string.IsNullOrWhiteSpace(normalizedName) || string.Equals(normalizedName, b, StringComparison.Ordinal))
		{
			return false;
		}
		return (allowedFamilies ?? new List<AllowedLoadedFamilyIdentity>()).Any([SpecialName] (AllowedLoadedFamilyIdentity x) => x != null && string.Equals(Normalize(x.Name), b, StringComparison.Ordinal));
	}

	private static bool LoadedFamilyCategoriesMatch(string left, string right)
	{
		string normalizedLeft = Normalize(left);
		string normalizedRight = Normalize(right);
		if (string.IsNullOrWhiteSpace(normalizedLeft) || string.IsNullOrWhiteSpace(normalizedRight))
		{
			return true;
		}
		return string.Equals(normalizedLeft, normalizedRight, StringComparison.Ordinal);
	}

	private static void GuardAgainstCopiedFamilies(Document targetDocument, ISet<int> familyStateBefore, IEnumerable<RoutingDependencyPreflightItem> dependencyItems = null)
	{
		//IL_007a: Unknown result type (might be due to invalid IL or missing references)
		HashSet<string> hashSet = new HashSet<string>(from x in dependencyItems ?? new List<RoutingDependencyPreflightItem>()
			select Normalize(x.SourceFamilyName) into x
			where !string.IsNullOrWhiteSpace(x)
			select x, StringComparer.Ordinal);
		List<Family> newFamilies = (from Family x in (IEnumerable)new FilteredElementCollector(targetDocument).OfClass(typeof(Family))
			where x != null && ((Element)x).Id != null && !familyStateBefore.Contains(RevitElementIdCompat.CompatIntegerValue(((Element)x).Id))
			select x).ToList();
		if (newFamilies.Count != 0)
		{
			List<string> unexpected = (from x in newFamilies.Where([SpecialName] (Family x) =>
				{
					string value = ((Element)x).Name ?? string.Empty;
					string text = Normalize(value);
					string b = Normalize(RemoveDuplicateSuffix(value));
					return (!hashSet.Contains(text) || !string.Equals(text, b, StringComparison.Ordinal)) ? true : false;
				})
				select ((Element)x).Name ?? string.Empty).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
			if (unexpected.Count > 0)
			{
				throw new InvalidOperationException("System type copy tried to create unexpected or duplicate-suffixed loadable families. The transaction was rolled back. Families: " + string.Join(", ", unexpected));
			}
		}
	}

	private static void GuardDependencyLoadDidNotCreateDuplicateFamilies(Document targetDocument, ISet<int> familyStateBefore, Family standardFamily, RoutingDependencyPreflightItem dependencyItem, Family loadedFamily, IEnumerable<AllowedLoadedFamilyIdentity> allowedFamilies, SystemTypeApplyExecutionItem resultItem)
	{
		//IL_0063: Unknown result type (might be due to invalid IL or missing references)
		_Closure_0024__119_002D0 arg = default(_Closure_0024__119_002D0);
		_Closure_0024__119_002D0 CS_0024_003C_003E8__locals8 = new _Closure_0024__119_002D0(arg);
		CS_0024_003C_003E8__locals8._0024VB_0024Local_familyStateBefore = familyStateBefore;
		CS_0024_003C_003E8__locals8._0024VB_0024Local_expectedName = (dependencyItem.SourceFamilyName ?? string.Empty).Trim();
		List<string> allowedFamilyNames = UniqueSortedNames((allowedFamilies ?? new List<AllowedLoadedFamilyIdentity>()).Select([SpecialName] (AllowedLoadedFamilyIdentity x) => x.Name ?? string.Empty));
		List<Family> newFamilies = (from Family x in (IEnumerable)new FilteredElementCollector(targetDocument).OfClass(typeof(Family))
			where x != null && ((Element)x).Id != null && !CS_0024_003C_003E8__locals8._0024VB_0024Local_familyStateBefore.Contains(RevitElementIdCompat.CompatIntegerValue(((Element)x).Id))
			select x).ToList();
		if (newFamilies.Count == 0)
		{
			AddFamilyLoadDiagnostic(resultItem, "DependencyFamilyLoad.Guard", CS_0024_003C_003E8__locals8._0024VB_0024Local_expectedName, ((loadedFamily != null) ? ((Element)loadedFamily).Name : null) ?? string.Empty, allowedFamilyNames, new List<string>(), new List<string>(), new List<string>(), "None");
			return;
		}
		List<string> newFamilyNames = UniqueSortedNames(newFamilies.Select([SpecialName] (Family x) => ((Element)x).Name ?? string.Empty));
		List<Family> unexpected = new List<Family>();
		List<string> allowedNewNames = new List<string>();
		foreach (Family family in newFamilies)
		{
			string familyName = ((Element)family).Name ?? string.Empty;
			if (LoadedFamilyIsAllowed(family, allowedFamilies))
			{
				allowedNewNames.Add(familyName);
			}
			else
			{
				unexpected.Add(family);
			}
		}
		if (unexpected.Count == 0)
		{
			AddFamilyLoadDiagnostic(resultItem, "DependencyFamilyLoad.Guard", CS_0024_003C_003E8__locals8._0024VB_0024Local_expectedName, ((loadedFamily != null) ? ((Element)loadedFamily).Name : null) ?? string.Empty, allowedFamilyNames, newFamilyNames, allowedNewNames, new List<string>(), "None");
			List<string> allowedNestedNames = allowedNewNames.Where([SpecialName] (string x) => !string.Equals(Normalize(x), Normalize(CS_0024_003C_003E8__locals8._0024VB_0024Local_expectedName), StringComparison.Ordinal)).ToList();
			if (allowedNestedNames.Count > 0)
			{
				AddResultMessage(resultItem, "Allowed standard nested dependency families loaded: " + string.Join(", ", UniqueSortedNames(allowedNestedNames)));
			}
			return;
		}
		List<string> unexpectedNames = unexpected.Select([SpecialName] (Family x) => ((Element)x).Name ?? string.Empty).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase)
			.ToList();
		DuplicateCleanupResult cleanup = TryDeleteFamilies(targetDocument, unexpected);
		AddFamilyLoadDiagnostic(resultItem, "DependencyFamilyLoad.Guard", CS_0024_003C_003E8__locals8._0024VB_0024Local_expectedName, ((loadedFamily != null) ? ((Element)loadedFamily).Name : null) ?? string.Empty, allowedFamilyNames, newFamilyNames, allowedNewNames, unexpectedNames, DescribeDuplicateCleanupAction(cleanup));
		throw new InvalidOperationException(BuildDuplicateCleanupMessage("Dependency family load tried to create duplicate or unexpected families instead of using the standard canonical family.", CS_0024_003C_003E8__locals8._0024VB_0024Local_expectedName, unexpectedNames, cleanup));
	}

	private static void GuardLoadedRoutingFamilyDidNotCreateDuplicateFamilies(Document targetDocument, ISet<int> familyStateBefore, Family standardFamily, Family loadedFamily, IEnumerable<AllowedLoadedFamilyIdentity> allowedFamilies, SystemTypeApplyExecutionItem resultItem)
	{
		//IL_0075: Unknown result type (might be due to invalid IL or missing references)
		_Closure_0024__120_002D0 arg = default(_Closure_0024__120_002D0);
		_Closure_0024__120_002D0 CS_0024_003C_003E8__locals9 = new _Closure_0024__120_002D0(arg);
		CS_0024_003C_003E8__locals9._0024VB_0024Local_familyStateBefore = familyStateBefore;
		if (targetDocument == null || CS_0024_003C_003E8__locals9._0024VB_0024Local_familyStateBefore == null || standardFamily == null)
		{
			return;
		}
		CS_0024_003C_003E8__locals9._0024VB_0024Local_expectedName = ((Element)standardFamily).Name ?? string.Empty;
		List<string> allowedFamilyNames = UniqueSortedNames((allowedFamilies ?? new List<AllowedLoadedFamilyIdentity>()).Select([SpecialName] (AllowedLoadedFamilyIdentity x) => x.Name ?? string.Empty));
		List<Family> newFamilies = (from Family x in (IEnumerable)new FilteredElementCollector(targetDocument).OfClass(typeof(Family))
			where x != null && ((Element)x).Id != null && !CS_0024_003C_003E8__locals9._0024VB_0024Local_familyStateBefore.Contains(RevitElementIdCompat.CompatIntegerValue(((Element)x).Id))
			select x).ToList();
		if (newFamilies.Count == 0)
		{
			AddFamilyLoadDiagnostic(resultItem, "RoutingFamilyLoad.Guard", CS_0024_003C_003E8__locals9._0024VB_0024Local_expectedName, ((loadedFamily != null) ? ((Element)loadedFamily).Name : null) ?? string.Empty, allowedFamilyNames, new List<string>(), new List<string>(), new List<string>(), "None");
			return;
		}
		List<string> newFamilyNames = UniqueSortedNames(newFamilies.Select([SpecialName] (Family x) => ((Element)x).Name ?? string.Empty));
		List<Family> unexpected = new List<Family>();
		List<string> allowedNewNames = new List<string>();
		foreach (Family family in newFamilies)
		{
			string familyName = ((Element)family).Name ?? string.Empty;
			if (LoadedFamilyIsAllowed(family, allowedFamilies))
			{
				allowedNewNames.Add(familyName);
			}
			else
			{
				unexpected.Add(family);
			}
		}
		if (unexpected.Count == 0)
		{
			AddFamilyLoadDiagnostic(resultItem, "RoutingFamilyLoad.Guard", CS_0024_003C_003E8__locals9._0024VB_0024Local_expectedName, ((loadedFamily != null) ? ((Element)loadedFamily).Name : null) ?? string.Empty, allowedFamilyNames, newFamilyNames, allowedNewNames, new List<string>(), "None");
			List<string> allowedNestedNames = allowedNewNames.Where([SpecialName] (string x) => !string.Equals(Normalize(x), Normalize(CS_0024_003C_003E8__locals9._0024VB_0024Local_expectedName), StringComparison.Ordinal)).ToList();
			if (allowedNestedNames.Count > 0)
			{
				AddResultMessage(resultItem, "Allowed standard nested routing families loaded: " + string.Join(", ", UniqueSortedNames(allowedNestedNames)));
			}
			return;
		}
		List<string> unexpectedNames = unexpected.Select([SpecialName] (Family x) => ((Element)x).Name ?? string.Empty).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase)
			.ToList();
		DuplicateCleanupResult cleanup = TryDeleteFamilies(targetDocument, unexpected);
		AddFamilyLoadDiagnostic(resultItem, "RoutingFamilyLoad.Guard", CS_0024_003C_003E8__locals9._0024VB_0024Local_expectedName, ((loadedFamily != null) ? ((Element)loadedFamily).Name : null) ?? string.Empty, allowedFamilyNames, newFamilyNames, allowedNewNames, unexpectedNames, DescribeDuplicateCleanupAction(cleanup));
		throw new InvalidOperationException(BuildDuplicateCleanupMessage("Apply-time routing family load tried to create duplicate or unexpected families instead of using the standard canonical family.", CS_0024_003C_003E8__locals9._0024VB_0024Local_expectedName, unexpectedNames, cleanup));
	}

	private static void AddFamilyLoadDiagnostic(SystemTypeApplyExecutionItem resultItem, string stage, string expectedFamilyName, string loadedFamilyName, IEnumerable<string> allowedFamilyNames, IEnumerable<string> actualNewFamilyNames, IEnumerable<string> allowedNewFamilyNames, IEnumerable<string> unexpectedFamilyNames, string cleanupAction)
	{
		AddResultMessage(resultItem, (stage ?? "DependencyFamilyLoad.Guard") + " - expected family=" + (expectedFamilyName ?? string.Empty) + " / loaded family=" + (loadedFamilyName ?? string.Empty) + " / allowed created family names=" + JoinNamesOrNone(allowedFamilyNames) + " / actual new family names=" + JoinNamesOrNone(actualNewFamilyNames) + " / allowed new family names=" + JoinNamesOrNone(allowedNewFamilyNames) + " / blocked new family names=" + JoinNamesOrNone(unexpectedFamilyNames) + " / cleanup action=" + (cleanupAction ?? string.Empty));
	}

	private static string JoinNamesOrNone(IEnumerable<string> names)
	{
		List<string> values = UniqueSortedNames(names);
		if (values.Count == 0)
		{
			return "(none)";
		}
		return string.Join(", ", values);
	}

	private static string DescribeDuplicateCleanupAction(DuplicateCleanupResult cleanup)
	{
		if (cleanup == null || cleanup.AttemptedNames.Count == 0)
		{
			return "None";
		}
		List<string> parts = new List<string>();
		if (cleanup.DeletedNames.Count > 0)
		{
			parts.Add("Removed: " + string.Join(", ", cleanup.DeletedNames));
		}
		if (cleanup.FailedNames.Count > 0)
		{
			parts.Add("Failed cleanup: " + string.Join(", ", cleanup.FailedNames));
		}
		if (!string.IsNullOrWhiteSpace(cleanup.ExceptionMessage))
		{
			parts.Add("Cleanup error: " + cleanup.ExceptionMessage);
		}
		if (parts.Count == 0)
		{
			return "Cleanup attempted: " + string.Join(", ", cleanup.AttemptedNames);
		}
		return string.Join(" / ", parts);
	}

	private static DuplicateCleanupResult TryDeleteFamilies(Document targetDocument, IEnumerable<Family> families)
	{
		//IL_00b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ba: Expected O, but got Unknown
		//IL_00bc: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fa: Unknown result type (might be due to invalid IL or missing references)
		DuplicateCleanupResult result = new DuplicateCleanupResult();
		List<VB_0024AnonymousType_0<ElementId, string>> targets = (from x in families ?? new List<Family>()
			where x != null && ((Element)x).Id != null
			select new VB_0024AnonymousType_0<ElementId, string>(((Element)x).Id, ((Element)x).Name ?? string.Empty)).ToList();
		result.AttemptedNames = UniqueSortedNames(targets.Select([SpecialName] (VB_0024AnonymousType_0<ElementId, string> x) => x.Name));
		DuplicateCleanupResult TryDeleteFamilies;
		if (targets.Count == 0)
		{
			TryDeleteFamilies = result;
		}
		else
		{
			try
			{
				Transaction tx = new Transaction(targetDocument, "KKY Family Browser Remove Duplicate Families");
				try
				{
					tx.Start();
					targetDocument.Delete((ICollection<ElementId>)targets.Select([SpecialName] (VB_0024AnonymousType_0<ElementId, string> x) => x.Id).ToList());
					tx.Commit();
				}
				finally
				{
					((IDisposable)tx)?.Dispose();
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				result.ExceptionMessage = ex2.Message;
				result.FailedNames = UniqueSortedNames(result.AttemptedNames);
				TryDeleteFamilies = result;
				ProjectData.ClearProjectError();
				goto IL_01c2;
			}
			List<string> deletedNames = new List<string>();
			List<string> failedNames = new List<string>();
			foreach (VB_0024AnonymousType_0<ElementId, string> target in targets)
			{
				if (ElementWasDeleted(targetDocument, target.Id))
				{
					deletedNames.Add(target.Name);
				}
				else
				{
					failedNames.Add(target.Name);
				}
			}
			result.DeletedNames = UniqueSortedNames(deletedNames);
			result.FailedNames = UniqueSortedNames(failedNames);
			TryDeleteFamilies = result;
		}
		goto IL_01c2;
		IL_01c2:
		return TryDeleteFamilies;
	}

	private static string BuildDuplicateCleanupMessage(string context, string expectedName, IEnumerable<string> createdNames, DuplicateCleanupResult cleanup)
	{
		if (cleanup != null && cleanup.AllDeleted)
		{
			return context + " The unexpected families were removed. Family: " + expectedName + " / Removed: " + string.Join(", ", cleanup.DeletedNames);
		}
		List<string> parts = new List<string>
		{
			"Family: " + expectedName,
			"Created: " + string.Join(", ", UniqueSortedNames(createdNames))
		};
		if (cleanup != null && cleanup.DeletedNames.Count > 0)
		{
			parts.Add("Removed: " + string.Join(", ", cleanup.DeletedNames));
		}
		if (cleanup != null && cleanup.FailedNames.Count > 0)
		{
			parts.Add("Failed cleanup: " + string.Join(", ", cleanup.FailedNames));
		}
		if (cleanup != null && !string.IsNullOrWhiteSpace(cleanup.ExceptionMessage))
		{
			parts.Add("Cleanup error: " + cleanup.ExceptionMessage);
		}
		return context + " Automatic cleanup failed for one or more families. Admin cleanup is required. " + string.Join(" / ", parts);
	}

	private static bool ElementWasDeleted(Document targetDocument, ElementId id)
	{
		bool ElementWasDeleted;
		try
		{
			ElementWasDeleted = targetDocument.GetElement(id) == null;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ElementWasDeleted = false;
			ProjectData.ClearProjectError();
		}
		return ElementWasDeleted;
	}

	private static List<string> UniqueSortedNames(IEnumerable<string> names)
	{
		return (from x in names ?? new List<string>()
			select (x ?? string.Empty).Trim() into x
			where !string.IsNullOrWhiteSpace(x)
			select x).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
	}

	private static string RemoveDuplicateSuffix(string value)
	{
		string input = (value ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(input))
		{
			return string.Empty;
		}
		checked
		{
			if (input.EndsWith(")", StringComparison.Ordinal))
			{
				int openIndex = input.LastIndexOf('(');
				if (openIndex > 0 && int.TryParse(input.Substring(openIndex + 1, input.Length - openIndex - 2), out var _))
				{
					return input.Substring(0, openIndex).Trim();
				}
			}
			int lastSpace = input.LastIndexOf(' ');
			if (lastSpace > 0 && lastSpace < input.Length - 1 && int.TryParse(input.Substring(lastSpace + 1), out var _))
			{
				return input.Substring(0, lastSpace).Trim();
			}
			return input;
		}
	}

	private static bool FingerprintEquals(string left, string right)
	{
		return string.Equals(Normalize(left), Normalize(right), StringComparison.Ordinal);
	}

	private static bool CategoryNamesMatch(string left, string right)
	{
		if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
		{
			return string.IsNullOrWhiteSpace(left) && string.IsNullOrWhiteSpace(right);
		}
		return string.Equals(Normalize(left), Normalize(right), StringComparison.Ordinal);
	}

	private static void ConsolidateObsoleteTypes(Document targetDocument, IDictionary<int, List<ElementId>> targetUsageMap, IEnumerable<ElementType> obsoleteTypes, ElementType canonicalType, ref int retypedCount, ref int deletedCount)
	{
		HashSet<int> processedTypeIds = new HashSet<int>();
		checked
		{
			foreach (ElementType obsoleteType in obsoleteTypes.Where([SpecialName] (ElementType x) => x != null))
			{
				if (((Element)obsoleteType).Id == null || ((Element)obsoleteType).Id == ((Element)canonicalType).Id)
				{
					continue;
				}
				int obsoleteTypeId = RevitElementIdCompat.CompatIntegerValue(((Element)obsoleteType).Id);
				if (processedTypeIds.Contains(obsoleteTypeId))
				{
					continue;
				}
				processedTypeIds.Add(obsoleteTypeId);
				List<ElementId> elementIds = null;
				if (targetUsageMap.TryGetValue(obsoleteTypeId, out elementIds))
				{
					foreach (ElementId elementId in elementIds.ToList())
					{
						Element instanceElement = targetDocument.GetElement(elementId);
						if (instanceElement != null)
						{
							instanceElement.ChangeTypeId(((Element)canonicalType).Id);
							retypedCount++;
						}
					}
					List<ElementId> canonicalUsage = null;
					int canonicalTypeId = RevitElementIdCompat.CompatIntegerValue(((Element)canonicalType).Id);
					if (!targetUsageMap.TryGetValue(canonicalTypeId, out canonicalUsage))
					{
						canonicalUsage = (targetUsageMap[canonicalTypeId] = new List<ElementId>());
					}
					canonicalUsage.AddRange(elementIds);
					targetUsageMap.Remove(obsoleteTypeId);
				}
				targetDocument.Delete(((Element)obsoleteType).Id);
				deletedCount++;
			}
		}
	}

	private static IEnumerable<ElementType> ResolveDuplicateTypes(Document targetDocument, SystemTypeSyncPlanItem syncItem)
	{
		List<ElementType> duplicates = new List<ElementType>();
		Dictionary<string, ElementType> currentTargetMap = BuildSystemTypeMap(targetDocument);
		foreach (string duplicateName in syncItem.RelatedDuplicateNames.OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase))
		{
			ElementType duplicateType = null;
			if (currentTargetMap.TryGetValue(SystemTypeIdentityService.BuildKey(syncItem.SystemFamilyKind, syncItem.CategoryName, duplicateName), out duplicateType) && duplicateType != null)
			{
				duplicates.Add(duplicateType);
			}
		}
		return duplicates;
	}

	private static ElementType ResolveStandardSystemType(IDictionary<string, ElementType> standardSystemMap, SystemTypeSyncPlanItem syncItem)
	{
		ElementType sourceType = null;
		standardSystemMap.TryGetValue(BuildIdentityKey(syncItem), out sourceType);
		return sourceType;
	}

	private static SystemTypeSyncPlanItem FindSyncPlanItem(SystemTypeSyncPlan syncPlan, string systemFamilyKind, string categoryName, string typeName)
	{
		return (syncPlan?.Items ?? new List<SystemTypeSyncPlanItem>()).FirstOrDefault([SpecialName] (SystemTypeSyncPlanItem x) => string.Equals(BuildIdentityKey(x), SystemTypeIdentityService.BuildKey(systemFamilyKind, categoryName, typeName), StringComparison.Ordinal));
	}

	private static IEnumerable<RoutingDependencyPreflightItem> FindDependencyItems(RoutingDependencyPreflightPlan dependencyPlan, string systemFamilyKind, string typeName)
	{
		return (dependencyPlan?.Items ?? new List<RoutingDependencyPreflightItem>()).Where([SpecialName] (RoutingDependencyPreflightItem x) => string.Equals(Normalize(x.SystemFamilyKind), Normalize(systemFamilyKind), StringComparison.Ordinal) && string.Equals(Normalize(x.SystemTypeName), Normalize(typeName), StringComparison.Ordinal)).OrderBy([SpecialName] (RoutingDependencyPreflightItem x) => Normalize(x.DependencyRole), StringComparer.Ordinal).ThenBy([SpecialName] (RoutingDependencyPreflightItem x) => Normalize(x.SourceFamilyName), StringComparer.Ordinal)
			.ThenBy([SpecialName] (RoutingDependencyPreflightItem x) => Normalize(x.SourceTypeName), StringComparer.Ordinal);
	}

	private static List<RoutingDependencyPreflightItem> CollectDependencyItems(RoutingDependencyPreflightPlan dependencyPlan)
	{
		return (dependencyPlan?.Items ?? new List<RoutingDependencyPreflightItem>()).Where([SpecialName] (RoutingDependencyPreflightItem x) => x != null).ToList();
	}

	private static ISet<string> CollectSystemTypeKeys(SystemTypeSyncPlan syncPlan)
	{
		HashSet<string> result = new HashSet<string>(StringComparer.Ordinal);
		if (syncPlan == null || syncPlan.Items == null)
		{
			return result;
		}
		foreach (SystemTypeSyncPlanItem item in syncPlan.Items)
		{
			if (item != null)
			{
				result.Add(BuildIdentityKey(item));
			}
		}
		return result;
	}

	private static IDictionary<int, List<ElementId>> EnsureTargetUsageMap(Document targetDocument, ref Dictionary<int, List<ElementId>> targetUsageMap)
	{
		if (targetUsageMap == null)
		{
			targetUsageMap = BuildTypeUsageMap(targetDocument);
		}
		return targetUsageMap;
	}

	private static Dictionary<string, ElementType> BuildSystemTypeMap(Document doc, ISet<string> requestedKeys = null)
	{
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		Dictionary<string, ElementType> result = new Dictionary<string, ElementType>(StringComparer.Ordinal);
		bool hasFilter = requestedKeys != null && requestedKeys.Count > 0;
		foreach (ElementType elementType in from ElementType x in (IEnumerable)new FilteredElementCollector(doc).WhereElementIsElementType()
			where x != null
			where !(x is FamilySymbol)
			select x)
		{
			string key = SystemTypeIdentityService.BuildKey(((object)elementType).GetType().Name, ResolveCategoryName((Element)(object)elementType), ((Element)elementType).Name);
			if (!hasFilter || requestedKeys.Contains(key))
			{
				result[key] = elementType;
			}
		}
		return result;
	}

	private static Dictionary<string, Family> BuildStandardFamilyMap(Document doc, IEnumerable<RoutingDependencyPreflightItem> dependencyItems = null)
	{
		//IL_00b1: Unknown result type (might be due to invalid IL or missing references)
		Dictionary<string, Family> result = new Dictionary<string, Family>(StringComparer.Ordinal);
		HashSet<string> requestedFamilyIds = new HashSet<string>(StringComparer.Ordinal);
		HashSet<string> requestedFamilyNames = new HashSet<string>(StringComparer.Ordinal);
		foreach (RoutingDependencyPreflightItem item in dependencyItems ?? new List<RoutingDependencyPreflightItem>())
		{
			if (item != null)
			{
				if (!string.IsNullOrWhiteSpace(item.SourceLibraryFamilyId))
				{
					requestedFamilyIds.Add(Normalize(item.SourceLibraryFamilyId));
				}
				if (!string.IsNullOrWhiteSpace(item.SourceFamilyName))
				{
					requestedFamilyNames.Add(Normalize(item.SourceFamilyName));
				}
			}
		}
		bool hasFilter = requestedFamilyIds.Count > 0 || requestedFamilyNames.Count > 0;
		foreach (Family family in from Family x in (IEnumerable)new FilteredElementCollector(doc).OfClass(typeof(Family))
			where x != null
			select x)
		{
			string familyId = Normalize(((Element)family).UniqueId ?? string.Empty);
			string familyName = Normalize(((Element)family).Name ?? string.Empty);
			if (!hasFilter || requestedFamilyIds.Contains(familyId) || requestedFamilyNames.Contains(familyName))
			{
				result[BuildFamilyKey(ResolveFamilyCategoryName(family), ((Element)family).Name)] = family;
				if (!string.IsNullOrWhiteSpace(familyId))
				{
					result[familyId] = family;
				}
			}
		}
		return result;
	}

	private static Family ResolveStandardDependencyFamily(IDictionary<string, Family> standardFamilyMap, RoutingDependencyPreflightItem dependencyItem)
	{
		if (standardFamilyMap == null || dependencyItem == null)
		{
			return null;
		}
		Family standardFamily = null;
		if (!string.IsNullOrWhiteSpace(dependencyItem.SourceLibraryFamilyId) && standardFamilyMap.TryGetValue(Normalize(dependencyItem.SourceLibraryFamilyId), out standardFamily) && standardFamily != null)
		{
			return standardFamily;
		}
		string text = Normalize(dependencyItem.SourceFamilyName);
		if (string.IsNullOrWhiteSpace(text))
		{
			return null;
		}
		List<Family> matches = standardFamilyMap.Values.Where([SpecialName] (Family x) => x != null && string.Equals(Normalize(((Element)x).Name), text, StringComparison.Ordinal)).Distinct().ToList();
		if (matches.Count == 1)
		{
			return matches[0];
		}
		if (matches.Count > 1)
		{
			throw new InvalidOperationException(T("Multiple standard dependency families share the same name. Use a canonical category identity before system type apply: ", "여러 표준 의존 패밀리가 같은 이름을 사용합니다. 시스템 타입 적용 전에 표준 카테고리 식별자를 정리하세요: ") + dependencyItem.SourceFamilyName);
		}
		return null;
	}

	private static Dictionary<int, List<ElementId>> BuildTypeUsageMap(Document doc)
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		Dictionary<int, List<ElementId>> result = new Dictionary<int, List<ElementId>>();
		foreach (Element instanceElement in from Element x in (IEnumerable)new FilteredElementCollector(doc).WhereElementIsNotElementType()
			where x != null
			select x)
		{
			ElementId typeId = instanceElement.GetTypeId();
			if (typeId != null && !(typeId == ElementId.InvalidElementId) && RevitElementIdCompat.CompatIntegerValue(typeId) > 0)
			{
				List<ElementId> elementIds = null;
				if (!result.TryGetValue(RevitElementIdCompat.CompatIntegerValue(typeId), out elementIds))
				{
					elementIds = new List<ElementId>();
					result[RevitElementIdCompat.CompatIntegerValue(typeId)] = elementIds;
				}
				elementIds.Add(instanceElement.Id);
			}
		}
		return result;
	}

	private static string GenerateTemporaryTypeName(Document targetDocument, string baseTypeName)
	{
		int sequence = 1;
		_Closure_0024__146_002D0 closure_0024__146_002D = default(_Closure_0024__146_002D0);
		while (true)
		{
			closure_0024__146_002D = new _Closure_0024__146_002D0(closure_0024__146_002D);
			closure_0024__146_002D._0024VB_0024Local_candidate = baseTypeName + " [KKY_FB_BACKUP_" + sequence.ToString(CultureInfo.InvariantCulture) + "]";
			if (!BuildSystemTypeMap(targetDocument).Values.Any(closure_0024__146_002D._Lambda_0024__0))
			{
				break;
			}
			sequence = checked(sequence + 1);
		}
		return closure_0024__146_002D._0024VB_0024Local_candidate;
	}

	private static string BuildDependencyFamilyKey(RoutingDependencyPreflightItem item)
	{
		if (!string.IsNullOrWhiteSpace(item.SourceLibraryFamilyId))
		{
			return Normalize(item.SourceLibraryFamilyId);
		}
		return BuildFamilyKey(string.Empty, item.SourceFamilyName);
	}

	private static string BuildFamilyKey(string categoryName, string familyName)
	{
		return Normalize(categoryName) + "|" + Normalize(familyName);
	}

	private static string BuildIdentityKey(SystemTypeSyncPlanItem item)
	{
		return SystemTypeIdentityService.BuildKey(item.SystemFamilyKind, item.CategoryName, item.SourceTypeName);
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

	private static string ResolveExceptionMessage(Exception ex)
	{
		if (ex == null)
		{
			return string.Empty;
		}
		if (ex is TargetInvocationException { InnerException: not null } invocation)
		{
			return invocation.InnerException.Message;
		}
		if (ex.InnerException != null && !string.IsNullOrWhiteSpace(ex.InnerException.Message))
		{
			return ex.InnerException.Message;
		}
		return ex.Message;
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

	private static bool IsSupportedDependencyAction(string actionName)
	{
		switch (Normalize(actionName))
		{
		case "reuseloadeddependency":
		case "loadmissingdependencyfamily":
		case "reloadfamilyoverwrite":
		case "reuseandcleanupduplicatetypes":
		case "promoteorrenamedependencytype":
			return true;
		default:
			return false;
		}
	}

	private static void AddSystemApplyLog(SystemTypeApplyExecutionItem resultItem, string stage, string detail)
	{
		string line = (stage ?? "SystemApply") + (string.IsNullOrWhiteSpace(detail) ? string.Empty : (" - " + detail));
		AddResultMessage(resultItem, line);
	}

	private static void AddResultMessage(SystemTypeApplyExecutionItem resultItem, string message)
	{
		if (resultItem != null && !string.IsNullOrWhiteSpace(message))
		{
			resultItem.Messages.Add(message);
		}
	}

	private static void RegenerateAfterFamilyLoad(Document targetDocument, SystemTypeApplyExecutionItem resultItem, string stage, string familyName = "")
	{
		if (targetDocument == null)
		{
			return;
		}
		if (!targetDocument.IsModifiable)
		{
			AddSystemApplyLog(resultItem, stage + ".Skipped", "Document is not modifiable after family load." + (string.IsNullOrWhiteSpace(familyName) ? string.Empty : (" Family=" + familyName)));
			return;
		}
		try
		{
			targetDocument.Regenerate();
			AddSystemApplyLog(resultItem, stage, "Target document regenerated after family load." + (string.IsNullOrWhiteSpace(familyName) ? string.Empty : (" Family=" + familyName)));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			AddSystemApplyLog(resultItem, stage + ".Failed", ResolveExceptionMessage(ex2) + (string.IsNullOrWhiteSpace(familyName) ? string.Empty : (" Family=" + familyName)));
			ProjectData.ClearProjectError();
		}
	}

	private static string BuildBlockedReason(SystemSyncExecutionItem executionPlanItem)
	{
		if (executionPlanItem == null)
		{
			return "Manual review is required before this system type can be synchronized.";
		}
		if (executionPlanItem.BlockingReasons != null && executionPlanItem.BlockingReasons.Count > 0)
		{
			return string.Join("\r\n", executionPlanItem.BlockingReasons);
		}
		if (!string.IsNullOrWhiteSpace(executionPlanItem.Summary))
		{
			return executionPlanItem.Summary;
		}
		return "Manual review is required before this system type can be synchronized.";
	}

	private static ApplySummarySnapshot CaptureApplySummarySnapshot(SystemTypeApplyExecutionSummary summary)
	{
		if (summary == null)
		{
			return new ApplySummarySnapshot();
		}
		return new ApplySummarySnapshot
		{
			CreatedCount = summary.CreatedCount,
			OverwrittenCount = summary.OverwrittenCount,
			ConsolidatedCount = summary.ConsolidatedCount,
			DependencyLoadedCount = summary.DependencyLoadedCount,
			RetypedElementCount = summary.RetypedElementCount,
			DeletedObsoleteTypeCount = summary.DeletedObsoleteTypeCount,
			TrackingRefreshedCount = summary.TrackingRefreshedCount,
			BlockedCount = summary.BlockedCount,
			SkippedCount = summary.SkippedCount,
			FailedCount = summary.FailedCount
		};
	}

	private static void RestoreApplySummarySnapshot(SystemTypeApplyExecutionSummary summary, ApplySummarySnapshot snapshot)
	{
		if (summary != null && snapshot != null)
		{
			summary.CreatedCount = snapshot.CreatedCount;
			summary.OverwrittenCount = snapshot.OverwrittenCount;
			summary.ConsolidatedCount = snapshot.ConsolidatedCount;
			summary.DependencyLoadedCount = snapshot.DependencyLoadedCount;
			summary.RetypedElementCount = snapshot.RetypedElementCount;
			summary.DeletedObsoleteTypeCount = snapshot.DeletedObsoleteTypeCount;
			summary.TrackingRefreshedCount = snapshot.TrackingRefreshedCount;
			summary.BlockedCount = snapshot.BlockedCount;
			summary.SkippedCount = snapshot.SkippedCount;
			summary.FailedCount = snapshot.FailedCount;
		}
	}

	private static void RestoreRefreshedDependencyFamilies(ISet<string> refreshedDependencyFamilies, ISet<string> snapshot)
	{
		if (refreshedDependencyFamilies == null || snapshot == null)
		{
			return;
		}
		refreshedDependencyFamilies.Clear();
		foreach (string value in snapshot)
		{
			refreshedDependencyFamilies.Add(value);
		}
	}

	private static void TrimDependencyActions(SystemTypeApplyExecutionItem resultItem, int countBefore)
	{
		if (resultItem?.DependencyActions != null && countBefore >= 0 && countBefore < resultItem.DependencyActions.Count)
		{
			resultItem.DependencyActions.RemoveRange(countBefore, checked(resultItem.DependencyActions.Count - countBefore));
		}
	}

	private static void TryRollback(Transaction transaction)
	{
		//IL_0005: Unknown result type (might be due to invalid IL or missing references)
		//IL_000b: Invalid comparison between Unknown and I4
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		if (transaction == null)
		{
			return;
		}
		try
		{
			if ((int)transaction.GetStatus() == 1)
			{
				transaction.RollBack();
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static void TryRollback(TransactionGroup transactionGroup)
	{
		//IL_0005: Unknown result type (might be due to invalid IL or missing references)
		//IL_000b: Invalid comparison between Unknown and I4
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		if (transactionGroup == null)
		{
			return;
		}
		try
		{
			if ((int)transactionGroup.GetStatus() == 1)
			{
				transactionGroup.RollBack();
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private static string Normalize(string value)
	{
		return SystemTypeIdentityService.Normalize(value);
	}
}
