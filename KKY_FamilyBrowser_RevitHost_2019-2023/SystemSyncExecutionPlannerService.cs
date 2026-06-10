using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;

public sealed class SystemSyncExecutionPlannerService
{
	[CompilerGenerated]
	internal sealed class _Closure_0024__5_002D0
	{
		public SystemSyncExecutionItem _0024VB_0024Local_executionItem;

		public _Closure_0024__5_002D0(_Closure_0024__5_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_executionItem = arg0._0024VB_0024Local_executionItem;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__4(SystemSyncExecutionStep x)
		{
			return !string.Equals(x.TargetKind, _0024VB_0024Local_executionItem.SystemFamilyKind, StringComparison.OrdinalIgnoreCase);
		}
	}

	private SystemSyncExecutionPlannerService()
	{
	}

	public static SystemSyncExecutionPlan BuildPlan(SystemTypeSyncPlan syncPlan, RoutingDependencyPreflightPlan dependencyPlan)
	{
		SystemSyncExecutionPlan plan = new SystemSyncExecutionPlan
		{
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
		};
		if (syncPlan == null)
		{
			return plan;
		}
		Dictionary<string, List<RoutingDependencyPreflightItem>> dependencyMap = (dependencyPlan?.Items ?? new List<RoutingDependencyPreflightItem>()).GroupBy([SpecialName] (RoutingDependencyPreflightItem x) => BuildKey(x.SystemFamilyKind, x.SystemTypeName), StringComparer.Ordinal).ToDictionary([SpecialName] (IGrouping<string, RoutingDependencyPreflightItem> x) => x.Key, [SpecialName] (IGrouping<string, RoutingDependencyPreflightItem> x) => x.OrderBy([SpecialName] (RoutingDependencyPreflightItem y) => Normalize(y.DependencyRole), StringComparer.Ordinal).ThenBy([SpecialName] (RoutingDependencyPreflightItem y) => Normalize(y.SourceFamilyName), StringComparer.Ordinal).ThenBy([SpecialName] (RoutingDependencyPreflightItem y) => Normalize(y.SourceTypeName), StringComparer.Ordinal)
			.ToList(), StringComparer.Ordinal);
		foreach (SystemTypeSyncPlanItem syncItem in syncPlan.Items.OrderBy([SpecialName] (SystemTypeSyncPlanItem x) => Normalize(x.SystemFamilyKind), StringComparer.Ordinal).ThenBy([SpecialName] (SystemTypeSyncPlanItem x) => Normalize(x.SourceTypeName), StringComparer.Ordinal))
		{
			SystemSyncExecutionItem executionItem = new SystemSyncExecutionItem
			{
				SystemFamilyKind = syncItem.SystemFamilyKind,
				CategoryName = syncItem.CategoryName,
				SystemTypeName = syncItem.SourceTypeName,
				SyncAction = syncItem.Action,
				RelatedDuplicateNames = (syncItem.RelatedDuplicateNames ?? new List<string>()).ToList()
			};
			List<RoutingDependencyPreflightItem> dependencies = null;
			if (!dependencyMap.TryGetValue(BuildKey(syncItem.SystemFamilyKind, syncItem.SourceTypeName), out dependencies))
			{
				dependencies = new List<RoutingDependencyPreflightItem>();
			}
			if (!IsSkipMissingTypeAction(syncItem.Action))
			{
				AppendDependencySteps(executionItem, dependencies);
			}
			AppendSystemTypeSteps(executionItem, syncItem);
			FinalizeExecutionItem(executionItem);
			plan.Items.Add(executionItem);
		}
		return plan;
	}

	private static void AppendDependencySteps(SystemSyncExecutionItem executionItem, IEnumerable<RoutingDependencyPreflightItem> dependencies)
	{
		foreach (RoutingDependencyPreflightItem dependency in dependencies)
		{
			SystemSyncExecutionStep stepItem = new SystemSyncExecutionStep
			{
				SequenceNo = checked(executionItem.Steps.Count + 1),
				TargetKind = dependency.DependencyRole,
				TargetName = dependency.SourceFamilyName + " : " + dependency.SourceTypeName
			};
			switch (Normalize(dependency.Action))
			{
			case "reuseloadeddependency":
				stepItem.Action = "ValidateDependency";
				stepItem.Status = "Ready";
				stepItem.Notes = dependency.Reason;
				break;
			case "reuseandcleanupduplicatetypes":
				stepItem.Action = "ReloadStandardFamilyFirst";
				stepItem.Status = "ApprovalRequired";
				stepItem.Notes = dependency.Reason;
				break;
			case "loadmissingdependencyfamily":
				stepItem.Action = "ApplyStandardFamilyFirst";
				stepItem.Status = "ApprovalRequired";
				stepItem.Notes = dependency.Reason;
				break;
			case "reloadfamilyoverwrite":
				stepItem.Action = "ReloadStandardFamilyFirst";
				stepItem.Status = "ApprovalRequired";
				stepItem.Notes = dependency.Reason;
				break;
			case "promoteorrenamedependencytype":
				stepItem.Action = "ReloadStandardFamilyFirst";
				stepItem.Status = "ApprovalRequired";
				stepItem.Notes = dependency.Reason;
				break;
			case "manualreviewnameonlymatch":
				stepItem.Action = "ReviewDependencyCategory";
				stepItem.Status = "ManualReview";
				stepItem.Notes = dependency.Reason;
				break;
			default:
				stepItem.Action = "ReloadStandardFamilyFirst";
				stepItem.Status = "ApprovalRequired";
				stepItem.Notes = dependency.Reason;
				break;
			}
			executionItem.Steps.Add(stepItem);
		}
	}

	private static void AppendSystemTypeSteps(SystemSyncExecutionItem executionItem, SystemTypeSyncPlanItem syncItem)
	{
		bool hasManualReview = executionItem.Steps.Any([SpecialName] (SystemSyncExecutionStep x) => string.Equals(x.Status, "ManualReview", StringComparison.OrdinalIgnoreCase));
		bool hasFoundationBlocked = executionItem.Steps.Any([SpecialName] (SystemSyncExecutionStep x) => string.Equals(x.Status, "FoundationBlocked", StringComparison.OrdinalIgnoreCase));
		bool hasApprovalSteps = executionItem.Steps.Any([SpecialName] (SystemSyncExecutionStep x) => string.Equals(x.Status, "ApprovalRequired", StringComparison.OrdinalIgnoreCase));
		bool hasDuplicateSystemTypes = syncItem.RelatedDuplicateNames.Count > 0;
		checked
		{
			SystemSyncExecutionStep applyStep = new SystemSyncExecutionStep
			{
				SequenceNo = executionItem.Steps.Count + 1,
				TargetKind = syncItem.SystemFamilyKind,
				TargetName = syncItem.SourceTypeName
			};
			switch (Normalize(syncItem.Action))
			{
			case "keepdestination":
				applyStep.Action = (hasDuplicateSystemTypes ? "CleanupDuplicateSystemTypes" : "KeepExistingSystemType");
				applyStep.Status = DetermineApplyStatus(hasManualReview, hasFoundationBlocked, hasApprovalSteps || hasDuplicateSystemTypes);
				applyStep.Notes = syncItem.Reason;
				break;
			case "overwritedestination":
				applyStep.Action = "OverwriteSystemTypeInPlace";
				applyStep.Status = DetermineApplyStatus(hasManualReview, hasFoundationBlocked, needsApproval: true);
				applyStep.Notes = syncItem.Reason;
				break;
			case "createmissingtype":
				applyStep.Action = "CreateCanonicalSystemType";
				applyStep.Status = DetermineApplyStatus(hasManualReview, hasFoundationBlocked, needsApproval: true);
				applyStep.Notes = syncItem.Reason;
				break;
			case "skipmissingtype":
				applyStep.Action = "SkipMissingSystemType";
				applyStep.Status = "Skipped";
				applyStep.Notes = syncItem.Reason;
				break;
			case "consolidateduplicatesuffixtypes":
				applyStep.Action = "ConsolidateDuplicateSystemTypes";
				applyStep.Status = DetermineApplyStatus(hasManualReview, hasFoundationBlocked, needsApproval: true);
				applyStep.Notes = syncItem.Reason;
				break;
			case "manualreview":
				applyStep.Action = "ReviewSystemType";
				applyStep.Status = "ManualReview";
				applyStep.Notes = syncItem.Reason;
				break;
			default:
				applyStep.Action = "ReviewSystemType";
				applyStep.Status = DetermineApplyStatus(hasManualReview: true, hasFoundationBlocked, needsApproval: true);
				applyStep.Notes = syncItem.Reason;
				break;
			}
			executionItem.Steps.Add(applyStep);
			if (hasDuplicateSystemTypes && !string.Equals(Normalize(syncItem.Action), "keepdestination", StringComparison.Ordinal))
			{
				executionItem.Steps.Add(new SystemSyncExecutionStep
				{
					SequenceNo = executionItem.Steps.Count + 1,
					Action = "CleanupDuplicateSystemTypes",
					Status = DetermineApplyStatus(hasManualReview, hasFoundationBlocked, needsApproval: true),
					TargetKind = syncItem.SystemFamilyKind,
					TargetName = string.Join(", ", syncItem.RelatedDuplicateNames),
					Notes = "Duplicate-suffix system types should be remapped and removed after the canonical type is updated."
				});
			}
		}
	}

	private static string DetermineApplyStatus(bool hasManualReview, bool hasFoundationBlocked, bool needsApproval)
	{
		if (hasManualReview)
		{
			return "Blocked";
		}
		if (hasFoundationBlocked)
		{
			return "Blocked";
		}
		if (needsApproval)
		{
			return "ApprovalRequired";
		}
		return "Ready";
	}

	private static void FinalizeExecutionItem(SystemSyncExecutionItem executionItem)
	{
		_Closure_0024__5_002D0 arg = default(_Closure_0024__5_002D0);
		_Closure_0024__5_002D0 CS_0024_003C_003E8__locals30 = new _Closure_0024__5_002D0(arg);
		CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem = executionItem;
		List<SystemSyncExecutionStep> manualReviewSteps = CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.Steps.Where([SpecialName] (SystemSyncExecutionStep x) => string.Equals(x.Status, "ManualReview", StringComparison.OrdinalIgnoreCase)).ToList();
		List<SystemSyncExecutionStep> foundationBlockedSteps = CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.Steps.Where([SpecialName] (SystemSyncExecutionStep x) => string.Equals(x.Status, "FoundationBlocked", StringComparison.OrdinalIgnoreCase)).ToList();
		List<SystemSyncExecutionStep> approvalSteps = CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.Steps.Where([SpecialName] (SystemSyncExecutionStep x) => string.Equals(x.Status, "ApprovalRequired", StringComparison.OrdinalIgnoreCase)).ToList();
		CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.HasManualReview = manualReviewSteps.Count > 0;
		CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.RequiresLoadableFoundation = foundationBlockedSteps.Count > 0;
		CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.RequiresApproval = approvalSteps.Count > 0;
		if (CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.Steps.Count > 0 && CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.Steps.All([SpecialName] (SystemSyncExecutionStep x) => string.Equals(x.Status, "Skipped", StringComparison.OrdinalIgnoreCase) || string.Equals(x.Action, "SkipMissingSystemType", StringComparison.OrdinalIgnoreCase)))
		{
			CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.ExecutionStatus = "SkippedMissing";
			CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.Summary = "The registered standard system type is not loaded in the current project, so it is left uncreated.";
			return;
		}
		foreach (SystemSyncExecutionStep stepItem in foundationBlockedSteps)
		{
			CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.FoundationBlockingReasons.Add(stepItem.Action + ": " + stepItem.Notes);
			CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.BlockingReasons.Add(stepItem.Action + ": " + stepItem.Notes);
		}
		foreach (SystemSyncExecutionStep stepItem2 in manualReviewSteps)
		{
			CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.BlockingReasons.Add(stepItem2.Action + ": " + stepItem2.Notes);
		}
		if (CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.RequiresLoadableFoundation)
		{
			CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.ExecutionStatus = "Blocked";
			CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.Summary = "Loadable family foundation must match the standard before this system type can be synchronized.";
		}
		else if (CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.HasManualReview)
		{
			CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.ExecutionStatus = "Blocked";
			CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.Summary = "Manual review is required before this system type can be synchronized.";
		}
		else if (CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.RequiresApproval)
		{
			CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.ExecutionStatus = "ApprovalRequired";
			if (approvalSteps.Any([SpecialName] (SystemSyncExecutionStep x) => !string.Equals(x.TargetKind, CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.SystemFamilyKind, StringComparison.OrdinalIgnoreCase)))
			{
				CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.Summary = "This system type can proceed after dependency and overwrite approvals are granted in order.";
			}
			else
			{
				CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.Summary = "This system type is ready for approval and in-place update.";
			}
		}
		else if (CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.Steps.All([SpecialName] (SystemSyncExecutionStep x) => string.Equals(x.Action, "KeepExistingSystemType", StringComparison.OrdinalIgnoreCase) || string.Equals(x.Action, "ValidateDependency", StringComparison.OrdinalIgnoreCase)))
		{
			CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.ExecutionStatus = "NoChange";
			CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.Summary = "The loaded dependencies and system type already match the canonical source.";
		}
		else
		{
			CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.ExecutionStatus = "Ready";
			CS_0024_003C_003E8__locals30._0024VB_0024Local_executionItem.Summary = "The system type is ready to be applied without additional approvals.";
		}
	}

	private static string BuildKey(string systemFamilyKind, string typeName)
	{
		return Normalize(systemFamilyKind) + "|" + Normalize(typeName);
	}

	private static bool IsSkipMissingTypeAction(string action)
	{
		return string.Equals(Normalize(action), "skipmissingtype", StringComparison.Ordinal);
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
