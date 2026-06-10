using System;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;

public sealed class LoadableFamilySyncPlannerService
{
	private LoadableFamilySyncPlannerService()
	{
	}

	public static LoadableFamilySyncPlan BuildPlan(ProjectStandardComparisonReport comparisonReport, string comparisonPath, bool allowFingerprintOverwrite = false, bool allowMissingFamilyLoad = false)
	{
		if (comparisonReport == null)
		{
			throw new ArgumentNullException("comparisonReport");
		}
		LoadableFamilySyncPlan plan = new LoadableFamilySyncPlan
		{
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			ProjectDocumentTitle = comparisonReport.Project.DocumentTitle,
			ProjectDocumentPath = comparisonReport.Project.DocumentPath,
			StandardDisplayName = comparisonReport.Standard.DisplayName,
			ComparisonPath = comparisonPath
		};
		checked
		{
			foreach (LoadableFamilyComparisonItem item in comparisonReport.LoadableFamilies.OrderBy([SpecialName] (LoadableFamilyComparisonItem x) => Normalize(x.CategoryName) + "|" + Normalize(x.FamilyName), StringComparer.Ordinal))
			{
				if (FamilyBrowserFamilyClassificationService.IsTypeManagedFamilyLike(item.CategoryName, string.Empty, item.FamilyName))
				{
					continue;
				}
				string plannedAction = "KeepCurrent";
				string executionMode = "Skip";
				string notes = item.Notes;
				switch (Normalize(item.Status))
				{
				case "loadavailable":
					if (allowMissingFamilyLoad)
					{
						plannedAction = "LoadFromStandard";
						executionMode = "Load";
						notes = AppendNote(notes, "Registered standard family is not loaded in the current project. It will be loaded because the user explicitly selected this family.");
						plan.Summary.LoadCount++;
					}
					else
					{
						plannedAction = "SkipMissingFromProject";
						executionMode = "Skip";
						notes = AppendNote(notes, "Registered standard family is not loaded in the current project. Missing project families are not loaded automatically; only existing project families are updated.");
						plan.Summary.SkipCount++;
					}
					break;
				case "updateavailable":
					plannedAction = "ReloadFromStandard";
					executionMode = "Reload";
					plan.Summary.ReloadCount++;
					break;
				case "loadedwithoutversionstamp":
				case "stampnormalizationneeded":
					plannedAction = "RefreshTracking";
					executionMode = "StampOnly";
					plan.Summary.StampOnlyCount++;
					break;
				case "differentfromstandard":
				case "locallymodified":
				case "versionconflict":
					if (allowFingerprintOverwrite)
					{
						plannedAction = "ReloadFromStandard";
						executionMode = "Reload";
						notes = AppendNote(notes, "Admin bulk apply will overwrite the project family because its fingerprint differs from the registered standard.");
						plan.Summary.ReloadCount++;
					}
					else
					{
						plannedAction = "BlockedManualReview";
						executionMode = "Blocked";
						plan.Summary.BlockedCount++;
					}
					break;
				case "categorymismatch":
					plannedAction = "BlockedCategoryMismatch";
					executionMode = "Blocked";
					notes = AppendNote(notes, "Category mismatch is not a fingerprint overwrite target. Review the family category before applying a standard family.");
					plan.Summary.BlockedCount++;
					break;
				case "manualreview":
					plannedAction = "BlockedManualReview";
					executionMode = "Blocked";
					notes = AppendNote(notes, "Manual review is required before applying this family.");
					plan.Summary.BlockedCount++;
					break;
				case "projectonly":
					plannedAction = "KeepProjectOnly";
					executionMode = "Skip";
					plan.Summary.SkipCount++;
					break;
				default:
					plannedAction = "KeepCurrent";
					executionMode = "Skip";
					plan.Summary.SkipCount++;
					break;
				}
				plan.Items.Add(new LoadableFamilySyncPlanItem
				{
					IdentityKey = item.IdentityKey,
					FamilyName = item.FamilyName,
					CategoryName = item.CategoryName,
					ComparisonStatus = item.Status,
					PlannedAction = plannedAction,
					ExecutionMode = executionMode,
					Notes = notes
				});
			}
			return plan;
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

	private static string AppendNote(string existingNotes, string note)
	{
		if (string.IsNullOrWhiteSpace(existingNotes))
		{
			return note;
		}
		if (string.IsNullOrWhiteSpace(note))
		{
			return existingNotes;
		}
		return existingNotes.Trim() + " " + note;
	}
}
