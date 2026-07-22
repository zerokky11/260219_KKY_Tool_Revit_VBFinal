using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;

public sealed class ProjectStandardComparisonService
{
	private class ComparisonClassificationResult
	{
		public string Status { get; set; }

		public string Notes { get; set; }

		public ComparisonClassificationResult()
		{
			Status = string.Empty;
			Notes = string.Empty;
		}
	}

	private class SignatureLineDelta
	{
		public List<string> StandardOnly { get; set; }

		public List<string> ProjectOnly { get; set; }

		public SignatureLineDelta()
		{
			StandardOnly = new List<string>();
			ProjectOnly = new List<string>();
		}
	}

	private class LookupCsvSignatureTableInfo
	{
		public string Name { get; set; }

		public int RowCount { get; set; }

		public int ColumnCount { get; set; }

		public bool HasRowCount { get; set; }

		public bool HasColumnCount { get; set; }

		public bool Missing { get; set; }

		public bool HasError { get; set; }

		public string Signature { get; set; }

		public LookupCsvSignatureTableInfo()
		{
			Name = string.Empty;
			Signature = string.Empty;
		}
	}

	private class NestedLabelSignatureInfo
	{
		public string InstanceKey { get; set; }

		public string InstanceDisplay { get; set; }

		public string ParameterName { get; set; }

		public string LabelName { get; set; }

		public string NestedFamilyName { get; set; }

		public string NestedTypeName { get; set; }

		public string NestedCategoryName { get; set; }

		public string StorageName { get; set; }

		public string RoleName { get; set; }

		public string Formula { get; set; }

		public string ComparisonValue { get; set; }

		public NestedLabelSignatureInfo()
		{
			InstanceKey = string.Empty;
			InstanceDisplay = string.Empty;
			ParameterName = string.Empty;
			LabelName = string.Empty;
			NestedFamilyName = string.Empty;
			NestedTypeName = string.Empty;
			NestedCategoryName = string.Empty;
			StorageName = string.Empty;
			RoleName = string.Empty;
			Formula = string.Empty;
			ComparisonValue = string.Empty;
		}
	}

	private class RoutingPreferenceSignatureRule
	{
		public string GroupName { get; set; }

		public string RuleIndex { get; set; }

		public string PartClass { get; set; }

		public string PartCategory { get; set; }

		public string FamilyKey { get; set; }

		public string FamilyName { get; set; }

		public string TypeName { get; set; }

		public string FamilyFingerprint { get; set; }

		public string TypeFingerprint { get; set; }

		public string PartFingerprint { get; set; }

		public string CriteriaSignature { get; set; }

		public string RuleKey => (GroupName ?? string.Empty).Trim().ToLowerInvariant() + "|" + (RuleIndex ?? string.Empty).Trim().ToLowerInvariant();

		public RoutingPreferenceSignatureRule()
		{
			GroupName = string.Empty;
			RuleIndex = string.Empty;
			PartClass = string.Empty;
			PartCategory = string.Empty;
			FamilyKey = string.Empty;
			FamilyName = string.Empty;
			TypeName = string.Empty;
			FamilyFingerprint = string.Empty;
			TypeFingerprint = string.Empty;
			PartFingerprint = string.Empty;
			CriteriaSignature = string.Empty;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__63_002D0
	{
		public string _0024VB_0024Local_typeName;

		public _Closure_0024__63_002D0(_Closure_0024__63_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_typeName = arg0._0024VB_0024Local_typeName;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(string x)
		{
			return string.Equals(x, _0024VB_0024Local_typeName, StringComparison.OrdinalIgnoreCase);
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__71_002D0
	{
		public string _0024VB_0024Local_difference;

		public _Closure_0024__71_002D0(_Closure_0024__71_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_difference = arg0._0024VB_0024Local_difference;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(string x)
		{
			return string.Equals(Normalize(x), Normalize(_0024VB_0024Local_difference), StringComparison.Ordinal);
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__9_002D0
	{
		public HashSet<string> _0024VB_0024Local_nestedFamilyNames;

		public _Closure_0024__9_002D0(_Closure_0024__9_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_nestedFamilyNames = arg0._0024VB_0024Local_nestedFamilyNames;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__1(StandardLoadableFamilySnapshotItem x)
		{
			return !IsHiddenNestedLoadableChild(x, _0024VB_0024Local_nestedFamilyNames);
		}

		[SpecialName]
		internal bool _Lambda_0024__3(ProjectLoadableFamilySnapshotItem x)
		{
			return !IsHiddenNestedLoadableChild(x, _0024VB_0024Local_nestedFamilyNames);
		}
	}

	private const string IdentityModeName = "NameCategoryTypeAndCanonicalContentFingerprint";

	private ProjectStandardComparisonService()
	{
	}

	public static ProjectStandardComparisonReport BuildReport(StandardLibraryRegistrationRecord registration, string snapshotPath, StandardLibrarySnapshot standardSnapshot, string projectSnapshotPath, ProjectContentSnapshot projectSnapshot, ProjectTrackingCatalog trackingCatalog = null, bool compareDetailedSystemTypeComponents = true)
	{
		ProjectTrackingCatalog effectiveTrackingCatalog = trackingCatalog;
		string trackingState = "NoTrackingCatalog";
		if (effectiveTrackingCatalog != null)
		{
			if (string.Equals(Normalize(effectiveTrackingCatalog.SourceId), Normalize(registration.SourceId), StringComparison.Ordinal))
			{
				trackingState = "TrackingCatalogLoaded";
			}
			else
			{
				effectiveTrackingCatalog = null;
				trackingState = "TrackingCatalogSourceMismatch";
			}
		}
		RecoverStandardLoadableSignatureDebugMetadata(snapshotPath, standardSnapshot);
		RecoverProjectLoadableSignatureDebugMetadata(projectSnapshotPath, projectSnapshot);
		ProjectStandardComparisonReport report = new ProjectStandardComparisonReport
		{
			IdentityMode = "NameCategoryTypeAndCanonicalContentFingerprint",
			TrackingState = trackingState,
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			Standard = new ProjectStandardReference
			{
				SourceId = registration.SourceId,
				DisplayName = registration.DisplayName,
				SourceKind = registration.SourceKind,
				ResolvedPath = registration.ResolvedPath,
				SnapshotPath = snapshotPath,
				RevitVersion = registration.RevitVersion
			},
			Project = new ProjectReference
			{
				DocumentTitle = projectSnapshot.DocumentTitle,
				DocumentPath = projectSnapshot.DocumentPath,
				RevitVersion = projectSnapshot.RevitVersion,
				SnapshotPath = projectSnapshotPath
			}
		};
		report.ProjectLoadableSignatureFailures = BuildProjectLoadableSignatureFailures(projectSnapshot);
		report.LoadableFamilies = BuildLoadableFamilyComparisons(standardSnapshot, projectSnapshot, effectiveTrackingCatalog);
		report.SystemTypes = BuildSystemTypeComparisons(standardSnapshot, projectSnapshot, effectiveTrackingCatalog, compareDetailedSystemTypeComponents);
		ProjectStandardComparisonSummary summary = new ProjectStandardComparisonSummary();
		checked
		{
			foreach (LoadableFamilyComparisonItem loadableFamily in report.LoadableFamilies)
			{
				switch (loadableFamily.Status)
				{
				case "LoadedLatest":
					summary.LoadableLatestCount++;
					break;
				case "LoadAvailable":
					summary.LoadableLoadAvailableCount++;
					break;
				case "DifferentFromStandard":
				case "LoadedWithoutVersionStamp":
				case "StampNormalizationNeeded":
				case "UpdateAvailable":
				case "LocallyModified":
				case "VersionConflict":
				case "CategoryMismatch":
				case "ManualReview":
				case "NestedMissingFromParent":
				case "NestedExtraInParent":
					summary.LoadableDifferentCount++;
					break;
				case "ProjectOnly":
					summary.LoadableProjectOnlyCount++;
					break;
				}
			}
			foreach (SystemTypeComparisonItem systemType in report.SystemTypes)
			{
				switch (systemType.Status)
				{
				case "LoadedLatest":
					summary.SystemLatestCount++;
					break;
				case "LoadAvailable":
					summary.SystemLoadAvailableCount++;
					break;
				case "DifferentFromStandard":
				case "LoadedWithoutVersionStamp":
				case "StampNormalizationNeeded":
				case "UpdateAvailable":
				case "LocallyModified":
				case "VersionConflict":
				case "CategoryMismatch":
				case "ManualReview":
					summary.SystemDifferentCount++;
					break;
				case "ProjectOnly":
					summary.SystemProjectOnlyCount++;
					break;
				}
			}
			summary.LoadableSignatureFailureCount = ((report.ProjectLoadableSignatureFailures != null) ? report.ProjectLoadableSignatureFailures.Count : 0);
			report.Summary = summary;
			return report;
		}
	}

	private static List<ProjectLoadableSignatureFailureItem> BuildProjectLoadableSignatureFailures(ProjectContentSnapshot projectSnapshot)
	{
		if (projectSnapshot == null)
		{
			return new List<ProjectLoadableSignatureFailureItem>();
		}
		return ProjectSnapshotCaptureService.BuildLoadableSignatureFailures(projectSnapshot.LoadableFamilies).OrderBy<ProjectLoadableSignatureFailureItem, string>([SpecialName] (ProjectLoadableSignatureFailureItem x) => x.CategoryName ?? string.Empty, StringComparer.OrdinalIgnoreCase).ThenBy<ProjectLoadableSignatureFailureItem, string>([SpecialName] (ProjectLoadableSignatureFailureItem x) => x.FamilyName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
			.ToList();
	}

	private static void RecoverStandardLoadableSignatureDebugMetadata(string snapshotPath, StandardLibrarySnapshot snapshot)
	{
		if (snapshot == null || snapshot.LoadableFamilies == null || snapshot.LoadableFamilies.Count == 0)
		{
			return;
		}
		List<string> signaturePaths = (from x in snapshot.LoadableFamilies
			where x != null
			select x.ContentSignatureDebugPath ?? string.Empty into x
			where !string.IsNullOrWhiteSpace(x)
			select x).ToList();
		List<string> signatureRoots = BuildSignatureDebugRootCandidates(snapshotPath, signaturePaths);
		Dictionary<string, List<LoadableSignatureDebugRecord>> signatureIndex = FingerprintDebugSignatureStore.BuildLoadableSignatureIndex(signaturePaths, signatureRoots);
		foreach (StandardLoadableFamilySnapshotItem item in snapshot.LoadableFamilies)
		{
			if (item == null || string.IsNullOrWhiteSpace(item.FamilyName))
			{
				continue;
			}
			LoadableSignatureDebugRecord record = FingerprintDebugSignatureStore.FindBestLoadableSignatureRecord(signatureIndex, item.CategoryName, item.FamilyName);
			if (record != null)
			{
				if (SignaturePathNeedsRecovery(item.ContentSignatureDebugPath))
				{
					item.ContentSignatureDebugPath = record.Path ?? string.Empty;
				}
				if (string.IsNullOrWhiteSpace(item.ContentFingerprint) && !string.IsNullOrWhiteSpace(record.Fingerprint))
				{
					item.ContentFingerprint = record.Fingerprint;
					item.ContentFingerprintFailureReason = string.Empty;
				}
				else if (string.IsNullOrWhiteSpace(item.ContentFingerprintFailureReason) && !string.IsNullOrWhiteSpace(record.ErrorMessage))
				{
					item.ContentFingerprintFailureReason = record.ErrorMessage;
				}
			}
		}
	}

	private static void RecoverProjectLoadableSignatureDebugMetadata(string snapshotPath, ProjectContentSnapshot snapshot)
	{
		if (snapshot == null || snapshot.LoadableFamilies == null || snapshot.LoadableFamilies.Count == 0)
		{
			return;
		}
		List<string> signaturePaths = (from x in snapshot.LoadableFamilies
			where x != null
			select x.ContentSignatureDebugPath ?? string.Empty into x
			where !string.IsNullOrWhiteSpace(x)
			select x).ToList();
		if (snapshot.LoadableSignatureFailures != null)
		{
			signaturePaths.AddRange(from x in snapshot.LoadableSignatureFailures
				where x != null
				select x.ContentSignatureDebugPath ?? string.Empty into x
				where !string.IsNullOrWhiteSpace(x)
				select x);
		}
		List<string> signatureRoots = BuildSignatureDebugRootCandidates(snapshotPath, signaturePaths);
		Dictionary<string, List<LoadableSignatureDebugRecord>> signatureIndex = FingerprintDebugSignatureStore.BuildLoadableSignatureIndex(signaturePaths, signatureRoots);
		foreach (ProjectLoadableFamilySnapshotItem item in snapshot.LoadableFamilies)
		{
			if (item == null || string.IsNullOrWhiteSpace(item.FamilyName))
			{
				continue;
			}
			LoadableSignatureDebugRecord record = FingerprintDebugSignatureStore.FindBestLoadableSignatureRecord(signatureIndex, item.CategoryName, item.FamilyName);
			if (record != null)
			{
				if (SignaturePathNeedsRecovery(item.ContentSignatureDebugPath))
				{
					item.ContentSignatureDebugPath = record.Path ?? string.Empty;
				}
				if (string.IsNullOrWhiteSpace(item.ContentFingerprint) && !string.IsNullOrWhiteSpace(record.Fingerprint))
				{
					item.ContentFingerprint = record.Fingerprint;
					item.ContentFingerprintFailureReason = string.Empty;
				}
				else if (string.IsNullOrWhiteSpace(item.ContentFingerprintFailureReason) && !string.IsNullOrWhiteSpace(record.ErrorMessage))
				{
					item.ContentFingerprintFailureReason = record.ErrorMessage;
				}
			}
		}
	}

	private static List<string> BuildSignatureDebugRootCandidates(string snapshotPath, IEnumerable<string> signaturePaths)
	{
		List<string> roots = new List<string>();
		if (signaturePaths != null)
		{
			foreach (string signaturePath in signaturePaths)
			{
				try
				{
					string expandedPath = Environment.ExpandEnvironmentVariables((signaturePath ?? string.Empty).Trim());
					if (string.IsNullOrWhiteSpace(expandedPath))
					{
						continue;
					}
					string parentFolder = Path.GetDirectoryName(expandedPath);
					if (string.IsNullOrWhiteSpace(parentFolder))
					{
						continue;
					}
					AddDistinctPath(roots, parentFolder);
					string runFolder = Path.GetDirectoryName(parentFolder);
					if (!string.IsNullOrWhiteSpace(runFolder))
					{
						AddDistinctPath(roots, runFolder);
						string debugFolder = Path.GetDirectoryName(runFolder);
						if (!string.IsNullOrWhiteSpace(debugFolder))
						{
							AddDistinctPath(roots, debugFolder);
						}
					}
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
		}
		try
		{
			string expandedSnapshotPath = Environment.ExpandEnvironmentVariables((snapshotPath ?? string.Empty).Trim());
			if (!string.IsNullOrWhiteSpace(expandedSnapshotPath))
			{
				string snapshotFolder = Path.GetDirectoryName(expandedSnapshotPath);
				if (!string.IsNullOrWhiteSpace(snapshotFolder))
				{
					AddDistinctPath(roots, Path.Combine(snapshotFolder, "FingerprintDebug"));
					string parentFolder2 = Path.GetDirectoryName(snapshotFolder);
					if (!string.IsNullOrWhiteSpace(parentFolder2))
					{
						AddDistinctPath(roots, Path.Combine(parentFolder2, "FingerprintDebug"));
					}
				}
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		return roots;
	}

	private static bool SignaturePathNeedsRecovery(string signaturePath)
	{
		bool SignaturePathNeedsRecovery;
		if (string.IsNullOrWhiteSpace(signaturePath))
		{
			SignaturePathNeedsRecovery = true;
		}
		else
		{
			try
			{
				SignaturePathNeedsRecovery = !File.Exists(Environment.ExpandEnvironmentVariables(signaturePath.Trim()));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				SignaturePathNeedsRecovery = true;
				ProjectData.ClearProjectError();
			}
		}
		return SignaturePathNeedsRecovery;
	}

	private static void AddDistinctPath(List<string> paths, string path)
	{
		if (paths != null && !string.IsNullOrWhiteSpace(path))
		{
			string text = path.Trim();
			if (!paths.Any([SpecialName] (string x) => string.Equals(x, text, StringComparison.OrdinalIgnoreCase)))
			{
				paths.Add(text);
			}
		}
	}

	private static List<LoadableFamilyComparisonItem> BuildLoadableFamilyComparisons(StandardLibrarySnapshot standardSnapshot, ProjectContentSnapshot projectSnapshot, ProjectTrackingCatalog trackingCatalog)
	{
		List<LoadableFamilyComparisonItem> results = new List<LoadableFamilyComparisonItem>();
		List<StandardLoadableFamilySnapshotItem> standardLoadableFamilies = (from x in standardSnapshot.LoadableFamilies
			where !FamilyBrowserFamilyClassificationService.IsTypeManagedFamilyLike(x.CategoryName, x.CategoryId, x.FamilyName)
			select x).ToList();
		List<ProjectLoadableFamilySnapshotItem> projectLoadableFamilies = (from x in projectSnapshot.LoadableFamilies
			where !FamilyBrowserFamilyClassificationService.IsTypeManagedFamilyLike(x.CategoryName, x.CategoryId, x.FamilyName)
			select x).ToList();
		Dictionary<string, ProjectLoadableFamilySnapshotItem> projectMap = BuildFirstMap(projectLoadableFamilies, [SpecialName] (ProjectLoadableFamilySnapshotItem x) => BuildLoadableMatchKey(x));
		Dictionary<string, List<ProjectLoadableFamilySnapshotItem>> projectNameMap = BuildGroupedMap(projectLoadableFamilies, [SpecialName] (ProjectLoadableFamilySnapshotItem x) => Normalize(x.FamilyName));
		Dictionary<string, List<StandardLoadableFamilySnapshotItem>> standardNameMap = BuildGroupedMap(standardLoadableFamilies, [SpecialName] (StandardLoadableFamilySnapshotItem x) => Normalize(x.FamilyName));
		HashSet<string> matchedProjectKeys = new HashSet<string>(StringComparer.Ordinal);
		Dictionary<string, TrackedLoadableFamilyState> trackingMap = BuildTrackedLoadableMap(trackingCatalog);
		foreach (StandardLoadableFamilySnapshotItem standardFamily in standardLoadableFamilies.OrderBy<StandardLoadableFamilySnapshotItem, string>([SpecialName] (StandardLoadableFamilySnapshotItem x) => BuildLoadableMatchKey(x), StringComparer.Ordinal))
		{
			string key = BuildLoadableIdentityKey(standardFamily);
			string matchKey = BuildLoadableMatchKey(standardFamily);
			ProjectLoadableFamilySnapshotItem projectFamily = null;
			projectMap.TryGetValue(matchKey, out projectFamily);
			if (projectFamily == null)
			{
				projectFamily = FindProjectLoadableCategoryMatch(standardFamily, projectNameMap);
			}
			if (projectFamily != null)
			{
				matchedProjectKeys.Add(BuildLoadableMatchKey(projectFamily));
				List<string> missingTypes = standardFamily.TypeNames.Except<string>(projectFamily.TypeNames, StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
				List<string> extraTypes = projectFamily.TypeNames.Except<string>(standardFamily.TypeNames, StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
				AppendSignatureInventoryTypeDifferences(standardFamily, projectFamily, missingTypes, extraTypes);
				string standardFingerprint = ProjectSnapshotFingerprintService.BuildLoadableFingerprint(standardFamily);
				string comparisonProjectFingerprint = ProjectSnapshotFingerprintService.BuildLoadableFingerprint(projectFamily);
				if (LoadableFamiliesEquivalentForComparison(standardFamily, projectFamily, missingTypes, extraTypes))
				{
					comparisonProjectFingerprint = standardFingerprint;
				}
				TrackedLoadableFamilyState trackedState = null;
				trackingMap.TryGetValue(key, out trackedState);
				ComparisonClassificationResult classification = (IsProjectLoadableContentFingerprintMissing(standardFamily, projectFamily) ? new ComparisonClassificationResult
				{
					Status = "ManualReview",
					Notes = (string.IsNullOrWhiteSpace(projectFamily.ContentFingerprintFailureReason) ? "Project family content fingerprint was not created. Re-run the current model precise check and inspect the project signature diagnostics before comparing or applying this family." : projectFamily.ContentFingerprintFailureReason)
				} : ClassifyTrackedStatus(trackedState?.ApprovedFingerprint, trackedState?.ApprovedStandardStamp, standardFingerprint, standardSnapshot.CapturedAtUtc, comparisonProjectFingerprint));
				results.Add(new LoadableFamilyComparisonItem
				{
					IdentityKey = key,
					FamilyName = standardFamily.FamilyName,
					CategoryName = standardFamily.CategoryName,
					Status = classification.Status,
					StandardFingerprint = standardFingerprint,
					ProjectFingerprint = comparisonProjectFingerprint,
					StandardContentFingerprint = (standardFamily.ContentFingerprint ?? string.Empty),
					ProjectContentFingerprint = (projectFamily.ContentFingerprint ?? string.Empty),
					StandardContentSignatureDebugPath = (standardFamily.ContentSignatureDebugPath ?? string.Empty),
					ProjectContentSignatureDebugPath = (projectFamily.ContentSignatureDebugPath ?? string.Empty),
					StandardContentFingerprintFailureReason = (standardFamily.ContentFingerprintFailureReason ?? string.Empty),
					ProjectContentFingerprintFailureReason = (projectFamily.ContentFingerprintFailureReason ?? string.Empty),
					FingerprintDifferenceSummary = BuildLoadableFingerprintDifferenceSummary(standardFamily, projectFamily, standardFingerprint, comparisonProjectFingerprint),
					FingerprintDifferenceDetails = BuildLoadableFingerprintDifferenceDetails(standardFamily, projectFamily, standardFingerprint, comparisonProjectFingerprint, missingTypes, extraTypes),
					StandardTypeCount = standardFamily.TypeCount,
					ProjectTypeCount = projectFamily.TypeCount,
					ProjectInstanceCount = projectFamily.InstanceCount,
					MissingTypeNames = missingTypes,
					ExtraTypeNames = extraTypes,
					ProjectTypeNames = CloneStringList(projectFamily.TypeNames),
					ProjectParameters = CloneParameterSnapshotItems(projectFamily.Parameters),
					ProjectNestedLoadableFamilies = BuildNestedLoadableFamiliesFromSignature(projectFamily.ContentSignatureDebugPath, projectFamily.FamilyName),
					Notes = classification.Notes,
					IsNestedLoadableChild = IsActualNestedLoadableChild(standardFamily),
					NestedLoadableFamilies = CloneNestedLoadableFamilies(standardFamily.NestedLoadableFamilies)
				});
				continue;
			}
			List<ProjectLoadableFamilySnapshotItem> categoryMismatches = FindProjectLoadableCategoryMismatches(standardFamily, projectNameMap);
			if (categoryMismatches.Count > 0)
			{
				foreach (ProjectLoadableFamilySnapshotItem mismatch in categoryMismatches)
				{
					matchedProjectKeys.Add(BuildLoadableMatchKey(mismatch));
				}
				results.Add(new LoadableFamilyComparisonItem
				{
					IdentityKey = key,
					FamilyName = standardFamily.FamilyName,
					CategoryName = standardFamily.CategoryName,
					Status = "CategoryMismatch",
					StandardFingerprint = ProjectSnapshotFingerprintService.BuildLoadableFingerprint(standardFamily),
					StandardContentFingerprint = (standardFamily.ContentFingerprint ?? string.Empty),
					StandardContentSignatureDebugPath = (standardFamily.ContentSignatureDebugPath ?? string.Empty),
					StandardTypeCount = standardFamily.TypeCount,
					ProjectTypeCount = categoryMismatches.Sum([SpecialName] (ProjectLoadableFamilySnapshotItem x) => x.TypeCount),
					ProjectInstanceCount = categoryMismatches.Sum([SpecialName] (ProjectLoadableFamilySnapshotItem x) => x.InstanceCount),
					ExtraTypeNames = categoryMismatches.SelectMany([SpecialName] (ProjectLoadableFamilySnapshotItem x) => x.TypeNames).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase)
						.ToList(),
					ProjectTypeNames = categoryMismatches.SelectMany([SpecialName] (ProjectLoadableFamilySnapshotItem x) => x.TypeNames ?? new List<string>()).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase)
						.ToList(),
					ProjectParameters = categoryMismatches.SelectMany([SpecialName] (ProjectLoadableFamilySnapshotItem x) => x.Parameters ?? new List<StandardFamilyParameterSnapshotItem>()).ToList(),
					ProjectNestedLoadableFamilies = categoryMismatches.SelectMany([SpecialName] (ProjectLoadableFamilySnapshotItem x) => BuildNestedLoadableFamiliesFromSignature(x.ContentSignatureDebugPath, x.FamilyName)).ToList(),
					Notes = BuildLoadableCategoryMismatchNote(standardFamily, categoryMismatches),
					IsNestedLoadableChild = IsActualNestedLoadableChild(standardFamily),
					NestedLoadableFamilies = CloneNestedLoadableFamilies(standardFamily.NestedLoadableFamilies)
				});
			}
			else
			{
				results.Add(new LoadableFamilyComparisonItem
				{
					IdentityKey = key,
					FamilyName = standardFamily.FamilyName,
					CategoryName = standardFamily.CategoryName,
					Status = "LoadAvailable",
					StandardFingerprint = ProjectSnapshotFingerprintService.BuildLoadableFingerprint(standardFamily),
					StandardContentFingerprint = (standardFamily.ContentFingerprint ?? string.Empty),
					StandardContentSignatureDebugPath = (standardFamily.ContentSignatureDebugPath ?? string.Empty),
					StandardTypeCount = standardFamily.TypeCount,
					ProjectTypeCount = 0,
					MissingTypeNames = standardFamily.TypeNames.OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList(),
					Notes = "Not loaded in current project.",
					IsNestedLoadableChild = IsActualNestedLoadableChild(standardFamily),
					NestedLoadableFamilies = CloneNestedLoadableFamilies(standardFamily.NestedLoadableFamilies)
				});
			}
		}
		foreach (ProjectLoadableFamilySnapshotItem projectFamily2 in projectLoadableFamilies.OrderBy<ProjectLoadableFamilySnapshotItem, string>([SpecialName] (ProjectLoadableFamilySnapshotItem x) => BuildLoadableMatchKey(x), StringComparer.Ordinal))
		{
			string key2 = BuildLoadableIdentityKey(projectFamily2);
			string matchKey2 = BuildLoadableMatchKey(projectFamily2);
			if (matchedProjectKeys.Contains(matchKey2))
			{
				continue;
			}
			List<StandardLoadableFamilySnapshotItem> standardCategoryMismatches = FindStandardLoadableCategoryMismatches(projectFamily2, standardNameMap);
			if (standardCategoryMismatches.Count > 0)
			{
				results.Add(new LoadableFamilyComparisonItem
				{
					IdentityKey = key2,
					FamilyName = projectFamily2.FamilyName,
					CategoryName = projectFamily2.CategoryName,
					Status = "CategoryMismatch",
					ProjectFingerprint = ProjectSnapshotFingerprintService.BuildLoadableFingerprint(projectFamily2),
					ProjectContentFingerprint = (projectFamily2.ContentFingerprint ?? string.Empty),
					ProjectContentSignatureDebugPath = (projectFamily2.ContentSignatureDebugPath ?? string.Empty),
					StandardTypeCount = standardCategoryMismatches.Sum([SpecialName] (StandardLoadableFamilySnapshotItem x) => x.TypeCount),
					ProjectTypeCount = projectFamily2.TypeCount,
					ProjectInstanceCount = projectFamily2.InstanceCount,
					ExtraTypeNames = projectFamily2.TypeNames.OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList(),
					ProjectTypeNames = CloneStringList(projectFamily2.TypeNames),
					ProjectParameters = CloneParameterSnapshotItems(projectFamily2.Parameters),
					ProjectNestedLoadableFamilies = BuildNestedLoadableFamiliesFromSignature(projectFamily2.ContentSignatureDebugPath, projectFamily2.FamilyName),
					Notes = BuildProjectLoadableCategoryMismatchNote(projectFamily2, standardCategoryMismatches)
				});
			}
			else
			{
				results.Add(new LoadableFamilyComparisonItem
				{
					IdentityKey = key2,
					FamilyName = projectFamily2.FamilyName,
					CategoryName = projectFamily2.CategoryName,
					Status = "ProjectOnly",
					ProjectFingerprint = ProjectSnapshotFingerprintService.BuildLoadableFingerprint(projectFamily2),
					ProjectContentFingerprint = (projectFamily2.ContentFingerprint ?? string.Empty),
					ProjectContentSignatureDebugPath = (projectFamily2.ContentSignatureDebugPath ?? string.Empty),
					StandardTypeCount = 0,
					ProjectTypeCount = projectFamily2.TypeCount,
					ProjectInstanceCount = projectFamily2.InstanceCount,
					ExtraTypeNames = projectFamily2.TypeNames.OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList(),
					ProjectTypeNames = CloneStringList(projectFamily2.TypeNames),
					ProjectParameters = CloneParameterSnapshotItems(projectFamily2.Parameters),
					ProjectNestedLoadableFamilies = BuildNestedLoadableFamiliesFromSignature(projectFamily2.ContentSignatureDebugPath, projectFamily2.FamilyName),
					Notes = "Project family was not found in the registered standard snapshot."
				});
			}
		}
		NestedLoadableFamilyDifferencePropagationService.Apply(results);
		return results;
	}

	private static bool LoadableFamiliesEquivalentForComparison(StandardLoadableFamilySnapshotItem standardFamily, ProjectLoadableFamilySnapshotItem projectFamily, IEnumerable<string> missingTypes, IEnumerable<string> extraTypes)
	{
		if (standardFamily == null || projectFamily == null)
		{
			return false;
		}
		if (!string.Equals(Normalize(standardFamily.CategoryName), Normalize(projectFamily.CategoryName), StringComparison.Ordinal))
		{
			return false;
		}
		if (!string.Equals(Normalize(standardFamily.CategoryGroup), Normalize(projectFamily.CategoryGroup), StringComparison.Ordinal))
		{
			return false;
		}
		if (!string.Equals(Normalize(standardFamily.FamilyName), Normalize(projectFamily.FamilyName), StringComparison.Ordinal))
		{
			return false;
		}
		if (standardFamily.TypeCount != projectFamily.TypeCount)
		{
			return false;
		}
		if ((missingTypes ?? Enumerable.Empty<string>()).Any([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)) || (extraTypes ?? Enumerable.Empty<string>()).Any([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)))
		{
			return false;
		}
		if (standardFamily.IsShared != projectFamily.IsShared)
		{
			return false;
		}
		if (!string.Equals(BuildParameterDifferenceSignature(standardFamily.Parameters), BuildParameterDifferenceSignature(projectFamily.Parameters), StringComparison.Ordinal))
		{
			return false;
		}
		string standardContentFingerprint = Normalize(standardFamily.ContentFingerprint);
		string projectContentFingerprint = Normalize(projectFamily.ContentFingerprint);
		if (standardContentFingerprint.Length == 0 && projectContentFingerprint.Length == 0)
		{
			return true;
		}
		if (string.Equals(standardContentFingerprint, projectContentFingerprint, StringComparison.Ordinal))
		{
			return true;
		}
		if (IsProjectLoadableContentFingerprintMissing(standardFamily, projectFamily))
		{
			return false;
		}
		return ContentSignatureSourcesHaveSameComparableLines(standardFamily.ContentSignatureDebugPath, projectFamily.ContentSignatureDebugPath);
	}

	private static List<string> BuildLoadableFingerprintDifferenceSummary(StandardLoadableFamilySnapshotItem standardFamily, ProjectLoadableFamilySnapshotItem projectFamily, string standardFingerprint, string projectFingerprint)
	{
		List<string> result = new List<string>();
		if (standardFamily == null || projectFamily == null)
		{
			return result;
		}
		if (string.Equals(standardFingerprint ?? string.Empty, projectFingerprint ?? string.Empty, StringComparison.OrdinalIgnoreCase))
		{
			return result;
		}
		bool projectContentFingerprintMissing = IsProjectLoadableContentFingerprintMissing(standardFamily, projectFamily);
		if (projectContentFingerprintMissing)
		{
			result.Add(string.IsNullOrWhiteSpace(projectFamily.ContentFingerprintFailureReason) ? "Project fingerprint missing." : ("Project fingerprint missing: " + projectFamily.ContentFingerprintFailureReason));
		}
		if (!string.Equals(Normalize(standardFamily.CategoryName), Normalize(projectFamily.CategoryName), StringComparison.Ordinal))
		{
			result.Add("Category differs.");
		}
		if (!string.Equals(Normalize(standardFamily.CategoryGroup), Normalize(projectFamily.CategoryGroup), StringComparison.Ordinal))
		{
			result.Add("Category group differs.");
		}
		if (!string.Equals(Normalize(standardFamily.FamilyName), Normalize(projectFamily.FamilyName), StringComparison.Ordinal))
		{
			result.Add("Family name differs.");
		}
		if (standardFamily.TypeCount != projectFamily.TypeCount)
		{
			result.Add("Type count differs: standard=" + standardFamily.TypeCount.ToString(CultureInfo.InvariantCulture) + ", project=" + projectFamily.TypeCount.ToString(CultureInfo.InvariantCulture) + ".");
		}
		List<string> missingTypes = (standardFamily.TypeNames ?? new List<string>()).Except<string>(projectFamily.TypeNames ?? new List<string>(), StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
		List<string> extraTypes = (projectFamily.TypeNames ?? new List<string>()).Except<string>(standardFamily.TypeNames ?? new List<string>(), StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
		AppendSignatureInventoryTypeDifferences(standardFamily, projectFamily, missingTypes, extraTypes);
		if (missingTypes.Count > 0)
		{
			result.Add("Missing project types: " + string.Join(", ", missingTypes.Take(12)) + ((missingTypes.Count > 12) ? " ..." : string.Empty));
		}
		if (extraTypes.Count > 0)
		{
			result.Add("Extra project types: " + string.Join(", ", extraTypes.Take(12)) + ((extraTypes.Count > 12) ? " ..." : string.Empty));
		}
		if (standardFamily.IsShared != projectFamily.IsShared)
		{
			result.Add("Shared flag differs: standard=" + standardFamily.IsShared + ", project=" + projectFamily.IsShared + ".");
		}
		if (!string.Equals(standardFamily.ContentFingerprint ?? string.Empty, projectFamily.ContentFingerprint ?? string.Empty, StringComparison.OrdinalIgnoreCase) && !projectContentFingerprintMissing && !ContentSignatureSourcesHaveSameComparableLines(standardFamily.ContentSignatureDebugPath, projectFamily.ContentSignatureDebugPath))
		{
			if (missingTypes.Count > 0 || extraTypes.Count > 0)
			{
				result.Add("Type inventory differs.");
			}
			else
			{
				List<string> signatureDiffs = BuildContentSignatureDifferenceSummary(standardFamily.ContentSignatureDebugPath, projectFamily.ContentSignatureDebugPath);
				if (signatureDiffs.Count > 0)
				{
					result.AddRange(signatureDiffs);
				}
				else
				{
					result.Add("Fingerprint differs.");
				}
			}
		}
		if (!string.IsNullOrWhiteSpace(standardFamily.ContentSignatureDebugPath) && string.IsNullOrWhiteSpace(projectFamily.ContentSignatureDebugPath))
		{
			result.Add("Project signature diagnostics missing.");
		}
		string a = BuildParameterDifferenceSignature(standardFamily.Parameters);
		string projectParameterSignature = BuildParameterDifferenceSignature(projectFamily.Parameters);
		if (!string.Equals(a, projectParameterSignature, StringComparison.Ordinal))
		{
			result.Add("Parameter differs.");
		}
		if (result.Count == 0)
		{
			List<string> signatureDiffs2 = BuildContentSignatureDifferenceSummary(standardFamily.ContentSignatureDebugPath, projectFamily.ContentSignatureDebugPath);
			if (signatureDiffs2.Count > 0)
			{
				result.AddRange(signatureDiffs2);
			}
			else
			{
				result.Add("Fingerprint differs.");
			}
		}
		return result;
	}

	private static List<LoadableFingerprintDifferenceDetailItem> BuildLoadableFingerprintDifferenceDetails(StandardLoadableFamilySnapshotItem standardFamily, ProjectLoadableFamilySnapshotItem projectFamily, string standardFingerprint, string projectFingerprint, IEnumerable<string> missingTypes, IEnumerable<string> extraTypes)
	{
		List<LoadableFingerprintDifferenceDetailItem> result = new List<LoadableFingerprintDifferenceDetailItem>();
		if (standardFamily == null || projectFamily == null)
		{
			return result;
		}
		if (string.Equals(standardFingerprint ?? string.Empty, projectFingerprint ?? string.Empty, StringComparison.OrdinalIgnoreCase))
		{
			return result;
		}
		if (!string.Equals(Normalize(standardFamily.CategoryName), Normalize(projectFamily.CategoryName), StringComparison.Ordinal))
		{
			AddDifferenceDetail(result, "category", "different", standardFamily.CategoryName, projectFamily.CategoryName, string.Empty);
		}
		if (!string.Equals(Normalize(standardFamily.CategoryGroup), Normalize(projectFamily.CategoryGroup), StringComparison.Ordinal))
		{
			AddDifferenceDetail(result, "category group", "different", standardFamily.CategoryGroup, projectFamily.CategoryGroup, string.Empty);
		}
		if (!string.Equals(Normalize(standardFamily.FamilyName), Normalize(projectFamily.FamilyName), StringComparison.Ordinal))
		{
			AddDifferenceDetail(result, "family name", "different", standardFamily.FamilyName, projectFamily.FamilyName, string.Empty);
		}
		if (standardFamily.TypeCount != projectFamily.TypeCount)
		{
			AddDifferenceDetail(result, "family types", "count", standardFamily.TypeCount.ToString(CultureInfo.InvariantCulture), projectFamily.TypeCount.ToString(CultureInfo.InvariantCulture), "Type count differs.");
		}
		List<string> missingList = (missingTypes ?? Enumerable.Empty<string>()).Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase)
			.ToList();
		List<string> extraList = (extraTypes ?? Enumerable.Empty<string>()).Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase)
			.ToList();
		if (missingList.Count > 0 || extraList.Count > 0)
		{
			AddDifferenceDetail(result, "family types", "inventory", LimitListDisplay(missingList, 6), LimitListDisplay(extraList, 6), "Standard-only types are missing in the project; project-only types are extra.");
		}
		if (standardFamily.IsShared != projectFamily.IsShared)
		{
			AddDifferenceDetail(result, "shared flag", "different", standardFamily.IsShared.ToString(), projectFamily.IsShared.ToString(), string.Empty);
		}
		if (!string.Equals(BuildParameterDifferenceSignature(standardFamily.Parameters), BuildParameterDifferenceSignature(projectFamily.Parameters), StringComparison.Ordinal))
		{
			AppendDifferenceDetails(result, BuildParameterDifferenceDetails(standardFamily.Parameters, projectFamily.Parameters), 10);
		}
		if (IsProjectLoadableContentFingerprintMissing(standardFamily, projectFamily))
		{
			AddDifferenceDetail(result, "content fingerprint", "missing", ShortHash(standardFamily.ContentFingerprint), string.IsNullOrWhiteSpace(projectFamily.ContentFingerprintFailureReason) ? "(missing)" : projectFamily.ContentFingerprintFailureReason, "Project family content fingerprint was not created.");
		}
		else if (!string.Equals(standardFamily.ContentFingerprint ?? string.Empty, projectFamily.ContentFingerprint ?? string.Empty, StringComparison.OrdinalIgnoreCase) && !ContentSignatureSourcesHaveSameComparableLines(standardFamily.ContentSignatureDebugPath, projectFamily.ContentSignatureDebugPath))
		{
			AppendDifferenceDetails(result, BuildContentSignatureDifferenceDetails(standardFamily.ContentSignatureDebugPath, projectFamily.ContentSignatureDebugPath, standardFamily.ContentFingerprint, projectFamily.ContentFingerprint), 12);
		}
		if (result.Count == 0)
		{
			AddDifferenceDetail(result, "fingerprint", "different", ShortHash(standardFingerprint), ShortHash(projectFingerprint), "Fingerprint differs, but no concise field-level difference was identified. Inspect signature diagnostics.");
		}
		return result.Take(14).ToList();
	}

	private static List<LoadableFingerprintDifferenceDetailItem> BuildParameterDifferenceDetails(IEnumerable<StandardFamilyParameterSnapshotItem> standardParameters, IEnumerable<StandardFamilyParameterSnapshotItem> projectParameters)
	{
		List<LoadableFingerprintDifferenceDetailItem> result = new List<LoadableFingerprintDifferenceDetailItem>();
		Dictionary<string, List<StandardFamilyParameterSnapshotItem>> standardGroups = BuildParameterDifferenceGroups(standardParameters);
		Dictionary<string, List<StandardFamilyParameterSnapshotItem>> projectGroups = BuildParameterDifferenceGroups(projectParameters);
		List<string> allKeys = standardGroups.Keys.Union<string>(projectGroups.Keys, StringComparer.Ordinal).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal).ToList();
		foreach (string key in allKeys)
		{
			List<StandardFamilyParameterSnapshotItem> standardGroup = null;
			List<StandardFamilyParameterSnapshotItem> projectGroup = null;
			standardGroups.TryGetValue(key, out standardGroup);
			projectGroups.TryGetValue(key, out projectGroup);
			if (standardGroup == null)
			{
				AddDifferenceDetail(result, "parameters/formulas", "project-only", "-", FormatParameterGroupForDiff(projectGroup), "Parameter exists only in the project scan.");
			}
			else if (projectGroup == null)
			{
				AddDifferenceDetail(result, "parameters/formulas", "standard-only", FormatParameterGroupForDiff(standardGroup), "-", "Parameter exists only in the standard scan.");
			}
			else
			{
				string a = BuildParameterGroupComparableSignature(standardGroup);
				string projectSignature = BuildParameterGroupComparableSignature(projectGroup);
				if (!string.Equals(a, projectSignature, StringComparison.Ordinal))
				{
					AddDifferenceDetail(result, "parameters/formulas", "modified", FormatParameterGroupForDiff(standardGroup), FormatParameterGroupForDiff(projectGroup), "Parameter definition, shared flag, GUID, or formula differs.");
				}
			}
			if (result.Count >= 8)
			{
				break;
			}
		}
		if (result.Count == 0)
		{
			AddDifferenceDetail(result, "parameters/formulas", "different", "parameter signature differs", "parameter signature differs", "Parameter signature differed, but no compact parameter row could be isolated.");
		}
		return result;
	}

	private static Dictionary<string, List<StandardFamilyParameterSnapshotItem>> BuildParameterDifferenceGroups(IEnumerable<StandardFamilyParameterSnapshotItem> parameters)
	{
		Dictionary<string, List<StandardFamilyParameterSnapshotItem>> result = new Dictionary<string, List<StandardFamilyParameterSnapshotItem>>(StringComparer.Ordinal);
		if (parameters == null)
		{
			return result;
		}
		foreach (StandardFamilyParameterSnapshotItem parameter in FamilyParameterSnapshotNormalizationService.DeduplicateDefinitions(parameters))
		{
			if (parameter == null || string.IsNullOrWhiteSpace(parameter.Name))
			{
				continue;
			}
			string key = BuildParameterDifferenceKey(parameter);
			if (!string.IsNullOrWhiteSpace(key))
			{
				List<StandardFamilyParameterSnapshotItem> group = null;
				if (!result.TryGetValue(key, out group))
				{
					group = (result[key] = new List<StandardFamilyParameterSnapshotItem>());
				}
				group.Add(parameter);
			}
		}
		return result;
	}

	private static string BuildParameterDifferenceKey(StandardFamilyParameterSnapshotItem parameter)
	{
		return FamilyParameterSnapshotNormalizationService.BuildDefinitionIdentityKey(parameter);
	}

	private static string BuildParameterGroupComparableSignature(IEnumerable<StandardFamilyParameterSnapshotItem> parameters)
	{
		return FamilyParameterSnapshotNormalizationService.BuildComparableDefinitionSignature(parameters);
	}

	private static string FormatParameterGroupForDiff(IEnumerable<StandardFamilyParameterSnapshotItem> parameters)
	{
		if (parameters == null)
		{
			return "-";
		}
		List<StandardFamilyParameterSnapshotItem> items = parameters.Where([SpecialName] (StandardFamilyParameterSnapshotItem x) => x != null).OrderBy<StandardFamilyParameterSnapshotItem, string>([SpecialName] (StandardFamilyParameterSnapshotItem x) => Normalize(x.TypeName), StringComparer.Ordinal).ThenBy<StandardFamilyParameterSnapshotItem, string>([SpecialName] (StandardFamilyParameterSnapshotItem x) => Normalize(x.Name), StringComparer.Ordinal)
			.ToList();
		if (items.Count == 0)
		{
			return "-";
		}
		StandardFamilyParameterSnapshotItem first = items[0];
		List<string> parts = new List<string>
		{
			string.IsNullOrWhiteSpace(first.Name) ? "-" : first.Name,
			string.IsNullOrWhiteSpace(first.Scope) ? "-" : first.Scope,
			first.IsInstance ? "instance" : "type/family",
			first.IsShared ? "shared" : "family",
			string.IsNullOrWhiteSpace(first.StorageType) ? "-" : first.StorageType
		};
		List<string> identities = (from x in items
			select ResolvePortableParameterIdentity(x) into x
			where !string.IsNullOrWhiteSpace(x)
			select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).Take(2).ToList();
		if (identities.Count > 0)
		{
			parts.Add(string.Join(", ", identities));
		}
		List<string> formulas = (from x in items
			select (x.Formula ?? string.Empty).Trim() into x
			where !string.IsNullOrWhiteSpace(x)
			select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).Take(3).ToList();
		if (formulas.Count > 0)
		{
			parts.Add("formula=" + string.Join("; ", formulas));
		}
		if (items.Count > 1)
		{
			parts.Add("rows=" + items.Count.ToString(CultureInfo.InvariantCulture));
		}
		return ShortDiffCell(string.Join(" / ", parts), 260);
	}

	private static List<LoadableFingerprintDifferenceDetailItem> BuildContentSignatureDifferenceDetails(string standardSignaturePath, string projectSignaturePath, string standardContentFingerprint, string projectContentFingerprint)
	{
		List<LoadableFingerprintDifferenceDetailItem> result = new List<LoadableFingerprintDifferenceDetailItem>();
		List<string> standardLines = ReadSignatureSourceLines(standardSignaturePath);
		List<string> projectLines = ReadSignatureSourceLines(projectSignaturePath);
		Dictionary<string, List<string>> standardDebugIndex = ReadSignatureElementDebugIndex(standardSignaturePath);
		Dictionary<string, List<string>> projectDebugIndex = ReadSignatureElementDebugIndex(projectSignaturePath);
		if (standardLines.Count == 0 || projectLines.Count == 0)
		{
			AddDifferenceDetail(result, "signature diagnostics", "missing", (standardLines.Count == 0) ? (standardSignaturePath ?? string.Empty) : "readable", (projectLines.Count == 0) ? (projectSignaturePath ?? string.Empty) : "readable", "One or both signature detail files could not be read.");
			return result;
		}
		SignatureLineDelta signatureLineDelta = BuildSignatureLineDelta(standardLines, projectLines);
		List<string> standardOnly = signatureLineDelta.StandardOnly;
		List<string> projectOnly = signatureLineDelta.ProjectOnly;
		if (standardOnly.Count == 0 && projectOnly.Count == 0)
		{
			return result;
		}
		Dictionary<string, List<string>> standardGroups = standardOnly.GroupBy<string, string>([SpecialName] (string x) => DescribeSignatureLine(x), StringComparer.Ordinal).ToDictionary<IGrouping<string, string>, string, List<string>>([SpecialName] (IGrouping<string, string> x) => x.Key, [SpecialName] (IGrouping<string, string> x) => x.ToList(), StringComparer.Ordinal);
		Dictionary<string, List<string>> projectGroups = projectOnly.GroupBy<string, string>([SpecialName] (string x) => DescribeSignatureLine(x), StringComparer.Ordinal).ToDictionary<IGrouping<string, string>, string, List<string>>([SpecialName] (IGrouping<string, string> x) => x.Key, [SpecialName] (IGrouping<string, string> x) => x.ToList(), StringComparer.Ordinal);
		List<string> groupNames = (from x in standardGroups.Keys.Union<string>(projectGroups.Keys, StringComparer.Ordinal)
			where !IsFamilyTypesSignatureGroup(x)
			orderby SignatureDifferenceGroupOrder(x)
			select x).ThenBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal).ToList();
		foreach (string groupName in groupNames.Take(8))
		{
			List<string> standardGroup = null;
			List<string> projectGroup = null;
			standardGroups.TryGetValue(groupName, out standardGroup);
			projectGroups.TryGetValue(groupName, out projectGroup);
			if (string.Equals(groupName, "nested labels", StringComparison.Ordinal))
			{
				List<LoadableFingerprintDifferenceDetailItem> nestedLabelDetails = BuildNestedLabelDifferenceDetails(standardGroup, projectGroup);
				if (nestedLabelDetails.Count > 0)
				{
					AppendDifferenceDetails(result, nestedLabelDetails, 14);
					continue;
				}
			}
			if (string.Equals(groupName, "lookup tables", StringComparison.Ordinal))
			{
				List<LoadableFingerprintDifferenceDetailItem> lookupCsvDetails = BuildLookupCsvDifferenceDetails(standardGroup, projectGroup);
				if (lookupCsvDetails.Count > 0)
				{
					AppendDifferenceDetails(result, lookupCsvDetails, 14);
					continue;
				}
			}
			bool isMaterialGroup = IsMaterialSignatureGroup(groupName);
			if (ShouldBuildPairedSignatureDifferenceDetails(groupName))
			{
				List<LoadableFingerprintDifferenceDetailItem> pairedDetails = BuildPairedSignatureDifferenceDetails(groupName, standardGroup, projectGroup, standardDebugIndex, projectDebugIndex);
				if (pairedDetails.Count > 0)
				{
					AppendDifferenceDetails(result, pairedDetails, 14);
					continue;
				}
			}
			string standardValue = FormatSignatureGroupForDiff(standardGroup, isMaterialGroup ? null : standardDebugIndex);
			string projectValue = FormatSignatureGroupForDiff(projectGroup, isMaterialGroup ? null : projectDebugIndex);
			int standardCount = standardGroup?.Count ?? 0;
			int projectCount = projectGroup?.Count ?? 0;
			string standardIds = (isMaterialGroup ? string.Empty : FormatSignatureDebugIdentities(standardGroup, standardDebugIndex));
			string projectIds = (isMaterialGroup ? string.Empty : FormatSignatureDebugIdentities(projectGroup, projectDebugIndex));
			string details = "standard-only " + standardCount.ToString(CultureInfo.InvariantCulture) + ", project-only " + projectCount.ToString(CultureInfo.InvariantCulture);
			if (!string.IsNullOrWhiteSpace(standardIds))
			{
				details = details + ", standard ids " + standardIds;
			}
			if (!string.IsNullOrWhiteSpace(projectIds))
			{
				details = details + ", project ids " + projectIds;
			}
			AddDifferenceDetail(result, groupName, "signature-source", standardValue, projectValue, details);
		}
		if (groupNames.Count > 8)
		{
			AddDifferenceDetail(result, "additional groups", "omitted", "-", "-", checked(groupNames.Count - 8).ToString(CultureInfo.InvariantCulture) + " additional signature difference groups omitted.");
		}
		return result;
	}

	private static List<LoadableFingerprintDifferenceDetailItem> BuildLookupCsvDifferenceDetails(IList<string> standardGroup, IList<string> projectGroup)
	{
		List<LoadableFingerprintDifferenceDetailItem> result = new List<LoadableFingerprintDifferenceDetailItem>();
		Dictionary<string, LookupCsvSignatureTableInfo> standardTables = BuildLookupCsvSignatureTableMap(standardGroup);
		Dictionary<string, LookupCsvSignatureTableInfo> projectTables = BuildLookupCsvSignatureTableMap(projectGroup);
		List<string> keys = standardTables.Keys.Union(projectTables.Keys, StringComparer.Ordinal).OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal).ToList();
		foreach (string key in keys.Take(10))
		{
			LookupCsvSignatureTableInfo standardTable = null;
			LookupCsvSignatureTableInfo projectTable = null;
			standardTables.TryGetValue(key, out standardTable);
			projectTables.TryGetValue(key, out projectTable);
			if (standardTable == null)
			{
				AddDifferenceDetail(result, "lookup csv", "project-only", "CSV: no", FormatLookupCsvTableForDiff(projectTable), "Lookup CSV table exists only in the current project scan.");
			}
			else if (projectTable == null)
			{
				AddDifferenceDetail(result, "lookup csv", "standard-only", FormatLookupCsvTableForDiff(standardTable), "CSV: no", "Lookup CSV table exists only in the standard scan.");
			}
			else if (!LookupCsvCountsMatch(standardTable, projectTable))
			{
				AddDifferenceDetail(result, "lookup csv", "modified", FormatLookupCsvTableForDiff(standardTable), FormatLookupCsvTableForDiff(projectTable), "Lookup CSV row/column count differs.");
			}
			else if (!string.Equals(standardTable.Signature ?? string.Empty, projectTable.Signature ?? string.Empty, StringComparison.Ordinal))
			{
				AddDifferenceDetail(result, "lookup csv", "modified", FormatLookupCsvTableForDiff(standardTable), FormatLookupCsvTableForDiff(projectTable), "Lookup CSV content differs.");
			}
			if (result.Count >= 10)
			{
				break;
			}
		}
		if (result.Count == 0 && !LookupCsvSignatureLineSetsMatch(standardGroup, projectGroup))
		{
			AddDifferenceDetail(result, "lookup csv", "signature-source", FormatLookupCsvGroupForDiff(standardGroup), FormatLookupCsvGroupForDiff(projectGroup), "Lookup CSV signature differs.");
		}
		return result;
	}

	private static bool LookupCsvSignatureLineSetsMatch(IEnumerable<string> standardGroup, IEnumerable<string> projectGroup)
	{
		List<string> standardLines = NormalizeLookupCsvSignatureLines(standardGroup);
		List<string> projectLines = NormalizeLookupCsvSignatureLines(projectGroup);
		return standardLines.SequenceEqual(projectLines, StringComparer.Ordinal);
	}

	private static List<string> NormalizeLookupCsvSignatureLines(IEnumerable<string> lines)
	{
		if (lines == null)
		{
			return new List<string>();
		}
		return lines.Select([SpecialName] (string x) => NormalizeSignatureLine(x)).Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)).OrderBy([SpecialName] (string x) => x, StringComparer.Ordinal).ToList();
	}

	private static Dictionary<string, LookupCsvSignatureTableInfo> BuildLookupCsvSignatureTableMap(IEnumerable<string> lines)
	{
		Dictionary<string, LookupCsvSignatureTableInfo> result = new Dictionary<string, LookupCsvSignatureTableInfo>(StringComparer.Ordinal);
		foreach (LookupCsvSignatureTableInfo table in ParseLookupCsvSignatureTables(lines))
		{
			if (table == null || string.IsNullOrWhiteSpace(table.Name))
			{
				continue;
			}
			result[Normalize(table.Name)] = table;
		}
		return result;
	}

	private static List<LookupCsvSignatureTableInfo> ParseLookupCsvSignatureTables(IEnumerable<string> lines)
	{
		List<LookupCsvSignatureTableInfo> result = new List<LookupCsvSignatureTableInfo>();
		if (lines == null)
		{
			return result;
		}
		foreach (string rawLine in lines)
		{
			string line = NormalizeSignatureLine(rawLine);
			if (string.IsNullOrWhiteSpace(line))
			{
				continue;
			}
			if (line.StartsWith("lookup-tables=", StringComparison.Ordinal))
			{
				line = line.Substring("lookup-tables=".Length).Trim();
				if (line.Length == 0)
				{
					continue;
				}
			}
			if (line.StartsWith("table=", StringComparison.Ordinal))
			{
				LookupCsvSignatureTableInfo table = ParseLookupCsvTableLine(line);
				if (table != null && !string.IsNullOrWhiteSpace(table.Name))
				{
					result.Add(table);
				}
			}
			else if (line.StartsWith("lookup-table-error=", StringComparison.Ordinal))
			{
				result.Add(new LookupCsvSignatureTableInfo
				{
					Name = "lookup-table-error",
					HasError = true,
					Signature = line
				});
			}
		}
		return result;
	}

	private static LookupCsvSignatureTableInfo ParseLookupCsvTableLine(string line)
	{
		string text = NormalizeSignatureLine(line);
		if (text.StartsWith("lookup-tables=", StringComparison.Ordinal))
		{
			text = text.Substring("lookup-tables=".Length).Trim();
		}
		if (!text.StartsWith("table=", StringComparison.Ordinal))
		{
			return null;
		}
		string body = text.Substring("table=".Length);
		string[] tokens = body.Split('|');
		if (tokens.Length == 0)
		{
			return null;
		}
		LookupCsvSignatureTableInfo result = new LookupCsvSignatureTableInfo
		{
			Name = (tokens[0] ?? string.Empty).Trim(),
			Signature = text
		};
		foreach (string rawToken in tokens.Skip(1))
		{
			string token = (rawToken ?? string.Empty).Trim();
			if (token.Length == 0)
			{
				continue;
			}
			if (string.Equals(token, "missing", StringComparison.OrdinalIgnoreCase))
			{
				result.Missing = true;
			}
			else if (token.StartsWith("error=", StringComparison.OrdinalIgnoreCase))
			{
				result.HasError = true;
			}
			else if (token.StartsWith("columns=", StringComparison.OrdinalIgnoreCase))
			{
				result.ColumnCount = CountLookupCsvColumns(token.Substring("columns=".Length));
				result.HasColumnCount = true;
			}
			else if (token.StartsWith("rows=", StringComparison.OrdinalIgnoreCase))
			{
				result.RowCount = CountLookupCsvRows(token.Substring("rows=".Length));
				result.HasRowCount = true;
			}
		}
		return result;
	}

	private static int CountLookupCsvColumns(string value)
	{
		if (string.IsNullOrWhiteSpace(value))
		{
			return 0;
		}
		return value.Split(';').Count([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x));
	}

	private static int CountLookupCsvRows(string value)
	{
		if (string.IsNullOrWhiteSpace(value))
		{
			return 0;
		}
		return value.Split(';').Count([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x));
	}

	private static bool LookupCsvCountsMatch(LookupCsvSignatureTableInfo standardTable, LookupCsvSignatureTableInfo projectTable)
	{
		if (standardTable == null || projectTable == null)
		{
			return standardTable == projectTable;
		}
		if (standardTable.HasRowCount != projectTable.HasRowCount || standardTable.HasColumnCount != projectTable.HasColumnCount)
		{
			return false;
		}
		return standardTable.RowCount == projectTable.RowCount && standardTable.ColumnCount == projectTable.ColumnCount && standardTable.Missing == projectTable.Missing && standardTable.HasError == projectTable.HasError;
	}

	private static string FormatLookupCsvTableForDiff(LookupCsvSignatureTableInfo table)
	{
		if (table == null)
		{
			return "CSV: no";
		}
		List<string> parts = new List<string> { "CSV: yes" };
		if (!string.IsNullOrWhiteSpace(table.Name))
		{
			parts.Add("table=" + table.Name);
		}
		if (table.Missing)
		{
			parts.Add("missing");
		}
		if (table.HasError)
		{
			parts.Add("error");
		}
		if (table.HasRowCount || table.HasColumnCount)
		{
			parts.Add(table.RowCount.ToString(CultureInfo.InvariantCulture) + " rows x " + table.ColumnCount.ToString(CultureInfo.InvariantCulture) + " columns");
		}
		return ShortDiffCell(string.Join(" / ", parts), 180);
	}

	private static string FormatLookupCsvGroupForDiff(IEnumerable<string> lines)
	{
		List<LookupCsvSignatureTableInfo> tables = ParseLookupCsvSignatureTables(lines);
		if (tables.Count == 0)
		{
			return "CSV: no";
		}
		return ShortDiffCell(string.Join("; ", tables.Take(4).Select([SpecialName] (LookupCsvSignatureTableInfo x) => FormatLookupCsvTableForDiff(x))), 220);
	}

	private static bool ShouldBuildPairedSignatureDifferenceDetails(string groupName)
	{
		string key = Normalize(groupName);
		if (string.Equals(key, "elements/material", StringComparison.Ordinal))
		{
			return false;
		}
		return key.StartsWith("elements/", StringComparison.Ordinal) || string.Equals(key, "elements", StringComparison.Ordinal) || string.Equals(key, "nested/loadable instances", StringComparison.Ordinal) || string.Equals(key, "nested/loadable types", StringComparison.Ordinal) || string.Equals(key, "nested/loadable families", StringComparison.Ordinal) || string.Equals(key, "connectors", StringComparison.Ordinal) || string.Equals(key, "geometry", StringComparison.Ordinal);
	}

	private static bool IsMaterialSignatureGroup(string groupName)
	{
		return string.Equals(Normalize(groupName), "elements/material", StringComparison.Ordinal);
	}

	private static List<LoadableFingerprintDifferenceDetailItem> BuildPairedSignatureDifferenceDetails(string groupName, IList<string> standardGroup, IList<string> projectGroup, Dictionary<string, List<string>> standardDebugIndex, Dictionary<string, List<string>> projectDebugIndex)
	{
		List<LoadableFingerprintDifferenceDetailItem> result = new List<LoadableFingerprintDifferenceDetailItem>();
		IList<string> standardLines = standardGroup ?? new List<string>();
		IList<string> projectLines = projectGroup ?? new List<string>();
		int rowCount = Math.Max(standardLines.Count, projectLines.Count);
		if (rowCount <= 0)
		{
			return result;
		}
		int maxRows = Math.Min(rowCount, 10);
		checked
		{
			int num = maxRows - 1;
			for (int index = 0; index <= num; index++)
			{
				string standardLine = ((index < standardLines.Count) ? AppendSignatureDebugIdentity(standardLines[index], standardDebugIndex) : "-");
				string projectLine = ((index < projectLines.Count) ? AppendSignatureDebugIdentity(projectLines[index], projectDebugIndex) : "-");
				string kind;
				string details;
				if (index < standardLines.Count && index < projectLines.Count)
				{
					kind = "modified";
					details = "Signature value differs.";
				}
				else if (index < standardLines.Count)
				{
					kind = "standard-only";
					details = "Only in standard.";
				}
				else
				{
					kind = "project-only";
					details = "Only in project.";
				}
				AddDifferenceDetail(result, groupName, kind, standardLine, projectLine, details);
			}
			if (rowCount > maxRows)
			{
				AddDifferenceDetail(result, groupName, "omitted", "-", "-", (rowCount - maxRows).ToString(CultureInfo.InvariantCulture) + " additional differences omitted.");
			}
			return result;
		}
	}

	private static List<LoadableFingerprintDifferenceDetailItem> BuildNestedLabelDifferenceDetails(IEnumerable<string> standardGroup, IEnumerable<string> projectGroup)
	{
		List<LoadableFingerprintDifferenceDetailItem> result = new List<LoadableFingerprintDifferenceDetailItem>();
		Dictionary<string, NestedLabelSignatureInfo> standardLabels = BuildNestedLabelInfoMap(standardGroup);
		Dictionary<string, NestedLabelSignatureInfo> projectLabels = BuildNestedLabelInfoMap(projectGroup);
		if (standardLabels.Count == 0 && projectLabels.Count == 0)
		{
			return result;
		}
		List<string> keys = standardLabels.Keys.Union<string>(projectLabels.Keys, StringComparer.Ordinal).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal).ToList();
		foreach (string key in keys)
		{
			NestedLabelSignatureInfo standardInfo = null;
			NestedLabelSignatureInfo projectInfo = null;
			standardLabels.TryGetValue(key, out standardInfo);
			projectLabels.TryGetValue(key, out projectInfo);
			if (standardInfo != null && projectInfo != null)
			{
				if (!string.Equals(standardInfo.ComparisonValue, projectInfo.ComparisonValue, StringComparison.Ordinal))
				{
					AddDifferenceDetail(result, "nested labels", "modified", FormatNestedLabelDiffCell(standardInfo), FormatNestedLabelDiffCell(projectInfo), BuildNestedLabelDifferenceReason(standardInfo, projectInfo));
				}
			}
			else if (standardInfo != null)
			{
				AddDifferenceDetail(result, "nested labels", "standard-only", FormatNestedLabelDiffCell(standardInfo), "-", "Nested label is missing in project.");
			}
			else if (projectInfo != null)
			{
				AddDifferenceDetail(result, "nested labels", "project-only", "-", FormatNestedLabelDiffCell(projectInfo), "Nested label exists only in project.");
			}
		}
		return result;
	}

	private static Dictionary<string, NestedLabelSignatureInfo> BuildNestedLabelInfoMap(IEnumerable<string> lines)
	{
		Dictionary<string, NestedLabelSignatureInfo> result = new Dictionary<string, NestedLabelSignatureInfo>(StringComparer.Ordinal);
		if (lines == null)
		{
			return result;
		}
		Dictionary<string, int> duplicateCounts = new Dictionary<string, int>(StringComparer.Ordinal);
		foreach (string item in ExpandNestedLabelEntries(lines))
		{
			NestedLabelSignatureInfo info = ParseNestedLabelSignatureEntry(item);
			if (info != null && !string.IsNullOrWhiteSpace(info.InstanceKey))
			{
				string key = info.InstanceKey;
				int duplicateCount = 0;
				duplicateCounts.TryGetValue(key, out duplicateCount);
				duplicateCounts[key] = checked(duplicateCount + 1);
				if (duplicateCount > 0)
				{
					key = key + "#" + duplicateCount.ToString(CultureInfo.InvariantCulture);
				}
				result[key] = info;
			}
		}
		return result;
	}

	private static List<string> ExpandNestedLabelEntries(IEnumerable<string> lines)
	{
		List<string> result = new List<string>();
		if (lines == null)
		{
			return result;
		}
		foreach (string line2 in lines)
		{
			string line = (line2 ?? string.Empty).Trim();
			if (line.Length == 0)
			{
				continue;
			}
			if (line.StartsWith("nested-labels=", StringComparison.OrdinalIgnoreCase))
			{
				line = line.Substring("nested-labels=".Length).Trim();
			}
			string[] array = line.Split(';');
			for (int i = 0; i < array.Length; i = checked(i + 1))
			{
				string entry = (array[i] ?? string.Empty).Trim();
				if (entry.Length > 0)
				{
					result.Add(entry);
				}
			}
		}
		return result;
	}

	private static NestedLabelSignatureInfo ParseNestedLabelSignatureEntry(string entry)
	{
		string text = (entry ?? string.Empty).Trim();
		if (text.Length == 0)
		{
			return null;
		}
		int arrowIndex = text.IndexOf("=>", StringComparison.Ordinal);
		if (arrowIndex < 0)
		{
			return null;
		}
		string left = text.Substring(0, arrowIndex).Trim();
		checked
		{
			string right = text.Substring(arrowIndex + 2).Trim();
			if (left.Length == 0)
			{
				return null;
			}
			List<string> leftParts = SplitSignatureSegments(left);
			List<string> rightParts = SplitSignatureSegments(right);
			string parameterName = ((leftParts.Count > 0) ? leftParts[leftParts.Count - 1] : string.Empty);
			string labelName = GetSignatureSegmentValue(rightParts, "label");
			string nestedFamilyName = GetSignatureSegmentValue(rightParts, "nested-family");
			string nestedTypeName = GetSignatureSegmentValue(rightParts, "nested-type");
			string nestedCategoryName = GetSignatureSegmentValue(rightParts, "nested-category");
			string storageName = string.Empty;
			string roleName = string.Empty;
			string formula = string.Empty;
			if (string.IsNullOrWhiteSpace(labelName) && rightParts.Count > 0)
			{
				labelName = rightParts[0];
				storageName = ((rightParts.Count > 1) ? rightParts[1] : string.Empty);
				roleName = ((rightParts.Count > 2) ? FormatNestedLabelRole(rightParts[2]) : string.Empty);
				formula = ((rightParts.Count > 3) ? rightParts[3] : string.Empty);
			}
			string comparison = ((!string.IsNullOrWhiteSpace(nestedFamilyName) || !string.IsNullOrWhiteSpace(nestedTypeName) || !string.IsNullOrWhiteSpace(nestedCategoryName)) ? (labelName + "|" + nestedFamilyName + "|" + nestedTypeName + "|" + nestedCategoryName) : string.Join("|", rightParts));
			return new NestedLabelSignatureInfo
			{
				InstanceKey = Normalize(left),
				InstanceDisplay = BuildNestedLabelInstanceDisplay(leftParts),
				ParameterName = parameterName,
				LabelName = labelName,
				NestedFamilyName = nestedFamilyName,
				NestedTypeName = nestedTypeName,
				NestedCategoryName = nestedCategoryName,
				StorageName = storageName,
				RoleName = roleName,
				Formula = formula,
				ComparisonValue = Normalize(comparison)
			};
		}
	}

	private static List<string> SplitSignatureSegments(string value)
	{
		return (from x in (value ?? string.Empty).Split('|')
			select (x ?? string.Empty).Trim() into x
			where !string.IsNullOrWhiteSpace(x)
			select x).ToList();
	}

	private static string GetSignatureSegmentValue(IEnumerable<string> segments, string key)
	{
		if (segments == null || string.IsNullOrWhiteSpace(key))
		{
			return string.Empty;
		}
		string prefix = key.Trim() + "=";
		foreach (string segment in segments)
		{
			string text = (segment ?? string.Empty).Trim();
			if (text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
			{
				return text.Substring(prefix.Length).Trim();
			}
		}
		return string.Empty;
	}

	private static string BuildNestedLabelInstanceDisplay(IList<string> parts)
	{
		if (parts == null || parts.Count == 0)
		{
			return "Instance";
		}
		string familyName = ((parts.Count > 1) ? parts[1] : string.Empty);
		string typeName = ((parts.Count > 2) ? parts[2] : string.Empty);
		string instanceName = ((parts.Count > 3) ? parts[3] : string.Empty);
		string parameterName = ((parts.Count > 4) ? parts[4] : string.Empty);
		List<string> labels = new List<string>();
		if (!string.IsNullOrWhiteSpace(familyName))
		{
			labels.Add("Family " + familyName);
		}
		if (!string.IsNullOrWhiteSpace(typeName) && !string.Equals(typeName, familyName, StringComparison.OrdinalIgnoreCase))
		{
			labels.Add("Type " + typeName);
		}
		if (!string.IsNullOrWhiteSpace(instanceName) && !string.Equals(instanceName, typeName, StringComparison.OrdinalIgnoreCase))
		{
			labels.Add("Instance " + instanceName);
		}
		if (!string.IsNullOrWhiteSpace(parameterName))
		{
			labels.Add("Parameter " + parameterName);
		}
		if (labels.Count == 0)
		{
			labels.Add(string.Join(" · ", parts.Take(Math.Min(4, parts.Count))));
		}
		return ShortDiffCell(string.Join(" · ", labels), 160);
	}

	private static string FormatNestedLabelRole(string value)
	{
		string text = Normalize(value);
		if (string.Equals(text, "true", StringComparison.Ordinal))
		{
			return "instance";
		}
		if (string.Equals(text, "false", StringComparison.Ordinal))
		{
			return "type";
		}
		return (value ?? string.Empty).Trim();
	}

	private static string FormatNestedLabelDiffCell(NestedLabelSignatureInfo info)
	{
		if (info == null)
		{
			return "-";
		}
		List<string> parts = new List<string>();
		if (!string.IsNullOrWhiteSpace(info.InstanceDisplay))
		{
			parts.Add(info.InstanceDisplay);
		}
		string labelText = (string.IsNullOrWhiteSpace(info.LabelName) ? "-" : info.LabelName);
		parts.Add("Label parameter " + labelText);
		if (!string.IsNullOrWhiteSpace(info.NestedFamilyName) || !string.IsNullOrWhiteSpace(info.NestedTypeName))
		{
			List<string> valueParts = new List<string>();
			if (!string.IsNullOrWhiteSpace(info.NestedFamilyName))
			{
				valueParts.Add(info.NestedFamilyName);
			}
			if (!string.IsNullOrWhiteSpace(info.NestedTypeName))
			{
				valueParts.Add(info.NestedTypeName);
			}
			parts.Add("Value " + string.Join(" · ", valueParts));
		}
		List<string> attributes = new List<string>();
		if (!string.IsNullOrWhiteSpace(info.NestedCategoryName))
		{
			attributes.Add("category=" + info.NestedCategoryName);
		}
		if (!string.IsNullOrWhiteSpace(info.RoleName))
		{
			attributes.Add(info.RoleName);
		}
		if (!string.IsNullOrWhiteSpace(info.StorageName))
		{
			attributes.Add(info.StorageName);
		}
		if (!string.IsNullOrWhiteSpace(info.Formula))
		{
			attributes.Add("formula=" + info.Formula);
		}
		if (attributes.Count > 0)
		{
			parts.Add(string.Join(", ", attributes));
		}
		return ShortDiffCell(string.Join(" · ", parts), 260);
	}

	private static string BuildNestedLabelDifferenceReason(NestedLabelSignatureInfo standardInfo, NestedLabelSignatureInfo projectInfo)
	{
		if (standardInfo == null || projectInfo == null)
		{
			return "Nested label differs.";
		}
		List<string> changes = new List<string>();
		if (!string.Equals(Normalize(standardInfo.LabelName), Normalize(projectInfo.LabelName), StringComparison.Ordinal))
		{
			changes.Add("Label parameter");
		}
		if (!string.Equals(Normalize(standardInfo.NestedFamilyName), Normalize(projectInfo.NestedFamilyName), StringComparison.Ordinal))
		{
			changes.Add("nested family");
		}
		if (!string.Equals(Normalize(standardInfo.NestedTypeName), Normalize(projectInfo.NestedTypeName), StringComparison.Ordinal))
		{
			changes.Add("nested type");
		}
		if (!string.Equals(Normalize(standardInfo.NestedCategoryName), Normalize(projectInfo.NestedCategoryName), StringComparison.Ordinal))
		{
			changes.Add("nested category");
		}
		if (changes.Count == 0 && !string.Equals(Normalize(standardInfo.RoleName), Normalize(projectInfo.RoleName), StringComparison.Ordinal))
		{
			changes.Add("type/instance");
		}
		if (changes.Count == 0 && !string.Equals(Normalize(standardInfo.StorageName), Normalize(projectInfo.StorageName), StringComparison.Ordinal))
		{
			changes.Add("storage");
		}
		if (changes.Count == 0 && !string.Equals(Normalize(standardInfo.Formula), Normalize(projectInfo.Formula), StringComparison.Ordinal))
		{
			changes.Add("formula");
		}
		if (changes.Count == 0)
		{
			return "Label definition differs.";
		}
		return "Nested label differs: " + string.Join(", ", changes);
	}

	private static void AddDifferenceDetail(List<LoadableFingerprintDifferenceDetailItem> result, string area, string differenceKind, string standardValue, string projectValue, string details)
	{
		result?.Add(new LoadableFingerprintDifferenceDetailItem
		{
			Area = (area ?? string.Empty),
			DifferenceKind = (differenceKind ?? string.Empty),
			StandardValue = ShortDiffCell(string.IsNullOrWhiteSpace(standardValue) ? "-" : standardValue, 320),
			ProjectValue = ShortDiffCell(string.IsNullOrWhiteSpace(projectValue) ? "-" : projectValue, 320),
			Details = ShortDiffCell(details ?? string.Empty, 260)
		});
	}

	private static void AppendDifferenceDetails(List<LoadableFingerprintDifferenceDetailItem> target, IEnumerable<LoadableFingerprintDifferenceDetailItem> source, int maxTotal)
	{
		if (target == null || source == null)
		{
			return;
		}
		foreach (LoadableFingerprintDifferenceDetailItem item in source)
		{
			if (item != null)
			{
				target.Add(item);
				if (target.Count >= maxTotal)
				{
					break;
				}
			}
		}
	}

	private static string FormatSignatureGroupForDiff(IEnumerable<string> lines, Dictionary<string, List<string>> debugIndex = null)
	{
		if (lines == null)
		{
			return "-";
		}
		List<string> items = (from x in lines.Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)).Take(2)
			select ShortSignatureCell(AppendSignatureDebugIdentity(x, debugIndex))).ToList();
		if (items.Count == 0)
		{
			return "-";
		}
		return string.Join("; ", items);
	}

	private static string AppendSignatureDebugIdentity(string line, Dictionary<string, List<string>> debugIndex)
	{
		string text = (line ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(text) || debugIndex == null || debugIndex.Count == 0)
		{
			return text;
		}
		List<string> identities = null;
		if (!debugIndex.TryGetValue(NormalizeSignatureLine(text), out identities) || identities == null || identities.Count == 0)
		{
			return text;
		}
		return text + " [" + string.Join(", ", identities.Take(2)) + ((identities.Count > 2) ? (", +" + checked(identities.Count - 2).ToString(CultureInfo.InvariantCulture)) : string.Empty) + "]";
	}

	private static string FormatSignatureDebugIdentities(IEnumerable<string> lines, Dictionary<string, List<string>> debugIndex)
	{
		if (lines == null || debugIndex == null || debugIndex.Count == 0)
		{
			return string.Empty;
		}
		List<string> values = new List<string>();
		foreach (string line in lines)
		{
			if (!string.IsNullOrWhiteSpace(line))
			{
				List<string> identities = null;
				if (debugIndex.TryGetValue(NormalizeSignatureLine(line), out identities) && identities != null)
				{
					values.AddRange(identities);
				}
			}
		}
		List<string> distinctValues = values.Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)).Distinct<string>(StringComparer.OrdinalIgnoreCase).Take(6)
			.ToList();
		if (distinctValues.Count == 0)
		{
			return string.Empty;
		}
		return string.Join(", ", distinctValues);
	}

	private static string LimitListDisplay(IEnumerable<string> items, int maxCount)
	{
		if (items == null)
		{
			return "-";
		}
		List<string> list = items.Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)).Take(Math.Max(1, maxCount)).ToList();
		if (list.Count == 0)
		{
			return "-";
		}
		string text = string.Join(", ", list);
		int total = items.Count([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x));
		if (total > list.Count)
		{
			text = text + " (+" + checked(total - list.Count).ToString(CultureInfo.InvariantCulture) + " more)";
		}
		return text;
	}

	private static string ShortHash(string value)
	{
		string text = (value ?? string.Empty).Trim();
		if (text.Length == 0)
		{
			return "-";
		}
		if (text.Length <= 14)
		{
			return text;
		}
		return text.Substring(0, 14) + "...";
	}

	private static string ShortDiffCell(string value, int maxLength)
	{
		string text = (value ?? string.Empty).Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ")
			.Replace("\t", " ")
			.Trim();
		if (text.Length <= maxLength)
		{
			return text;
		}
		return text.Substring(0, Math.Max(0, checked(maxLength - 3))) + "...";
	}

	private static List<string> BuildContentSignatureDifferenceSummary(string standardSignaturePath, string projectSignaturePath)
	{
		List<string> result = new List<string>();
		List<string> standardLines = ReadSignatureSourceLines(standardSignaturePath);
		List<string> projectLines = ReadSignatureSourceLines(projectSignaturePath);
		if (standardLines.Count == 0 || projectLines.Count == 0)
		{
			if (standardLines.Count == 0 && projectLines.Count == 0)
			{
				result.Add("Signature diagnostics missing.");
			}
			else if (standardLines.Count == 0)
			{
				result.Add("Standard signature diagnostics missing.");
			}
			else
			{
				result.Add("Project signature diagnostics missing.");
			}
			return result;
		}
		SignatureLineDelta signatureLineDelta = BuildSignatureLineDelta(standardLines, projectLines);
		List<string> standardOnly = signatureLineDelta.StandardOnly;
		List<string> projectOnly = signatureLineDelta.ProjectOnly;
		if (standardOnly.Count == 0 && projectOnly.Count == 0)
		{
			return result;
		}
		List<IGrouping<string, string>> groupedStandardOnly = (from x in standardOnly.GroupBy<string, string>([SpecialName] (string x) => DescribeSignatureLine(x), StringComparer.Ordinal)
			where !IsFamilyTypesSignatureGroup(x.Key)
			orderby SignatureDifferenceGroupOrder(x.Key)
			select x).ThenBy<IGrouping<string, string>, string>([SpecialName] (IGrouping<string, string> x) => x.Key, StringComparer.Ordinal).ToList();
		List<IGrouping<string, string>> groupedProjectOnly = (from x in projectOnly.GroupBy<string, string>([SpecialName] (string x) => DescribeSignatureLine(x), StringComparer.Ordinal)
			where !IsFamilyTypesSignatureGroup(x.Key)
			orderby SignatureDifferenceGroupOrder(x.Key)
			select x).ThenBy<IGrouping<string, string>, string>([SpecialName] (IGrouping<string, string> x) => x.Key, StringComparer.Ordinal).ToList();
		foreach (IGrouping<string, string> group in groupedStandardOnly.Take(6))
		{
			result.Add("Signature differs [" + group.Key + "]: standard-only " + group.Count().ToString(CultureInfo.InvariantCulture));
		}
		foreach (IGrouping<string, string> group2 in groupedProjectOnly.Take(6))
		{
			result.Add("Signature differs [" + group2.Key + "]: project-only " + group2.Count().ToString(CultureInfo.InvariantCulture));
		}
		int omitted = checked(Math.Max(0, groupedStandardOnly.Count - 6) + Math.Max(0, groupedProjectOnly.Count - 6));
		if (omitted > 0)
		{
			result.Add("Additional signature difference groups omitted: " + omitted.ToString(CultureInfo.InvariantCulture));
		}
		return result;
	}

	private static List<string> ReadSignatureSourceLines(string signaturePath)
	{
		List<string> result = new List<string>();
		List<string> ReadSignatureSourceLines;
		if (string.IsNullOrWhiteSpace(signaturePath) || !File.Exists(signaturePath))
		{
			ReadSignatureSourceLines = result;
		}
		else
		{
			try
			{
				bool afterMarker = false;
				foreach (string rawLine in File.ReadLines(signaturePath))
				{
					if (rawLine == null)
					{
						continue;
					}
					string line = rawLine.Trim();
					if (!afterMarker)
					{
						if (string.Equals(line, "----- signature-source -----", StringComparison.OrdinalIgnoreCase))
						{
							afterMarker = true;
						}
					}
					else if (line.Length != 0)
					{
						result.Add(NormalizeSignatureLine(line));
					}
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ReadSignatureSourceLines = new List<string>();
				ProjectData.ClearProjectError();
				goto IL_00f5;
			}
			ReadSignatureSourceLines = result.Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal).ToList();
		}
		goto IL_00f5;
		IL_00f5:
		return ReadSignatureSourceLines;
	}

	private static Dictionary<string, List<string>> ReadSignatureElementDebugIndex(string signaturePath)
	{
		Dictionary<string, List<string>> result = new Dictionary<string, List<string>>(StringComparer.Ordinal);
		Dictionary<string, List<string>> ReadSignatureElementDebugIndex;
		if (string.IsNullOrWhiteSpace(signaturePath) || !File.Exists(signaturePath))
		{
			ReadSignatureElementDebugIndex = result;
		}
		else
		{
			try
			{
				bool inDebugSection = false;
				foreach (string rawLine in File.ReadLines(signaturePath))
				{
					if (rawLine == null)
					{
						continue;
					}
					string line = rawLine.Trim();
					if (line.Length == 0)
					{
						continue;
					}
					if (string.Equals(line, "----- signature-debug -----", StringComparison.OrdinalIgnoreCase))
					{
						inDebugSection = true;
					}
					else if (line.StartsWith("----- ", StringComparison.Ordinal))
					{
						if (inDebugSection)
						{
							break;
						}
					}
					else
					{
						if (!inDebugSection)
						{
							continue;
						}
						string[] parts = line.Split(new char[1] { '\t' }, 3);
						if (parts.Length < 3 || !string.Equals(parts[0].Trim(), "element", StringComparison.OrdinalIgnoreCase))
						{
							continue;
						}
						string key = NormalizeSignatureLine(parts[1]);
						string identity = FormatDebugIdentityForDisplay(parts[2]);
						if (!string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(identity))
						{
							List<string> list = null;
							if (!result.TryGetValue(key, out list) || list == null)
							{
								list = (result[key] = new List<string>());
							}
							if (!list.Contains<string>(identity, StringComparer.OrdinalIgnoreCase))
							{
								list.Add(identity);
							}
						}
					}
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ReadSignatureElementDebugIndex = new Dictionary<string, List<string>>(StringComparer.Ordinal);
				ProjectData.ClearProjectError();
				goto IL_0159;
			}
			ReadSignatureElementDebugIndex = result;
		}
		goto IL_0159;
		IL_0159:
		return ReadSignatureElementDebugIndex;
	}

	private static string FormatDebugIdentityForDisplay(string value)
	{
		string text = (value ?? string.Empty).Trim();
		if (text.Length == 0)
		{
			return string.Empty;
		}
		string idValue = ExtractDebugTokenValue(text, "id=");
		string categoryValue = ExtractDebugTokenValue(text, "category=");
		string nestedFamilyValue = ExtractDebugTokenValue(text, "nestedFamily=");
		string nestedTypeValue = ExtractDebugTokenValue(text, "nestedType=");
		string familyValue = ExtractDebugTokenValue(text, "family=");
		List<string> result = new List<string>();
		if (!string.IsNullOrWhiteSpace(idValue))
		{
			result.Add("ID " + idValue);
		}
		AddDebugDisplayToken(result, "nestedFamily", nestedFamilyValue);
		AddDebugDisplayToken(result, "nestedType", nestedTypeValue);
		if (string.IsNullOrWhiteSpace(nestedFamilyValue))
		{
			AddDebugDisplayToken(result, "family", familyValue);
		}
		AddDebugDisplayToken(result, "category", categoryValue);
		if (result.Count == 0)
		{
			return ShortDebugIdentity(text, 48);
		}
		return string.Join(" / ", result);
	}

	private static void AddDebugDisplayToken(IList<string> result, string key, string value)
	{
		if (result != null && !string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(value))
		{
			result.Add(key + "=\"" + value.Replace('"', '\'').Trim() + "\"");
		}
	}

	private static string ExtractDebugTokenValue(string text, string marker)
	{
		if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(marker))
		{
			return string.Empty;
		}
		int index = text.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
		if (index < 0)
		{
			return string.Empty;
		}
		checked
		{
			index += marker.Length;
			if (index < text.Length && text[index] == '"')
			{
				index++;
				int quotedEndIndex = text.IndexOf('"', index);
				if (quotedEndIndex < 0 || quotedEndIndex <= index)
				{
					return string.Empty;
				}
				return text.Substring(index, quotedEndIndex - index).Trim();
			}
			int endIndex = text.IndexOf(' ', index);
			if (endIndex < 0)
			{
				endIndex = text.Length;
			}
			if (endIndex <= index)
			{
				return string.Empty;
			}
			return text.Substring(index, endIndex - index).Trim();
		}
	}

	private static string ShortDebugIdentity(string value, int maxLength)
	{
		string text = (value ?? string.Empty).Trim();
		if (text.Length <= maxLength)
		{
			return text;
		}
		return text.Substring(0, Math.Max(0, checked(maxLength - 3))) + "...";
	}

	private static bool ContentSignatureSourcesHaveSameComparableLines(string standardSignaturePath, string projectSignaturePath)
	{
		List<string> standardLines = ReadSignatureSourceLines(standardSignaturePath);
		List<string> projectLines = ReadSignatureSourceLines(projectSignaturePath);
		if (standardLines.Count == 0 || projectLines.Count == 0)
		{
			return false;
		}
		return standardLines.SequenceEqual<string>(projectLines, StringComparer.Ordinal);
	}

	private static SignatureLineDelta BuildSignatureLineDelta(IEnumerable<string> standardLines, IEnumerable<string> projectLines)
	{
		Dictionary<string, int> standardCounts = BuildSignatureLineCountMap(standardLines);
		Dictionary<string, int> projectCounts = BuildSignatureLineCountMap(projectLines);
		SignatureLineDelta result = new SignatureLineDelta();
		List<string> allLines = standardCounts.Keys.Union<string>(projectCounts.Keys, StringComparer.Ordinal).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal).ToList();
		checked
		{
			foreach (string line in allLines)
			{
				int standardCount = 0;
				int projectCount = 0;
				standardCounts.TryGetValue(line, out standardCount);
				projectCounts.TryGetValue(line, out projectCount);
				if (standardCount > projectCount)
				{
					int num = standardCount - projectCount;
					for (int index = 1; index <= num; index++)
					{
						result.StandardOnly.Add(line);
					}
				}
				else if (projectCount > standardCount)
				{
					int num2 = projectCount - standardCount;
					for (int index2 = 1; index2 <= num2; index2++)
					{
						result.ProjectOnly.Add(line);
					}
				}
			}
			return result;
		}
	}

	private static Dictionary<string, int> BuildSignatureLineCountMap(IEnumerable<string> lines)
	{
		Dictionary<string, int> result = new Dictionary<string, int>(StringComparer.Ordinal);
		if (lines == null)
		{
			return result;
		}
		foreach (string line in lines)
		{
			string key = NormalizeSignatureLine(line);
			if (!string.IsNullOrWhiteSpace(key))
			{
				int count = 0;
				result.TryGetValue(key, out count);
				result[key] = checked(count + 1);
			}
		}
		return result;
	}

	private static string NormalizeSignatureLine(string line)
	{
		if (line == null)
		{
			return string.Empty;
		}
		return NormalizeMaterialSignatureLineByName(line.Trim().ToLowerInvariant());
	}

	private static string NormalizeMaterialSignatureLineByName(string line)
	{
		if (string.IsNullOrWhiteSpace(line))
		{
			return string.Empty;
		}
		string working = line.Trim();
		if (working.StartsWith("elements=", StringComparison.Ordinal))
		{
			working = working.Substring("elements=".Length).Trim();
		}
		if (!working.StartsWith("material|", StringComparison.Ordinal))
		{
			return line;
		}
		string[] tokens = working.Split('|');
		if (tokens.Length < 3)
		{
			return line;
		}
		string materialName = (tokens[2] ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(materialName))
		{
			return line;
		}
		return "material|material|" + materialName;
	}

	private static string DescribeSignatureLine(string line)
	{
		string value = Normalize(line);
		if (value.StartsWith("content-signature-version=", StringComparison.Ordinal))
		{
			return "signature version";
		}
		if (value.StartsWith("family=", StringComparison.Ordinal))
		{
			return "family name";
		}
		if (value.StartsWith("category=", StringComparison.Ordinal))
		{
			return "category";
		}
		if (value.StartsWith("lookup-tables=", StringComparison.Ordinal) || value.StartsWith("table=", StringComparison.Ordinal) || value.StartsWith("lookup-table-error=", StringComparison.Ordinal))
		{
			return "lookup tables";
		}
		string elementGroup = DescribeElementSignatureGroup(value);
		if (!string.IsNullOrWhiteSpace(elementGroup))
		{
			return elementGroup;
		}
		if (value.StartsWith("params=", StringComparison.Ordinal) || value.Contains("|guid:") || value.Contains("formula") || value.Contains("parameter"))
		{
			return "parameters/formulas";
		}
		if (value.StartsWith("types=", StringComparison.Ordinal))
		{
			return "family types";
		}
		if (value.StartsWith("nested-labels=", StringComparison.Ordinal) || value.Contains("=>"))
		{
			return "nested labels";
		}
		if (value.Contains("familyinstance") || value.Contains("symbol") || value.Contains("nested"))
		{
			return "nested/loadable instances";
		}
		if (value.Contains("solid:") || value.Contains("mesh:") || value.Contains("curve:") || value.Contains("geometry:"))
		{
			return "geometry";
		}
		if (value.Contains("connector"))
		{
			return "connectors";
		}
		return "elements";
	}

	private static bool IsFamilyTypesSignatureGroup(string groupName)
	{
		return string.Equals(Normalize(groupName).Replace(" ", string.Empty).Replace("-", string.Empty).Replace("_", string.Empty)
			.Replace("/", string.Empty), "familytypes", StringComparison.Ordinal);
	}

	private static string DescribeElementSignatureGroup(string value)
	{
		if (string.IsNullOrWhiteSpace(value) || value.IndexOf('|') < 0 || value.StartsWith("params=", StringComparison.Ordinal) || value.StartsWith("types=", StringComparison.Ordinal) || value.StartsWith("nested-labels=", StringComparison.Ordinal))
		{
			return string.Empty;
		}
		string firstToken = value.Split('|')[0].Trim();
		switch (firstToken)
		{
		case "familyinstance":
			return "nested/loadable instances";
		case "familysymbol":
		case "familytype":
			return "nested/loadable types";
		case "family":
			return "nested/loadable families";
		case "dimension":
			return "elements/dimension";
		case "referenceplane":
			return "elements/reference plane";
		case "modelcurve":
		case "modelline":
		case "modelarc":
		case "modelellipse":
		case "modelnurbsspline":
			return "elements/model line";
		case "detailcurve":
		case "detailline":
		case "detailarc":
		case "detailellipse":
		case "detailnurbsspline":
			return "elements/detail line";
		case "textnote":
		case "textelement":
			return "elements/text";
		case "filledregion":
			return "elements/filled region";
		case "extrusion":
		case "blend":
		case "sweep":
		case "sweptblend":
		case "revolution":
		case "form":
		case "genericform":
		case "freeformelement":
			return "elements/form";
		default:
			if (firstToken.Length == 0 || firstToken.Contains("=") || firstToken.Contains(":"))
			{
				return string.Empty;
			}
			if (firstToken.Contains("connector"))
			{
				return "connectors";
			}
			return "elements/" + firstToken;
		}
	}

	private static int SignatureDifferenceGroupOrder(string groupName)
	{
		string normalizedGroup = Normalize(groupName);
		if (normalizedGroup.StartsWith("elements/", StringComparison.Ordinal))
		{
			return 9;
		}
		return normalizedGroup switch
		{
			"family types" => 0,
			"parameters/formulas" => 1,
			"lookup tables" => 2,
			"nested labels" => 3,
			"nested/loadable instances" => 4,
			"nested/loadable types" => 5,
			"nested/loadable families" => 6,
			"connectors" => 7,
			"geometry" => 8,
			"elements" => 9,
			_ => 10,
		};
	}

	private static string ShortSignatureLine(string line)
	{
		if (string.IsNullOrWhiteSpace(line))
		{
			return "(empty)";
		}
		string trimmed = line.Trim().Replace("|", " / ");
		if (trimmed.Length <= 180)
		{
			return "'" + trimmed + "'";
		}
		return "'" + trimmed.Substring(0, 180) + "...'";
	}

	private static string ShortSignatureCell(string line)
	{
		if (string.IsNullOrWhiteSpace(line))
		{
			return "-";
		}
		string trimmed = line.Trim().Replace("|", " / ");
		if (trimmed.Length <= 220)
		{
			return trimmed;
		}
		return trimmed.Substring(0, 217) + "...";
	}

	private static bool IsProjectLoadableContentFingerprintMissing(StandardLoadableFamilySnapshotItem standardFamily, ProjectLoadableFamilySnapshotItem projectFamily)
	{
		if (standardFamily == null || projectFamily == null)
		{
			return false;
		}
		return !string.IsNullOrWhiteSpace(standardFamily.ContentFingerprint) && string.IsNullOrWhiteSpace(projectFamily.ContentFingerprint);
	}

	private static void AppendSignatureInventoryTypeDifferences(StandardLoadableFamilySnapshotItem standardFamily, ProjectLoadableFamilySnapshotItem projectFamily, List<string> missingTypes, List<string> extraTypes)
	{
		if (standardFamily == null || projectFamily == null)
		{
			return;
		}
		List<string> standardTypes = ExtractFamilyManagerTypeNamesFromSignature(standardFamily.ContentSignatureDebugPath);
		List<string> projectTypes = ExtractFamilyManagerTypeNamesFromSignature(projectFamily.ContentSignatureDebugPath);
		if (standardTypes.Count == 0 || projectTypes.Count == 0)
		{
			return;
		}
		foreach (string familyTypeName in standardTypes.Except<string>(projectTypes, StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase))
		{
			AddDistinctTypeName(missingTypes, familyTypeName);
		}
		foreach (string familyTypeName2 in projectTypes.Except<string>(standardTypes, StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase))
		{
			AddDistinctTypeName(extraTypes, familyTypeName2);
		}
	}

	private static void AddDistinctTypeName(List<string> items, string typeName)
	{
		if (items != null && !string.IsNullOrWhiteSpace(typeName) && !items.Any([SpecialName] (string x) => string.Equals(x, typeName, StringComparison.OrdinalIgnoreCase)))
		{
			items.Add(typeName);
		}
	}

	private static List<string> CloneStringList(IEnumerable<string> values)
	{
		if (values == null)
		{
			return new List<string>();
		}
		return (from x in values
			where !string.IsNullOrWhiteSpace(x)
			select x.Trim()).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
	}

	private static List<StandardFamilyParameterSnapshotItem> CloneParameterSnapshotItems(IEnumerable<StandardFamilyParameterSnapshotItem> parameters)
	{
		List<StandardFamilyParameterSnapshotItem> result = new List<StandardFamilyParameterSnapshotItem>();
		if (parameters == null)
		{
			return result;
		}
		foreach (StandardFamilyParameterSnapshotItem item in parameters)
		{
			if (item != null)
			{
				result.Add(new StandardFamilyParameterSnapshotItem
				{
					Scope = (item.Scope ?? string.Empty),
					TypeName = (item.TypeName ?? string.Empty),
					Name = (item.Name ?? string.Empty),
					StorageType = (item.StorageType ?? string.Empty),
					ValuePreview = (item.ValuePreview ?? string.Empty),
					Formula = (item.Formula ?? string.Empty),
					IsInstance = item.IsInstance,
					IsReadOnly = item.IsReadOnly,
					IsShared = item.IsShared,
					ParameterId = (item.ParameterId ?? string.Empty),
					ExternalGuid = (item.ExternalGuid ?? string.Empty)
				});
			}
		}
		return result;
	}

	private static List<StandardNestedLoadableFamilySnapshotItem> BuildNestedLoadableFamiliesFromSignature(string signaturePath)
	{
		Dictionary<string, StandardNestedLoadableFamilySnapshotItem> result = new Dictionary<string, StandardNestedLoadableFamilySnapshotItem>(StringComparer.Ordinal);
		List<StandardNestedLoadableFamilySnapshotItem> BuildNestedLoadableFamiliesFromSignature;
		string resolvedSignaturePath = string.IsNullOrWhiteSpace(signaturePath) ? string.Empty : Environment.ExpandEnvironmentVariables(signaturePath.Trim());
		if (string.IsNullOrWhiteSpace(resolvedSignaturePath) || !File.Exists(resolvedSignaturePath))
		{
			BuildNestedLoadableFamiliesFromSignature = new List<StandardNestedLoadableFamilySnapshotItem>();
		}
		else
		{
			try
			{
				foreach (StandardNestedLoadableFamilySnapshotItem debugItem in BuildNestedLoadableFamiliesFromSignatureDebug(resolvedSignaturePath))
				{
					string debugKey = BuildNestedLoadableSignatureKey(debugItem.CategoryName, debugItem.FamilyName);
					if (!string.IsNullOrWhiteSpace(debugKey) && !result.ContainsKey(debugKey))
					{
						result.Add(debugKey, debugItem);
					}
				}
				bool inSource = false;
				_Closure_0024__63_002D0 closure_0024__63_002D = default(_Closure_0024__63_002D0);
				foreach (string rawLine in File.ReadLines(resolvedSignaturePath))
				{
					closure_0024__63_002D = new _Closure_0024__63_002D0(closure_0024__63_002D);
					if (rawLine == null)
					{
						continue;
					}
					string line = rawLine.Trim();
					if (!inSource)
					{
						if (string.Equals(line, "----- signature-source -----", StringComparison.OrdinalIgnoreCase))
						{
							inSource = true;
						}
						continue;
					}
					if (line.StartsWith("----- ", StringComparison.Ordinal))
					{
						break;
					}
					if (!line.StartsWith("familyinstance|", StringComparison.OrdinalIgnoreCase) && !line.StartsWith("familysymbol|", StringComparison.OrdinalIgnoreCase) && !line.StartsWith("family|", StringComparison.OrdinalIgnoreCase))
					{
						continue;
					}
					List<string> tokens = SplitSignatureTokens(line);
					string familyName = GetSignatureToken(tokens, 2);
					if (string.IsNullOrWhiteSpace(familyName))
					{
						continue;
					}
					string categoryName = GetSignatureToken(tokens, 1);
					string typeName = GetSignatureToken(tokens, 3);
					string key = BuildNestedLoadableSignatureKey(categoryName, familyName);
					if (string.IsNullOrWhiteSpace(key) || result.ContainsKey(key))
					{
						continue;
					}
					StandardNestedLoadableFamilySnapshotItem item = (result[key] = new StandardNestedLoadableFamilySnapshotItem
					{
						FamilyName = familyName.Trim(),
						CategoryName = categoryName.Trim(),
						CategoryId = string.Empty,
						CategoryGroup = FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(string.Empty, categoryName, string.Empty, familyName),
						IsShared = true
					});
					if (!string.IsNullOrWhiteSpace(typeName))
					{
						item.TypeNames.Add(typeName.Trim());
					}
					continue;
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				BuildNestedLoadableFamiliesFromSignature = new List<StandardNestedLoadableFamilySnapshotItem>();
				ProjectData.ClearProjectError();
				goto IL_0288;
			}
			foreach (StandardNestedLoadableFamilySnapshotItem value in result.Values)
			{
				value.TypeNames = CloneStringList(value.TypeNames);
				value.TypeCount = value.TypeNames.Count;
			}
			BuildNestedLoadableFamiliesFromSignature = result.Values.OrderBy<StandardNestedLoadableFamilySnapshotItem, string>([SpecialName] (StandardNestedLoadableFamilySnapshotItem x) => Normalize(x.CategoryName), StringComparer.Ordinal).ThenBy<StandardNestedLoadableFamilySnapshotItem, string>([SpecialName] (StandardNestedLoadableFamilySnapshotItem x) => Normalize(x.FamilyName), StringComparer.Ordinal).ToList();
		}
		goto IL_0288;
		IL_0288:
		return BuildNestedLoadableFamiliesFromSignature;
	}

	private static List<StandardNestedLoadableFamilySnapshotItem> BuildNestedLoadableFamiliesFromSignature(string signaturePath, string parentFamilyName)
	{
		List<StandardNestedLoadableFamilySnapshotItem> items = BuildNestedLoadableFamiliesFromSignature(signaturePath);
		string parentToken = Normalize(parentFamilyName);
		if (items == null || string.IsNullOrWhiteSpace(parentToken))
		{
			return items ?? new List<StandardNestedLoadableFamilySnapshotItem>();
		}
		return items.Where([SpecialName] (StandardNestedLoadableFamilySnapshotItem x) => x != null && !string.Equals(Normalize(x.FamilyName), parentToken, StringComparison.Ordinal)).ToList();
	}

	private static List<StandardNestedLoadableFamilySnapshotItem> BuildNestedLoadableFamiliesFromSignatureDebug(string signaturePath)
	{
		Dictionary<string, StandardNestedLoadableFamilySnapshotItem> result = new Dictionary<string, StandardNestedLoadableFamilySnapshotItem>(StringComparer.Ordinal);
		if (string.IsNullOrWhiteSpace(signaturePath) || !File.Exists(signaturePath))
		{
			return new List<StandardNestedLoadableFamilySnapshotItem>();
		}
		try
		{
			bool inDebug = false;
			foreach (string rawLine in File.ReadLines(signaturePath))
			{
				string line = (rawLine ?? string.Empty).Trim();
				if (!inDebug)
				{
					if (string.Equals(line, "----- signature-debug -----", StringComparison.OrdinalIgnoreCase))
					{
						inDebug = true;
					}
					continue;
				}
				if (line.StartsWith("----- ", StringComparison.Ordinal))
				{
					break;
				}
				if (!line.StartsWith("element\t", StringComparison.OrdinalIgnoreCase))
				{
					continue;
				}
				string[] parts = line.Split('\t');
				if (parts.Length < 3)
				{
					continue;
				}
				string signatureText = parts[1] ?? string.Empty;
				if (!signatureText.StartsWith("familyinstance|", StringComparison.OrdinalIgnoreCase) && !signatureText.StartsWith("familysymbol|", StringComparison.OrdinalIgnoreCase) && !signatureText.StartsWith("family|", StringComparison.OrdinalIgnoreCase))
				{
					continue;
				}
				List<string> tokens = SplitSignatureTokens(signatureText);
				string identity = parts[2] ?? string.Empty;
				string familyName = FirstNonEmptyToken(ExtractSignatureDebugQuotedValue(identity, "nestedFamily"), ExtractSignatureDebugQuotedValue(identity, "family"), GetSignatureToken(tokens, 2));
				if (string.IsNullOrWhiteSpace(familyName))
				{
					continue;
				}
				string categoryName = FirstNonEmptyToken(ExtractSignatureDebugQuotedValue(identity, "category"), GetSignatureToken(tokens, 1));
				string typeName = FirstNonEmptyToken(ExtractSignatureDebugQuotedValue(identity, "nestedType"), ExtractSignatureDebugQuotedValue(identity, "type"), GetSignatureToken(tokens, 3));
				string key = BuildNestedLoadableSignatureKey(categoryName, familyName);
				if (string.IsNullOrWhiteSpace(key))
				{
					continue;
				}
				StandardNestedLoadableFamilySnapshotItem item = null;
				if (!result.TryGetValue(key, out item) || item == null)
				{
					item = (result[key] = new StandardNestedLoadableFamilySnapshotItem
					{
						FamilyName = familyName.Trim(),
						CategoryName = categoryName.Trim(),
						CategoryId = string.Empty,
						CategoryGroup = FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(string.Empty, categoryName, string.Empty, familyName),
						IsShared = true
					});
				}
				if (!string.IsNullOrWhiteSpace(typeName) && !item.TypeNames.Any((string x) => string.Equals(x, typeName, StringComparison.OrdinalIgnoreCase)))
				{
					item.TypeNames.Add(typeName.Trim());
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		foreach (StandardNestedLoadableFamilySnapshotItem value in result.Values)
		{
			value.TypeNames = CloneStringList(value.TypeNames);
			value.TypeCount = value.TypeNames.Count;
		}
		return result.Values.OrderBy<StandardNestedLoadableFamilySnapshotItem, string>((StandardNestedLoadableFamilySnapshotItem x) => Normalize(x.CategoryName), StringComparer.Ordinal).ThenBy<StandardNestedLoadableFamilySnapshotItem, string>((StandardNestedLoadableFamilySnapshotItem x) => Normalize(x.FamilyName), StringComparer.Ordinal).ToList();
	}

	private static string BuildNestedLoadableSignatureKey(string categoryName, string familyName)
	{
		string familyKey = Normalize(familyName);
		if (string.IsNullOrWhiteSpace(familyKey))
		{
			return string.Empty;
		}
		return Normalize(categoryName) + "|" + familyKey;
	}

	private static string ExtractSignatureDebugQuotedValue(string value, string key)
	{
		if (string.IsNullOrWhiteSpace(value) || string.IsNullOrWhiteSpace(key))
		{
			return string.Empty;
		}
		string marker = key + "=\"";
		int start = value.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
		if (start < 0)
		{
			return string.Empty;
		}
		start += marker.Length;
		int end = value.IndexOf('"', start);
		if (end < 0)
		{
			end = value.Length;
		}
		if (end <= start)
		{
			return string.Empty;
		}
		return value.Substring(start, end - start).Trim();
	}

	private static List<string> SplitSignatureTokens(string value)
	{
		string text = value ?? string.Empty;
		int bracketIndex = text.IndexOf("[", StringComparison.Ordinal);
		if (bracketIndex >= 0)
		{
			text = text.Substring(0, bracketIndex);
		}
		return (from x in text.Split('|')
			select (x ?? string.Empty).Trim()).ToList();
	}

	private static string GetSignatureToken(IList<string> tokens, int index)
	{
		if (tokens == null || index < 0 || index >= tokens.Count)
		{
			return string.Empty;
		}
		return tokens[index];
	}

	private static string FirstNonEmptyToken(params string[] values)
	{
		if (values == null)
		{
			return string.Empty;
		}
		foreach (string value in values)
		{
			if (!string.IsNullOrWhiteSpace(value))
			{
				return value.Trim();
			}
		}
		return string.Empty;
	}

	private static List<string> ExtractFamilyManagerTypeNamesFromSignature(string signaturePath)
	{
		List<string> result = new List<string>();
		List<string> ExtractFamilyManagerTypeNamesFromSignature;
		if (string.IsNullOrWhiteSpace(signaturePath) || !File.Exists(signaturePath))
		{
			ExtractFamilyManagerTypeNamesFromSignature = result;
		}
		else
		{
			try
			{
				string typeLine = (from line in File.ReadLines(signaturePath)
					where line != null
					select line.Trim() into line
					where line.StartsWith("types=", StringComparison.OrdinalIgnoreCase)
					where !line.StartsWith("types=params=", StringComparison.OrdinalIgnoreCase)
					select line).LastOrDefault();
				if (string.IsNullOrWhiteSpace(typeLine))
				{
					ExtractFamilyManagerTypeNamesFromSignature = result;
					goto IL_019b;
				}
				result = (from x in typeLine.Substring("types=".Length).Split(';')
					select (x ?? string.Empty).Trim() into x
					where !string.IsNullOrWhiteSpace(x)
					select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ExtractFamilyManagerTypeNamesFromSignature = new List<string>();
				ProjectData.ClearProjectError();
				goto IL_019b;
			}
			ExtractFamilyManagerTypeNamesFromSignature = result;
		}
		goto IL_019b;
		IL_019b:
		return ExtractFamilyManagerTypeNamesFromSignature;
	}

	private static string BuildParameterDifferenceSignature(IEnumerable<StandardFamilyParameterSnapshotItem> parameters)
	{
		return FamilyParameterSnapshotNormalizationService.BuildComparableDefinitionSignature(parameters);
	}

	private static string ResolvePortableParameterIdentity(StandardFamilyParameterSnapshotItem parameter)
	{
		if (parameter == null)
		{
			return string.Empty;
		}
		if (!string.IsNullOrWhiteSpace(parameter.ExternalGuid))
		{
			return "guid:" + parameter.ExternalGuid.Trim();
		}
		if (!string.IsNullOrWhiteSpace(parameter.ParameterId) && parameter.ParameterId.Trim().StartsWith("-", StringComparison.Ordinal))
		{
			return "builtin:" + parameter.ParameterId.Trim();
		}
		return string.Empty;
	}

	private static List<SystemTypeComparisonItem> BuildSystemTypeComparisons(StandardLibrarySnapshot standardSnapshot, ProjectContentSnapshot projectSnapshot, ProjectTrackingCatalog trackingCatalog, bool compareDetailedSystemTypeComponents)
	{
		List<SystemTypeComparisonItem> results = new List<SystemTypeComparisonItem>();
		Dictionary<string, ProjectSystemTypeSnapshotItem> projectMap = BuildFirstMap(projectSnapshot.SystemTypes, [SpecialName] (ProjectSystemTypeSnapshotItem x) => BuildSystemMatchKey(x));
		Dictionary<string, List<ProjectSystemTypeSnapshotItem>> projectNameMap = BuildGroupedMap(projectSnapshot.SystemTypes, [SpecialName] (ProjectSystemTypeSnapshotItem x) => Normalize(x.TypeClassName) + "|" + Normalize(x.TypeName));
		Dictionary<string, List<StandardSystemTypeSnapshotItem>> standardNameMap = BuildGroupedMap(standardSnapshot.SystemTypes, [SpecialName] (StandardSystemTypeSnapshotItem x) => Normalize(x.TypeClassName) + "|" + Normalize(x.TypeName));
		HashSet<string> matchedProjectKeys = new HashSet<string>(StringComparer.Ordinal);
		Dictionary<string, TrackedSystemTypeState> trackingMap = BuildTrackedSystemMap(trackingCatalog);
		foreach (StandardSystemTypeSnapshotItem standardType in standardSnapshot.SystemTypes.OrderBy<StandardSystemTypeSnapshotItem, string>([SpecialName] (StandardSystemTypeSnapshotItem x) => BuildSystemMatchKey(x), StringComparer.Ordinal))
		{
			string key = BuildSystemKey(standardType.TypeClassName, standardType.CategoryName, standardType.TypeName);
			string matchKey = BuildSystemMatchKey(standardType);
			ProjectSystemTypeSnapshotItem projectType = null;
			projectMap.TryGetValue(matchKey, out projectType);
			if (projectType == null)
			{
				projectType = FindProjectSystemTypeCategoryMatch(standardType, projectNameMap);
			}
			if (projectType != null)
			{
				matchedProjectKeys.Add(BuildSystemMatchKey(projectType));
				bool detailedComponentsRequested = compareDetailedSystemTypeComponents && SystemTypeDetailedComponentSnapshotService.SupportsDetailedComponents(standardType.TypeClassName);
				bool detailedComponentsReady = detailedComponentsRequested && standardType.DetailedComponentsCaptured && projectType.DetailedComponentsCaptured;
				string standardFingerprint = ProjectSnapshotFingerprintService.BuildSystemFingerprint(standardType, detailedComponentsReady);
				string projectFingerprint = ProjectSnapshotFingerprintService.BuildSystemFingerprint(projectType, detailedComponentsReady);
				TrackedSystemTypeState trackedState = null;
				trackingMap.TryGetValue(key, out trackedState);
				ComparisonClassificationResult classification = ClassifyTrackedStatus(trackedState?.ApprovedFingerprint, trackedState?.ApprovedStandardStamp, standardFingerprint, standardSnapshot.CapturedAtUtc, projectFingerprint);
				List<string> differenceSummary = BuildSystemTypeDifferenceSummary(standardType, projectType, standardFingerprint, projectFingerprint, detailedComponentsRequested, detailedComponentsReady);
				classification = OverrideSystemClassificationForRoutingDependencyDifferences(classification, differenceSummary);
				classification = OverrideSystemClassificationForDetailedComponentDifferences(classification, differenceSummary);
				classification = OverrideSystemClassificationForCurtainPanelDifferences(classification, differenceSummary);
				results.Add(new SystemTypeComparisonItem
				{
					IdentityKey = key,
					TypeName = standardType.TypeName,
					CategoryName = standardType.CategoryName,
					TypeClassName = standardType.TypeClassName,
					Status = classification.Status,
					StandardFingerprint = standardFingerprint,
					ProjectFingerprint = projectFingerprint,
					DifferenceSummary = differenceSummary,
					SupportsRoutingDependencies = standardType.SupportsRoutingDependencies,
					StandardRoutingPreferenceSignature = standardType.RoutingPreferenceSignature ?? string.Empty,
					ProjectRoutingPreferenceSignature = projectType.RoutingPreferenceSignature ?? string.Empty,
					Notes = CombineSystemTypeNotes(classification.Notes, differenceSummary),
					DetailSummary = BuildSystemTypeDetailSummary(standardType, projectType, compareDetailedSystemTypeComponents),
					Layers = CloneSystemTypeLayers(standardType.Layers)
				});
				continue;
			}
			List<ProjectSystemTypeSnapshotItem> categoryMismatches = FindProjectSystemTypeCategoryMismatches(standardType, projectNameMap);
			if (categoryMismatches.Count > 0)
			{
				foreach (ProjectSystemTypeSnapshotItem mismatch in categoryMismatches)
				{
					matchedProjectKeys.Add(BuildSystemMatchKey(mismatch));
				}
				results.Add(new SystemTypeComparisonItem
				{
					IdentityKey = key,
					TypeName = standardType.TypeName,
					CategoryName = standardType.CategoryName,
					TypeClassName = standardType.TypeClassName,
					Status = "CategoryMismatch",
					SupportsRoutingDependencies = standardType.SupportsRoutingDependencies,
					StandardRoutingPreferenceSignature = standardType.RoutingPreferenceSignature ?? string.Empty,
					ProjectRoutingPreferenceSignature = categoryMismatches.FirstOrDefault()?.RoutingPreferenceSignature ?? string.Empty,
					Notes = BuildSystemTypeCategoryMismatchNote(standardType, categoryMismatches),
					DetailSummary = BuildSystemTypeDetailSummary(standardType, categoryMismatches.FirstOrDefault(), compareDetailedSystemTypeComponents),
					Layers = CloneSystemTypeLayers(standardType.Layers)
				});
			}
			else
			{
				results.Add(new SystemTypeComparisonItem
				{
					IdentityKey = key,
					TypeName = standardType.TypeName,
					CategoryName = standardType.CategoryName,
					TypeClassName = standardType.TypeClassName,
					Status = "LoadAvailable",
					StandardFingerprint = ProjectSnapshotFingerprintService.BuildSystemFingerprint(standardType, ShouldUseDetailedComponents(standardType, compareDetailedSystemTypeComponents)),
					SupportsRoutingDependencies = standardType.SupportsRoutingDependencies,
					StandardRoutingPreferenceSignature = standardType.RoutingPreferenceSignature ?? string.Empty,
					Notes = "Not loaded in current project.",
					DetailSummary = BuildSystemTypeDetailSummary(standardType, null, compareDetailedSystemTypeComponents),
					Layers = CloneSystemTypeLayers(standardType.Layers)
				});
			}
		}
		foreach (ProjectSystemTypeSnapshotItem projectType2 in projectSnapshot.SystemTypes.OrderBy<ProjectSystemTypeSnapshotItem, string>([SpecialName] (ProjectSystemTypeSnapshotItem x) => BuildSystemMatchKey(x), StringComparer.Ordinal))
		{
			string key2 = BuildSystemKey(projectType2.TypeClassName, projectType2.CategoryName, projectType2.TypeName);
			string matchKey2 = BuildSystemMatchKey(projectType2);
			if (!matchedProjectKeys.Contains(matchKey2))
			{
				List<StandardSystemTypeSnapshotItem> standardCategoryMismatches = FindStandardSystemTypeCategoryMismatches(projectType2, standardNameMap);
				if (standardCategoryMismatches.Count > 0)
				{
					results.Add(new SystemTypeComparisonItem
					{
						IdentityKey = key2,
						TypeName = projectType2.TypeName,
						CategoryName = projectType2.CategoryName,
						TypeClassName = projectType2.TypeClassName,
						Status = "CategoryMismatch",
						SupportsRoutingDependencies = projectType2.SupportsRoutingDependencies,
						StandardRoutingPreferenceSignature = standardCategoryMismatches.FirstOrDefault()?.RoutingPreferenceSignature ?? string.Empty,
						ProjectRoutingPreferenceSignature = projectType2.RoutingPreferenceSignature ?? string.Empty,
						Notes = BuildProjectSystemTypeCategoryMismatchNote(projectType2, standardCategoryMismatches),
						DetailSummary = BuildSystemTypeDetailSummary(standardCategoryMismatches.FirstOrDefault(), projectType2, compareDetailedSystemTypeComponents),
						Layers = CloneSystemTypeLayers(standardCategoryMismatches.FirstOrDefault()?.Layers)
					});
				}
				else
				{
					results.Add(new SystemTypeComparisonItem
					{
						IdentityKey = key2,
						TypeName = projectType2.TypeName,
						CategoryName = projectType2.CategoryName,
						TypeClassName = projectType2.TypeClassName,
						Status = "ProjectOnly",
						ProjectFingerprint = ProjectSnapshotFingerprintService.BuildSystemFingerprint(projectType2, ShouldUseDetailedComponents(projectType2, compareDetailedSystemTypeComponents)),
						SupportsRoutingDependencies = projectType2.SupportsRoutingDependencies,
						ProjectRoutingPreferenceSignature = projectType2.RoutingPreferenceSignature ?? string.Empty,
						DetailSummary = BuildSystemTypeDetailSummary(null, projectType2, compareDetailedSystemTypeComponents),
						Notes = "Project system type was not found in the registered standard snapshot."
					});
				}
			}
		}
		return results;
	}

	private static string SelectSystemTypeDetailSummary(string standardDetailSummary, string projectDetailSummary)
	{
		if (!string.IsNullOrWhiteSpace(standardDetailSummary))
		{
			return standardDetailSummary;
		}
		return projectDetailSummary ?? string.Empty;
	}

	private static bool ShouldUseDetailedComponents(StandardSystemTypeSnapshotItem item, bool requested)
	{
		return requested && item != null && item.DetailedComponentsCaptured && SystemTypeDetailedComponentSnapshotService.SupportsDetailedComponents(item.TypeClassName);
	}

	private static bool ShouldUseDetailedComponents(ProjectSystemTypeSnapshotItem item, bool requested)
	{
		return requested && item != null && item.DetailedComponentsCaptured && SystemTypeDetailedComponentSnapshotService.SupportsDetailedComponents(item.TypeClassName);
	}

	private static string BuildSystemTypeDetailSummary(StandardSystemTypeSnapshotItem standardType, ProjectSystemTypeSnapshotItem projectType, bool compareDetailedComponents)
	{
		string detail = SelectSystemTypeDetailSummary(standardType?.DetailSummary, projectType?.DetailSummary);
		string typeClassName = standardType?.TypeClassName ?? projectType?.TypeClassName ?? string.Empty;
		bool supportsOptionalComponents = compareDetailedComponents && SystemTypeDetailedComponentSnapshotService.SupportsDetailedComponents(typeClassName);
		bool hasRequiredCurtainPanelComponents = SystemTypeDetailedComponentSnapshotService.HasRequiredCurtainPanelComponents(standardType?.DetailedComponents) || SystemTypeDetailedComponentSnapshotService.HasRequiredCurtainPanelComponents(projectType?.DetailedComponents);
		if (!supportsOptionalComponents && !hasRequiredCurtainPanelComponents)
		{
			return detail;
		}
		List<string> lines = new List<string>();
		if (!string.IsNullOrWhiteSpace(detail))
		{
			lines.Add(detail.TrimEnd());
		}
		if (hasRequiredCurtainPanelComponents)
		{
			lines.Add("@section\tcurtain-component-differences");
			if (standardType != null && projectType != null)
			{
				if (!standardType.DetailedComponentsCaptured || !projectType.DetailedComponentsCaptured)
				{
					lines.Add("@row\tcurtain-component-differences\tScan status\tA new precise scan is required before curtain panel dependency comparison.");
				}
				else
				{
					List<string> curtainRows = BuildDetailedComponentDifferenceRows(
						(standardType.DetailedComponents ?? new List<SystemTypeDetailedComponentSnapshotItem>()).Where(SystemTypeDetailedComponentSnapshotService.IsRequiredCurtainPanelComponent),
						(projectType.DetailedComponents ?? new List<SystemTypeDetailedComponentSnapshotItem>()).Where(SystemTypeDetailedComponentSnapshotService.IsRequiredCurtainPanelComponent),
						100,
						"curtain-component-differences");
					if (curtainRows.Count == 0)
					{
						lines.Add("@row\tcurtain-component-differences\tComparison\tMatches standard");
					}
					else
					{
						lines.AddRange(curtainRows);
					}
				}
			}
		}
		if (!supportsOptionalComponents)
		{
			return string.Join(Environment.NewLine, lines);
		}
		lines.Add("@section\tcomponent-differences");
		if (standardType == null || projectType == null)
		{
			return string.Join("\r\n", lines);
		}
		if (!standardType.DetailedComponentsCaptured || !projectType.DetailedComponentsCaptured)
		{
			lines.Add("@row\tcomponent-differences\tScan status\tA new precise scan is required before detailed component comparison.");
			return string.Join(Environment.NewLine, lines);
		}
		List<string> rows = BuildDetailedComponentDifferenceRows(
			(standardType.DetailedComponents ?? new List<SystemTypeDetailedComponentSnapshotItem>()).Where(x => !SystemTypeDetailedComponentSnapshotService.IsRequiredCurtainPanelComponent(x)),
			(projectType.DetailedComponents ?? new List<SystemTypeDetailedComponentSnapshotItem>()).Where(x => !SystemTypeDetailedComponentSnapshotService.IsRequiredCurtainPanelComponent(x)),
			100);
		if (rows.Count == 0)
		{
			lines.Add("@row\tcomponent-differences\tComparison\tMatches standard");
		}
		else
		{
			lines.AddRange(rows);
		}
		return string.Join(Environment.NewLine, lines);
	}

	private static List<string> BuildDetailedComponentDifferenceRows(IEnumerable<SystemTypeDetailedComponentSnapshotItem> standardItems, IEnumerable<SystemTypeDetailedComponentSnapshotItem> projectItems, int limit, string sectionName = "component-differences")
	{
		Dictionary<string, SystemTypeDetailedComponentSnapshotItem> standardMap = (standardItems ?? Enumerable.Empty<SystemTypeDetailedComponentSnapshotItem>()).Where(x => x != null).GroupBy(SystemTypeDetailedComponentSnapshotService.BuildIdentityKey, StringComparer.Ordinal).ToDictionary(x => x.Key, x => x.First(), StringComparer.Ordinal);
		Dictionary<string, SystemTypeDetailedComponentSnapshotItem> projectMap = (projectItems ?? Enumerable.Empty<SystemTypeDetailedComponentSnapshotItem>()).Where(x => x != null).GroupBy(SystemTypeDetailedComponentSnapshotService.BuildIdentityKey, StringComparer.Ordinal).ToDictionary(x => x.Key, x => x.First(), StringComparer.Ordinal);
		List<string> output = new List<string>();
		int omitted = 0;
		int differenceCount = 0;
		foreach (string key in standardMap.Keys.Union(projectMap.Keys, StringComparer.Ordinal).OrderBy(x => x, StringComparer.Ordinal))
		{
			SystemTypeDetailedComponentSnapshotItem standard = standardMap.ContainsKey(key) ? standardMap[key] : null;
			SystemTypeDetailedComponentSnapshotItem project = projectMap.ContainsKey(key) ? projectMap[key] : null;
			if (standard != null && project != null && string.Equals(SystemTypeDetailedComponentSnapshotService.BuildComparableValue(standard), SystemTypeDetailedComponentSnapshotService.BuildComparableValue(project), StringComparison.Ordinal))
			{
				continue;
			}
			if (differenceCount >= Math.Max(1, limit))
			{
				omitted++;
				continue;
			}
			differenceCount++;
			string role = standard?.RoleName ?? project?.RoleName ?? key;
			string path = standard?.Path ?? project?.Path ?? string.Empty;
			string standardValue = standard == null ? "Missing" : SystemTypeDetailedComponentSnapshotService.BuildDisplayValue(standard);
			string projectValue = project == null ? "Missing" : SystemTypeDetailedComponentSnapshotService.BuildDisplayValue(project);
			output.Add(BuildStructuredDetailedComponentDifferenceRow(sectionName, role, path, standard, project));
			output.Add("@row\t" + CleanSystemDetailValue(sectionName) + "\t" + CleanSystemDetailValue(role + " · " + path) + "\t" + CleanSystemDetailValue("Standard: " + standardValue + " | Current: " + projectValue));
		}
		if (omitted > 0)
		{
			output.Add("@row\t" + CleanSystemDetailValue(sectionName) + "\tMore differences\t" + omitted.ToString(CultureInfo.InvariantCulture) + " additional item(s)");
		}
		return output;
	}

	private static string BuildStructuredDetailedComponentDifferenceRow(string sectionName, string role, string path, SystemTypeDetailedComponentSnapshotItem standard, SystemTypeDetailedComponentSnapshotItem project)
	{
		return string.Join("\t", new[]
		{
			"@component-diff",
			CleanSystemDetailValue(sectionName),
			CleanSystemDetailValue(role),
			CleanSystemDetailValue(standard?.ValueKind ?? "Missing"),
			CleanSystemDetailValue(standard?.RawValue),
			CleanSystemDetailValue(standard == null ? "Missing" : SystemTypeDetailedComponentSnapshotService.BuildDisplayValue(standard)),
			CleanSystemDetailValue(project?.ValueKind ?? "Missing"),
			CleanSystemDetailValue(project?.RawValue),
			CleanSystemDetailValue(project == null ? "Missing" : SystemTypeDetailedComponentSnapshotService.BuildDisplayValue(project)),
			CleanSystemDetailValue(path)
		});
	}

	private static string CleanSystemDetailValue(string value)
	{
		return (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim();
	}

	private static List<string> BuildSystemTypeDifferenceSummary(StandardSystemTypeSnapshotItem standardType, ProjectSystemTypeSnapshotItem projectType, string standardFingerprint, string projectFingerprint, bool detailedComponentsRequested, bool detailedComponentsReady)
	{
		List<string> result = new List<string>();
		if (standardType == null || projectType == null)
		{
			return result;
		}
		List<string> dependencyDifferences = BuildRoutingDependencyFingerprintDifferenceSummary(standardType.RoutingPreferenceSignature, projectType.RoutingPreferenceSignature);
		bool requiredCurtainPanelComparison = SystemTypeDetailedComponentSnapshotService.HasRequiredCurtainPanelComponents(standardType.DetailedComponents) || SystemTypeDetailedComponentSnapshotService.HasRequiredCurtainPanelComponents(projectType.DetailedComponents);
		bool requiredCurtainPanelReady = requiredCurtainPanelComparison && standardType.DetailedComponentsCaptured && projectType.DetailedComponentsCaptured;
		if (detailedComponentsRequested && !detailedComponentsReady)
		{
			result.Add("Detailed component comparison requires a new precise scan");
		}
		if (requiredCurtainPanelComparison && !requiredCurtainPanelReady)
		{
			result.Add("Curtain panel dependency comparison requires a new precise scan");
		}
		if (string.Equals(standardFingerprint ?? string.Empty, projectFingerprint ?? string.Empty, StringComparison.OrdinalIgnoreCase))
		{
			return result;
		}
		AddSystemValueDifference(result, "Classification", "분류", standardType.ClassificationCode, projectType.ClassificationCode);
		AddSystemValueDifference(result, "Segment", "Segment", standardType.SegmentName, projectType.SegmentName);
		AddSystemValueDifference(result, "Material", "재질", standardType.MaterialName, projectType.MaterialName);
		AddSystemValueDifference(result, "Shape", "형상", standardType.Shape, projectType.Shape);
		if (!string.Equals(ProjectSnapshotFingerprintService.NormalizeRoutingPreferenceSignature(standardType.RoutingPreferenceSignature), ProjectSnapshotFingerprintService.NormalizeRoutingPreferenceSignature(projectType.RoutingPreferenceSignature), StringComparison.Ordinal))
		{
			List<string> routingDifferences = BuildRoutingPreferenceDifferenceSummary(standardType.RoutingPreferenceSignature, projectType.RoutingPreferenceSignature);
			if (routingDifferences.Count > 0)
			{
				foreach (string difference in routingDifferences)
				{
					result.Add(difference);
				}
			}
			else
			{
				result.Add("Routing Preference differs");
			}
		}
		using (List<string>.Enumerator enumerator2 = dependencyDifferences.GetEnumerator())
		{
			_Closure_0024__71_002D0 closure_0024__71_002D = default(_Closure_0024__71_002D0);
			while (enumerator2.MoveNext())
			{
				closure_0024__71_002D = new _Closure_0024__71_002D0(closure_0024__71_002D);
				closure_0024__71_002D._0024VB_0024Local_difference = enumerator2.Current;
				if (!result.Any(closure_0024__71_002D._Lambda_0024__0))
				{
					result.Add(closure_0024__71_002D._0024VB_0024Local_difference);
				}
			}
		}
		if (!string.Equals(Normalize(standardType.CompoundStructureSignature), Normalize(projectType.CompoundStructureSignature), StringComparison.Ordinal))
		{
			result.Add("Layer differs");
		}
		if (detailedComponentsReady)
		{
			result.AddRange(BuildDetailedComponentDifferenceSummary(
				(standardType.DetailedComponents ?? new List<SystemTypeDetailedComponentSnapshotItem>()).Where(x => !SystemTypeDetailedComponentSnapshotService.IsRequiredCurtainPanelComponent(x)),
				(projectType.DetailedComponents ?? new List<SystemTypeDetailedComponentSnapshotItem>()).Where(x => !SystemTypeDetailedComponentSnapshotService.IsRequiredCurtainPanelComponent(x))));
		}
		if (requiredCurtainPanelReady)
		{
			result.AddRange(BuildCurtainPanelDifferenceSummary(standardType.DetailedComponents, projectType.DetailedComponents));
		}
		if (result.Count == 0)
		{
			result.Add("System Type fingerprint differs");
		}
		return result;
	}

	private static List<string> BuildDetailedComponentDifferenceSummary(IEnumerable<SystemTypeDetailedComponentSnapshotItem> standardItems, IEnumerable<SystemTypeDetailedComponentSnapshotItem> projectItems)
	{
		List<string> rows = BuildDetailedComponentDifferenceRows(standardItems, projectItems, 8);
		return rows.Where(x => x.StartsWith("@component-diff\t", StringComparison.Ordinal)).Select(x =>
		{
			string[] parts = x.Split('\t');
			return parts.Length >= 4 ? "Detailed component differs: " + parts[2] : "Detailed component differs";
		}).ToList();
	}

	private static List<string> BuildCurtainPanelDifferenceSummary(IEnumerable<SystemTypeDetailedComponentSnapshotItem> standardItems, IEnumerable<SystemTypeDetailedComponentSnapshotItem> projectItems)
	{
		List<string> rows = BuildDetailedComponentDifferenceRows(
			(standardItems ?? Enumerable.Empty<SystemTypeDetailedComponentSnapshotItem>()).Where(SystemTypeDetailedComponentSnapshotService.IsRequiredCurtainPanelComponent),
			(projectItems ?? Enumerable.Empty<SystemTypeDetailedComponentSnapshotItem>()).Where(SystemTypeDetailedComponentSnapshotService.IsRequiredCurtainPanelComponent),
			8,
			"curtain-component-differences");
		return rows.Where(x => x.StartsWith("@component-diff\t", StringComparison.Ordinal)).Select(x =>
		{
			string[] parts = x.Split('\t');
			return parts.Length >= 4 ? "Curtain panel dependency differs: " + parts[2] : "Curtain panel dependency differs";
		}).ToList();
	}

	private static ComparisonClassificationResult OverrideSystemClassificationForDetailedComponentDifferences(ComparisonClassificationResult classification, IEnumerable<string> differenceSummary)
	{
		if (classification == null)
		{
			classification = new ComparisonClassificationResult();
		}
		List<string> differences = (differenceSummary ?? Enumerable.Empty<string>()).Where(x => Normalize(x).Contains("detailed component")).ToList();
		if (differences.Count == 0)
		{
			return classification;
		}
		if (differences.Any(x => Normalize(x).Contains("requires a new precise scan")))
		{
			return new ComparisonClassificationResult
			{
				Status = "ManualReview",
				Notes = "Detailed System Type components are enabled, but this Railing/Stair snapshot predates component capture. Run a new precise scan."
			};
		}
		string status = Normalize(classification.Status);
		if (status == "loadedlatest" || status == "stampnormalizationneeded")
		{
			return new ComparisonClassificationResult
			{
				Status = "UpdateAvailable",
				Notes = "Railing/Stair detailed component configuration differs from the current standard."
			};
		}
		return classification;
	}

	private static ComparisonClassificationResult OverrideSystemClassificationForCurtainPanelDifferences(ComparisonClassificationResult classification, IEnumerable<string> differenceSummary)
	{
		if (classification == null)
		{
			classification = new ComparisonClassificationResult();
		}
		List<string> differences = (differenceSummary ?? Enumerable.Empty<string>()).Where(x => Normalize(x).Contains("curtain panel dependency")).ToList();
		if (differences.Count == 0)
		{
			return classification;
		}
		if (differences.Any(x => Normalize(x).Contains("requires a new precise scan")))
		{
			return new ComparisonClassificationResult
			{
				Status = "ManualReview",
				Notes = "Curtain panel dependency data is mandatory, but this snapshot predates capture. Run a new precise scan."
			};
		}
		string status = Normalize(classification.Status);
		if (status == "loadedlatest" || status == "stampnormalizationneeded")
		{
			return new ComparisonClassificationResult
			{
				Status = "UpdateAvailable",
				Notes = "Curtain panel family/type dependency differs from the current standard."
			};
		}
		return classification;
	}

	private static ComparisonClassificationResult OverrideSystemClassificationForRoutingDependencyDifferences(ComparisonClassificationResult classification, IEnumerable<string> differenceSummary)
	{
		if (classification == null)
		{
			classification = new ComparisonClassificationResult();
		}
		if (!HasRoutingDependencyFingerprintDifference(differenceSummary))
		{
			return classification;
		}
		string left = Normalize(classification.Status);
		if (Operators.CompareString(left, "loadedlatest", TextCompare: false) == 0 || Operators.CompareString(left, "stampnormalizationneeded", TextCompare: false) == 0)
		{
			return new ComparisonClassificationResult
			{
				Status = "UpdateAvailable",
				Notes = "Routing dependency family/type fingerprint differs from the current standard; reload the standard dependency families before applying this system type."
			};
		}
		return classification;
	}

	private static bool HasRoutingDependencyFingerprintDifference(IEnumerable<string> differenceSummary)
	{
		return differenceSummary?.Any([SpecialName] (string x) => Normalize(x).Contains("routingdependencyfingerprint")) ?? false;
	}

	private static List<string> BuildRoutingDependencyFingerprintDifferenceSummary(string standardSignature, string projectSignature)
	{
		List<string> result = new List<string>();
		Dictionary<string, RoutingPreferenceSignatureRule> standardRules = BuildRoutingPreferenceRuleMap(standardSignature);
		Dictionary<string, RoutingPreferenceSignatureRule> projectRules = BuildRoutingPreferenceRuleMap(projectSignature);
		List<string> commonKeys = standardRules.Keys.Intersect<string>(projectRules.Keys, StringComparer.Ordinal).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal).ToList();
		foreach (string key in commonKeys)
		{
			RoutingPreferenceSignatureRule standardRule = standardRules[key];
			RoutingPreferenceSignatureRule projectRule = projectRules[key];
			if (standardRule != null && projectRule != null && string.Equals(FormatRoutingPreferencePart(standardRule), FormatRoutingPreferencePart(projectRule), StringComparison.Ordinal) && !RoutingPreferencePartFingerprintMatches(standardRule, projectRule))
			{
				result.Add(BuildRoutingDependencyFingerprintDifferenceLine(standardRule, projectRule));
				if (result.Count >= 10)
				{
					break;
				}
			}
		}
		if (commonKeys.Count > 10 && result.Count >= 10)
		{
			result.Add("RoutingDependencyFingerprint differs: standard additional rules=" + checked(commonKeys.Count - 10).ToString(CultureInfo.InvariantCulture) + " / project see detailed report");
		}
		return result;
	}

	private static string BuildRoutingDependencyFingerprintDifferenceLine(RoutingPreferenceSignatureRule standardRule, RoutingPreferenceSignatureRule projectRule)
	{
		string standardValue = FormatRoutingPreferenceRuleLabel(standardRule) + " " + FormatRoutingPreferencePart(standardRule) + " " + FormatRoutingPreferenceFingerprintBrief(standardRule);
		string projectValue = FormatRoutingPreferenceRuleLabel(projectRule) + " " + FormatRoutingPreferencePart(projectRule) + " " + FormatRoutingPreferenceFingerprintBrief(projectRule);
		return "RoutingDependencyFingerprint differs: standard " + DisplaySystemDiffValue(standardValue) + " / project " + DisplaySystemDiffValue(projectValue);
	}

	private static List<string> BuildRoutingPreferenceDifferenceSummary(string standardSignature, string projectSignature)
	{
		List<string> result = new List<string>();
		Dictionary<string, RoutingPreferenceSignatureRule> standardRules = BuildRoutingPreferenceRuleMap(standardSignature);
		Dictionary<string, RoutingPreferenceSignatureRule> projectRules = BuildRoutingPreferenceRuleMap(projectSignature);
		List<string> allKeys = standardRules.Keys.Union<string>(projectRules.Keys, StringComparer.Ordinal).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.Ordinal).ToList();
		foreach (string key in allKeys)
		{
			RoutingPreferenceSignatureRule standardRule = null;
			RoutingPreferenceSignatureRule projectRule = null;
			standardRules.TryGetValue(key, out standardRule);
			projectRules.TryGetValue(key, out projectRule);
			if (standardRule == null)
			{
				result.Add(BuildRoutingPreferenceDifferenceLine("Routing Preference", "(none)", FormatRoutingPreferenceRule(projectRule)));
			}
			else if (projectRule == null)
			{
				result.Add(BuildRoutingPreferenceDifferenceLine("Routing Preference", FormatRoutingPreferenceRule(standardRule), "(none)"));
			}
			else
			{
				if (!string.Equals(FormatRoutingPreferencePart(standardRule), FormatRoutingPreferencePart(projectRule), StringComparison.Ordinal))
				{
					result.Add(BuildRoutingPreferenceDifferenceLine("Routing Preference", FormatRoutingPreferenceRule(standardRule), FormatRoutingPreferenceRule(projectRule)));
				}
				if (!string.Equals(Normalize(standardRule.CriteriaSignature), Normalize(projectRule.CriteriaSignature), StringComparison.Ordinal))
				{
					result.Add(BuildRoutingPreferenceDifferenceLine("Routing Criteria", FormatRoutingPreferenceRuleLabel(standardRule) + " " + DisplaySystemDiffValue(standardRule.CriteriaSignature), FormatRoutingPreferenceRuleLabel(projectRule) + " " + DisplaySystemDiffValue(projectRule.CriteriaSignature)));
				}
				if (!RoutingPreferencePartFingerprintMatches(standardRule, projectRule) && string.Equals(FormatRoutingPreferencePart(standardRule), FormatRoutingPreferencePart(projectRule), StringComparison.Ordinal))
				{
					result.Add(BuildRoutingPreferenceDifferenceLine("Routing Part Fingerprint", FormatRoutingPreferenceRuleLabel(standardRule) + " " + FormatRoutingPreferenceFingerprintBrief(standardRule), FormatRoutingPreferenceRuleLabel(projectRule) + " " + FormatRoutingPreferenceFingerprintBrief(projectRule)));
				}
			}
			if (result.Count >= 10)
			{
				break;
			}
		}
		if (allKeys.Count > 10)
		{
			result.Add("Routing Preference differs: standard additional differences=" + checked(allKeys.Count - 10).ToString(CultureInfo.InvariantCulture) + " / project see detailed report");
		}
		return result;
	}

	private static Dictionary<string, RoutingPreferenceSignatureRule> BuildRoutingPreferenceRuleMap(string signature)
	{
		Dictionary<string, RoutingPreferenceSignatureRule> result = new Dictionary<string, RoutingPreferenceSignatureRule>(StringComparer.Ordinal);
		if (string.IsNullOrWhiteSpace(signature))
		{
			return result;
		}
		string[] array = signature.Replace("\r", "\n").Split(new string[1] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
		for (int i = 0; i < array.Length; i = checked(i + 1))
		{
			RoutingPreferenceSignatureRule rule = ParseRoutingPreferenceSignatureRule(array[i]);
			if (rule != null)
			{
				result[rule.RuleKey] = rule;
			}
		}
		return result;
	}

	private static RoutingPreferenceSignatureRule ParseRoutingPreferenceSignatureRule(string line)
	{
		if (string.IsNullOrWhiteSpace(line))
		{
			return null;
		}
		string[] parts = line.Split('|');
		if (parts.Length < 2)
		{
			return null;
		}
		RoutingPreferenceSignatureRule rule = new RoutingPreferenceSignatureRule
		{
			GroupName = parts[0].Trim(),
			RuleIndex = parts[1].Trim()
		};
		if (parts.Length > 2)
		{
			rule.PartClass = parts[2].Trim();
		}
		if (parts.Length > 3)
		{
			rule.PartCategory = parts[3].Trim();
		}
		if (parts.Length > 4)
		{
			rule.FamilyKey = parts[4].Trim();
		}
		if (parts.Length > 5)
		{
			rule.FamilyName = parts[5].Trim();
		}
		if (parts.Length > 6)
		{
			rule.TypeName = parts[6].Trim();
		}
		if (parts.Length > 7)
		{
			rule.FamilyFingerprint = parts[7].Trim();
		}
		if (parts.Length > 8)
		{
			rule.TypeFingerprint = parts[8].Trim();
		}
		if (parts.Length > 9)
		{
			rule.PartFingerprint = parts[9].Trim();
		}
		if (parts.Length > 10)
		{
			rule.CriteriaSignature = string.Join("|", parts.Skip(10)).Trim();
		}
		return rule;
	}

	private static string BuildRoutingPreferenceDifferenceLine(string label, string standardValue, string projectValue)
	{
		return label + " differs: standard " + DisplaySystemDiffValue(standardValue) + " / project " + DisplaySystemDiffValue(projectValue);
	}

	private static string FormatRoutingPreferenceRule(RoutingPreferenceSignatureRule rule)
	{
		if (rule == null)
		{
			return "(none)";
		}
		return FormatRoutingPreferenceRuleLabel(rule) + " " + FormatRoutingPreferencePart(rule);
	}

	private static string FormatRoutingPreferenceRuleLabel(RoutingPreferenceSignatureRule rule)
	{
		if (rule == null)
		{
			return string.Empty;
		}
		return (string.IsNullOrWhiteSpace(rule.GroupName) ? "rule" : rule.GroupName) + "[" + (string.IsNullOrWhiteSpace(rule.RuleIndex) ? "?" : rule.RuleIndex) + "]";
	}

	private static string FormatRoutingPreferencePart(RoutingPreferenceSignatureRule rule)
	{
		if (rule == null)
		{
			return "(none)";
		}
		List<string> parts = new List<string>();
		if (!string.IsNullOrWhiteSpace(rule.FamilyName))
		{
			parts.Add(rule.FamilyName);
		}
		if (!string.IsNullOrWhiteSpace(rule.TypeName))
		{
			parts.Add(rule.TypeName);
		}
		if (parts.Count == 0 && !string.IsNullOrWhiteSpace(rule.PartCategory))
		{
			parts.Add(rule.PartCategory);
		}
		if (parts.Count == 0 && !string.IsNullOrWhiteSpace(rule.PartClass))
		{
			parts.Add(rule.PartClass);
		}
		if (parts.Count == 0)
		{
			parts.Add("(none)");
		}
		return string.Join(" / ", parts);
	}

	private static bool RoutingPreferencePartFingerprintMatches(RoutingPreferenceSignatureRule standardRule, RoutingPreferenceSignatureRule projectRule)
	{
		if (standardRule == null || projectRule == null)
		{
			return standardRule == projectRule;
		}
		return string.Equals(Normalize(standardRule.FamilyFingerprint), Normalize(projectRule.FamilyFingerprint), StringComparison.Ordinal) && string.Equals(Normalize(standardRule.TypeFingerprint), Normalize(projectRule.TypeFingerprint), StringComparison.Ordinal) && string.Equals(Normalize(standardRule.PartFingerprint), Normalize(projectRule.PartFingerprint), StringComparison.Ordinal);
	}

	private static string FormatRoutingPreferenceFingerprintBrief(RoutingPreferenceSignatureRule rule)
	{
		if (rule == null)
		{
			return "(none)";
		}
		List<string> parts = new List<string>();
		if (!string.IsNullOrWhiteSpace(rule.FamilyFingerprint))
		{
			parts.Add("family=" + ShortDiagnosticValue(rule.FamilyFingerprint, 10));
		}
		if (!string.IsNullOrWhiteSpace(rule.TypeFingerprint))
		{
			parts.Add("type=" + ShortDiagnosticValue(rule.TypeFingerprint, 10));
		}
		if (!string.IsNullOrWhiteSpace(rule.PartFingerprint))
		{
			parts.Add("part=" + ShortDiagnosticValue(rule.PartFingerprint, 10));
		}
		if (parts.Count == 0)
		{
			return "(none)";
		}
		return string.Join(", ", parts);
	}

	private static void AddSystemValueDifference(ICollection<string> target, string englishLabel, string koreanLabel, string standardValue, string projectValue)
	{
		if (target != null && !string.Equals(Normalize(standardValue), Normalize(projectValue), StringComparison.Ordinal))
		{
			target.Add(englishLabel + " differs: standard " + DisplaySystemDiffValue(standardValue) + " / project " + DisplaySystemDiffValue(projectValue));
		}
	}

	private static string DisplaySystemDiffValue(string value)
	{
		if (string.IsNullOrWhiteSpace(value))
		{
			return "(none)";
		}
		string normalized = NormalizeMultiline(value);
		if (normalized.Length <= 80)
		{
			return normalized;
		}
		return normalized.Substring(0, 77) + "...";
	}

	private static string NormalizeMultiline(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ")
			.Trim();
	}

	private static string CombineSystemTypeNotes(string primaryNote, IEnumerable<string> differenceSummary)
	{
		List<string> parts = new List<string>();
		if (!string.IsNullOrWhiteSpace(primaryNote))
		{
			parts.Add(primaryNote.Trim());
		}
		List<string> diffs = (differenceSummary ?? Enumerable.Empty<string>()).Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)).Distinct<string>(StringComparer.OrdinalIgnoreCase).ToList();
		if (diffs.Count > 0)
		{
			parts.Add("차이: " + string.Join(" / ", diffs));
		}
		return string.Join(" | ", parts);
	}

	private static List<StandardSystemTypeLayerSnapshotItem> CloneSystemTypeLayers(IEnumerable<StandardSystemTypeLayerSnapshotItem> layers)
	{
		List<StandardSystemTypeLayerSnapshotItem> result = new List<StandardSystemTypeLayerSnapshotItem>();
		if (layers == null)
		{
			return result;
		}
		foreach (StandardSystemTypeLayerSnapshotItem layer in layers)
		{
			if (layer != null)
			{
				result.Add(new StandardSystemTypeLayerSnapshotItem
				{
					Index = layer.Index,
					FunctionName = (layer.FunctionName ?? string.Empty),
					MaterialName = (layer.MaterialName ?? string.Empty),
					ThicknessDisplay = (layer.ThicknessDisplay ?? string.Empty),
					ThicknessFeet = layer.ThicknessFeet,
					IsCore = layer.IsCore,
					IsStructuralMaterial = layer.IsStructuralMaterial,
					IsVariable = layer.IsVariable
				});
			}
		}
		return result;
	}

	private static List<ProjectLoadableFamilySnapshotItem> FindProjectLoadableCategoryMismatches(StandardLoadableFamilySnapshotItem standardFamily, Dictionary<string, List<ProjectLoadableFamilySnapshotItem>> projectNameMap)
	{
		if (standardFamily == null || projectNameMap == null)
		{
			return new List<ProjectLoadableFamilySnapshotItem>();
		}
		List<ProjectLoadableFamilySnapshotItem> candidates = null;
		if (!projectNameMap.TryGetValue(Normalize(standardFamily.FamilyName), out candidates) || candidates == null || candidates.Count == 0)
		{
			return new List<ProjectLoadableFamilySnapshotItem>();
		}
		return candidates.Where([SpecialName] (ProjectLoadableFamilySnapshotItem x) => !LoadableCategoryMatches(standardFamily, x)).OrderBy<ProjectLoadableFamilySnapshotItem, string>([SpecialName] (ProjectLoadableFamilySnapshotItem x) => BuildLoadableMatchKey(x), StringComparer.Ordinal).ToList();
	}

	private static ProjectLoadableFamilySnapshotItem FindProjectLoadableCategoryMatch(StandardLoadableFamilySnapshotItem standardFamily, Dictionary<string, List<ProjectLoadableFamilySnapshotItem>> projectNameMap)
	{
		if (standardFamily == null || projectNameMap == null)
		{
			return null;
		}
		List<ProjectLoadableFamilySnapshotItem> candidates = null;
		if (!projectNameMap.TryGetValue(Normalize(standardFamily.FamilyName), out candidates) || candidates == null || candidates.Count == 0)
		{
			return null;
		}
		return candidates.Where([SpecialName] (ProjectLoadableFamilySnapshotItem x) => LoadableCategoryMatches(standardFamily, x)).OrderBy<ProjectLoadableFamilySnapshotItem, string>([SpecialName] (ProjectLoadableFamilySnapshotItem x) => BuildLoadableMatchKey(x), StringComparer.Ordinal).FirstOrDefault();
	}

	private static List<StandardLoadableFamilySnapshotItem> FindStandardLoadableCategoryMismatches(ProjectLoadableFamilySnapshotItem projectFamily, Dictionary<string, List<StandardLoadableFamilySnapshotItem>> standardNameMap)
	{
		if (projectFamily == null || standardNameMap == null)
		{
			return new List<StandardLoadableFamilySnapshotItem>();
		}
		List<StandardLoadableFamilySnapshotItem> candidates = null;
		if (!standardNameMap.TryGetValue(Normalize(projectFamily.FamilyName), out candidates) || candidates == null || candidates.Count == 0)
		{
			return new List<StandardLoadableFamilySnapshotItem>();
		}
		return candidates.Where([SpecialName] (StandardLoadableFamilySnapshotItem x) => !LoadableCategoryMatches(x, projectFamily)).OrderBy<StandardLoadableFamilySnapshotItem, string>([SpecialName] (StandardLoadableFamilySnapshotItem x) => BuildLoadableMatchKey(x), StringComparer.Ordinal).ToList();
	}

	private static string BuildLoadableCategoryMismatchNote(StandardLoadableFamilySnapshotItem standardFamily, List<ProjectLoadableFamilySnapshotItem> projectFamilies)
	{
		List<string> projectCategories = (from x in projectFamilies ?? new List<ProjectLoadableFamilySnapshotItem>()
			select x.CategoryName ?? string.Empty into x
			where !string.IsNullOrWhiteSpace(x)
			select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
		return "Family name matches a project family, but the category differs. Standard category: " + (standardFamily?.CategoryName ?? string.Empty) + " / Project category: " + ((projectCategories.Count == 0) ? string.Empty : string.Join(", ", projectCategories)) + ". Treat this as a category mismatch and review it before loading or overwriting.";
	}

	private static string BuildProjectLoadableCategoryMismatchNote(ProjectLoadableFamilySnapshotItem projectFamily, List<StandardLoadableFamilySnapshotItem> standardFamilies)
	{
		List<string> standardCategories = (from x in standardFamilies ?? new List<StandardLoadableFamilySnapshotItem>()
			select x.CategoryName ?? string.Empty into x
			where !string.IsNullOrWhiteSpace(x)
			select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
		return "Project family name exists in the standard, but the category differs. Standard category: " + ((standardCategories.Count == 0) ? string.Empty : string.Join(", ", standardCategories)) + " / Project category: " + (projectFamily?.CategoryName ?? string.Empty) + ". Treat this as a category mismatch and review it before loading or overwriting.";
	}

	private static bool IsHiddenNestedLoadableChild(StandardLoadableFamilySnapshotItem item, HashSet<string> nestedFamilyNames)
	{
		if (item == null || !item.IsNestedLoadableChild)
		{
			if (item == null || nestedFamilyNames == null)
			{
				return false;
			}
			if (!nestedFamilyNames.Contains(Normalize(item.FamilyName)))
			{
				return false;
			}
		}
		return true;
	}

	private static bool IsHiddenNestedLoadableChild(ProjectLoadableFamilySnapshotItem item, HashSet<string> nestedFamilyNames)
	{
		if (item == null || nestedFamilyNames == null)
		{
			return false;
		}
		return nestedFamilyNames.Contains(Normalize(item.FamilyName));
	}

	private static HashSet<string> BuildNestedLoadableNameSet(StandardLibrarySnapshot snapshot)
	{
		HashSet<string> result = new HashSet<string>(StringComparer.Ordinal);
		if (snapshot == null || snapshot.LoadableFamilies == null)
		{
			return result;
		}
		foreach (StandardLoadableFamilySnapshotItem parentItem in snapshot.LoadableFamilies)
		{
			if (parentItem == null)
			{
				continue;
			}
			foreach (StandardNestedLoadableFamilySnapshotItem signatureChild in BuildNestedLoadableFamiliesFromSignature(parentItem.ContentSignatureDebugPath, parentItem.FamilyName))
			{
				if (signatureChild != null && !string.IsNullOrWhiteSpace(signatureChild.FamilyName))
				{
					string signatureFamilyName = Normalize(signatureChild.FamilyName);
					if (!string.IsNullOrWhiteSpace(signatureFamilyName))
					{
						result.Add(signatureFamilyName);
					}
				}
			}
			if (parentItem.IsNestedLoadableChild)
			{
				string nestedFamilyName = Normalize(parentItem.FamilyName);
				if (!string.IsNullOrWhiteSpace(nestedFamilyName))
				{
					result.Add(nestedFamilyName);
				}
			}
			if (parentItem.NestedLoadableFamilies == null)
			{
				continue;
			}
			foreach (StandardNestedLoadableFamilySnapshotItem child in parentItem.NestedLoadableFamilies)
			{
				if (child != null && IsModelNestedLoadableChild(child))
				{
					string familyName = Normalize((child == null) ? string.Empty : child.FamilyName);
					if (!string.IsNullOrWhiteSpace(familyName))
					{
						result.Add(familyName);
					}
				}
			}
		}
		return result;
	}

	private static bool IsModelLoadableFamily(StandardLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return false;
		}
		return string.Equals(FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(item.CategoryGroup, item.CategoryName, item.CategoryId, item.FamilyName), "Model", StringComparison.OrdinalIgnoreCase);
	}

	private static bool IsModelLoadableFamily(ProjectLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return false;
		}
		return string.Equals(FamilyBrowserFamilyClassificationService.ResolveCategoryGroup(item.CategoryGroup, item.CategoryName, item.CategoryId, item.FamilyName), "Model", StringComparison.OrdinalIgnoreCase);
	}

	private static bool IsModelNestedLoadableChild(StandardNestedLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return false;
		}
		return IsNestedLoadableFamilyCandidate(item.CategoryGroup, item.CategoryName, item.CategoryId, item.FamilyName);
	}

	private static bool IsNestedLoadableFamilyCandidate(string categoryGroup, string categoryName, string categoryId, string familyName)
	{
		if (string.IsNullOrWhiteSpace(familyName))
		{
			return false;
		}
		return !FamilyBrowserFamilyClassificationService.IsTypeManagedFamilyLike(categoryName, categoryId, familyName);
	}

	private static bool IsActualNestedLoadableChild(StandardLoadableFamilySnapshotItem item)
	{
		if (item != null && item.IsNestedLoadableChild)
		{
			return true;
		}
		return false;
	}

	private static List<StandardNestedLoadableFamilySnapshotItem> CloneNestedLoadableFamilies(IEnumerable<StandardNestedLoadableFamilySnapshotItem> items)
	{
		List<StandardNestedLoadableFamilySnapshotItem> result = new List<StandardNestedLoadableFamilySnapshotItem>();
		if (items == null)
		{
			return result;
		}
		foreach (StandardNestedLoadableFamilySnapshotItem item in items)
		{
			if (item != null && IsModelNestedLoadableChild(item))
			{
				result.Add(new StandardNestedLoadableFamilySnapshotItem
				{
					FamilyName = (item.FamilyName ?? string.Empty),
					CategoryName = (item.CategoryName ?? string.Empty),
					CategoryId = (item.CategoryId ?? string.Empty),
					CategoryGroup = (item.CategoryGroup ?? string.Empty),
					TypeCount = item.TypeCount,
					TypeNames = (item.TypeNames ?? new List<string>()).ToList(),
					IsShared = item.IsShared
				});
			}
		}
		return result;
	}

	private static List<ProjectSystemTypeSnapshotItem> FindProjectSystemTypeCategoryMismatches(StandardSystemTypeSnapshotItem standardType, Dictionary<string, List<ProjectSystemTypeSnapshotItem>> projectNameMap)
	{
		if (standardType == null || projectNameMap == null)
		{
			return new List<ProjectSystemTypeSnapshotItem>();
		}
		List<ProjectSystemTypeSnapshotItem> candidates = null;
		string nameKey = Normalize(standardType.TypeClassName) + "|" + Normalize(standardType.TypeName);
		if (!projectNameMap.TryGetValue(nameKey, out candidates) || candidates == null || candidates.Count == 0)
		{
			return new List<ProjectSystemTypeSnapshotItem>();
		}
		return candidates.Where([SpecialName] (ProjectSystemTypeSnapshotItem x) => !SystemTypeCategoryMatches(standardType, x)).OrderBy<ProjectSystemTypeSnapshotItem, string>([SpecialName] (ProjectSystemTypeSnapshotItem x) => BuildSystemMatchKey(x), StringComparer.Ordinal).ToList();
	}

	private static ProjectSystemTypeSnapshotItem FindProjectSystemTypeCategoryMatch(StandardSystemTypeSnapshotItem standardType, Dictionary<string, List<ProjectSystemTypeSnapshotItem>> projectNameMap)
	{
		if (standardType == null || projectNameMap == null)
		{
			return null;
		}
		List<ProjectSystemTypeSnapshotItem> candidates = null;
		string nameKey = Normalize(standardType.TypeClassName) + "|" + Normalize(standardType.TypeName);
		if (!projectNameMap.TryGetValue(nameKey, out candidates) || candidates == null || candidates.Count == 0)
		{
			return null;
		}
		return candidates.Where([SpecialName] (ProjectSystemTypeSnapshotItem x) => SystemTypeCategoryMatches(standardType, x)).OrderBy<ProjectSystemTypeSnapshotItem, string>([SpecialName] (ProjectSystemTypeSnapshotItem x) => BuildSystemMatchKey(x), StringComparer.Ordinal).FirstOrDefault();
	}

	private static List<StandardSystemTypeSnapshotItem> FindStandardSystemTypeCategoryMismatches(ProjectSystemTypeSnapshotItem projectType, Dictionary<string, List<StandardSystemTypeSnapshotItem>> standardNameMap)
	{
		if (projectType == null || standardNameMap == null)
		{
			return new List<StandardSystemTypeSnapshotItem>();
		}
		List<StandardSystemTypeSnapshotItem> candidates = null;
		string nameKey = Normalize(projectType.TypeClassName) + "|" + Normalize(projectType.TypeName);
		if (!standardNameMap.TryGetValue(nameKey, out candidates) || candidates == null || candidates.Count == 0)
		{
			return new List<StandardSystemTypeSnapshotItem>();
		}
		return candidates.Where([SpecialName] (StandardSystemTypeSnapshotItem x) => !SystemTypeCategoryMatches(x, projectType)).OrderBy<StandardSystemTypeSnapshotItem, string>([SpecialName] (StandardSystemTypeSnapshotItem x) => BuildSystemMatchKey(x), StringComparer.Ordinal).ToList();
	}

	private static string BuildSystemTypeCategoryMismatchNote(StandardSystemTypeSnapshotItem standardType, List<ProjectSystemTypeSnapshotItem> projectTypes)
	{
		List<string> projectCategories = (from x in projectTypes ?? new List<ProjectSystemTypeSnapshotItem>()
			select x.CategoryName ?? string.Empty into x
			where !string.IsNullOrWhiteSpace(x)
			select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
		return "System type name matches a project type, but the category differs. Standard category: " + (standardType?.CategoryName ?? string.Empty) + " / Project category: " + ((projectCategories.Count == 0) ? string.Empty : string.Join(", ", projectCategories)) + ". Treat this as a system type category mismatch and review it before applying.";
	}

	private static string BuildProjectSystemTypeCategoryMismatchNote(ProjectSystemTypeSnapshotItem projectType, List<StandardSystemTypeSnapshotItem> standardTypes)
	{
		List<string> standardCategories = (from x in standardTypes ?? new List<StandardSystemTypeSnapshotItem>()
			select x.CategoryName ?? string.Empty into x
			where !string.IsNullOrWhiteSpace(x)
			select x).Distinct<string>(StringComparer.OrdinalIgnoreCase).OrderBy<string, string>([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
		return "Project system type name exists in the standard, but the category differs. Standard category: " + ((standardCategories.Count == 0) ? string.Empty : string.Join(", ", standardCategories)) + " / Project category: " + (projectType?.CategoryName ?? string.Empty) + ". Treat this as a system type category mismatch and review it before applying.";
	}

	private static Dictionary<string, TrackedLoadableFamilyState> BuildTrackedLoadableMap(ProjectTrackingCatalog trackingCatalog)
	{
		if (trackingCatalog == null || trackingCatalog.LoadableFamilies == null)
		{
			return new Dictionary<string, TrackedLoadableFamilyState>(StringComparer.Ordinal);
		}
		Dictionary<string, TrackedLoadableFamilyState> result = new Dictionary<string, TrackedLoadableFamilyState>(StringComparer.Ordinal);
		foreach (TrackedLoadableFamilyState item in trackingCatalog.LoadableFamilies)
		{
			if (item != null && !string.IsNullOrWhiteSpace(item.IdentityKey))
			{
				result[item.IdentityKey] = item;
			}
		}
		return result;
	}

	private static Dictionary<string, TrackedSystemTypeState> BuildTrackedSystemMap(ProjectTrackingCatalog trackingCatalog)
	{
		if (trackingCatalog == null || trackingCatalog.SystemTypes == null)
		{
			return new Dictionary<string, TrackedSystemTypeState>(StringComparer.Ordinal);
		}
		Dictionary<string, TrackedSystemTypeState> result = new Dictionary<string, TrackedSystemTypeState>(StringComparer.Ordinal);
		foreach (TrackedSystemTypeState item in trackingCatalog.SystemTypes)
		{
			if (item != null && !string.IsNullOrWhiteSpace(item.IdentityKey))
			{
				result[item.IdentityKey] = item;
			}
		}
		return result;
	}

	private static Dictionary<string, List<T>> BuildGroupedMap<T>(IEnumerable<T> items, Func<T, string> keySelector)
	{
		Dictionary<string, List<T>> result = new Dictionary<string, List<T>>(StringComparer.Ordinal);
		if (items == null)
		{
			return result;
		}
		foreach (T item in items)
		{
			string key = keySelector(item);
			List<T> bucket = null;
			if (!result.TryGetValue(key, out bucket))
			{
				bucket = new List<T>();
				result.Add(key, bucket);
			}
			bucket.Add(item);
		}
		return result;
	}

	private static Dictionary<string, T> BuildFirstMap<T>(IEnumerable<T> items, Func<T, string> keySelector)
	{
		Dictionary<string, T> result = new Dictionary<string, T>(StringComparer.Ordinal);
		if (items == null)
		{
			return result;
		}
		foreach (T item in items)
		{
			string key = keySelector(item);
			if (!result.ContainsKey(key))
			{
				result.Add(key, item);
			}
		}
		return result;
	}

	private static ComparisonClassificationResult ClassifyTrackedStatus(string approvedFingerprint, string approvedStandardStamp, string currentStandardFingerprint, string currentStandardStamp, string currentProjectFingerprint)
	{
		string normalizedApprovedFingerprint = Normalize(approvedFingerprint);
		string normalizedApprovedStamp = Normalize(approvedStandardStamp);
		string normalizedCurrentStandardFingerprint = Normalize(currentStandardFingerprint);
		string normalizedCurrentStandardStamp = Normalize(currentStandardStamp);
		string normalizedCurrentProjectFingerprint = Normalize(currentProjectFingerprint);
		if (normalizedApprovedFingerprint.Length == 0)
		{
			if (string.Equals(normalizedCurrentProjectFingerprint, normalizedCurrentStandardFingerprint, StringComparison.Ordinal))
			{
				return new ComparisonClassificationResult
				{
					Status = "LoadedLatest",
					Notes = "Project content matches the current standard snapshot. Tracking stamp is missing because this item may have been loaded outside Family Browser; no reload is required."
				};
			}
			return new ComparisonClassificationResult
			{
				Status = "DifferentFromStandard",
				Notes = "Project content differs from the standard snapshot and no tracked approval stamp exists yet."
			};
		}
		if (string.Equals(normalizedCurrentProjectFingerprint, normalizedCurrentStandardFingerprint, StringComparison.Ordinal))
		{
			if (string.Equals(normalizedApprovedFingerprint, normalizedCurrentStandardFingerprint, StringComparison.Ordinal) && string.Equals(normalizedApprovedStamp, normalizedCurrentStandardStamp, StringComparison.Ordinal))
			{
				return new ComparisonClassificationResult
				{
					Status = "LoadedLatest",
					Notes = "Project content matches the current standard snapshot and tracked approval stamp."
				};
			}
			return new ComparisonClassificationResult
			{
				Status = "StampNormalizationNeeded",
				Notes = BuildStampNormalizationNote(approvedFingerprint, approvedStandardStamp, currentStandardFingerprint, currentStandardStamp)
			};
		}
		if (string.Equals(normalizedCurrentProjectFingerprint, normalizedApprovedFingerprint, StringComparison.Ordinal) && !string.Equals(normalizedCurrentStandardFingerprint, normalizedApprovedFingerprint, StringComparison.Ordinal))
		{
			return new ComparisonClassificationResult
			{
				Status = "UpdateAvailable",
				Notes = "Project content still matches the previously approved stamp, but the standard snapshot has changed."
			};
		}
		if (string.Equals(normalizedCurrentStandardFingerprint, normalizedApprovedFingerprint, StringComparison.Ordinal) && !string.Equals(normalizedCurrentProjectFingerprint, normalizedApprovedFingerprint, StringComparison.Ordinal))
		{
			return new ComparisonClassificationResult
			{
				Status = "LocallyModified",
				Notes = "Tracked approval matches the standard snapshot, but the current project content has drifted."
			};
		}
		return new ComparisonClassificationResult
		{
			Status = "VersionConflict",
			Notes = "Project content, tracked approval stamp, and current standard snapshot are inconsistent."
		};
	}

	private static string BuildMissingTrackingStampNote(string approvedStandardStamp, string currentStandardFingerprint, string currentStandardStamp)
	{
		List<string> details = new List<string>();
		details.Add("Tracking refresh reason: tracked fingerprint is missing");
		details.Add("Project content: matches current standard");
		details.Add("Tracked fingerprint: " + ShortDiagnosticValue(string.Empty, 12));
		details.Add("Current standard fingerprint: " + ShortDiagnosticValue(currentStandardFingerprint, 12));
		if (Normalize(approvedStandardStamp).Length == 0)
		{
			details.Add("Tracked standard stamp: " + ShortDiagnosticValue(string.Empty, 24));
		}
		else
		{
			details.Add("Tracked standard stamp: " + ShortDiagnosticValue(approvedStandardStamp, 24));
		}
		details.Add("Current standard stamp: " + ShortDiagnosticValue(currentStandardStamp, 24));
		return string.Join(" | ", details);
	}

	private static string BuildStampNormalizationNote(string approvedFingerprint, string approvedStandardStamp, string currentStandardFingerprint, string currentStandardStamp)
	{
		string normalizedApprovedFingerprint = Normalize(approvedFingerprint);
		string normalizedApprovedStamp = Normalize(approvedStandardStamp);
		string normalizedCurrentStandardFingerprint = Normalize(currentStandardFingerprint);
		string normalizedCurrentStandardStamp = Normalize(currentStandardStamp);
		List<string> reasons = new List<string>();
		bool fingerprintMismatch = !string.Equals(normalizedApprovedFingerprint, normalizedCurrentStandardFingerprint, StringComparison.Ordinal);
		bool stampMismatch = !string.Equals(normalizedApprovedStamp, normalizedCurrentStandardStamp, StringComparison.Ordinal);
		if (normalizedApprovedFingerprint.Length == 0)
		{
			reasons.Add("tracked fingerprint is missing");
		}
		else if (fingerprintMismatch)
		{
			reasons.Add("tracked fingerprint differs from current standard fingerprint");
		}
		if (normalizedApprovedStamp.Length == 0)
		{
			reasons.Add("tracked standard stamp is missing");
		}
		else if (stampMismatch)
		{
			reasons.Add("tracked standard stamp differs from current standard snapshot");
		}
		if (reasons.Count == 0)
		{
			reasons.Add("tracked approval stamp is incomplete");
		}
		List<string> details = new List<string>();
		details.Add("Tracking refresh reason: " + string.Join(", ", reasons));
		details.Add("Project content: matches current standard");
		if (fingerprintMismatch || normalizedApprovedFingerprint.Length == 0)
		{
			details.Add("Tracked fingerprint: " + ShortDiagnosticValue(approvedFingerprint, 12));
			details.Add("Current standard fingerprint: " + ShortDiagnosticValue(currentStandardFingerprint, 12));
		}
		if (stampMismatch || normalizedApprovedStamp.Length == 0)
		{
			details.Add("Tracked standard stamp: " + ShortDiagnosticValue(approvedStandardStamp, 24));
			details.Add("Current standard stamp: " + ShortDiagnosticValue(currentStandardStamp, 24));
		}
		return string.Join(" | ", details);
	}

	private static string ShortDiagnosticValue(string value, int maxLength = 18)
	{
		if (string.IsNullOrWhiteSpace(value))
		{
			return "(missing)";
		}
		string trimmed = value.Trim();
		int safeLength = Math.Max(4, maxLength);
		if (trimmed.Length <= safeLength)
		{
			return trimmed;
		}
		return trimmed.Substring(0, safeLength) + "...";
	}

	private static string BuildLoadableKey(string categoryName, string familyName)
	{
		return BuildLoadableIdentityKey(categoryName, familyName);
	}

	private static string BuildLoadableIdentityKey(string categoryName, string familyName)
	{
		return Normalize(categoryName) + "|" + Normalize(familyName);
	}

	private static string BuildLoadableIdentityKey(StandardLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return ProjectStandardComparisonService.BuildLoadableIdentityKey(item.CategoryName, item.FamilyName);
	}

	private static string BuildLoadableIdentityKey(ProjectLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return ProjectStandardComparisonService.BuildLoadableIdentityKey(item.CategoryName, item.FamilyName);
	}

	private static string BuildLoadableMatchKey(StandardLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return BuildCategoryIdentity(item.CategoryId, item.CategoryName) + "|" + Normalize(item.FamilyName);
	}

	private static string BuildLoadableMatchKey(ProjectLoadableFamilySnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return BuildCategoryIdentity(item.CategoryId, item.CategoryName) + "|" + Normalize(item.FamilyName);
	}

	private static string BuildCategoryIdentity(string categoryId, string categoryName)
	{
		string normalizedId = Normalize(categoryId);
		if (normalizedId.Length > 0)
		{
			return "id:" + normalizedId;
		}
		return "name:" + Normalize(categoryName);
	}

	private static bool LoadableCategoryMatches(StandardLoadableFamilySnapshotItem standardFamily, ProjectLoadableFamilySnapshotItem projectFamily)
	{
		if (standardFamily == null || projectFamily == null)
		{
			return false;
		}
		string standardCategoryId = Normalize(standardFamily.CategoryId);
		string projectCategoryId = Normalize(projectFamily.CategoryId);
		if (standardCategoryId.Length > 0 && projectCategoryId.Length > 0)
		{
			return string.Equals(standardCategoryId, projectCategoryId, StringComparison.Ordinal);
		}
		return string.Equals(Normalize(standardFamily.CategoryName), Normalize(projectFamily.CategoryName), StringComparison.Ordinal);
	}

	private static string BuildSystemKey(string typeClassName, string categoryName, string typeName)
	{
		return Normalize(typeClassName) + "|" + Normalize(categoryName) + "|" + Normalize(typeName);
	}

	private static string BuildSystemMatchKey(StandardSystemTypeSnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return Normalize(item.TypeClassName) + "|" + BuildCategoryIdentity(item.CategoryId, item.CategoryName) + "|" + Normalize(item.TypeName);
	}

	private static string BuildSystemMatchKey(ProjectSystemTypeSnapshotItem item)
	{
		if (item == null)
		{
			return string.Empty;
		}
		return Normalize(item.TypeClassName) + "|" + BuildCategoryIdentity(item.CategoryId, item.CategoryName) + "|" + Normalize(item.TypeName);
	}

	private static bool SystemTypeCategoryMatches(StandardSystemTypeSnapshotItem standardType, ProjectSystemTypeSnapshotItem projectType)
	{
		if (standardType == null || projectType == null)
		{
			return false;
		}
		string standardCategoryId = Normalize(standardType.CategoryId);
		string projectCategoryId = Normalize(projectType.CategoryId);
		if (standardCategoryId.Length > 0 && projectCategoryId.Length > 0)
		{
			return string.Equals(standardCategoryId, projectCategoryId, StringComparison.Ordinal);
		}
		return string.Equals(Normalize(standardType.CategoryName), Normalize(projectType.CategoryName), StringComparison.Ordinal);
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
