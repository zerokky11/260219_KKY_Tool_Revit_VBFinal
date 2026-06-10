using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using Microsoft.VisualBasic.CompilerServices;

public sealed class LoadableFamilySyncExecutionService
{
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

	[CompilerGenerated]
	internal sealed class _Closure_0024__6_002D0
	{
		public ISet<int> _0024VB_0024Local_familyStateBefore;

		public string _0024VB_0024Local_expectedName;

		public _Closure_0024__6_002D0(_Closure_0024__6_002D0 arg0)
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
	internal sealed class _Closure_0024__8_002D0
	{
		public string _0024VB_0024Local_normalizedName;

		public Func<Family, bool> _0024I0;

		public _Closure_0024__8_002D0(_Closure_0024__8_002D0 arg0)
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

	private LoadableFamilySyncExecutionService()
	{
	}

	public static LoadableFamilySyncExecutionReport Execute(Document targetDocument, Document standardDocument, LoadableFamilySyncPlan plan, Action<int, int, string> progress = null)
	{
		if (targetDocument == null)
		{
			throw new ArgumentNullException("targetDocument");
		}
		if (plan == null)
		{
			throw new ArgumentNullException("plan");
		}
		LoadableFamilySyncExecutionReport report = new LoadableFamilySyncExecutionReport
		{
			GeneratedAtUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
			ProjectDocumentTitle = plan.ProjectDocumentTitle,
			ProjectDocumentPath = (string.IsNullOrWhiteSpace(plan.ProjectDocumentPath) ? ProjectSnapshotStore.ResolveProjectIdentityPath(targetDocument) : plan.ProjectDocumentPath),
			StandardDisplayName = plan.StandardDisplayName,
			ComparisonPath = plan.ComparisonPath
		};
		List<LoadableFamilySyncPlanItem> planItems = plan.Items ?? new List<LoadableFamilySyncPlanItem>();
		bool num = planItems.Any([SpecialName] (LoadableFamilySyncPlanItem x) =>
		{
			string a = Normalize(x.ExecutionMode);
			return string.Equals(a, "load", StringComparison.Ordinal) || string.Equals(a, "reload", StringComparison.Ordinal);
		});
		Dictionary<string, Family> standardFamilyMap = null;
		if (num)
		{
			if (standardDocument == null)
			{
				throw new InvalidOperationException(T("A standard RVT document is required for loadable family execution.", "로더블 패밀리 실행에는 표준 RVT 문서가 필요합니다."));
			}
			standardFamilyMap = BuildStandardFamilyMap(standardDocument);
		}
		LoadableFamilyLoadOptions loadOptions = new LoadableFamilyLoadOptions(overwriteParameterValues: true);
		int total = Math.Max(1, planItems.Count);
		ReportProgress(progress, 0, total, T("Preparing family load execution...", "패밀리 로드 실행 준비 중..."));
		checked
		{
			using (FamilyBrowserNativeCommandGuardService.BeginTrustedOperation("Family Browser loadable family sync"))
			{
				int num2 = planItems.Count - 1;
				for (int itemIndex = 0; itemIndex <= num2; itemIndex++)
				{
					LoadableFamilySyncPlanItem planItem = planItems[itemIndex];
					ReportProgress(progress, itemIndex + 1, total, T("Processing family ", "패밀리 처리 중 ") + (itemIndex + 1).ToString(CultureInfo.InvariantCulture) + "/" + total.ToString(CultureInfo.InvariantCulture) + ": " + (planItem.FamilyName ?? string.Empty));
					LoadableFamilySyncExecutionItem executionItem = new LoadableFamilySyncExecutionItem
					{
						IdentityKey = planItem.IdentityKey,
						FamilyName = planItem.FamilyName,
						CategoryName = planItem.CategoryName,
						ComparisonStatus = planItem.ComparisonStatus,
						PlannedAction = planItem.PlannedAction,
						ExecutionMode = planItem.ExecutionMode,
						Outcome = "Skipped",
						Details = planItem.Notes
					};
					switch (Normalize(planItem.ExecutionMode))
					{
					case "load":
					case "reload":
					{
						string key = BuildKey(planItem.CategoryName, planItem.FamilyName);
						Family standardFamily = null;
						if (standardFamilyMap == null || !standardFamilyMap.TryGetValue(key, out standardFamily) || standardFamily == null)
						{
							executionItem.Outcome = "Failed";
							executionItem.Details = T("The family was not found in the registered standard RVT.", "등록된 표준 RVT에서 해당 패밀리를 찾지 못했습니다.");
							report.Summary.FailedCount++;
							break;
						}
						if (standardFamily.IsInPlace)
						{
							executionItem.Outcome = "Failed";
							executionItem.Details = T("In-place families cannot be synchronized through this command.", "내부 패밀리는 이 명령으로 동기화할 수 없습니다.");
							report.Summary.FailedCount++;
							break;
						}
						if (!standardFamily.IsEditable)
						{
							executionItem.Outcome = "Failed";
							executionItem.Details = T("The family exists in the standard RVT but is not editable through the Revit API.", "패밀리가 표준 RVT에 있지만 Revit API로 편집할 수 없습니다.");
							report.Summary.FailedCount++;
							break;
						}
						Document familyDoc = null;
						try
						{
							EnsureNoCategoryConflictBeforeLoad(targetDocument, standardFamily, planItem);
							ISet<int> familyStateBefore = CaptureFamilyNameState(targetDocument);
							familyDoc = standardDocument.EditFamily(standardFamily);
							List<AllowedLoadedFamilyIdentity> allowedLoadedFamilies = BuildAllowedLoadedFamilyIdentities(standardFamily, familyDoc, standardDocument, new List<string>
							{
								planItem.FamilyName,
								((Element)standardFamily).Name
							});
							Family loadedFamily = familyDoc.LoadFamily(targetDocument, (IFamilyLoadOptions)(object)loadOptions);
							if (loadedFamily == null)
							{
								executionItem.Outcome = "Failed";
								executionItem.Details = T("Revit returned no family reference after load.", "로드 후 Revit이 패밀리 참조를 반환하지 않았습니다.");
								report.Summary.FailedCount++;
								break;
							}
							GuardFamilyLoadDidNotCreateDuplicateFamilies(targetDocument, familyStateBefore, standardFamily, planItem, loadedFamily, allowedLoadedFamilies, executionItem);
							if (string.Equals(Normalize(planItem.ExecutionMode), "load", StringComparison.Ordinal))
							{
								executionItem.Outcome = "Loaded";
								AddExecutionDetail(executionItem, T("Family loaded from the registered standard RVT.", "등록된 표준 RVT에서 패밀리를 로드했습니다."));
								report.Summary.LoadedCount++;
							}
							else
							{
								executionItem.Outcome = "Reloaded";
								AddExecutionDetail(executionItem, T("Family reloaded from the registered standard RVT with overwrite enabled.", "등록된 표준 RVT에서 패밀리를 덮어쓰기 옵션으로 다시 로드했습니다."));
								report.Summary.ReloadedCount++;
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							executionItem.Outcome = "Failed";
							executionItem.Details = (string.IsNullOrWhiteSpace(executionItem.Details) ? ex2.Message : (ex2.Message + " | " + executionItem.Details));
							report.Summary.FailedCount++;
							ProjectData.ClearProjectError();
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
						break;
					}
					case "stamponly":
						executionItem.Outcome = "StampRefreshPending";
						executionItem.Details = T("No family reload is required. Tracking metadata should be refreshed.", "패밀리 재로드는 필요 없으며 추적 메타데이터 갱신만 필요합니다.");
						break;
					case "blocked":
						executionItem.Outcome = "Blocked";
						executionItem.Details = (string.IsNullOrWhiteSpace(planItem.Notes) ? T("This family requires manual review before overwrite.", "이 패밀리는 덮어쓰기 전에 수동 검토가 필요합니다.") : planItem.Notes);
						report.Summary.BlockedCount++;
						break;
					default:
						executionItem.Outcome = "Skipped";
						executionItem.Details = (string.IsNullOrWhiteSpace(planItem.Notes) ? T("No loadable family action is required.", "실행할 로더블 패밀리 작업이 없습니다.") : planItem.Notes);
						report.Summary.SkippedCount++;
						break;
					}
					report.Items.Add(executionItem);
				}
			}
			ReportProgress(progress, total, total, T("Family load execution finished.", "패밀리 로드 실행이 완료되었습니다."));
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

	private static void EnsureNoCategoryConflictBeforeLoad(Document targetDocument, Family standardFamily, LoadableFamilySyncPlanItem planItem)
	{
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		if (targetDocument == null || standardFamily == null || planItem == null)
		{
			return;
		}
		string text = (((Element)standardFamily).Name ?? string.Empty).Trim();
		string text2 = ResolveCategoryName(standardFamily);
		if (!string.IsNullOrWhiteSpace(text) && !string.IsNullOrWhiteSpace(text2))
		{
			List<string> conflicts = (from Family x in (IEnumerable)new FilteredElementCollector(targetDocument).OfClass(typeof(Family))
				where x != null
				where string.Equals(Normalize(((Element)x).Name), Normalize(text), StringComparison.Ordinal)
				where !CategoryNamesMatch(ResolveCategoryName(x), text2)
				select ResolveCategoryName(x)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase).ToList();
			if (conflicts.Count != 0)
			{
				throw new InvalidOperationException(T("A project family with the same name exists in a different category. Review it before loading the standard family. Family: ", "현재 프로젝트에 같은 이름이지만 카테고리가 다른 패밀리가 있습니다. 표준 패밀리를 로드하기 전에 검토하세요. 패밀리: ") + text + T(" / Standard category: ", " / 표준 카테고리: ") + text2 + T(" / Project category: ", " / 프로젝트 카테고리: ") + string.Join(", ", conflicts));
			}
		}
	}

	private static ISet<int> CaptureFamilyNameState(Document targetDocument)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		return new HashSet<int>(from Family x in (IEnumerable)new FilteredElementCollector(targetDocument).OfClass(typeof(Family))
			where x != null && ((Element)x).Id != null
			select RevitElementIdCompat.CompatIntegerValue(((Element)x).Id));
	}

	private static void GuardFamilyLoadDidNotCreateDuplicateFamilies(Document targetDocument, ISet<int> familyStateBefore, Family standardFamily, LoadableFamilySyncPlanItem planItem, Family loadedFamily, IEnumerable<AllowedLoadedFamilyIdentity> allowedFamilies, LoadableFamilySyncExecutionItem executionItem)
	{
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		_Closure_0024__6_002D0 arg = default(_Closure_0024__6_002D0);
		_Closure_0024__6_002D0 CS_0024_003C_003E8__locals8 = new _Closure_0024__6_002D0(arg);
		CS_0024_003C_003E8__locals8._0024VB_0024Local_familyStateBefore = familyStateBefore;
		CS_0024_003C_003E8__locals8._0024VB_0024Local_expectedName = (((standardFamily != null) ? ((Element)standardFamily).Name : null) ?? planItem.FamilyName).Trim();
		List<string> allowedFamilyNames = UniqueSortedNames((allowedFamilies ?? new List<AllowedLoadedFamilyIdentity>()).Select([SpecialName] (AllowedLoadedFamilyIdentity x) => x.Name ?? string.Empty));
		List<Family> newFamilies = (from Family x in (IEnumerable)new FilteredElementCollector(targetDocument).OfClass(typeof(Family))
			where x != null && ((Element)x).Id != null && !CS_0024_003C_003E8__locals8._0024VB_0024Local_familyStateBefore.Contains(RevitElementIdCompat.CompatIntegerValue(((Element)x).Id))
			select x).ToList();
		if (newFamilies.Count == 0)
		{
			AddFamilyLoadDiagnostic(executionItem, CS_0024_003C_003E8__locals8._0024VB_0024Local_expectedName, ((loadedFamily != null) ? ((Element)loadedFamily).Name : null) ?? string.Empty, allowedFamilyNames, new List<string>(), new List<string>(), new List<string>(), "None");
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
			AddFamilyLoadDiagnostic(executionItem, CS_0024_003C_003E8__locals8._0024VB_0024Local_expectedName, ((loadedFamily != null) ? ((Element)loadedFamily).Name : null) ?? string.Empty, allowedFamilyNames, newFamilyNames, allowedNewNames, new List<string>(), "None");
			List<string> allowedNestedNames = allowedNewNames.Where([SpecialName] (string x) => !string.Equals(Normalize(x), Normalize(CS_0024_003C_003E8__locals8._0024VB_0024Local_expectedName), StringComparison.Ordinal)).ToList();
			if (allowedNestedNames.Count > 0)
			{
				AddExecutionDetail(executionItem, T("Allowed standard nested/helper families loaded: ", "허용된 표준 하위/보조 패밀리를 함께 로드했습니다: ") + string.Join(", ", UniqueSortedNames(allowedNestedNames)));
			}
			return;
		}
		List<string> unexpectedNames = unexpected.Select([SpecialName] (Family x) => ((Element)x).Name ?? string.Empty).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy([SpecialName] (string x) => x, StringComparer.OrdinalIgnoreCase)
			.ToList();
		DuplicateCleanupResult cleanup = TryDeleteFamilies(targetDocument, unexpected);
		AddFamilyLoadDiagnostic(executionItem, CS_0024_003C_003E8__locals8._0024VB_0024Local_expectedName, ((loadedFamily != null) ? ((Element)loadedFamily).Name : null) ?? string.Empty, allowedFamilyNames, newFamilyNames, allowedNewNames, unexpectedNames, DescribeDuplicateCleanupAction(cleanup));
		throw new InvalidOperationException(BuildDuplicateCleanupMessage(T("Family load created duplicate or unexpected families instead of using the standard canonical family.", "패밀리 로드 중 표준 원본 패밀리 대신 중복 또는 예상하지 못한 패밀리가 생성되었습니다."), CS_0024_003C_003E8__locals8._0024VB_0024Local_expectedName, unexpectedNames, cleanup));
	}

	private static List<AllowedLoadedFamilyIdentity> BuildAllowedLoadedFamilyIdentities(Family sourceFamily, Document familyDocument, Document standardDocument, IEnumerable<string> explicitFamilyNames)
	{
		//IL_0115: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_0085: Unknown result type (might be due to invalid IL or missing references)
		List<AllowedLoadedFamilyIdentity> result = new List<AllowedLoadedFamilyIdentity>();
		AddAllowedLoadedFamilyIdentity(result, sourceFamily);
		List<string> requestedNames = new List<string>();
		if (sourceFamily != null)
		{
			requestedNames.Add(((Element)sourceFamily).Name);
		}
		requestedNames.AddRange(explicitFamilyNames ?? new List<string>());
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
		_Closure_0024__8_002D0 arg = default(_Closure_0024__8_002D0);
		_Closure_0024__8_002D0 CS_0024_003C_003E8__locals2 = new _Closure_0024__8_002D0(arg);
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
			string categoryName = ResolveCategoryName(family);
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
		string left = ResolveCategoryName(family);
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

	private static void AddFamilyLoadDiagnostic(LoadableFamilySyncExecutionItem executionItem, string expectedFamilyName, string loadedFamilyName, IEnumerable<string> allowedFamilyNames, IEnumerable<string> actualNewFamilyNames, IEnumerable<string> allowedNewFamilyNames, IEnumerable<string> unexpectedFamilyNames, string cleanupAction)
	{
		AddExecutionDetail(executionItem, "FamilyLoad.Guard - expected family=" + (expectedFamilyName ?? string.Empty) + " / loaded family=" + (loadedFamilyName ?? string.Empty) + " / allowed created family names=" + JoinNamesOrNone(allowedFamilyNames) + " / actual new family names=" + JoinNamesOrNone(actualNewFamilyNames) + " / allowed new family names=" + JoinNamesOrNone(allowedNewFamilyNames) + " / blocked new family names=" + JoinNamesOrNone(unexpectedFamilyNames) + " / cleanup action=" + (cleanupAction ?? string.Empty));
	}

	private static void AddExecutionDetail(LoadableFamilySyncExecutionItem executionItem, string message)
	{
		if (executionItem != null && !string.IsNullOrWhiteSpace(message))
		{
			if (string.IsNullOrWhiteSpace(executionItem.Details))
			{
				executionItem.Details = message;
			}
			else
			{
				LoadableFamilySyncExecutionItem loadableFamilySyncExecutionItem;
				(loadableFamilySyncExecutionItem = executionItem).Details = loadableFamilySyncExecutionItem.Details + " | " + message;
			}
		}
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
			return T("None", "없음");
		}
		if (cleanup.AllDeleted)
		{
			return T("Removed: ", "삭제됨: ") + string.Join(", ", cleanup.DeletedNames);
		}
		if (cleanup.FailedNames.Count > 0)
		{
			return T("Cleanup failed: ", "정리 실패: ") + string.Join(", ", cleanup.FailedNames);
		}
		return T("Cleanup attempted", "정리 시도됨");
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
			return context + " " + T("The unexpected families were removed. Family: ", "예상하지 못한 패밀리를 삭제했습니다. 패밀리: ") + expectedName + T(" / Removed: ", " / 삭제됨: ") + string.Join(", ", cleanup.DeletedNames);
		}
		List<string> parts = new List<string>
		{
			T("Family: ", "패밀리: ") + expectedName,
			T("Created: ", "생성됨: ") + string.Join(", ", UniqueSortedNames(createdNames))
		};
		if (cleanup != null && cleanup.DeletedNames.Count > 0)
		{
			parts.Add(T("Removed: ", "삭제됨: ") + string.Join(", ", cleanup.DeletedNames));
		}
		if (cleanup != null && cleanup.FailedNames.Count > 0)
		{
			parts.Add(T("Failed cleanup: ", "정리 실패: ") + string.Join(", ", cleanup.FailedNames));
		}
		if (cleanup != null && !string.IsNullOrWhiteSpace(cleanup.ExceptionMessage))
		{
			parts.Add(T("Cleanup error: ", "정리 오류: ") + cleanup.ExceptionMessage);
		}
		return context + " " + T("Automatic cleanup failed for one or more families. Admin cleanup is required. ", "일부 패밀리 자동 정리에 실패했습니다. 관리자 정리가 필요합니다. ") + string.Join(" / ", parts);
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

	private static Dictionary<string, Family> BuildStandardFamilyMap(Document doc)
	{
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		Dictionary<string, Family> result = new Dictionary<string, Family>(StringComparer.Ordinal);
		foreach (Family family in (from Family x in (IEnumerable)new FilteredElementCollector(doc).OfClass(typeof(Family))
			where x != null
			select x).OrderBy([SpecialName] (Family x) => BuildKey(ResolveCategoryName(x), ((Element)x).Name), StringComparer.Ordinal))
		{
			result[BuildKey(ResolveCategoryName(family), ((Element)family).Name)] = family;
		}
		return result;
	}

	private static string ResolveCategoryName(Family family)
	{
		string ResolveCategoryName;
		try
		{
			Category familyCategory = family.FamilyCategory;
			ResolveCategoryName = ((familyCategory != null) ? familyCategory.Name : null) ?? string.Empty;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ResolveCategoryName = string.Empty;
			ProjectData.ClearProjectError();
		}
		return ResolveCategoryName;
	}

	private static string BuildKey(string categoryName, string familyName)
	{
		return Normalize(categoryName) + "|" + Normalize(familyName);
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
}
