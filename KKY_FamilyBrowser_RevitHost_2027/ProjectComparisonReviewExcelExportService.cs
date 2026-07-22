using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

public sealed class ProjectComparisonReviewExcelExportService
{
	private sealed class DifferenceExportFields
	{
		public string Area { get; set; }

		public string StandardValue { get; set; }

		public string ProjectValue { get; set; }

		public string Summary { get; set; }

		public DifferenceExportFields()
		{
			Area = string.Empty;
			StandardValue = string.Empty;
			ProjectValue = string.Empty;
			Summary = string.Empty;
		}
	}

	private const string DefaultSheetName = "Review";

	private static readonly HashSet<string> ReviewStatuses = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
	{
		"DifferentFromStandard", "LoadedWithoutVersionStamp", "StampNormalizationNeeded", "UpdateAvailable", "LocallyModified", "VersionConflict", "CategoryMismatch", "ProjectOnly", "ManualReview", "ReviewNeeded",
		"ApprovalRequired"
	};

	private ProjectComparisonReviewExcelExportService()
	{
	}

	public static int CountReviewRows(ProjectStandardComparisonReport report)
	{
		if (report == null)
		{
			return 0;
		}
		int count = 0;
		checked
		{
			if (report.LoadableFamilies != null)
			{
				count += report.LoadableFamilies.Where([SpecialName] (LoadableFamilyComparisonItem item) => item != null && IsReviewExportStatus(item.Status)).Count();
			}
			if (report.ProjectLoadableSignatureFailures != null)
			{
				count += report.ProjectLoadableSignatureFailures.Where([SpecialName] (ProjectLoadableSignatureFailureItem item) => item != null).Count();
			}
			if (report.SystemTypes != null)
			{
				count += report.SystemTypes.Where([SpecialName] (SystemTypeComparisonItem item) => item != null && IsReviewExportStatus(item.Status)).Count();
			}
			return count;
		}
	}

	public static ProjectComparisonReviewExcelExportResult SaveReviewList(string outputPath, ProjectStandardComparisonReport report, string disciplineLabel, string comparisonReportPath, bool korean = false)
	{
		return SaveReviewList(outputPath, report, disciplineLabel, comparisonReportPath, null, korean);
	}

	public static ProjectComparisonReviewExcelExportResult SaveReviewList(string outputPath, ProjectStandardComparisonReport report, string disciplineLabel, string comparisonReportPath, IEnumerable<FamilyThumbnailAutoConfirmedDialogRecord> autoHandledDialogs, bool korean = false)
	{
		if (report == null)
		{
			throw new ArgumentNullException("report");
		}
		List<List<string>> rows = BuildRows(report, disciplineLabel, comparisonReportPath, korean);
		string sheetName = (korean ? "검토결과" : "Review");
		List<string> headers = BuildHeaders(korean);
		List<List<string>> dialogRows = BuildAutoHandledDialogRows(autoHandledDialogs);
		string dialogSheetName = (korean ? "스캔경고" : "ScanDialogs");
		string workbookPath = EnsureWorkbookOutputPath(outputPath);
		WriteWorkbook(workbookPath, sheetName, headers, rows, dialogSheetName, BuildAutoHandledDialogHeaders(korean), dialogRows);
		return new ProjectComparisonReviewExcelExportResult
		{
			OutputPath = workbookPath,
			RowCount = checked(rows.Count + dialogRows.Count),
			SheetName = sheetName
		};
	}

	private static List<string> BuildAutoHandledDialogHeaders(bool korean)
	{
		return korean
			? new List<string> { "처리 시각(UTC)", "카테고리", "패밀리", "자동 동작", "판정 사유", "처리 결과", "사용 가능 버튼", "오류/경고 내용" }
			: new List<string> { "Handled At (UTC)", "Category", "Family", "Automatic Action", "Reason", "Result", "Available Buttons", "Error / Warning Text" };
	}

	private static List<List<string>> BuildAutoHandledDialogRows(IEnumerable<FamilyThumbnailAutoConfirmedDialogRecord> records)
	{
		List<List<string>> rows = new List<List<string>>();
		foreach (FamilyThumbnailAutoConfirmedDialogRecord record in records ?? Enumerable.Empty<FamilyThumbnailAutoConfirmedDialogRecord>())
		{
			if (record == null)
			{
				continue;
			}
			rows.Add(new List<string>
			{
				record.ConfirmedAtUtc ?? string.Empty,
				record.CategoryName ?? string.Empty,
				record.FamilyName ?? string.Empty,
				record.ActionTaken ?? string.Empty,
				record.Reason ?? string.Empty,
				record.OverrideResult ?? string.Empty,
				record.AvailableButtons ?? string.Empty,
				record.DialogText ?? string.Empty
			});
		}
		return rows;
	}

	private static List<string> BuildHeaders(bool korean)
	{
		if (korean)
		{
			return new List<string>
			{
				"공종", "항목 구분", "상태", "카테고리", "패밀리", "타입", "시스템 패밀리", "표준 타입 수", "프로젝트 타입 수", "프로젝트 인스턴스 수",
				"비고", "차이 항목", "표준", "프로젝트", "차이 요약", "표준 RVT", "프로젝트", "비교 리포트"
			};
		}
		if (korean)
		{
			return new List<string>
			{
				"공종", "항목 구분", "상태", "카테고리", "패밀리", "타입", "시스템 패밀리", "표준 타입 수", "프로젝트 타입 수", "프로젝트 인스턴스 수",
				"비고", "표준 RVT", "프로젝트", "비교 리포트"
			};
		}
		return new List<string>
		{
			"Discipline", "Item Kind", "Status", "Category", "Family", "Type", "System Family", "Standard Type Count", "Project Type Count", "Project Instance Count",
			"Notes", "Difference Item", "Standard", "Project", "Difference Summary", "Standard RVT", "Project", "Comparison Report"
		};
	}

	private static List<List<string>> BuildRows(ProjectStandardComparisonReport report, string disciplineLabel, string comparisonReportPath, bool korean)
	{
		List<List<string>> rows = new List<List<string>>();
		string discipline = FirstNonEmpty(disciplineLabel, report?.Standard?.DisplayName ?? string.Empty, "-");
		string standardName = FirstNonEmpty(report?.Standard?.DisplayName ?? string.Empty, report?.Standard?.ResolvedPath ?? string.Empty, "-");
		string projectName = FirstNonEmpty(report?.Project?.DocumentTitle ?? string.Empty, report?.Project?.DocumentPath ?? string.Empty, "-");
		if (report.LoadableFamilies != null)
		{
			foreach (LoadableFamilyComparisonItem item in report.LoadableFamilies.Where([SpecialName] (LoadableFamilyComparisonItem x) => x != null && IsReviewExportStatus(x.Status)).OrderBy<LoadableFamilyComparisonItem, string>([SpecialName] (LoadableFamilyComparisonItem x) => x.CategoryName ?? string.Empty, StringComparer.OrdinalIgnoreCase).ThenBy<LoadableFamilyComparisonItem, string>([SpecialName] (LoadableFamilyComparisonItem x) => x.FamilyName ?? string.Empty, StringComparer.OrdinalIgnoreCase))
			{
				DifferenceExportFields difference = BuildLoadableDifferenceFields(item, korean);
				rows.Add(new List<string>
				{
					discipline,
					korean ? "로더블 패밀리" : "Loadable Family",
					DisplayStatus(item.Status, korean),
					item.CategoryName ?? string.Empty,
					item.FamilyName ?? string.Empty,
					string.Empty,
					string.Empty,
					item.StandardTypeCount.ToString(CultureInfo.InvariantCulture),
					item.ProjectTypeCount.ToString(CultureInfo.InvariantCulture),
					item.ProjectInstanceCount.ToString(CultureInfo.InvariantCulture),
					BuildLoadableNotes(item, korean),
					difference.Area,
					difference.StandardValue,
					difference.ProjectValue,
					difference.Summary,
					standardName,
					projectName,
					comparisonReportPath ?? string.Empty
				});
			}
		}
		if (report.SystemTypes != null)
		{
			foreach (SystemTypeComparisonItem item2 in report.SystemTypes.Where([SpecialName] (SystemTypeComparisonItem x) => x != null && IsReviewExportStatus(x.Status)).OrderBy<SystemTypeComparisonItem, string>([SpecialName] (SystemTypeComparisonItem x) => x.CategoryName ?? string.Empty, StringComparer.OrdinalIgnoreCase).ThenBy<SystemTypeComparisonItem, string>([SpecialName] (SystemTypeComparisonItem x) => x.TypeName ?? string.Empty, StringComparer.OrdinalIgnoreCase))
			{
				DifferenceExportFields difference2 = BuildSystemDifferenceFields(item2, korean);
				rows.Add(new List<string>
				{
					discipline,
					korean ? "시스템 타입" : "System Type",
					DisplayStatus(item2.Status, korean),
					item2.CategoryName ?? string.Empty,
					string.Empty,
					item2.TypeName ?? string.Empty,
					item2.TypeClassName ?? string.Empty,
					string.Empty,
					string.Empty,
					string.Empty,
					BuildSystemNotes(item2, korean),
					difference2.Area,
					difference2.StandardValue,
					difference2.ProjectValue,
					difference2.Summary,
					standardName,
					projectName,
					comparisonReportPath ?? string.Empty
				});
			}
		}
		if (report.ProjectLoadableSignatureFailures != null)
		{
			foreach (ProjectLoadableSignatureFailureItem item3 in report.ProjectLoadableSignatureFailures.Where([SpecialName] (ProjectLoadableSignatureFailureItem x) => x != null).OrderBy<ProjectLoadableSignatureFailureItem, string>([SpecialName] (ProjectLoadableSignatureFailureItem x) => x.CategoryName ?? string.Empty, StringComparer.OrdinalIgnoreCase).ThenBy<ProjectLoadableSignatureFailureItem, string>([SpecialName] (ProjectLoadableSignatureFailureItem x) => x.FamilyName ?? string.Empty, StringComparer.OrdinalIgnoreCase))
			{
				rows.Add(new List<string>
				{
					discipline,
					korean ? "프로젝트 Fingerprint 실패" : "Project Fingerprint Failure",
					korean ? "Fingerprint 누락" : "Fingerprint Missing",
					item3.CategoryName ?? string.Empty,
					item3.FamilyName ?? string.Empty,
					string.Empty,
					string.Empty,
					string.Empty,
					item3.TypeCount.ToString(CultureInfo.InvariantCulture),
					item3.InstanceCount.ToString(CultureInfo.InvariantCulture),
					BuildSignatureFailureNotes(item3, korean),
					korean ? "Fingerprint 생성" : "Fingerprint Creation",
					string.Empty,
					item3.Reason ?? string.Empty,
					korean ? "프로젝트 Fingerprint 생성 실패" : "Project fingerprint was not created",
					standardName,
					projectName,
					comparisonReportPath ?? string.Empty
				});
			}
		}
		return rows;
	}

	private static DifferenceExportFields BuildLoadableDifferenceFields(LoadableFamilyComparisonItem item, bool korean)
	{
		DifferenceExportFields fields = new DifferenceExportFields();
		if (item == null)
		{
			return fields;
		}
		List<string> summaries = new List<string>();
		if (item.FingerprintDifferenceDetails != null)
		{
			foreach (LoadableFingerprintDifferenceDetailItem detail in item.FingerprintDifferenceDetails)
			{
				if (detail != null)
				{
					if (string.IsNullOrWhiteSpace(fields.Area))
					{
						DifferenceExportFields concise = BuildConciseDifferenceFields(detail, korean);
						fields.Area = concise.Area;
						fields.StandardValue = concise.StandardValue;
						fields.ProjectValue = concise.ProjectValue;
					}
					AddIfNotEmpty(summaries, DifferenceBrief(detail, korean));
					if (summaries.Count >= 3)
					{
						break;
					}
				}
			}
		}
		if (summaries.Count == 0)
		{
			if (HasProjectFingerprintMissing(item))
			{
				fields.Area = (korean ? "Fingerprint 누락" : "Fingerprint Missing");
				fields.ProjectValue = FirstNonEmpty(item.ProjectContentFingerprintFailureReason, "(missing)");
				AddIfNotEmpty(summaries, korean ? "Fingerprint 누락" : "Fingerprint missing");
			}
			else if (item.MissingTypeNames != null && item.MissingTypeNames.Count > 0)
			{
				fields.Area = (korean ? "누락 타입" : "Missing Type");
				fields.StandardValue = string.Join(", ", item.MissingTypeNames.Take(10));
				AddIfNotEmpty(summaries, (korean ? "누락 타입: " : "Missing type: ") + item.MissingTypeNames[0]);
			}
			else if (item.ExtraTypeNames != null && item.ExtraTypeNames.Count > 0)
			{
				fields.Area = (korean ? "추가 타입" : "Extra Type");
				fields.ProjectValue = string.Join(", ", item.ExtraTypeNames.Take(10));
				AddIfNotEmpty(summaries, (korean ? "추가 타입: " : "Extra type: ") + item.ExtraTypeNames[0]);
			}
			else if (item.FingerprintDifferenceSummary != null)
			{
				AddIfNotEmpty(summaries, BuildConciseSummaryDifferenceNote(item.FingerprintDifferenceSummary.FirstOrDefault([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)), korean));
			}
		}
		fields.Summary = JoinDistinct(summaries);
		return fields;
	}

	private static DifferenceExportFields BuildSystemDifferenceFields(SystemTypeComparisonItem item, bool korean)
	{
		DifferenceExportFields fields = new DifferenceExportFields();
		if (item == null)
		{
			return fields;
		}
		List<string> summaries = new List<string>();
		if (item.DifferenceSummary != null)
		{
			foreach (string difference in item.DifferenceSummary.Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)).Take(3))
			{
				if (string.IsNullOrWhiteSpace(fields.Area))
				{
					PopulateSystemDifferenceFields(fields, difference, korean);
				}
				AddIfNotEmpty(summaries, BuildSystemDifferenceSummaryText(difference, korean));
			}
		}
		if (string.IsNullOrWhiteSpace(fields.Area))
		{
			string status = Normalize(item.Status);
			if (Operators.CompareString(status, "categorymismatch", TextCompare: false) == 0)
			{
				fields.Area = (korean ? "카테고리" : "Category");
				fields.StandardValue = item.CategoryName ?? string.Empty;
				fields.ProjectValue = ExtractCategoryMismatchProjectValue(item.Notes);
				AddIfNotEmpty(summaries, korean ? "같은 이름의 시스템 타입이 다른 카테고리에 있습니다." : "Same-name system type exists under a different category.");
			}
			else if (Operators.CompareString(status, "projectonly", TextCompare: false) == 0)
			{
				fields.Area = (korean ? "표준 목록" : "Standard List");
				fields.ProjectValue = item.TypeName ?? string.Empty;
				AddIfNotEmpty(summaries, korean ? "표준 목록에 없는 프로젝트 전용 시스템 타입입니다." : "Project-only system type is not in the selected standard.");
			}
			else if (!string.Equals(Normalize(item.StandardFingerprint), Normalize(item.ProjectFingerprint), StringComparison.Ordinal))
			{
				fields.Area = "Fingerprint";
				fields.StandardValue = ShortDiffValue(item.StandardFingerprint, 32);
				fields.ProjectValue = ShortDiffValue(item.ProjectFingerprint, 32);
				AddIfNotEmpty(summaries, korean ? "시스템 타입 Fingerprint 다름" : "System Type fingerprint differs");
			}
		}
		fields.Summary = JoinDistinct(summaries);
		return fields;
	}

	private static void PopulateSystemDifferenceFields(DifferenceExportFields fields, string difference, bool korean)
	{
		string text = CleanDiffCell(difference);
		if (fields == null || string.IsNullOrWhiteSpace(text))
		{
			return;
		}
		Match match = Regex.Match(text, "^(.*?)\\s+differs:\\s+standard\\s+(.*?)\\s+/\\s+project\\s+(.*)$", RegexOptions.IgnoreCase);
		if (match.Success)
		{
			fields.Area = SystemDifferenceAreaLabel(match.Groups[1].Value, korean);
			fields.StandardValue = ShortDiffValue(match.Groups[2].Value, 120);
			fields.ProjectValue = ShortDiffValue(match.Groups[3].Value, 120);
			return;
		}
		if (text.IndexOf("Layer differs", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			fields.Area = korean ? "레이어" : "Layer";
		}
		else if (text.IndexOf("Routing", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			fields.Area = korean ? "라우팅 환경설정" : "Routing Preference";
		}
		else if (text.IndexOf("fingerprint", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			fields.Area = "Fingerprint";
		}
		else
		{
			fields.Area = korean ? "시스템 타입 차이" : "System Type Difference";
		}
		fields.Summary = BuildSystemDifferenceSummaryText(text, korean);
	}

	private static string BuildSystemDifferenceSummaryText(string difference, bool korean)
	{
		string text = CleanDiffCell(difference);
		if (string.IsNullOrWhiteSpace(text))
		{
			return string.Empty;
		}
		Match match = Regex.Match(text, "^(.*?)\\s+differs:\\s+standard\\s+(.*?)\\s+/\\s+project\\s+(.*)$", RegexOptions.IgnoreCase);
		if (match.Success)
		{
			return SystemDifferenceAreaLabel(match.Groups[1].Value, korean);
		}
		if (korean)
		{
			text = text.Replace("RoutingDependencyFingerprint differs", "라우팅 의존 패밀리 Fingerprint 다름").Replace("Routing Preference differs", "라우팅 환경설정 다름").Replace("Routing Criteria differs", "라우팅 기준 다름")
				.Replace("Routing Part Fingerprint differs", "라우팅 부품 Fingerprint 다름")
				.Replace("Layer differs", "레이어 다름")
				.Replace("System Type fingerprint differs", "시스템 타입 Fingerprint 다름")
				.Replace("standard", "표준")
				.Replace("project", "프로젝트");
		}
		return ShortDiffValue(text, 90);
	}

	private static string SystemDifferenceAreaLabel(string area, bool korean)
	{
		string normalized = Normalize(area);
		if (normalized.IndexOf("routingdependencyfingerprint", StringComparison.Ordinal) >= 0)
		{
			return korean ? "라우팅 의존 패밀리 Fingerprint" : "Routing Dependency Fingerprint";
		}
		if (normalized.IndexOf("routing part fingerprint", StringComparison.Ordinal) >= 0)
		{
			return korean ? "라우팅 부품 Fingerprint" : "Routing Part Fingerprint";
		}
		if (normalized.IndexOf("routing criteria", StringComparison.Ordinal) >= 0)
		{
			return korean ? "라우팅 기준" : "Routing Criteria";
		}
		if (normalized.IndexOf("routing preference", StringComparison.Ordinal) >= 0)
		{
			return korean ? "라우팅 환경설정" : "Routing Preference";
		}
		if (normalized.IndexOf("classification", StringComparison.Ordinal) >= 0)
		{
			return korean ? "분류" : "Classification";
		}
		if (normalized.IndexOf("segment", StringComparison.Ordinal) >= 0)
		{
			return "Segment";
		}
		if (normalized.IndexOf("material", StringComparison.Ordinal) >= 0)
		{
			return korean ? "재질" : "Material";
		}
		if (normalized.IndexOf("shape", StringComparison.Ordinal) >= 0)
		{
			return korean ? "형상" : "Shape";
		}
		if (normalized.IndexOf("layer", StringComparison.Ordinal) >= 0)
		{
			return korean ? "레이어" : "Layer";
		}
		if (normalized.IndexOf("fingerprint", StringComparison.Ordinal) >= 0)
		{
			return "Fingerprint";
		}
		return FirstNonEmpty(area, korean ? "시스템 타입 차이" : "System Type Difference");
	}

	private static string ExtractCategoryMismatchProjectValue(string notes)
	{
		string text = CleanDiffCell(notes);
		if (string.IsNullOrWhiteSpace(text))
		{
			return string.Empty;
		}
		Match match = Regex.Match(text, "Project categor(?:y|ies)\\s*:\\s*(.*)$", RegexOptions.IgnoreCase);
		if (match.Success)
		{
			return ShortDiffValue(match.Groups[1].Value, 80);
		}
		return ShortDiffValue(text, 80);
	}

	private static string BuildConciseSummaryDifferenceNote(string reason, bool korean)
	{
		string text = CleanDiffCell(reason);
		if (string.IsNullOrWhiteSpace(text))
		{
			return string.Empty;
		}
		string normalized = Normalize(text);
		if (normalized.IndexOf("signature detail paths", StringComparison.OrdinalIgnoreCase) >= 0 || normalized.IndexOf("fingerprint 상세 경로", StringComparison.OrdinalIgnoreCase) >= 0 || normalized.IndexOf("contentsignaturedebugpath", StringComparison.OrdinalIgnoreCase) >= 0 || normalized.IndexOf("standardsignature=", StringComparison.OrdinalIgnoreCase) >= 0 || normalized.IndexOf("projectsignature=", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return string.Empty;
		}
		if (normalized.IndexOf("type count", StringComparison.OrdinalIgnoreCase) >= 0 || normalized.IndexOf("타입 개수", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return korean ? "타입 수" : "Type count";
		}
		if (normalized.IndexOf("missing project types", StringComparison.OrdinalIgnoreCase) >= 0 || normalized.IndexOf("누락 타입", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return korean ? "누락 타입" : "Missing type";
		}
		if (normalized.IndexOf("extra project types", StringComparison.OrdinalIgnoreCase) >= 0 || normalized.IndexOf("추가 타입", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return korean ? "추가 타입" : "Extra type";
		}
		if (normalized.IndexOf("parameter", StringComparison.OrdinalIgnoreCase) >= 0 || normalized.IndexOf("파라미터", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return korean ? "파라미터" : "Parameter";
		}
		if (normalized.IndexOf("connector", StringComparison.OrdinalIgnoreCase) >= 0 || normalized.IndexOf("커넥터", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return korean ? "커넥터" : "Connector";
		}
		if (normalized.IndexOf("geometry", StringComparison.OrdinalIgnoreCase) >= 0 || normalized.IndexOf("형상", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return korean ? "형상" : "Geometry";
		}
		if (normalized.IndexOf("nested", StringComparison.OrdinalIgnoreCase) >= 0 || normalized.IndexOf("하위", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return korean ? "하위 패밀리" : "Nested family";
		}
		if (normalized.IndexOf("fingerprint", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return "Fingerprint";
		}
		return ShortDiffValue(text, 42);
	}

	private static DifferenceExportFields BuildConciseDifferenceFields(LoadableFingerprintDifferenceDetailItem detail, bool korean)
	{
		DifferenceExportFields fields = new DifferenceExportFields();
		if (detail == null)
		{
			return fields;
		}
		string areaKey = NormalizeDiffToken(detail.Area);
		string kindKey = NormalizeDiffToken(detail.DifferenceKind);
		if (areaKey.StartsWith("elements", StringComparison.Ordinal) && !string.Equals(areaKey, "elements", StringComparison.Ordinal))
		{
			fields.Area = DifferenceAreaLabel(detail.Area, detail.DifferenceKind, korean);
			fields.StandardValue = FirstNonEmpty(SignatureValueCell(detail.StandardValue, korean), SignatureCountCell(detail.Details, standardSide: true, korean));
			fields.ProjectValue = FirstNonEmpty(SignatureValueCell(detail.ProjectValue, korean), SignatureCountCell(detail.Details, standardSide: false, korean));
			fields.Summary = SignatureDifferenceNote(detail, korean);
			if (string.IsNullOrWhiteSpace(fields.Summary))
			{
				fields.Summary = fields.Area;
			}
			return fields;
		}
		switch (areaKey)
		{
		case "familytypes":
			if (Operators.CompareString(kindKey, "count", TextCompare: false) == 0)
			{
				fields.Area = (korean ? "타입 수" : "Type Count");
				fields.StandardValue = ShortDiffValue(detail.StandardValue, 18);
				fields.ProjectValue = ShortDiffValue(detail.ProjectValue, 18);
				fields.Summary = fields.Area;
			}
			else if (!IsDiffBlank(detail.StandardValue))
			{
				fields.Area = (korean ? "누락 타입" : "Missing Type");
				fields.StandardValue = ShortDiffValue(detail.StandardValue, 80);
				fields.Summary = fields.Area + ": " + fields.StandardValue;
			}
			else if (!IsDiffBlank(detail.ProjectValue))
			{
				fields.Area = (korean ? "추가 타입" : "Extra Type");
				fields.ProjectValue = ShortDiffValue(detail.ProjectValue, 80);
				fields.Summary = fields.Area + ": " + fields.ProjectValue;
			}
			else
			{
				fields.Area = (korean ? "타입 구성" : "Type List");
				fields.StandardValue = ShortDiffValue(detail.StandardValue, 60);
				fields.ProjectValue = ShortDiffValue(detail.ProjectValue, 60);
				fields.Summary = fields.Area;
			}
			break;
		case "parametersformulas":
		{
			string changedLabels = BuildParameterChangedLabelText(detail.StandardValue, detail.ProjectValue, korean);
			if (Operators.CompareString(kindKey, "standardonly", TextCompare: false) != 0)
			{
				if (Operators.CompareString(kindKey, "projectonly", TextCompare: false) == 0)
				{
					fields.Area = (korean ? "추가 파라미터" : "Extra Parameter");
					fields.ProjectValue = BuildParameterCompactCell(detail.ProjectValue, string.Empty, korean);
				}
				else
				{
					fields.Area = (korean ? "파라미터" : "Parameter");
					fields.StandardValue = BuildParameterCompactCell(detail.StandardValue, changedLabels, korean);
					fields.ProjectValue = BuildParameterCompactCell(detail.ProjectValue, changedLabels, korean);
					fields.Summary = changedLabels;
				}
			}
			else
			{
				fields.Area = (korean ? "누락 파라미터" : "Missing Parameter");
				fields.StandardValue = BuildParameterCompactCell(detail.StandardValue, string.Empty, korean);
			}
			if (string.IsNullOrWhiteSpace(fields.Summary))
			{
				fields.Summary = fields.Area;
			}
			break;
		}
		case "category":
			fields.Area = (korean ? "카테고리" : "Category");
			fields.StandardValue = ShortDiffValue(detail.StandardValue, 60);
			fields.ProjectValue = ShortDiffValue(detail.ProjectValue, 60);
			break;
		case "categorygroup":
			fields.Area = (korean ? "카테고리 그룹" : "Category Group");
			fields.StandardValue = ShortDiffValue(detail.StandardValue, 60);
			fields.ProjectValue = ShortDiffValue(detail.ProjectValue, 60);
			break;
		case "familyname":
			fields.Area = (korean ? "패밀리명" : "Family Name");
			fields.StandardValue = ShortDiffValue(detail.StandardValue, 60);
			fields.ProjectValue = ShortDiffValue(detail.ProjectValue, 60);
			break;
		case "sharedflag":
			fields.Area = (korean ? "공유 여부" : "Shared");
			fields.StandardValue = ShortDiffValue(detail.StandardValue, 24);
			fields.ProjectValue = ShortDiffValue(detail.ProjectValue, 24);
			break;
		case "contentfingerprint":
			fields.Area = (korean ? "Fingerprint 누락" : "Fingerprint Missing");
			fields.ProjectValue = ShortDiffValue(detail.ProjectValue, 70);
			break;
		case "storedfingerprint":
			fields.Area = "Fingerprint";
			fields.StandardValue = (korean ? "원문 동일" : "Source same");
			fields.ProjectValue = (korean ? "해시 다름" : "Hash differs");
			fields.Summary = (korean ? "재스캔" : "Rescan");
			break;
		case "signaturediagnostics":
			fields.Area = (korean ? "진단 누락" : "Diagnostics");
			fields.StandardValue = SignaturePresence(detail.StandardValue, korean);
			fields.ProjectValue = SignaturePresence(detail.ProjectValue, korean);
			break;
		case "nestedlabels":
			fields.Area = (korean ? "하위 Label" : "Nested Label");
			fields.StandardValue = NestedLabelDiffCell(detail.StandardValue);
			fields.ProjectValue = NestedLabelDiffCell(detail.ProjectValue);
			fields.Summary = NestedLabelDiffNote(detail.Details, korean);
			if (IsDiffBlank(fields.StandardValue) && IsDiffBlank(fields.ProjectValue))
			{
				fields.StandardValue = SignatureCountCell(detail.Details, standardSide: true, korean);
				fields.ProjectValue = SignatureCountCell(detail.Details, standardSide: false, korean);
			}
			break;
		case "nestedloadableinstances":
			fields.Area = (korean ? "하위 패밀리" : "Nested Family");
			fields.StandardValue = FirstNonEmpty(SignatureValueCell(detail.StandardValue, korean), SignatureCountCell(detail.Details, standardSide: true, korean));
			fields.ProjectValue = FirstNonEmpty(SignatureValueCell(detail.ProjectValue, korean), SignatureCountCell(detail.Details, standardSide: false, korean));
			fields.Summary = SignatureDifferenceNote(detail, korean);
			break;
		case "connectors":
			fields.Area = (korean ? "커넥터" : "Connector");
			fields.StandardValue = FirstNonEmpty(SignatureValueCell(detail.StandardValue, korean), SignatureCountCell(detail.Details, standardSide: true, korean));
			fields.ProjectValue = FirstNonEmpty(SignatureValueCell(detail.ProjectValue, korean), SignatureCountCell(detail.Details, standardSide: false, korean));
			fields.Summary = SignatureDifferenceNote(detail, korean);
			break;
		case "geometry":
			fields.Area = (korean ? "형상" : "Geometry");
			fields.StandardValue = FirstNonEmpty(SignatureValueCell(detail.StandardValue, korean), SignatureCountCell(detail.Details, standardSide: true, korean));
			fields.ProjectValue = FirstNonEmpty(SignatureValueCell(detail.ProjectValue, korean), SignatureCountCell(detail.Details, standardSide: false, korean));
			fields.Summary = SignatureDifferenceNote(detail, korean);
			break;
		case "elements":
			fields.Area = (korean ? "요소" : "Element");
			fields.StandardValue = FirstNonEmpty(SignatureValueCell(detail.StandardValue, korean), SignatureCountCell(detail.Details, standardSide: true, korean));
			fields.ProjectValue = FirstNonEmpty(SignatureValueCell(detail.ProjectValue, korean), SignatureCountCell(detail.Details, standardSide: false, korean));
			fields.Summary = SignatureDifferenceNote(detail, korean);
			break;
		default:
			fields.Area = DifferenceAreaLabel(detail.Area, detail.DifferenceKind, korean);
			fields.StandardValue = ShortDiffValue(detail.StandardValue, 60);
			fields.ProjectValue = ShortDiffValue(detail.ProjectValue, 60);
			fields.Summary = ShortDiffValue(detail.Details, 42);
			break;
		}
		if (string.IsNullOrWhiteSpace(fields.Summary))
		{
			if (!IsDiffBlank(fields.StandardValue) && IsDiffBlank(fields.ProjectValue))
			{
				fields.Summary = fields.Area + ": " + fields.StandardValue;
			}
			else if (IsDiffBlank(fields.StandardValue) && !IsDiffBlank(fields.ProjectValue))
			{
				fields.Summary = fields.Area + ": " + fields.ProjectValue;
			}
			else if (!IsDiffBlank(fields.StandardValue) && !IsDiffBlank(fields.ProjectValue))
			{
				fields.Summary = fields.Area;
			}
		}
		return fields;
	}

	private static string DifferenceBrief(LoadableFingerprintDifferenceDetailItem detail, bool korean)
	{
		if (detail == null)
		{
			return string.Empty;
		}
		DifferenceExportFields fields = BuildConciseDifferenceFields(detail, korean);
		return FirstNonEmpty(fields.Summary, fields.Area);
	}

	private static bool IsDiffBlank(string value)
	{
		string text = CleanDiffCell(value);
		if (!string.IsNullOrWhiteSpace(text))
		{
			return string.Equals(text, "-", StringComparison.Ordinal);
		}
		return true;
	}

	private static string SignatureValueCell(string value, bool korean)
	{
		string text = CleanDiffCell(value);
		if (string.IsNullOrWhiteSpace(text) || string.Equals(text, "-", StringComparison.Ordinal))
		{
			return string.Empty;
		}
		string readable = ReadableSignatureValueCell(text, korean);
		if (!string.IsNullOrWhiteSpace(readable))
		{
			return ShortDiffValue(readable, 140);
		}
		text = Regex.Replace(text, "\\s*/\\s*UID\\s+[^,]+", string.Empty, RegexOptions.IgnoreCase);
		text = Regex.Replace(text, "\\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\\b", string.Empty, RegexOptions.IgnoreCase);
		text = text.Replace("|", " · ").Replace(":", " ");
		text = Regex.Replace(text, "\\s{2,}", " ");
		text = text.Trim().Trim(new char[3] { '/', ',', '·' });
		if (string.IsNullOrWhiteSpace(text))
		{
			return string.Empty;
		}
		return ShortDiffValue(text, 120);
	}

	private static string ReadableSignatureValueCell(string value, bool korean)
	{
		string text = CleanDiffCell(value);
		if (string.IsNullOrWhiteSpace(text) || string.Equals(text, "-", StringComparison.Ordinal))
		{
			return string.Empty;
		}
		string ids = ExtractSignatureDebugIds(text);
		text = Regex.Replace(text, "\\[[^\\]]*\\]", string.Empty).Trim();
		List<string> tokens = (from x in text.Split('|')
			select (x ?? string.Empty).Trim() into x
			where x.Length > 0
			select x).ToList();
		if (tokens.Count == 0)
		{
			return string.Empty;
		}
		string kind = NormalizeDiffToken(tokens[0]);
		List<string> parts = new List<string>();
		if (!string.IsNullOrWhiteSpace(ids))
		{
			parts.Add("ID: " + ids);
		}
		switch (kind)
		{
		case "familyinstance":
			AddSignatureField(parts, korean ? "카테고리" : "Category", GetSignatureToken(tokens, 1));
			AddSignatureField(parts, korean ? "패밀리" : "Family", GetSignatureToken(tokens, 2));
			AddSignatureField(parts, korean ? "타입" : "Type", FirstNonEmpty(GetSignatureToken(tokens, 5), GetSignatureToken(tokens, 2)));
			break;
		case "familysymbol":
			AddSignatureField(parts, korean ? "카테고리" : "Category", GetSignatureToken(tokens, 1));
			AddSignatureField(parts, korean ? "타입" : "Type", GetSignatureToken(tokens, 2));
			break;
		case "dimension":
			AddSignatureField(parts, korean ? "카테고리" : "Category", GetSignatureToken(tokens, 1));
			AddSignatureField(parts, korean ? "이름" : "Name", GetSignatureToken(tokens, 2));
			AddSignatureField(parts, korean ? "타입" : "Type", FirstNonEmpty(GetSignatureToken(tokens, 5), GetSignatureToken(tokens, 3)));
			break;
		case "material":
			AddSignatureField(parts, korean ? "재료" : "Material", GetSignatureToken(tokens, 2));
			break;
		case "element":
			AddSignatureField(parts, korean ? "카테고리" : "Category", GetSignatureToken(tokens, 1));
			AddSignatureField(parts, korean ? "이름" : "Name", GetSignatureToken(tokens, 2));
			AddSignatureField(parts, korean ? "타입" : "Type", GetSignatureToken(tokens, 5));
			break;
		default:
			AddSignatureField(parts, korean ? "카테고리" : "Category", GetSignatureToken(tokens, 1));
			AddSignatureField(parts, korean ? "이름" : "Name", GetSignatureToken(tokens, 2));
			AddSignatureField(parts, korean ? "타입" : "Type", GetSignatureToken(tokens, 5));
			break;
		}
		if (parts.Count == 0)
		{
			return string.Empty;
		}
		return string.Join(" · ", parts);
	}

	private static string GetSignatureToken(IList<string> tokens, int index)
	{
		if (tokens == null || index < 0 || index >= tokens.Count)
		{
			return string.Empty;
		}
		return tokens[index];
	}

	private static void AddSignatureField(IList<string> parts, string label, string value)
	{
		string text = CleanSignatureDisplayValue(value);
		if (!string.IsNullOrWhiteSpace(text) && !string.Equals(text, "-", StringComparison.Ordinal))
		{
			parts.Add(label + ": " + text);
		}
	}

	private static string CleanSignatureDisplayValue(string value)
	{
		string text = CleanDiffCell(value);
		if (string.IsNullOrWhiteSpace(text))
		{
			return string.Empty;
		}
		text = Regex.Replace(text, "\\s*/\\s*UID\\s+[^,]+", string.Empty, RegexOptions.IgnoreCase);
		text = Regex.Replace(text, "\\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\\b", string.Empty, RegexOptions.IgnoreCase);
		text = Regex.Replace(text, "^(dimensiontype|familytype|familysymbol|elementtype|material|element)\\s*/+\\s*", string.Empty, RegexOptions.IgnoreCase);
		text = text.Replace("//", "/").Replace("|", " · ");
		text = Regex.Replace(text, "\\s*/\\s*", "/");
		text = Regex.Replace(text, "\\s{2,}", " ");
		text = text.Trim().Trim(new char[3] { '/', ',', '·' });
		return text.Trim();
	}

	private static string ExtractSignatureDebugIds(string value)
	{
		string text = value ?? string.Empty;
		if (string.IsNullOrWhiteSpace(text))
		{
			return string.Empty;
		}
		MatchCollection matches = Regex.Matches(text, "ID\\s+([0-9]+)", RegexOptions.IgnoreCase);
		if (matches == null || matches.Count == 0)
		{
			return string.Empty;
		}
		List<string> ids = new List<string>();
		foreach (Match match in matches)
		{
			if (match != null && match.Groups.Count >= 2)
			{
				string idValue = match.Groups[1].Value;
				if (!string.IsNullOrWhiteSpace(idValue) && !ids.Contains<string>(idValue, StringComparer.OrdinalIgnoreCase))
				{
					ids.Add(idValue);
				}
			}
		}
		string moreText = string.Empty;
		Match moreMatch = Regex.Match(text, "\\+([0-9]+)", RegexOptions.IgnoreCase);
		if (moreMatch.Success)
		{
			moreText = " +" + moreMatch.Groups[1].Value;
		}
		return string.Join(", ", ids.Take(4)) + moreText;
	}

	private static string SignatureDifferenceNote(LoadableFingerprintDifferenceDetailItem detail, bool korean)
	{
		return NormalizeDiffToken((detail == null) ? string.Empty : detail.DifferenceKind) switch
		{
			"standardonly" => korean ? "표준에만 있음" : "Only in standard", 
			"projectonly" => korean ? "프로젝트에만 있음" : "Only in project", 
			"modified" => korean ? "값 다름" : "Value differs", 
			"omitted" => ShortDiffValue((detail == null) ? string.Empty : detail.Details, 60), 
			_ => ProjectComparisonReviewExcelExportService.SignatureDifferenceNote((detail == null) ? string.Empty : detail.Details, korean), 
		};
	}

	private static string SignatureDifferenceNote(string details, bool korean)
	{
		int standardCount = ExtractSignatureSideCount(details, "standard-only ");
		int projectCount = ExtractSignatureSideCount(details, "project-only ");
		List<string> parts = new List<string>();
		if (standardCount > 0)
		{
			parts.Add((korean ? "표준에만 " : "standard only ") + standardCount.ToString(CultureInfo.InvariantCulture));
		}
		if (projectCount > 0)
		{
			parts.Add((korean ? "프로젝트에만 " : "project only ") + projectCount.ToString(CultureInfo.InvariantCulture));
		}
		if (parts.Count == 0)
		{
			return string.Empty;
		}
		return ShortDiffValue(string.Join(" · ", parts), 60);
	}

	private static string NestedLabelDiffCell(string value)
	{
		string text = CleanDiffCell(value);
		if (string.IsNullOrWhiteSpace(text) || string.Equals(text, "-", StringComparison.Ordinal))
		{
			return "-";
		}
		text = text.Replace("nested-labels=", string.Empty).Replace("=>", " Label ").Replace("|", " · ");
		text = Regex.Replace(text, "\\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\\b", string.Empty, RegexOptions.IgnoreCase);
		text = Regex.Replace(text, "\\s{2,}", " ");
		text = text.Replace("/ /", "/").Trim().Trim(new char[3] { '/', ',', '·' });
		if (string.IsNullOrWhiteSpace(text))
		{
			return "-";
		}
		return ShortDiffValue(text, 120);
	}

	private static string NestedLabelDiffNote(string details, bool korean)
	{
		string text = CleanDiffCell(details);
		if (string.IsNullOrWhiteSpace(text))
		{
			return korean ? "Label 다름" : "Label differs";
		}
		if (korean)
		{
			text = text.Replace("Nested label differs: ", string.Empty).Replace("Nested label differs.", "Label 다름").Replace("Nested label is missing in project.", "프로젝트 누락")
				.Replace("Nested label exists only in project.", "프로젝트에만 있음")
				.Replace("Label definition differs.", "Label 정의 다름")
				.Replace("Label name", "Label 이름")
				.Replace("type/instance", "타입/인스턴스")
				.Replace("storage", "저장형식")
				.Replace("formula", "공식");
		}
		return ShortDiffValue(text, 60);
	}

	private static string SignaturePresence(string value, bool korean)
	{
		if (IsDiffBlank(value))
		{
			return korean ? "없음" : "Missing";
		}
		if (string.Equals(CleanDiffCell(value), "readable", StringComparison.OrdinalIgnoreCase))
		{
			return korean ? "있음" : "Present";
		}
		return korean ? "없음" : "Missing";
	}

	private static string SignatureCountCell(string details, bool standardSide, bool korean)
	{
		int count = ExtractSignatureSideCount(details, standardSide ? "standard-only " : "project-only ");
		if (count <= 0)
		{
			return "-";
		}
		return count.ToString(CultureInfo.InvariantCulture) + (korean ? "개" : " item");
	}

	private static int ExtractSignatureSideCount(string details, string marker)
	{
		string text = details ?? string.Empty;
		int index = text.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
		if (index < 0)
		{
			return 0;
		}
		checked
		{
			index += marker.Length;
			StringBuilder digits = new StringBuilder();
			for (; index < text.Length && char.IsDigit(text[index]); index++)
			{
				digits.Append(text[index]);
			}
			if (int.TryParse(digits.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
			{
				return value;
			}
			return 0;
		}
	}

	private static string BuildParameterCompactCell(string value, string changedLabels, bool korean)
	{
		List<string> segments = SplitDiffValue(value);
		if (segments.Count == 0)
		{
			return "-";
		}
		string name = segments[0];
		if (string.IsNullOrWhiteSpace(changedLabels))
		{
			return ShortDiffValue(name, 54);
		}
		string labelKey = NormalizeDiffToken(changedLabels);
		if (labelKey.IndexOf(NormalizeDiffToken(korean ? "공식" : "Formula"), StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return ShortDiffValue(name + " / " + FirstNonEmpty(GetParameterFormula(segments), korean ? "공식 없음" : "No formula"), 72);
		}
		if (labelKey.IndexOf(NormalizeDiffToken(korean ? "공유" : "Shared"), StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return ShortDiffValue(name + " / " + GetParameterSharedLabel(segments, korean), 72);
		}
		if (labelKey.IndexOf("guid", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return ShortDiffValue(name + " / " + GetParameterGuidLabel(segments, korean), 72);
		}
		if (labelKey.IndexOf(NormalizeDiffToken(korean ? "타입/인스턴스" : "Type/Instance"), StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return ShortDiffValue(name + " / " + GetParameterRoleLabel(segments, korean), 72);
		}
		if (labelKey.IndexOf(NormalizeDiffToken(korean ? "자료형" : "Data Type"), StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return ShortDiffValue(name + " / " + GetParameterStorageLabel(segments), 72);
		}
		return ShortDiffValue(name, 54);
	}

	private static string BuildParameterChangedLabelText(string standardValue, string projectValue, bool korean)
	{
		List<string> standardSegments = SplitDiffValue(standardValue);
		List<string> projectSegments = SplitDiffValue(projectValue);
		if (standardSegments.Count == 0 || projectSegments.Count == 0)
		{
			return string.Empty;
		}
		List<string> labels = new List<string>();
		if (!string.Equals(GetParameterFormula(standardSegments), GetParameterFormula(projectSegments), StringComparison.OrdinalIgnoreCase))
		{
			labels.Add(korean ? "공식" : "Formula");
		}
		if (!string.Equals(GetParameterSharedToken(standardSegments), GetParameterSharedToken(projectSegments), StringComparison.OrdinalIgnoreCase))
		{
			labels.Add(korean ? "공유" : "Shared");
		}
		if (!string.Equals(GetParameterGuidToken(standardSegments), GetParameterGuidToken(projectSegments), StringComparison.OrdinalIgnoreCase))
		{
			labels.Add("GUID");
		}
		if (!string.Equals(GetParameterRoleToken(standardSegments), GetParameterRoleToken(projectSegments), StringComparison.OrdinalIgnoreCase))
		{
			labels.Add(korean ? "타입/인스턴스" : "Type/Instance");
		}
		if (!string.Equals(GetParameterStorageLabel(standardSegments), GetParameterStorageLabel(projectSegments), StringComparison.OrdinalIgnoreCase))
		{
			labels.Add(korean ? "자료형" : "Data Type");
		}
		if (labels.Count == 0)
		{
			labels.Add(korean ? "정의" : "Definition");
		}
		return string.Join(", ", labels.Take(2));
	}

	private static List<string> SplitDiffValue(string value)
	{
		string text = CleanDiffCell(value);
		if (string.IsNullOrWhiteSpace(text) || string.Equals(text, "-", StringComparison.Ordinal))
		{
			return new List<string>();
		}
		return (from x in text.Split(new string[1] { " / " }, StringSplitOptions.None)
			select x.Trim() into x
			where x.Length > 0
			select x).ToList();
	}

	private static string GetParameterRoleToken(IList<string> segments)
	{
		if (segments == null || segments.Count <= 2)
		{
			return string.Empty;
		}
		return segments[2];
	}

	private static string GetParameterRoleLabel(IList<string> segments, bool korean)
	{
		if (NormalizeDiffToken(GetParameterRoleToken(segments)).IndexOf("instance", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return korean ? "인스턴스" : "Instance";
		}
		return korean ? "타입" : "Type";
	}

	private static string GetParameterSharedToken(IList<string> segments)
	{
		if (segments == null || segments.Count <= 3)
		{
			return string.Empty;
		}
		return segments[3];
	}

	private static string GetParameterSharedLabel(IList<string> segments, bool korean)
	{
		if (NormalizeDiffToken(GetParameterSharedToken(segments)).IndexOf("shared", StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return korean ? "공유" : "Shared";
		}
		return korean ? "패밀리" : "Family";
	}

	private static string GetParameterStorageLabel(IList<string> segments)
	{
		if (segments == null || segments.Count <= 4)
		{
			return string.Empty;
		}
		return segments[4];
	}

	private static string GetParameterGuidToken(IList<string> segments)
	{
		if (segments == null)
		{
			return string.Empty;
		}
		foreach (string segment in segments)
		{
			if (segment != null && segment.Trim().StartsWith("guid:", StringComparison.OrdinalIgnoreCase))
			{
				return segment.Trim();
			}
		}
		return string.Empty;
	}

	private static string GetParameterGuidLabel(IList<string> segments, bool korean)
	{
		if (string.IsNullOrWhiteSpace(GetParameterGuidToken(segments)))
		{
			return korean ? "GUID 없음" : "No GUID";
		}
		return "GUID";
	}

	private static string GetParameterFormula(IList<string> segments)
	{
		if (segments == null)
		{
			return string.Empty;
		}
		foreach (string segment in segments)
		{
			if (segment != null)
			{
				string text = segment.Trim();
				if (text.StartsWith("formula=", StringComparison.OrdinalIgnoreCase))
				{
					return text.Substring("formula=".Length).Trim();
				}
			}
		}
		return string.Empty;
	}

	private static string DifferenceDetailOrPair(LoadableFingerprintDifferenceDetailItem detail, bool korean)
	{
		if (detail != null && !string.IsNullOrWhiteSpace(detail.Details))
		{
			return TranslateExportDetail(detail.Details, korean);
		}
		return StandardProjectPair((detail == null) ? string.Empty : detail.StandardValue, (detail == null) ? string.Empty : detail.ProjectValue, korean, 44);
	}

	private static string DifferenceAreaLabel(string area, string kind, bool korean)
	{
		string normalizedArea = Normalize(area);
		if (normalizedArea.StartsWith("elements/", StringComparison.Ordinal))
		{
			return (korean ? "요소 - " : "Element - ") + ElementSignatureKindLabel(normalizedArea.Substring("elements/".Length), korean);
		}
		return NormalizeDiffToken(area) switch
		{
			"familytypes" => korean ? "타입 구성" : "Type Inventory", 
			"parametersformulas" => korean ? "파라미터 / 공식" : "Parameters / Formulas", 
			"category" => korean ? "카테고리" : "Category", 
			"categorygroup" => korean ? "카테고리 그룹" : "Category Group", 
			"familyname" => korean ? "패밀리명" : "Family Name", 
			"sharedflag" => korean ? "공유 여부" : "Shared Flag", 
			"contentfingerprint" => korean ? "Content Fingerprint" : "Content Fingerprint", 
			"storedfingerprint" => korean ? "저장된 Fingerprint" : "Stored Fingerprint", 
			"signaturediagnostics" => korean ? "Signature 진단" : "Signature Diagnostics", 
			"connectors" => korean ? "커넥터" : "Connectors", 
			"geometry" => korean ? "형상" : "Geometry", 
			"nestedlabels" => korean ? "하위 패밀리 Label" : "Nested Labels", 
			"nestedloadableinstances" => korean ? "하위/로드 패밀리 인스턴스" : "Nested / Loadable Instances", 
			_ => FirstNonEmpty(area, korean ? "차이" : "Difference"), 
		};
	}

	private static string ElementSignatureKindLabel(string value, bool korean)
	{
		return Normalize(value) switch
		{
			"dimension" => korean ? "치수선" : "Dimension", 
			"reference plane" => korean ? "참조 평면" : "Reference Plane", 
			"model line" => korean ? "모델 선" : "Model Line", 
			"detail line" => korean ? "상세 선" : "Detail Line", 
			"text" => korean ? "문자" : "Text", 
			"filled region" => korean ? "채우기 영역" : "Filled Region", 
			"form" => korean ? "형상 요소" : "Form", 
			_ => HumanizeDiffToken(value), 
		};
	}

	private static string HumanizeDiffToken(string value)
	{
		string text = (value ?? string.Empty).Trim();
		if (text.Length == 0)
		{
			return string.Empty;
		}
		text = text.Replace("_", " ").Replace("-", " ").Replace("/", " ");
		StringBuilder builder = new StringBuilder();
		checked
		{
			int num = text.Length - 1;
			for (int index = 0; index <= num; index++)
			{
				char ch = text[index];
				if (index > 0 && char.IsUpper(ch) && !char.IsWhiteSpace(text[index - 1]))
				{
					builder.Append(' ');
				}
				builder.Append(ch);
			}
			return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(builder.ToString().ToLowerInvariant());
		}
	}

	private static string StandardProjectPair(string standardValue, string projectValue, bool korean, int maxLength)
	{
		return (korean ? "표준 " : "standard ") + ShortDiffValue(standardValue, maxLength) + (korean ? ", 프로젝트 " : ", project ") + ShortDiffValue(projectValue, maxLength);
	}

	private static string ExtractDiffName(string value)
	{
		string text = CleanDiffCell(value);
		if (string.IsNullOrWhiteSpace(text) || string.Equals(text, "-", StringComparison.Ordinal))
		{
			return string.Empty;
		}
		string[] segments = text.Split(new string[1] { " / " }, StringSplitOptions.None);
		if (segments.Length == 0)
		{
			return text;
		}
		return segments[0].Trim();
	}

	private static string TranslateExportDetail(string value, bool korean)
	{
		string text = CleanDiffCell(value);
		if (!korean)
		{
			return text;
		}
		return text.Replace("standard-only", "표준에만 있음").Replace("project-only", "프로젝트에만 있음").Replace("additional signature difference groups omitted.", "개의 추가 signature 차이 그룹 생략");
	}

	private static string CleanDiffCell(string value)
	{
		return (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Replace("\t", " ")
			.Trim();
	}

	private static string ShortDiffValue(string value, int maxLength)
	{
		string text = CleanDiffCell(value);
		if (string.IsNullOrWhiteSpace(text))
		{
			return "-";
		}
		if (maxLength <= 3 || text.Length <= maxLength)
		{
			return text;
		}
		return text.Substring(0, checked(maxLength - 3)) + "...";
	}

	private static string NormalizeDiffToken(string value)
	{
		return Normalize(value).Replace("/", string.Empty).Replace("\\", string.Empty).Replace(" ", string.Empty)
			.Replace("-", string.Empty)
			.Replace("_", string.Empty)
			.Replace(".", string.Empty);
	}

	private static bool IsReviewExportStatus(string status)
	{
		return ReviewStatuses.Contains((status ?? string.Empty).Trim());
	}

	private static string BuildLoadableNotes(LoadableFamilyComparisonItem item, bool korean)
	{
		List<string> parts = new List<string>();
		if (HasProjectFingerprintMissing(item))
		{
			AddIfNotEmpty(parts, korean ? "프로젝트 패밀리 Fingerprint가 생성되지 않았습니다." : "Project family fingerprint was not created.");
			if (!string.IsNullOrWhiteSpace(item.ProjectContentFingerprintFailureReason))
			{
				AddIfNotEmpty(parts, korean ? ("프로젝트 패밀리 Fingerprint가 생성되지 않았습니다: " + item.ProjectContentFingerprintFailureReason) : ("Project family fingerprint was not created: " + item.ProjectContentFingerprintFailureReason));
			}
		}
		else
		{
			AddIfNotEmpty(parts, StatusReason(item.Status, korean));
		}
		AddIfNotEmpty(parts, item.Notes);
		if (item.FingerprintDifferenceDetails != null)
		{
			foreach (LoadableFingerprintDifferenceDetailItem detail in item.FingerprintDifferenceDetails.Take(3))
			{
				AddIfNotEmpty(parts, DifferenceBrief(detail, korean));
			}
		}
		if (item.MissingTypeNames != null && item.MissingTypeNames.Count > 0)
		{
			AddIfNotEmpty(parts, (korean ? "누락 타입: " : "Missing type: ") + string.Join(", ", item.MissingTypeNames.Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)).Take(8)));
		}
		if (item.ExtraTypeNames != null && item.ExtraTypeNames.Count > 0)
		{
			AddIfNotEmpty(parts, (korean ? "추가 타입: " : "Extra type: ") + string.Join(", ", item.ExtraTypeNames.Where([SpecialName] (string x) => !string.IsNullOrWhiteSpace(x)).Take(8)));
		}
		return JoinDistinct(parts);
	}

	private static bool HasProjectFingerprintMissing(LoadableFamilyComparisonItem item)
	{
		if (item == null)
		{
			return false;
		}
		return !string.IsNullOrWhiteSpace(item.StandardContentFingerprint) && string.IsNullOrWhiteSpace(item.ProjectContentFingerprint);
	}

	private static string BuildSystemNotes(SystemTypeComparisonItem item, bool korean)
	{
		List<string> parts = new List<string>();
		AddIfNotEmpty(parts, StatusReason(item.Status, korean));
		AddIfNotEmpty(parts, item.Notes);
		if (item.Layers != null && item.Layers.Count > 0)
		{
			string layerText = string.Join(" | ", (from x in item.Layers.Select([SpecialName] (StandardSystemTypeLayerSnapshotItem layer) =>
				{
					if (layer == null)
					{
						return string.Empty;
					}
					string text = ((layer.Index > 0) ? (layer.Index.ToString(CultureInfo.InvariantCulture) + ". ") : string.Empty);
					string text2 = FirstNonEmpty(layer.MaterialName, layer.FunctionName, "-");
					string text3 = ((!string.IsNullOrWhiteSpace(layer.ThicknessDisplay)) ? (" / " + layer.ThicknessDisplay) : ((layer.ThicknessFeet > 0.0) ? (" / " + layer.ThicknessFeet.ToString("0.###", CultureInfo.InvariantCulture) + " ft") : string.Empty));
					return text + text2 + text3;
				})
				where !string.IsNullOrWhiteSpace(x)
				select x).Take(20));
			AddIfNotEmpty(parts, (korean ? "Layer: " : "Layers: ") + layerText);
		}
		return JoinDistinct(parts);
	}

	private static string BuildSignatureFailureNotes(ProjectLoadableSignatureFailureItem item, bool korean)
	{
		List<string> parts = new List<string>();
		AddIfNotEmpty(parts, korean ? "프로젝트 패밀리 Fingerprint 또는 Signature 진단 파일이 생성되지 않았습니다." : "Project family fingerprint or signature diagnostic file was not created.");
		AddIfNotEmpty(parts, (korean ? "실패 종류: " : "Failure kind: ") + (item.FailureKind ?? string.Empty));
		AddIfNotEmpty(parts, (korean ? "사유: " : "Reason: ") + (item.Reason ?? string.Empty));
		return JoinDistinct(parts);
	}

	private static string StatusReason(string status, bool korean)
	{
		switch (Normalize(status))
		{
		case "differentfromstandard":
			return korean ? "차이 있음" : "Different";
		case "loadedwithoutversionstamp":
			return korean ? "Family Browser 추적 스탬프가 없습니다." : "Family Browser tracking stamp is missing.";
		case "stampnormalizationneeded":
			return korean ? "내용은 표준과 같지만 추적 스탬프 갱신이 필요합니다." : "Content matches the standard, but the tracking stamp needs refresh.";
		case "updateavailable":
			return korean ? "프로젝트는 이전 승인본과 같지만 표준이 변경되었습니다." : "Project matches the previous approved version, but the standard changed.";
		case "locallymodified":
			return korean ? "표준 승인 이후 프로젝트에서 로컬 수정이 감지되었습니다." : "Project content drifted after the standard approval.";
		case "versionconflict":
			return korean ? "프로젝트, 승인 스탬프, 현재 표준 정보가 서로 일치하지 않습니다." : "Project, approval stamp, and current standard are inconsistent.";
		case "categorymismatch":
			return korean ? "같은 이름의 항목이 다른 카테고리로 존재합니다." : "Same-name item exists under a different category.";
		case "projectonly":
			return korean ? "표준 목록에 없는 프로젝트 전용 항목입니다." : "Project-only item is not in the selected standard.";
		case "manualreview":
		case "reviewneeded":
		case "approvalrequired":
			return korean ? "수동 검토가 필요한 항목입니다." : "Manual review is required.";
		default:
			return string.Empty;
		}
	}

	private static string DisplayStatus(string status, bool korean)
	{
		return Normalize(status) switch
		{
			"differentfromstandard" => korean ? "표준과 다름" : "Different From Standard", 
			"loadedwithoutversionstamp" => korean ? "스탬프 없음" : "Stamp Missing", 
			"stampnormalizationneeded" => korean ? "스탬프 갱신 필요" : "Stamp Refresh Needed", 
			"updateavailable" => korean ? "업데이트 가능" : "Update Available", 
			"locallymodified" => korean ? "로컬 수정됨" : "Locally Modified", 
			"versionconflict" => korean ? "버전 충돌" : "Version Conflict", 
			"categorymismatch" => korean ? "카테고리 불일치" : "Category Mismatch", 
			"projectonly" => korean ? "검토 필요" : "Review Needed", 
			"manualreview" => korean ? "수동 검토" : "Manual Review", 
			"reviewneeded" => korean ? "검토 필요" : "Review Needed", 
			"approvalrequired" => korean ? "승인 필요" : "Approval Required", 
			_ => status ?? string.Empty, 
		};
	}

	private static void AddIfNotEmpty(List<string> parts, string value)
	{
		if (parts != null && !string.IsNullOrWhiteSpace(value))
		{
			parts.Add(value.Trim());
		}
	}

	private static string JoinDistinct(IEnumerable<string> parts)
	{
		if (parts == null)
		{
			return string.Empty;
		}
		return string.Join(" | ", (from x in parts
			where !string.IsNullOrWhiteSpace(x)
			select x.Trim()).Distinct<string>(StringComparer.OrdinalIgnoreCase));
	}

	private static string FirstNonEmpty(params string[] values)
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

	private static string Normalize(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Trim().ToLowerInvariant();
	}

	private static string EnsureWorkbookOutputPath(string outputPath)
	{
		if (string.IsNullOrWhiteSpace(outputPath))
		{
			return outputPath;
		}
		if (string.Equals(Path.GetExtension(outputPath), ".xlsx", StringComparison.OrdinalIgnoreCase))
		{
			return outputPath;
		}
		return Path.ChangeExtension(outputPath, ".xlsx");
	}

	private static void WriteWorkbook(string outputPath, string sheetName, List<string> headers, List<List<string>> rows, string dialogSheetName, List<string> dialogHeaders, List<List<string>> dialogRows)
	{
		if (string.IsNullOrWhiteSpace(outputPath))
		{
			throw new ArgumentException("Output Excel path is empty.", "outputPath");
		}
		string outputFolder = Path.GetDirectoryName(outputPath);
		if (!string.IsNullOrWhiteSpace(outputFolder))
		{
			Directory.CreateDirectory(outputFolder);
		}
		if (File.Exists(outputPath))
		{
			File.Delete(outputPath);
		}
		using FileStream stream = new FileStream(outputPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
		using ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create);
		bool includeDialogSheet = dialogRows != null && dialogRows.Count > 0;
		AddEntry(archive, "[Content_Types].xml", BuildContentTypesXml(includeDialogSheet));
		AddEntry(archive, "_rels/.rels", BuildRootRelationshipsXml());
		AddEntry(archive, "docProps/app.xml", BuildAppXml());
		AddEntry(archive, "docProps/core.xml", BuildCoreXml());
		AddEntry(archive, "xl/workbook.xml", BuildWorkbookXml(sheetName, dialogSheetName, includeDialogSheet));
		AddEntry(archive, "xl/_rels/workbook.xml.rels", BuildWorkbookRelationshipsXml(includeDialogSheet));
		AddEntry(archive, "xl/styles.xml", BuildStylesXml());
		AddEntry(archive, "xl/worksheets/sheet1.xml", BuildWorksheetXml(headers, rows));
		if (includeDialogSheet)
		{
			AddEntry(archive, "xl/worksheets/sheet2.xml", BuildWorksheetXml(dialogHeaders, dialogRows));
		}
	}

	private static void AddEntry(ZipArchive archive, string entryName, string content)
	{
		ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
		using StreamWriter writer = new StreamWriter(entry.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
		writer.Write(content);
	}

	private static string BuildContentTypesXml(bool includeDialogSheet)
	{
		string secondSheet = includeDialogSheet ? "<Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" : string.Empty;
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" + secondSheet + "</Types>";
	}

	private static string BuildRootRelationshipsXml()
	{
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/></Relationships>";
	}

	private static string BuildWorkbookXml(string sheetName, string dialogSheetName, bool includeDialogSheet)
	{
		string secondSheet = includeDialogSheet ? ("<sheet name=\"" + XmlEscape(string.IsNullOrWhiteSpace(dialogSheetName) ? "ScanDialogs" : dialogSheetName) + "\" sheetId=\"2\" r:id=\"rId2\"/>") : string.Empty;
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"" + XmlEscape(string.IsNullOrWhiteSpace(sheetName) ? "Review" : sheetName) + "\" sheetId=\"1\" r:id=\"rId1\"/>" + secondSheet + "</sheets></workbook>";
	}

	private static string BuildWorkbookRelationshipsXml(bool includeDialogSheet)
	{
		string secondSheet = includeDialogSheet ? "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>" : string.Empty;
		string stylesRelationshipId = includeDialogSheet ? "rId3" : "rId2";
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" + secondSheet + "<Relationship Id=\"" + stylesRelationshipId + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>";
	}

	private static string BuildStylesXml()
	{
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"2\"><font><sz val=\"10\"/><name val=\"Malgun Gothic\"/></font><font><b/><sz val=\"10\"/><name val=\"Malgun Gothic\"/></font></fonts><fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFEAF4EF\"/><bgColor indexed=\"64\"/></patternFill></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"3\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment wrapText=\"1\" vertical=\"top\"/></xf></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>";
	}

	private static string BuildAppXml()
	{
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application>KKY Family Browser</Application></Properties>";
	}

	private static string BuildCoreXml()
	{
		string stamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
		return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:creator>KKY Family Browser</dc:creator><cp:lastModifiedBy>KKY Family Browser</cp:lastModifiedBy><dcterms:created xsi:type=\"dcterms:W3CDTF\">" + stamp + "</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + stamp + "</dcterms:modified></cp:coreProperties>";
	}

	private static string BuildWorksheetXml(List<string> headers, List<List<string>> rows)
	{
		checked
		{
			int rowCount = (rows?.Count ?? 0) + 1;
			int colCount = headers?.Count ?? 0;
			string lastRef = ColumnName(Math.Max(1, colCount)) + Math.Max(1, rowCount).ToString(CultureInfo.InvariantCulture);
			StringBuilder builder = new StringBuilder();
			builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
			builder.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
			builder.Append("<dimension ref=\"A1:" + lastRef + "\"/>");
			builder.Append("<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>");
			builder.Append(BuildColumnsXml(colCount));
			builder.Append("<sheetData>");
			builder.Append(BuildRowXml(1, headers, 1));
			if (rows != null)
			{
				int num = rows.Count - 1;
				for (int i = 0; i <= num; i++)
				{
					builder.Append(BuildRowXml(i + 2, rows[i], 2));
				}
			}
			builder.Append("</sheetData>");
			builder.Append("<autoFilter ref=\"A1:" + lastRef + "\"/>");
			builder.Append("</worksheet>");
			return builder.ToString();
		}
	}

	private static string BuildColumnsXml(int colCount)
	{
		StringBuilder builder = new StringBuilder();
		builder.Append("<cols>");
		for (int i = 1; i <= colCount; i = checked(i + 1))
		{
			int width;
			switch (i)
			{
			case 1:
			case 2:
			case 3:
			case 4:
				width = 18;
				break;
			case 5:
			case 6:
			case 7:
				width = 30;
				break;
			case 8:
			case 9:
			case 10:
				width = 16;
				break;
			case 11:
				width = 86;
				break;
			case 12:
				width = 24;
				break;
			case 13:
			case 14:
				width = 42;
				break;
			case 15:
				width = 64;
				break;
			case 16:
			case 17:
			case 18:
				width = 42;
				break;
			default:
				width = 18;
				break;
			}
			builder.Append("<col min=\"" + i.ToString(CultureInfo.InvariantCulture) + "\" max=\"" + i.ToString(CultureInfo.InvariantCulture) + "\" width=\"" + width.ToString(CultureInfo.InvariantCulture) + "\" customWidth=\"1\"/>");
		}
		builder.Append("</cols>");
		return builder.ToString();
	}

	private static string BuildRowXml(int rowIndex, IList<string> values, int styleIndex)
	{
		StringBuilder builder = new StringBuilder();
		builder.Append("<row r=\"" + rowIndex.ToString(CultureInfo.InvariantCulture) + "\">");
		checked
		{
			if (values != null)
			{
				int num = values.Count - 1;
				for (int col = 0; col <= num; col++)
				{
					builder.Append(BuildCellXml(rowIndex, col + 1, values[col], styleIndex));
				}
			}
			builder.Append("</row>");
			return builder.ToString();
		}
	}

	private static string BuildCellXml(int rowIndex, int columnIndex, string value, int styleIndex)
	{
		string styleText = ((styleIndex > 0) ? (" s=\"" + styleIndex.ToString(CultureInfo.InvariantCulture) + "\"") : string.Empty);
		return "<c r=\"" + ColumnName(columnIndex) + rowIndex.ToString(CultureInfo.InvariantCulture) + "\" t=\"inlineStr\"" + styleText + "><is><t xml:space=\"preserve\">" + XmlEscape(value) + "</t></is></c>";
	}

	private static string ColumnName(int columnIndex)
	{
		int value = columnIndex;
		StringBuilder builder = new StringBuilder();
		while (value > 0)
		{
			checked
			{
				value--;
				builder.Insert(0, Strings.ChrW(65 + unchecked(value % 26)));
			}
			value /= 26;
		}
		return builder.ToString();
	}

	private static string XmlEscape(string value)
	{
		return SecurityElement.Escape(value ?? string.Empty);
	}
}
