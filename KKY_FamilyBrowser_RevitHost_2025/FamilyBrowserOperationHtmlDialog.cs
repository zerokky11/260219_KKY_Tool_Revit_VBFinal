using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

internal static class FamilyBrowserOperationHtmlDialog
{
	public static DialogResult ShowFamilyLoadConfirmation(IWin32Window owner, bool isKorean, string caption, string scopeText, bool isSelectedApply, bool allowFingerprintOverwrite, int skippedLatestFamilyCount, int skippedMissingFamilyCount, int loadCount, int reloadCount, int stampOnlyCount, int blockedCount, string categoryMismatchSummary)
	{
		List<DialogStat> stats = new List<DialogStat>();
		if (isSelectedApply)
		{
			stats.Add(new DialogStat(Tx(isKorean, "Load selected", "선택 로드"), loadCount, "good"));
			stats.Add(new DialogStat(Tx(isKorean, "Existing updates", "기존 업데이트"), Tx(isKorean, "Admin only", "관리자 설정"), "neutral"));
		}
		else
		{
			stats.Add(new DialogStat(Tx(isKorean, "Reload different", "변경 재로드"), reloadCount, "good"));
			stats.Add(new DialogStat(Tx(isKorean, "Tracking only", "추적만 갱신"), stampOnlyCount, "info"));
			stats.Add(new DialogStat(Tx(isKorean, "Latest kept", "최신 유지"), skippedLatestFamilyCount, "neutral"));
			stats.Add(new DialogStat(Tx(isKorean, "Missing skipped", "누락 건너뜀"), skippedMissingFamilyCount, "neutral"));
		}
		stats.Add(new DialogStat(Tx(isKorean, "Blocked", "차단"), blockedCount, blockedCount > 0 ? "warn" : "neutral"));
		List<string> notes = new List<string>();
		if (isSelectedApply)
		{
			notes.Add(Tx(isKorean, "Only selected missing standard families are loaded. Existing loaded families are updated from Admin Settings.", "선택한 누락 표준 패밀리만 로드합니다. 이미 로드된 기존 패밀리 업데이트는 관리자 설정에서 처리합니다."));
		}
		else
		{
			notes.Add(allowFingerprintOverwrite ? Tx(isKorean, "Existing project families will be updated from the registered standard. Missing families are not created here.", "현재 프로젝트에 이미 로드된 패밀리를 등록된 표준 기준으로 업데이트합니다. 누락 패밀리는 여기서 새로 만들지 않습니다.") : Tx(isKorean, "Standard family changes will be applied only to existing loaded families.", "기존에 로드된 패밀리에만 표준 패밀리 변경을 적용합니다."));
		}
		AddOptionalMultiline(notes, categoryMismatchSummary);
		string html = BuildHtml(isKorean, caption, isSelectedApply ? Tx(isKorean, "Load selected standard families?", "선택 표준 패밀리를 로드할까?") : Tx(isKorean, "Update existing standard families?", "기존 표준 패밀리를 업데이트할까?"), scopeText, stats, notes, null, false, MessageBoxIcon.Question);
		return ShowChoice(owner, isKorean, caption, html, MessageBoxIcon.Question, Tx(isKorean, "Apply", "적용"), Tx(isKorean, "Cancel", "취소"));
	}

	public static DialogResult ShowSystemTypeApplyConfirmation(IWin32Window owner, bool isKorean, string caption, string scopeText, int createCount, int skippedMissingSystemCount, int overwriteCount, int consolidateCount, int missingDependencyCount, int reloadDependencyCount, int unsupportedCount)
	{
		List<DialogStat> stats = new List<DialogStat>
		{
			new DialogStat(Tx(isKorean, "Create new", "신규 생성"), createCount, createCount > 0 ? "good" : "neutral"),
			new DialogStat(Tx(isKorean, "Overwrite", "덮어쓰기"), overwriteCount, overwriteCount > 0 ? "good" : "neutral"),
			new DialogStat(Tx(isKorean, "Consolidate", "중복 정리"), consolidateCount, consolidateCount > 0 ? "info" : "neutral"),
			new DialogStat(Tx(isKorean, "Dependency load", "의존 로드"), missingDependencyCount + reloadDependencyCount, (missingDependencyCount + reloadDependencyCount) > 0 ? "info" : "neutral"),
			new DialogStat(Tx(isKorean, "Skipped missing", "누락 건너뜀"), skippedMissingSystemCount, skippedMissingSystemCount > 0 ? "warn" : "neutral"),
			new DialogStat(Tx(isKorean, "Review only", "검토만"), unsupportedCount, unsupportedCount > 0 ? "warn" : "neutral")
		};
		List<string> notes = new List<string>
		{
			Tx(isKorean, "Required dependency families are loaded or reloaded from the registered standard RVT before system types are applied.", "필요한 의존 패밀리는 시스템 타입 적용 전에 등록된 표준 RVT에서 먼저 로드하거나 재로드합니다."),
			Tx(isKorean, "Unsupported system type changes remain review-only and will not mutate the model.", "자동 적용 미지원 시스템 타입 변경은 검토 항목으로 남기고 모델은 변경하지 않습니다.")
		};
		string html = BuildHtml(isKorean, caption, Tx(isKorean, "Apply standard system type changes?", "표준 시스템 타입 변경을 적용할까?"), scopeText, stats, notes, null, false, MessageBoxIcon.Question);
		return ShowChoice(owner, isKorean, caption, html, MessageBoxIcon.Question, Tx(isKorean, "Apply", "적용"), Tx(isKorean, "Cancel", "취소"));
	}

	public static DialogResult ShowFamilyLoadResult(IWin32Window owner, bool isKorean, string caption, string headlineText, string scopeText, LoadableFamilySyncExecutionReport report, string reportPath)
	{
		LoadableFamilySyncExecutionSummary summary = report == null ? null : report.Summary;
		bool hasIssue = summary != null && (summary.FailedCount > 0 || summary.BlockedCount > 0);
		List<DialogStat> stats = new List<DialogStat>
		{
			new DialogStat(Tx(isKorean, "Loaded", "로드됨"), summary == null ? 0 : summary.LoadedCount, "good"),
			new DialogStat(Tx(isKorean, "Reloaded", "재로드됨"), summary == null ? 0 : summary.ReloadedCount, "info"),
			new DialogStat(Tx(isKorean, "Tracking", "추적 갱신"), summary == null ? 0 : summary.TrackingRefreshedCount, "neutral"),
			new DialogStat(Tx(isKorean, "Skipped", "건너뜀"), summary == null ? 0 : summary.SkippedCount, "neutral"),
			new DialogStat(Tx(isKorean, "Blocked", "차단"), summary == null ? 0 : summary.BlockedCount, summary != null && summary.BlockedCount > 0 ? "warn" : "neutral"),
			new DialogStat(Tx(isKorean, "Failed", "실패"), summary == null ? 0 : summary.FailedCount, summary != null && summary.FailedCount > 0 ? "bad" : "neutral")
		};
		List<DialogRow> rows = new List<DialogRow>();
		foreach (LoadableFamilySyncExecutionItem item in OrderedFamilyItems(report))
		{
			rows.Add(new DialogRow(DisplayOutcome(isKorean, item.Outcome), DisplayFamilyAction(isKorean, item.ExecutionMode, item.PlannedAction), item.CategoryName, item.FamilyName, item.Details));
		}
		string html = BuildHtml(isKorean, caption, string.IsNullOrWhiteSpace(headlineText) ? Tx(isKorean, "Family load finished.", "패밀리 로드가 완료되었습니다.") : headlineText, scopeText, stats, new List<string> { Tx(isKorean, "The item table is sorted by issue severity, then category and family name.", "항목 표는 문제 심각도, 카테고리, 패밀리 이름 순으로 정렬됩니다.") }, rows, true, hasIssue ? MessageBoxIcon.Exclamation : MessageBoxIcon.Asterisk);
		return ShowResult(owner, isKorean, caption, html, hasIssue ? MessageBoxIcon.Exclamation : MessageBoxIcon.Asterisk, reportPath, rows, true, "KKY-FamilyBrowser-Family-Load-Result", "FamilyLoadResult");
	}

	public static DialogResult ShowSystemTypeApplyResult(IWin32Window owner, bool isKorean, string caption, string scopeText, SystemTypeApplyExecutionReport report, string reportPath)
	{
		SystemTypeApplyExecutionSummary summary = report == null ? null : report.Summary;
		bool hasIssue = summary != null && (summary.FailedCount > 0 || summary.BlockedCount > 0);
		List<DialogStat> stats = new List<DialogStat>
		{
			new DialogStat(Tx(isKorean, "Created", "생성"), summary == null ? 0 : summary.CreatedCount, "good"),
			new DialogStat(Tx(isKorean, "Overwritten", "덮어쓰기"), summary == null ? 0 : summary.OverwrittenCount, "info"),
			new DialogStat(Tx(isKorean, "Consolidated", "중복 정리"), summary == null ? 0 : summary.ConsolidatedCount, "info"),
			new DialogStat(Tx(isKorean, "Dependencies", "의존 갱신"), summary == null ? 0 : summary.DependencyLoadedCount, "neutral"),
			new DialogStat(Tx(isKorean, "Retyped elements", "타입 재지정 요소"), summary == null ? 0 : summary.RetypedElementCount, "neutral"),
			new DialogStat(Tx(isKorean, "Failed", "실패"), summary == null ? 0 : summary.FailedCount, summary != null && summary.FailedCount > 0 ? "bad" : "neutral")
		};
		List<string> notes = new List<string>();
		notes.Add((summary != null && summary.TrackingRefreshedCount > 0) ? Tx(isKorean, "Tracking was stamped after a clean apply.", "정상 적용 후 추적 스탬프를 기록했습니다.") : Tx(isKorean, "Tracking stamp was skipped because no clean mutation needed stamping or an issue remained.", "스탬프할 정상 변경이 없거나 문제가 남아 추적 스탬프를 건너뛰었습니다."));
		notes.Add(Tx(isKorean, "Post-apply review and diagnostic JSON paths are saved for admin diagnostics.", "적용 후 재검토와 진단 JSON 경로는 관리자 진단용으로 저장됩니다."));
		List<DialogRow> rows = new List<DialogRow>();
		foreach (SystemTypeApplyExecutionItem item in OrderedSystemItems(report))
		{
			string details = item.Details;
			if (string.IsNullOrWhiteSpace(details) && item.Messages != null && item.Messages.Count > 0)
			{
				details = item.Messages[item.Messages.Count - 1];
			}
			rows.Add(new DialogRow(DisplayOutcome(isKorean, item.Outcome), DisplaySystemAction(isKorean, item.SyncAction), item.CategoryName, item.SystemFamilyKind, item.SystemTypeName, details));
		}
		string html = BuildHtml(isKorean, caption, Tx(isKorean, "Standard system type apply completed.", "표준 시스템 타입 적용이 완료되었습니다."), scopeText, stats, notes, rows, false, hasIssue ? MessageBoxIcon.Exclamation : MessageBoxIcon.Asterisk);
		return ShowResult(owner, isKorean, caption, html, hasIssue ? MessageBoxIcon.Exclamation : MessageBoxIcon.Asterisk, reportPath, rows, false, "KKY-FamilyBrowser-System-Type-Apply-Result", "SystemTypeApply");
	}

	private static DialogResult ShowChoice(IWin32Window owner, bool isKorean, string caption, string html, MessageBoxIcon icon, string positiveText, string negativeText)
	{
		using (FamilyBrowserHtmlDialogHost dialog = new FamilyBrowserHtmlDialogHost(isKorean, caption, html, MessageBoxButtons.YesNo, icon, MessageBoxDefaultButton.Button1, positiveText, negativeText, string.Empty, string.Empty, new Size(1120, 760), new Size(820, 560), true))
		{
			return dialog.ShowDialog(owner);
		}
	}

	private static DialogResult ShowResult(IWin32Window owner, bool isKorean, string caption, string html, MessageBoxIcon icon, string reportPath, IList<DialogRow> rows, bool familyRows, string filePrefix, string sheetName)
	{
		using (FamilyBrowserHtmlDialogHost dialog = new FamilyBrowserHtmlDialogHost(isKorean, caption, html, MessageBoxButtons.OK, icon, MessageBoxDefaultButton.Button1, string.Empty, string.Empty, string.Empty, reportPath, new Size(1120, 760), new Size(820, 560), true, Tx(isKorean, "Export Excel", "Excel 내보내기")))
		{
			dialog.AuxiliaryActionRequested += delegate
			{
				ExportResultRows(dialog, dialog, isKorean, rows, familyRows, filePrefix, sheetName);
			};
			return dialog.ShowDialog(owner);
		}
	}

	private static void ExportResultRows(IWin32Window owner, FamilyBrowserHtmlDialogHost host, bool isKorean, IList<DialogRow> rows, bool familyRows, string filePrefix, string sheetName)
	{
		List<string> headers = familyRows
			? new List<string> { Tx(isKorean, "Result", "결과"), Tx(isKorean, "Action", "작업"), Tx(isKorean, "Category", "카테고리"), Tx(isKorean, "Family", "패밀리"), Tx(isKorean, "Details", "상세") }
			: new List<string> { Tx(isKorean, "Result", "결과"), Tx(isKorean, "Action", "작업"), Tx(isKorean, "Category", "카테고리"), Tx(isKorean, "System family", "시스템 패밀리"), Tx(isKorean, "System type", "시스템 타입"), Tx(isKorean, "Details", "상세") };
		List<List<string>> exportRows = new List<List<string>>();
		foreach (DialogRow row in rows ?? new List<DialogRow>())
		{
			List<string> values = new List<string> { row.Result, row.Action, row.Category };
			if (!familyRows)
			{
				values.Add(row.Kind);
			}
			values.Add(row.Name);
			values.Add(row.Details);
			exportRows.Add(values);
		}
		FamilyBrowserResultExcelExportUi.SaveRows(owner, host, isKorean, FamilyBrowserResultExcelExportUi.TimestampedFileName(filePrefix), sheetName, headers, exportRows);
	}

	private static string BuildHtml(bool isKorean, string caption, string headline, string scopeText, IList<DialogStat> stats, IList<string> notes, IList<DialogRow> rows, bool familyRows, MessageBoxIcon icon)
	{
		FamilyBrowserUiTheme theme = FamilyBrowserUiThemeService.Load();
		StringBuilder sb = new StringBuilder();
		sb.AppendLine("<!doctype html><html><head><meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\"/><meta charset=\"utf-8\"/>");
		sb.AppendLine("<style>");
		sb.AppendLine("html,body{margin:0;padding:0;background:#f5f7fb;color:#111827;font-family:'Malgun Gothic','Segoe UI',sans-serif;font-size:14px;}");
		sb.AppendLine(".wrap{padding:22px 24px 26px 24px;}.hero{border:1px solid #d6dde9;border-left:5px solid " + AccentHex(icon) + ";background:#fff;padding:18px 20px;margin-bottom:16px;box-shadow:0 2px 8px rgba(17,24,39,.07);}.eyebrow{font-size:12px;color:#64748b;font-weight:700;margin-bottom:5px;}h1{font-size:21px;line-height:1.25;margin:0;color:#111827;}p.scope{white-space:pre-wrap;margin:9px 0 0 0;color:#475569;font-size:13px;}");
		sb.AppendLine(".stats{margin:0 -6px 16px -6px;}.stat{display:inline-block;vertical-align:top;width:15.8%;min-width:118px;margin:0 6px 10px 6px;background:#fff;border:1px solid #d6dde9;border-top:4px solid #8aa3cc;padding:12px 12px 11px 12px;box-sizing:border-box;}.stat.good{border-top-color:#16845d}.stat.info{border-top-color:#2f6bff}.stat.warn{border-top-color:#d39a1b}.stat.bad{border-top-color:#c94d3e}.stat .label{font-size:12px;color:#64748b;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.stat .value{font-size:24px;line-height:1.1;font-weight:800;color:#111827;margin-top:7px;}");
		sb.AppendLine(".section{background:#fff;border:1px solid #d6dde9;margin-top:14px;padding:15px 16px;}h2{margin:0 0 10px 0;font-size:15px;color:#253b62;}ul{margin:0;padding-left:18px;}li{margin:4px 0;color:#475569;line-height:1.45;}.table-scroll{max-height:330px;overflow:auto;border:1px solid #cfd9ea;background:#fff;}table{width:100%;border-collapse:collapse;table-layout:fixed;}th{position:relative;background:#eef4ff;color:#273b5d;text-align:left;font-size:12px;padding:9px;border-bottom:1px solid #cfd9ea;}td{padding:9px;border-bottom:1px solid #e7edf7;vertical-align:top;word-break:break-word;line-height:1.35;}td.result{font-weight:800}.empty{padding:26px;text-align:center;color:#64748b;}.path-note{font-size:12px;color:#64748b;margin-top:10px;}");
		sb.AppendLine(FamilyBrowserUiThemeService.ThemeCss());
		sb.AppendLine("</style></head><body data-theme=\"" + Html(FamilyBrowserUiThemeService.Code(theme)) + "\" class=\"fb-operation-dialog " + Html(FamilyBrowserUiThemeService.BodyClass(theme)) + "\"><div class=\"wrap\">");
		sb.Append("<div class=\"hero\"><div class=\"eyebrow\">").Append(Html(caption)).Append("</div><h1>").Append(Html(headline)).Append("</h1>");
		if (!string.IsNullOrWhiteSpace(scopeText))
		{
			sb.Append("<p class=\"scope\">").Append(Html(NormalizeScope(scopeText))).Append("</p>");
		}
		sb.AppendLine("</div>");
		sb.AppendLine("<div class=\"stats\">");
		foreach (DialogStat stat in stats ?? new List<DialogStat>())
		{
			sb.Append("<div class=\"stat ").Append(Html(stat.Tone)).Append("\"><div class=\"label\">").Append(Html(stat.Label)).Append("</div><div class=\"value\">").Append(Html(stat.Value)).Append("</div></div>");
		}
		sb.AppendLine("</div>");
		if (notes != null && notes.Count > 0)
		{
			sb.Append("<div class=\"section\"><h2>").Append(Html(Tx(isKorean, "Review before continuing", "계속하기 전 확인"))).Append("</h2><ul>");
			foreach (string note in notes.Where(delegate(string x) { return !string.IsNullOrWhiteSpace(x); }))
			{
				sb.Append("<li>").Append(Html(note.Trim())).Append("</li>");
			}
			sb.AppendLine("</ul></div>");
		}
		if (rows != null)
		{
			sb.Append("<div class=\"section\"><h2>").Append(Html(Tx(isKorean, "Item result table", "항목 결과 표"))).Append("</h2><div class=\"table-scroll\">");
			if (rows.Count == 0)
			{
				sb.Append("<div class=\"empty\">").Append(Html(Tx(isKorean, "No item-level records were written.", "항목별 기록이 없습니다."))).Append("</div>");
			}
			else
			{
				sb.Append("<table><thead><tr><th style=\"width:92px\">").Append(Html(Tx(isKorean, "Result", "결과"))).Append("</th><th style=\"width:128px\">").Append(Html(Tx(isKorean, "Action", "작업"))).Append("</th><th style=\"width:160px\">").Append(Html(Tx(isKorean, "Category", "카테고리"))).Append("</th>");
				sb.Append(familyRows ? ("<th style=\"width:220px\">" + Html(Tx(isKorean, "Family", "패밀리")) + "</th>") : ("<th style=\"width:150px\">" + Html(Tx(isKorean, "System family", "시스템 패밀리")) + "</th><th style=\"width:220px\">" + Html(Tx(isKorean, "System type", "시스템 타입")) + "</th>"));
				sb.Append("<th>").Append(Html(Tx(isKorean, "Details", "상세"))).Append("</th></tr></thead><tbody>");
				foreach (DialogRow row in rows.Take(300))
				{
					sb.Append("<tr><td class=\"result\">").Append(Html(Blank(row.Result))).Append("</td><td>").Append(Html(Blank(row.Action))).Append("</td><td>").Append(Html(Blank(row.Category))).Append("</td>");
					if (familyRows)
					{
						sb.Append("<td>").Append(Html(Blank(row.Name))).Append("</td>");
					}
					else
					{
						sb.Append("<td>").Append(Html(Blank(row.Kind))).Append("</td><td>").Append(Html(Blank(row.Name))).Append("</td>");
					}
					sb.Append("<td>").Append(Html(Blank(row.Details))).Append("</td></tr>");
				}
				sb.Append("</tbody></table>");
			}
			sb.AppendLine("</div></div>");
		}
		sb.AppendLine("</div></body></html>");
		return sb.ToString();
	}

	private sealed class DialogStat
	{
		public string Label { get; private set; }
		public string Value { get; private set; }
		public string Tone { get; private set; }

		public DialogStat(string label, int value, string tone)
			: this(label, value.ToString(CultureInfo.InvariantCulture), tone)
		{
		}

		public DialogStat(string label, string value, string tone)
		{
			Label = label ?? string.Empty;
			Value = value ?? string.Empty;
			Tone = string.IsNullOrWhiteSpace(tone) ? "neutral" : tone;
		}
	}

	private sealed class DialogRow
	{
		public string Result { get; private set; }
		public string Action { get; private set; }
		public string Category { get; private set; }
		public string Kind { get; private set; }
		public string Name { get; private set; }
		public string Details { get; private set; }

		public DialogRow(string result, string action, string category, string name, string details)
			: this(result, action, category, string.Empty, name, details)
		{
		}

		public DialogRow(string result, string action, string category, string kind, string name, string details)
		{
			Result = result ?? string.Empty;
			Action = action ?? string.Empty;
			Category = category ?? string.Empty;
			Kind = kind ?? string.Empty;
			Name = name ?? string.Empty;
			Details = details ?? string.Empty;
		}
	}

	private static IEnumerable<LoadableFamilySyncExecutionItem> OrderedFamilyItems(LoadableFamilySyncExecutionReport report)
	{
		return ((report == null || report.Items == null) ? new List<LoadableFamilySyncExecutionItem>() : report.Items).Where(delegate(LoadableFamilySyncExecutionItem x) { return x != null; }).OrderBy(delegate(LoadableFamilySyncExecutionItem x) { return OutcomeSortKey(x.Outcome); }).ThenBy(delegate(LoadableFamilySyncExecutionItem x) { return x.CategoryName ?? string.Empty; }, StringComparer.OrdinalIgnoreCase).ThenBy(delegate(LoadableFamilySyncExecutionItem x) { return x.FamilyName ?? string.Empty; }, StringComparer.OrdinalIgnoreCase);
	}

	private static IEnumerable<SystemTypeApplyExecutionItem> OrderedSystemItems(SystemTypeApplyExecutionReport report)
	{
		return ((report == null || report.Items == null) ? new List<SystemTypeApplyExecutionItem>() : report.Items).Where(delegate(SystemTypeApplyExecutionItem x) { return x != null; }).OrderBy(delegate(SystemTypeApplyExecutionItem x) { return OutcomeSortKey(x.Outcome); }).ThenBy(delegate(SystemTypeApplyExecutionItem x) { return x.CategoryName ?? string.Empty; }, StringComparer.OrdinalIgnoreCase).ThenBy(delegate(SystemTypeApplyExecutionItem x) { return x.SystemTypeName ?? string.Empty; }, StringComparer.OrdinalIgnoreCase);
	}

	private static int OutcomeSortKey(string outcome)
	{
		switch (Normalize(outcome))
		{
		case "failed":
		case "error":
			return 0;
		case "blocked":
			return 1;
		case "skipped":
			return 2;
		case "loaded":
		case "reloaded":
		case "created":
		case "overwritten":
		case "consolidated":
			return 4;
		default:
			return 3;
		}
	}

	private static string DisplayOutcome(bool isKorean, string outcome)
	{
		switch (Normalize(outcome))
		{
		case "loaded":
			return Tx(isKorean, "Loaded", "로드됨");
		case "reloaded":
			return Tx(isKorean, "Reloaded", "재로드됨");
		case "created":
			return Tx(isKorean, "Created", "생성");
		case "overwritten":
			return Tx(isKorean, "Overwritten", "덮어쓰기");
		case "consolidated":
			return Tx(isKorean, "Consolidated", "중복 정리");
		case "skipped":
			return Tx(isKorean, "Skipped", "건너뜀");
		case "blocked":
			return Tx(isKorean, "Blocked", "차단");
		case "failed":
			return Tx(isKorean, "Failed", "실패");
		default:
			return string.IsNullOrWhiteSpace(outcome) ? "-" : outcome;
		}
	}

	private static string DisplayFamilyAction(bool isKorean, string executionMode, string plannedAction)
	{
		switch (Normalize(executionMode))
		{
		case "load":
			return Tx(isKorean, "Load new", "신규 로드");
		case "reload":
			return Tx(isKorean, "Reload", "재로드");
		case "stamponly":
			return Tx(isKorean, "Tracking only", "추적만");
		case "blocked":
			return Tx(isKorean, "Blocked", "차단");
		case "skip":
			return Tx(isKorean, "Skip", "건너뜀");
		default:
			return string.IsNullOrWhiteSpace(plannedAction) ? "-" : plannedAction;
		}
	}

	private static string DisplaySystemAction(bool isKorean, string action)
	{
		switch (Normalize(action))
		{
		case "createmissingtype":
			return Tx(isKorean, "Create new", "신규 생성");
		case "overwritedestination":
			return Tx(isKorean, "Overwrite", "덮어쓰기");
		case "consolidateduplicatesuffixtypes":
			return Tx(isKorean, "Consolidate duplicates", "중복 정리");
		case "keepcurrent":
			return Tx(isKorean, "Keep current", "현재 유지");
		case "skipmissingtype":
			return Tx(isKorean, "Skip missing", "누락 건너뜀");
		default:
			return string.IsNullOrWhiteSpace(action) ? "-" : action;
		}
	}

	private static void AddOptionalMultiline(ICollection<string> notes, string text)
	{
		if (notes == null || string.IsNullOrWhiteSpace(text))
		{
			return;
		}
		foreach (string line in text.Replace("\r\n", "\n").Replace("\r", "\n").Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries))
		{
			if (!string.IsNullOrWhiteSpace(line))
			{
				notes.Add(line.Trim());
			}
		}
	}

	private static string BuildReportHint(bool isKorean, string reportPath)
	{
		return string.IsNullOrWhiteSpace(reportPath) ? Tx(isKorean, "No report path was recorded.", "저장된 리포트 경로가 없습니다.") : Tx(isKorean, "Diagnostic report saved.", "진단 리포트가 저장되었습니다.") + " " + reportPath;
	}

	private static string AccentHex(MessageBoxIcon icon)
	{
		switch (icon)
		{
		case MessageBoxIcon.Exclamation:
			return "#d99a17";
		case MessageBoxIcon.Error:
			return "#c94d3e";
		case MessageBoxIcon.Question:
			return "#2f6bff";
		default:
			return "#2f6bff";
		}
	}

	private static string Tx(bool isKorean, string englishText, string koreanText)
	{
		return isKorean ? (koreanText ?? string.Empty) : (englishText ?? string.Empty);
	}

	private static string Html(string value)
	{
		if (value == null)
		{
			return string.Empty;
		}
		return value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
	}

	private static string Blank(string value)
	{
		return string.IsNullOrWhiteSpace(value) ? "-" : value.Trim();
	}

	private static string Normalize(string value)
	{
		return value == null ? string.Empty : value.Trim().ToLowerInvariant();
	}

	private static string NormalizeScope(string value)
	{
		return (value ?? string.Empty).Trim();
	}
}
