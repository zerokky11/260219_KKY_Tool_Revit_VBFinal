using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

internal static class FamilyBrowserMessageHtmlRenderer
{
	public static string Build(bool isKorean, string message, string caption, MessageBoxIcon icon, FamilyBrowserUiTheme? explicitTheme = null)
	{
		FamilyBrowserUiTheme theme = explicitTheme ?? FamilyBrowserUiThemeService.Load();
		ParsedMessage parsed = Parse(message, caption, isKorean);
		string kind = ResolveKind(icon);
		string accent = ResolveAccent(icon);
		string tint = ResolveTint(icon);
		StringBuilder html = new StringBuilder();
		html.AppendLine("<!doctype html><html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"/><meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\"/><meta charset=\"utf-8\"/>");
		html.AppendLine("<style>");
		html.AppendLine("*{box-sizing:border-box}html,body{margin:0;padding:0;background:#f5f7fb;color:#111827;font-family:'Malgun Gothic','Segoe UI',sans-serif;font-size:14px;}body{overflow-x:hidden;overflow-y:auto}.wrap{padding:18px 20px 24px 20px}.hero{background:#fff;border:1px solid #d6dde9;border-left:5px solid " + accent + ";box-shadow:0 2px 8px rgba(17,24,39,.07);padding:17px 18px;margin:0 0 13px 0}.hero-table{width:100%;border-collapse:collapse;table-layout:fixed}.hero-icon-cell{width:52px;vertical-align:top}.hero-icon{display:block;width:38px;height:38px;line-height:38px;text-align:center;color:#fff;background:" + accent + ";font-size:18px;font-weight:800;border-radius:6px}.hero-copy{vertical-align:middle}.eyebrow{color:#64748b;font-size:11px;font-weight:700;margin:0 0 4px 0}h1{margin:0;color:#111827;font-size:20px;line-height:1.35;font-weight:800;word-wrap:break-word}.intro{margin:9px 0 0 0;color:#475569;font-size:13px;line-height:1.55;white-space:pre-wrap;word-wrap:break-word}.result-groups{margin-top:12px}.result-group{background:#fff;border:1px solid #d6dde9;margin-top:11px;padding:14px 15px}.result-group-title{margin:0 0 11px 0;color:#253b62;font-size:14px;font-weight:800}.result-metrics{margin:0 -5px 6px -5px}.result-metric{display:inline-block;vertical-align:top;width:calc(25% - 10px);min-width:128px;margin:0 5px 10px 5px;padding:11px 12px;border:1px solid #d6dde9;border-top:4px solid #8aa3cc;background:#fff}.result-metric.good{border-top-color:#16845d}.result-metric.info{border-top-color:#2f6bff}.result-metric.warn{border-top-color:#d39a1b}.result-metric.bad{border-top-color:#c94d3e}.result-metric-label{color:#64748b;font-size:11px;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.result-metric-value{margin-top:6px;color:#111827;font-size:22px;line-height:1.1;font-weight:800}.result-info-table,.result-output-table{width:100%;border-collapse:collapse;table-layout:fixed;border:1px solid #dce3ee;background:#fff}.result-info-table th,.result-output-table th{width:180px;padding:8px 10px;border-bottom:1px solid #e5eaf2;background:#eef4ff;color:#475569;text-align:left;font-size:12px;font-weight:800;vertical-align:top}.result-info-table td,.result-output-table td{padding:8px 10px;border-bottom:1px solid #e5eaf2;color:#1f2937;vertical-align:top;word-wrap:break-word}.result-output-table{margin-top:10px}.result-output-table td{font-family:Consolas,'Malgun Gothic',monospace;font-size:12px;word-break:break-all}.result-notes{margin:10px 0 0 0;padding:0 0 0 19px;color:#475569}.result-notes li{margin:5px 0;line-height:1.5}.section{background:#fff;border:1px solid #d6dde9;border-left:4px solid #94a3b8;margin:11px 0 0 0;padding:13px 15px}.section.cause{border-left-color:#c95846;background:#fffdfc}.section.action{border-left-color:#2f6bff;background:#f8faff}.section.admin{border-left-color:#2f6bff;background:#f8faff}.section.technical{border-left-color:#64748b;background:#f8fafc}.section-title{font-size:12px;line-height:1.2;color:#475569;font-weight:800;margin:0 0 7px 0}.section.cause .section-title{color:#9d3f32}.section.action .section-title,.section.admin .section-title{color:#1f55d1}.section-copy{font-size:14px;line-height:1.55;color:#1f2937;white-space:pre-wrap;word-wrap:break-word}.section.admin .section-copy{word-break:break-all}.meta-list{margin-top:10px;border-top:1px solid #dce3ee}.meta-row{display:table;width:100%;table-layout:fixed;border-bottom:1px solid #e6ebf2}.meta-label,.meta-value{display:table-cell;padding:7px 8px;vertical-align:top}.meta-label{width:112px;color:#475569;font-size:12px;font-weight:700;background:#eef4ff}.meta-value{font-family:Consolas,'Malgun Gothic',monospace;font-size:12px;color:#1f2937;word-break:break-all;background:#fff}.technical-scroll{max-height:180px;overflow:auto;border:1px solid #d6dde9;background:#fff}.technical-scroll pre{margin:0;padding:11px 12px;color:#475569;font-family:Consolas,'Malgun Gothic',monospace;font-size:12px;line-height:1.45;white-space:pre;min-width:100%}.empty{color:#7a8d86}.footer-note{margin:12px 2px 0 2px;color:#64748b;font-size:11px;line-height:1.45}.tone-band{height:3px;background:" + tint + ";margin-top:14px}@media(max-width:760px){.result-metric{width:calc(50% - 10px)}.result-info-table th,.result-output-table th{width:132px}}@media(max-width:620px){.wrap{padding:14px}.hero{padding:14px}.hero-icon-cell{width:46px}h1{font-size:18px}.meta-label{width:92px}.result-metric{width:calc(100% - 10px)}}");
		html.AppendLine(FamilyBrowserUiThemeService.ThemeCss());
		html.AppendLine("</style></head>");
		html.Append("<body data-theme=\"").Append(Html(FamilyBrowserUiThemeService.Code(theme))).Append("\" data-message-kind=\"").Append(Html(kind)).Append("\" data-message-structured=\"").Append(parsed.Sections.Count > 0 || parsed.ResultGroups.Count > 0 ? "true" : "false").Append("\" data-message-auto-result=\"").Append(parsed.ResultGroups.Count > 0 ? "true" : "false").Append("\" class=\"fb-message-dialog ").Append(Html(FamilyBrowserUiThemeService.BodyClass(theme))).AppendLine("\"><div class=\"wrap\">");
		html.Append("<div class=\"hero\"><table class=\"hero-table\" role=\"presentation\"><tr><td class=\"hero-icon-cell\"><span class=\"hero-icon\">").Append(Html(ResolveGlyph(icon))).Append("</span></td><td class=\"hero-copy\"><div class=\"eyebrow\">").Append(Html(ResolveKindLabel(isKorean, icon))).Append("</div><h1 id=\"messageHeadline\">").Append(Html(parsed.Headline)).Append("</h1>");
		if (!string.IsNullOrWhiteSpace(parsed.Intro))
		{
			html.Append("<div id=\"messageIntro\" class=\"intro\">").Append(Html(parsed.Intro)).Append("</div>");
		}
		html.AppendLine("</td></tr></table></div>");
		AppendResultGroups(html, parsed.ResultGroups, isKorean);
		for (int i = 0; i < parsed.Sections.Count; i++)
		{
			AppendSection(html, parsed.Sections[i], isKorean);
		}
		if (parsed.Sections.Count == 0 && parsed.ResultGroups.Count == 0 && string.IsNullOrWhiteSpace(parsed.Intro))
		{
			html.Append("<div class=\"footer-note\">").Append(Html(isKorean ? "아래 버튼을 눌러 작업을 계속합니다." : "Use the button below to continue.")).AppendLine("</div>");
		}
		html.Append("<div class=\"tone-band\"></div></div></body></html>");
		return html.ToString();
	}

	public static bool ContainsStructuredSections(string message)
	{
		ParsedMessage parsed = Parse(message, string.Empty, true);
		return parsed.Sections.Count > 0 || parsed.ResultGroups.Count > 0;
	}

	public static string FindPrimaryOutputPath(string message)
	{
		string[] lines = NormalizeNewlines(message).Split(new char[] { '\n' });
		foreach (string rawLine in lines)
		{
			ResultItem item;
			if (TryParseResultItem(rawLine, out item) && IsResultPath(item.Label, item.Value))
			{
				string candidate = (item.Value ?? string.Empty).Trim().Trim('"');
				if (candidate.StartsWith("\\\\", StringComparison.Ordinal) || (candidate.Length > 2 && char.IsLetter(candidate[0]) && candidate[1] == ':' && (candidate[2] == '\\' || candidate[2] == '/')))
				{
					return candidate;
				}
			}
		}
		return string.Empty;
	}

	private static void AppendSection(StringBuilder html, MessageSection section, bool isKorean)
	{
		string sectionId = ResolveSectionElementId(section.Key);
		html.Append("<div id=\"").Append(sectionId).Append("\" class=\"section ").Append(Html(section.Tone)).Append("\"><div class=\"section-title\">").Append(Html(DisplayHeading(section.Key, isKorean))).Append("</div>");
		if (string.Equals(section.Key, "technical", StringComparison.OrdinalIgnoreCase))
		{
			html.Append("<div class=\"technical-scroll\"><pre id=\"messageTechnicalDetail\">").Append(Html(Blank(section.Body))).Append("</pre></div>");
		}
		else
		{
			AppendRegularSectionBody(html, section.Body, isKorean);
		}
		html.AppendLine("</div>");
	}

	private static void AppendRegularSectionBody(StringBuilder html, string body, bool isKorean)
	{
		string normalized = NormalizeNewlines(body);
		string[] lines = normalized.Split(new char[] { '\n' });
		StringBuilder copy = new StringBuilder();
		List<MetaLine> metadata = new List<MetaLine>();
		foreach (string rawLine in lines)
		{
			MetaLine meta;
			if (TryParseMetaLine(rawLine, isKorean, out meta))
			{
				metadata.Add(meta);
				continue;
			}
			if (copy.Length > 0)
			{
				copy.AppendLine();
			}
			copy.Append(rawLine ?? string.Empty);
		}
		string copyText = copy.ToString().Trim();
		if (!string.IsNullOrWhiteSpace(copyText))
		{
			html.Append("<div class=\"section-copy\">").Append(Html(copyText)).Append("</div>");
		}
		if (metadata.Count > 0)
		{
			html.Append("<div class=\"meta-list\">");
			foreach (MetaLine meta in metadata)
			{
				html.Append("<div class=\"meta-row\"><div class=\"meta-label\">").Append(Html(meta.Label)).Append("</div><div id=\"").Append(Html(meta.ElementId)).Append("\" class=\"meta-value\">").Append(Html(Blank(meta.Value))).Append("</div></div>");
			}
			html.Append("</div>");
		}
		if (string.IsNullOrWhiteSpace(copyText) && metadata.Count == 0)
		{
			html.Append("<div class=\"section-copy empty\">-</div>");
		}
	}

	private static bool TryParseMetaLine(string rawLine, bool isKorean, out MetaLine meta)
	{
		meta = null;
		string line = (rawLine ?? string.Empty).Trim();
		string value;
		if (TryReadAfterPrefix(line, "로그:", out value) || TryReadAfterPrefix(line, "Log:", out value))
		{
			meta = new MetaLine(isKorean ? "로그" : "Log", value, "messageLogPath");
			return true;
		}
		if (TryReadAfterPrefix(line, "지원 코드:", out value) || TryReadAfterPrefix(line, "Support code:", out value))
		{
			meta = new MetaLine(isKorean ? "지원 코드" : "Support code", value, "messageSupportCode");
			return true;
		}
		if (TryReadAfterPrefix(line, "경로:", out value) || TryReadAfterPrefix(line, "Path:", out value))
		{
			meta = new MetaLine(isKorean ? "경로" : "Path", value, "messagePath");
			return true;
		}
		return false;
	}

	private static bool TryReadAfterPrefix(string line, string prefix, out string value)
	{
		value = string.Empty;
		if (!line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
		{
			return false;
		}
		value = line.Substring(prefix.Length).Trim();
		return true;
	}

	private static ParsedMessage Parse(string message, string caption, bool isKorean)
	{
		string normalized = NormalizeNewlines(message);
		string[] lines = normalized.Split(new char[] { '\n' });
		int first = -1;
		for (int i = 0; i < lines.Length; i++)
		{
			if (!string.IsNullOrWhiteSpace(lines[i]))
			{
				first = i;
				break;
			}
		}
		ParsedMessage parsed = new ParsedMessage();
		if (first < 0)
		{
			parsed.Headline = string.IsNullOrWhiteSpace(caption) ? (isKorean ? "작업 결과" : "Action result") : caption.Trim();
			return parsed;
		}

		string firstKey;
		bool firstIsHeading = TryResolveHeading(lines[first], out firstKey);
		parsed.Headline = firstIsHeading ? (string.IsNullOrWhiteSpace(caption) ? (isKorean ? "작업 결과" : "Action result") : caption.Trim()) : lines[first].Trim();
		StringBuilder intro = new StringBuilder();
		MessageSection current = null;
		int start = firstIsHeading ? first : first + 1;
		for (int i = start; i < lines.Length; i++)
		{
			string key;
			if (TryResolveHeading(lines[i], out key))
			{
				current = new MessageSection(key, ResolveTone(key));
				parsed.Sections.Add(current);
				continue;
			}
			if (current != null)
			{
				AppendLine(current.Builder, lines[i]);
			}
			else
			{
				AppendLine(intro, lines[i]);
			}
		}
		parsed.Intro = intro.ToString().Trim();
		foreach (MessageSection section in parsed.Sections)
		{
			section.Body = section.Builder.ToString().Trim();
		}
		if (parsed.Sections.Count == 0)
		{
			AnalyzeAutomaticResult(parsed, isKorean);
		}
		return parsed;
	}

	private static void AppendResultGroups(StringBuilder html, IList<ResultGroup> groups, bool isKorean)
	{
		if (groups == null || groups.Count == 0)
		{
			return;
		}
		html.Append("<div id=\"messageResultGroups\" class=\"result-groups\">");
		bool metricIdWritten = false;
		bool contextIdWritten = false;
		bool outputIdWritten = false;
		for (int i = 0; i < groups.Count; i++)
		{
			ResultGroup group = groups[i];
			html.Append("<div class=\"result-group\"><div class=\"result-group-title\">").Append(Html(string.IsNullOrWhiteSpace(group.Title) ? (isKorean ? "결과 요약" : "Result summary") : group.Title)).Append("</div>");
			if (group.Metrics.Count > 0)
			{
				html.Append("<div").Append(metricIdWritten ? string.Empty : " id=\"messageMetricGrid\"").Append(" class=\"result-metrics\">");
				metricIdWritten = true;
				foreach (ResultItem item in group.Metrics)
				{
					html.Append("<div class=\"result-metric ").Append(Html(item.Tone)).Append("\"><div class=\"result-metric-label\" title=\"").Append(Html(item.Label)).Append("\">").Append(Html(item.Label)).Append("</div><div class=\"result-metric-value\">").Append(Html(item.Value)).Append("</div></div>");
				}
				html.Append("</div>");
			}
			if (group.Facts.Count > 0)
			{
				html.Append("<table").Append(contextIdWritten ? string.Empty : " id=\"messageContextTable\"").Append(" class=\"result-info-table\"><tbody>");
				contextIdWritten = true;
				foreach (ResultItem item in group.Facts)
				{
					html.Append("<tr><th>").Append(Html(item.Label)).Append("</th><td>").Append(Html(Blank(item.Value))).Append("</td></tr>");
				}
				html.Append("</tbody></table>");
			}
			if (group.Paths.Count > 0)
			{
				html.Append("<table").Append(outputIdWritten ? string.Empty : " id=\"messageOutputList\"").Append(" class=\"result-output-table\"><tbody>");
				outputIdWritten = true;
				foreach (ResultItem item in group.Paths)
				{
					html.Append("<tr><th>").Append(Html(item.Label)).Append("</th><td title=\"").Append(Html(item.Value)).Append("\">").Append(Html(Blank(item.Value))).Append("</td></tr>");
				}
				html.Append("</tbody></table>");
			}
			if (group.Notes.Count > 0)
			{
				html.Append("<ul class=\"result-notes\">");
				foreach (string note in group.Notes)
				{
					html.Append("<li>").Append(Html(note)).Append("</li>");
				}
				html.Append("</ul>");
			}
			html.Append("</div>");
		}
		html.Append("</div>");
	}

	private static void AnalyzeAutomaticResult(ParsedMessage parsed, bool isKorean)
	{
		string source = NormalizeNewlines(parsed.Intro);
		string[] lines = source.Split(new char[] { '\n' });
		List<ResultGroup> groups = new List<ResultGroup>();
		ResultGroup current = new ResultGroup(string.Empty);
		groups.Add(current);
		int itemCount = 0;
		for (int i = 0; i < lines.Length; i++)
		{
			string line = (lines[i] ?? string.Empty).Trim();
			if (string.IsNullOrWhiteSpace(line))
			{
				continue;
			}
			ResultItem item;
			if (TryParseResultItem(line, out item))
			{
				AddResultItem(current, item);
				itemCount++;
				continue;
			}
			if (LooksLikeResultGroupHeading(lines, i))
			{
				current = new ResultGroup(TrimBullet(line));
				groups.Add(current);
				continue;
			}
			current.Notes.Add(TrimBullet(line));
		}
		if (itemCount < 2)
		{
			return;
		}
		parsed.Intro = string.Empty;
		foreach (ResultGroup group in groups)
		{
			if (group.HasContent)
			{
				parsed.ResultGroups.Add(group);
			}
		}
	}

	private static bool LooksLikeResultGroupHeading(string[] lines, int index)
	{
		string line = (lines[index] ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(line) || line.Length > 80 || line.IndexOf(':') >= 0 || line.IndexOf('：') >= 0 || line.StartsWith("-", StringComparison.Ordinal))
		{
			return false;
		}
		for (int i = index + 1; i < lines.Length; i++)
		{
			string next = (lines[i] ?? string.Empty).Trim();
			if (string.IsNullOrWhiteSpace(next))
			{
				continue;
			}
			ResultItem item;
			return TryParseResultItem(next, out item);
		}
		return false;
	}

	private static bool TryParseResultItem(string rawLine, out ResultItem item)
	{
		item = null;
		string line = TrimBullet(rawLine);
		int colon = line.IndexOf(':');
		int fullColon = line.IndexOf('：');
		if (colon < 0 || (fullColon >= 0 && fullColon < colon))
		{
			colon = fullColon;
		}
		if (colon <= 0 || colon > 80 || colon >= line.Length - 1)
		{
			return false;
		}
		string label = line.Substring(0, colon).Trim();
		string value = line.Substring(colon + 1).Trim();
		if (string.IsNullOrWhiteSpace(label) || string.IsNullOrWhiteSpace(value))
		{
			return false;
		}
		item = new ResultItem(label, value, ResolveResultTone(label));
		return true;
	}

	private static void AddResultItem(ResultGroup group, ResultItem item)
	{
		if (IsResultPath(item.Label, item.Value))
		{
			group.Paths.Add(item);
		}
		else if (IsNumericResultValue(item.Value))
		{
			group.Metrics.Add(item);
		}
		else
		{
			group.Facts.Add(item);
		}
	}

	private static bool IsNumericResultValue(string value)
	{
		double number;
		return double.TryParse((value ?? string.Empty).Trim().Replace(",", string.Empty), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out number);
	}

	private static bool IsResultPath(string label, string value)
	{
		string key = ((label ?? string.Empty) + " " + (value ?? string.Empty)).ToLowerInvariant();
		return key.IndexOf("path", StringComparison.Ordinal) >= 0 || key.IndexOf("report", StringComparison.Ordinal) >= 0 || key.IndexOf("output", StringComparison.Ordinal) >= 0 || key.IndexOf("registry", StringComparison.Ordinal) >= 0 || key.IndexOf("snapshot", StringComparison.Ordinal) >= 0 || key.IndexOf("경로", StringComparison.Ordinal) >= 0 || key.IndexOf("리포트", StringComparison.Ordinal) >= 0 || key.IndexOf("보고서", StringComparison.Ordinal) >= 0 || key.IndexOf("결과 파일", StringComparison.Ordinal) >= 0 || key.IndexOf("스냅샷", StringComparison.Ordinal) >= 0 || key.IndexOf("\\\\", StringComparison.Ordinal) >= 0 || key.IndexOf(":\\", StringComparison.Ordinal) >= 0;
	}

	private static string ResolveResultTone(string label)
	{
		string key = (label ?? string.Empty).ToLowerInvariant();
		if (ContainsAny(key, "failed", "error", "blocked", "different", "missing", "실패", "오류", "차단", "다름", "누락")) return "bad";
		if (ContainsAny(key, "review", "approval", "skipped", "project only", "검토", "승인", "건너뜀", "프로젝트 전용", "조치 필요")) return "warn";
		if (ContainsAny(key, "loaded", "reloaded", "created", "latest", "no change", "success", "로드", "재로드", "생성", "기준 일치", "변경 없음", "성공")) return "good";
		if (ContainsAny(key, "ready", "available", "tracking", "준비", "가능", "추적", "갱신")) return "info";
		return "neutral";
	}

	private static bool ContainsAny(string value, params string[] needles)
	{
		foreach (string needle in needles)
		{
			if (value.IndexOf(needle, StringComparison.Ordinal) >= 0) return true;
		}
		return false;
	}

	private static string TrimBullet(string value)
	{
		string text = (value ?? string.Empty).Trim();
		while (text.StartsWith("-", StringComparison.Ordinal) || text.StartsWith("•", StringComparison.Ordinal))
		{
			text = text.Substring(1).TrimStart();
		}
		return text;
	}

	private static void AppendLine(StringBuilder builder, string line)
	{
		if (builder.Length > 0)
		{
			builder.AppendLine();
		}
		builder.Append((line ?? string.Empty).TrimEnd());
	}

	private static bool TryResolveHeading(string raw, out string key)
	{
		key = string.Empty;
		string value = (raw ?? string.Empty).Trim().TrimEnd(':', '：').Trim().ToLowerInvariant();
		switch (value)
		{
			case "실패 이유":
			case "원인":
			case "why it failed":
			case "cause":
				key = "cause";
				return true;
			case "지금 할 일":
			case "조치":
			case "다음 작업":
			case "what to do now":
			case "action":
			case "next step":
				key = "action";
				return true;
			case "관리자에게 전달할 정보":
			case "관리자 확인":
			case "관리자 조치":
			case "send this to the administrator":
			case "administrator information":
			case "admin action":
				key = "admin";
				return true;
			case "기술 정보":
			case "기술 상세":
			case "technical detail":
			case "technical information":
				key = "technical";
				return true;
			case "영향":
			case "impact":
				key = "impact";
				return true;
			case "처리 결과":
			case "결과":
			case "result":
			case "action result":
				key = "result";
				return true;
			case "확인 사항":
			case "안내":
			case "주의":
			case "things to check":
			case "notice":
			case "warning":
				key = "notice";
				return true;
			case "세부 정보":
			case "상세 정보":
			case "details":
				key = "details";
				return true;
			default:
				return false;
		}
	}

	private static string DisplayHeading(string key, bool isKorean)
	{
		switch (key)
		{
			case "cause":
				return isKorean ? "실패 이유" : "Why it failed";
			case "action":
				return isKorean ? "지금 할 일" : "What to do now";
			case "admin":
				return isKorean ? "관리자에게 전달할 정보" : "Administrator information";
			case "technical":
				return isKorean ? "기술 정보" : "Technical detail";
			case "impact":
				return isKorean ? "영향" : "Impact";
			case "result":
				return isKorean ? "처리 결과" : "Result";
			case "notice":
				return isKorean ? "확인 사항" : "Things to check";
			case "details":
				return isKorean ? "세부 정보" : "Details";
			default:
				return key ?? string.Empty;
		}
	}

	private static string ResolveTone(string key)
	{
		switch (key)
		{
			case "cause":
				return "cause";
			case "action":
				return "action";
			case "admin":
				return "admin";
			case "technical":
				return "technical";
			default:
				return "neutral";
		}
	}

	private static string ResolveKind(MessageBoxIcon icon)
	{
		if (icon == MessageBoxIcon.Hand)
		{
			return "error";
		}
		if (icon == MessageBoxIcon.Exclamation)
		{
			return "warning";
		}
		if (icon == MessageBoxIcon.Question)
		{
			return "question";
		}
		return "information";
	}

	private static string ResolveKindLabel(bool isKorean, MessageBoxIcon icon)
	{
		if (icon == MessageBoxIcon.Hand)
		{
			return isKorean ? "오류" : "Error";
		}
		if (icon == MessageBoxIcon.Exclamation)
		{
			return isKorean ? "주의" : "Attention";
		}
		if (icon == MessageBoxIcon.Question)
		{
			return isKorean ? "확인" : "Confirmation";
		}
		return isKorean ? "안내" : "Information";
	}

	private static string ResolveAccent(MessageBoxIcon icon)
	{
		if (icon == MessageBoxIcon.Hand)
		{
			return "#c94d3e";
		}
		if (icon == MessageBoxIcon.Exclamation)
		{
			return "#d39a1b";
		}
		if (icon == MessageBoxIcon.Question)
		{
			return "#2f6bff";
		}
		return "#2f6bff";
	}

	private static string ResolveTint(MessageBoxIcon icon)
	{
		if (icon == MessageBoxIcon.Hand)
		{
			return "#f3d7d2";
		}
		if (icon == MessageBoxIcon.Exclamation)
		{
			return "#f5e6bd";
		}
		if (icon == MessageBoxIcon.Question)
		{
			return "#dbe8ff";
		}
		return "#dbe8ff";
	}

	private static string ResolveGlyph(MessageBoxIcon icon)
	{
		if (icon == MessageBoxIcon.Hand)
		{
			return "X";
		}
		if (icon == MessageBoxIcon.Exclamation)
		{
			return "!";
		}
		if (icon == MessageBoxIcon.Question)
		{
			return "?";
		}
		return "i";
	}

	private static string ToIdSuffix(string key)
	{
		if (string.IsNullOrWhiteSpace(key))
		{
			return "General";
		}
		return char.ToUpperInvariant(key[0]) + key.Substring(1);
	}

	private static string ResolveSectionElementId(string key)
	{
		switch (key)
		{
			case "cause":
				return "messageSectionCause";
			case "action":
				return "messageSectionAction";
			case "admin":
				return "messageSectionAdmin";
			case "technical":
				return "messageSectionTechnical";
			default:
				return "messageSection" + ToIdSuffix(key);
		}
	}

	private static string NormalizeNewlines(string value)
	{
		return (value ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n");
	}

	private static string Blank(string value)
	{
		return string.IsNullOrWhiteSpace(value) ? "-" : value.Trim();
	}

	private static string Html(string value)
	{
		return (value ?? string.Empty).Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&#39;");
	}

	private sealed class ParsedMessage
	{
		public string Headline = string.Empty;
		public string Intro = string.Empty;
		public readonly List<MessageSection> Sections = new List<MessageSection>();
		public readonly List<ResultGroup> ResultGroups = new List<ResultGroup>();
	}

	private sealed class ResultGroup
	{
		public readonly string Title;
		public readonly List<ResultItem> Metrics = new List<ResultItem>();
		public readonly List<ResultItem> Facts = new List<ResultItem>();
		public readonly List<ResultItem> Paths = new List<ResultItem>();
		public readonly List<string> Notes = new List<string>();
		public bool HasContent { get { return Metrics.Count > 0 || Facts.Count > 0 || Paths.Count > 0 || Notes.Count > 0; } }

		public ResultGroup(string title)
		{
			Title = title ?? string.Empty;
		}
	}

	private sealed class ResultItem
	{
		public readonly string Label;
		public readonly string Value;
		public readonly string Tone;

		public ResultItem(string label, string value, string tone)
		{
			Label = label ?? string.Empty;
			Value = value ?? string.Empty;
			Tone = string.IsNullOrWhiteSpace(tone) ? "neutral" : tone;
		}
	}

	private sealed class MessageSection
	{
		public readonly string Key;
		public readonly string Tone;
		public readonly StringBuilder Builder = new StringBuilder();
		public string Body = string.Empty;

		public MessageSection(string key, string tone)
		{
			Key = key ?? string.Empty;
			Tone = tone ?? "neutral";
		}
	}

	private sealed class MetaLine
	{
		public readonly string Label;
		public readonly string Value;
		public readonly string ElementId;

		public MetaLine(string label, string value, string elementId)
		{
			Label = label ?? string.Empty;
			Value = value ?? string.Empty;
			ElementId = elementId ?? string.Empty;
		}
	}
}
