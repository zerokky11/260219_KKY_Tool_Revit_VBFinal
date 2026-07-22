using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

internal sealed class FileGuardHtmlConfigurationForm : Form
{
	private static string BuildLegacyWebBrowserLayoutCss()
	{
		StringBuilder css = new StringBuilder();
		css.Append("*{box-sizing:border-box;}html,body{overflow:auto!important;}button,input{font-family:'Malgun Gothic','Segoe UI',Arial,sans-serif;}");
		css.Append(".top{display:block!important;position:relative!important;padding:12px 20px 11px 34px!important;min-height:64px!important;box-sizing:border-box!important;overflow:visible!important;}.mark{position:absolute!important;left:24px!important;top:18px!important;bottom:18px!important;width:6px!important;height:auto!important;}h1{font-size:20px!important;line-height:1.25!important;margin:0 0 6px!important;white-space:normal!important;max-height:none!important;overflow:visible!important;}.hint{font-size:12px!important;line-height:1.45!important;white-space:normal!important;max-height:none!important;overflow:visible!important;}");
		css.Append(".body{display:block!important;overflow:auto!important;}.panel{margin-bottom:10px!important;}.toolbar{display:block!important;white-space:normal!important;}.toolbar>*{margin:3px 5px 6px 0!important;vertical-align:middle!important;}.toolbar .search{display:inline-block!important;width:420px!important;max-width:100%!important;min-width:260px!important;margin-left:0!important;}.toolbar .search b{display:inline-block!important;margin-right:8px!important;}input[type=text]{height:36px!important;line-height:34px!important;}");
		css.Append("button{display:inline-block!important;min-height:32px!important;height:auto!important;line-height:1.2!important;padding:7px 11px!important;text-align:center!important;white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important;max-width:220px!important;}.tablewrap{height:auto!important;min-height:260px!important;max-height:none!important;overflow:auto!important;}th,td{line-height:1.35!important;}.footer{display:table!important;width:100%!important;box-sizing:border-box!important;}.footer .count{display:table-cell!important;vertical-align:middle!important;}.footer>div:last-child{display:table-cell!important;text-align:right!important;vertical-align:middle!important;white-space:normal!important;}");
		return css.ToString();
	}

	private sealed class FileGuardHtmlRow
	{
		public int Index { get; set; }

		public bool Selected { get; set; }

		public string FilePath { get; set; }

		public string FileName { get; set; }

		public string RelativePath { get; set; }

		public string Discipline { get; set; }

		public bool Protect { get; set; }

		public bool BlockFamily { get; set; }

		public bool BlockType { get; set; }

		public bool BlockNestedOnly { get; set; }

		public bool TrackElements { get; set; }
	}

	private string _rootFolder;

	private readonly bool _isKorean;

	private readonly FamilyBrowserStandardPolicy _standardPolicy;

	private readonly List<FamilyBrowserStandardLibrarySlot> _disciplineOptions;

	private readonly string _defaultDiscipline;

	private readonly List<FileGuardHtmlRow> _rows;

	private readonly WebBrowser _browser;

	private FamilyBrowserFileGuardPolicy _resultPolicy;

	public FileGuardHtmlConfigurationForm(FamilyBrowserFileGuardPolicy currentGuard, bool isKorean)
		: this(currentGuard, null, isKorean)
	{
	}

	public FileGuardHtmlConfigurationForm(FamilyBrowserFileGuardPolicy currentGuard, FamilyBrowserStandardPolicy standardPolicy, bool isKorean)
		: this(currentGuard?.RootFolder ?? string.Empty, new List<string>(), currentGuard, standardPolicy, isKorean)
	{
	}

	public FileGuardHtmlConfigurationForm(string rootFolder, List<string> rvtFiles, FamilyBrowserFileGuardPolicy currentGuard, bool isKorean)
		: this(rootFolder, rvtFiles, currentGuard, null, isKorean)
	{
	}

	public FileGuardHtmlConfigurationForm(string rootFolder, List<string> rvtFiles, FamilyBrowserFileGuardPolicy currentGuard, FamilyBrowserStandardPolicy standardPolicy, bool isKorean)
	{
		_rootFolder = FirstNonEmpty(rootFolder, currentGuard?.RootFolder);
		_isKorean = isKorean;
		_standardPolicy = standardPolicy ?? new FamilyBrowserStandardPolicy();
		_disciplineOptions = FamilyBrowserFileGuardDisciplineService.GetSelectableSlots(_standardPolicy);
		_defaultDiscipline = FamilyBrowserFileGuardDisciplineService.ResolveAssignedDiscipline(_standardPolicy, null, allowLegacyFallback: true);
		_rows = BuildRows(_rootFolder, rvtFiles, currentGuard, _defaultDiscipline);
		Text = Tx("File-specific Guard", "파일별 권한 적용");
		AutoScaleMode = AutoScaleMode.Dpi;
		AutoScaleDimensions = new SizeF(96f, 96f);
		Font = new Font(_isKorean ? "Malgun Gothic" : "Segoe UI", 10f, FontStyle.Regular, GraphicsUnit.Point);
		StartPosition = FormStartPosition.CenterParent;
		FormBorderStyle = FormBorderStyle.Sizable;
		ShowInTaskbar = false;
		MinimizeBox = false;
		MaximizeBox = true;
		Rectangle workingArea = Screen.FromPoint(Cursor.Position).WorkingArea;
		ClientSize = new Size(Math.Max(1100, Math.Min(1480, workingArea.Width - 100)), Math.Max(720, Math.Min(940, workingArea.Height - 100)));
		MinimumSize = new Size(1040, 660);
		_browser = new WebBrowser
		{
			Dock = DockStyle.Fill,
			ScriptErrorsSuppressed = true,
			AllowNavigation = true,
			AllowWebBrowserDrop = false,
			IsWebBrowserContextMenuEnabled = false,
			WebBrowserShortcutsEnabled = false
		};
		_browser.Navigating += BrowserNavigating;
		Controls.Add(_browser);
	}

	protected override void OnShown(EventArgs e)
	{
		base.OnShown(e);
		RenderBrowser();
	}

	public FamilyBrowserFileGuardPolicy BuildPolicy()
	{
		return _resultPolicy ?? BuildPolicyFromRows();
	}

	private void BrowserNavigating(object sender, WebBrowserNavigatingEventArgs e)
	{
		if (e == null || e.Url == null)
		{
			return;
		}
		string command = string.Empty;
		if (string.Equals(e.Url.Scheme, "kkyfileguard", StringComparison.OrdinalIgnoreCase))
		{
			command = e.Url.Host ?? string.Empty;
		}
		else
		{
			string raw = e.Url.AbsoluteUri ?? string.Empty;
			if (raw.StartsWith("about:kkyfileguard://", StringComparison.OrdinalIgnoreCase))
			{
				command = raw.Substring("about:kkyfileguard://".Length).Trim('/');
			}
			else if (raw.StartsWith("about:kkyfileguard:", StringComparison.OrdinalIgnoreCase))
			{
				command = raw.Substring("about:kkyfileguard:".Length).Trim('/', ' ');
			}
		}
		if (string.IsNullOrWhiteSpace(command))
		{
			return;
		}
		command = Uri.UnescapeDataString(command.Trim('/', ' '));
		e.Cancel = true;
		if (string.Equals(command, "save", StringComparison.OrdinalIgnoreCase))
		{
			try
			{
				CaptureBrowserState();
				_resultPolicy = BuildPolicyFromRows();
				DialogResult = DialogResult.OK;
				Close();
			}
			catch (Exception ex)
			{
				FamilyBrowserModernMessageDialog.Show(this, _isKorean, ex.Message, Tx("File-specific Guard", "파일별 권한 적용"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
			}
			return;
		}
		if (string.Equals(command, "cancel", StringComparison.OrdinalIgnoreCase))
		{
			DialogResult = DialogResult.Cancel;
			Close();
			return;
		}
		CaptureBrowserState();
		if (string.Equals(command, "import-excel", StringComparison.OrdinalIgnoreCase))
		{
			try
			{
				ImportExcel();
				RenderBrowser();
			}
			catch (Exception ex)
			{
				FamilyBrowserModernMessageDialog.Show(this, _isKorean, ex.Message, Tx("Import File Guard Excel", "파일별 권한 Excel 가져오기"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
			}
			return;
		}
		if (string.Equals(command, "export-excel", StringComparison.OrdinalIgnoreCase))
		{
			try
			{
				ExportExcelTemplate();
				RenderBrowser();
			}
			catch (Exception ex)
			{
				FamilyBrowserModernMessageDialog.Show(this, _isKorean, ex.Message, Tx("Save File Guard Excel", "파일별 권한 Excel 저장"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
			}
			return;
		}
		if (string.Equals(command, "add-file", StringComparison.OrdinalIgnoreCase))
		{
			AddRvtFilesFromDialog();
			RenderBrowser();
			return;
		}
		if (string.Equals(command, "add-folder", StringComparison.OrdinalIgnoreCase))
		{
			AddRvtFilesFromFolderDialog();
			RenderBrowser();
			return;
		}
		if (string.Equals(command, "remove-selected", StringComparison.OrdinalIgnoreCase))
		{
			_rows.RemoveAll((FileGuardHtmlRow row) => row.Selected);
			ReindexRows();
			RenderBrowser();
			return;
		}
		if (string.Equals(command, "clear-all", StringComparison.OrdinalIgnoreCase))
		{
			if (FamilyBrowserModernMessageDialog.Show(this, _isKorean, Tx("Remove every RVT from the file-specific guard list?", "파일별 권한 적용 목록의 RVT를 모두 지울까요?"), Tx("File-specific Guard", "파일별 권한 적용"), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
			{
				_rows.Clear();
				RenderBrowser();
			}
		}
	}

	private void RenderBrowser()
	{
		ReindexRows();
		_browser.DocumentText = BuildHtml();
	}

	private string BuildHtml()
	{
		FamilyBrowserUiTheme theme = FamilyBrowserUiThemeService.Load();
		StringBuilder builder = new StringBuilder();
		builder.Append("<!doctype html><html><head><meta charset='utf-8'><meta http-equiv='X-UA-Compatible' content='IE=edge'><style>");
		builder.Append("html,body{margin:0;height:100%;font-family:'").Append(_isKorean ? "Malgun Gothic" : "Segoe UI").Append("',sans-serif;background:#f4f8f6;color:#20342d;font-size:13px;}");
		builder.Append(".shell{height:100%;display:flex;flex-direction:column;}");
		builder.Append(".top{background:#263731;color:#fff;padding:18px 24px 16px;display:flex;align-items:center;gap:16px;}");
		builder.Append(".mark{width:6px;align-self:stretch;background:#25b27f;border-radius:3px;}");
		builder.Append("h1{font-size:21px;line-height:1.2;margin:0 0 5px;font-weight:800;}");
		builder.Append(".hint{font-size:12px;color:#d9ebe4;line-height:1.45;}");
		builder.Append(".body{padding:16px 20px 14px;display:flex;flex-direction:column;gap:10px;min-height:0;flex:1;}");
		builder.Append(".panel{border:1px solid #cddbd5;background:#fff;border-radius:8px;padding:11px 13px;}");
		builder.Append(".toolbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;}");
		builder.Append(".toolbar .search{flex:1;min-width:330px;display:flex;align-items:center;gap:8px;margin-left:auto;}");
		builder.Append("input[type=text],select{box-sizing:border-box;border:1px solid #b8c8c1;border-radius:5px;padding:7px 9px;font-size:13px;background:#fff;color:#20342d;}input[type=text]{width:100%;padding:9px 10px;}select{min-height:34px;}");
		builder.Append("button{border:1px solid #c5d3cd;background:#fff;color:#21372f;border-radius:6px;padding:7px 11px;font-weight:700;font-size:13px;cursor:pointer;}");
		builder.Append("button.primary{background:#1f996d;border-color:#1b825d;color:#fff;}");
		builder.Append("button.danger{color:#a13d31;border-color:#d4b7b2;background:#fff8f7;}");
		builder.Append(".tablewrap{flex:1;min-height:0;border:1px solid #c7d6cf;background:#fff;overflow:auto;border-radius:8px;}");
		builder.Append("table{width:100%;min-width:1600px;border-collapse:collapse;table-layout:fixed;}");
		builder.Append("th{position:sticky;top:0;background:#eaf2ee;color:#20342d;text-align:left;font-size:12px;padding:9px 8px;border-bottom:1px solid #c5d3cd;z-index:1;}");
		builder.Append("td{padding:8px;border-bottom:1px solid #e2ebe7;vertical-align:middle;word-break:break-all;}");
		builder.Append("tr:nth-child(even) td{background:#fbfdfc;}");
		builder.Append("tr:hover td{background:#eef8f3;}");
		builder.Append(".c{width:58px;text-align:center;word-break:normal;}");
		builder.Append(".apply{width:66px;text-align:center;word-break:normal;}");
		builder.Append(".name{width:270px;font-weight:700;}");
		builder.Append(".path{width:auto;color:#4b5e57;}");
		builder.Append(".discipline{width:170px;word-break:keep-all;}.discipline select{width:100%;}.guard{width:168px;text-align:center;word-break:keep-all;}.track-guard{width:190px;text-align:center;word-break:keep-all;}.nested-guard{width:210px;text-align:center;word-break:keep-all;}");
		builder.Append(".empty{padding:42px;text-align:center;color:#60786f;font-weight:700;}");
		builder.Append(".footer{display:flex;justify-content:space-between;align-items:center;padding:8px 16px 10px;background:#fff;border-top:1px solid #dbe6e1;}");
		builder.Append(".count{color:#547066;font-size:12px;}");
		builder.Append(BuildLegacyWebBrowserLayoutCss());
		builder.Append(FamilyBrowserUiThemeService.ThemeCss());
		builder.Append("</style><script>");
		builder.Append(FamilyBrowserOverflowTitleScript.Script());
		builder.Append("var rowCount=").Append(_rows.Count.ToString(CultureInfo.InvariantCulture)).Append(";");
		builder.Append("function el(id){return document.getElementById(id);}");
		builder.Append("function isRow(tr){return tr&&tr.getAttribute('data-row')==='1';}");
		builder.Append("function visible(tr){return tr&&tr.style.display!=='none';}");
		builder.Append("function updateCount(){var trs=document.getElementsByTagName('tr');var v=0;for(var i=0;i<trs.length;i++){if(isRow(trs[i])&&visible(trs[i]))v++;}if(el('count'))el('count').innerHTML=v+' / '+rowCount;}");
		builder.Append("function filterRows(){var q=(el('search').value||'').toLowerCase();var t=q.split(/\\s+/);var trs=document.getElementsByTagName('tr');for(var i=0;i<trs.length;i++){var tr=trs[i];if(!isRow(tr))continue;var s=(tr.getAttribute('data-search')||'').toLowerCase();var ok=true;for(var j=0;j<t.length;j++){if(t[j]&&s.indexOf(t[j])<0){ok=false;break;}}tr.style.display=ok?'':'none';}updateCount();}");
		builder.Append("function setVisibleApply(value){for(var i=0;i<rowCount;i++){var tr=el('row_'+i);if(!visible(tr))continue;el('protect_'+i).checked=value;if(value){el('family_'+i).checked=true;el('type_'+i).checked=true;el('tracking_'+i).checked=true;}}}");
		builder.Append("function setVisibleSelected(value){for(var i=0;i<rowCount;i++){var tr=el('row_'+i);if(!visible(tr))continue;el('select_'+i).checked=value;}}");
		builder.Append("function setVisibleDiscipline(){var s=el('bulkDiscipline');if(!s||!s.value)return;for(var i=0;i<rowCount;i++){var tr=el('row_'+i);if(!visible(tr))continue;el('discipline_'+i).value=s.value;}}");
		builder.Append("function serialize(){var rows=[];for(var i=0;i<rowCount;i++){rows.push(i+','+(el('select_'+i).checked?'1':'0')+','+(el('protect_'+i).checked?'1':'0')+','+(el('family_'+i).checked?'1':'0')+','+(el('type_'+i).checked?'1':'0')+','+(el('nested_'+i).checked?'1':'0')+','+(el('tracking_'+i).checked?'1':'0')+','+encodeURIComponent(el('discipline_'+i).value||''));}el('payload').value=rows.join(';');}");
		builder.Append("function command(name){serialize();window.location.href='kkyfileguard://'+name;return false;}");
		builder.Append("window.onload=function(){updateCount();};");
		builder.Append("</script></head><body data-theme='").Append(H(FamilyBrowserUiThemeService.Code(theme))).Append("' class='fb-file-guard ").Append(H(FamilyBrowserUiThemeService.BodyClass(theme))).Append("'><div class='shell'>");
		builder.Append("<div class='top'><div class='mark'></div><div><h1>").Append(H(Tx("File-specific Revit Command Guard", "파일별 Revit 기본 명령 차단"))).Append("</h1><div class='hint'>");
		builder.Append(H(Tx("Review the current guarded RVT list first. Add individual RVT files or add every RVT under a selected folder, then remove only the items you do not need.", "현재 권한 적용 RVT 목록을 먼저 확인하고, RVT 파일을 하나씩 추가하거나 폴더 하위의 RVT를 한 번에 추가한 뒤 필요 없는 항목만 삭제하세요.")));
		builder.Append("</div></div></div>");
		builder.Append("<div class='body'>");
		builder.Append("<div class='panel toolbar'>");
		builder.Append("<button class='primary' onclick=\"return command('add-file');\">").Append(H(Tx("Add RVT File", "RVT 파일 추가"))).Append("</button>");
		builder.Append("<button onclick=\"return command('add-folder');\">").Append(H(Tx("Add RVTs From Folder", "폴더에서 RVT 추가"))).Append("</button>");
		builder.Append("<button onclick='setVisibleSelected(true);return false;'>").Append(H(Tx("Select Visible", "표시 항목 선택"))).Append("</button>");
		builder.Append("<button onclick='setVisibleSelected(false);return false;'>").Append(H(Tx("Clear Selection", "선택 해제"))).Append("</button>");
		builder.Append("<button class='danger' onclick=\"return command('remove-selected');\">").Append(H(Tx("Remove Selected", "선택 삭제"))).Append("</button>");
		builder.Append("<button class='danger' onclick=\"return command('clear-all');\">").Append(H(Tx("Remove All", "전체 삭제"))).Append("</button>");
		builder.Append("<button onclick=\"return command('import-excel');\">").Append(H(Tx("Import Excel", "Excel 일괄 불러오기"))).Append("</button>");
		builder.Append("<button onclick=\"return command('export-excel');\">").Append(H(Tx("Save Excel Template", "Excel 양식 저장"))).Append("</button>");
		builder.Append("<div class='search'><b>").Append(H(Tx("Search", "검색"))).Append("</b><input id='search' type='text' onkeyup='filterRows()' placeholder='").Append(H(Tx("RVT file or path", "RVT 파일명 또는 경로"))).Append("'></div>");
		builder.Append("</div>");
		builder.Append("<div class='panel toolbar'>");
		builder.Append("<button onclick='setVisibleApply(true);return false;'>").Append(H(Tx("Guard Visible", "표시 항목 적용"))).Append("</button>");
		builder.Append("<button onclick='setVisibleApply(false);return false;'>").Append(H(Tx("Do Not Guard Visible", "표시 항목 적용 해제"))).Append("</button>");
		builder.Append("<select id='bulkDiscipline'>").Append(BuildDisciplineOptionsHtml(string.Empty, includeEmpty: true)).Append("</select>");
		builder.Append("<button onclick='setVisibleDiscipline();return false;'>").Append(H(Tx("Assign Trade To Visible", "표시 항목 공종 일괄 지정"))).Append("</button>");
		builder.Append("<span class='count' title='").Append(H(Tx("Element tracking records creation, modification, and deletion on successful Save or Synchronize. Nested-only placement blocking requires a new precise standard scan.", "요소 변경 추적은 저장 또는 동기화 성공 시 생성·수정·삭제를 기록합니다. 하위 전용 단독 배치 금지는 새 표준 정밀 스캔이 필요합니다."))).Append("'>").Append(H(Tx("New and legacy registered RVTs use element tracking by default; clear the per-file checkbox only when tracking is not required.", "신규 및 기존 등록 RVT는 요소 변경 추적을 기본 사용합니다. 추적하지 않을 파일만 해당 체크를 해제하세요."))).Append("</span>");
		builder.Append("</div>");
		builder.Append("<div class='tablewrap'>");
		if (_rows.Count == 0)
		{
			builder.Append("<div class='empty'>").Append(H(Tx("No guarded RVT files yet. Add an RVT file or folder to start.", "아직 권한 적용 RVT가 없습니다. RVT 파일 또는 폴더를 추가하세요."))).Append("</div>");
		}
		else
		{
			builder.Append("<table><thead><tr><th class='c'>").Append(H(Tx("Select", "선택"))).Append("</th><th class='apply'>").Append(H(Tx("Apply", "적용"))).Append("</th><th class='name'>").Append(H(Tx("RVT file", "RVT 파일"))).Append("</th><th class='path'>").Append(H(Tx("Path", "경로"))).Append("</th><th class='discipline' title='").Append(H(Tx("The assigned trade selects the standard RVT used for automatic Model Check and nested-family review.", "지정 공종의 표준 RVT로 최초 자동 모델검사와 하위패밀리 검토를 실행합니다."))).Append("'>").Append(H(Tx("Trade", "공종"))).Append("</th><th class='track-guard' title='").Append(H(Tx("Record element creation, modification, and deletion for this RVT.", "이 RVT의 요소 생성·수정·삭제를 기록합니다."))).Append("'>").Append(H(Tx("Track element changes", "요소 생성·수정·삭제 추적"))).Append("</th><th class='guard'>").Append(H(Tx("Family load/edit", "패밀리 로드/편집"))).Append("</th><th class='guard'>").Append(H(Tx("Type changes", "타입 변경"))).Append("</th><th class='nested-guard' title='").Append(H(Tx("Block direct project placement of families used only as nested children in the assigned trade standard.", "지정 공종 표준에서 하위 구성으로만 쓰이는 패밀리를 프로젝트에 직접 배치하지 못하게 합니다."))).Append("'>").Append(H(Tx("Block nested-only standalone placement", "하위 전용 패밀리 단독 모델링 금지"))).Append("</th></tr></thead><tbody>");
			foreach (FileGuardHtmlRow row in _rows)
			{
				string search = (row.FileName + " " + row.RelativePath + " " + row.FilePath + " " + ResolveDisciplineLabel(row.Discipline)).Replace("'", " ");
				builder.Append("<tr id='row_").Append(row.Index.ToString(CultureInfo.InvariantCulture)).Append("' data-row='1' data-search='").Append(H(search)).Append("'>");
				builder.Append("<td class='c'><input type='checkbox' id='select_").Append(row.Index.ToString(CultureInfo.InvariantCulture)).Append("'").Append(row.Selected ? " checked" : string.Empty).Append("></td>");
				builder.Append("<td class='apply'><input type='checkbox' id='protect_").Append(row.Index.ToString(CultureInfo.InvariantCulture)).Append("'").Append(row.Protect ? " checked" : string.Empty).Append("></td>");
				builder.Append("<td class='name'>").Append(H(row.FileName)).Append("</td>");
				builder.Append("<td class='path'>").Append(H(row.FilePath)).Append("</td>");
				builder.Append("<td class='discipline'><select id='discipline_").Append(row.Index.ToString(CultureInfo.InvariantCulture)).Append("'>").Append(BuildDisciplineOptionsHtml(row.Discipline, includeEmpty: false)).Append("</select></td>");
				builder.Append("<td class='track-guard'><input type='checkbox' id='tracking_").Append(row.Index.ToString(CultureInfo.InvariantCulture)).Append("'").Append(row.TrackElements ? " checked" : string.Empty).Append("></td>");
				builder.Append("<td class='guard'><input type='checkbox' id='family_").Append(row.Index.ToString(CultureInfo.InvariantCulture)).Append("'").Append(row.BlockFamily ? " checked" : string.Empty).Append("></td>");
				builder.Append("<td class='guard'><input type='checkbox' id='type_").Append(row.Index.ToString(CultureInfo.InvariantCulture)).Append("'").Append(row.BlockType ? " checked" : string.Empty).Append("></td>");
				builder.Append("<td class='nested-guard'><input type='checkbox' id='nested_").Append(row.Index.ToString(CultureInfo.InvariantCulture)).Append("'").Append(row.BlockNestedOnly ? " checked" : string.Empty).Append("></td></tr>");
			}
			builder.Append("</tbody></table>");
		}
		builder.Append("</div><input type='hidden' id='payload'></div>");
		builder.Append("<div class='footer'><div id='count' class='count'>0 / 0</div><div><button onclick=\"return command('cancel');\">").Append(H(Tx("Cancel", "취소"))).Append("</button> <button class='primary' onclick=\"return command('save');\">").Append(H(Tx("Save", "저장"))).Append("</button></div></div>");
		builder.Append("</div></body></html>");
		return builder.ToString();
	}

	private void CaptureBrowserState()
	{
		HtmlDocument document = _browser.Document;
		if (document == null)
		{
			return;
		}
		HtmlElement payloadElement = document.GetElementById("payload");
		string payload = payloadElement?.GetAttribute("value") ?? string.Empty;
		foreach (string entry in payload.Split(new char[1] { ';' }, StringSplitOptions.RemoveEmptyEntries))
		{
			string[] parts = entry.Split(',');
			if (parts.Length < 8 || !int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out int index))
			{
				continue;
			}
			if (index < 0 || index >= _rows.Count)
			{
				continue;
			}
			_rows[index].Selected = string.Equals(parts[1], "1", StringComparison.Ordinal);
			_rows[index].Protect = string.Equals(parts[2], "1", StringComparison.Ordinal);
			_rows[index].BlockFamily = string.Equals(parts[3], "1", StringComparison.Ordinal);
			_rows[index].BlockType = string.Equals(parts[4], "1", StringComparison.Ordinal);
			_rows[index].BlockNestedOnly = string.Equals(parts[5], "1", StringComparison.Ordinal);
			_rows[index].TrackElements = string.Equals(parts[6], "1", StringComparison.Ordinal);
			_rows[index].Discipline = Uri.UnescapeDataString(parts[7] ?? string.Empty);
		}
	}

	private FamilyBrowserFileGuardPolicy BuildPolicyFromRows()
	{
		List<FamilyBrowserFileGuardTarget> targets = new List<FamilyBrowserFileGuardTarget>();
		string nowText = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		string rootFolder = ResolvePolicyRootFolder();
		foreach (FileGuardHtmlRow row in _rows)
		{
			if (row.Protect && !string.IsNullOrWhiteSpace(row.FilePath))
			{
				FamilyBrowserStandardLibrarySlot assignedSlot = FamilyBrowserFileGuardDisciplineService.ResolveSlot(_standardPolicy, row.Discipline);
				if (assignedSlot == null)
				{
					throw new InvalidOperationException(Tx("Select a registered trade for every guarded RVT.", "권한을 적용할 모든 RVT에 등록된 공종을 선택하세요.") + Environment.NewLine + row.FileName);
				}
				targets.Add(new FamilyBrowserFileGuardTarget
				{
					Enabled = true,
					FileName = Path.GetFileName(row.FilePath),
					CentralPath = row.FilePath,
					RelativePath = MakeRelativePath(rootFolder, row.FilePath),
					Discipline = assignedSlot.Discipline ?? string.Empty,
					BlockFamilyLoadAndEdit = row.BlockFamily,
					BlockTypeChanges = row.BlockType,
					BlockNestedOnlyStandalonePlacement = row.BlockNestedOnly,
					TrackElementChanges = row.TrackElements,
					TrackElementChangesConfigured = true,
					LastUpdatedUtc = nowText,
					LastUpdatedBy = Environment.UserName
				});
			}
		}
		return new FamilyBrowserFileGuardPolicy
		{
			Enabled = targets.Count > 0,
			RootFolder = rootFolder,
			Targets = targets,
			LastUpdatedUtc = nowText,
			LastUpdatedBy = Environment.UserName
		};
	}

	private FamilyBrowserFileGuardPolicy BuildExcelPolicyFromRows()
	{
		string nowText = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
		string rootFolder = ResolvePolicyRootFolder();
		List<FamilyBrowserFileGuardTarget> targets = new List<FamilyBrowserFileGuardTarget>();
		foreach (FileGuardHtmlRow row in _rows)
		{
			if (row == null || string.IsNullOrWhiteSpace(row.FilePath))
			{
				continue;
			}
			FamilyBrowserStandardLibrarySlot slot = FamilyBrowserFileGuardDisciplineService.ResolveSlot(_standardPolicy, row.Discipline);
			targets.Add(new FamilyBrowserFileGuardTarget
			{
				Enabled = row.Protect,
				FileName = Path.GetFileName(row.FilePath),
				CentralPath = row.FilePath,
				RelativePath = MakeRelativePath(rootFolder, row.FilePath),
				Discipline = slot == null ? ((row.Discipline ?? string.Empty).Trim()) : (slot.Discipline ?? string.Empty),
				BlockFamilyLoadAndEdit = row.BlockFamily,
				BlockTypeChanges = row.BlockType,
				BlockNestedOnlyStandalonePlacement = row.BlockNestedOnly,
				TrackElementChanges = row.TrackElements,
				TrackElementChangesConfigured = true,
				LastUpdatedUtc = nowText,
				LastUpdatedBy = Environment.UserName
			});
		}
		if (targets.Count == 0)
		{
			targets.Add(new FamilyBrowserFileGuardTarget
			{
				Enabled = false,
				FileName = "Project_Central.rvt",
				CentralPath = @"\\server\BIM\Project_Central.rvt",
				RelativePath = "Project_Central.rvt",
				Discipline = _defaultDiscipline,
				BlockFamilyLoadAndEdit = true,
				BlockTypeChanges = true,
				BlockNestedOnlyStandalonePlacement = false,
				TrackElementChanges = true,
				TrackElementChangesConfigured = true,
				LastUpdatedUtc = nowText,
				LastUpdatedBy = Environment.UserName
			});
		}
		return new FamilyBrowserFileGuardPolicy
		{
			Enabled = targets.Any(delegate(FamilyBrowserFileGuardTarget target) { return target != null && target.Enabled; }),
			RootFolder = rootFolder,
			Targets = targets,
			LastUpdatedUtc = nowText,
			LastUpdatedBy = Environment.UserName
		};
	}

	private void ImportExcel()
	{
		using (OpenFileDialog dialog = new OpenFileDialog())
		{
			dialog.Title = Tx("Import file-specific guards from Excel", "파일별 권한 Excel 일괄 불러오기");
			dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
			dialog.Multiselect = false;
			dialog.CheckFileExists = true;
			dialog.RestoreDirectory = true;
			if (dialog.ShowDialog(this) != DialogResult.OK)
			{
				return;
			}
			if (_rows.Count > 0 && FamilyBrowserModernMessageDialog.Show(this, _isKorean, Tx("Replace the current list with the RVT rows from this workbook?", "현재 목록을 Excel의 RVT 행으로 교체할까요?"), Tx("Import Excel", "Excel 일괄 불러오기"), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
			{
				return;
			}
			FamilyBrowserFileGuardExcelImportResult result = FamilyBrowserFileGuardExcelService.Import(dialog.FileName, _standardPolicy, Environment.UserName, _isKorean);
			if (result.ImportedRowCount == 0)
			{
				throw new InvalidDataException(Tx("No valid RVT rows were found in the workbook.", "Excel에서 유효한 RVT 행을 찾지 못했습니다.") + Environment.NewLine + string.Join(Environment.NewLine, result.Warnings.Take(8)));
			}
			if (!string.IsNullOrWhiteSpace(result.Policy.RootFolder))
			{
				_rootFolder = result.Policy.RootFolder;
			}
			_rows.Clear();
			_rows.AddRange(BuildRows(_rootFolder, new List<string>(), result.Policy, _defaultDiscipline));
			string message = Tx("Excel import completed.", "Excel 일괄 불러오기를 완료했습니다.") + Environment.NewLine + Tx("Rows: ", "행: ") + result.ImportedRowCount.ToString(CultureInfo.InvariantCulture);
			if (result.Warnings.Count > 0)
			{
				message += Environment.NewLine + Environment.NewLine + Tx("Review:", "확인 필요:") + Environment.NewLine + string.Join(Environment.NewLine, result.Warnings.Take(8));
			}
			FamilyBrowserModernMessageDialog.Show(this, _isKorean, message, Tx("Import Excel", "Excel 일괄 불러오기"), MessageBoxButtons.OK, result.Warnings.Count == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
		}
	}

	private void ExportExcelTemplate()
	{
		using (SaveFileDialog dialog = new SaveFileDialog())
		{
			dialog.Title = Tx("Save file guard Excel template", "파일별 권한 Excel 양식 저장");
			dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
			dialog.DefaultExt = "xlsx";
			dialog.AddExtension = true;
			dialog.OverwritePrompt = true;
			dialog.RestoreDirectory = true;
			dialog.FileName = "FamilyBrowser_FileGuards_Template_" + DateTime.Now.ToString("yyyyMMdd-HHmm", CultureInfo.InvariantCulture) + ".xlsx";
			if (dialog.ShowDialog(this) != DialogResult.OK)
			{
				return;
			}
			FamilyBrowserPermissionExcelExportResult result = FamilyBrowserPermissionExcelExportService.ExportFileGuardPolicy(BuildExcelPolicyFromRows(), dialog.FileName, _isKorean);
			FamilyBrowserModernMessageDialog.Show(this, _isKorean, Tx("Excel template saved.", "Excel 양식을 저장했습니다.") + Environment.NewLine + result.OutputPath, Tx("Save Excel Template", "Excel 양식 저장"), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
		}
	}

	private string BuildDisciplineOptionsHtml(string selectedDiscipline, bool includeEmpty)
	{
		StringBuilder builder = new StringBuilder();
		if (includeEmpty)
		{
			builder.Append("<option value=''>").Append(H(Tx("Select trade", "공종 선택"))).Append("</option>");
		}
		string selectedKey = FamilyBrowserPolicyKey.Normalize(selectedDiscipline);
		foreach (FamilyBrowserStandardLibrarySlot slot in _disciplineOptions)
		{
			if (slot == null)
			{
				continue;
			}
			string value = slot.Discipline ?? string.Empty;
			bool selected = string.Equals(FamilyBrowserPolicyKey.Normalize(value), selectedKey, StringComparison.OrdinalIgnoreCase);
			builder.Append("<option value='").Append(H(value)).Append("'").Append(selected ? " selected" : string.Empty).Append(">").Append(H(FamilyBrowserStandardPolicyStore.ResolveSlotDisplayName(slot, _isKorean))).Append("</option>");
		}
		return builder.ToString();
	}

	private string ResolveDisciplineLabel(string discipline)
	{
		FamilyBrowserStandardLibrarySlot slot = FamilyBrowserFileGuardDisciplineService.ResolveSlot(_standardPolicy, discipline);
		return slot == null ? (discipline ?? string.Empty) : FamilyBrowserStandardPolicyStore.ResolveSlotDisplayName(slot, _isKorean);
	}

	private void AddRvtFilesFromDialog()
	{
		using OpenFileDialog dialog = new OpenFileDialog();
		dialog.Title = Tx("Add RVT files to the guard list", "권한 적용 목록에 추가할 RVT 파일을 선택하세요");
		dialog.Filter = "Revit Project (*.rvt)|*.rvt|All files (*.*)|*.*";
		dialog.Multiselect = true;
		dialog.CheckFileExists = true;
		dialog.CheckPathExists = true;
		dialog.RestoreDirectory = true;
		string initialFolder = ResolveInitialFolder();
		if (!string.IsNullOrWhiteSpace(initialFolder) && Directory.Exists(initialFolder))
		{
			dialog.InitialDirectory = initialFolder;
		}
		if (dialog.ShowDialog(this) == DialogResult.OK)
		{
			AddRvtFiles(dialog.FileNames ?? Array.Empty<string>(), null);
		}
	}

	private void AddRvtFilesFromFolderDialog()
	{
		string folder = BrowseForFolderWithExplorer(Tx("Select a folder. Every RVT under it will be added to the list.", "폴더를 선택하세요. 하위 RVT 파일을 목록에 추가합니다."), ResolveInitialFolder());
		if (string.IsNullOrWhiteSpace(folder))
		{
			return;
		}
		if (string.IsNullOrWhiteSpace(_rootFolder))
		{
			_rootFolder = folder;
		}
		Cursor previousCursor = Cursor.Current;
		try
		{
			Cursor.Current = Cursors.WaitCursor;
			AddRvtFiles(EnumerateRvtFiles(folder), folder);
		}
		finally
		{
			Cursor.Current = previousCursor;
		}
	}

	private void AddRvtFiles(IEnumerable<string> filePaths, string relativeRoot)
	{
		if (filePaths == null)
		{
			return;
		}
		foreach (string filePath in filePaths)
		{
			if (string.IsNullOrWhiteSpace(filePath) || !string.Equals(Path.GetExtension(filePath), ".rvt", StringComparison.OrdinalIgnoreCase))
			{
				continue;
			}
			AddOrUpdateRow(filePath, relativeRoot);
		}
		_rows.Sort((FileGuardHtmlRow left, FileGuardHtmlRow right) => string.Compare(left?.FilePath ?? string.Empty, right?.FilePath ?? string.Empty, StringComparison.OrdinalIgnoreCase));
		ReindexRows();
	}

	private void AddOrUpdateRow(string filePath, string relativeRoot)
	{
		string fullPath = NormalizeFullPathForDisplay(filePath);
		if (string.IsNullOrWhiteSpace(fullPath))
		{
			return;
		}
		string lookup = NormalizePathForLookup(fullPath);
		FileGuardHtmlRow existing = _rows.FirstOrDefault((FileGuardHtmlRow row) => string.Equals(NormalizePathForLookup(row.FilePath), lookup, StringComparison.OrdinalIgnoreCase));
		if (existing != null)
		{
			existing.Protect = true;
			existing.BlockFamily = true;
			existing.BlockType = true;
			existing.TrackElements = true;
			existing.Selected = false;
			return;
		}
		_rows.Add(new FileGuardHtmlRow
		{
			Index = _rows.Count,
			Selected = false,
			FilePath = fullPath,
			FileName = Path.GetFileName(fullPath),
			RelativePath = MakeRelativePath(FirstNonEmpty(relativeRoot, _rootFolder), fullPath),
			Discipline = _defaultDiscipline,
			Protect = true,
			BlockFamily = true,
			BlockType = true,
			BlockNestedOnly = false,
			TrackElements = true
		});
	}

	private string ResolvePolicyRootFolder()
	{
		if (!string.IsNullOrWhiteSpace(_rootFolder) && Directory.Exists(_rootFolder))
		{
			return _rootFolder;
		}
		List<string> folders = _rows.Select((FileGuardHtmlRow row) => SafeDirectoryName(row.FilePath)).Where((string path) => !string.IsNullOrWhiteSpace(path)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
		if (folders.Count == 1)
		{
			return folders[0];
		}
		return _rootFolder ?? string.Empty;
	}

	private string ResolveInitialFolder()
	{
		if (!string.IsNullOrWhiteSpace(_rootFolder) && Directory.Exists(_rootFolder))
		{
			return _rootFolder;
		}
		foreach (FileGuardHtmlRow row in _rows)
		{
			string folder = SafeDirectoryName(row.FilePath);
			if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
			{
				return folder;
			}
		}
		return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
	}

	private void ReindexRows()
	{
		for (int i = 0; i < _rows.Count; i++)
		{
			_rows[i].Index = i;
		}
	}

	private static List<FileGuardHtmlRow> BuildRows(string rootFolder, List<string> rvtFiles, FamilyBrowserFileGuardPolicy currentGuard, string defaultDiscipline)
	{
		List<FileGuardHtmlRow> rows = new List<FileGuardHtmlRow>();
		Dictionary<string, FamilyBrowserFileGuardTarget> targetByPath = BuildTargetLookup(currentGuard, rootFolder);
		HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		foreach (string filePath in rvtFiles ?? new List<string>())
		{
			AddRowFromFile(rows, seen, rootFolder, filePath, ResolveCurrentTarget(targetByPath, filePath, rootFolder), defaultDiscipline);
		}
		if (currentGuard != null && currentGuard.Targets != null)
		{
			foreach (FamilyBrowserFileGuardTarget target in currentGuard.Targets)
			{
				string filePath = ResolveTargetFilePath(rootFolder, target);
				FamilyBrowserFileGuardTarget effectiveTarget = ResolveCurrentTarget(targetByPath, filePath, rootFolder) ?? target;
				AddRowFromFile(rows, seen, rootFolder, filePath, effectiveTarget, defaultDiscipline);
			}
		}
		rows.Sort((FileGuardHtmlRow left, FileGuardHtmlRow right) => string.Compare(left?.FilePath ?? string.Empty, right?.FilePath ?? string.Empty, StringComparison.OrdinalIgnoreCase));
		for (int i = 0; i < rows.Count; i++)
		{
			rows[i].Index = i;
		}
		return rows;
	}

	private static void AddRowFromFile(List<FileGuardHtmlRow> rows, HashSet<string> seen, string rootFolder, string filePath, FamilyBrowserFileGuardTarget current, string defaultDiscipline)
	{
		if (rows == null || seen == null || string.IsNullOrWhiteSpace(filePath))
		{
			return;
		}
		string normalizedPath = NormalizeFullPathForDisplay(filePath);
		string key = FamilyBrowserFileGuardPathMatcher.BuildStablePolicyPathKey(normalizedPath);
		if (string.IsNullOrWhiteSpace(key))
		{
			key = NormalizePathForLookup(normalizedPath);
		}
		if (string.IsNullOrWhiteSpace(key) || !seen.Add(key))
		{
			return;
		}
		rows.Add(new FileGuardHtmlRow
		{
			Index = rows.Count,
			Selected = false,
			FilePath = normalizedPath,
			FileName = Path.GetFileName(normalizedPath),
			RelativePath = MakeRelativePath(rootFolder, normalizedPath),
			Discipline = current == null ? (defaultDiscipline ?? string.Empty) : (current.Discipline ?? string.Empty),
			Protect = current?.Enabled ?? false,
			BlockFamily = current?.BlockFamilyLoadAndEdit ?? true,
			BlockType = current?.BlockTypeChanges ?? true,
			BlockNestedOnly = current?.BlockNestedOnlyStandalonePlacement ?? false,
			TrackElements = current?.TrackElementChanges ?? true
		});
	}

	private static FamilyBrowserFileGuardTarget ResolveCurrentTarget(Dictionary<string, FamilyBrowserFileGuardTarget> targetByPath, string filePath, string rootFolder)
	{
		if (targetByPath == null || string.IsNullOrWhiteSpace(filePath))
		{
			return null;
		}
		FamilyBrowserFileGuardTarget current = null;
		string stableKey = FamilyBrowserFileGuardPathMatcher.BuildStablePolicyPathKey(filePath);
		if (!string.IsNullOrWhiteSpace(stableKey))
		{
			targetByPath.TryGetValue(stableKey, out current);
		}
		if (current == null && !targetByPath.TryGetValue(NormalizePathForLookup(filePath), out current))
		{
			targetByPath.TryGetValue(NormalizePathForLookup(MakeRelativePath(rootFolder, filePath)), out current);
		}
		bool rooted = false;
		try
		{
			rooted = Path.IsPathRooted(filePath);
		}
		catch
		{
		}
		if (current == null && !rooted)
		{
			targetByPath.TryGetValue(NormalizeNameForLookup(Path.GetFileName(filePath)), out current);
		}
		return current;
	}

	private static string ResolveTargetFilePath(string rootFolder, FamilyBrowserFileGuardTarget target)
	{
		if (target == null)
		{
			return string.Empty;
		}
		if (!string.IsNullOrWhiteSpace(target.CentralPath))
		{
			return target.CentralPath;
		}
		if (!string.IsNullOrWhiteSpace(rootFolder) && !string.IsNullOrWhiteSpace(target.RelativePath))
		{
			try
			{
				return Path.Combine(rootFolder, target.RelativePath);
			}
			catch
			{
			}
		}
		return FirstNonEmpty(target.FileName, target.RelativePath);
	}

	private static Dictionary<string, FamilyBrowserFileGuardTarget> BuildTargetLookup(FamilyBrowserFileGuardPolicy currentGuard, string rootFolder)
	{
		Dictionary<string, FamilyBrowserFileGuardTarget> result = new Dictionary<string, FamilyBrowserFileGuardTarget>(StringComparer.OrdinalIgnoreCase);
		if (currentGuard == null || currentGuard.Targets == null)
		{
			return result;
		}
		foreach (FamilyBrowserFileGuardTarget target in currentGuard.Targets)
		{
			if (target == null)
			{
				continue;
			}
			AddTargetLookup(result, NormalizePathForLookup(target.CentralPath), target);
			AddTargetLookup(result, NormalizePathForLookup(target.RelativePath), target);
			AddTargetLookup(result, NormalizeNameForLookup(target.FileName), target);
			string resolvedPath = ResolveTargetFilePath(FirstNonEmpty(rootFolder, currentGuard.RootFolder), target);
			AddTargetLookup(result, FamilyBrowserFileGuardPathMatcher.BuildStablePolicyPathKey(resolvedPath), target);
		}
		return result;
	}

	private static void AddTargetLookup(Dictionary<string, FamilyBrowserFileGuardTarget> result, string key, FamilyBrowserFileGuardTarget target)
	{
		if (result == null || string.IsNullOrWhiteSpace(key) || target == null)
		{
			return;
		}
		FamilyBrowserFileGuardTarget existing;
		if (result.TryGetValue(key, out existing))
		{
			result[key] = FamilyBrowserFileGuardPathMatcher.MergeConservativeTargets(new FamilyBrowserFileGuardTarget[] { existing, target });
		}
		else
		{
			result.Add(key, target);
		}
	}

	private string BrowseForFolderWithExplorer(string title, string initialFolder)
	{
		using (OpenFileDialog dialog = new OpenFileDialog())
		{
			dialog.Title = title ?? string.Empty;
			dialog.CheckFileExists = false;
			dialog.CheckPathExists = true;
			dialog.ValidateNames = false;
			dialog.AddExtension = false;
			dialog.DereferenceLinks = true;
			dialog.Multiselect = false;
			dialog.RestoreDirectory = true;
			dialog.FileName = Tx("Select this folder", "이 폴더 선택");
			if (!string.IsNullOrWhiteSpace(initialFolder) && Directory.Exists(initialFolder))
			{
				dialog.InitialDirectory = initialFolder;
			}
			if (!TryEnableFolderPickMode(dialog))
			{
				return BrowseForFolderFallback(title, initialFolder);
			}
			if (dialog.ShowDialog(this) != DialogResult.OK)
			{
				return string.Empty;
			}
			string selected = (dialog.FileName ?? string.Empty).Trim();
			if (Directory.Exists(selected))
			{
				return selected;
			}
			string folder = SafeDirectoryName(selected);
			if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
			{
				return folder;
			}
		}
		return string.Empty;
	}

	private string BrowseForFolderFallback(string title, string initialFolder)
	{
		using FolderBrowserDialog folderDialog = new FolderBrowserDialog();
		folderDialog.Description = title ?? string.Empty;
		folderDialog.ShowNewFolderButton = false;
		if (!string.IsNullOrWhiteSpace(initialFolder) && Directory.Exists(initialFolder))
		{
			folderDialog.SelectedPath = initialFolder;
		}
		if (folderDialog.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(folderDialog.SelectedPath))
		{
			return string.Empty;
		}
		return folderDialog.SelectedPath;
	}

	private static bool TryEnableFolderPickMode(OpenFileDialog dialog)
	{
		if (dialog == null)
		{
			return false;
		}
		try
		{
			FieldInfo field = typeof(FileDialog).GetField("options", BindingFlags.Instance | BindingFlags.NonPublic);
			if (field == null)
			{
				return false;
			}
			int options = Convert.ToInt32(field.GetValue(dialog), CultureInfo.InvariantCulture);
			field.SetValue(dialog, options | 0x20);
			return true;
		}
		catch
		{
			return false;
		}
	}

	private static List<string> EnumerateRvtFiles(string rootFolder)
	{
		List<string> result = new List<string>();
		if (string.IsNullOrWhiteSpace(rootFolder) || !Directory.Exists(rootFolder))
		{
			return result;
		}
		Stack<string> pending = new Stack<string>();
		pending.Push(rootFolder);
		while (pending.Count > 0)
		{
			string folder = pending.Pop();
			try
			{
				foreach (string file in Directory.GetFiles(folder, "*.rvt"))
				{
					result.Add(file);
				}
			}
			catch
			{
			}
			try
			{
				foreach (string child in Directory.GetDirectories(folder))
				{
					pending.Push(child);
				}
			}
			catch
			{
			}
		}
		return result.OrderBy((string path) => path, StringComparer.OrdinalIgnoreCase).ToList();
	}

	private static string MakeRelativePath(string rootFolder, string filePath)
	{
		if (string.IsNullOrWhiteSpace(rootFolder) || string.IsNullOrWhiteSpace(filePath))
		{
			return filePath ?? string.Empty;
		}
		try
		{
			string root = Path.GetFullPath(rootFolder).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
			string full = Path.GetFullPath(filePath);
			Uri uri = new Uri(root);
			Uri fileUri = new Uri(full);
			return Uri.UnescapeDataString(uri.MakeRelativeUri(fileUri).ToString()).Replace('/', Path.DirectorySeparatorChar);
		}
		catch
		{
			return Path.GetFileName(filePath);
		}
	}

	private static string NormalizeFullPathForDisplay(string value)
	{
		string text = (value ?? string.Empty).Trim();
		if (text.Length == 0)
		{
			return string.Empty;
		}
		try
		{
			if (Path.IsPathRooted(text))
			{
				return Path.GetFullPath(text);
			}
		}
		catch
		{
		}
		return text;
	}

	private static string NormalizePathForLookup(string value)
	{
		return FamilyBrowserFileGuardPathMatcher.BuildStablePolicyPathKey(value);
	}

	private static string NormalizeNameForLookup(string value)
	{
		string text = (value ?? string.Empty).Trim();
		if (text.Length == 0)
		{
			return string.Empty;
		}
		try
		{
			text = Path.GetFileNameWithoutExtension(text);
		}
		catch
		{
		}
		return text.ToLowerInvariant();
	}

	private static string SafeDirectoryName(string filePath)
	{
		try
		{
			return Path.GetDirectoryName(filePath ?? string.Empty) ?? string.Empty;
		}
		catch
		{
			return string.Empty;
		}
	}

	private static string FirstNonEmpty(params string[] values)
	{
		foreach (string value in values ?? Array.Empty<string>())
		{
			if (!string.IsNullOrWhiteSpace(value))
			{
				return value.Trim();
			}
		}
		return string.Empty;
	}

	private static string H(string value)
	{
		return System.Net.WebUtility.HtmlEncode(value ?? string.Empty);
	}

	private string Tx(string en, string ko)
	{
		return _isKorean ? ko : en;
	}
}
