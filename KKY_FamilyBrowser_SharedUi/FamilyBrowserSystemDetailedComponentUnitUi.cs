using System;
using System.Text;

public static class FamilyBrowserSystemDetailedComponentUnitUi
{
	private const string Style =
		".system-component-head{display:table;width:100%;min-height:50px;background:#eff4ff;border-bottom:1px solid #bfcce1;table-layout:fixed;}" +
		".system-component-title,.system-component-unit-cell{display:table-cell;padding:9px 11px;vertical-align:middle;}" +
		".system-component-title{color:#24416f;font-size:14px;font-weight:900;text-align:left;}" +
		".system-component-unit-cell{width:210px;text-align:right;}" +
		".system-component-value-main{color:#17233a;font-weight:800;line-height:1.4;word-break:keep-all;overflow-wrap:anywhere;}" +
		".system-component-meta{margin-top:4px;color:#687994;font-size:11px;line-height:1.35;word-break:keep-all;overflow-wrap:anywhere;}" +
		".system-component-diff-line{display:table;width:100%;table-layout:fixed;margin:0 0 5px 0;}" +
		".system-component-diff-line:last-child{margin-bottom:0;}" +
		".system-component-diff-label,.system-component-diff-value{display:table-cell;vertical-align:top;}" +
		".system-component-diff-label{width:76px;color:#526581;font-size:11px;font-weight:900;white-space:nowrap;}" +
		".system-component-diff-value{color:#17233a;font-size:13px;font-weight:800;word-break:keep-all;overflow-wrap:anywhere;}" +
		".system-component-diff-line.current .system-component-diff-value{color:#9b3c22;}" +
		".system-component-length-value{font-variant-numeric:tabular-nums;}" +
		"@media(max-width:900px){.system-component-title,.system-component-unit-cell{display:block;width:auto;text-align:left;}.system-component-unit-cell{padding-top:0;}}";

	public static string Script(bool english)
	{
		StringBuilder script = new StringBuilder();
		script.Append("(function(){if(document.getElementById('kkyfb-system-component-unit-style'))return;var css=")
			.Append(JsString(Style))
			.Append(";var style=document.createElement('style');style.id='kkyfb-system-component-unit-style';style.type='text/css';if(style.styleSheet){style.styleSheet.cssText=css;}else{style.appendChild(document.createTextNode(css));}(document.getElementsByTagName('head')[0]||document.documentElement).appendChild(style);})();");

		script.Append("var systemComponentUnitLabel=").Append(JsString(english ? "Value unit" : "값 단위")).Append(";")
			.Append("var systemComponentStandardLabel=").Append(JsString(english ? "Standard" : "기준")).Append(";")
			.Append("var systemComponentCurrentLabel=").Append(JsString(english ? "Current" : "현재")).Append(";")
			.Append("var systemComponentUnitSequence=0;");

		script.Append(@"function systemComponentTrim(value){return trimParam(String(value==null?'':value));}"
			+ @"function systemComponentIsLength(kind){return systemComponentTrim(kind).toLowerCase()=='length';}"
			+ @"function systemComponentLegacyRow(p){return{structured:false,name:systemComponentTrim(p[2]||''),value:systemComponentTrim(p.slice(3).join(' ')||'')};}"
			+ @"function systemComponentRecord(p){return{structured:true,isDiff:false,name:systemComponentTrim(p[2]||''),valueKind:systemComponentTrim(p[3]||''),raw:systemComponentTrim(p[4]||''),display:systemComponentTrim(p[5]||''),reference:systemComponentTrim(p[6]||''),path:systemComponentTrim(p.slice(7).join(' ')||'')};}"
			+ @"function systemComponentDifferenceRecord(p){return{structured:true,isDiff:true,name:systemComponentTrim(p[2]||''),standardKind:systemComponentTrim(p[3]||''),standardRaw:systemComponentTrim(p[4]||''),standardDisplay:systemComponentTrim(p[5]||''),projectKind:systemComponentTrim(p[6]||''),projectRaw:systemComponentTrim(p[7]||''),projectDisplay:systemComponentTrim(p[8]||''),path:systemComponentTrim(p.slice(9).join(' ')||'')};}"
			+ @"function parseSystemComponentRecords(raw,configSection,differenceSection){var structuredConfig=[],structuredDiff=[],legacyConfig=[],legacyDiff=[];configSection=systemComponentTrim(configSection).toLowerCase();differenceSection=systemComponentTrim(differenceSection).toLowerCase();var lines=String(raw||'').split(/\r?\n/);for(var i=0;i<lines.length;i++){var p=(lines[i]||'').split('\t');var section=systemComponentTrim(p[1]||'').toLowerCase();if(p[0]=='@component'&&section==configSection){structuredConfig.push(systemComponentRecord(p));continue;}if(p[0]=='@component-diff'&&section==differenceSection){structuredDiff.push(systemComponentDifferenceRecord(p));continue;}if(p[0]!='@row')continue;if(section==configSection)legacyConfig.push(systemComponentLegacyRow(p));else if(section==differenceSection)legacyDiff.push(systemComponentLegacyRow(p));}return{components:structuredConfig.length?structuredConfig:legacyConfig,diffs:structuredDiff.length?structuredDiff:legacyDiff};}"
			+ @"parseSystemComponentRows=function(raw){return parseSystemComponentRecords(raw,'components','component-differences');};"
			+ @"parseCurtainPanelComponentRows=function(raw){return parseSystemComponentRecords(raw,'curtain-components','curtain-component-differences');};"
			+ @"function systemComponentValueHtml(kind,raw,display){raw=systemComponentTrim(raw);display=systemComponentTrim(display)||'-';if(systemComponentIsLength(kind)&&raw!==''&&isFinite(parseFloat(raw))){return '<span class=""system-routing-size-value system-component-length-value"" data-feet=""'+escHtml(raw)+'"">'+escHtml(formatSystemRoutingSize(raw,systemRoutingDisplayUnit))+'</span>';}return '<span>'+escHtml(display)+'</span>';}"
			+ @"function systemComponentRowHasLength(row){if(!row||!row.structured)return false;if(row.isDiff)return systemComponentIsLength(row.standardKind)||systemComponentIsLength(row.projectKind);return systemComponentIsLength(row.valueKind);}"
			+ @"function systemComponentRowsHaveLength(rows){for(var i=0;i<(rows||[]).length;i++){if(systemComponentRowHasLength(rows[i]))return true;}return false;}"
			+ @"function systemComponentUnitToolbar(){systemComponentUnitSequence++;var id='systemComponentUnit'+systemComponentUnitSequence;return '<div class=""system-routing-unit-toolbar"" data-unit-scope=""components""><label class=""system-routing-unit-label"" for=""'+id+'"">'+escHtml(systemComponentUnitLabel)+'</label><select id=""'+id+'"" class=""system-routing-unit-select system-component-unit-select"" onchange=""return changeSystemRoutingUnit(this)""><option value=""mm""'+(systemRoutingDisplayUnit=='mm'?' selected=""selected""':'')+'>mm</option><option value=""in""'+(systemRoutingDisplayUnit=='in'?' selected=""selected""':'')+'>in</option></select></div>';}"
			+ @"function systemComponentConfigValueHtml(row){if(!row.structured)return '<div class=""system-component-value-main"">'+escHtml(row.value||'-')+'</div>';var html='<div class=""system-component-value-main"">'+systemComponentValueHtml(row.valueKind,row.raw,row.display)+'</div>';if(row.reference)html+='<div class=""system-component-meta"">'+escHtml(row.reference)+'</div>';if(row.path)html+='<div class=""system-component-meta"">'+escHtml(row.path)+'</div>';return html;}"
			+ @"function systemComponentDifferenceValueHtml(row){if(!row.structured)return '<div class=""system-component-value-main"">'+escHtml(row.value||'-')+'</div>';var html='<div class=""system-component-diff-line standard""><span class=""system-component-diff-label"">'+escHtml(systemComponentStandardLabel)+'</span><span class=""system-component-diff-value"">'+systemComponentValueHtml(row.standardKind,row.standardRaw,row.standardDisplay)+'</span></div>';html+='<div class=""system-component-diff-line current""><span class=""system-component-diff-label"">'+escHtml(systemComponentCurrentLabel)+'</span><span class=""system-component-diff-value"">'+systemComponentValueHtml(row.projectKind,row.projectRaw,row.projectDisplay)+'</span></div>';if(row.path)html+='<div class=""system-component-meta"">'+escHtml(row.path)+'</div>';return html;}"
			+ @"function systemComponentRowTitle(row){if(!row)return '';if(!row.structured)return row.value||'';if(row.isDiff)return systemComponentStandardLabel+': '+(row.standardDisplay||row.standardRaw||'-')+' | '+systemComponentCurrentLabel+': '+(row.projectDisplay||row.projectRaw||'-')+(row.path?' | '+row.path:'');return (row.display||row.raw||'-')+(row.reference?' | '+row.reference:'')+(row.path?' | '+row.path:'');}"
			+ @"renderSystemComponentTable=function(title,rows,kind){if(!rows||!rows.length)return '';var hasLength=systemComponentRowsHaveLength(rows);var html='<div class=""system-component-block '+escHtml(kind||'')+'""><div class=""system-component-head""><div class=""system-component-title"">'+escHtml(title)+'</div><div class=""system-component-unit-cell"">'+(hasLength?systemComponentUnitToolbar():'')+'</div></div><div class=""system-component-table-wrap""><table class=""system-component-table '+escHtml(kind||'')+'"" data-system-component-table=""true""><colgroup><col style=""width:34%""/><col style=""width:66%""/></colgroup><tr><th>'+escHtml(systemComponentNameLabel)+'</th><th>'+escHtml(systemComponentValueLabel)+'</th></tr>';for(var i=0;i<rows.length;i++){var row=rows[i]||{};html+='<tr><td title=""'+escHtml(row.name||'')+'"">'+escHtml(row.name||'-')+'</td><td title=""'+escHtml(systemComponentRowTitle(row))+'"">'+(row.isDiff?systemComponentDifferenceValueHtml(row):systemComponentConfigValueHtml(row))+'</td></tr>';}return html+'</table></div></div>';};");

		return script.ToString();
	}

	private static string JsString(string value)
	{
		return "'" + (value ?? string.Empty)
			.Replace("\\", "\\\\")
			.Replace("'", "\\'")
			.Replace("\r", "\\r")
			.Replace("\n", "\\n")
			.Replace("</", "<\\/") + "'";
	}
}
