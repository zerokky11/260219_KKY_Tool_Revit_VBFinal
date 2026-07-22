using System;
using System.Text;

public static class FamilyBrowserSystemRoutingUnitUi
{
	private const string Style =
		".system-routing-unit-toolbar{display:inline-block;margin:0;text-align:right;white-space:nowrap;vertical-align:middle;}" +
		".system-routing-unit-label{display:inline-block;margin-right:7px;color:#4b5d79;font-size:12px;font-weight:800;vertical-align:middle;}" +
		".system-routing-unit-select{display:inline-block;width:82px;height:32px;padding:0 9px;border:1px solid #9fb5d6;border-radius:6px;background:#fff;color:#17233a;font-size:13px;font-weight:800;vertical-align:middle;}" +
		".system-routing-preference-toolbar{margin:0 0 8px;text-align:right;}" +
		".system-routing-preference-table{min-width:1040px!important;}" +
		".system-routing-range{width:360px!important;}" +
		".system-routing-criteria-wrap{width:100%;max-width:100%;overflow-x:auto;overflow-y:hidden;box-sizing:border-box;}" +
		".system-routing-criteria-table{width:100%;min-width:300px;border-collapse:collapse;table-layout:fixed;background:#fff;}" +
		".system-routing-preference-table .system-routing-criteria-table th,.system-routing-preference-table .system-routing-criteria-table td{padding:5px 7px!important;border:1px solid #dbe4f2!important;background:#fff!important;color:#273b5d!important;font-size:12px!important;line-height:1.35!important;text-align:left!important;vertical-align:middle!important;white-space:nowrap!important;word-break:normal!important;overflow-wrap:normal!important;}" +
		".system-routing-preference-table .system-routing-criteria-table th{background:#f3f6fb!important;color:#4b5d79!important;font-weight:800!important;}" +
		".system-routing-preference-table .system-routing-criteria-table th:first-child,.system-routing-preference-table .system-routing-criteria-table td:first-child{width:28%;text-align:center!important;}" +
		".system-routing-size-value,.system-layer-thickness-value{font-variant-numeric:tabular-nums;font-weight:700;color:#17233a;white-space:nowrap;}" +
		".system-routing-all-sizes{display:inline-block;padding:4px 7px;border:1px solid #dbe4f2;border-radius:4px;background:#f8faff;color:#526581;font-size:12px;font-weight:700;}" +
		".system-routing-criteria-raw{display:block;white-space:normal;word-break:break-word;color:#526581;font-size:12px;}" +
		".system-routing-layer-block{margin-top:14px!important;border:1px solid #bfcce1!important;border-radius:7px!important;background:#fff!important;overflow:hidden!important;}" +
		".system-routing-layer-head{display:table;width:100%;min-height:50px;background:#eff4ff;border-bottom:1px solid #bfcce1;table-layout:fixed;}" +
		".system-routing-layer-title-cell,.system-routing-layer-unit-cell{display:table-cell;padding:9px 11px;vertical-align:middle;}" +
		".system-routing-layer-title-cell{color:#24416f;font-size:14px;font-weight:900;text-align:left;}" +
		".system-routing-layer-unit-cell{width:210px;text-align:right;}" +
		".system-layer-direction{padding:7px 12px;background:#f8faff;color:#526581;font-size:12px;font-weight:800;letter-spacing:0;border-bottom:1px solid #dbe4f2;}" +
		".system-layer-direction.interior{border-top:1px solid #dbe4f2;border-bottom:0;text-align:right;}" +
		".system-layer-table-wrap{width:100%;overflow-x:auto;overflow-y:hidden;background:#fff;}" +
		".system-layer-composition-table{width:100%;min-width:700px;border-collapse:collapse;table-layout:fixed;background:#fff;}" +
		".system-layer-composition-table col.layer-index{width:66px;}.system-layer-composition-table col.layer-function{width:28%;}.system-layer-composition-table col.layer-material{width:auto;}.system-layer-composition-table col.layer-thickness{width:150px;}" +
		".system-layer-composition-table th{padding:8px 10px;border-bottom:1px solid #bfcce1;border-right:1px solid #dbe4f2;background:#eaf1fc;color:#273b5d;font-size:12px;font-weight:900;text-align:left;white-space:nowrap;}" +
		".system-layer-composition-table th:last-child{border-right:0;}.system-layer-composition-table th.layer-index{text-align:center;}" +
		".system-layer-composition-table td{padding:9px 10px;border-top:1px solid #e4eaf4;border-right:1px solid #edf1f7;color:#17233a;font-size:13px;line-height:1.35;vertical-align:middle;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}" +
		".system-layer-composition-table td:last-child{border-right:0;}.system-layer-composition-table td.layer-index{text-align:center;color:#4b5d79;font-weight:800;}" +
		".system-layer-composition-table tr.layer-core td{background:#f8fbff;}.system-layer-composition-table td.layer-function{font-weight:800;}" +
		".system-layer-core-boundary td{padding:5px 10px!important;background:#dfe8f7!important;border-top:1px solid #9fb5d6!important;border-bottom:1px solid #9fb5d6!important;color:#38577f!important;font-size:11px!important;font-weight:900!important;text-align:center!important;letter-spacing:0!important;}" +
		".system-layer-badge{display:inline-block;margin-left:6px;padding:2px 5px;border:1px solid #b8c9e3;border-radius:4px;background:#f3f7fd;color:#36557e;font-size:10px;font-weight:900;vertical-align:1px;}" +
		".system-layer-badge.variable{border-color:#e1b85d;background:#fff8e7;color:#8a5a00;}" +
		".system-layer-legacy-note{display:block;margin-top:3px;color:#718098;font-size:11px;font-weight:400;white-space:normal;}" +
		".detached-content .system-routing-preference-table{min-width:1080px!important;}" +
		".detached-content .system-layer-composition-table{min-width:760px;}" +
		"@media(max-width:900px){.system-routing-preference-table{min-width:940px!important;}.system-routing-range{width:330px!important;}.system-routing-layer-title-cell,.system-routing-layer-unit-cell{display:block;width:auto;text-align:left;}.system-routing-layer-unit-cell{padding-top:0;}.system-layer-composition-table{min-width:650px;}}";

	public static string Script(bool english, string initialUnit = null)
	{
		string unit = string.Equals((initialUnit ?? string.Empty).Trim(), "in", StringComparison.OrdinalIgnoreCase) ? "in" : "mm";
		StringBuilder script = new StringBuilder();
		script.Append("(function(){if(document.getElementById('kkyfb-system-routing-unit-style'))return;var css=")
			.Append(JsString(Style))
			.Append(";var style=document.createElement('style');style.id='kkyfb-system-routing-unit-style';style.type='text/css';if(style.styleSheet){style.styleSheet.cssText=css;}else{style.appendChild(document.createTextNode(css));}(document.getElementsByTagName('head')[0]||document.documentElement).appendChild(style);})();");

		script.Append("var systemRoutingUnitLabel=").Append(JsString(english ? "Size unit" : "사이즈 단위")).Append(";")
			.Append("var systemLayerUnitLabel=").Append(JsString(english ? "Thickness unit" : "두께 단위")).Append(";")
			.Append("var systemRoutingCriterionLabel=").Append(JsString(english ? "Criterion" : "조건")).Append(";")
			.Append("var systemRoutingMinimumLabel=").Append(JsString(english ? "Minimum" : "최소")).Append(";")
			.Append("var systemRoutingMaximumLabel=").Append(JsString(english ? "Maximum" : "최대")).Append(";")
			.Append("var systemRoutingNoLimitLabel=").Append(JsString(english ? "No limit" : "제한 없음")).Append(";")
			.Append("var systemRoutingSizeLabel=").Append(JsString(english ? "Size" : "사이즈")).Append(";")
			.Append("var systemLayerIndexLabel=").Append(JsString(english ? "Layer" : "레이어")).Append(";")
			.Append("var systemLayerFunctionLabel=").Append(JsString(english ? "Function" : "기능")).Append(";")
			.Append("var systemLayerMaterialLabel=").Append(JsString(english ? "Material" : "재료")).Append(";")
			.Append("var systemLayerThicknessLabel=").Append(JsString(english ? "Thickness" : "두께")).Append(";")
			.Append("var systemLayerExteriorLabel=").Append(JsString(english ? "Exterior side" : "외부측")).Append(";")
			.Append("var systemLayerInteriorLabel=").Append(JsString(english ? "Interior side" : "내부측")).Append(";")
			.Append("var systemLayerCoreBoundaryLabel=").Append(JsString(english ? "Core Boundary" : "코어 경계")).Append(";")
			.Append("var systemLayerStructuralLabel=").Append(JsString(english ? "Structural material" : "구조 재료")).Append(";")
			.Append("var systemLayerVariableLabel=").Append(JsString(english ? "Variable" : "가변")).Append(";")
			.Append("var systemLayerLegacyLabel=").Append(JsString(english ? "Legacy snapshot" : "이전 스냅샷")).Append(";")
			.Append("var systemRoutingDisplayUnit=").Append(JsString(unit)).Append(";")
			.Append("var systemRoutingUnitToolbarSequence=0;");

		script.Append(@"function parseSystemRoutingCriteria(raw){var rows=[];var parts=String(raw||'').split(/\s*;\s*/);for(var i=0;i<parts.length;i++){var piece=trimParam(parts[i]||'');if(!piece)continue;var indexMatch=piece.match(/\bsize\s+(\d+)/i);var minMatch=piece.match(/min\s*=\s*([^\s;]+)/i);var maxMatch=piece.match(/max\s*=\s*([^\s;]+)/i);if(minMatch||maxMatch){rows.push({index:indexMatch?indexMatch[1]:String(rows.length+1),min:minMatch?minMatch[1]:'',max:maxMatch?maxMatch[1]:'',raw:piece});}}return rows;}"
			+ @"function systemRoutingTrimNumber(value,decimals){var scale=Math.pow(10,decimals);var rounded=Math.round(value*scale)/scale;if(Math.abs(rounded)<(0.5/scale))rounded=0;var text=rounded.toFixed(decimals);text=text.replace(/(\.\d*?)0+$/,'$1').replace(/\.$/,'');return text;}"
			+ @"function formatSystemRoutingSize(raw,unit){var text=trimParam(String(raw==null?'':raw));if(!text)return '-';var value=parseFloat(text);if(!isFinite(value)||Math.abs(value)>=1e20)return systemRoutingNoLimitLabel;unit=(unit=='in')?'in':'mm';var converted=value*(unit=='in'?12:304.8);return systemRoutingTrimNumber(converted,unit=='in'?3:2)+' '+unit;}"
			+ @"function systemRoutingSizeValueHtml(raw){return '<span class=""system-routing-size-value"" data-feet=""'+escHtml(String(raw==null?'':raw))+'"">'+escHtml(formatSystemRoutingSize(raw,systemRoutingDisplayUnit))+'</span>'; }"
			+ @"function renderSystemRoutingCriteria(raw){var rows=parseSystemRoutingCriteria(raw);if(!rows.length)return '<span class=""system-routing-criteria-raw"">'+escHtml(raw||systemRoutingAllSizesLabel)+'</span>';var html='<div class=""system-routing-criteria-wrap""><table class=""system-routing-criteria-table"" data-system-routing-criteria-table=""true""><tr><th>'+escHtml(systemRoutingCriterionLabel)+'</th><th>'+escHtml(systemRoutingMinimumLabel)+'</th><th>'+escHtml(systemRoutingMaximumLabel)+'</th></tr>';for(var i=0;i<rows.length;i++){html+='<tr><td>'+escHtml(systemRoutingSizeLabel+' '+rows[i].index)+'</td><td>'+systemRoutingSizeValueHtml(rows[i].min)+'</td><td>'+systemRoutingSizeValueHtml(rows[i].max)+'</td></tr>';}return html+'</table></div>'; }"
			+ @"function systemRoutingUnitToolbar(scope){systemRoutingUnitToolbarSequence++;var id='systemMeasurementUnit'+systemRoutingUnitToolbarSequence;var label=(scope=='layers')?systemLayerUnitLabel:systemRoutingUnitLabel;return '<div class=""system-routing-unit-toolbar"" data-unit-scope=""'+escHtml(scope||'routing')+'""><label class=""system-routing-unit-label"" for=""'+id+'"">'+escHtml(label)+'</label><select id=""'+id+'"" class=""system-routing-unit-select"" onchange=""return changeSystemRoutingUnit(this)""><option value=""mm""'+(systemRoutingDisplayUnit=='mm'?' selected=""selected""':'')+'>mm</option><option value=""in""'+(systemRoutingDisplayUnit=='in'?' selected=""selected""':'')+'>in</option></select></div>'; }"
			+ @"function persistSystemRoutingUnit(unit){try{window.location='kkyfb:measurement-unit/'+encodeURIComponent(unit);}catch(e){try{window.location.href='kkyfb:measurement-unit/'+encodeURIComponent(unit);}catch(e2){}}}"
			+ @"function applySystemRoutingUnit(unit,persist){systemRoutingDisplayUnit=(unit=='in')?'in':'mm';var selects=document.getElementsByTagName('select');for(var i=0;i<selects.length;i++){if((' '+(selects[i].className||'')+' ').indexOf(' system-routing-unit-select ')>=0)selects[i].value=systemRoutingDisplayUnit;}var spans=document.getElementsByTagName('span');for(var j=0;j<spans.length;j++){var cls=' '+(spans[j].className||'')+' ';if(cls.indexOf(' system-routing-size-value ')>=0||cls.indexOf(' system-layer-thickness-value ')>=0){spans[j].innerHTML=escHtml(formatSystemRoutingSize(spans[j].getAttribute('data-feet')||'',systemRoutingDisplayUnit));}}if(persist)persistSystemRoutingUnit(systemRoutingDisplayUnit);return false;}"
			+ @"function changeSystemRoutingUnit(select){return applySystemRoutingUnit((select&&select.value=='in')?'in':'mm',true);}function setSystemDisplayUnitFromHost(unit){return applySystemRoutingUnit(unit,false);}"
			+ @"function systemLayerBool(value){value=trimParam(String(value||'')).toLowerCase();return value=='true'||value=='1'||value=='yes';}"
			+ @"function systemLayerFeetFromDisplay(value){var text=trimParam(String(value||''));var match=text.match(/(-?\d+(?:\.\d+)?(?:e[+-]?\d+)?)\s*(mm|in|ft|feet|foot)?/i);if(!match)return '';var number=parseFloat(match[1]);if(!isFinite(number))return '';var unit=(match[2]||'mm').toLowerCase();if(unit=='in')return String(number/12);if(unit=='ft'||unit=='feet'||unit=='foot')return String(number);return String(number/304.8);}"
			+ @"function parseSystemLayerModel(raw,fallback){var rows=[];var lines=String(raw||'').split(/\r?\n/);for(var i=0;i<lines.length;i++){var p=(lines[i]||'').split('\t');if(p[0]!='@layer')continue;var feet=trimParam(p[4]||'');var display=trimParam(p[5]||'');if(!feet)feet=systemLayerFeetFromDisplay(display);rows.push({index:trimParam(p[1]||String(rows.length+1)),functionName:trimParam(p[2]||'-')||'-',materialName:trimParam(p[3]||'-')||'-',feet:feet,display:display||'-',isCore:systemLayerBool(p[6]),isStructural:systemLayerBool(p[7]),isVariable:systemLayerBool(p[8]),legacy:false});}if(rows.length)return rows;fallback=fallback||[];for(var j=0;j<fallback.length;j++){var item=fallback[j]||{};var pieces=String(item.value||'').split(/\s+\/\s+/);var display='';if(pieces.length&&/(-?\d+(?:\.\d+)?(?:e[+-]?\d+)?)\s*(mm|in|ft|feet|foot)\s*$/i.test(pieces[pieces.length-1]))display=trimParam(pieces.pop());var fn=pieces.length?trimParam(pieces.shift()):'-';var material=pieces.length?trimParam(pieces.join(' / ')):'-';rows.push({index:String(item.name||'').replace(/^#/,'')||String(j+1),functionName:fn||'-',materialName:material||'-',feet:systemLayerFeetFromDisplay(display),display:display||String(item.value||'-'),isCore:false,isStructural:false,isVariable:false,legacy:true});}return rows;}"
			+ @"function systemLayerThicknessHtml(layer){if(layer.feet!=='')return '<span class=""system-layer-thickness-value"" data-feet=""'+escHtml(layer.feet)+'"">'+escHtml(formatSystemRoutingSize(layer.feet,systemRoutingDisplayUnit))+'</span>';return '<span title=""'+escHtml(layer.display||'-')+'"">'+escHtml(layer.display||'-')+'</span>'; }"
			+ @"function systemLayerBadge(text,cls){return '<span class=""system-layer-badge '+(cls||'')+'"">'+escHtml(text)+'</span>'; }"
			+ @"function systemLayerCoreBoundaryHtml(){return '<tr class=""system-layer-core-boundary""><td colspan=""4"">'+escHtml(systemLayerCoreBoundaryLabel)+'</td></tr>'; }"
			+ @"function renderSystemRoutingLayers(model,raw){var layers=parseSystemLayerModel(raw,model.layers);if(!layers.length)return '';var html='<div class=""system-routing-layer-block"" data-system-layer-composition=""true""><div class=""system-routing-layer-head""><div class=""system-routing-layer-title-cell"">'+escHtml(systemRoutingLayersLabel)+'</div><div class=""system-routing-layer-unit-cell"">'+systemRoutingUnitToolbar('layers')+'</div></div><div class=""system-layer-direction exterior"">'+escHtml(systemLayerExteriorLabel)+'</div><div class=""system-layer-table-wrap""><table class=""system-layer-composition-table"" data-system-layer-table=""true""><colgroup><col class=""layer-index""/><col class=""layer-function""/><col class=""layer-material""/><col class=""layer-thickness""/></colgroup><tr><th class=""layer-index"">#</th><th>'+escHtml(systemLayerFunctionLabel)+'</th><th>'+escHtml(systemLayerMaterialLabel)+'</th><th>'+escHtml(systemLayerThicknessLabel)+'</th></tr>';var coreOpen=false;for(var i=0;i<layers.length;i++){var layer=layers[i];if(layer.isCore&&!coreOpen){html+=systemLayerCoreBoundaryHtml();coreOpen=true;}else if(!layer.isCore&&coreOpen){html+=systemLayerCoreBoundaryHtml();coreOpen=false;}var badges='';if(layer.isStructural)badges+=systemLayerBadge(systemLayerStructuralLabel,'structural');if(layer.isVariable)badges+=systemLayerBadge(systemLayerVariableLabel,'variable');var legacy=layer.legacy?'<span class=""system-layer-legacy-note"">'+escHtml(systemLayerLegacyLabel)+'</span>':'';html+='<tr class=""system-layer-row'+(layer.isCore?' layer-core':'')+'""><td class=""layer-index"">'+escHtml(layer.index||String(i+1))+'</td><td class=""layer-function"" title=""'+escHtml(layer.functionName||'-')+'"">'+escHtml(layer.functionName||'-')+badges+'</td><td title=""'+escHtml(layer.materialName||'-')+'"">'+escHtml(layer.materialName||'-')+legacy+'</td><td>'+systemLayerThicknessHtml(layer)+'</td></tr>';}if(coreOpen)html+=systemLayerCoreBoundaryHtml();return html+'</table></div><div class=""system-layer-direction interior"">'+escHtml(systemLayerInteriorLabel)+'</div></div>'; }"
			+ @"renderSystemDetailTable=function(raw){if(!isSystemDetailRaw(raw))return '<div class=""parameter-empty"">'+escHtml(noSystemDetailLabel)+'</div>';var model=parseSystemRoutingModel(raw);var html='<div class=""system-routing-shell"" data-system-routing-preferences=""true"">'+renderSystemRoutingBasic(model);if(!model.routes.length){html+='<div class=""system-routing-empty"">'+escHtml(systemRoutingNoRulesLabel)+'</div>';}else{html+='<div class=""system-routing-preference-toolbar"">'+systemRoutingUnitToolbar('routing')+'</div><div class=""system-routing-table-wrap""><table class=""system-routing-preference-table"" data-system-routing-preference-table=""true""><colgroup><col class=""system-routing-order""/><col class=""system-routing-part""/><col class=""system-routing-class""/><col class=""system-routing-count""/><col class=""system-routing-range""/></colgroup><tr><th class=""system-routing-order"">'+escHtml(systemRoutingPriorityLabel)+'</th><th>'+escHtml(systemRoutingPartLabel)+'</th><th>'+escHtml(systemRoutingClassLabel)+'</th><th class=""system-routing-count"">'+escHtml(systemRoutingSizeCountLabel)+'</th><th>'+escHtml(systemRoutingRangeLabel)+'</th></tr>';var group='';for(var i=0;i<model.routes.length;i++){var row=model.routes[i],groupKey=systemRoutingKey(row.group);if(groupKey!=group){group=groupKey;html+='<tr class=""system-routing-group-row""><td colspan=""5""><span class=""system-routing-group-dot""></span>'+escHtml(systemRoutingGroupLabel(row.group))+'</td></tr>';}var p=parseInt(row.priority,10);var priority=isNaN(p)?(row.priority||'-'):String(p+1);var cls=row.partClass||'';if(row.category)cls=cls?cls+' / '+row.category:row.category;var rangeHtml=systemRoutingMissing(row.criteria)?'<span class=""system-routing-all-sizes"">'+escHtml(systemRoutingAllSizesLabel)+'</span>':renderSystemRoutingCriteria(row.criteria);html+='<tr class=""system-routing-rule-row""><td class=""system-routing-order"">'+escHtml(priority)+'</td><td class=""system-routing-part"" title=""'+escHtml(row.part)+'"">'+escHtml(row.part)+'</td><td class=""system-routing-class"" title=""'+escHtml(cls||'-')+'"">'+escHtml(cls||'-')+'</td><td class=""system-routing-count"">'+escHtml(systemRoutingMissing(row.sizeCount)?'-':row.sizeCount)+'</td><td class=""system-routing-range"">'+rangeHtml+'</td></tr>';}html+='</table></div>';}return html+renderSystemRoutingLayers(model,raw)+'</div>';};");

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
