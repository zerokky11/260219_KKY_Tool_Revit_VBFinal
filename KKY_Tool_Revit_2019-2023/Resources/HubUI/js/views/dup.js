// Resources/HubUI/js/views/dup.js
import { clear, div, toast, showExcelSavedDialog, chooseExcelMode, showCompletionSummaryDialog, closeCompletionSummaryDialog } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { onHost, post } from '../core/bridge.js';

/*
  dup.js v75
  - 목표: "결과가 계산되는데 UI에 표시가 안 되는" 문제를 끊기 위해
         Host → JS 이벤트 수신을 모든 onHost 시그니처에 대해 강제로 호환.
  - WebView2 구버전 호환: optional chaining(?.), nullish(??) 미사용.
*/

const RESP_ROWS_EVENTS = ['dup:list','dup:rows','duplicate:list'];
const EV_RUN_REQ     = 'dup:run';
const EV_DELETE_REQ  = 'duplicate:delete';
const EV_RESTORE_REQ = 'duplicate:restore';
const EV_SELECT_REQ  = 'duplicate:select';
const EV_EXPORT_REQ  = 'duplicate:export';

const EV_DELETED_ONE  = 'dup:deleted';
const EV_RESTORED_ONE = 'dup:restored';
const EV_EXPORTED     = 'dup:exported';

const LS_MODE    = 'kky_dup_mode';
const LS_TOL_MM  = 'kky_dup_tol_mm';
const LS_TOL_UNIT = 'kky_dup_tol_unit';
const LS_SCOPE   = 'kky_dup_scope_mode';
const LS_EXCL_KW = 'kky_dup_excl_kw';
const LS_EXPORT_PARAMS = 'kky_dup_export_params';
const LS_META    = 'kky_dup_meta_v1';
const LS_RULECFG = 'kky_dup_ruleset_v1';
const LS_RULELIB = 'kky_dup_rulelib_v1';
const LS_RULELIB_SEL = 'kky_dup_rulelib_sel_v1';

const DEFAULT_TOL_MM = 0.01; // 1/64 ft

const FIELD_OPTS = [
  { v:'category', t:'Category' },
  { v:'family',   t:'Family' },
  { v:'type',     t:'Type' },
  { v:'name',     t:'Name' },
  { v:'param',    t:'Parameter' }
];
const OP_OPTS = [
  { v:'contains',    t:'Contains' },
  { v:'equals',      t:'Equal' },
  { v:'startswith',  t:'StartsWith' },
  { v:'endswith',    t:'EndsWith' },
  { v:'notcontains', t:'NotContains' },
  { v:'notequals',   t:'NotEqual' }
];

function isObj(x){ return x && typeof x === 'object'; }
function co(a,b){ return (a === undefined || a === null) ? b : a; }
function asStr(x, d){ return (x === undefined || x === null) ? (d || '') : String(x); }
function asNum(x, d){ var n = Number(x); return (isFinite(n) && !isNaN(n)) ? n : (d || 0); }
function esc(s){
  return String(co(s,''))
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/\"/g,'&quot;').replace(/'/g,'&#39;');
}
function uniq(list){
  var out = [];
  var seen = {};
  for (var i=0;i<(list||[]).length;i++){
    var s = String(list[i] || '').trim();
    if (!s) continue;
    var k = s.toLowerCase();
    if (seen[k]) continue;
    seen[k] = 1;
    out.push(s);
  }
  return out;
}
function safeId(v){ return (v === undefined || v === null) ? '' : String(v); }
function uniqNums(list){
  var out = [];
  var seen = {};
  for (var i=0;i<(list||[]).length;i++){
    var n = asNum(list[i], 0);
    if (!n) continue;
    var k = String(n);
    if (seen[k]) continue;
    seen[k] = 1;
    out.push(n);
  }
  return out;
}
function parseIdsAttr(raw){
  if (Array.isArray(raw)) return uniqNums(raw);
  var s = String(co(raw, '')).trim();
  if (!s) return [];
  return uniqNums(s.split(/[\s,;|]+/));
}
function mmToFeet(mm){ return mm / 304.8; }
function mmToInch(mm){ return mm / 25.4; }
function inchToMm(inch){ return inch * 25.4; }
function clamp(n, a, b){ return Math.max(a, Math.min(b, n)); }
function ensureId(prefix){ return String(prefix || 'S') + Math.random().toString(16).slice(2,10); }
function normalizeTolUnit(unit){
  var v = String(co(unit, 'mm')).trim().toLowerCase();
  return v === 'inch' ? 'inch' : 'mm';
}
function formatTolNumber(value){
  var n = asNum(value, 0);
  if (!isFinite(n) || isNaN(n)) return '0';
  if (Math.abs(n) >= 1) return String(Math.round(n * 1000) / 1000).replace(/\.0+$/,'').replace(/(\.\d*?)0+$/,'$1');
  return String(Math.round(n * 10000) / 10000).replace(/\.0+$/,'').replace(/(\.\d*?)0+$/,'$1');
}

// ---- config ----
function defaultRuleConfig(){
  return { version: 1, sets: [], pairs: [], excludeSetIds: [], excludeFamilies: [], excludeCategories: [] };
}
function normalizeRuleConfig(cfg){
  var c = isObj(cfg) ? cfg : {};
  var out = {
    version: 1,
    sets: Array.isArray(c.sets) ? c.sets : [],
    pairs: Array.isArray(c.pairs) ? c.pairs : [],
    excludeSetIds: Array.isArray(c.excludeSetIds) ? c.excludeSetIds.map(String).filter(Boolean) : [],
    excludeFamilies: Array.isArray(c.excludeFamilies) ? c.excludeFamilies.map(String).filter(Boolean) : [],
    excludeCategories: Array.isArray(c.excludeCategories) ? c.excludeCategories.map(String).filter(Boolean) : []
  };

  out.sets = out.sets.map(function(s, idx){
    var ss = isObj(s) ? s : {};
    var id = asStr(ss.id, 'S' + (idx + 1));
    var name = asStr(ss.name, id);
    var logic = (ss.logic === 'and' || ss.logic === 'or') ? ss.logic : 'or';
    var groups = Array.isArray(ss.groups) ? ss.groups : [];
    return { id:id, name:name, logic:logic, groups: groups };
  });

  for (var i=0;i<out.sets.length;i++){
    var s2 = out.sets[i];
    s2.groups = (s2.groups || []).map(function(g){
      var gg = isObj(g) ? g : {};
      return { clauses: Array.isArray(gg.clauses) ? gg.clauses : [] };
    });
    for (var gi=0;gi<s2.groups.length;gi++){
      var g2 = s2.groups[gi];
      g2.clauses = (g2.clauses || []).map(function(cl){
        var cc = isObj(cl) ? cl : {};
        return {
          field: asStr(cc.field,'category'),
          op: asStr(cc.op,'contains'),
          value: asStr(cc.value,''),
          param: asStr(cc.param,'')
        };
      });
    }
  }

  out.pairs = out.pairs.map(function(p){
    var pp = isObj(p) ? p : {};
    return { a: asStr(pp.a,''), b: asStr(pp.b,''), enabled: (pp.enabled !== false) };
  }).filter(function(p){ return !!p.a && !!p.b; });

  out.excludeSetIds = (out.excludeSetIds || []).map(String).filter(Boolean);
  return out;
}
function loadRuleConfig(){
  try{
    var raw = localStorage.getItem(LS_RULECFG);
    if (!raw) return defaultRuleConfig();
    return normalizeRuleConfig(JSON.parse(raw));
  } catch(e){
    return defaultRuleConfig();
  }
}
function saveRuleConfig(cfg){
  var norm = normalizeRuleConfig(cfg);
  try{ localStorage.setItem(LS_RULECFG, JSON.stringify(norm)); } catch(e){}
  return norm;
}
function loadRuleLibrary(){
  try{
    var raw = localStorage.getItem(LS_RULELIB);
    if (!raw) return [];
    var arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch(e){ return []; }
}
function saveRuleLibrary(items){
  try{ localStorage.setItem(LS_RULELIB, JSON.stringify(Array.isArray(items) ? items : [])); } catch(e){}
}
function getSelectedRuleLibraryId(){
  try{ return asStr(localStorage.getItem(LS_RULELIB_SEL), ''); } catch(e){ return ''; }
}
function setSelectedRuleLibraryId(id){
  try{ localStorage.setItem(LS_RULELIB_SEL, asStr(id, '')); } catch(e){}
}
function getSelectedRuleLibraryItem(){
  var lib = loadRuleLibrary();
  var id = getSelectedRuleLibraryId();
  for (var i=0;i<lib.length;i++) if (String(lib[i].id) === String(id)) return lib[i];
  return lib.length ? lib[0] : null;
}
function upsertRuleLibraryItem(fileName, parsed){
  var lib = loadRuleLibrary();
  var name = asStr(fileName, '').trim() || 'DuplicateInspector_RuleSet.xml';
  var item = {
    id: ensureId('CFG'),
    name: name,
    importedAt: new Date().toISOString(),
    parsed: {
      tolMm: clamp(asNum(parsed.tolMm, DEFAULT_TOL_MM), 0.01, 1000),
      tolUnit: normalizeTolUnit(parsed.tolUnit || 'mm'),
      scopeMode: asStr(parsed.scopeMode, 'all'),
      excludeKeywords: asStr(parsed.excludeKeywords, ''),
      exportParams: asStr(parsed.exportParams, ''),
      ruleConfig: normalizeRuleConfig(parsed.ruleConfig)
    }
  };
  var found = -1;
  for (var i=0;i<lib.length;i++) if (String(lib[i].name).toLowerCase() === name.toLowerCase()) { found = i; break; }
  if (found >= 0) item.id = lib[found].id;
  if (found >= 0) lib[found] = item; else lib.unshift(item);
  saveRuleLibrary(lib);
  setSelectedRuleLibraryId(item.id);
  return item;
}
function applyRuleLibraryItem(id){
  var item = null;
  var lib = loadRuleLibrary();
  for (var i=0;i<lib.length;i++) if (String(lib[i].id) === String(id)) { item = lib[i]; break; }
  if (!item || !item.parsed) return false;
  try{ localStorage.setItem(LS_TOL_MM, String(item.parsed.tolMm)); } catch(e){}
  try{ localStorage.setItem(LS_TOL_UNIT, normalizeTolUnit(item.parsed.tolUnit || 'mm')); } catch(e1){}
  try{ localStorage.setItem(LS_SCOPE, asStr(item.parsed.scopeMode, 'all')); } catch(e2){}
  try{ localStorage.setItem(LS_EXCL_KW, asStr(item.parsed.excludeKeywords, '')); } catch(e3){}
  try{ localStorage.setItem(LS_EXPORT_PARAMS, asStr(item.parsed.exportParams, '')); } catch(e4){}
  saveRuleConfig(item.parsed.ruleConfig);
  setSelectedRuleLibraryId(item.id);
  return true;
}
function removeRuleLibraryItem(id){
  var lib = loadRuleLibrary().filter(function(x){ return String(x.id) !== String(id); });
  saveRuleLibrary(lib);
  if (String(getSelectedRuleLibraryId()) === String(id)) setSelectedRuleLibraryId(lib.length ? lib[0].id : '');
}
function clearRuleLibrary(){ saveRuleLibrary([]); setSelectedRuleLibraryId(''); }
function sanitizeRuleFileName(name){
  var base = asStr(name, '').trim() || 'DuplicateInspector_RuleSet.xml';
  base = base.replace(/[\/:*?"<>|]+/g, '_');
  if (!/\.xml$/i.test(base)) base += '.xml';
  return base;
}
var dupUiHooks = { renderRulePanel:null, renderAppliedBar:null, serializeRuleFile:null };
function callDupUiHook(name){
  try{
    var fn = dupUiHooks && dupUiHooks[name];
    if (typeof fn === 'function') return fn();
  }catch(e){}
  return undefined;
}
async function exportRuleFileInteractive(){
  var xml = callDupUiHook('serializeRuleFile');
  if (!xml){
    try{
      if (typeof serializeRuleFile === 'function') xml = serializeRuleFile();
      else if (typeof SerializeRuleFile === 'function') xml = SerializeRuleFile();
    }catch(e0){}
  }
  if (!xml){ toast('설정 XML 직렬화 함수를 찾지 못했습니다.', 'err', 2200); return; }
  var sel = getSelectedRuleLibraryItem();
  var suggested = sanitizeRuleFileName(sel && sel.name ? sel.name : 'DuplicateInspector_RuleSet.xml');
  if (window.showSaveFilePicker){
    try{
      var handle = await window.showSaveFilePicker({ suggestedName: suggested, types:[{ description:'XML 파일', accept:{ 'application/xml':['.xml'] } }] });
      var writable = await handle.createWritable();
      await writable.write(xml);
      await writable.close();
      toast('설정 XML 파일을 저장했습니다.', 'ok', 1400);
      return;
    } catch(ex){
      if (ex && ex.name === 'AbortError') return;
    }
  }
  try{
    var fileName = sanitizeRuleFileName(window.prompt('저장할 파일 이름을 입력하세요.', suggested) || suggested);
    if (!fileName) return;
    var blob = new Blob([xml], { type: 'application/xml;charset=utf-8' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    setTimeout(function(){ try{ URL.revokeObjectURL(a.href); }catch(e){} try{ document.body.removeChild(a); }catch(e2){} }, 0);
    toast('저장 위치 선택 API가 없어서 브라우저 다운로드로 저장했습니다.', 'warn', 2200);
  } catch(ex2){
    toast('설정 XML 파일 저장에 실패했습니다.', 'err', 2400);
  }
}
function loadMeta(){
  try{
    var raw = localStorage.getItem(LS_META);
    if (!raw) return {};
    var v = JSON.parse(raw);
    // v could be a JSON string if stored incorrectly in older builds
    if (typeof v === 'string'){
      var s = String(v || '').trim();
      if (s && (s[0] === '{' || s[0] === '[')){
        try{ v = JSON.parse(s); }catch(e){}
      }
    }
    return (v && typeof v === 'object') ? v : {};
  } catch(e){ return {}; }
}

function getExportParamText(){
  try{
    return String(localStorage.getItem(LS_EXPORT_PARAMS) || '').trim();
  } catch(e){
    return '';
  }
}

function getExportParamNames(){
  var raw = getExportParamText();
  if (!raw) return [];
  return uniq(raw.split(/[\r\n,;|]+/).map(function(x){ return String(x || '').trim(); }).filter(Boolean));
}

// ---- Exclude Picker (패밀리/시스템) ----
var exPickerEl = null;
var exPickerOpen = false;
var exPickerDraft = { fam: {}, cat: {} };
var exPickerMetaPollTimer = 0;

function resetExcludePickerDraftFromConfig(){
  var cfg = loadRuleConfig();
  var famSel = {};
  var catSel = {};
  (cfg.excludeFamilies || []).forEach(function(x){ var s = String(x || '').trim(); if (s) famSel[s] = 1; });
  (cfg.excludeCategories || []).forEach(function(x){ var s = String(x || '').trim(); if (s) catSel[s] = 1; });
  exPickerDraft = { fam: famSel, cat: catSel };
}

function syncExcludePickerDraftFromDom(){
  if (!exPickerEl) return;
  var famSel = Object.assign({}, (exPickerDraft && exPickerDraft.fam) ? exPickerDraft.fam : {});
  var catSel = Object.assign({}, (exPickerDraft && exPickerDraft.cat) ? exPickerDraft.cat : {});
  var cbs = exPickerEl.querySelectorAll('input[type="checkbox"]');
  for (var i=0;i<cbs.length;i++) {
    var cb = cbs[i];
    if (!cb) continue;
    var kind = cb.dataset ? cb.dataset.kind : '';
    var key = String(cb.value || '').trim();
    if (!key) continue;
    if (kind === 'fam') {
      if (cb.checked) famSel[key] = 1; else delete famSel[key];
    } else if (kind === 'cat') {
      if (cb.checked) catSel[key] = 1; else delete catSel[key];
    }
  }
  exPickerDraft = { fam: famSel, cat: catSel };
}

function shouldHideSystemTypeItem(cat, typ, label){
  var catLc = String(cat || '').trim().toLowerCase();
  var typLc = String(typ || '').trim().toLowerCase();
  var labelLc = String(label || '').trim().toLowerCase();
  function hasAny(s, arr){
    for (var i=0;i<arr.length;i++){ if (s.indexOf(arr[i]) >= 0) return true; }
    return false;
  }
  if (hasAny(catLc, ['rvt link','revit link','cad link','dwg link','import', 'camera', 'view']) || catLc === 'views') return true;
  if (hasAny(labelLc, ['dwg','camera','revit link','cad link','dwg link']) || hasAny(typLc, ['dwg','camera','revit link','cad link','dwg link'])) return true;
  return false;
}

function hasExcludePickerMeta(metaData){
  return !!(metaData && ((metaData.fams && metaData.fams.length) || (metaData.sysItems && metaData.sysItems.length) || (metaData.cats && metaData.cats.length)));
}

function stopExcludePickerMetaPoll(){
  if (exPickerMetaPollTimer){
    try{ clearInterval(exPickerMetaPollTimer); }catch(e){}
    exPickerMetaPollTimer = 0;
  }
}

function startExcludePickerMetaPoll(){
  stopExcludePickerMetaPoll();
  var tries = 0;
  exPickerMetaPollTimer = setInterval(function(){
    tries++;
    if (!exPickerOpen){
      stopExcludePickerMetaPoll();
      return;
    }
    try{ renderExcludePicker(); }catch(e){}
    var md = getExcludePickerMeta();
    if (hasExcludePickerMeta(md) || tries >= 40){
      stopExcludePickerMetaPoll();
    }
  }, 150);
}

function getExcludePickerMeta(){
  var meta = loadMeta() || {};
  var fams = uniq((Array.isArray(meta.modelFamilies) ? meta.modelFamilies : []).map(function(x){ return String(x || '').trim(); }).filter(Boolean));

  var sysRaw = Array.isArray(meta.systemTypes) ? meta.systemTypes : (Array.isArray(meta.systemFamilyTypes) ? meta.systemFamilyTypes : (Array.isArray(meta.systemTypeNames) ? meta.systemTypeNames : null));
  var sysItems = [];
  var seen = {};
  var cats = [];

  if (sysRaw && sysRaw.length){
    for (var si=0; si<sysRaw.length; si++) {
      var it = sysRaw[si];
      if (!it) continue;
      var cat = '';
      var typ = '';
      var label = '';
      var key = '';
      if (typeof it === 'string') {
        label = String(it || '').trim();
        key = label;
      } else if (typeof it === 'object') {
        cat = String(it.category || it.cat || '').trim();
        typ = String(it.type || it.name || it.typeName || '').trim();
        label = (cat && typ) ? (cat + ' : ' + typ) : (cat || typ);
        key = (cat && typ) ? label : (typ || cat || label);
      }
      label = String(label || '').trim();
      key = String(key || label).trim();
      if (!label || !key) continue;
      if (shouldHideSystemTypeItem(cat, typ, label)) continue;
      if (seen[key]) continue;
      seen[key] = 1;
      sysItems.push({ key:key, label:label });
    }
    sysItems.sort(function(a,b){ return String(a.label).localeCompare(String(b.label)); });
  } else {
    cats = uniq((Array.isArray(meta.systemCategories) ? meta.systemCategories : []).map(function(x){ return String(x || '').trim(); }).filter(Boolean));
    cats = cats.filter(function(c){ return !shouldHideSystemTypeItem(c, '', c); });
    cats.sort(function(a,b){ return String(a).localeCompare(String(b)); });
  }

  fams.sort(function(a,b){ return String(a).localeCompare(String(b)); });
  return { fams:fams, sysItems:sysItems, cats:cats };
}

function ensureExcludePicker(){
  var el = document.getElementById('dup-expicker');
  if (!el){
    el = document.createElement('div');
    el.id = 'dup-expicker';
    el.className = 'dup-expicker';
    el.style.display = 'none';

    var backdrop = document.createElement('div');
    backdrop.className = 'rm-backdrop';

    var win = document.createElement('div');
    win.className = 'rm-window';

    var head = document.createElement('div');
    head.className = 'rm-head';

    var headLeft = document.createElement('div');
    headLeft.className = 'rm-head-left';

    var t = document.createElement('div');
    t.className = 'rm-title';
    t.textContent = 'Exclude 목록 선택';

    var sub = document.createElement('div');
    sub.className = 'rm-sub';
    sub.textContent = '모델 패밀리와 시스템 목록을 선택해 간섭 검토에서 제외합니다.';

    headLeft.appendChild(t);
    headLeft.appendChild(sub);

    var headAct = document.createElement('div');
    headAct.className = 'rm-actions';

    var btnClose = document.createElement('button');
    btnClose.className = 'rp-btn';
    btnClose.textContent = '닫기';
    btnClose.dataset.act = 'close';

    headAct.appendChild(btnClose);

    head.appendChild(headLeft);
    head.appendChild(headAct);

    var body = document.createElement('div');
    body.className = 'rm-body';

    var sec = document.createElement('div');
    sec.className = 'rp-sec';
    var secTitle = document.createElement('div');
    secTitle.className = 'rp-sec-title';
    secTitle.textContent = '검색';
    sec.appendChild(secTitle);

    var picked = document.createElement('div');
    picked.className = 'expicked';
    picked.dataset.slot = 'picked';

    var grid = document.createElement('div');
    grid.className = 'rp-grid';
    var lbl = document.createElement('div');
    lbl.className = 'rp-label';
    lbl.textContent = '키워드';
    var q = document.createElement('input');
    q.className = 'rp-input';
    q.type = 'text';
    q.placeholder = '패밀리 / 시스템 이름 검색';
    q.dataset.bind = 'q';
    grid.appendChild(lbl);
    grid.appendChild(q);
    sec.appendChild(grid);

    var wrap = document.createElement('div');
    wrap.className = 'exfam-wrap';

    var colFam = document.createElement('div');
    colFam.className = 'exfam-col';
    var famTitle = document.createElement('div');
    famTitle.className = 'exfam-title';
    famTitle.textContent = '모델 패밀리';
    var famList = document.createElement('div');
    famList.className = 'exfam-list';
    famList.dataset.slot = 'fam';
    colFam.appendChild(famTitle);
    colFam.appendChild(famList);

    var colCat = document.createElement('div');
    colCat.className = 'exfam-col';
    var catTitle = document.createElement('div');
    catTitle.className = 'exfam-title';
    catTitle.textContent = '시스템(카테고리)';
    var catList = document.createElement('div');
    catList.className = 'exfam-list';
    catList.dataset.slot = 'cat';
    colCat.appendChild(catTitle);
    colCat.appendChild(catList);

    wrap.appendChild(colFam);
    wrap.appendChild(colCat);

    var actRow = document.createElement('div');
    actRow.className = 'exfam-actions';

    var btnAll = document.createElement('button');
    btnAll.className = 'rp-btn rp-btn--ghost';
    btnAll.textContent = '전체 체크';
    btnAll.dataset.act = 'all';

    var btnNone = document.createElement('button');
    btnNone.className = 'rp-btn rp-btn--ghost';
    btnNone.textContent = '전체 해제';
    btnNone.dataset.act = 'none';

    var btnApply = document.createElement('button');
    btnApply.className = 'rp-btn rp-btn--primary';
    btnApply.textContent = '적용';
    btnApply.dataset.act = 'apply';

    actRow.appendChild(btnAll);
    actRow.appendChild(btnNone);
    actRow.appendChild(btnApply);

    body.appendChild(sec);
    body.appendChild(picked);
    body.appendChild(wrap);
    body.appendChild(actRow);

    win.appendChild(head);
    win.appendChild(body);

    el.appendChild(backdrop);
    el.appendChild(win);

    document.body.appendChild(el);

    el.addEventListener('click', function(e){
      var t2 = e.target;
      var act = t2 && t2.dataset ? t2.dataset.act : '';
      if (act === 'close' || (t2 && t2.classList && t2.classList.contains('rm-backdrop'))){
        closeExcludePicker();
        return;
      }
      if (act === 'clearq'){
        try{ var qEl2 = el.querySelector('[data-bind="q"]'); if (qEl2) qEl2.value = ''; }catch(e2){}
        renderExcludePicker();
        return;
      }
      if (act === 'all'){
        var cbs = el.querySelectorAll('input[type="checkbox"]');
        for (var i=0;i<cbs.length;i++) cbs[i].checked = true;
        syncExcludePickerDraftFromDom();
        renderExcludePicker();
        return;
      }
      if (act === 'none'){
        var cbs2 = el.querySelectorAll('input[type="checkbox"]');
        for (var j=0;j<cbs2.length;j++) cbs2[j].checked = false;
        syncExcludePickerDraftFromDom();
        renderExcludePicker();
        return;
      }
      if (act === 'apply'){
        applyExcludePicker();
        return;
      }
    });

    el.addEventListener('input', function(e){
      var t3 = e.target;
      if (t3 && t3.dataset && t3.dataset.bind === 'q'){
        renderExcludePicker();
      }
    });

    el.addEventListener('change', function(e){
      var t4 = e.target;
      if (t4 && t4.type === 'checkbox'){
        syncExcludePickerDraftFromDom();
        renderExcludePicker();
      }
    });
  }
  exPickerEl = el;
}
function openExcludePicker(){
  ensureExcludePicker();
  resetExcludePickerDraftFromConfig();
  exPickerOpen = true;
  exPickerEl.style.display = 'block';
  exPickerEl.classList.add('is-open');

  try{
    var qEl = exPickerEl.querySelector('[data-bind="q"]');
    if (qEl) qEl.value = '';
  }catch(e){}

  var metaData = getExcludePickerMeta();
  renderExcludePicker();

  if (!hasExcludePickerMeta(metaData)) {
    startExcludePickerMetaPoll();
  } else {
    stopExcludePickerMetaPoll();
  }

  if (!metaRefreshPending){
    metaRefreshPending = true;
    try{ post(EV_RUN_REQ, { mode: readMode(), metaOnly: true }); }catch(e2){}
    startExcludePickerMetaPoll();
  }
}
function closeExcludePicker(){
  if (!exPickerEl) return;
  exPickerOpen = false;
  stopExcludePickerMetaPoll();
  exPickerEl.classList.remove('is-open');
  exPickerEl.style.display = 'none';
}
function renderExcludePicker(){
  if (!exPickerEl) return;
  var famSel = exPickerDraft && exPickerDraft.fam ? exPickerDraft.fam : {};
  var catSel = exPickerDraft && exPickerDraft.cat ? exPickerDraft.cat : {};

  var metaData = getExcludePickerMeta();
  var fams = metaData.fams;
  var sysItems = metaData.sysItems;
  var cats = metaData.cats;

  var qEl = exPickerEl.querySelector('[data-bind="q"]');
  var q = qEl ? String(qEl.value || '').trim().toLowerCase() : '';

  var pickedBox = exPickerEl.querySelector('[data-slot="picked"]');
  if (pickedBox){
    var famPicked = Object.keys(famSel).sort(function(a,b){ return String(a).localeCompare(String(b)); });
    var catPicked = Object.keys(catSel).sort(function(a,b){ return String(a).localeCompare(String(b)); });
    var chips = [];
    for (var pf=0; pf<famPicked.length; pf++) chips.push('<span class="expicked-chip">패밀리 · ' + esc(famPicked[pf]) + '</span>');
    for (var pc=0; pc<catPicked.length; pc++) chips.push('<span class="expicked-chip">시스템 · ' + esc(catPicked[pc]) + '</span>');
    pickedBox.innerHTML = '<div class="rp-sec-title">선택된 제외 목록</div>' +
      '<div class="rp-sec-sub">체크한 항목이 여기 누적됩니다. 적용을 누르면 현재 목록으로 저장됩니다.</div>' +
      '<div class="expicked-meta"><span class="expicked-count">패밀리 ' + famPicked.length + '개</span><span class="expicked-count">시스템 ' + catPicked.length + '개</span></div>' +
      '<div class="expicked-list">' + (chips.length ? chips.join('') : '<span class="rp-emptyline">아직 체크한 제외 항목이 없습니다.</span>') + '</div>';
  }

  function mk(kind, val, checked){
    var lab = document.createElement('label');
    lab.className = 'exfam-item';
    var cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.dataset.kind = kind;
    cb.value = String(val);
    cb.checked = !!checked;
    var sp = document.createElement('span');
    sp.className = 'exfam-text';
    sp.textContent = String(val);
    lab.appendChild(cb);
    lab.appendChild(sp);
    return lab;
  }

  var famBox = exPickerEl.querySelector('[data-slot="fam"]');
  var catBox = exPickerEl.querySelector('[data-slot="cat"]');
  if (!famBox || !catBox) return;

  famBox.innerHTML = '';
  var famCount = 0;
  for (var i=0;i<fams.length;i++){
    var f = String(fams[i] || '');
    if (!f) continue;
    if (q && f.toLowerCase().indexOf(q) < 0) continue;
    famBox.appendChild(mk('fam', f, !!famSel[f]));
    famCount++;
  }
  if (!famCount){
    var e = document.createElement('div');
    e.className = 'exfam-empty';
    e.textContent = (fams.length ? '검색 조건으로 0개입니다. 검색어를 비워 보세요.' : '목록을 자동으로 불러오는 중입니다.');
    famBox.appendChild(e);
  }

  catBox.innerHTML = '';
  var catCount = 0;
  if (sysItems && sysItems.length){
    for (var j=0;j<sysItems.length;j++){
      var it = sysItems[j];
      if (!it) continue;
      var label = String(it.label || '').trim();
      var key = String(it.key || label).trim();
      if (!label) continue;
      if (q && label.toLowerCase().indexOf(q) < 0) continue;
      catBox.appendChild(mk('cat', key, !!catSel[key]));
      try{ catBox.lastChild.querySelector('.exfam-text').textContent = label; }catch(e){}
      catCount++;
    }
  } else {
    for (var j=0;j<cats.length;j++){
      var c = String(cats[j] || '');
      if (!c) continue;
      if (q && c.toLowerCase().indexOf(q) < 0) continue;
      catBox.appendChild(mk('cat', c, !!catSel[c]));
      catCount++;
    }
  }
  if (!catCount){
    var e2 = document.createElement('div');
    e2.className = 'exfam-empty';
    e2.textContent = ((sysItems && sysItems.length) || (cats && cats.length) ? '검색 조건으로 0개입니다. 검색어를 비워 보세요.' : '목록을 자동으로 불러오는 중입니다.');
    catBox.appendChild(e2);
  }
}
function applyExcludePicker(){
  if (!exPickerEl) return;
  syncExcludePickerDraftFromDom();
  var cfg = loadRuleConfig();
  var fam = Object.keys((exPickerDraft && exPickerDraft.fam) ? exPickerDraft.fam : {}).sort(function(a,b){ return String(a).localeCompare(String(b)); });
  var cat = Object.keys((exPickerDraft && exPickerDraft.cat) ? exPickerDraft.cat : {}).sort(function(a,b){ return String(a).localeCompare(String(b)); });

  cfg.excludeFamilies = fam;
  cfg.excludeCategories = cat;
  saveRuleConfig(cfg);
  toast('Exclude 목록이 적용되었습니다.', 'ok', 1200);
  closeExcludePicker();
  callDupUiHook('renderRulePanel');
  callDupUiHook('renderAppliedBar');
}

// ---- Host 이벤트 수신 “강제 호환” ----
function normIn(a,b){
  // 반환: { ev, payload }
  if (typeof a === 'string'){
    return { ev: a, payload: b };
  }
  if (isObj(a) && (a.ev || a.event || a.name)){
    var ev = a.ev || a.event || a.name || '';
    var payload = (a.payload !== undefined) ? a.payload :
                  (a.data !== undefined) ? a.data :
                  (a.body !== undefined) ? a.body : b;
    return { ev: ev, payload: payload };
  }
  // event-specific onHost(ev, fn) 형태면 a가 payload일 수 있음
  return { ev: '', payload: a };
}

function extractList(payload){
  if (!payload) return [];
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload.rows)) return payload.rows;
  if (Array.isArray(payload.data)) return payload.data;
  if (Array.isArray(payload.list)) return payload.list;
  return [];
}

// ---------------------- main view ----------------------
export function renderDup(root){
  var target = root || document.getElementById('view-root') || document.getElementById('app');
  if (!target) return;

  ensureStyles();
  ensureExcludePicker();

  clear(target);

  var mode = readMode();
  var scopeMode = readScopeMode();

  var busy = false;
  var exporting = false;
  var lastCompletionSignature = '';

  var rows = [];
  var groups = [];
  var deleted = {};
  var expanded = {};

  var lastResult = null;
  var lastPairs = [];
  var metaRefreshPending = false;
  var lastRun = null;

  var page = div('dup-page feature-shell');
  var header = div('feature-header dup-toolbar');
  var heading = div('feature-heading');
  var actions = div('feature-actions');

  var runBtn = makeBtn('검토 시작', 'card-action-btn');
  var exportBtn = makeBtn('엑셀 내보내기', 'card-action-btn');
  exportBtn.disabled = true;

  var btnDup = makeBtn('중복검토', 'dup-mode-btn');
  var btnClash = makeBtn('자체간섭', 'dup-mode-btn');
  var settingsBtn = makeBtn('Rule/Set 설정', 'dup-settings-btn');

  var modeShell = div('dup-mode-shell');
  var modeLabel = div('dup-mode-label');
  modeLabel.textContent = '검토 모드';
  var modeSwitch = div('dup-mode-switch');
  modeSwitch.appendChild(btnDup);
  modeSwitch.appendChild(btnClash);
  modeShell.appendChild(modeLabel);
  modeShell.appendChild(modeSwitch);

  var settingsShell = div('dup-settings-shell');
  var settingsLabel = div('dup-mode-label');
  settingsLabel.textContent = '설정';
  settingsShell.appendChild(settingsLabel);
  settingsShell.appendChild(settingsBtn);

  var summaryBar = div('dup-summarybar sticky hidden');
  var appliedBar = div('dup-appliedbar hidden');
  var body = div('dup-body');

  actions.appendChild(runBtn);
  actions.appendChild(modeShell);
  actions.appendChild(settingsShell);
  actions.appendChild(exportBtn);

  header.appendChild(heading);
  header.appendChild(actions);

  page.appendChild(header);
  page.appendChild(summaryBar);
  page.appendChild(appliedBar);
  page.appendChild(body);
  target.appendChild(page);

  var rulePanel = buildRulePanel();
  page.appendChild(rulePanel);

  dupUiHooks.renderRulePanel = function(){ renderRulePanel(); };
  dupUiHooks.renderAppliedBar = function(){ renderAppliedBar(); };
  dupUiHooks.serializeRuleFile = function(){ return serializeRuleFile(); };

  syncModeButtons();
  syncHeading();
  renderIntro();
  renderAppliedBar();

  // --- event routing ---
  function route(ev, payload){
    try{
      if (!ev) return;

      if (RESP_ROWS_EVENTS.indexOf(ev) >= 0){
        setLoading(false);
        handleRows(extractList(payload));
        return;
      }

      if (ev === 'dup:result'){
        setLoading(false);
        lastResult = payload || null;
        // 결과가 list 없이 result만 오는 케이스 대비: payload.rows/data도 처리
        var list = extractList(payload);
        if (list.length){
          handleRows(list);
        } else if (!rows.length && !lastPairs.length && lastRun){
          exportBtn.disabled = true;
          showNoResultsState();
          refreshSummary();
          renderAppliedBar();
        } else {
          paint();
          refreshSummary();
          renderAppliedBar();
        }
        showDupCompletionDialog();
        return;
      }

      if (ev === 'dup:pairs'){
        setLoading(false);
        lastPairs = Array.isArray(payload) ? payload : extractList(payload);
        lastPairs = dedupPairs(lastPairs);
        exportBtn.disabled = (busy || (!rows.length && !lastPairs.length));
        if (!lastPairs.length && !rows.length && lastRun){
          showNoResultsState();
        } else {
          paint();
        }
        refreshSummary();
        renderAppliedBar();
        return;
      }

      if (ev === EV_DELETED_ONE){
        var id = safeId(payload && payload.id);
        if (id) { deleted[id] = 1; paintRowStates(); refreshSummary(); }
        return;
      }
      if (ev === EV_RESTORED_ONE){
        var id2 = safeId(payload && payload.id);
        if (id2) { delete deleted[id2]; paintRowStates(); refreshSummary(); }
        return;
      }

      if (ev === EV_EXPORTED){
        exporting = false;
        ProgressDialog.hide();
        var path = asStr(payload && payload.path, '');
        if (path){
          showExcelSavedDialog('엑셀로 내보냈습니다.', path, function(p){
            if (p) post('excel:open', { path: p });
          });
        } else {
          toast(asStr(payload && payload.message, '엑셀 내보내기 실패'), 'err', 2800);
        }
        exportBtn.disabled = (busy || (!rows.length && !lastPairs.length));
        return;
      }

      if (ev === 'dup:progress'){
        handleExcelProgress(payload || {});
        return;
      }

      if (ev === 'dup:meta'){
        try{
          var meta = payload;
          // payload may arrive as JSON string depending on bridge implementation
          if (typeof meta === 'string'){
            var s = String(meta || '').trim();
            if (s && (s[0] === '{' || s[0] === '[')){
              try{ meta = JSON.parse(s); }catch(e){}
            }
          }
          // sometimes wrapped
          if (meta && typeof meta === 'object' && meta.meta && typeof meta.meta === 'object') meta = meta.meta;
          if (!meta || typeof meta !== 'object') meta = {};
          try{ localStorage.setItem(LS_META, JSON.stringify(meta)); } catch(e){}
          metaRefreshPending = false;
          if (exPickerOpen){
            renderExcludePicker();
            stopExcludePickerMetaPoll();
          }
          renderRulePanel();
        } catch(e2){}
        return;
      }

      if (ev === 'host:warn'){
        setLoading(false);
        var msg = asStr(payload && payload.message, '');
        if (msg) toast(msg, 'warn', 3200);
        return;
      }
      if (ev === 'host:error' || ev === 'revit:error'){
        setLoading(false);
        exporting = false;
        ProgressDialog.hide();
        toast(asStr(payload && payload.message, '오류가 발생했습니다.'), 'err', 3400);
        exportBtn.disabled = (busy || (!rows.length && !lastPairs.length));
        return;
      }
    } catch(ex){
      try{ console.error('[dup] route error', ev, ex); } catch(e){}
      toast('dup.js 오류: ' + asStr(ex && ex.message, 'unknown'), 'err', 3200);
    }
  }

  // 1) 이벤트명 구독(가능한 경우)
  onHost('dup:meta', function(p){ route('dup:meta', p); });

  for (var i=0;i<RESP_ROWS_EVENTS.length;i++){
    (function(evName){
      onHost(evName, function(p){ route(evName, p); });
    })(RESP_ROWS_EVENTS[i]);
  }

  onHost('dup:result', function(p){ route('dup:result', p); });
  onHost('dup:pairs', function(p){ route('dup:pairs', p); });
  onHost('dup:progress', function(p){ route('dup:progress', p); });
  onHost(EV_DELETED_ONE, function(p){ route(EV_DELETED_ONE, p); });
  onHost(EV_RESTORED_ONE, function(p){ route(EV_RESTORED_ONE, p); });
  onHost(EV_EXPORTED, function(p){ route(EV_EXPORTED, p); });
  onHost('host:warn', function(p){ route('host:warn', p); });
  onHost('host:error', function(p){ route('host:error', p); });
  onHost('revit:error', function(p){ route('revit:error', p); });

  // 2) wildcard 구독(환경에 따라 이게 유일하게 동작하는 케이스 대비)
  onHost(function(a,b){
    var x = normIn(a,b);
    if (x.ev) route(x.ev, x.payload);
  });

  // ---- UI events ----
  runBtn.addEventListener('click', function(){
    if (busy) return;
    closeCompletionSummaryDialog();
    lastCompletionSignature = '';

    rows = [];
    groups = [];
    deleted = {};
    expanded = {};
    lastPairs = [];
    lastResult = null;
    lastRun = null;
    renderAppliedBar();
    body.innerHTML = '';
    body.appendChild(buildSkeleton(6));
    exportBtn.disabled = true;

    setLoading(true);

    var cfg = loadRuleConfig();
    var kw = getExcludeKeywords();
    var merged = uniq([].concat(kw, cfg.excludeFamilies || [], cfg.excludeCategories || []));

    var tolFeet = readTolFeet();
    lastRun = {
      mode: mode,
      tolMm: readTolMm(),
      tolFeet: tolFeet,
      scopeMode: scopeMode,
      excludeKeywordsRaw: getExcludeKeywords(),
      excludeKeywordsMerged: merged,
      ruleConfig: cfg
    };
    renderAppliedBar();
    post(EV_RUN_REQ, {
      mode: mode,
      tolFeet: tolFeet,
      scopeMode: scopeMode,
      excludeKeywords: merged,
      ruleConfig: cfg
    });

    // "응답 없음" 안내
    var t = setTimeout(function(){
      if (!busy) return;
      setLoading(false);
      body.innerHTML = '';
      renderIntro();
      toast('응답이 없습니다. (Host→JS 이벤트 미수신)', 'err', 3400);
    }, 12000);

    // 어떤 결과 이벤트든 오면 타임아웃 해제
    onHost(function(a,b){
      var x = normIn(a,b);
      if (!x.ev) return;
      if (RESP_ROWS_EVENTS.indexOf(x.ev) >= 0 || x.ev === 'dup:result' || x.ev === 'dup:pairs' || x.ev === 'host:error' || x.ev === 'revit:error'){
        try{ clearTimeout(t); } catch(e){}
      }
    });
  });

  exportBtn.addEventListener('click', function(){
    if (exporting) return;
    if (!rows.length && !lastPairs.length) return;

    exporting = true;
    exportBtn.disabled = true;
    var exportParamNames = getExportParamNames();

    chooseExcelMode(function(excelMode){
      post(EV_EXPORT_REQ, {
        excelMode: excelMode || 'fast',
        exportParamNames: exportParamNames
      });
    });
  });

  btnDup.addEventListener('click', function(){ setMode('duplicate'); });
  btnClash.addEventListener('click', function(){ setMode('clash'); });

  settingsBtn.addEventListener('click', function(){
    var open = !rulePanel.classList.contains('is-open');
    rulePanel.classList.toggle('is-open', open);
    settingsBtn.classList.toggle('is-active', open);
    if (open){
      renderRulePanel();
      renderAppliedBar();
      if (!metaRefreshPending){
        metaRefreshPending = true;
        try{ post(EV_RUN_REQ, { mode: readMode(), metaOnly: true }); }catch(e){}
      }
    }
  });

  body.addEventListener('click', function(e){
    var t = e.target;

    var sel = t && t.closest ? t.closest('[data-act="sel"]') : null;
    if (sel){
      var id = asNum(sel.dataset.id, 0);
      if (id > 0) post(EV_SELECT_REQ, { id: id, zoom: true });
      return;
    }

    var act = t && t.closest ? t.closest('[data-act]') : null;
    if (!act) return;
    var a = act.dataset.act;

    if (a === 'sel-group' || a === 'sel-pair'){
      var ids0 = parseIdsAttr(act.dataset.ids);
      if (ids0.length) post(EV_SELECT_REQ, { ids: ids0, zoom: true });
      return;
    }

    var id2 = asNum(act.dataset.id, 0);
    if (!id2) return;

    if (a === 'zoom'){
      post(EV_SELECT_REQ, { id: id2, zoom: true });
      return;
    }
    if (a === 'del'){
      post(EV_DELETE_REQ, { id: id2, ids: [id2] });
      return;
    }
    if (a === 'res'){
      post(EV_RESTORE_REQ, { id: id2, ids: [id2] });
      return;
    }
  });

  // ---- functions ----
  function setMode(m){
    if (busy) return;
    if (m !== 'duplicate' && m !== 'clash') return;
    mode = m;
    try{ localStorage.setItem(LS_MODE, m); } catch(e){}
    syncModeButtons();
    syncHeading();

    rows = [];
    groups = [];
    deleted = {};
    expanded = {};
    lastPairs = [];
    lastResult = null;

    body.innerHTML = '';
    renderIntro();
    refreshSummary();
    renderAppliedBar();
    exportBtn.disabled = true;
  }

  function syncModeButtons(){
    btnDup.classList.toggle('is-active', mode === 'duplicate');
    btnClash.classList.toggle('is-active', mode === 'clash');
  }

  function syncHeading(){
    var title = (mode === 'clash') ? '자체간섭 검토' : '중복검토';
    var sub2 = (mode === 'clash')
      ? '같은 파일 내 자체간섭 후보를 A↔B 쌍으로 표시합니다.'
      : '중복 요소 후보를 그룹별로 확인하고 삭제/되돌리기를 관리합니다.';
    heading.innerHTML =
      '<span class="feature-kicker">Duplicate Inspector</span>' +
      '<h2 class="feature-title">' + esc(title) + '</h2>' +
      '<p class="feature-sub">' + esc(sub2) + '</p>';
  }

  function setLoading(on){
    busy = !!on;
    runBtn.disabled = busy;
    btnDup.disabled = busy;
    btnClash.disabled = busy;
    settingsBtn.disabled = busy;
    exportBtn.disabled = busy || exporting || (!rows.length && !lastPairs.length);
    runBtn.textContent = busy ? '검토 중…' : '검토 시작';
  }

  function showNoResultsState(){
    body.innerHTML = '';
    var empty = div('dup-emptycard');
    empty.innerHTML = '<div class="empty-emoji">✅</div><h3 class="empty-title">' +
      ((mode === 'clash') ? '간섭이 없습니다' : '중복이 없습니다') +
      '</h3><p class="empty-sub">검토 결과가 0건입니다. 오류가 아니라, 현재 조건에서 검토 대상이 발견되지 않은 상태입니다.</p>';
    body.appendChild(empty);
  }

  function handleRows(list){
    rows = (Array.isArray(list) ? list : []).map(normalizeRow);
    groups = buildGroups(rows);

    expanded = {};
    for (var i=0;i<groups.length;i++) expanded[groups[i].key] = 1;

    exportBtn.disabled = (busy || (!rows.length && !lastPairs.length));
    if (!rows.length && !lastPairs.length){
      showNoResultsState();
    } else {
      paint();
    }
    refreshSummary();
    renderAppliedBar();
  }

  function paint(){
    body.innerHTML = '';

    if (lastResult && lastResult.truncated){
      var info = div('dup-info');
      var shown = asNum(lastResult.shown, 0);
      var total = asNum(lastResult.total, 0);
      info.innerHTML = '<div class="t">표시 제한</div><div class="s">결과가 많아 상위 ' + shown + '건만 표시합니다. 전체(' + total + '건)는 엑셀 내보내기에서 확인하세요.</div>';
      body.appendChild(info);
    }

    if (mode === 'clash' && lastPairs && lastPairs.length){
      paintPairs(lastPairs);
      return;
    }

    paintGroups(groups);
    paintRowStates();
  }

  function collectGroupPairIds(items){
    var ids = [];
    for (var i=0;i<(items||[]).length;i++){
      var p = items[i] || {};
      ids.push(p.aId);
      ids.push(p.bId);
    }
    return uniqNums(ids);
  }

  function collectPairIds(p){
    p = p || {};
    return uniqNums([p.aId, p.bId]);
  }

  function paintPairs(pairs){
    var byGroup = {};
    for (var i=0;i<pairs.length;i++){
      var p = pairs[i] || {};
      var gk = asStr(p.groupKey, 'C0000');
      if (!byGroup[gk]) byGroup[gk] = [];
      byGroup[gk].push(p);
    }

    var keys = Object.keys(byGroup);
    keys.sort(function(a,b){ return byGroup[b].length - byGroup[a].length; });

    for (var gi=0; gi<keys.length; gi++){
      var gk2 = keys[gi];
      var items = byGroup[gk2];

      var card = div('dup-grp');
      var h = div('grp-h');

      var left = div('grp-txt');
      left.innerHTML =
        '<div class="grp-title"><span class="grp-badge">간섭 그룹 ' + (gi + 1) + '</span>' +
        '<span class="grp-meta">' + esc(gk2) + '</span></div>' +
        '<div class="grp-count">' + items.length + '쌍</div>';

      var right = div('grp-actions');
      var grpIds = collectGroupPairIds(items);
      if (grpIds.length){
        var grpSel = makeBtn('그룹 선택', 'control-chip chip-btn subtle');
        grpSel.dataset.act = 'sel-group';
        grpSel.dataset.ids = grpIds.join(',');
        right.appendChild(grpSel);
      }

      h.appendChild(left);
      h.appendChild(right);
      card.appendChild(h);

      var list = div('pair-cardlist');

      var seen = {};
      for (var i2=0;i2<items.length;i2++){
        var p2 = items[i2] || {};
        var aId = safeId(p2.aId);
        var bId = safeId(p2.bId);
        var na = asNum(aId, NaN), nb = asNum(bId, NaN);
        var x = (isFinite(na) && !isNaN(na) && isFinite(nb) && !isNaN(nb)) ? Math.min(na, nb) : aId;
        var y = (isFinite(na) && !isNaN(na) && isFinite(nb) && !isNaN(nb)) ? Math.max(na, nb) : bId;
        var k = gk2 + ':' + String(x) + '-' + String(y);
        if (seen[k]) continue;
        seen[k] = 1;

        var aInfo = (asStr(p2.aCategory,'') + ' · ' + asStr(p2.aFamily,'') + (p2.aType ? ' : ' + asStr(p2.aType,'') : '')).trim();
        var bInfo = (asStr(p2.bCategory,'') + ' · ' + asStr(p2.bFamily,'') + (p2.bType ? ' : ' + asStr(p2.bType,'') : '')).trim();

        var pc = div('pair-card');
        var pairIds = collectPairIds(p2);
        pc.innerHTML =
          '<div class="pair-main">' +
            '<div class="pair-side"><div class="pair-lbl">A</div>' +
              '<div class="pair-id"><button class="table-action-btn" data-act="sel" data-id="' + esc(aId) + '">' + esc(aId) + '</button></div>' +
              '<div class="pair-meta">' + esc(aInfo || '—') + '</div></div>' +
            '<div class="pair-mid">↔</div>' +
            '<div class="pair-side"><div class="pair-lbl">B</div>' +
              '<div class="pair-id"><button class="table-action-btn" data-act="sel" data-id="' + esc(bId) + '">' + esc(bId) + '</button></div>' +
              '<div class="pair-meta">' + esc(bInfo || '—') + '</div></div>' +
          '</div>';

        var pairComment = asStr(p2.comment || p2.note || p2.reason || '', '');
        if (pairComment){
          var pairNote = div('pair-note');
          pairNote.textContent = pairComment;
          pc.appendChild(pairNote);
        }

        if (pairIds.length){
          var pairActions = div('pair-actions');
          var pairSel = makeBtn('페어 선택', 'table-action-btn');
          pairSel.dataset.act = 'sel-pair';
          pairSel.dataset.ids = pairIds.join(',');
          pairActions.appendChild(pairSel);
          pc.appendChild(pairActions);
        }

        list.appendChild(pc);
      }

      card.appendChild(list);
      body.appendChild(card);
    }
  }

  function paintGroups(gs){
    for (var i=0;i<gs.length;i++){
      var g = gs[i];

      var card = div('dup-grp');
      var h = div('grp-h');

      var left = div('grp-txt');
      left.innerHTML =
        '<div class="grp-title"><span class="grp-badge">' + (mode === 'clash' ? '간섭 그룹' : '중복 그룹') + ' ' + (i + 1) + '</span>' +
        '<span class="grp-meta">' + esc(buildGroupMeta(g)) + '</span></div>' +
        '<div class="grp-count">' + g.rows.length + '개</div>';

      var right = div('grp-actions');
      var tg = makeBtn(expanded[g.key] ? '접기' : '펼치기', 'control-chip chip-btn subtle');
      tg.addEventListener('click', (function(key){
        return function(){
          if (expanded[key]) delete expanded[key];
          else expanded[key] = 1;
          paint();
          refreshSummary();
        };
      })(g.key));
      right.appendChild(tg);

      h.appendChild(left);
      h.appendChild(right);

      card.appendChild(h);

      var tbl = div('grp-body');

      var sh = div('dup-subhead');
      sh.appendChild(cell('', 'ck'));
      sh.appendChild(cell('Element ID', 'th'));
      sh.appendChild(cell('Category', 'th'));
      sh.appendChild(cell('Family', 'th'));
      sh.appendChild(cell('Type', 'th'));
      sh.appendChild(cell('연결', 'th conn'));
      sh.appendChild(cell('작업', 'th right'));
      tbl.appendChild(sh);

      if (expanded[g.key]){
        for (var r=0;r<g.rows.length;r++){
          tbl.appendChild(renderRow(g.rows[r]));
        }
      }

      card.appendChild(tbl);
      body.appendChild(card);
    }
  }

  function renderRow(r){
    var row = div('dup-row');
    row.dataset.id = r.id;

    var ckCell = cell('', 'ck');
    var ck = document.createElement('input');
    ck.type = 'checkbox';
    ck.className = 'ckbox';
    ckCell.textContent = '';
    ckCell.appendChild(ck);

    row.appendChild(ckCell);
    row.appendChild(cell(co(r.id,'-'), 'td mono right'));
    row.appendChild(cell(co(r.category,'—'), 'td'));
    row.appendChild(cell(co(r.family, (r.category ? (r.category + ' Type') : '—')), 'td ell'));
    row.appendChild(cell(co(r.type,'—'), 'td ell'));

    var conn = document.createElement('div');
    conn.className = 'conn-cell';
    var ids = Array.isArray(r.connectedIds) ? r.connectedIds : [];
    conn.textContent = ids.length ? String(ids.length) : '0';
    if (ids.length) conn.title = ids.join(', ');
    row.appendChild(cell(conn, 'td mono right conn'));

    var act = div('row-actions');
    var zoom = makeBtn('선택/줌', 'table-action-btn');
    zoom.dataset.act = 'zoom';
    zoom.dataset.id = r.id;

    var del = makeBtn(deleted[r.id] ? '되돌리기' : '삭제', 'table-action-btn ' + (deleted[r.id] ? 'restore' : 'table-action-btn--danger'));
    del.dataset.act = deleted[r.id] ? 'res' : 'del';
    del.dataset.id = r.id;

    act.appendChild(zoom);
    act.appendChild(del);
    row.appendChild(cell(act, 'td right'));

    row.classList.toggle('is-deleted', !!deleted[r.id]);

    return row;
  }

  function paintRowStates(){
    var els = body.querySelectorAll('.dup-row');
    for (var i=0;i<els.length;i++){
      var el = els[i];
      var id = el.dataset ? el.dataset.id : '';
      var isDel = !!deleted[id];
      el.classList.toggle('is-deleted', isDel);

      var btn = el.querySelector('[data-act="del"], [data-act="res"]');
      if (btn){
        btn.textContent = isDel ? '되돌리기' : '삭제';
        btn.dataset.act = isDel ? 'res' : 'del';
        btn.classList.toggle('restore', isDel);
        btn.classList.toggle('table-action-btn--danger', !isDel);
      }
    }
  }

  function refreshSummary(){
    summaryBar.innerHTML = '';
    var gCount = groups.length;
    var eCount = 0;
    for (var i=0;i<groups.length;i++) eCount += groups[i].rows.length;

    var visible = (busy || eCount > 0 || (mode === 'clash' && lastPairs.length > 0));
    summaryBar.classList.toggle('hidden', !visible);

    var gLabel = (mode === 'clash') ? '간섭 그룹' : '중복 그룹';
    summaryBar.appendChild(chip(gLabel + ' ' + gCount));
    summaryBar.appendChild(chip('요소 ' + eCount));
    if (mode === 'clash'){
      summaryBar.appendChild(chip('쌍 ' + (lastPairs ? lastPairs.length : 0)));
    }

  
  }

  function showDupCompletionDialog(){
    if (!lastResult || typeof lastResult !== 'object') return;

    var modeKey = String(asStr(lastResult.mode, mode) || mode || 'duplicate').toLowerCase();
    var isClashMode = modeKey === 'clash';
    var scanCount = asNum(lastResult.scan, 0);
    var groupCount = asNum(lastResult.groups, groups.length);
    var candidateCount = asNum(lastResult.candidates, isClashMode ? lastPairs.length : rows.length);
    var shownCount = asNum(lastResult.shown, isClashMode ? lastPairs.length : rows.length);
    var totalCount = asNum(lastResult.total, isClashMode ? lastPairs.length : rows.length);
    var signature = [
      modeKey,
      scanCount,
      groupCount,
      candidateCount,
      shownCount,
      totalCount,
      lastResult.truncated ? '1' : '0'
    ].join('|');
    if (signature === lastCompletionSignature) return;
    lastCompletionSignature = signature;

    var notes = [];
    if (lastResult.truncated){
      notes.push('결과가 많아 화면에는 상위 ' + shownCount + '건만 표시했습니다. 전체 ' + totalCount + '건은 엑셀 내보내기에서 확인할 수 있습니다.');
    }

    showCompletionSummaryDialog({
      title: isClashMode ? '자체간섭 검토 완료' : '중복 검토 완료',
      message: '검토 결과를 요약했습니다. 필요하면 바로 엑셀로 내보내세요.',
      summaryItems: [
        { label: '검사 대상 수', value: String(scanCount) },
        { label: '그룹 수', value: String(groupCount) },
        { label: isClashMode ? '간섭 후보 수' : '삭제 후보 수', value: String(candidateCount) },
        { label: '표시 건수 / 전체 건수', value: shownCount + ' / ' + totalCount }
      ],
      notes: notes,
      exportDisabled: !!exportBtn.disabled,
      onExport: function(){
        exportBtn.click();
      }
    });
  }


  function buildGroupMeta(g){
    var cats = uniq((g.rows||[]).map(function(r){ return r.category || '—'; }));
    var fams = uniq((g.rows||[]).map(function(r){ return r.family || (r.category ? (r.category + ' Type') : '—'); }));
    var types = uniq((g.rows||[]).map(function(r){ return r.type || '—'; }));

    var catOut = (cats.length === 1) ? cats[0] : ('혼합(' + cats.length + ')');
    var famOut = (fams.length === 1) ? fams[0] : ('혼합(' + fams.length + ')');
    var typOut = (types.length === 1) ? types[0] : ('혼합(' + types.length + ')');
    return catOut + ' · ' + famOut + ' · ' + typOut;
  }

  function normalizeRow(r){
    var id = safeId(co(r.elementId, co(r.ElementId, co(r.id, r.Id))));
    var category = asStr(co(r.category, r.Category), '');
    var family = asStr(co(r.family, r.Family), '');
    var type = asStr(co(r.type, r.Type), '');
    var groupKey = asStr(co(r.groupKey, r.GroupKey), '');
    var connRaw = co(r.connectedIds, co(r.ConnectedIds, []));
    var connectedIds = Array.isArray(connRaw) ? connRaw.map(String) : [];
    return { id: id, category: category, family: family, type: type, groupKey: groupKey, connectedIds: connectedIds };
  }

  function buildGroups(rs){
    var list = Array.isArray(rs) ? rs : [];
    var hasKey = false;
    for (var i=0;i<list.length;i++){ if (list[i] && list[i].groupKey){ hasKey = true; break; } }
    if (hasKey){
      var map = {};
      for (var j=0;j<list.length;j++){
        var r = list[j];
        var k = r.groupKey || '_';
        if (!map[k]) map[k] = { key: k, rows: [] };
        map[k].rows.push(r);
      }
      return Object.keys(map).map(function(k){ return map[k]; });
    }
    var map2 = {};
    for (var k2=0;k2<list.length;k2++){
      var rr = list[k2];
      var sig = [String(rr.id)].concat((rr.connectedIds||[]).map(String)).filter(Boolean).sort().join(',');
      var key = [rr.category||'', rr.family||'', rr.type||'', sig].join('|');
      if (!map2[key]) map2[key] = { key: key, rows: [] };
      map2[key].rows.push(rr);
    }
    return Object.keys(map2).map(function(k){ return map2[k]; });
  }

  function dedupPairs(pairs){
    var out = [];
    var seen = {};
    for (var i=0;i<(pairs||[]).length;i++){
      var p = pairs[i] || {};
      var gk = asStr(p.groupKey,'');
      var aId = safeId(p.aId);
      var bId = safeId(p.bId);
      var na = asNum(aId, NaN), nb = asNum(bId, NaN);
      var x = (isFinite(na) && !isNaN(na) && isFinite(nb) && !isNaN(nb)) ? Math.min(na,nb) : aId;
      var y = (isFinite(na) && !isNaN(na) && isFinite(nb) && !isNaN(nb)) ? Math.max(na,nb) : bId;
      var key = gk + ':' + String(x) + '-' + String(y);
      if (seen[key]) continue;
      seen[key] = 1;
      out.push(p);
    }
    return out;
  }

  function renderIntro(){
    body.innerHTML = '';
    var hero = div('dup-hero');
    hero.innerHTML = '<h3 class="hero-title">' + (mode === 'clash' ? '자체간섭 검토를 시작해 보세요' : '중복검토를 시작해 보세요') + '</h3>' +
                     '<p class="hero-sub">' + (mode === 'clash' ? '같은 파일 내 자체간섭 후보를 쌍(A↔B)으로 보여줍니다.' : '모델의 중복 요소를 그룹으로 묶어 보여줍니다.') + '</p>';
    body.appendChild(hero);
  }

  function getExcludeKeywords(){
    try{
      var raw = String(localStorage.getItem(LS_EXCL_KW) || '').trim();
      if (!raw) return [];
      return raw.split(',').map(function(s){ return s.trim(); }).filter(Boolean);
    } catch(e){ return []; }
  }

  function readTolMm(){
    try{
      var raw = localStorage.getItem(LS_TOL_MM);
      var n = Number(String(raw || '').trim());
      if (isFinite(n) && !isNaN(n) && n > 0) return clamp(n, 0.01, 1000);
    } catch(e){}
    return DEFAULT_TOL_MM;
  }
  function readTolUnit(){
    try{
      return normalizeTolUnit(localStorage.getItem(LS_TOL_UNIT));
    } catch(e){}
    return 'mm';
  }
  function getTolDisplayValue(mm, unit){
    return normalizeTolUnit(unit) === 'inch' ? mmToInch(mm) : mm;
  }
  function getTolMmFromDisplay(value, unit){
    var n = asNum(value, DEFAULT_TOL_MM);
    var mm = normalizeTolUnit(unit) === 'inch' ? inchToMm(n) : n;
    return clamp(mm, 0.01, 1000);
  }
  function readTolDisplayValue(unit){
    return getTolDisplayValue(readTolMm(), unit || readTolUnit());
  }
  function getTolDisplayMeta(unit){
    var u = normalizeTolUnit(unit || readTolUnit());
    if (u === 'inch') return { label:'inch', min:'0.0004', step:'0.0001' };
    return { label:'mm', min:'0.01', step:'0.1' };
  }
  function formatTolDisplayText(mm, unit){
    var meta = getTolDisplayMeta(unit);
    return formatTolNumber(getTolDisplayValue(mm, unit)) + ' ' + meta.label;
  }
  function readTolFeet(){
    var mm = readTolMm();
    return Math.max(0.000001, mmToFeet(mm));
  }

  function readMode(){
    try{
      var m = String(localStorage.getItem(LS_MODE) || '').trim();
      if (m === 'duplicate' || m === 'clash') return m;
    } catch(e){}
    return 'duplicate';
  }
  function readScopeMode(){
    try{
      var m = String(localStorage.getItem(LS_SCOPE) || '').trim();
      if (m === 'all' || m === 'scope' || m === 'exclude') return m;
    } catch(e){}
    return 'all';
  }

  function handleExcelProgress(p){
    if (!p){ ProgressDialog.hide(); return; }
    var phase = asStr(p.phase,'').toUpperCase();
    var total = asNum(p.total,0);
    var current = asNum(p.current,0);
    var pct = total > 0 ? Math.floor((current/total)*100) : asNum(p.percent,0);
    ProgressDialog.show('엑셀 내보내기', asStr(p.message,''));
    ProgressDialog.update(clamp(pct,0,100), asStr(p.message,''), asStr(p.detail,''));
    if (phase === 'DONE' || phase === 'ERROR'){
      setTimeout(function(){ ProgressDialog.hide(); }, 200);
    }
  }
  function renderAppliedBar(){
    if (!appliedBar) return;
    var cfg = loadRuleConfig();
    var merged = uniq([].concat(getExcludeKeywords(), cfg.excludeFamilies || [], cfg.excludeCategories || []));
    var setCount = (cfg.sets && cfg.sets.length) ? cfg.sets.length : 0;
    var pairCount = (cfg.pairs && cfg.pairs.length) ? cfg.pairs.length : 0;
    var pairEnabled = 0;
    if (cfg.pairs && cfg.pairs.length){
      for (var i=0;i<cfg.pairs.length;i++) if (cfg.pairs[i] && cfg.pairs[i].enabled !== false) pairEnabled++;
    }
    var exSetCount = (cfg.excludeSetIds && cfg.excludeSetIds.length) ? cfg.excludeSetIds.length : 0;
    var exFamCount = (cfg.excludeFamilies && cfg.excludeFamilies.length) ? cfg.excludeFamilies.length : 0;
    var exCatCount = (cfg.excludeCategories && cfg.excludeCategories.length) ? cfg.excludeCategories.length : 0;
    var exportParams = getExportParamNames();
    var scopeTxt = (readScopeMode() === 'scope') ? '선택만' : (readScopeMode() === 'exclude') ? '선택 제외' : '전체';
    var modeTxt = (mode === 'clash') ? '자체간섭' : '중복검토';
    var selectedFile = getSelectedRuleLibraryItem();

    appliedBar.classList.remove('hidden');
    appliedBar.innerHTML = '';

    var label = document.createElement('div');
    label.className = 'ap-label';
    label.textContent = '현재 적용 규칙';
    appliedBar.appendChild(label);

    var wrap = document.createElement('div');
    wrap.className = 'ap-chips';
    function addChip(txt){
      var c = document.createElement('span');
      c.className = 'ap-chip';
      c.textContent = txt;
      wrap.appendChild(c);
    }
    addChip('모드 ' + modeTxt);
    addChip('허용오차 ' + formatTolDisplayText(readTolMm(), readTolUnit()));
    addChip('범위 ' + scopeTxt);
    addChip('Set ' + setCount);
    addChip('Pair ' + pairEnabled + '/' + pairCount);
    addChip('Exclude Set ' + exSetCount);
    addChip('Exclude 목록 ' + (exFamCount + exCatCount));
    addChip('키워드 ' + merged.length);
    addChip('속성 추출 ' + exportParams.length);
    appliedBar.appendChild(wrap);

    var lines = [];
    if (selectedFile && selectedFile.name) lines.push('불러온 설정: ' + selectedFile.name);
    if (cfg.pairs && cfg.pairs.length){
      var names = [];
      for (var pi=0; pi<cfg.pairs.length; pi++){
        var p = cfg.pairs[pi];
        if (!p || p.enabled === false) continue;
        names.push((p.a || '__') + ' vs ' + (p.b || '__'));
        if (names.length >= 2) break;
      }
      if (names.length) lines.push('활성 Pair: ' + names.join(' · ') + (pairEnabled > 2 ? ' …' : ''));
    }
    if (exportParams.length){
      lines.push('추출 파라미터: ' + exportParams.slice(0, 3).join(', ') + (exportParams.length > 3 ? ' ...' : ''));
    }
    if (lines.length){
      var sub = document.createElement('div');
      sub.className = 'ap-sub';
      sub.textContent = lines.join('  |  ');
      appliedBar.appendChild(sub);
    }
  }

  function buildRulePanel(){
    var wrap = div('dup-rulemodal');
    wrap.innerHTML =
      '<div class="rm-backdrop" data-act="close"></div>' +
      '<div class="rm-window">' +
        '<div class="rm-head">' +
          '<div class="rm-head-left">' +
            '<div class="rm-title">Rule/Set 설정</div>' +
            '<div class="rm-sub">검토 규칙, Exclude Set, 제외 목록, 허용오차를 한곳에서 관리합니다.</div>' +
          '</div>' +
          '<div class="rm-actions">' +
            '<button class="rp-btn rp-btn--ghost" data-act="export">XML 저장</button>' +
            '<button class="rp-btn rp-btn--ghost" data-act="import">XML 불러오기</button>' +
            '<button class="rp-btn rp-btn--primary" data-act="apply">적용</button>' +
            '<button class="rp-btn" data-act="close">닫기</button>' +
          '</div>' +
        '</div>' +
        '<div class="rm-body">' +
          '<section class="rp-overview">' +
            '<div class="rp-overview-title">설정 안내</div>' +
            '<div class="rp-overview-copy">중복검토와 자체간섭 검토는 상단 모드 전환으로 바꾸고, 여기서는 검토 규칙과 제외 대상을 관리합니다.</div>' +
          '</section>' +

          '<section class="rp-sec rp-sec--compact">' +
            '<div class="rp-sec-head">' +
              '<div><div class="rp-sec-title">공통 설정</div><div class="rp-sec-sub">허용오차, 선택 범위, 제외 키워드를 설정합니다. 목록은 창을 열 때마다 자동으로 최신 상태를 가져옵니다.</div></div>' +
            '</div>' +
            '<div class="rp-grid">' +
              '<div class="rp-label">허용오차</div><div class="rp-inline-field"><input class="rp-input" type="number" step="0.1" min="0.01" data-bind="tolValue"/><select class="rp-select rp-select--unit" data-bind="tolUnit"><option value="mm">mm</option><option value="inch">inch</option></select></div>' +
              '<div class="rp-label">범위(Selection)</div>' +
              '<select class="rp-select" data-bind="scopeMode"><option value="all">전체</option><option value="scope">선택한 요소만 검사</option><option value="exclude">선택한 요소는 제외</option></select>' +
              '<div class="rp-label">제외 키워드</div><input class="rp-input" type="text" placeholder="예: Dummy, Temp (콤마 구분)" data-bind="excludeKeywords"/>' +
              '<div class="rp-label">속성 추출</div><input class="rp-input" type="text" placeholder="예: Mark, Comments (콤마 구분)" data-bind="exportParamNames"/>' +
            '</div>' +
            '<div class="rp-hint">이 창을 열면 프로젝트에 실제로 존재하는 파라미터/패밀리/시스템 목록을 자동으로 최신 상태로 가져옵니다.</div>' +
          '</section>' +

          '<section class="rp-sec">' +
            '<div class="rp-sec-head">' +
              '<div><div class="rp-sec-title">Set 정의</div><div class="rp-sec-sub">Set은 그룹 OR · 조건 AND 구조입니다.</div></div>' +
              '<div class="rp-section-actions"><button class="rp-btn rp-btn--add" data-act="add-set">+ Set 추가</button></div>' +
            '</div>' +
            '<div data-slot="sets"></div>' +
          '</section>' +

          '<section class="rp-sec">' +
            '<div class="rp-sec-head">' +
              '<div><div class="rp-sec-title">Pair (Set vs Set)</div><div class="rp-sec-sub">A와 B에 속한 요소 사이의 자체간섭만 계산합니다.</div></div>' +
              '<div class="rp-section-actions"><button class="rp-btn rp-btn--add" data-act="add-pair">+ Pair 추가</button></div>' +
            '</div>' +
            '<div data-slot="pairs"></div>' +
          '</section>' +

          '<section class="rp-sec">' +
            '<div class="rp-sec-head">' +
              '<div><div class="rp-sec-title">Exclude Sets</div><div class="rp-sec-sub">여기에 등록한 Set은 모든 간섭 검토에서 제외됩니다.</div></div>' +
              '<div class="rp-section-actions"><button class="rp-btn rp-btn--add" data-act="add-exset">+ Exclude Set 추가</button></div>' +
            '</div>' +
            '<div class="rp-inline-note">Set을 먼저 만들어야 Exclude Set으로 등록할 수 있습니다.</div>' +
            '<div data-slot="exsets"></div>' +
          '</section>' +

          '<section class="rp-sec">' +
            '<div class="rp-sec-head">' +
              '<div><div class="rp-sec-title">Exclude 목록</div><div class="rp-sec-sub">패밀리/시스템 목록에서 직접 제외 대상을 선택합니다.</div></div>' +
              '<div class="rp-section-actions">' +
                '<button class="rp-btn rp-btn--accent" data-act="open-expicker">패밀리/시스템 목록 열기</button>' +
              '</div>' +
            '</div>' +
            '<div class="rp-exsummary" data-slot="exsum">—</div>' +
          '</section>' +

          '<section class="rp-sec rp-sec--compact">' +
            '<div class="rp-sec-head">' +
              '<div><div class="rp-sec-title">설정 파일</div><div class="rp-sec-sub">XML 저장은 저장 경로와 파일 이름을 직접 지정합니다. 불러온 XML은 파일 이름으로 목록에 누적되고, 선택 적용/삭제/비우기가 가능합니다.</div></div>' +
              '<div class="rp-section-actions">' +
                '<button class="rp-btn rp-btn--ghost" data-act="apply-lib">선택 적용</button>' +
                '<button class="rp-btn rp-btn--ghost" data-act="del-lib">선택 삭제</button>' +
                '<button class="rp-btn rp-btn--ghost" data-act="clear-lib">목록 비우기</button>' +
              '</div>' +
            '</div>' +
            '<div class="rp-filebox">' +
              '<div class="rp-filecopy">XML 불러오기 후 파일 이름 기준으로 목록이 생성됩니다.</div>' +
              '<div class="rp-filehint">목록에서 파일을 선택한 뒤 “선택 적용”으로 현재 설정에 반영하세요.</div>' +
              '<div data-slot="cfglist"></div>' +
            '</div>' +
            '<input type="file" data-bind="cfgfile" accept=".xml,.json,text/xml,application/xml,application/json" style="display:none" />' +
          '</section>' +
        '</div>' +
      '</div>';

    wrap.addEventListener('click', function(e){
      var t = e.target;
      var act = t && t.dataset ? t.dataset.act : '';
      if (!act && t && t.classList && t.classList.contains('rm-backdrop')){ closeRulePanel(); return; }

      if (act === 'close'){ closeRulePanel(); return; }
      if (act === 'open-expicker'){ openExcludePicker(); return; }
      if (act === 'export'){ exportRuleJson(); return; }
      if (act === 'import'){
        var fileEl = wrap.querySelector('[data-bind="cfgfile"]');
        if (fileEl) fileEl.click();
        return;
      }
      if (act === 'apply'){ applyRulePanel(); return; }
      if (act === 'apply-lib'){ applySelectedRuleLibrary(); return; }
      if (act === 'del-lib'){ deleteSelectedRuleLibrary(); return; }
      if (act === 'clear-lib'){ clearSelectedRuleLibrary(); return; }

      if (act === 'add-set'){ addSet(); renderRulePanel(); return; }
      if (act === 'add-group'){
        var si = asNum(t.dataset.si, -1);
        if (si >= 0){ addGroup(si); renderRulePanel(); }
        return;
      }
      if (act === 'add-clause'){
        var si2 = asNum(t.dataset.si, -1);
        var gi2 = asNum(t.dataset.gi, -1);
        if (si2 >= 0 && gi2 >= 0){ addClause(si2, gi2); renderRulePanel(); }
        return;
      }
      if (act === 'del-set'){
        var si3 = asNum(t.dataset.si, -1);
        if (si3 >= 0){ delSet(si3); renderRulePanel(); }
        return;
      }
      if (act === 'del-group'){
        var si4 = asNum(t.dataset.si, -1);
        var gi4 = asNum(t.dataset.gi, -1);
        if (si4 >= 0 && gi4 >= 0){ delGroup(si4, gi4); renderRulePanel(); }
        return;
      }
      if (act === 'del-clause'){
        var si5 = asNum(t.dataset.si, -1);
        var gi5 = asNum(t.dataset.gi, -1);
        var ci5 = asNum(t.dataset.ci, -1);
        if (si5 >= 0 && gi5 >= 0 && ci5 >= 0){ delClause(si5, gi5, ci5); renderRulePanel(); }
        return;
      }

      if (act === 'add-pair'){ addPair(); renderRulePanel(); return; }
      if (act === 'del-pair'){
        var pi = asNum(t.dataset.pi, -1);
        if (pi >= 0){ delPair(pi); renderRulePanel(); }
        return;
      }

      if (act === 'add-exset'){
        var c0 = loadRuleConfig();
        if (!c0.sets || !c0.sets.length){ toast('Exclude Sets를 추가하려면 먼저 Set을 만들어 주세요.', 'warn', 1800); return; }
        addExcludeSet(); renderRulePanel(); return;
      }
      if (act === 'del-exset'){
        var ei = asNum(t.dataset.ei, -1);
        if (ei >= 0){ delExcludeSet(ei); renderRulePanel(); }
        return;
      }
    });

    wrap.addEventListener('change', function(e){
      var t = e.target;
      if (!t || !t.dataset) return;

      if (t.dataset.bind === 'cfgfile'){
        handleRuleFileImport(t);
        return;
      }
      if (t.dataset.libSel){
        setSelectedRuleLibraryId(asStr(t.value, ''));
        renderRulePanel();
        return;
      }

      if (t.dataset.bind === 'scopeMode'){
        scopeMode = asStr(t.value,'all');
        try{ localStorage.setItem(LS_SCOPE, scopeMode); } catch(e2){}
        return;
      }
      if (t.dataset.bind === 'tolUnit'){
        var nextUnit = normalizeTolUnit(t.value);
        try{ localStorage.setItem(LS_TOL_UNIT, nextUnit); } catch(e3){}
        renderRulePanel();
        renderAppliedBar();
        return;
      }
      if (t.dataset.bind === 'tolValue'){
        var mm = getTolMmFromDisplay(t.value, readTolUnit());
        try{ localStorage.setItem(LS_TOL_MM, String(mm)); } catch(e4){}
        renderAppliedBar();
        return;
      }
      if (t.dataset.bind === 'excludeKeywords'){
        var kw = asStr(t.value,'').trim();
        try{ localStorage.setItem(LS_EXCL_KW, kw); } catch(e5){}
        return;
      }
      if (t.dataset.bind === 'exportParamNames'){
        var exportParamText = asStr(t.value,'').trim();
        try{ localStorage.setItem(LS_EXPORT_PARAMS, exportParamText); } catch(e6){}
        renderAppliedBar();
        return;
      }

      if (t.dataset.setName){
        var si = asNum(t.dataset.setName, -1);
        if (si >= 0){
          var cfg = loadRuleConfig();
          if (cfg.sets[si]) cfg.sets[si].name = asStr(t.value, cfg.sets[si].name);
          saveRuleConfig(cfg);
        }
        return;
      }
      if (t.dataset.setLogic){
        var si2 = asNum(t.dataset.setLogic, -1);
        if (si2 >= 0){
          var cfg2 = loadRuleConfig();
          if (cfg2.sets[si2]) cfg2.sets[si2].logic = (t.value === 'and' ? 'and' : 'or');
          saveRuleConfig(cfg2);
          renderRulePanel();
        }
        return;
      }

      if (t.dataset.clField){
        var parts = String(t.dataset.clField).split(':');
        var si3 = asNum(parts[0],-1), gi3 = asNum(parts[1],-1), ci3 = asNum(parts[2],-1);
        if (si3>=0 && gi3>=0 && ci3>=0){
          var cfg3 = loadRuleConfig();
          var cl = cfg3.sets[si3] && cfg3.sets[si3].groups[gi3] && cfg3.sets[si3].groups[gi3].clauses[ci3];
          if (cl){
            cl.field = String(t.value || 'category');
            saveRuleConfig(cfg3);
            renderRulePanel();
          }
        }
        return;
      }
      if (t.dataset.clOp){
        var parts2 = String(t.dataset.clOp).split(':');
        var si4 = asNum(parts2[0],-1), gi4 = asNum(parts2[1],-1), ci4 = asNum(parts2[2],-1);
        if (si4>=0 && gi4>=0 && ci4>=0){
          var cfg4 = loadRuleConfig();
          var cl2 = cfg4.sets[si4] && cfg4.sets[si4].groups[gi4] && cfg4.sets[si4].groups[gi4].clauses[ci4];
          if (cl2){ cl2.op = String(t.value || 'contains'); saveRuleConfig(cfg4); }
        }
        return;
      }
      if (t.dataset.clParam){
        var parts3 = String(t.dataset.clParam).split(':');
        var si5 = asNum(parts3[0],-1), gi5 = asNum(parts3[1],-1), ci5 = asNum(parts3[2],-1);
        if (si5>=0 && gi5>=0 && ci5>=0){
          var cfg5 = loadRuleConfig();
          var cl3 = cfg5.sets[si5] && cfg5.sets[si5].groups[gi5] && cfg5.sets[si5].groups[gi5].clauses[ci5];
          if (cl3){ cl3.param = String(t.value || ''); saveRuleConfig(cfg5); }
        }
        return;
      }
      if (t.dataset.clVal){
        var parts4 = String(t.dataset.clVal).split(':');
        var si6 = asNum(parts4[0],-1), gi6 = asNum(parts4[1],-1), ci6 = asNum(parts4[2],-1);
        if (si6>=0 && gi6>=0 && ci6>=0){
          var cfg6 = loadRuleConfig();
          var cl4 = cfg6.sets[si6] && cfg6.sets[si6].groups[gi6] && cfg6.sets[si6].groups[gi6].clauses[ci6];
          if (cl4){ cl4.value = String(t.value || ''); saveRuleConfig(cfg6); }
        }
        return;
      }

      if (t.dataset.pA){
        var piA = asNum(t.dataset.pA, -1);
        if (piA>=0){
          var cfg7 = loadRuleConfig();
          if (cfg7.pairs[piA]) cfg7.pairs[piA].a = String(t.value || '');
          saveRuleConfig(cfg7);
        }
        return;
      }
      if (t.dataset.pB){
        var piB = asNum(t.dataset.pB, -1);
        if (piB>=0){
          var cfg8 = loadRuleConfig();
          if (cfg8.pairs[piB]) cfg8.pairs[piB].b = String(t.value || '');
          saveRuleConfig(cfg8);
        }
        return;
      }
      if (t.dataset.pEn){
        var piE = asNum(t.dataset.pEn, -1);
        if (piE>=0){
          var cfg9 = loadRuleConfig();
          if (cfg9.pairs[piE]) cfg9.pairs[piE].enabled = !!t.checked;
          saveRuleConfig(cfg9);
        }
        return;
      }

      if (t.dataset.exSet){
        var eiS = asNum(t.dataset.exSet, -1);
        if (eiS>=0){
          var cfg10 = loadRuleConfig();
          cfg10.excludeSetIds[eiS] = String(t.value || '');
          saveRuleConfig(cfg10);
        }
        return;
      }
    });

    wrap.addEventListener('input', function(e){
      var t = e.target;
      if (!t || !t.dataset) return;
      if (t.dataset.bind === 'excludeKeywords'){
        var kw = asStr(t.value,'').trim();
        try{ localStorage.setItem(LS_EXCL_KW, kw); } catch(e2){}
      }
      if (t.dataset.bind === 'exportParamNames'){
        var exportParamText2 = asStr(t.value,'').trim();
        try{ localStorage.setItem(LS_EXPORT_PARAMS, exportParamText2); } catch(e3){}
        renderAppliedBar();
      }
      if (t.dataset.clVal){
        var parts = String(t.dataset.clVal).split(':');
        var si = asNum(parts[0],-1), gi=asNum(parts[1],-1), ci=asNum(parts[2],-1);
        if (si>=0 && gi>=0 && ci>=0){
          var cfg = loadRuleConfig();
          var cl = cfg.sets[si] && cfg.sets[si].groups[gi] && cfg.sets[si].groups[gi].clauses[ci];
          if (cl){ cl.value = String(t.value || ''); saveRuleConfig(cfg); }
        }
      }
      if (t.dataset.setName){
        var si2 = asNum(t.dataset.setName, -1);
        if (si2>=0){
          var cfg2 = loadRuleConfig();
          if (cfg2.sets[si2]) cfg2.sets[si2].name = String(t.value || '');
          saveRuleConfig(cfg2);
        }
      }
    });

    return wrap;
  }

  function xmlEsc(s){
    return String(co(s,''))
      .replace(/&/g,'&amp;')
      .replace(/</g,'&lt;')
      .replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;')
      .replace(/'/g,'&apos;');
  }
  function serializeRuleFile(){
    var cfg = loadRuleConfig();
    var lines = [];
    lines.push('<?xml version="1.0" encoding="utf-8"?>');
    lines.push('<DuplicateInspectorRuleSet version="2">');
    lines.push('  <Settings tolMm="' + xmlEsc(String(readTolMm())) + '" tolUnit="' + xmlEsc(readTolUnit()) + '" scopeMode="' + xmlEsc(readScopeMode()) + '" excludeKeywords="' + xmlEsc(getExcludeKeywords().join(', ')) + '" exportParams="' + xmlEsc(getExportParamNames().join(', ')) + '" />');
    lines.push('  <Sets>');
    for (var i=0;i<(cfg.sets||[]).length;i++){
      var s = cfg.sets[i] || {};
      lines.push('    <Set id="' + xmlEsc(asStr(s.id,'')) + '" name="' + xmlEsc(asStr(s.name,'')) + '" logic="' + xmlEsc(asStr(s.logic,'or')) + '">');
      var groups = Array.isArray(s.groups) ? s.groups : [];
      for (var gi=0; gi<groups.length; gi++){
        var g = groups[gi] || {};
        lines.push('      <Group>');
        var clauses = Array.isArray(g.clauses) ? g.clauses : [];
        for (var ci=0; ci<clauses.length; ci++){
          var cl = clauses[ci] || {};
          lines.push('        <Clause field="' + xmlEsc(asStr(cl.field,'category')) + '" op="' + xmlEsc(asStr(cl.op,'contains')) + '" param="' + xmlEsc(asStr(cl.param,'')) + '" value="' + xmlEsc(asStr(cl.value,'')) + '" />');
        }
        lines.push('      </Group>');
      }
      lines.push('    </Set>');
    }
    lines.push('  </Sets>');
    lines.push('  <Pairs>');
    for (var pi=0; pi<(cfg.pairs||[]).length; pi++){
      var p = cfg.pairs[pi] || {};
      lines.push('    <Pair a="' + xmlEsc(asStr(p.a,'__ALL__')) + '" b="' + xmlEsc(asStr(p.b,'__ALL__')) + '" enabled="' + (p.enabled === false ? 'false' : 'true') + '" />');
    }
    lines.push('  </Pairs>');
    lines.push('  <ExcludeSets>');
    for (var ei=0; ei<(cfg.excludeSetIds||[]).length; ei++){
      lines.push('    <SetRef id="' + xmlEsc(asStr(cfg.excludeSetIds[ei],'')) + '" />');
    }
    lines.push('  </ExcludeSets>');
    lines.push('  <ExcludeFamilies>');
    for (var fi=0; fi<(cfg.excludeFamilies||[]).length; fi++){
      lines.push('    <Item value="' + xmlEsc(asStr(cfg.excludeFamilies[fi],'')) + '" />');
    }
    lines.push('  </ExcludeFamilies>');
    lines.push('  <ExcludeCategories>');
    for (var ci2=0; ci2<(cfg.excludeCategories||[]).length; ci2++){
      lines.push('    <Item value="' + xmlEsc(asStr(cfg.excludeCategories[ci2],'')) + '" />');
    }
    lines.push('  </ExcludeCategories>');
    lines.push('</DuplicateInspectorRuleSet>');
    return lines.join('\n');
  }
  function readAttr(node, name, d){
    if (!node) return d || '';
    try{
      var v = node.getAttribute(name);
      return (v === undefined || v === null) ? (d || '') : String(v);
    }catch(e){ return d || ''; }
  }

  function SerializeRuleFile(){
    return serializeRuleFile();
  }
  function ParseRuleFileText(textRaw){
    return parseRuleFileText(textRaw);
  }

  function parseRuleFileText(textRaw){
    var text = String(textRaw || '').trim();
    if (!text) throw new Error('파일 내용이 비어 있습니다.');

    if (text.charAt(0) === '{'){
      var j = JSON.parse(text);
      if (j && j.ruleConfig){
        return {
          tolMm: clamp(asNum(j.tolMm, readTolMm()), 0.01, 1000),
          tolUnit: normalizeTolUnit(j.tolUnit || readTolUnit()),
          scopeMode: asStr(j.scopeMode, readScopeMode()),
          excludeKeywords: asStr(j.excludeKeywords, getExcludeKeywords().join(', ')),
          exportParams: asStr(j.exportParams, getExportParamText()),
          ruleConfig: normalizeRuleConfig(j.ruleConfig)
        };
      }
      return {
        tolMm: clamp(asNum(j.tolMm, readTolMm()), 0.01, 1000),
        tolUnit: normalizeTolUnit(j.tolUnit || readTolUnit()),
        scopeMode: readScopeMode(),
        excludeKeywords: getExcludeKeywords().join(', '),
        exportParams: asStr(j.exportParams, getExportParamText()),
        ruleConfig: normalizeRuleConfig(j)
      };
    }

    var parser = new DOMParser();
    var xml = parser.parseFromString(text, 'application/xml');
    if (xml.getElementsByTagName('parsererror').length){
      throw new Error('XML 형식이 올바르지 않습니다.');
    }

    var root = xml.documentElement;
    if (!root || root.nodeName !== 'DuplicateInspectorRuleSet'){
      throw new Error('지원하지 않는 설정 파일입니다.');
    }

    var settings = root.getElementsByTagName('Settings')[0];
    var out = {
      tolMm: clamp(asNum(readAttr(settings, 'tolMm', readTolMm()), readTolMm()), 0.01, 1000),
      tolUnit: normalizeTolUnit(readAttr(settings, 'tolUnit', readTolUnit())),
      scopeMode: asStr(readAttr(settings, 'scopeMode', readScopeMode()), readScopeMode()),
      excludeKeywords: asStr(readAttr(settings, 'excludeKeywords', getExcludeKeywords().join(', ')), ''),
      exportParams: asStr(readAttr(settings, 'exportParams', getExportParamText()), ''),
      ruleConfig: defaultRuleConfig()
    };

    var sets = [];
    var setNodes = root.getElementsByTagName('Sets')[0];
    var setEls = setNodes ? setNodes.getElementsByTagName('Set') : [];
    for (var si=0; si<setEls.length; si++){
      var se = setEls[si];
      var setObj = { id: asStr(readAttr(se,'id', ensureId('S')), ''), name: asStr(readAttr(se,'name','Set'), 'Set'), logic: (readAttr(se,'logic','or') === 'and' ? 'and' : 'or'), groups: [] };
      var child = se.firstElementChild;
      while (child){
        if (child.nodeName === 'Group'){
          var groupObj = { clauses: [] };
          var clauseEls = child.getElementsByTagName('Clause');
          for (var ci=0; ci<clauseEls.length; ci++){
            var ce = clauseEls[ci];
            groupObj.clauses.push({
              field: asStr(readAttr(ce,'field','category'),'category'),
              op: asStr(readAttr(ce,'op','contains'),'contains'),
              param: asStr(readAttr(ce,'param',''),''),
              value: asStr(readAttr(ce,'value',''),'')
            });
          }
          if (!groupObj.clauses.length) groupObj.clauses.push({ field:'category', op:'contains', param:'', value:'' });
          setObj.groups.push(groupObj);
        }
        child = child.nextElementSibling;
      }
      if (!setObj.groups.length) setObj.groups.push({ clauses:[{ field:'category', op:'contains', param:'', value:'' }] });
      sets.push(setObj);
    }

    var pairs = [];
    var pairNodes = root.getElementsByTagName('Pairs')[0];
    var pairEls = pairNodes ? pairNodes.getElementsByTagName('Pair') : [];
    for (var pi=0; pi<pairEls.length; pi++){
      var pe = pairEls[pi];
      pairs.push({ a: asStr(readAttr(pe,'a','__ALL__'),'__ALL__'), b: asStr(readAttr(pe,'b','__ALL__'),'__ALL__'), enabled: readAttr(pe,'enabled','true') !== 'false' });
    }

    var exSets = [];
    var exSetNodes = root.getElementsByTagName('ExcludeSets')[0];
    var exSetEls = exSetNodes ? exSetNodes.getElementsByTagName('SetRef') : [];
    for (var ei=0; ei<exSetEls.length; ei++) exSets.push(asStr(readAttr(exSetEls[ei],'id',''),''));

    var exFam = [];
    var exFamNodes = root.getElementsByTagName('ExcludeFamilies')[0];
    var exFamEls = exFamNodes ? exFamNodes.getElementsByTagName('Item') : [];
    for (var fi=0; fi<exFamEls.length; fi++) exFam.push(asStr(readAttr(exFamEls[fi],'value',''),''));

    var exCat = [];
    var exCatNodes = root.getElementsByTagName('ExcludeCategories')[0];
    var exCatEls = exCatNodes ? exCatNodes.getElementsByTagName('Item') : [];
    for (var ci3=0; ci3<exCatEls.length; ci3++) exCat.push(asStr(readAttr(exCatEls[ci3],'value',''),''));

    out.ruleConfig = normalizeRuleConfig({
      version: 1,
      sets: sets,
      pairs: pairs,
      excludeSetIds: exSets,
      excludeFamilies: exFam,
      excludeCategories: exCat
    });
    return out;
  }
  function handleRuleFileImport(inputEl){
    if (!inputEl || !inputEl.files || !inputEl.files.length) return;
    var file = inputEl.files[0];
    var reader = new FileReader();
    reader.onload = function(){
      try{
        var parsed = parseRuleFileText(reader.result || '');
        var item = upsertRuleLibraryItem(asStr(file && file.name, 'DuplicateInspector_RuleSet.xml'), parsed);
        renderRulePanel();
        try{ renderAppliedBar(); }catch(e0){}
        toast('XML을 목록에 불러왔습니다. 목록에서 선택 후 적용하세요.', 'ok', 1600);
      } catch(ex){
        toast(asStr(ex && ex.message, '설정 파일을 읽지 못했습니다.'), 'err', 2600);
      }
      try{ inputEl.value = ''; }catch(e4){}
    };
    reader.onerror = function(){
      toast('설정 파일을 읽지 못했습니다.', 'err', 2400);
      try{ inputEl.value = ''; }catch(e5){}
    };
    reader.readAsText(file, 'utf-8');
  }

  function closeRulePanel(){
    rulePanel.classList.remove('is-open');
    settingsBtn.classList.remove('is-active');
  }

  function renderRulePanel(){
    var cfg = loadRuleConfig();
    var meta = loadMeta();
    var params = Array.isArray(meta.parameters) ? meta.parameters : [];

    var tolUnit = readTolUnit();
    var tolEl = rulePanel.querySelector('[data-bind="tolValue"]');
    if (tolEl){
      var tolMeta = getTolDisplayMeta(tolUnit);
      tolEl.value = formatTolNumber(readTolDisplayValue(tolUnit));
      tolEl.min = tolMeta.min;
      tolEl.step = tolMeta.step;
    }
    var tolUnitEl = rulePanel.querySelector('[data-bind="tolUnit"]');
    if (tolUnitEl) tolUnitEl.value = tolUnit;

    var scopeEl = rulePanel.querySelector('[data-bind="scopeMode"]');
    if (scopeEl) scopeEl.value = scopeMode;

    var kwEl = rulePanel.querySelector('[data-bind="excludeKeywords"]');
    var exportParamEl = rulePanel.querySelector('[data-bind="exportParamNames"]');
    if (kwEl){
      try{ kwEl.value = String(localStorage.getItem(LS_EXCL_KW) || ''); } catch(e){ kwEl.value = ''; }
    }
    if (exportParamEl){
      try{ exportParamEl.value = String(localStorage.getItem(LS_EXPORT_PARAMS) || ''); } catch(e2){ exportParamEl.value = ''; }
    }

    function buildExcludeSummaryHtml(){
      var chips = [];
      var fams = cfg.excludeFamilies || [];
      var cats = cfg.excludeCategories || [];
      for (var i=0;i<fams.length;i++) chips.push('<span class="rp-chip"><span class="rp-chip-kind">패밀리</span>' + esc(String(fams[i] || '')) + '</span>');
      for (var j=0;j<cats.length;j++) chips.push('<span class="rp-chip"><span class="rp-chip-kind">시스템</span>' + esc(String(cats[j] || '')) + '</span>');
      var head = '<div class="rp-exsummary-top"><span>패밀리 ' + fams.length + '개</span><span>시스템 ' + cats.length + '개</span></div>';
      if (!chips.length) return head + '<div class="rp-emptyline">현재 제외 목록이 없습니다.</div>';
      return head + '<div class="rp-chiplist">' + chips.join('') + '</div>';
    }

    var exSum = rulePanel.querySelector('[data-slot="exsum"]');
    if (exSum) exSum.innerHTML = buildExcludeSummaryHtml();

    var cfgList = rulePanel.querySelector('[data-slot="cfglist"]');
    if (cfgList){
      var lib = loadRuleLibrary();
      var selId = getSelectedRuleLibraryId();
      if (!selId && lib.length){ selId = lib[0].id; setSelectedRuleLibraryId(selId); }
      cfgList.innerHTML = '';
      if (!lib.length){
        cfgList.innerHTML = '<div class="rp-emptyline">불러온 XML이 없습니다.</div>';
      } else {
        for (var li=0; li<lib.length; li++){
          var item = lib[li];
          var row = document.createElement('label');
          row.className = 'rp-fileitem' + (String(item.id) === String(selId) ? ' is-selected' : '');
          row.innerHTML = '<input type="radio" name="dup-rulelib" data-lib-sel="1" value="' + esc(item.id) + '" ' + (String(item.id) === String(selId) ? 'checked' : '') + ' />' +
                          '<div class="rp-fileitem-body"><div class="rp-fileitem-title">' + esc(item.name) + '</div><div class="rp-fileitem-sub">Set ' + (((item.parsed && item.parsed.ruleConfig && item.parsed.ruleConfig.sets) || []).length) + ' · Pair ' + (((item.parsed && item.parsed.ruleConfig && item.parsed.ruleConfig.pairs) || []).length) + '</div></div>';
          cfgList.appendChild(row);
        }
      }
    }

    var addExBtn = rulePanel.querySelector('[data-act="add-exset"]');
    if (addExBtn) addExBtn.disabled = !cfg.sets.length;

    function countClauses(groups){
      var total = 0;
      for (var i=0;i<(groups||[]).length;i++) total += ((groups[i] && groups[i].clauses) ? groups[i].clauses.length : 0);
      return total;
    }
    function buildSetOptions(){
      var opts = [];
      opts.push({v:'__ALL__', t:'__ALL__ (전체)'});
      for (var i=0;i<cfg.sets.length;i++){
        var s = cfg.sets[i];
        opts.push({v:s.id, t:(s.name || s.id)});
      }
      return opts;
    }

    var setsSlot = rulePanel.querySelector('[data-slot="sets"]');
    if (setsSlot){
      setsSlot.innerHTML = '';
      for (var si=0; si<cfg.sets.length; si++){
        var s = cfg.sets[si];
        var groupCount = (s.groups || []).length;
        var clauseCount = countClauses(s.groups || []);

        var card = document.createElement('div');
        card.className = 'rp-card rp-card--set';

        var head = document.createElement('div');
        head.className = 'rp-card-head';

        var left = document.createElement('div');
        left.className = 'rp-card-main';

        var badges = document.createElement('div');
        badges.className = 'rp-badges';
        badges.innerHTML = '<span class="rp-badge">SET ' + (si+1) + '</span>' +
          '<span class="rp-soft">ID ' + esc(String(s.id || '')) + '</span>' +
          '<span class="rp-soft">그룹 ' + groupCount + '개</span>' +
          '<span class="rp-soft">조건 ' + clauseCount + '개</span>';

        var name = document.createElement('input');
        name.className = 'rp-input rp-input--title';
        name.value = String(s.name || '');
        name.dataset.setName = String(si);
        name.placeholder = 'Set 이름';

        var logicRow = document.createElement('div');
        logicRow.className = 'rp-card-subline';
        logicRow.innerHTML = '<span class="rp-inline-label">그룹 결합</span>';

        var logicSel = document.createElement('select');
        logicSel.className = 'rp-select rp-select--sm';
        logicSel.dataset.setLogic = String(si);
        var o1 = document.createElement('option'); o1.value='or'; o1.textContent='OR (기본)';
        var o2 = document.createElement('option'); o2.value='and'; o2.textContent='AND';
        logicSel.appendChild(o1); logicSel.appendChild(o2);
        logicSel.value = (s.logic === 'and') ? 'and' : 'or';
        logicRow.appendChild(logicSel);

        left.appendChild(badges);
        left.appendChild(name);
        left.appendChild(logicRow);

        var right = document.createElement('div');
        right.className = 'rp-card-tools';
        var del = document.createElement('button');
        del.className = 'rp-x';
        del.textContent = '삭제';
        del.dataset.act = 'del-set';
        del.dataset.si = String(si);
        right.appendChild(del);

        head.appendChild(left);
        head.appendChild(right);
        card.appendChild(head);

        var groupsWrap = document.createElement('div');
        groupsWrap.className = 'rp-groups';

        for (var gi=0; gi<(s.groups||[]).length; gi++){
          var g = s.groups[gi];
          var gBox = document.createElement('div');
          gBox.className = 'rp-group';

          var gh = document.createElement('div');
          gh.className = 'rp-group-head';
          gh.innerHTML = '<div class="rp-group-title">OR 그룹 ' + (gi+1) + '<span class="rp-group-sub">조건 ' + ((g.clauses||[]).length) + '개 · 내부 AND</span></div>';

          var gdel = document.createElement('button');
          gdel.className = 'rp-x';
          gdel.textContent = '삭제';
          gdel.dataset.act = 'del-group';
          gdel.dataset.si = String(si);
          gdel.dataset.gi = String(gi);
          gh.appendChild(gdel);
          gBox.appendChild(gh);

          var clausesWrap = document.createElement('div');
          clausesWrap.className = 'rp-clauses';

          for (var ci=0; ci<(g.clauses||[]).length; ci++){
            var cl = g.clauses[ci];
            var row = document.createElement('div');
            row.className = 'rp-clause';

            var fSel = document.createElement('select');
            fSel.className = 'rp-select rp-select--sm';
            fSel.dataset.clField = String(si)+':'+String(gi)+':'+String(ci);
            for (var k=0;k<FIELD_OPTS.length;k++){
              var o = document.createElement('option');
              o.value = FIELD_OPTS[k].v;
              o.textContent = FIELD_OPTS[k].t;
              fSel.appendChild(o);
            }
            fSel.value = cl.field || 'category';

            var opSel = document.createElement('select');
            opSel.className = 'rp-select rp-select--sm';
            opSel.dataset.clOp = String(si)+':'+String(gi)+':'+String(ci);
            for (var k2=0;k2<OP_OPTS.length;k2++){
              var o22 = document.createElement('option');
              o22.value = OP_OPTS[k2].v;
              o22.textContent = OP_OPTS[k2].t;
              opSel.appendChild(o22);
            }
            opSel.value = cl.op || 'contains';

            var pSel = document.createElement('select');
            pSel.className = 'rp-select rp-select--sm rp-param';
            pSel.dataset.clParam = String(si)+':'+String(gi)+':'+String(ci);
            var p0 = document.createElement('option'); p0.value=''; p0.textContent='(param)';
            pSel.appendChild(p0);
            for (var k3=0;k3<params.length;k3++){
              var pn = String(params[k3] || '');
              if (!pn) continue;
              var o3 = document.createElement('option');
              o3.value = pn;
              o3.textContent = pn;
              pSel.appendChild(o3);
            }
            pSel.value = cl.param || '';
            pSel.style.display = (fSel.value === 'param') ? '' : 'none';

            var vInp = document.createElement('input');
            vInp.className = 'rp-input rp-input--sm';
            vInp.dataset.clVal = String(si)+':'+String(gi)+':'+String(ci);
            vInp.value = String(cl.value || '');
            vInp.placeholder = '비교 값';

            var cdel = document.createElement('button');
            cdel.className = 'rp-x';
            cdel.textContent = '삭제';
            cdel.dataset.act = 'del-clause';
            cdel.dataset.si = String(si);
            cdel.dataset.gi = String(gi);
            cdel.dataset.ci = String(ci);

            row.appendChild(fSel);
            row.appendChild(opSel);
            row.appendChild(pSel);
            row.appendChild(vInp);
            row.appendChild(cdel);
            clausesWrap.appendChild(row);
          }

          gBox.appendChild(clausesWrap);

          var gActions = document.createElement('div');
          gActions.className = 'rp-inline-actions';
          var addC = document.createElement('button');
          addC.className = 'rp-btn rp-btn--tiny';
          addC.textContent = '+ 조건 추가';
          addC.dataset.act = 'add-clause';
          addC.dataset.si = String(si);
          addC.dataset.gi = String(gi);
          gActions.appendChild(addC);
          gBox.appendChild(gActions);

          groupsWrap.appendChild(gBox);
        }

        card.appendChild(groupsWrap);

        var addGWrap = document.createElement('div');
        addGWrap.className = 'rp-inline-actions';
        var addG = document.createElement('button');
        addG.className = 'rp-btn rp-btn--tiny';
        addG.textContent = '+ OR 그룹 추가';
        addG.dataset.act = 'add-group';
        addG.dataset.si = String(si);
        addGWrap.appendChild(addG);
        card.appendChild(addGWrap);
        setsSlot.appendChild(card);
      }

      if (!cfg.sets.length){
        var h = document.createElement('div');
        h.className = 'rp-emptycard';
        h.innerHTML = '<div class="rp-emptytitle">등록된 Set이 없습니다.</div><div class="rp-emptysub">+ Set 추가를 눌러 필터 기준을 먼저 만들어 주세요.</div>';
        setsSlot.appendChild(h);
      }
    }

    var pairsSlot = rulePanel.querySelector('[data-slot="pairs"]');
    if (pairsSlot){
      pairsSlot.innerHTML = '';
      var setOpts = buildSetOptions();
      for (var pi=0; pi<cfg.pairs.length; pi++){
        var p = cfg.pairs[pi];
        var row = document.createElement('div');
        row.className = 'rp-pair';

        var pairMain = document.createElement('div');
        pairMain.className = 'rp-pair-main';

        var aSel = document.createElement('select');
        aSel.className = 'rp-select rp-select--sm';
        aSel.dataset.pA = String(pi);
        for (var k=0;k<setOpts.length;k++){
          var o = document.createElement('option');
          o.value = setOpts[k].v;
          o.textContent = setOpts[k].t;
          aSel.appendChild(o);
        }
        aSel.value = p.a || '__ALL__';

        var vs = document.createElement('span');
        vs.className = 'rp-vs';
        vs.textContent = 'vs';

        var bSel = document.createElement('select');
        bSel.className = 'rp-select rp-select--sm';
        bSel.dataset.pB = String(pi);
        for (var k2=0;k2<setOpts.length;k2++){
          var o2 = document.createElement('option');
          o2.value = setOpts[k2].v;
          o2.textContent = setOpts[k2].t;
          bSel.appendChild(o2);
        }
        bSel.value = p.b || '__ALL__';

        pairMain.appendChild(aSel);
        pairMain.appendChild(vs);
        pairMain.appendChild(bSel);

        var tools = document.createElement('div');
        tools.className = 'rp-pair-tools';
        var lab = document.createElement('label');
        lab.className = 'rp-chk';
        var cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.dataset.pEn = String(pi);
        cb.checked = (p.enabled !== false);
        lab.appendChild(cb);
        lab.appendChild(document.createTextNode('사용'));

        var del = document.createElement('button');
        del.className = 'rp-x';
        del.textContent = '삭제';
        del.dataset.act = 'del-pair';
        del.dataset.pi = String(pi);
        tools.appendChild(lab);
        tools.appendChild(del);

        row.appendChild(pairMain);
        row.appendChild(tools);
        pairsSlot.appendChild(row);
      }
      if (!cfg.pairs.length){
        var h2 = document.createElement('div');
        h2.className = 'rp-emptycard';
        h2.innerHTML = '<div class="rp-emptytitle">등록된 Pair가 없습니다.</div><div class="rp-emptysub">Set을 만든 뒤 + Pair 추가로 비교 대상을 지정하세요.</div>';
        pairsSlot.appendChild(h2);
      }
    }

    var exSlot = rulePanel.querySelector('[data-slot="exsets"]');
    if (exSlot){
      exSlot.innerHTML = '';
      var opts = [];
      for (var i2=0;i2<cfg.sets.length;i2++){
        var s2 = cfg.sets[i2];
        opts.push({v:s2.id, t:(s2.name || s2.id)});
      }
      for (var ei=0; ei<cfg.excludeSetIds.length; ei++){
        var sid = cfg.excludeSetIds[ei];
        var row2 = document.createElement('div');
        row2.className = 'rp-exrow';
        var sel = document.createElement('select');
        sel.className = 'rp-select rp-select--sm';
        sel.dataset.exSet = String(ei);
        for (var kk=0;kk<opts.length;kk++){
          var oo = document.createElement('option');
          oo.value = opts[kk].v;
          oo.textContent = opts[kk].t;
          sel.appendChild(oo);
        }
        sel.value = sid || (opts[0] ? opts[0].v : '');
        var del2 = document.createElement('button');
        del2.className = 'rp-x';
        del2.textContent = '삭제';
        del2.dataset.act = 'del-exset';
        del2.dataset.ei = String(ei);
        row2.appendChild(sel);
        row2.appendChild(del2);
        exSlot.appendChild(row2);
      }
      if (!cfg.excludeSetIds.length){
        var h3 = document.createElement('div');
        h3.className = 'rp-emptyline';
        h3.textContent = cfg.sets.length ? '등록된 Exclude Set이 없습니다.' : 'Set을 먼저 만들어야 Exclude Set을 등록할 수 있습니다.';
        exSlot.appendChild(h3);
      }
    }
    renderAppliedBar();
  }

  function exportRuleJson(){
    exportRuleFileInteractive();
  }
  function importRuleJson(){
    var inputEl = rulePanel.querySelector('[data-bind="cfgfile"]');
    if (inputEl) inputEl.click();
  }
  function applySelectedRuleLibrary(){
    var item = getSelectedRuleLibraryItem();
    if (!item){ toast('적용할 XML 목록이 없습니다.', 'warn', 1800); return; }
    if (!applyRuleLibraryItem(item.id)){ toast('선택한 XML을 적용하지 못했습니다.', 'err', 2200); return; }
    scopeMode = readScopeMode();
    renderRulePanel();
    renderAppliedBar();
    toast('선택한 XML 설정을 적용했습니다.', 'ok', 1300);
  }
  function deleteSelectedRuleLibrary(){
    var item = getSelectedRuleLibraryItem();
    if (!item){ toast('삭제할 XML을 먼저 선택하세요.', 'warn', 1800); return; }
    removeRuleLibraryItem(item.id);
    renderRulePanel();
    renderAppliedBar();
    toast('선택한 XML 목록을 삭제했습니다.', 'ok', 1200);
  }
  function clearSelectedRuleLibrary(){
    var lib = loadRuleLibrary();
    if (!lib.length){ toast('비울 XML 목록이 없습니다.', 'warn', 1600); return; }
    clearRuleLibrary();
    renderRulePanel();
    renderAppliedBar();
    toast('불러온 XML 목록을 비웠습니다.', 'ok', 1200);
  }
  function applyRulePanel(){
    var tolEl = rulePanel.querySelector('[data-bind="tolValue"]');
    var tolUnitEl = rulePanel.querySelector('[data-bind="tolUnit"]');
    var scopeEl = rulePanel.querySelector('[data-bind="scopeMode"]');
    var kwEl = rulePanel.querySelector('[data-bind="excludeKeywords"]');
    var exportParamEl = rulePanel.querySelector('[data-bind="exportParamNames"]');

    var tolUnit = normalizeTolUnit(tolUnitEl ? tolUnitEl.value : readTolUnit());
    var mm = getTolMmFromDisplay(tolEl ? tolEl.value : DEFAULT_TOL_MM, tolUnit);
    try{ localStorage.setItem(LS_TOL_MM, String(mm)); } catch(e){}
    try{ localStorage.setItem(LS_TOL_UNIT, tolUnit); } catch(e0){}

    scopeMode = asStr(scopeEl ? scopeEl.value : 'all', 'all');
    try{ localStorage.setItem(LS_SCOPE, scopeMode); } catch(e2){}

    var kw = asStr(kwEl ? kwEl.value : '', '').trim();
    try{ localStorage.setItem(LS_EXCL_KW, kw); } catch(e3){}
    var exportParamText = asStr(exportParamEl ? exportParamEl.value : '', '').trim();
    try{ localStorage.setItem(LS_EXPORT_PARAMS, exportParamText); } catch(e4){}

    toast('설정을 적용했습니다.', 'ok', 1200);
    try{ renderRulePanel(); }catch(e5){}
    try{ renderAppliedBar(); }catch(e6){}
  }

  function addSet(){
    var cfg = loadRuleConfig();
    cfg.sets.push({
      id: ensureId('S'),
      name: 'Set',
      logic: 'or',
      groups: [{ clauses: [{ field:'category', op:'contains', value:'', param:'' }] }]
    });
    saveRuleConfig(cfg);
    toast('Set가 추가되었습니다.', 'ok', 1200);
  }
  function delSet(si){
    var cfg = loadRuleConfig();
    if (si < 0 || si >= cfg.sets.length) return;
    var removedId = cfg.sets[si].id;
    cfg.sets.splice(si,1);
    cfg.excludeSetIds = (cfg.excludeSetIds||[]).filter(function(x){ return String(x) !== String(removedId); });
    cfg.pairs = (cfg.pairs||[]).filter(function(p){ return String(p.a) !== String(removedId) && String(p.b) !== String(removedId); });
    saveRuleConfig(cfg);
    toast('Set가 삭제되었습니다.', 'ok', 1200);
  }
  function addGroup(si){
    var cfg = loadRuleConfig();
    if (!cfg.sets[si]) return;
    cfg.sets[si].groups = cfg.sets[si].groups || [];
    cfg.sets[si].groups.push({ clauses: [{ field:'category', op:'contains', value:'', param:'' }] });
    saveRuleConfig(cfg);
  }
  function delGroup(si, gi){
    var cfg = loadRuleConfig();
    if (!cfg.sets[si] || !cfg.sets[si].groups[gi]) return;
    cfg.sets[si].groups.splice(gi,1);
    saveRuleConfig(cfg);
  }
  function addClause(si, gi){
    var cfg = loadRuleConfig();
    if (!cfg.sets[si] || !cfg.sets[si].groups[gi]) return;
    cfg.sets[si].groups[gi].clauses = cfg.sets[si].groups[gi].clauses || [];
    cfg.sets[si].groups[gi].clauses.push({ field:'category', op:'contains', value:'', param:'' });
    saveRuleConfig(cfg);
  }
  function delClause(si, gi, ci){
    var cfg = loadRuleConfig();
    if (!cfg.sets[si] || !cfg.sets[si].groups[gi] || !cfg.sets[si].groups[gi].clauses[ci]) return;
    cfg.sets[si].groups[gi].clauses.splice(ci,1);
    saveRuleConfig(cfg);
  }
  function addPair(){
    var cfg = loadRuleConfig();
    var base = cfg.sets.length ? cfg.sets[0].id : '__ALL__';
    cfg.pairs.push({ a: base, b: base, enabled: true });
    saveRuleConfig(cfg);
    toast('Pair가 추가되었습니다.', 'ok', 1200);
  }
  function delPair(pi){
    var cfg = loadRuleConfig();
    if (pi < 0 || pi >= cfg.pairs.length) return;
    cfg.pairs.splice(pi,1);
    saveRuleConfig(cfg);
  }
  function addExcludeSet(){
    var cfg = loadRuleConfig();
    if (!cfg.sets.length){
      toast('Set이 없습니다. 먼저 Set을 추가하세요.', 'warn', 1600);
      return;
    }
    cfg.excludeSetIds.push(String(cfg.sets[0].id));
    saveRuleConfig(cfg);
    toast('Exclude Set이 추가되었습니다.', 'ok', 1200);
  }
  function delExcludeSet(ei){
    var cfg = loadRuleConfig();
    if (ei < 0 || ei >= cfg.excludeSetIds.length) return;
    cfg.excludeSetIds.splice(ei,1);
    saveRuleConfig(cfg);
  }

  function makeBtn(text, cls){
    var b = document.createElement('button');
    b.type = 'button';
    b.className = cls || '';
    b.textContent = text;
    return b;
  }
  function chip(text){
    var c = div('chip');
    c.textContent = text;
    return c;
  }
  function cell(content, cls){
    var c = document.createElement('div');
    c.className = 'cell ' + (cls || '');
    // content may be a Node (button/div) or plain text
    if (content && typeof content === 'object' && content.nodeType) {
      c.textContent = '';
      c.appendChild(content);
    } else {
      c.textContent = String((content === undefined || content === null) ? '' : content);
    }
    return c;
  }
  function buildSkeleton(n){
    var wrap = div('dup-skeleton');
    var cnt = asNum(n, 6);
    for (var i=0;i<cnt;i++){
      var line = div('sk-row');
      line.appendChild(div('sk-chip'));
      line.appendChild(div('sk-id'));
      line.appendChild(div('sk-wide'));
      line.appendChild(div('sk-wide'));
      line.appendChild(div('sk-act'));
      wrap.appendChild(line);
    }
    return wrap;
  }
}

// ---- styles ----
function ensureStyles(){
  if (document.getElementById('dup-style-v76')) return;
  var st = document.createElement('style');
  st.id = 'dup-style-v76';

  var css = [
    '.dup-toolbar{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;flex-wrap:wrap;}',
    '.feature-actions{display:flex;align-items:flex-end;gap:12px;flex-wrap:wrap;}',
    '.dup-mode-shell,.dup-settings-shell{display:flex;flex-direction:column;gap:6px;min-width:0;}',
    '.dup-mode-label{font-size:11px;font-weight:800;letter-spacing:.04em;text-transform:uppercase;opacity:.68;padding-left:2px;}',
    '.dup-mode-switch{display:inline-flex;align-items:center;padding:4px;border-radius:16px;background:var(--surface-subtle);border:1px solid var(--border-soft);box-shadow:inset 0 1px 0 color-mix(in srgb, var(--text-main) 8%, transparent);}',
    '.dup-mode-btn{height:38px;padding:0 16px;border:none;background:transparent;border-radius:12px;font-weight:800;cursor:pointer;transition:all .18s ease;white-space:nowrap;color:inherit;}',
    '.dup-mode-btn.is-active{background:linear-gradient(180deg,var(--accent),var(--accent-strong));color:var(--accentText);box-shadow:var(--shadow-accent-soft);}',
    '.dup-mode-btn:disabled{opacity:.5;cursor:default;}',
    '.dup-settings-btn{height:38px;padding:0 16px;border-radius:12px;border:1px solid var(--border-accent-soft);background:var(--surface-note);font-weight:800;cursor:pointer;white-space:nowrap;color:inherit;}',
    '.dup-settings-btn.is-active{background:var(--surface-note-strong);border-color:var(--border-accent-strong);}',
    '.dup-row.is-deleted{opacity:.55;}',
    '.row-actions{display:flex;gap:10px;flex-wrap:nowrap;}',
    '.row-actions .table-action-btn{white-space:nowrap;min-width:88px;height:32px;line-height:32px;padding:0 12px;display:inline-flex;align-items:center;justify-content:center;}',
    '.pair-cardlist{padding:12px 14px 16px 14px;display:flex;flex-direction:column;gap:12px;}',
    '.pair-card{display:flex;flex-direction:column;gap:10px;border:1px solid var(--border-soft);border-radius:16px;background:var(--surface-elevated);padding:12px 14px;box-shadow:var(--shadow-soft);}',
    '.pair-main{display:grid;grid-template-columns:1fr 44px 1fr;gap:12px;align-items:center;}',
    '.pair-mid{text-align:center;opacity:.55;font-size:20px;font-weight:700;}',
    '.pair-lbl{font-weight:800;font-size:12px;opacity:.72;margin-bottom:4px;}',
    '.pair-meta{margin-top:6px;font-size:12px;opacity:.86;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}',
    '.pair-actions{display:flex;justify-content:flex-end;}',
    '.pair-note{padding:10px 12px;border-radius:12px;background:color-mix(in srgb, var(--warning) 14%, transparent);border:1px solid color-mix(in srgb, var(--warning) 34%, transparent);font-size:12px;font-weight:700;color:inherit;}',
    '.conn-cell{display:inline-block;min-width:18px;padding:2px 6px;border-radius:999px;border:1px solid var(--border-soft);text-align:center;}',

    '.dup-rulemodal{position:fixed;inset:0;z-index:9000;display:none;}',
    '.dup-rulemodal.is-open{display:block;}',
    '.dup-rulemodal .rm-backdrop{position:absolute;inset:0;background:var(--overlay-scrim);backdrop-filter:blur(4px);}',
    '.dup-rulemodal .rm-window{position:fixed;inset:20px;border-radius:24px;overflow:hidden;border:1px solid var(--border-soft);background:var(--surface-modal);box-shadow:var(--shadow-card);display:flex;flex-direction:column;}',
    '.dup-rulemodal .rm-head{display:flex;align-items:center;justify-content:space-between;gap:14px;padding:18px 20px;border-bottom:1px solid var(--border-soft);background:var(--surface-elevated);}',
    '.dup-rulemodal .rm-title{font-size:20px;font-weight:900;letter-spacing:-.02em;}',
    '.dup-rulemodal .rm-sub{font-size:13px;opacity:.78;margin-top:4px;line-height:1.45;}',
    '.dup-rulemodal .rm-actions{display:flex;gap:8px;flex-wrap:wrap;align-items:center;justify-content:flex-end;}',
    '.dup-rulemodal .rm-body{padding:20px 20px 24px 20px;overflow:auto;background:var(--surface-panel);display:flex;flex-direction:column;gap:14px;}',
    '.dup-rulemodal .rp-overview{padding:16px 18px;border-radius:18px;background:var(--surface-note);border:1px solid var(--border-accent-soft);}',
    '.dup-rulemodal .rp-overview-title{font-size:14px;font-weight:900;margin-bottom:6px;}',
    '.dup-rulemodal .rp-overview-copy{font-size:13px;line-height:1.55;opacity:.86;}',
    '.dup-rulemodal .rp-sec{padding:16px 18px;border-radius:18px;border:1px solid var(--border-soft);background:var(--surface-elevated);box-shadow:var(--shadow-soft);}',
    '.dup-rulemodal .rp-sec--compact{padding:14px 16px;}',
    '.dup-rulemodal .rp-sec-head{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:12px;}',
    '.dup-rulemodal .rp-sec-title{font-size:16px;font-weight:900;letter-spacing:-.01em;}',
    '.dup-rulemodal .rp-sec-sub{font-size:12px;opacity:.72;line-height:1.45;margin-top:4px;}',
    '.dup-rulemodal .rp-section-actions{display:flex;gap:8px;flex-wrap:wrap;align-items:center;}',
    '.dup-rulemodal .rp-grid{display:grid;grid-template-columns:170px minmax(0,1fr);gap:10px 14px;align-items:center;}',
    '.dup-rulemodal .rp-label{font-size:12px;font-weight:700;opacity:.82;}',
    '.dup-rulemodal .rp-inline-field{display:grid;grid-template-columns:minmax(0,1fr) 120px;gap:8px;align-items:center;}',
    '.dup-rulemodal .rp-hint{font-size:12px;opacity:.78;margin-top:10px;line-height:1.45;}',
    '.dup-rulemodal .rp-inline-note{margin-bottom:10px;padding:10px 12px;border-radius:12px;background:var(--surface-help);border:1px dashed var(--border-accent-soft);font-size:12px;line-height:1.45;}',
    '.dup-rulemodal .rp-btn{height:38px;padding:0 14px;border-radius:12px;border:1px solid var(--border-soft);background:var(--surface-control);color:inherit;cursor:pointer;font-weight:800;}',
    '.dup-rulemodal .rp-btn:hover{transform:translateY(-1px);box-shadow:var(--shadow-soft);}',
    '.dup-rulemodal .rp-btn:disabled{opacity:.45;cursor:not-allowed;transform:none;box-shadow:none;}',
    '.dup-rulemodal .rp-btn--ghost{background:var(--surface-subtle);}',
    '.dup-rulemodal .rp-btn--accent{background:var(--surface-note);border-color:var(--border-accent-soft);}',
    '.dup-rulemodal .rp-btn--add{height:34px;padding:0 12px;border-radius:10px;background:var(--surface-note);border-color:var(--border-accent-soft);}',
    '.dup-rulemodal .rp-btn--primary{background:linear-gradient(180deg,var(--accent),var(--accent-strong));color:var(--accentText);border-color:var(--accent);box-shadow:var(--shadow-accent-soft);}',
    '.dup-rulemodal .rp-btn--tiny{height:30px;padding:0 10px;border-radius:10px;font-size:12px;}',
    '.dup-rulemodal .rp-input,.dup-rulemodal .rp-select{width:100%;height:38px;padding:0 12px;border-radius:12px;border:1px solid var(--border-soft);background:var(--surface-control);color:inherit;box-sizing:border-box;}',
    '.dup-rulemodal .rp-select--unit{min-width:0;}',
    '.dup-rulemodal .rp-input:focus,.dup-rulemodal .rp-select:focus,.dup-expicker .rp-input:focus{outline:none;border-color:var(--border-accent-strong);box-shadow:0 0 0 4px var(--focus-ring);}',
    '.dup-rulemodal .rp-input--sm,.dup-rulemodal .rp-select--sm{height:34px;padding:0 10px;border-radius:10px;}',
    '.dup-rulemodal .rp-input--title{height:40px;font-size:15px;font-weight:800;}',
    '.dup-rulemodal .rp-card{padding:14px;border-radius:18px;border:1px solid var(--border-soft);margin-bottom:12px;background:var(--surface-elevated);box-shadow:var(--shadow-soft);}',
    '.dup-rulemodal .rp-card-head{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin-bottom:12px;}',
    '.dup-rulemodal .rp-card-main{flex:1;min-width:0;display:flex;flex-direction:column;gap:8px;}',
    '.dup-rulemodal .rp-card-tools{display:flex;align-items:center;gap:8px;}',
    '.dup-rulemodal .rp-badges{display:flex;gap:8px;flex-wrap:wrap;}',
    '.dup-rulemodal .rp-badge{display:inline-flex;align-items:center;height:24px;padding:0 10px;border-radius:999px;background:var(--surface-note);color:inherit;font-size:11px;font-weight:900;letter-spacing:.03em;}',
    '.dup-rulemodal .rp-soft{display:inline-flex;align-items:center;height:24px;padding:0 10px;border-radius:999px;background:var(--surface-subtle);font-size:11px;font-weight:700;opacity:.86;}',
    '.dup-rulemodal .rp-card-subline{display:flex;align-items:center;gap:8px;flex-wrap:wrap;}',
    '.dup-rulemodal .rp-inline-label{font-size:12px;font-weight:800;opacity:.72;}',
    '.dup-rulemodal .rp-x{height:32px;padding:0 12px;border-radius:10px;border:1px solid color-mix(in srgb, var(--danger) 36%, transparent);background:var(--danger-soft);color:inherit;font-size:12px;font-weight:800;cursor:pointer;}',
    '.dup-rulemodal .rp-group{padding:12px;border-radius:16px;border:1px dashed var(--border-accent-soft);margin-top:10px;background:var(--surface-help);}',
    '.dup-rulemodal .rp-group-head{display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:10px;}',
    '.dup-rulemodal .rp-group-title{font-weight:900;font-size:13px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;}',
    '.dup-rulemodal .rp-group-sub{font-size:11px;font-weight:700;opacity:.68;}',
    '.dup-rulemodal .rp-clauses{display:flex;flex-direction:column;gap:8px;}',
    '.dup-rulemodal .rp-clause{display:grid;grid-template-columns:1.1fr 1fr 1.2fr 1.5fr auto;gap:8px;align-items:center;}',
    '.dup-rulemodal .rp-param{min-width:120px;}',
    '.dup-rulemodal .rp-inline-actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px;}',
    '.dup-rulemodal .rp-pair,.dup-rulemodal .rp-exrow{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:12px 14px;border-radius:14px;border:1px solid var(--border-soft);background:var(--surface-help);margin-bottom:10px;}',
    '.dup-rulemodal .rp-pair-main{display:grid;grid-template-columns:minmax(0,1fr) 34px minmax(0,1fr);gap:8px;align-items:center;flex:1;}',
    '.dup-rulemodal .rp-pair-tools{display:flex;align-items:center;gap:10px;}',
    '.dup-rulemodal .rp-vs{display:inline-flex;align-items:center;justify-content:center;height:34px;border-radius:10px;background:var(--surface-subtle);font-size:12px;font-weight:900;opacity:.76;}',
    '.dup-rulemodal .rp-chk{display:flex;gap:6px;align-items:center;font-size:12px;font-weight:800;opacity:.9;white-space:nowrap;}',
    '.dup-rulemodal .rp-exsummary{display:flex;flex-direction:column;gap:10px;min-height:52px;}',
    '.dup-rulemodal .rp-exsummary-top{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;font-weight:800;opacity:.74;}',
    '.dup-rulemodal .rp-chiplist{display:flex;gap:8px;flex-wrap:wrap;}',
    '.dup-rulemodal .rp-chip{display:inline-flex;align-items:center;gap:8px;padding:7px 10px;border-radius:999px;background:var(--surface-note);border:1px solid var(--border-accent-soft);font-size:12px;font-weight:700;}',
    '.dup-rulemodal .rp-chip-kind{font-size:10px;font-weight:900;letter-spacing:.04em;text-transform:uppercase;opacity:.7;}',
    '.dup-rulemodal .rp-filebox{display:flex;flex-direction:column;gap:6px;padding:12px 14px;border-radius:14px;background:var(--surface-subtle);border:1px solid var(--border-soft);}',
    '.dup-rulemodal .rp-filecopy{font-size:13px;font-weight:800;}',
    '.dup-rulemodal .rp-filehint{font-size:12px;opacity:.74;line-height:1.45;}',
    '.dup-rulemodal .rp-emptycard{padding:16px;border-radius:16px;border:1px dashed var(--border-soft);background:var(--surface-empty);}',
    '.dup-rulemodal .rp-emptytitle{font-size:14px;font-weight:900;margin-bottom:6px;}',
    '.dup-rulemodal .rp-emptysub,.dup-rulemodal .rp-emptyline{font-size:12px;opacity:.76;line-height:1.45;}',
    '.dup-expicker{position:fixed;inset:0;z-index:9100;}',
    '.dup-expicker .rm-backdrop{position:absolute;inset:0;background:var(--overlay-scrim);backdrop-filter:blur(4px);}',
    '.dup-expicker .rm-window{position:fixed;inset:26px;border-radius:24px;overflow:hidden;border:1px solid var(--border-soft);background:var(--surface-modal);box-shadow:var(--shadow-card);display:flex;flex-direction:column;}',
    '.dup-expicker .rm-head{position:sticky;top:0;z-index:2;display:flex;align-items:center;justify-content:space-between;gap:12px;padding:18px 20px;border-bottom:1px solid var(--border-soft);background:var(--surface-elevated);}',
    '.dup-expicker .rm-body{padding:18px 20px 20px 20px;overflow:auto;background:var(--surface-panel);}',
    '.dup-expicker .expicked{display:flex;flex-direction:column;gap:10px;padding:14px 16px;margin-bottom:14px;border:1px solid var(--border-accent-soft);border-radius:16px;background:var(--surface-elevated);}',
    '.dup-expicker .expicked-meta{display:flex;gap:10px;flex-wrap:wrap;}',
    '.dup-expicker .expicked-count{display:inline-flex;align-items:center;height:28px;padding:0 10px;border-radius:999px;background:var(--surface-note);color:var(--accent);font-size:12px;font-weight:800;}',
    '.dup-expicker .expicked-list{display:flex;gap:8px;flex-wrap:wrap;}',
    '.dup-expicker .expicked-chip{display:inline-flex;align-items:center;min-height:30px;padding:6px 10px;border-radius:999px;background:var(--surface-subtle);font-size:12px;font-weight:700;color:inherit;}',
    '.dup-expicker .rp-btn{height:38px;padding:0 14px;border-radius:12px;border:1px solid var(--border-soft);background:var(--surface-control);cursor:pointer;font-weight:800;color:inherit;}',
    '.dup-expicker .rp-btn--ghost{background:var(--surface-subtle);}',
    '.dup-expicker .rp-btn--primary{background:linear-gradient(180deg,var(--accent),var(--accent-strong));color:var(--accentText);font-weight:900;border-color:var(--accent);}',
    '.dup-expicker .rp-input{width:100%;height:40px;padding:0 14px;border-radius:12px;border:1px solid var(--border-soft);background:var(--surface-control);box-sizing:border-box;color:inherit;}',
    '.exfam-wrap{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:12px;}',
    '@media (max-width:980px){.exfam-wrap{grid-template-columns:1fr;}}',
    '.exfam-col{border:1px solid var(--border-soft);border-radius:18px;background:var(--surface-elevated);overflow:hidden;box-shadow:var(--shadow-soft);}',
    '.exfam-title{padding:12px 14px;font-weight:900;border-bottom:1px solid var(--border-soft);background:var(--surface-note);}',
    '.exfam-list{padding:10px 12px;max-height:420px;overflow:auto;display:flex;flex-direction:column;gap:4px;}',
    '.exfam-item{display:flex;gap:10px;align-items:center;padding:8px 8px;border-radius:12px;transition:background .15s ease;}',
    '.exfam-item:hover{background:var(--surface-note);}',
    '.exfam-item input{width:16px;height:16px;}',
    '.exfam-text{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:13px;font-weight:700;}',
    '.exfam-empty{opacity:.76;font-size:12px;padding:10px 4px;line-height:1.45;}',
    '.exfam-actions{display:flex;gap:8px;margin-top:14px;justify-content:flex-end;flex-wrap:wrap;}',
    '.dup-appliedbar{margin:10px 0 0 0;padding:10px 12px;border-radius:14px;border:1px solid var(--border-soft);background:var(--surface-elevated);backdrop-filter: blur(6px);}',
    '.dup-appliedbar.hidden{display:none;}',
    '.ap-label{font-weight:800;font-size:12px;opacity:.75;margin-bottom:6px;}',
    '.ap-chips{display:flex;gap:8px;flex-wrap:wrap;align-items:center;}',
    '.ap-chip{display:inline-flex;align-items:center;padding:4px 10px;border-radius:999px;border:1px solid var(--border-accent-soft);background:var(--surface-note);font-size:12px;opacity:.95;}',
    '.ap-sub{margin-top:6px;font-size:12px;opacity:.7;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}',
    '.grp-body{border-top:1px solid var(--border-soft);}',
    '.grp-body .dup-subhead,.grp-body .dup-row{display:grid;grid-template-columns:44px 92px 1.1fr 1.2fr 1.2fr 76px 190px;gap:10px;align-items:center;padding:8px 12px;}',
    '.grp-body .dup-subhead{position:sticky;top:0;background:var(--surface-help);z-index:1;}',
    '.grp-body .dup-subhead .cell{font-weight:800;font-size:12px;opacity:.78;}',
    '.grp-body .dup-row{border-top:1px solid var(--border-soft);background:var(--surface-control);}',
    '.grp-body .dup-row:hover{background:var(--surface-note);}',
    '.grp-body .cell{min-width:0;}',
    '.grp-body .cell.ell{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}',
    '.grp-body .cell.right{text-align:right;}',
    '.grp-body .cell.mono{font-variant-numeric:tabular-nums;}',
    '.grp-body .cell.ck{display:flex;justify-content:center;}',
    '.grp-body .ckbox{width:16px;height:16px;}',
    '.grp-body .cell.conn{display:flex;justify-content:flex-end;}',
    '.row-actions{display:flex;gap:10px;justify-content:flex-end;}',
    '.row-actions .table-action-btn{white-space:nowrap;min-width:92px;height:32px;line-height:32px;padding:0 12px;}',
    '@media (max-width:1200px){.dup-rulemodal .rp-clause{grid-template-columns:1fr 1fr 1fr 1.2fr auto;}}',
    '@media (max-width:1100px){.grp-body .dup-subhead,.grp-body .dup-row{grid-template-columns:44px 92px 1fr 1fr 1fr 66px 176px;gap:8px;}}',
    '@media (max-width:980px){.dup-rulemodal .rm-head,.dup-expicker .rm-head{align-items:flex-start;}.dup-rulemodal .rm-actions,.dup-expicker .rm-actions{width:100%;justify-content:flex-start;}.dup-rulemodal .rp-grid{grid-template-columns:1fr;}.dup-rulemodal .rp-inline-field{grid-template-columns:1fr;}.dup-rulemodal .rp-clause{grid-template-columns:1fr;}.dup-rulemodal .rp-pair,.dup-rulemodal .rp-exrow{flex-direction:column;align-items:stretch;}.dup-rulemodal .rp-pair-main{grid-template-columns:1fr;}.dup-rulemodal .rp-vs{width:100%;}.dup-toolbar{align-items:stretch;}.feature-actions{align-items:stretch;}}',
    '@media (max-width:860px){.grp-body .dup-subhead,.grp-body .dup-row{grid-template-columns:44px 92px 1fr 1fr 0.9fr 0.7fr 160px;}}'
  ].join('\n');

  st.textContent = css;
  document.head.appendChild(st);
}
