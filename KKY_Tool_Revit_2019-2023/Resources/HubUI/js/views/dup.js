// Resources/HubUI/js/views/dup.js
import { clear, div, toast, showExcelSavedDialog, chooseExcelMode } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { onHost, post } from '../core/bridge.js';

/* v59: 파싱 오류 방지용 full replace (?. / ?? 미사용) */

const RESP_ROWS_EVENTS = ['dup:list','dup:rows','duplicate:list'];
const EV_RUN_REQ='dup:run', EV_DELETE_REQ='duplicate:delete', EV_RESTORE_REQ='duplicate:restore', EV_SELECT_REQ='duplicate:select', EV_EXPORT_REQ='duplicate:export';
const EV_DELETED_ONE='dup:deleted', EV_RESTORED_ONE='dup:restored', EV_EXPORTED='dup:exported';
const LS_MODE='kky_dup_mode', LS_TOL_MM='kky_dup_tol_mm', LS_SCOPE='kky_dup_scope_mode', LS_EXCL_KW='kky_dup_excl_kw', LS_META='kky_dup_meta_v1', LS_RULECFG='kky_dup_ruleset_v1';
const DEFAULT_TOL_MM=4.7625;

function isObj(x){return x&&typeof x==='object';}
function co(a,b){return (a===undefined||a===null)?b:a;}
function asStr(x,d){return (x===undefined||x===null)?(d||''):String(x);}
function asNum(x,d){var n=Number(x);return Number.isFinite(n)?n:(d||0);}
function esc(s){return String(co(s,'')).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
function uniq(list){var out=[],seen={};for(var i=0;i<(list||[]).length;i++){var s=String(list[i]||'').trim();if(!s)continue;var k=s.toLowerCase();if(seen[k])continue;seen[k]=1;out.push(s);}return out;}
function clamp(n,a,b){return Math.max(a,Math.min(b,n));}
function mmToFeet(mm){return mm/304.8;}
function readMode(){try{var m=String(localStorage.getItem(LS_MODE)||'').trim();if(m==='duplicate'||m==='clash')return m;}catch(e){}return 'duplicate';}
function readScopeMode(){try{var m=String(localStorage.getItem(LS_SCOPE)||'').trim();if(m==='all'||m==='scope'||m==='exclude')return m;}catch(e){}return 'all';}
function readTolMm(){try{var raw=localStorage.getItem(LS_TOL_MM);var n=Number(String(raw||'').trim());if(Number.isFinite(n)&&n>0)return clamp(n,0.01,1000);}catch(e){}return DEFAULT_TOL_MM;}
function readTolFeet(){return Math.max(0.000001, mmToFeet(readTolMm()));}
function getExcludeKeywords(){try{var raw=String(localStorage.getItem(LS_EXCL_KW)||'').trim();if(!raw)return [];return raw.split(',').map(function(s){return s.trim();}).filter(Boolean);}catch(e){return [];}}

function defaultRuleConfig(){return {version:1,sets:[],pairs:[],excludeSetIds:[],excludeFamilies:[],excludeCategories:[]};}
function normalizeRuleConfig(cfg){
  var c=isObj(cfg)?cfg:{};
  var out = {
    version:1,
    sets:Array.isArray(c.sets)?c.sets:[],
    pairs:Array.isArray(c.pairs)?c.pairs:[],
    excludeSetIds:Array.isArray(c.excludeSetIds)?c.excludeSetIds.map(String).filter(Boolean):[],
    excludeFamilies:Array.isArray(c.excludeFamilies)?c.excludeFamilies.map(String).filter(Boolean):[],
    excludeCategories:Array.isArray(c.excludeCategories)?c.excludeCategories.map(String).filter(Boolean):[]
  };

  // sets normalize
  out.sets = (out.sets||[]).map(function(s, idx){
    var ss=isObj(s)?s:{};
    var id = asStr(ss.id, 'S'+(idx+1));
    var name = asStr(ss.name, id);
    var logic = (ss.logic==='and' || ss.logic==='or') ? ss.logic : 'or';
    var groups = Array.isArray(ss.groups)?ss.groups:[];
    return { id:id, name:name, logic:logic, groups:groups };
  });
  (out.sets||[]).forEach(function(s){
    s.groups = (s.groups||[]).map(function(g){
      var gg=isObj(g)?g:{};
      return { clauses: Array.isArray(gg.clauses)?gg.clauses:[] };
    });
    (s.groups||[]).forEach(function(g){
      g.clauses = (g.clauses||[]).map(function(cl){
        var cc=isObj(cl)?cl:{};
        return {
          field: asStr(cc.field,'category'),
          op: asStr(cc.op,'contains'),
          value: asStr(cc.value,''),
          param: asStr(cc.param,'')
        };
      });
    });
  });

  // pairs normalize
  out.pairs = (out.pairs||[]).map(function(p){
    var pp=isObj(p)?p:{};
    return { a: asStr(pp.a,''), b: asStr(pp.b,''), enabled: (pp.enabled!==false) };
  }).filter(function(p){ return !!p.a && !!p.b; });

  return out;
}


function loadRuleConfig(){try{var raw=localStorage.getItem(LS_RULECFG);if(!raw)return defaultRuleConfig();return normalizeRuleConfig(JSON.parse(raw));}catch(e){return defaultRuleConfig();}}
function saveRuleConfig(cfg){var norm=normalizeRuleConfig(cfg);try{localStorage.setItem(LS_RULECFG,JSON.stringify(norm));}catch(e){}return norm;}
function loadMeta(){try{var raw=localStorage.getItem(LS_META);if(!raw)return {};var m=JSON.parse(raw);return isObj(m)?m:{};}catch(e){return {};}}

let exPickerEl=null, exPickerOpen=false;
function ensureExcludePicker(){
  var el=document.getElementById('dup-expicker');
  if(!el){
    el=document.createElement('div');el.id='dup-expicker';el.className='dup-expicker';el.style.display='none';
    el.innerHTML = '<div class="rm-backdrop" data-act="close"></div><div class="rm-window"><div class="rm-head"><div class="rm-head-left"><div class="rm-title">Exclude 목록 선택</div><div class="rm-sub">모델 패밀리 / 시스템 카테고리를 체크해 제외합니다.</div></div><div class="rm-actions"><button class="rp-btn rp-btn--ghost" data-act="refresh">목록 새로고침</button><button class="rp-btn" data-act="close">닫기</button></div></div><div class="rm-body"><div class="rp-sec"><div class="rp-sec-title">검색</div><div class="rp-grid"><div class="rp-label">키워드</div><input class="rp-input" type="text" placeholder="검색…" data-bind="q"/></div></div><div class="exfam-wrap"><div class="exfam-col"><div class="exfam-title">모델 패밀리</div><div class="exfam-list" data-slot="fam"></div></div><div class="exfam-col"><div class="exfam-title">시스템(카테고리)</div><div class="exfam-list" data-slot="cat"></div></div></div><div class="exfam-actions"><button class="rp-btn rp-btn--ghost" data-act="all">전체 체크</button><button class="rp-btn rp-btn--ghost" data-act="none">전체 해제</button><button class="rp-btn rp-btn--primary" data-act="apply">적용</button></div></div></div>';
    document.body.appendChild(el);
    el.addEventListener('click',function(e){
      var t=e.target;var act=(t&&t.dataset)?t.dataset.act:'';
      if(act==='close'||(t&&t.classList&&t.classList.contains('rm-backdrop'))){closeExcludePicker();return;}
      if(act==='refresh'){post(EV_RUN_REQ,{mode:readMode(),metaOnly:true});return;}
      if(act==='all'){var cbs=el.querySelectorAll('input[type="checkbox"]');for(var i=0;i<cbs.length;i++)cbs[i].checked=true;return;}
      if(act==='none'){var cbs2=el.querySelectorAll('input[type="checkbox"]');for(var j=0;j<cbs2.length;j++)cbs2[j].checked=false;return;}
      if(act==='apply'){applyExcludePicker();return;}
    });
    el.addEventListener('input',function(e){
      var t=e.target;if(t&&t.dataset&&t.dataset.bind==='q'){renderExcludePicker();}
    });
  }
  exPickerEl=el;
}
function openExcludePicker(){
  ensureExcludePicker();
  exPickerOpen=true;
  exPickerEl.style.display='block';
  exPickerEl.classList.add('is-open');
  var meta=loadMeta();
  var fams=Array.isArray(meta.modelFamilies)?meta.modelFamilies:[];
  var cats=Array.isArray(meta.systemCategories)?meta.systemCategories:[];
  if(!fams.length&&!cats.length){post(EV_RUN_REQ,{mode:readMode(),metaOnly:true});}
  renderExcludePicker();
}
function closeExcludePicker(){
  if(!exPickerEl)return;
  exPickerOpen=false;
  exPickerEl.classList.remove('is-open');
  exPickerEl.style.display='none';
}
function renderExcludePicker(){
  if(!exPickerEl)return;
  var cfg=loadRuleConfig();
  var famSel={},catSel={};
  (cfg.excludeFamilies||[]).forEach(function(x){famSel[String(x)]=1;});
  (cfg.excludeCategories||[]).forEach(function(x){catSel[String(x)]=1;});
  var meta=loadMeta();
  var fams=uniq(Array.isArray(meta.modelFamilies)?meta.modelFamilies:[]);
  var cats=uniq(Array.isArray(meta.systemCategories)?meta.systemCategories:[]);
  fams.sort(function(a,b){return String(a).localeCompare(String(b));});
  cats.sort(function(a,b){return String(a).localeCompare(String(b));});
  var qEl=exPickerEl.querySelector('[data-bind="q"]');
  var q=qEl?String(qEl.value||'').trim().toLowerCase():'';
  function mk(kind,val,checked){
    return '<label class="exfam-item"><input type="checkbox" data-kind="'+kind+'" value="'+esc(val)+'" '+(checked?'checked':'')+'/><span class="exfam-text">'+esc(val)+'</span></label>';
  }
  var famBox=exPickerEl.querySelector('[data-slot="fam"]');
  var catBox=exPickerEl.querySelector('[data-slot="cat"]');
  if(!famBox||!catBox)return;
  var fh=[];for(var i=0;i<fams.length;i++){var f=String(fams[i]||'');if(!f)continue;if(q&&f.toLowerCase().indexOf(q)<0)continue;fh.push(mk('fam',f,!!famSel[f]));}
  famBox.innerHTML = fh.length?fh.join(''):'<div class="exfam-empty">목록이 비어있습니다. “목록 새로고침”을 누르세요.</div>';
  var ch=[];for(var j=0;j<cats.length;j++){var c=String(cats[j]||'');if(!c)continue;if(q&&c.toLowerCase().indexOf(q)<0)continue;ch.push(mk('cat',c,!!catSel[c]));}
  catBox.innerHTML = ch.length?ch.join(''):'<div class="exfam-empty">목록이 비어있습니다. “목록 새로고침”을 누르세요.</div>';
}
function applyExcludePicker(){
  if(!exPickerEl)return;
  var cfg=loadRuleConfig();
  var fam=[],cat=[];
  var cbs=exPickerEl.querySelectorAll('input[type="checkbox"]');
  for(var i=0;i<cbs.length;i++){
    var cb=cbs[i];
    if(!cb.checked)continue;
    var kind=cb.dataset?cb.dataset.kind:'';
    if(kind==='fam')fam.push(String(cb.value||''));
    else if(kind==='cat')cat.push(String(cb.value||''));
  }
  cfg.excludeFamilies=fam;
  cfg.excludeCategories=cat;
  saveRuleConfig(cfg);
  toast('Exclude 목록이 적용되었습니다.','ok',1200);
  closeExcludePicker();
}

export function renderDup(root){
  ensureStyles();
  ensureExcludePicker();
  var target = root || document.getElementById('view-root') || document.getElementById('app');
  if(!target)return;
  clear(target);

  var mode = readMode();
  var scopeMode = readScopeMode();
  var busy=false, exporting=false;
  var rows=[], groups=[], deleted={}, expanded={}, lastResult=null, lastPairs=[];

  var page=div('dup-page feature-shell');
  var header=div('feature-header dup-toolbar');
  var heading=div('feature-heading');
  var actions=div('feature-actions');

  var runBtn=mkBtn('검토 시작','card-action-btn');
  var exportBtn=mkBtn('엑셀 내보내기','card-action-btn'); exportBtn.disabled=true;
  var btnDup=mkBtn('중복','control-chip chip-btn subtle');
  var btnClash=mkBtn('자체간섭','control-chip chip-btn subtle');
  var settingsBtn=mkBtn('규칙/Set','control-chip chip-btn subtle');

  var summaryBar=div('dup-summarybar sticky hidden');
  var body=div('dup-body');

  actions.append(runBtn, btnDup, btnClash, settingsBtn, exportBtn);
  header.append(heading, actions);
  page.append(header, summaryBar, body);
  target.append(page);

  // rule panel
  var rulePanel = div('dup-rulemodal');
  rulePanel.innerHTML =
    '<div class="rm-backdrop" data-act="close"></div>' +
    '<div class="rm-window">' +
      '<div class="rm-head">' +
        '<div class="rm-head-left"><div class="rm-title">규칙 / Set 설정</div><div class="rm-sub">Set / Pair / Exclude / 허용오차 / Import·Export</div></div>' +
        '<div class="rm-actions">' +
          '<button class="rp-btn rp-btn--ghost" data-act="refresh">목록 새로고침</button>' +
          '<button class="rp-btn rp-btn--ghost" data-act="export">Export</button>' +
          '<button class="rp-btn rp-btn--ghost" data-act="import">Import</button>' +
          '<button class="rp-btn" data-act="close">닫기</button>' +
        '</div>' +
      '</div>' +
      '<div class="rm-body">' +
        '<section class="rp-sec">' +
          '<div class="rp-sec-title">공통 설정</div>' +
          '<div class="rp-grid">' +
            '<div class="rp-label">허용오차(mm)</div><input class="rp-input" type="number" step="0.1" min="0.01" data-bind="tolMm"/>' +
            '<div class="rp-label">범위(Selection)</div><select class="rp-select" data-bind="scopeMode"><option value="all">전체</option><option value="scope">선택한 요소만 검사</option><option value="exclude">선택한 요소는 제외</option></select>' +
            '<div class="rp-label">제외 키워드</div><input class="rp-input" type="text" placeholder="예: Dummy, Temp (콤마)" data-bind="excludeKeywords"/>' +
          '</div>' +
          '<div class="rp-hint">※ 파라미터/패밀리/시스템 목록은 “목록 새로고침” 후 채워집니다.</div>' +
        '</section>' +

        '<section class="rp-sec">' +
          '<div class="rp-sec-title">Set 정의</div>' +
          '<div class="rp-hint">Set = (OR 그룹) - (AND 조건). 각 조건은 Category/Family/Type/Name/Parameter 필터입니다.</div>' +
          '<div data-slot="sets"></div>' +
          '<button class="rp-btn rp-btn--add" data-act="add-set">+ Set 추가</button>' +
        '</section>' +

        '<section class="rp-sec">' +
          '<div class="rp-sec-title">Pair (Set vs Set)</div>' +
          '<div class="rp-hint">A vs B: A에 속한 요소 ↔ B에 속한 요소 간 간섭만 계산합니다.</div>' +
          '<div data-slot="pairs"></div>' +
          '<button class="rp-btn rp-btn--add" data-act="add-pair">+ Pair 추가</button>' +
        '</section>' +

        '<section class="rp-sec">' +
          '<div class="rp-sec-title">Exclude Sets</div>' +
          '<div class="rp-hint">등록된 Set에 포함되는 요소는 모든 검토(중복/간섭)에서 제외합니다.</div>' +
          '<div data-slot="exsets"></div>' +
          '<button class="rp-btn rp-btn--add" data-act="add-exset">+ Exclude Set 추가</button>' +
        '</section>' +

        '<section class="rp-sec">' +
          '<div class="rp-sec-title">Exclude 목록(패밀리/시스템)</div>' +
          '<div class="rp-grid">' +
            '<div class="rp-label">선택</div><div><button class="rp-btn rp-btn--add" data-act="open-expicker">패밀리/시스템 목록 열기…</button></div>' +
            '<div class="rp-label">현재 제외</div><div class="rp-hint" data-slot="exsum">—</div>' +
          '</div>' +
        '</section>' +

        '<section class="rp-sec">' +
          '<div class="rp-sec-title">Import / Export JSON</div>' +
          '<textarea class="rp-textarea" data-bind="json"></textarea>' +
        '</section>' +

        '<div class="rp-foot"><button class="rp-btn rp-btn--primary" data-act="apply">적용</button></div>' +
      '</div>' +
    '</div>';
  page.append(rulePanel);


  function syncModeUI(){
    btnDup.classList.toggle('is-active', mode==='duplicate');
    btnClash.classList.toggle('is-active', mode==='clash');
    var title = (mode==='clash')?'자체간섭 검토':'중복검토';
    var sub = (mode==='clash')?'같은 파일 내 자체간섭 후보를 A↔B 쌍으로 표시합니다.':'중복 요소 후보를 그룹별로 확인하고 삭제/되돌리기를 관리합니다.';
    heading.innerHTML = '<span class="feature-kicker">Duplicate Inspector</span><h2 class="feature-title">'+esc(title)+'</h2><p class="feature-sub">'+esc(sub)+'</p>';
  }
  syncModeUI();
  renderIntro();

  function setLoading(on){
    busy=!!on;
    runBtn.disabled=busy;
    btnDup.disabled=busy;
    btnClash.disabled=busy;
    settingsBtn.disabled=busy;
    exportBtn.disabled = busy || exporting || (!rows.length && !lastPairs.length);
    runBtn.textContent = busy?'검토 중…':'검토 시작';
  }

  onHost('dup:meta', function(payload){
    try{
      var meta=isObj(payload)?payload:{};
      try{localStorage.setItem(LS_META, JSON.stringify(meta));}catch(e){}
      toast('목록을 갱신했습니다.','ok',1200);
      if(exPickerOpen)renderExcludePicker();
      if(rulePanel && rulePanel.classList && rulePanel.classList.contains('is-open')){ try{ renderRulePanel(); }catch(e){} }
    }catch(e){}
  });

  onHost(function(msg){
    var ev = msg && msg.ev ? msg.ev : '';
    var payload = msg ? msg.payload : null;

    if (RESP_ROWS_EVENTS.indexOf(ev) >= 0){
      setLoading(false);
      var list = (payload && Array.isArray(payload.rows)) ? payload.rows :
                 (payload && Array.isArray(payload.data)) ? payload.data :
                 (Array.isArray(payload) ? payload : []);
      handleRows(list);
      return;
    }

    if (ev === 'dup:result'){
      setLoading(false);
      lastResult = isObj(payload)?payload:null;
      paint();
      refreshSummary();
      return;
    }

    if (ev === 'dup:pairs'){
      setLoading(false);
      lastPairs = Array.isArray(payload)?payload:[];
      paint();
      refreshSummary();
      return;
    }

    if (ev === EV_DELETED_ONE){
      var id = safeId(payload && payload.id);
      if(id){deleted[id]=1; paintRowStates();}
      return;
    }
    if (ev === EV_RESTORED_ONE){
      var id2 = safeId(payload && payload.id);
      if(id2){delete deleted[id2]; paintRowStates();}
      return;
    }

    if (ev === EV_EXPORTED){
      exporting=false; ProgressDialog.hide();
      var path = asStr(payload && payload.path,'');
      if(path){
        showExcelSavedDialog('엑셀로 내보냈습니다.', path, function(p){ if(p) post('excel:open',{path:p}); });
      }else{
        toast(asStr(payload && payload.message,'엑셀 내보내기 실패'),'err',2800);
      }
      exportBtn.disabled = busy || (!rows.length && !lastPairs.length);
      return;
    }

    if (ev === 'dup:progress'){
      handleExcelProgress(payload||{});
      return;
    }

    if (ev === 'host:warn'){
      setLoading(false);
      var m = asStr(payload && payload.message,'');
      if(m) toast(m,'warn',3200);
      return;
    }
    if (ev === 'host:error' || ev === 'revit:error'){
      setLoading(false);
      exporting=false; ProgressDialog.hide();
      toast(asStr(payload && payload.message,'오류가 발생했습니다.'),'err',3400);
      return;
    }
  });

  runBtn.addEventListener('click', function(){
    if(busy)return;
    rows=[];groups=[];deleted={};expanded={};lastPairs=[];lastResult=null;
    body.innerHTML=''; body.append(buildSkeleton(6));
    exportBtn.disabled=true;
    setLoading(true);

    var cfg=loadRuleConfig();
    var kw=getExcludeKeywords();
    var merged=uniq([].concat(kw, cfg.excludeFamilies||[], cfg.excludeCategories||[]));
    post(EV_RUN_REQ, {mode:mode, tolFeet:readTolFeet(), scopeMode:scopeMode, excludeKeywords:merged, ruleConfig:cfg});

    var t=setTimeout(function(){
      if(!busy)return;
      setLoading(false);
      body.innerHTML='';
      renderIntro();
      toast('응답이 없습니다. Add-in 이벤트/런타임을 확인하세요.','err',3200);
    }, 12000);

    var off = onHost(function(m){
      var e = m && m.ev ? m.ev : '';
      if(RESP_ROWS_EVENTS.indexOf(e)>=0 || e==='dup:result' || e==='dup:pairs' || e==='host:error' || e==='revit:error'){
        try{clearTimeout(t);}catch(err){}
        try{off();}catch(err2){}
      }
    });
  });

  exportBtn.addEventListener('click', function(){
    if(exporting)return;
    if(!rows.length && !lastPairs.length)return;
    exporting=true;
    exportBtn.disabled=true;
    chooseExcelMode(function(excelMode){
      post(EV_EXPORT_REQ, {excelMode: excelMode || 'fast'});
    });
  });

  btnDup.addEventListener('click', function(){
    if(busy)return;
    mode='duplicate'; try{localStorage.setItem(LS_MODE,mode);}catch(e){}
    syncModeUI();
    rows=[];groups=[];deleted={};expanded={};lastPairs=[];lastResult=null;
    body.innerHTML=''; renderIntro(); refreshSummary(); exportBtn.disabled=true;
  });
  btnClash.addEventListener('click', function(){
    if(busy)return;
    mode='clash'; try{localStorage.setItem(LS_MODE,mode);}catch(e){}
    syncModeUI();
    rows=[];groups=[];deleted={};expanded={};lastPairs=[];lastResult=null;
    body.innerHTML=''; renderIntro(); refreshSummary(); exportBtn.disabled=true;
  });

  settingsBtn.addEventListener('click', function(){
    var open = !rulePanel.classList.contains('is-open');
    rulePanel.classList.toggle('is-open', open);
    settingsBtn.classList.toggle('is-active', open);
    if(open){
      try{ renderRulePanel(); }catch(e){}
      var tolEl=rulePanel.querySelector('[data-bind="tolMm"]'); if(tolEl) tolEl.value=String(readTolMm());
      var scEl=rulePanel.querySelector('[data-bind="scopeMode"]'); if(scEl) scEl.value=scopeMode;
      var kwEl=rulePanel.querySelector('[data-bind="excludeKeywords"]'); if(kwEl){ try{kwEl.value=String(localStorage.getItem(LS_EXCL_KW)||'');}catch(e){kwEl.value='';}}
    }
  });

  rulePanel.addEventListener('click', function(e){
    var t=e.target; var act=(t&&t.dataset)?t.dataset.act:'';
    if(act==='close' || (t&&t.classList&&t.classList.contains('rm-backdrop'))){
      rulePanel.classList.remove('is-open');
      settingsBtn.classList.remove('is-active');
      return;
    }

    if(act==='refresh'){ post(EV_RUN_REQ,{mode:readMode(),metaOnly:true}); return; }
    if(act==='open-expicker'){ openExcludePicker(); return; }

    if(act==='export'){ exportRuleJson(); return; }
    if(act==='import'){ importRuleJson(); return; }
    if(act==='apply'){ applyRulePanel(); return; }

    if(act==='add-set'){ addSet(); renderRulePanel(); return; }
    if(act==='del-set'){ delSet(asNum(t.dataset.si,-1)); renderRulePanel(); return; }
    if(act==='add-group'){ addGroup(asNum(t.dataset.si,-1)); renderRulePanel(); return; }
    if(act==='del-group'){ delGroup(asNum(t.dataset.si,-1), asNum(t.dataset.gi,-1)); renderRulePanel(); return; }
    if(act==='add-clause'){ addClause(asNum(t.dataset.si,-1), asNum(t.dataset.gi,-1)); renderRulePanel(); return; }
    if(act==='del-clause'){ delClause(asNum(t.dataset.si,-1), asNum(t.dataset.gi,-1), asNum(t.dataset.ci,-1)); renderRulePanel(); return; }

    if(act==='add-pair'){ addPair(); renderRulePanel(); return; }
    if(act==='del-pair'){ delPair(asNum(t.dataset.pi,-1)); renderRulePanel(); return; }

    if(act==='add-exset'){ addExcludeSet(); renderRulePanel(); return; }
    if(act==='del-exset'){ delExcludeSet(asNum(t.dataset.ei,-1)); renderRulePanel(); return; }
  });

  rulePanel.addEventListener('change', function(e){
    var t=e.target;
    if(!t || !t.dataset) return;

    // globals
    if(t.dataset.bind==='tolMm'){
      var mm = clamp(asNum(t.value, DEFAULT_TOL_MM), 0.01, 1000);
      try{localStorage.setItem(LS_TOL_MM, String(mm));}catch(ex){}
      return;
    }
    if(t.dataset.bind==='scopeMode'){
      scopeMode = asStr(t.value,'all');
      try{localStorage.setItem(LS_SCOPE, scopeMode);}catch(ex2){}
      return;
    }
    if(t.dataset.bind==='excludeKeywords'){
      try{localStorage.setItem(LS_EXCL_KW, String(t.value||''));}catch(ex3){}
      return;
    }

    // set name
    if(t.dataset.setName){
      var si = asNum(t.dataset.setName,-1);
      var cfg = loadRuleConfig();
      if(cfg.sets && cfg.sets[si]){ cfg.sets[si].name = String(t.value||''); saveRuleConfig(cfg); }
      return;
    }
    if(t.dataset.setLogic){
      var si2 = asNum(t.dataset.setLogic,-1);
      var cfg2 = loadRuleConfig();
      if(cfg2.sets && cfg2.sets[si2]){ cfg2.sets[si2].logic = (t.value==='and'?'and':'or'); saveRuleConfig(cfg2); }
      return;
    }

    // clause edits
    if(t.dataset.clField){
      var p = String(t.dataset.clField).split(':');
      var si3=asNum(p[0],-1), gi3=asNum(p[1],-1), ci3=asNum(p[2],-1);
      var cfg3 = loadRuleConfig();
      var cl = cfg3.sets && cfg3.sets[si3] && cfg3.sets[si3].groups && cfg3.sets[si3].groups[gi3] && cfg3.sets[si3].groups[gi3].clauses && cfg3.sets[si3].groups[gi3].clauses[ci3];
      if(cl){ cl.field = String(t.value||'category'); saveRuleConfig(cfg3); renderRulePanel(); }
      return;
    }
    if(t.dataset.clOp){
      var p2 = String(t.dataset.clOp).split(':');
      var si4=asNum(p2[0],-1), gi4=asNum(p2[1],-1), ci4=asNum(p2[2],-1);
      var cfg4 = loadRuleConfig();
      var cl2 = cfg4.sets && cfg4.sets[si4] && cfg4.sets[si4].groups && cfg4.sets[si4].groups[gi4] && cfg4.sets[si4].groups[gi4].clauses && cfg4.sets[si4].groups[gi4].clauses[ci4];
      if(cl2){ cl2.op = String(t.value||'contains'); saveRuleConfig(cfg4); }
      return;
    }
    if(t.dataset.clParam){
      var p3 = String(t.dataset.clParam).split(':');
      var si5=asNum(p3[0],-1), gi5=asNum(p3[1],-1), ci5=asNum(p3[2],-1);
      var cfg5 = loadRuleConfig();
      var cl3 = cfg5.sets && cfg5.sets[si5] && cfg5.sets[si5].groups && cfg5.sets[si5].groups[gi5] && cfg5.sets[si5].groups[gi5].clauses && cfg5.sets[si5].groups[gi5].clauses[ci5];
      if(cl3){ cl3.param = String(t.value||''); saveRuleConfig(cfg5); }
      return;
    }
    if(t.dataset.clVal){
      var p4 = String(t.dataset.clVal).split(':');
      var si6=asNum(p4[0],-1), gi6=asNum(p4[1],-1), ci6=asNum(p4[2],-1);
      var cfg6 = loadRuleConfig();
      var cl4 = cfg6.sets && cfg6.sets[si6] && cfg6.sets[si6].groups && cfg6.sets[si6].groups[gi6] && cfg6.sets[si6].groups[gi6].clauses && cfg6.sets[si6].groups[gi6].clauses[ci6];
      if(cl4){ cl4.value = String(t.value||''); saveRuleConfig(cfg6); }
      return;
    }

    // pairs
    if(t.dataset.pA){
      var pi=asNum(t.dataset.pA,-1);
      var cfg7 = loadRuleConfig();
      if(cfg7.pairs && cfg7.pairs[pi]){ cfg7.pairs[pi].a = String(t.value||''); saveRuleConfig(cfg7); }
      return;
    }
    if(t.dataset.pB){
      var pi2=asNum(t.dataset.pB,-1);
      var cfg8 = loadRuleConfig();
      if(cfg8.pairs && cfg8.pairs[pi2]){ cfg8.pairs[pi2].b = String(t.value||''); saveRuleConfig(cfg8); }
      return;
    }
    if(t.dataset.pEn){
      var pi3=asNum(t.dataset.pEn,-1);
      var cfg9 = loadRuleConfig();
      if(cfg9.pairs && cfg9.pairs[pi3]){ cfg9.pairs[pi3].enabled = !!t.checked; saveRuleConfig(cfg9); }
      return;
    }

    // exclude sets
    if(t.dataset.exSet){
      var ei=asNum(t.dataset.exSet,-1);
      var cfg10 = loadRuleConfig();
      if(cfg10.excludeSetIds && ei>=0 && ei<cfg10.excludeSetIds.length){ cfg10.excludeSetIds[ei] = String(t.value||''); saveRuleConfig(cfg10); }
      return;
    }
  });

  rulePanel.addEventListener('input', function(e){
    var t=e.target;
    if(!t || !t.dataset) return;

    if(t.dataset.bind==='excludeKeywords'){
      try{localStorage.setItem(LS_EXCL_KW, String(t.value||''));}catch(ex){}
      return;
    }
    if(t.dataset.setName){
      var si = asNum(t.dataset.setName,-1);
      var cfg = loadRuleConfig();
      if(cfg.sets && cfg.sets[si]){ cfg.sets[si].name = String(t.value||''); saveRuleConfig(cfg); }
      return;
    }
    if(t.dataset.clVal){
      var p = String(t.dataset.clVal).split(':');
      var si2=asNum(p[0],-1), gi2=asNum(p[1],-1), ci2=asNum(p[2],-1);
      var cfg2 = loadRuleConfig();
      var cl = cfg2.sets && cfg2.sets[si2] && cfg2.sets[si2].groups && cfg2.sets[si2].groups[gi2] && cfg2.sets[si2].groups[gi2].clauses && cfg2.sets[si2].groups[gi2].clauses[ci2];
      if(cl){ cl.value = String(t.value||''); saveRuleConfig(cfg2); }
      return;
    }
  });

  function ensureId(prefix){
    return String(prefix||'S') + Math.random().toString(16).slice(2,10);
  }

  function getMetaParams(){
    var meta = loadMeta();
    var ps = (meta && Array.isArray(meta.parameters)) ? meta.parameters : [];
    ps = uniq(ps);
    ps.sort(function(a,b){ return String(a).localeCompare(String(b)); });
    return ps;
  }

  function optHtml(list, sel){
    var out='';
    for(var i=0;i<list.length;i++){
      var v=list[i].v, t=list[i].t;
      out += '<option value="'+esc(v)+'"'+(String(sel)===String(v)?' selected':'')+'>'+esc(t)+'</option>';
    }
    return out;
  }
  function optHtmlStr(list, sel){
    var out='';
    for(var i=0;i<list.length;i++){
      var v=String(list[i]||'');
      if(!v) continue;
      out += '<option value="'+esc(v)+'"'+(String(sel)===String(v)?' selected':'')+'>'+esc(v)+'</option>';
    }
    return out;
  }

  function renderRulePanel(){
    try{
      var cfg = loadRuleConfig();
      var tolEl=rulePanel.querySelector('[data-bind="tolMm"]'); if(tolEl) tolEl.value = String(readTolMm());
      var scEl=rulePanel.querySelector('[data-bind="scopeMode"]'); if(scEl) scEl.value = scopeMode;
      var kwEl=rulePanel.querySelector('[data-bind="excludeKeywords"]'); if(kwEl){ try{kwEl.value=String(localStorage.getItem(LS_EXCL_KW)||'');}catch(ex){kwEl.value='';} }

      var exSum=rulePanel.querySelector('[data-slot="exsum"]');
      if(exSum){
        exSum.textContent = '패밀리 ' + (cfg.excludeFamilies?cfg.excludeFamilies.length:0) + '개, 시스템 ' + (cfg.excludeCategories?cfg.excludeCategories.length:0) + '개';
      }

      var params = getMetaParams();

      // sets
      var setsSlot = rulePanel.querySelector('[data-slot="sets"]');
      if(setsSlot){
        var html='';
        for(var si=0; si<(cfg.sets||[]).length; si++){
          var s = cfg.sets[si] || {};
          var logic = (s.logic==='and')?'and':'or';
          html += '<div class="rp-card">' +
            '<div class="rp-card-head">' +
              '<input class="rp-input rp-input--sm" data-set-name="'+si+'" value="'+esc(s.name||'')+'"/>' +
              '<button class="rp-x" data-act="del-set" data-si="'+si+'">×</button>' +
              '<div class="rp-card-sub">id: '+esc(s.id||'')+'</div>' +
            '</div>' +
            '<div class="rp-row">' +
              '<div class="rp-mini">그룹 결합</div>' +
              '<select class="rp-select rp-select--sm" data-set-logic="'+si+'">' +
                '<option value="or"'+(logic==='or'?' selected':'')+'>OR</option>' +
                '<option value="and"'+(logic==='and'?' selected':'')+'>AND</option>' +
              '</select>' +
              '<button class="rp-btn rp-btn--tiny" data-act="add-group" data-si="'+si+'">+ OR 그룹</button>' +
            '</div>';

          var groups = (s.groups && Array.isArray(s.groups)) ? s.groups : [];
          for(var gi=0; gi<groups.length; gi++){
            var g = groups[gi] || {};
            var clauses = (g.clauses && Array.isArray(g.clauses)) ? g.clauses : [];
            html += '<div class="rp-group">' +
              '<div class="rp-group-head">' +
                '<div class="rp-group-title">그룹 '+(gi+1)+' (AND)</div>' +
                '<button class="rp-x" data-act="del-group" data-si="'+si+'" data-gi="'+gi+'">×</button>' +
              '</div>';

            for(var ci=0; ci<clauses.length; ci++){
              var cl = clauses[ci] || {};
              var field = cl.field || 'category';
              var op = cl.op || 'contains';
              var val = cl.value || '';
              var param = cl.param || '';
              var paramHidden = (field==='param') ? '' : ' is-hidden';
              html += '<div class="rp-clause">' +
                '<select class="rp-select rp-select--sm" data-cl-field="'+si+':'+gi+':'+ci+'">' +
                  optHtml([
                    {v:'category',t:'Category'},{v:'family',t:'Family'},{v:'type',t:'Type'},{v:'name',t:'Name'},{v:'param',t:'Parameter'}
                  ], field) +
                '</select>' +
                '<select class="rp-select rp-select--sm" data-cl-op="'+si+':'+gi+':'+ci+'">' +
                  optHtml([
                    {v:'contains',t:'Contains'},{v:'equals',t:'Equal'},{v:'startswith',t:'StartsWith'},{v:'endswith',t:'EndsWith'},{v:'notcontains',t:'NotContains'},{v:'notequals',t:'NotEqual'}
                  ], op) +
                '</select>' +
                '<select class="rp-select rp-select--sm rp-param'+paramHidden+'" data-cl-param="'+si+':'+gi+':'+ci+'">' +
                  '<option value=""'+(param===''?' selected':'')+'>(param)</option>' + optHtmlStr(params, param) +
                '</select>' +
                '<input class="rp-input rp-input--sm" data-cl-val="'+si+':'+gi+':'+ci+'" value="'+esc(val)+'"/>' +
                '<button class="rp-x" data-act="del-clause" data-si="'+si+'" data-gi="'+gi+'" data-ci="'+ci+'">×</button>' +
              '</div>';
            }

            html += '<button class="rp-btn rp-btn--tiny" data-act="add-clause" data-si="'+si+'" data-gi="'+gi+'">+ 조건</button>';
            html += '</div>';
          }

          html += '</div>';
        }
        if(!(cfg.sets||[]).length){
          html = '<div class="rp-hint">Set이 없습니다. + Set 추가를 눌러 생성하세요.</div>';
        }
        setsSlot.innerHTML = html;
      }

      // pairs
      var pairsSlot = rulePanel.querySelector('[data-slot="pairs"]');
      if(pairsSlot){
        var setOpts = [{v:'__ALL__',t:'__ALL__ (전체)'}];
        for(var i=0;i<(cfg.sets||[]).length;i++){
          var ss = cfg.sets[i]||{};
          setOpts.push({v:ss.id||'', t:(ss.name||ss.id||'')});
        }
        var html2='';
        for(var pi=0; pi<(cfg.pairs||[]).length; pi++){
          var p = cfg.pairs[pi] || {};
          html2 += '<div class="rp-pair">' +
            '<select class="rp-select rp-select--sm" data-p-a="'+pi+'">' + optHtml(setOpts, p.a||'__ALL__') + '</select>' +
            '<span class="rp-vs">vs</span>' +
            '<select class="rp-select rp-select--sm" data-p-b="'+pi+'">' + optHtml(setOpts, p.b||'__ALL__') + '</select>' +
            '<label class="rp-chk"><input type="checkbox" data-p-en="'+pi+'"'+(p.enabled!==false?' checked':'')+'/> 사용</label>' +
            '<button class="rp-x" data-act="del-pair" data-pi="'+pi+'">×</button>' +
          '</div>';
        }
        if(!(cfg.pairs||[]).length) html2 = '<div class="rp-hint">Pair가 없습니다. + Pair 추가를 눌러 생성하세요.</div>';
        pairsSlot.innerHTML = html2;
      }

      // exsets
      var exSlot = rulePanel.querySelector('[data-slot="exsets"]');
      if(exSlot){
        var html3='';
        if(!(cfg.sets||[]).length){
          html3 = '<div class="rp-hint">Set이 없으면 Exclude Set을 등록할 수 없습니다.</div>';
        } else {
          for(var ei=0; ei<(cfg.excludeSetIds||[]).length; ei++){
            var sid = cfg.excludeSetIds[ei] || '';
            html3 += '<div class="rp-exrow">' +
              '<select class="rp-select rp-select--sm" data-ex-set="'+ei+'">' + optHtmlStr((cfg.sets||[]).map(function(x){return x.id;}), sid) + '</select>' +
              '<button class="rp-x" data-act="del-exset" data-ei="'+ei+'">×</button>' +
            '</div>';
          }
          if(!(cfg.excludeSetIds||[]).length) html3 = '<div class="rp-hint">등록된 Exclude Set이 없습니다.</div>';
        }
        exSlot.innerHTML = html3;
      }
    }catch(e){}
  }

  function exportRuleJson(){
    var ta = rulePanel.querySelector('[data-bind="json"]');
    if(!ta) return;
    ta.value = JSON.stringify(loadRuleConfig(), null, 2);
    toast('Export JSON 생성','ok',1200);
  }
  function importRuleJson(){
    var ta = rulePanel.querySelector('[data-bind="json"]');
    if(!ta) return;
    var cfg = null;
    try{ cfg = JSON.parse(String(ta.value||'').trim()); }catch(e){ cfg = null; }
    if(!cfg){ toast('Import JSON 형식이 올바르지 않습니다.','err',2400); return; }
    saveRuleConfig(cfg);
    renderRulePanel();
    toast('Import 적용 완료','ok',1200);
  }
  function applyRulePanel(){
    var tolEl=rulePanel.querySelector('[data-bind="tolMm"]');
    var scEl=rulePanel.querySelector('[data-bind="scopeMode"]');
    var kwEl=rulePanel.querySelector('[data-bind="excludeKeywords"]');

    var mm = clamp(asNum(tolEl?tolEl.value:DEFAULT_TOL_MM, DEFAULT_TOL_MM), 0.01, 1000);
    try{localStorage.setItem(LS_TOL_MM, String(mm));}catch(e){}
    scopeMode = asStr(scEl?scEl.value:'all','all');
    try{localStorage.setItem(LS_SCOPE, scopeMode);}catch(e2){}
    try{localStorage.setItem(LS_EXCL_KW, String(kwEl?kwEl.value:''));}catch(e3){}
    toast('설정 적용 완료','ok',1200);
  }

  function addSet(){
    var cfg = loadRuleConfig();
    cfg.sets = cfg.sets || [];
    cfg.sets.push({
      id: ensureId('S'),
      name: 'Set',
      logic: 'or',
      groups: [{ clauses: [{ field:'category', op:'contains', value:'', param:'' }] }]
    });
    saveRuleConfig(cfg);
    toast('Set가 추가되었습니다.','ok',1200);
  }
  function delSet(si){
    var cfg = loadRuleConfig();
    if(!cfg.sets || si<0 || si>=cfg.sets.length) return;
    var removedId = cfg.sets[si].id;
    cfg.sets.splice(si,1);
    cfg.excludeSetIds = (cfg.excludeSetIds||[]).filter(function(x){ return String(x)!==String(removedId); });
    cfg.pairs = (cfg.pairs||[]).filter(function(p){ return String(p.a)!==String(removedId) && String(p.b)!==String(removedId); });
    saveRuleConfig(cfg);
  }
  function addGroup(si){
    var cfg = loadRuleConfig();
    if(!cfg.sets || !cfg.sets[si]) return;
    cfg.sets[si].groups = cfg.sets[si].groups || [];
    cfg.sets[si].groups.push({ clauses: [{ field:'category', op:'contains', value:'', param:'' }] });
    saveRuleConfig(cfg);
  }
  function delGroup(si, gi){
    var cfg = loadRuleConfig();
    if(!cfg.sets || !cfg.sets[si] || !cfg.sets[si].groups || gi<0 || gi>=cfg.sets[si].groups.length) return;
    cfg.sets[si].groups.splice(gi,1);
    saveRuleConfig(cfg);
  }
  function addClause(si, gi){
    var cfg = loadRuleConfig();
    if(!cfg.sets || !cfg.sets[si] || !cfg.sets[si].groups || !cfg.sets[si].groups[gi]) return;
    cfg.sets[si].groups[gi].clauses = cfg.sets[si].groups[gi].clauses || [];
    cfg.sets[si].groups[gi].clauses.push({ field:'category', op:'contains', value:'', param:'' });
    saveRuleConfig(cfg);
  }
  function delClause(si, gi, ci){
    var cfg = loadRuleConfig();
    if(!cfg.sets || !cfg.sets[si] || !cfg.sets[si].groups || !cfg.sets[si].groups[gi] || !cfg.sets[si].groups[gi].clauses) return;
    if(ci<0 || ci>=cfg.sets[si].groups[gi].clauses.length) return;
    cfg.sets[si].groups[gi].clauses.splice(ci,1);
    saveRuleConfig(cfg);
  }

  function addPair(){
    var cfg = loadRuleConfig();
    cfg.pairs = cfg.pairs || [];
    var base = (cfg.sets && cfg.sets.length) ? cfg.sets[0].id : '__ALL__';
    cfg.pairs.push({ a: base, b: base, enabled: true });
    saveRuleConfig(cfg);
    toast('Pair가 추가되었습니다.','ok',1200);
  }
  function delPair(pi){
    var cfg = loadRuleConfig();
    if(!cfg.pairs || pi<0 || pi>=cfg.pairs.length) return;
    cfg.pairs.splice(pi,1);
    saveRuleConfig(cfg);
  }

  function addExcludeSet(){
    var cfg = loadRuleConfig();
    cfg.excludeSetIds = cfg.excludeSetIds || [];
    if(!cfg.sets || !cfg.sets.length){ toast('Set이 없습니다. 먼저 Set을 추가하세요.','warn',1600); return; }
    cfg.excludeSetIds.push(String(cfg.sets[0].id));
    saveRuleConfig(cfg);
    toast('Exclude Set이 추가되었습니다.','ok',1200);
  }
  function delExcludeSet(ei){
    var cfg = loadRuleConfig();
    if(!cfg.excludeSetIds || ei<0 || ei>=cfg.excludeSetIds.length) return;
    cfg.excludeSetIds.splice(ei,1);
    saveRuleConfig(cfg);
  }



  function handleRows(list){
    rows = (Array.isArray(list)?list:[]).map(normalizeRow);
    groups = buildGroups(rows);
    expanded = {};
    for(var i=0;i<groups.length;i++) expanded[groups[i].key]=1;
    exportBtn.disabled = busy || (!rows.length && !lastPairs.length);
    paint(); refreshSummary();
    if(!rows.length){
      body.innerHTML='';
      var empty=div('dup-emptycard');
      empty.innerHTML='<div class="empty-emoji">✅</div><h3 class="empty-title">'+(mode==='clash'?'간섭이 없습니다':'중복이 없습니다')+'</h3><p class="empty-sub">검토 결과가 0건입니다.</p>';
      body.append(empty);
    }
  }

  function paint(){
    body.innerHTML='';
    if(lastResult && lastResult.truncated){
      var info=div('dup-info');
      info.innerHTML='<div class="t">표시 제한</div><div class="s">결과가 많아 상위 '+asNum(lastResult.shown,0)+'건만 표시합니다. 전체('+asNum(lastResult.total,0)+'건)는 엑셀 내보내기에서 확인하세요.</div>';
      body.append(info);
    }
    if(mode==='clash' && lastPairs && lastPairs.length){
      paintPairs(lastPairs);
      return;
    }
    paintGroups(groups);
    paintRowStates();
  }

  function paintPairs(pairs){
    var byGroup={};
    for(var i=0;i<pairs.length;i++){
      var p=pairs[i]||{};
      var gk=asStr(p.groupKey,'C0000');
      if(!byGroup[gk]) byGroup[gk]=[];
      byGroup[gk].push(p);
    }
    var keys=Object.keys(byGroup);
    keys.sort(function(a,b){return byGroup[b].length - byGroup[a].length;});
    for(var gi=0;gi<keys.length;gi++){
      var gk=keys[gi];
      var items=byGroup[gk];
      var card=div('dup-grp');
      var h=div('grp-h');
      var left=div('grp-txt');
      left.innerHTML='<div class="grp-title"><span class="grp-badge">간섭 그룹 '+(gi+1)+'</span><span class="grp-meta">'+esc(gk)+'</span></div><div class="grp-count">'+items.length+'쌍</div>';
      h.append(left, div('grp-actions'));
      card.append(h);
      var list=div('pair-cardlist');
      var seen={};
      for(var j=0;j<items.length;j++){
        var p2=items[j]||{};
        var aId=safeId(p2.aId), bId=safeId(p2.bId);
        var na=asNum(aId,NaN), nb=asNum(bId,NaN);
        var x=(Number.isFinite(na)&&Number.isFinite(nb))?Math.min(na,nb):aId;
        var y=(Number.isFinite(na)&&Number.isFinite(nb))?Math.max(na,nb):bId;
        var k=gk+':'+String(x)+'-'+String(y);
        if(seen[k])continue; seen[k]=1;
        var aInfo=(asStr(p2.aCategory,'')+' · '+asStr(p2.aFamily,'')+(p2.aType?' : '+asStr(p2.aType,''):'')).trim();
        var bInfo=(asStr(p2.bCategory,'')+' · '+asStr(p2.bFamily,'')+(p2.bType?' : '+asStr(p2.bType,''):'')).trim();
        var pc=div('pair-card');
        pc.innerHTML='<div class="pair-side"><div class="pair-lbl">A</div><div class="pair-id"><button class="table-action-btn" data-act="sel" data-id="'+esc(aId)+'">'+esc(aId)+'</button></div><div class="pair-meta">'+esc(aInfo||'—')+'</div></div><div class="pair-mid">↔</div><div class="pair-side"><div class="pair-lbl">B</div><div class="pair-id"><button class="table-action-btn" data-act="sel" data-id="'+esc(bId)+'">'+esc(bId)+'</button></div><div class="pair-meta">'+esc(bInfo||'—')+'</div></div>';
        list.append(pc);
      }
      card.append(list);
      body.append(card);
    }
  }

  function paintGroups(gs){
    for(var i=0;i<gs.length;i++){
      var g=gs[i];
      var card=div('dup-grp');
      var h=div('grp-h');
      var left=div('grp-txt');
      left.innerHTML='<div class="grp-title"><span class="grp-badge">'+(mode==='clash'?'간섭 그룹':'중복 그룹')+' '+(i+1)+'</span><span class="grp-meta">'+esc(buildGroupMeta(g))+'</span></div><div class="grp-count">'+g.rows.length+'개</div>';
      var right=div('grp-actions');
      (function(key){
        var tg=mkBtn(expanded[key]?'접기':'펼치기','control-chip chip-btn subtle');
        tg.addEventListener('click',function(){
          if(expanded[key]) delete expanded[key]; else expanded[key]=1;
          paint(); refreshSummary();
        });
        right.append(tg);
      })(g.key);
      h.append(left,right);
      card.append(h);

      var tbl=div('grp-body');
      var sh=div('dup-subhead');
      sh.append(cell('','ck'), cell('Element ID','th'), cell('Category','th'), cell('Family','th'), cell('Type','th'), cell('연결','th conn'), cell('작업','th right'));
      tbl.append(sh);
      if(expanded[g.key]){
        for(var r=0;r<g.rows.length;r++) tbl.append(renderRow(g.rows[r]));
      }
      card.append(tbl);
      body.append(card);
    }
  }

  function renderRow(r){
    var row=div('dup-row');
    row.dataset.id=r.id;

    var ckCell=cell('', 'ck');
    var ck=document.createElement('input'); ck.type='checkbox'; ck.className='ckbox';
    ckCell.textContent=''; ckCell.append(ck);

    row.append(ckCell);
    row.append(cell(co(r.id,'-'),'td mono right'));
    row.append(cell(co(r.category,'—'),'td'));
    row.append(cell(co(r.family,(r.category?(r.category+' Type'):'—')),'td ell'));
    row.append(cell(co(r.type,'—'),'td ell'));

    var conn=document.createElement('div'); conn.className='conn-cell';
    var ids=Array.isArray(r.connectedIds)?r.connectedIds:[];
    conn.textContent=ids.length?String(ids.length):'0';
    if(ids.length) conn.title=ids.join(', ');
    row.append(cell(conn,'td mono right conn'));

    var act=div('row-actions');
    var zoom=mkBtn('선택/줌','table-action-btn'); zoom.dataset.act='zoom'; zoom.dataset.id=r.id;
    var del=mkBtn(deleted[r.id]?'되돌리기':'삭제','table-action-btn '+(deleted[r.id]?'restore':'table-action-btn--danger'));
    del.dataset.act=deleted[r.id]?'res':'del'; del.dataset.id=r.id;
    act.append(zoom,del);
    row.append(cell(act,'td right'));
    row.classList.toggle('is-deleted', !!deleted[r.id]);
    return row;
  }

  function paintRowStates(){
    var els=body.querySelectorAll('.dup-row');
    for(var i=0;i<els.length;i++){
      var el=els[i];
      var id=el.dataset?el.dataset.id:'';
      var isDel=!!deleted[id];
      el.classList.toggle('is-deleted', isDel);
      var btn=el.querySelector('[data-act="del"],[data-act="res"]');
      if(btn){
        btn.textContent=isDel?'되돌리기':'삭제';
        btn.dataset.act=isDel?'res':'del';
        btn.classList.toggle('restore',isDel);
        btn.classList.toggle('table-action-btn--danger',!isDel);
      }
    }
  }

  function refreshSummary(){
    summaryBar.innerHTML='';
    var gCount=groups.length; var eCount=0;
    for(var i=0;i<groups.length;i++) eCount += groups[i].rows.length;
    var visible = busy || eCount>0 || (mode==='clash' && lastPairs.length>0);
    summaryBar.classList.toggle('hidden', !visible);
    var gLabel=(mode==='clash')?'간섭 그룹':'중복 그룹';
    summaryBar.append(chip(gLabel+' '+gCount), chip('요소 '+eCount));
    if(mode==='clash') summaryBar.append(chip('쌍 '+(lastPairs?lastPairs.length:0)));
  }

  function buildGroupMeta(g){
    var cats=uniq((g.rows||[]).map(function(r){return r.category||'—';}));
    var fams=uniq((g.rows||[]).map(function(r){return r.family||(r.category?(r.category+' Type'):'—');}));
    var types=uniq((g.rows||[]).map(function(r){return r.type||'—';}));
    var catOut=(cats.length===1)?cats[0]:'혼합('+cats.length+')';
    var famOut=(fams.length===1)?fams[0]:'혼합('+fams.length+')';
    var typOut=(types.length===1)?types[0]:'혼합('+types.length+')';
    return catOut+' · '+famOut+' · '+typOut;
  }

  function normalizeRow(r){
    var id=safeId(co(r.elementId, co(r.ElementId, co(r.id, r.Id))));
    var category=asStr(co(r.category, r.Category),'');
    var family=asStr(co(r.family, r.Family),'');
    var type=asStr(co(r.type, r.Type),'');
    var groupKey=asStr(co(r.groupKey, r.GroupKey),'');
    var connRaw=co(r.connectedIds, co(r.ConnectedIds, []));
    var connectedIds=Array.isArray(connRaw)?connRaw.map(String):[];
    return {id:id,category:category,family:family,type:type,groupKey:groupKey,connectedIds:connectedIds};
  }

  function buildGroups(rs){
    var list=Array.isArray(rs)?rs:[];
    var hasKey=list.some(function(x){return !!x.groupKey;});
    if(hasKey){
      var map={};
      for(var i=0;i<list.length;i++){
        var r=list[i]; var k=r.groupKey||'_';
        if(!map[k]) map[k]={key:k,rows:[]};
        map[k].rows.push(r);
      }
      return Object.keys(map).map(function(k){return map[k];});
    }
    var map2={};
    for(var j=0;j<list.length;j++){
      var r2=list[j];
      var sig=[String(r2.id)].concat((r2.connectedIds||[]).map(String)).filter(Boolean).sort().join(',');
      var key=[r2.category||'', r2.family||'', r2.type||'', sig].join('|');
      if(!map2[key]) map2[key]={key:key,rows:[]};
      map2[key].rows.push(r2);
    }
    return Object.keys(map2).map(function(k){return map2[k];});
  }

  function renderIntro(){
    body.innerHTML='';
    var hero=div('dup-hero');
    hero.innerHTML='<h3 class="hero-title">'+(mode==='clash'?'자체간섭 검토를 시작해 보세요':'중복검토를 시작해 보세요')+'</h3><p class="hero-sub">'+(mode==='clash'?'같은 파일 내 자체간섭 후보를 쌍(A↔B)으로 보여줍니다.':'모델의 중복 요소를 그룹으로 묶어 보여줍니다.')+'</p>';
    body.append(hero);
  }

  function handleExcelProgress(p){
    if(!p){ProgressDialog.hide();return;}
    var phase=asStr(p.phase,'').toUpperCase();
    var total=asNum(p.total,0), current=asNum(p.current,0);
    var pct=(total>0)?Math.floor((current/total)*100):asNum(p.percent,0);
    ProgressDialog.show('엑셀 내보내기', asStr(p.message,''));
    ProgressDialog.update(clamp(pct,0,100), asStr(p.message,''), asStr(p.detail,''));
    if(phase==='DONE'||phase==='ERROR'){setTimeout(function(){ProgressDialog.hide();},200);}
  }

  function mkBtn(text, cls){
    var b=document.createElement('button');
    b.type='button';
    b.className=cls||'';
    b.textContent=text;
    return b;
  }
  function chip(text){
    var c=div('chip');
    c.textContent=text;
    return c;
  }
  function cell(content, cls){
    var c=document.createElement('div');
    c.className='cell '+(cls||'');
    if(content && content.nodeType) c.append(content);
    else c.textContent=String(content||'');
    return c;
  }
  function buildSkeleton(n){
    var wrap=div('dup-skeleton');
    var cnt=asNum(n,6);
    for(var i=0;i<cnt;i++){
      var line=div('sk-row');
      line.append(div('sk-chip'), div('sk-id'), div('sk-wide'), div('sk-wide'), div('sk-act'));
      wrap.append(line);
    }
    return wrap;
  }
}

function ensureStyles(){
  if(document.getElementById('dup-style-v60')) return;
  var st=document.createElement('style');
  st.id='dup-style-v60';
  st.textContent = [
    '.control-chip.is-active{background:rgba(76,111,255,.14);border-color:rgba(76,111,255,.45);font-weight:600;}',
    '.dup-row.is-deleted{opacity:.55;}',
    '.row-actions{display:flex;gap:10px;flex-wrap:nowrap;}',
    '.row-actions .table-action-btn{white-space:nowrap;min-width:88px;height:32px;line-height:32px;padding:0 12px;display:inline-flex;align-items:center;justify-content:center;}',
    '.pair-cardlist{padding:10px 12px 14px 12px;display:flex;flex-direction:column;gap:10px;}',
    '.pair-card{display:grid;grid-template-columns:1fr 44px 1fr;gap:10px;align-items:center;border:1px solid rgba(0,0,0,.08);border-radius:14px;background:rgba(255,255,255,.96);padding:10px 12px;}',
    '.pair-mid{text-align:center;opacity:.6;font-size:18px;}',
    '.pair-lbl{font-weight:800;font-size:12px;opacity:.7;margin-bottom:4px;}',
    '.pair-meta{margin-top:6px;font-size:12px;opacity:.85;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}',
    '.conn-cell{display:inline-block;min-width:18px;padding:2px 6px;border-radius:999px;border:1px solid rgba(0,0,0,.12);text-align:center;}',
    '.dup-rulemodal{position:fixed;inset:0;z-index:9000;display:none;}',
    '.dup-rulemodal.is-open{display:block;}',
    '.dup-rulemodal .rm-backdrop{position:absolute;inset:0;background:rgba(0,0,0,.22);}',
    '.dup-rulemodal .rm-window{position:fixed;inset:22px;border-radius:18px;overflow:hidden;border:1px solid rgba(0,0,0,.12);background:#fff;box-shadow:0 20px 60px rgba(0,0,0,.28);display:flex;flex-direction:column;}',
    '.dup-rulemodal .rm-head{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px 16px;border-bottom:1px solid rgba(0,0,0,.12);}',
    '.dup-rulemodal .rm-title{font-weight:800;}',
    '.dup-rulemodal .rm-sub{font-size:12px;opacity:.75;margin-top:2px;}',
    '.dup-rulemodal .rm-actions{display:flex;gap:8px;flex-wrap:wrap;align-items:center;}',
    '.dup-rulemodal .rm-body{padding:14px 16px 70px 16px;overflow:auto;background:rgba(245,246,250,.92);}',
    '.dup-rulemodal .rp-sec{margin:10px 0 14px 0;padding:12px;border-radius:14px;border:1px solid rgba(0,0,0,.12);background:#fff;}',
    '.dup-rulemodal .rp-sec-title{font-weight:800;margin-bottom:8px;}',
    '.dup-rulemodal .rp-grid{display:grid;grid-template-columns:160px 1fr;gap:10px 12px;align-items:center;}',
    '.dup-rulemodal .rp-label{font-size:12px;opacity:.82;}',
    '.dup-rulemodal .rp-hint{font-size:12px;opacity:.8;margin-top:8px;line-height:1.35;}',
    '.dup-rulemodal .rp-foot{position:sticky;bottom:0;background:linear-gradient(180deg,transparent,rgba(245,246,250,.95));padding-top:10px;}',
    '.dup-rulemodal .rp-btn{padding:6px 10px;border-radius:10px;border:1px solid rgba(0,0,0,.12);background:#fff;color:inherit;cursor:pointer;}',
    '.dup-rulemodal .rp-btn--ghost{background:transparent;opacity:.9;}',
    '.dup-rulemodal .rp-btn--add{margin-top:8px;width:100%;background:rgba(76,111,255,.12);}',
    '.dup-rulemodal .rp-btn--primary{background:rgba(76,111,255,.92);color:#fff;font-weight:800;border-color:rgba(76,111,255,.78);}',
    '.dup-rulemodal .rp-input,.dup-rulemodal .rp-select{width:100%;padding:6px 8px;border-radius:10px;border:1px solid rgba(0,0,0,.12);background:#fff;color:inherit;}',
    '.dup-expicker{position:fixed;inset:0;z-index:9100;}',
    '.dup-expicker .rm-backdrop{position:absolute;inset:0;background:rgba(0,0,0,.22);}',
    '.dup-expicker .rm-window{position:fixed;inset:22px;border-radius:18px;overflow:hidden;border:1px solid rgba(0,0,0,.12);background:#fff;box-shadow:0 20px 60px rgba(0,0,0,.28);display:flex;flex-direction:column;}',
    '.dup-expicker .rm-head{position:sticky;top:0;z-index:2;display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px 16px;border-bottom:1px solid rgba(0,0,0,.12);background:#fff;}',
    '.dup-expicker .rm-body{padding:14px 16px;overflow:auto;background:rgba(245,246,250,.92);}',
    '.exfam-wrap{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:10px;}',
    '@media (max-width:980px){.exfam-wrap{grid-template-columns:1fr;}}',
    '.exfam-col{border:1px solid rgba(0,0,0,.12);border-radius:14px;background:#fff;overflow:hidden;}',
    '.exfam-title{padding:10px 12px;font-weight:800;border-bottom:1px solid rgba(0,0,0,.12);background:rgba(76,111,255,.08);}',
    '.exfam-list{padding:10px 12px;max-height:340px;overflow:auto;}',
    '.exfam-item{display:flex;gap:10px;align-items:center;padding:6px 4px;border-radius:10px;}',
    '.exfam-item:hover{background:rgba(76,111,255,.08);}',
    '.exfam-text{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}',
    '.exfam-empty{opacity:.75;font-size:12px;padding:6px 0;}',
    '.exfam-actions{display:flex;gap:8px;margin-top:10px;justify-content:flex-end;}'
    '.rp-row{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:8px;}',
    '.rp-mini{font-size:12px;opacity:.82;min-width:88px;}',
    '.rp-group{padding:8px;border-radius:12px;border:1px dashed rgba(0,0,0,.18);margin-top:10px;background:#fff;}',
    '.rp-group-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;}',
    '.rp-group-title{font-weight:700;font-size:12px;opacity:.9;}',
    '.rp-clause{display:grid;grid-template-columns:1fr 1fr 1.1fr 1.4fr auto;gap:6px;align-items:center;margin-bottom:6px;}',
    '.rp-param.is-hidden{display:none;}',
  ].join('\n');
  document.head.appendChild(st);
}
