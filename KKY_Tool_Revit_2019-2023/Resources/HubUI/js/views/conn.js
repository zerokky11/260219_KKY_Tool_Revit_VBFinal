// Resources/HubUI/js/views/conn.js
// Connector Diagnostics view (fix2 규약 준수)
// - 버튼/이벤트: connector:run / connector:save-excel
// - 단위: 서버는 inch 고정, UI가 mm 선택 시 전송 전에 inch로 변환 / 표시 시 mm 변환
// - ParamName은 표에서 숨김(설정에 존재하므로)
// - UX: 결과영역은 검토 시작 전 숨김 → [검토 시작] 후 안내문 노출 → 데이터 수신 시 필터+표 노출
// - 강조: Status별 톤은 Value1/Value2/Status 셀만 '캡슐형 테두리'로 표시

import { clear, div, tdText, toast, setBusy, showExcelSavedDialog, chooseExcelMode, showCompletionSummaryDialog, closeCompletionSummaryDialog } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { post, onHost } from '../core/bridge.js';

const SKEY = 'kky_conn_opts';
const INCH_TO_MM = 25.4;
const MAX_PREVIEW_ROWS = 150;
const EXCEL_PHASE_WEIGHT = { EXCEL_INIT: 0.05, EXCEL_WRITE: 0.85, EXCEL_SAVE: 0.08, AUTOFIT: 0.02, DONE: 1, ERROR: 1 };
const EXCEL_PHASE_ORDER = ['EXCEL_INIT', 'EXCEL_WRITE', 'EXCEL_SAVE', 'AUTOFIT', 'DONE'];
let lastExcelPct = 0;

/* ---------- 옵션 ---------- */
function loadOpts() {
  const defaults = { tol: 1.0, unit: 'inch', param: 'Comments', reviewParams: ['Comments'], extraParams: '', targetFilter: '', excludeEndDummy: false };
  try {
    return Object.assign({}, defaults, JSON.parse(localStorage.getItem(SKEY) || '{}'));
  } catch { return defaults; }
}
function saveOpts(o) { localStorage.setItem(SKEY, JSON.stringify(o)); }

/* ---------- 단위 ---------- */
const toMm = (inch)=> Number.isFinite(+inch) ? (+inch * INCH_TO_MM) : inch;

/* ---------- Status 매핑 ---------- */
function statusKind(s){
  const t = String(s||'').trim().toLowerCase();
  if (/\b(mis-?match|error|err|fail|invalid|false)\b/.test(t)) return 'bad';
  if (t.includes('연결 필요') || t.includes('연결 대상') || t.includes('shared parameter')) return 'warn';
  if (/\b(warn|warning|minor|check)\b/.test(t)) return 'warn';
  if (/\b(ok|connected|valid|true)\b/.test(t)) return 'ok';
  return 'info';
}

/* ---------- 렌더 ---------- */
export function renderConn(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app'); clear(target);
  const topbar = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar'); if (topbar) topbar.classList.add('hub-topbar');

  const opts = loadOpts();
  const state = {
    rowsInch: [],
    mismatchRows: [],
    mismatchTotal: 0,
    mismatchPreviewCount: 0,
    mismatchHasMore: false,
    notConnectedRows: [],
    notConnectedTotal: 0,
    notConnectedPreviewCount: 0,
    notConnectedHasMore: false,
    hasRun: false,
    tab: 'mismatch',
    totalCount: 0,
    extraParams: [],
    reviewParams: normalizeReviewParams(opts.reviewParams && opts.reviewParams.length ? opts.reviewParams : parseCsvParams(opts.param || 'Comments')),
    paramList: [],
    paramItems: []
  };
  let exporting = false;
  let acceptRunProgress = false;
  let acceptExcelProgress = false;
  state.extraParams = parseExtraParams(opts.extraParams);
  const page = div('conn-page feature-shell');

  const header = div('feature-header');
  const heading = div('feature-heading');
  heading.innerHTML = `
    <span class="feature-kicker">Connector Diagnostics</span>
    <h2 class="feature-title">파라미터 연속성 검토</h2>
    <p class="feature-sub">공유 파라미터 txt 목록에서 검토 대상을 선택하고 파이프/덕트 연결 객체의 값 연속성을 확인합니다.</p>`;

  const run = cardBtn('검토 시작', onRun);
  const save = cardBtn('엑셀 내보내기', onExport);
  run.classList.add('btn-primary');
  save.classList.add('btn-outline');
  save.id = 'btnConnSave';
  save.disabled = true;

  const actions = div('feature-actions');
  actions.append(run, save);
  header.append(heading, actions);
  page.append(header);

  // 설정/작업 (sticky)
  const rowSettings = div('conn-row settings conn-sticky feature-controls');

  const cardSettings = div('conn-card section section-settings');
  const grid = div('conn-grid');
  const targetFilterInput = makeText(opts.targetFilter || '', 'ex) PM1=Value;PM2=Value2');
  const excludeEndDummy = makeCheckbox(opts.excludeEndDummy === true);
  targetFilterInput.title = targetFilterInput.value || targetFilterInput.placeholder || '';

  grid.append(
    kv('허용범위', makeNumber(opts.tol ?? 1.0)),
    kv('단위', makeUnit(opts.unit || 'inch')),
    kv('선택된 검토 파라미터', makeText((normalizeReviewParams(opts.reviewParams && opts.reviewParams.length ? opts.reviewParams : parseCsvParams(opts.param || 'Comments'))).join(', '), '목록에서 파라미터를 추가하세요')),
    kv('추가 추출 파라미터', makeText(opts.extraParams || '', 'PM1, PM2, ... 복수 입력 가능')),
    kv('검토 대상 필터', targetFilterInput),
    kv('End_ + Dummy 패밀리 제외', excludeEndDummy)
  );
  cardSettings.append(h1('설정'), grid);

  const paramInput = grid.querySelectorAll('input[type="text"]')[0];
  const paramTools = div('conn-param-tools');
  const paramActionRow = div('conn-param-actions');
  const paramSuggest = document.createElement('input');
  paramSuggest.type = 'text';
  paramSuggest.setAttribute('list', 'conn-param-datalist');
  paramSuggest.placeholder = '공유 파라미터 검색/선택';
  paramSuggest.style.width = '100%';
  paramSuggest.style.padding = '10px 12px';
  paramSuggest.style.borderRadius = '12px';
  paramSuggest.style.border = '1px solid rgba(148,163,184,.28)';
  const paramDataList = document.createElement('datalist');
  paramDataList.id = 'conn-param-datalist';

  const addParamsBtn = cardBtn('추가', () => {
    state.reviewParams = mergeReviewParams(state.reviewParams, paramSuggest.value || '');
    renderReviewParamChips();
    paramSuggest.value = '';
    commit();
  });
  addParamsBtn.classList.add('btn-outline');

  const clearParamsBtn = cardBtn('전체 비우기', () => {
    state.reviewParams = [];
    renderReviewParamChips();
    renderParamSearchResults();
    commit();
  });
  clearParamsBtn.classList.add('btn-outline');

  const chipsWrap = div('conn-param-chips');
  chipsWrap.style.display = 'flex';
  chipsWrap.style.flexWrap = 'wrap';
  chipsWrap.style.gap = '6px';
  chipsWrap.style.marginTop = '6px';
  const paramResultMeta = document.createElement('div');
  paramResultMeta.className = 'conn-param-result-meta';
  paramResultMeta.style.color = 'var(--muted,#64748b)';
  paramResultMeta.style.fontSize = '12px';
  const paramResultWrap = div('familylink-target-list');
  paramResultWrap.style.maxHeight = '240px';
  paramResultWrap.style.overflow = 'auto';
  paramResultWrap.style.border = '1px solid rgba(148,163,184,.18)';
  paramResultWrap.style.borderRadius = '12px';
  paramResultWrap.style.padding = '6px';

  paramActionRow.style.display = 'grid';
  paramActionRow.style.gridTemplateColumns = '1fr auto auto';
  paramActionRow.style.gap = '6px';
  paramActionRow.append(paramSuggest, addParamsBtn, clearParamsBtn);
  paramTools.append(paramActionRow, paramResultMeta, paramResultWrap, chipsWrap, paramDataList);

  const paramKv = paramInput.closest('.conn-kv');
  if (paramKv) paramKv.append(paramTools);
  paramInput.readOnly = true;
  paramInput.style.background = 'rgba(15,23,42,.04)';

  const cardActions = div('conn-card section section-actions');
  cardActions.innerHTML = '<div class="conn-title">결과 검토</div>';
  const excelHelp = document.createElement('ul');
  excelHelp.className = 'conn-excel-hint';
  excelHelp.innerHTML = `
    <li><strong>Connection Type</strong>: Proximity - 허용범위 내 연결 필요, Physical - 물리적 연결</li>
    <li><strong>Status</strong>: 연결 필요 / Mismatch / Shared Parameter 등록 필요 / 연결 대상 객체 없음</li>
    <li><strong>ParamCompare</strong>: Match / Mismatch / BothEmpty / N/A</li>
    <li><strong>Value1 / Value2</strong>: 비교 대상의 Parameter 값</li>`;

  const filterGuideBtn = cardBtn('And/Or 필터 사용방법', onOpenFilterGuide);
  const filterGuideModal = createFilterGuideModal();

  cardActions.append(excelHelp, filterGuideBtn, filterGuideModal.overlay);

  rowSettings.append(cardSettings, cardActions);

  // 검토 결과 (sticky)
  const cardResults = div('conn-card section section-results conn-sticky feature-results-panel');
  const resultsTitle = h1('검토 결과');
  const summary = div('conn-summary');
  const badgeAll = chip('총 결과', '0');
  const badgeFiltered = chip('표시 중', '0');
  summary.append(badgeAll, badgeFiltered);

  const tabBar = div('conn-tabs');
  const tabs = [
    { key: 'mismatch', label: 'Mismatch' },
    { key: 'not-connected', label: 'Not Connected' }
  ];
  const tabButtons = new Map();

  tabs.forEach(({ key, label }) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'conn-tab';
    btn.dataset.tab = key;
    btn.textContent = label;
    btn.addEventListener('click', () => setTab(key));
    tabButtons.set(key, btn);
    tabBar.append(btn);
  });

  const resultHead = div('feature-results-head');
  resultHead.append(resultsTitle, tabBar, summary);

  // 안내문(최초 숨김 – [검토 시작] 때만 표시)
  const emptyGuide = div('conn-empty');
  emptyGuide.setAttribute('aria-live','polite');
  emptyGuide.textContent = '🧩 검토를 시작하려면 상단에서 기준을 설정하고 [검토 시작]을 눌러주세요.';
  const previewNotice = div('conn-preview-note');
  previewNotice.style.display = 'none';

  cardResults.append(resultHead, emptyGuide, previewNotice);

  // 결과 표 (최초 숨김)
  const tableWrap = div('conn-tablewrap');
  const table = document.createElement('table'); table.className = 'conn-table';
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  table.append(thead, tbody);
  tableWrap.append(table);

  // 최초엔 결과 섹션 자체를 숨김
  cardResults.style.display = 'none';
  tableWrap.style.display = 'none';
  emptyGuide.style.display = 'none';

  cardResults.append(tableWrap);
  page.append(rowSettings, cardResults);
  target.append(page);

  // refs
  const tol = grid.querySelector('input[type="number"]');
  const unit = grid.querySelector('select');
  const textInputs = grid.querySelectorAll('input[type="text"]');
  const param = paramInput;
  const extra = textInputs[1];
  const targetFilter = textInputs[2];

  const checkInputs = grid.querySelectorAll('input[type="checkbox"]');
  const excludeCheckbox = checkInputs[0];

  const commit = () => {
    const finalParams = getFinalReviewParams();
    param.value = finalParams.join(', ');
    saveOpts({
      tol: parseFloat(tol.value || '1') || 1,
      unit: String(unit.value),
      param: finalParams.join(',') || 'Comments',
      reviewParams: finalParams,
      extraParams: String(extra.value || ''),
      targetFilter: String(targetFilter.value || ''),
      excludeEndDummy: Boolean(excludeCheckbox.checked)
    });
  };
  tol.addEventListener('change', () => { commit(); if(state.hasRun) paint(); });
  unit.addEventListener('change', () => { commit(); if(state.hasRun) paint(); });
  param.addEventListener('change', commit);
  extra.addEventListener('change', commit);
  targetFilter.addEventListener('change', commit);
  excludeCheckbox.addEventListener('change', commit);

  /* ---- Head (ParamName 숨김) ---- */
  function headerDefs() {
    const isMm = String(unit.value) === 'mm';
    const distHeader = isMm ? 'Distance (mm)' : 'Distance (inch)';
    const base = [
      { key: 'Id1', label: 'Id1', classes: ['mono'] },
      { key: 'Id2', label: 'Id2', classes: ['mono'] },
      { key: 'Category1', label: 'Category1' },
      { key: 'Category2', label: 'Category2' },
      { key: 'Family1', label: 'Family1', classes: ['dim'] },
      { key: 'Family2', label: 'Family2', classes: ['dim'] },
      { key: 'Distance (inch)', label: distHeader, classes: ['num'] },
      { key: 'ConnectionType', label: 'ConnectionType' },
      { key: 'Value1', label: 'Value1', classes: ['dim', 'tone-cell'] },
      { key: 'Value2', label: 'Value2', classes: ['dim', 'tone-cell'] },
      { key: 'ParamCompare', label: 'ParamCompare', classes: ['tone-cell'] },
      { key: 'Status', label: 'Status', classes: ['tone-cell'] },
      { key: 'ErrorMessage', label: 'ErrorMessage', classes: ['dim'] }
    ];

    return base;
  }

  function paintHead() {
    const headers = headerDefs();
    const tr = document.createElement('tr');
    headers.forEach(h => {
      const th = document.createElement('th');
      th.textContent = h.label;
      if (Array.isArray(h.classes)) th.className = h.classes.join(' ');
      tr.append(th);
    });
    thead.innerHTML = '';
    thead.append(tr);
  }

  /* ---- Body ---- */
  function paintBody() {
    while (tbody.firstChild) tbody.removeChild(tbody.firstChild);
    const isMm = String(unit.value) === "mm";
    const headers = headerDefs();

    const { rows: activeRows, total: activeTotal, previewCount, hasMore } = getActiveMeta();

    badgeAll.querySelector(".num").textContent = String(activeTotal);
    badgeFiltered.querySelector(".num").textContent = String(activeRows.length);

    if (hasMore) {
      previewNotice.textContent = `미리보기에서는 상위 ${previewCount}건만 표시됩니다. 전체 ${activeTotal}건은 엑셀 내보내기로 확인하세요.`;
      previewNotice.style.display = 'block';
    } else {
      previewNotice.textContent = '';
      previewNotice.style.display = 'none';
    }

    if (activeRows.length === 0) {
      const tr = document.createElement("tr");
      const td = document.createElement("td");
      td.colSpan = headers.length;
      td.textContent = "해당 조건의 결과가 없습니다.";
      td.className = "conn-empty-row";
      tr.append(td);
      tbody.append(tr);
      updateSaveDisabled();
      return;
    }

    if (activeRows.length > MAX_PREVIEW_ROWS) {
      const tr = document.createElement("tr");
      const td = document.createElement("td");
      td.colSpan = headers.length;
      td.textContent = "결과가 150개 이상입니다. 미리보기 대신 엑셀 내보내기를 이용해 주세요.";
      td.className = "conn-empty-row";
      tr.append(td);
      tbody.append(tr);
      updateSaveDisabled();
      return;
    }

    activeRows.forEach(r => {
      const tr = document.createElement("tr");

      const statusVal = r.Status ?? r.status;
      const cells = headers.map(h => {
        let v = r[h.key];
        if (h.key === 'Distance (inch)') {
          let dist = (r["Distance (inch)"] ?? r.DistanceInch ?? "");
          if (isMm && dist !== "") {
            const converted = toMm(dist);
            dist = Number.isFinite(converted) ? converted.toFixed(4) : converted;
          }
          v = dist;
        }
        return v;
      });

      cells.forEach((v, idx) => {
        const td = tdText(v);
        const defs = headers[idx];
        if (defs && Array.isArray(defs.classes)) td.classList.add(...defs.classes);

        if (defs && defs.key && (defs.key === 'Value1' || defs.key === 'Value2' || defs.key === 'ParamCompare' || defs.key === 'Status')) {
          const kind = statusKind(statusVal);
          td.classList.add("tone-cell",
            kind==='ok'?'tone-ok':kind==='warn'?'tone-warn':kind==='bad'?'tone-bad':'tone-info');
        }
        tr.append(td);
      });

      tbody.append(tr);
    });

    updateSaveDisabled();
  }

  function paint(){
    paintHead();
    paintBody();
  }

  function applyIncomingRows(payload){
    const rows = (payload && Array.isArray(payload.rows)) ? payload.rows : [];
    state.extraParams = parseExtraParams(payload && payload.extraParams);
    const mismatchSection = (payload && payload.mismatch) || {};
    const nearSection = (payload && payload.near) || {};

    const cleaned = Array.isArray(rows) ? rows : [];
    const mismatchFromCleaned = cleaned.filter(isMismatchIssue);
    const nearFromCleaned = cleaned.filter(isNotConnectedIssue);

    const mismatchPreview = Array.isArray(mismatchSection.rows) ? mismatchSection.rows : mismatchFromCleaned;
    const nearPreview = Array.isArray(nearSection.rows) ? nearSection.rows : nearFromCleaned;

    state.rowsInch = cleaned;
    state.mismatchTotal = Number(mismatchSection.total) || mismatchFromCleaned.length;
    state.notConnectedTotal = Number(nearSection.total) || nearFromCleaned.length;

    state.mismatchRows = mismatchPreview.slice(0, MAX_PREVIEW_ROWS);
    state.notConnectedRows = nearPreview.slice(0, MAX_PREVIEW_ROWS);

    state.mismatchPreviewCount = Number(mismatchSection.previewCount) || Math.min(state.mismatchRows.length, Math.max(state.mismatchTotal, state.mismatchRows.length), MAX_PREVIEW_ROWS);
    state.notConnectedPreviewCount = Number(nearSection.previewCount) || Math.min(state.notConnectedRows.length, Math.max(state.notConnectedTotal, state.notConnectedRows.length), MAX_PREVIEW_ROWS);

    state.mismatchHasMore = (mismatchSection.hasMore === true) || state.mismatchTotal > MAX_PREVIEW_ROWS;
    state.notConnectedHasMore = (nearSection.hasMore === true) || state.notConnectedTotal > MAX_PREVIEW_ROWS;

    const totalFromPayload = Number(payload && payload.total);
    state.totalCount = (Number.isFinite(totalFromPayload) && totalFromPayload > 0)
      ? totalFromPayload
      : (cleaned.length > 0 ? cleaned.length : (state.mismatchTotal + state.notConnectedTotal));

    setTab('mismatch', { silent: true });

    // 전환: 안내문 숨김 → 표 표시
    emptyGuide.style.display = 'none';
    tableWrap.style.display = 'block';

    paint();
  }



  function renderReviewParamChips() {
    chipsWrap.innerHTML = '';
    const items = normalizeReviewParams(state.reviewParams);
    state.reviewParams = items;
    param.value = items.join(', ');
    if (!items.length) {
      const empty = document.createElement('span');
      empty.className = 'conn-chip';
      empty.textContent = '선택된 파라미터 없음';
      chipsWrap.append(empty);
      return;
    }
    items.forEach((name, idx) => {
      const chipEl = document.createElement('span');
      chipEl.className = 'conn-chip';
      chipEl.style.display = 'inline-flex';
      chipEl.style.alignItems = 'center';
      chipEl.style.gap = '6px';

      const t = document.createElement('span');
      t.textContent = name;
      const del = document.createElement('button');
      del.type = 'button';
      del.textContent = '×';
      del.className = 'conn-tab';
      del.style.padding = '0 6px';
      del.style.lineHeight = '1';
      del.addEventListener('click', () => {
        state.reviewParams = state.reviewParams.filter((_, i) => i !== idx);
        renderReviewParamChips();
        renderParamSearchResults();
        commit();
      });
      chipEl.append(t, del);
      chipsWrap.append(chipEl);
    });
  }

  function renderParamSearchResults() {
    const query = String(paramSuggest.value || '').trim().toLowerCase();
    const selected = new Set(normalizeReviewParams(state.reviewParams).map((x) => x.toLowerCase()));
    const items = Array.isArray(state.paramItems) ? state.paramItems : [];
    const filtered = items.filter((item) => {
      if (!query) return true;
      const hay = `${item.name || ''} ${item.groupName || ''}`.toLowerCase();
      return hay.includes(query);
    });
    paramResultMeta.textContent = items.length ? `공유 파라미터 ${filtered.length}/${items.length}개` : '공유 파라미터 목록이 없습니다.';
    paramResultWrap.innerHTML = '';
    if (!items.length) {
      const empty = div('familylink-target-empty');
      empty.textContent = '공유 파라미터 목록을 불러오는 중이거나, 현재 사용할 수 없습니다.';
      paramResultWrap.append(empty);
      return;
    }
    if (!filtered.length) {
      const empty = div('familylink-target-empty');
      empty.textContent = '검색 결과가 없습니다.';
      paramResultWrap.append(empty);
      return;
    }
    filtered.slice(0, 120).forEach((item) => {
      const row = document.createElement('button');
      row.type = 'button';
      row.className = 'familylink-target-row';
      row.style.display = 'grid';
      row.style.gridTemplateColumns = '1fr auto';
      row.style.alignItems = 'center';
      row.style.width = '100%';
      row.style.textAlign = 'left';
      row.style.border = '0';
      row.style.background = selected.has(String(item.name || '').toLowerCase()) ? 'rgba(59,130,246,.12)' : 'transparent';
      row.style.borderRadius = '10px';
      row.style.padding = '8px 10px';
      row.style.cursor = 'pointer';

      const info = document.createElement('span');
      info.style.display = 'grid';
      const main = document.createElement('strong');
      main.textContent = item.name || '';
      const sub = document.createElement('small');
      sub.textContent = item.groupName ? `${item.groupName}${item.guid ? ' · ' + item.guid.slice(0, 8) : ''}` : (item.guid ? item.guid.slice(0, 8) : '');
      info.append(main, sub);

      const badge = document.createElement('span');
      badge.className = selected.has(String(item.name || '').toLowerCase()) ? 'chip chip--ok' : 'chip chip--info';
      badge.textContent = selected.has(String(item.name || '').toLowerCase()) ? '선택됨' : '추가';

      row.append(info, badge);
      row.addEventListener('click', () => {
        const next = normalizeReviewParams(state.reviewParams);
        const idx = next.findIndex((x) => x.toLowerCase() === String(item.name || '').toLowerCase());
        if (idx >= 0) next.splice(idx, 1); else next.push(item.name);
        state.reviewParams = next;
        renderReviewParamChips();
        renderParamSearchResults();
        commit();
      });
      paramResultWrap.append(row);
    });
  }

  function getFinalReviewParams() {
    const merged = normalizeReviewParams(state.reviewParams);
    if (merged.length > 0) return merged;
    const fallback = String(param.value || 'Comments').trim();
    return fallback ? parseCsvParams(fallback) : ['Comments'];
  }

  paramSuggest.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    state.reviewParams = mergeReviewParams(state.reviewParams, paramSuggest.value || '');
    renderReviewParamChips();
    paramSuggest.value = '';
    renderParamSearchResults();
    commit();
  });

  function onRun(){
    closeCompletionSummaryDialog();
    commit(); setBusy(true);
    acceptRunProgress = true;
    acceptExcelProgress = false;
    state.hasRun = true;

    // 결과 섹션 오픈 + 안내문 보이기
    cardResults.style.display = 'block';
    emptyGuide.style.display = 'flex';
    tableWrap.style.display = 'none';

    let sendTol = parseFloat(tol.value || '1');
    let sendUnit = String(unit.value || 'inch');
    if (sendUnit === 'mm') { if (!isFinite(sendTol)) sendTol = 1; sendTol = sendTol / INCH_TO_MM; sendUnit = 'inch'; }
    const finalParams = getFinalReviewParams();
    const paramsCsv = finalParams.join(',');
    post('connector:run', {
      tol: sendTol,
      unit: sendUnit,
      param: paramsCsv || String(param.value || 'Comments'),
      paramsCsv,
      extraParams: String(extra.value || ''),
      targetFilter: String(targetFilter.value || ''),
      excludeEndDummy: Boolean(excludeCheckbox.checked)
    });
  }

  function onExport() {
    if (exporting) return;
    chooseExcelMode((mode) => {
      exporting = true;
      acceptExcelProgress = true;
      updateSaveDisabled();
      run.disabled = true;
      const tab = state.tab || 'mismatch';
      const uiUnit = String(unit.value || 'inch');
      const finalParams = getFinalReviewParams();
      const paramsCsv = finalParams.join(',');
      post('connector:save-excel', {
        excelMode: mode || 'fast',
        tab,
        uiUnit,
        param: paramsCsv || String(param.value || 'Comments'),
        paramsCsv,
        extraParams: String(extra.value || ''),
        targetFilter: String(targetFilter.value || ''),
        excludeEndDummy: !!excludeCheckbox.checked
      });
    });
  }


  onHost(({ ev, payload }) => {
    switch (ev) {
      case 'connector:loaded':
        setBusy(false);
        // 결과 섹션 보장
        cardResults.style.display = 'block';
        applyIncomingRows(payload || {});
        break;
      case 'connector:done':
        acceptRunProgress = false;
        setBusy(false);
        ProgressDialog.hide();
        cardResults.style.display = 'block';
        applyIncomingRows(payload || {});
        if (payload?.ok !== false) {
          showConnectorCompletionDialog(payload || {});
        }
        break;
      case 'connector:progress':
        handleConnectorProgress(payload);
        break;
      case 'connector:param-list:done': {
        const list = Array.isArray(payload && payload.params) ? payload.params : [];
        const items = Array.isArray(payload && payload.items) ? payload.items : [];
        state.paramList = parseCsvParams(list.join(','));
        state.paramItems = normalizeParamItems(items.length ? items : state.paramList.map((name) => ({ name })));
        paramDataList.innerHTML = '';
        state.paramList.forEach((name) => {
          const opt = document.createElement('option');
          opt.value = name;
          paramDataList.append(opt);
        });
        renderParamSearchResults();
        break;
      }
      case 'connector:saved': {
        lastExcelPct = 0;
        exporting = false;
        acceptExcelProgress = false;
        run.disabled = false;
        ProgressDialog.hide();
        updateSaveDisabled();
        const p = (payload && payload.path) || '';
        if (p) {
          showExcelSavedDialog('엑셀 파일을 내보냈습니다.', p, (path) => {
            if (path) post('excel:open', { path });
          });
        } else {
          toast('엑셀 파일을 내보냈습니다.', 'ok', 2600);
        }
        break;
      }
      case 'revit:error':
        setBusy(false);
        exporting = false;
        acceptRunProgress = false;
        acceptExcelProgress = false;
        lastExcelPct = 0;
        ProgressDialog.hide();
        run.disabled = false;
        updateSaveDisabled();
        toast((payload && payload.message) || '오류가 발생했습니다.', 'err', 3200);
        break;
      case 'host:error':
        setBusy(false);
        exporting = false;
        acceptRunProgress = false;
        acceptExcelProgress = false;
        lastExcelPct = 0;
        ProgressDialog.hide();
        run.disabled = false;
        updateSaveDisabled();
        toast((payload && payload.message) || '호스트 오류가 발생했습니다.', 'err', 3200);
        break;
      default: break;
    }
  });

  /* helpers */
  function handleConnectorProgress(payload) {
    if (payload && (Object.prototype.hasOwnProperty.call(payload, 'pct') || Object.prototype.hasOwnProperty.call(payload, 'text'))) {
      handleRunProgress(payload || {});
      return;
    }
    handleExcelProgress(payload || {});
  }

  function handleRunProgress(payload) {
    if (!acceptRunProgress) return;
    const percent = typeof payload?.pct === 'number' ? payload.pct : 0;
    const message = payload?.text || '';
    if (percent <= 0 && !message) {
      ProgressDialog.hide();
      return;
    }
    ProgressDialog.show('커넥터 진단', message || '진행 중…');
    ProgressDialog.update(percent, message || '', '');
  }

  function handleExcelProgress(payload) {
    if (!acceptExcelProgress) return;
    if (!payload) {
      ProgressDialog.hide();
      exporting = false;
      run.disabled = false;
      lastExcelPct = 0;
      updateSaveDisabled();
      return;
    }
    const phase = normalizeExcelPhase(payload.phase);
    const total = Number(payload.total) || 0;
    const current = Number(payload.current) || 0;
    const percent = computeExcelPercent(phase, current, total, payload.phaseProgress);
    const subtitle = buildExcelSubtitle(phase, current, total);
    const detail = payload.message || '';

    exporting = phase !== 'DONE' && phase !== 'ERROR';
    run.disabled = exporting;
    updateSaveDisabled();

    ProgressDialog.show('커넥터 엑셀 내보내기', subtitle);
    ProgressDialog.update(percent, subtitle, detail);

    if (!exporting) {
      setTimeout(() => { ProgressDialog.hide(); lastExcelPct = 0; }, 260);
    }
  }

  function normalizeExcelPhase(phase) {
    return String(phase || '').trim().toUpperCase() || 'EXCEL_WRITE';
  }

  function computeExcelPercent(phase, current, total, phaseProgress) {
    const norm = normalizeExcelPhase(phase);
    if (norm === 'DONE') { lastExcelPct = 100; return 100; }
    if (norm === 'ERROR') return lastExcelPct;

    const completed = EXCEL_PHASE_ORDER.reduce((acc, key) => {
      if (key === norm) return acc;
      return acc + (EXCEL_PHASE_WEIGHT[key] || 0);
    }, 0);
    const weight = EXCEL_PHASE_WEIGHT[norm] || 0;
    const ratio = total > 0 ? Math.min(1, Math.max(0, current / total)) : 0;
    const staged = Math.max(ratio, clamp01(phaseProgress));
    const pct = Math.min(100, Math.max(lastExcelPct, (completed + weight * staged) * 100));
    lastExcelPct = pct;
    return pct;
  }

  function buildExcelSubtitle(phase, current, total) {
    const labelMap = {
      EXCEL_INIT: '엑셀 준비',
      EXCEL_WRITE: '엑셀 작성',
      EXCEL_SAVE: '파일 저장',
      AUTOFIT: 'AutoFit',
      DONE: '완료',
      ERROR: '오류'
    };
    const label = labelMap[normalizeExcelPhase(phase)] || '엑셀 작업';
    const count = total > 0 ? ` (${Math.max(current, 0)}/${total})` : '';
    return `${label}${count}`;
  }

  function clamp01(v) {
    const n = Number(v);
    if (!Number.isFinite(n)) return 0;
    return Math.max(0, Math.min(1, n));
  }

  function normalizeStatus(row){
    return String((row && (row.Status ?? row.status)) || '').trim().toUpperCase();
  }

  function isMismatchIssue(row){
    const status = normalizeStatus(row);
    if (status === 'MISMATCH') return true;
    if (status === 'SHARED PARAMETER 등록 필요'.toUpperCase()) return true;
    return false;
  }

  function isNotConnectedIssue(row){
    const status = normalizeStatus(row);
    if (status.includes('연결 필요')) return true;
    if (status.includes('연결 대상 객체 없음')) return true;
    const conn = normalizeConnectionType(row).toUpperCase();
    return conn.includes('PROXIMITY') || conn === 'NEAR';
  }

  function normalizeConnectionType(row){
    return String((row && (row.ConnectionType ?? row.connectionType ?? row.Type ?? row.type)) || '').trim();
  }

  function getActiveRows(){
    const base = state.tab === 'mismatch'
      ? state.mismatchRows
      : state.notConnectedRows;

    return Array.isArray(base) ? base : [];
  }

  function getActiveMeta(){
    if (state.tab === 'mismatch') {
      return {
        rows: getActiveRows(),
        total: state.mismatchTotal,
        previewCount: state.mismatchPreviewCount || getActiveRows().length,
        hasMore: state.mismatchHasMore
      };
    }
    return {
      rows: getActiveRows(),
      total: state.notConnectedTotal,
      previewCount: state.notConnectedPreviewCount || getActiveRows().length,
      hasMore: state.notConnectedHasMore
    };
  }

  function updateSaveDisabled(){
    const saveBtn = document.getElementById('btnConnSave');
    if (saveBtn) saveBtn.disabled = state.totalCount === 0 || exporting;
  }

  function showConnectorCompletionDialog(payload){
    const summary = payload && typeof payload.summary === 'object' ? payload.summary : {};
    const allCount = Number(summary.allCount);
    const displayedTotal = Number(summary.displayedTotal);
    const mismatchCount = Number(summary.mismatchCount);
    const nearCount = Number(summary.nearCount);
    const previewCount = Number(summary.previewCount);
    const normalCount = Number(summary.normalCount);

    const totalResults = Number.isFinite(allCount) ? allCount : Math.max(state.totalCount, state.rowsInch.length);
    const shownTotal = Number.isFinite(displayedTotal) ? displayedTotal : state.totalCount;
    const mismatch = Number.isFinite(mismatchCount) ? mismatchCount : state.mismatchTotal;
    const near = Number.isFinite(nearCount) ? nearCount : state.notConnectedTotal;
    const normal = Number.isFinite(normalCount) ? normalCount : Math.max(totalResults - mismatch - near, 0);
    const shown = Number.isFinite(previewCount) ? previewCount : Math.min(shownTotal, MAX_PREVIEW_ROWS);

    const notes = [];
    if (summary?.hasMore === true || payload?.hasMore === true) {
      notes.push(`미리보기에서는 상위 ${shown}건만 표시했습니다. 전체 ${shownTotal}건은 엑셀 내보내기로 확인할 수 있습니다.`);
    }

    showCompletionSummaryDialog({
      title: '커넥터 진단 완료',
      message: '검토 결과를 요약했습니다. 필요하면 바로 엑셀로 내보내세요.',
      summaryItems: [
        { label: '전체 결과 건수', value: String(totalResults) },
        { label: '오류/불일치 건수', value: String(mismatch) },
        { label: 'near/연결 필요 건수', value: String(near) },
        { label: '정상 건수', value: String(normal) },
        { label: '표시 건수 / 전체 건수', value: `${shown} / ${shownTotal}` }
      ],
      notes,
      exportDisabled: !!save.disabled,
      onExport: () => {
        save.click();
      }
    });
  }


  function setTab(tab, opts = {}){
    if (!tabButtons.has(tab)) return;
    state.tab = tab;
    tabButtons.forEach((btn, key) => {
      if (key === tab) btn.classList.add('is-active'); else btn.classList.remove('is-active');
    });
    if (!opts.silent) {
      paintBody();
    }
  }

  setTab('mismatch', { silent: true });
  renderReviewParamChips();
  renderParamSearchResults();
  try { post('connector:param-list', {}); } catch (_) {}

  function h1(t){ const e=document.createElement('div'); e.className='conn-title'; e.textContent=t; return e; }
  function kv(label, inputEl){ const wrap=document.createElement('div'); wrap.className='conn-kv'; const cap=document.createElement('label'); cap.textContent=label; wrap.append(cap,inputEl); return wrap; }
  function chip(label, numText){ const el=document.createElement('span'); el.className='conn-chip'; const t=document.createElement('span'); t.textContent=label; const n=document.createElement('span'); n.className='num'; n.textContent=numText; el.append(t,n); return el; }
  function cardBtn(text, onClick) {
    const b = document.createElement('button');
    b.textContent = text;
    b.className = 'card-action-btn';
    if (typeof onClick === 'function') b.addEventListener('click', onClick);
    return b;
  }
  function makeNumber(v){ const i=document.createElement('input'); i.type='number'; i.step='0.0001'; i.value=String(v); return i; }
  function makeUnit(v){ const s=document.createElement('select'); s.className='kkyt-select'; s.innerHTML='<option value="inch">inch</option><option value="mm">mm</option>'; s.value=String(v); return s; }
  function makeText(v, placeholder=''){ const i=document.createElement('input'); i.type='text'; i.value=String(v); if(placeholder) i.placeholder=placeholder; i.style.width='100%'; i.title=i.value||placeholder||''; i.addEventListener('input',()=>{ i.title=i.value||placeholder||'';}); return i; }
  function makeCheckbox(v){ const i=document.createElement('input'); i.type='checkbox'; i.checked=!!v; return i; }
  function parseCsvParams(raw){
    const txt = Array.isArray(raw) ? raw.join(',') : String(raw || '');
    const parts = txt.split(',').map(s => String(s || '').trim()).filter(Boolean);
    const seen = new Set();
    const out = [];
    for (const p of parts) {
      const key = p.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(p);
    }
    return out;
  }

  function normalizeReviewParams(raw) {
    return parseCsvParams(Array.isArray(raw) ? raw.join(',') : raw);
  }

  function normalizeParamItems(items) {
    const seen = new Set();
    const out = [];
    (Array.isArray(items) ? items : []).forEach((item) => {
      const name = String(item && item.name ? item.name : item || '').trim();
      if (!name) return;
      const key = name.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      out.push({
        name,
        groupName: String(item && item.groupName ? item.groupName : '').trim(),
        guid: String(item && item.guid ? item.guid : '').trim(),
        source: String(item && item.source ? item.source : '').trim()
      });
    });
    out.sort((a, b) => a.name.localeCompare(b.name, 'ko'));
    return out;
  }

  function mergeReviewParams(baseList, incomingRaw){
    const merged = [];
    const seen = new Set();
    const push = (v) => {
      const t = String(v || '').trim();
      if (!t) return;
      const key = t.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      merged.push(t);
    };
    (Array.isArray(baseList) ? baseList : []).forEach(push);
    parseCsvParams(incomingRaw).forEach(push);
    return merged;
  }

  function parseExtraParams(raw){
    const txt = Array.isArray(raw) ? raw.join(',') : String(raw || '');
    const parts = txt.split(',').map(s => String(s||'').trim()).filter(Boolean);
    const seen = new Set();
    const out = [];
    for (const p of parts) {
      const key = p.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(p);
    }
    return out;
  }

  function onOpenFilterGuide(){
    filterGuideModal.open();
  }

  function createFilterGuideModal(){
    const overlay = document.createElement('div');
    overlay.className = 'conn-filter-overlay';

    const dialog = document.createElement('div');
    dialog.className = 'conn-filter-modal';

    const headerEl = document.createElement('div');
    headerEl.className = 'conn-filter-modal__header';

    const title = document.createElement('h3');
    title.textContent = '검토대상 필터 사용방법';
    title.className = 'conn-filter-modal__title';

    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.textContent = '×';
    closeBtn.setAttribute('aria-label', '닫기');
    closeBtn.className = 'conn-filter-modal__close';

    headerEl.append(title, closeBtn);

    const body = document.createElement('div');
    body.className = 'conn-filter-modal__body';
    body.innerHTML = `
      <p class="conn-filter-modal__intro">열 이름과 값을 이용해 조건을 작성하면, 해당 조건을 만족하는 행만 엑셀에 포함됩니다.</p>
      <ul class="conn-excel-hint conn-filter-modal__list">
        <li><strong>기본 비교식</strong>: <code>PM1=Value</code> (대소문자 무시)</li>
        <li><strong>값에 공백</strong>: <code>PM_NAME='A B'</code> 또는 <code>PM_NAME="A B"</code></li>
        <li><strong>AND</strong>: <code>AND(cond1, cond2, ...)</code></li>
        <li><strong>OR</strong>: <code>OR(cond1, cond2, ...)</code></li>
        <li><strong>NOT</strong>: <code>NOT(cond)</code></li>
        <li><strong>구분자</strong>: 조건을 <code>,</code> 또는 <code>;</code> 로 나열하면 자동 AND 처리</li>
      </ul>
      <div class="conn-filter-modal__section">
        <div class="conn-filter-modal__section-title">콤마 생략도 허용되는 예시</div>
        <pre class="conn-filter-modal__code">OR(AND(PM1=1,PM2=2)AND(PM1=1,PM2=3))</pre>
      </div>
      <div class="conn-filter-modal__section">
        <div class="conn-filter-modal__section-title">대표 예시</div>
        <pre class="conn-filter-modal__code">OR(AND(PM1=1,PM2=2), AND(PM1=1,PM2=3))
→ PM1=1 이고 PM2가 2 또는 3</pre>
      </div>
      <div class="conn-filter-modal__section">
        <div class="conn-filter-modal__section-title">주의사항</div>
        <ul class="conn-filter-modal__list">
          <li><code>=</code> 비교만 지원(>, < 등 미지원)</li>
          <li>괄호/콤마 개수와 순서를 맞춰 주세요.</li>
          <li>파라미터명이 Revit 파라미터명과 정확히 일치해야 합니다.</li>
        </ul>
      </div>`;

    dialog.append(headerEl, body);
    overlay.append(dialog);

    function open(){
      overlay.classList.add('is-open');
      document.addEventListener('keydown', onKey);
    }

    function close(){
      overlay.classList.remove('is-open');
      document.removeEventListener('keydown', onKey);
    }

    function onKey(e){
      if (e.key === 'Escape') close();
    }

    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) close();
    });
    closeBtn.addEventListener('click', close);

    return { overlay, open, close };
  }
}
