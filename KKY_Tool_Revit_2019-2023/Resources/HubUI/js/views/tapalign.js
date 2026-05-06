import { clear, div, toast, setBusy, showExcelSavedDialog, chooseExcelMode, getLastExcelExportLocale } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { post, onHost } from '../core/bridge.js';

const SKEY = 'kky_tapalign_opts';
const COMMON_OPTIONS_KEY = 'kky.hub.commonOptions';
const DEFAULTS = { tol: 0.5, unit: 'mm', domain: 'all', featureTargetFilter: '' };
const EXCEL_PHASE_ORDER = ['EXCEL_INIT', 'EXCEL_WRITE', 'EXCEL_SAVE', 'AUTOFIT'];
const EXCEL_PHASE_WEIGHT = { EXCEL_INIT: 0.05, EXCEL_WRITE: 0.65, EXCEL_SAVE: 0.15, AUTOFIT: 0.15 };

let activeState = null;

export function renderTapAlign(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  const topbar = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar');
  if (topbar) topbar.classList.add('hub-topbar');

  const opts = loadOpts();
  const state = {
    rows: [],
    running: false,
    exporting: false,
    lastExcelPct: 0,
    hasRun: false,
    unit: opts.unit || DEFAULTS.unit,
    domain: opts.domain || DEFAULTS.domain,
    featureTargetFilter: typeof opts.featureTargetFilter === 'string' ? opts.featureTargetFilter : DEFAULTS.featureTargetFilter,
    resultUnit: opts.unit || DEFAULTS.unit,
    resultDomain: opts.domain || DEFAULTS.domain,
    extraHeaders: [],
    commonOptions: normalizeCommonOptions(null),
    commonLoadedFromLocal: false,
    refs: {}
  };

  activeState = state;
  bindListeners();

  const page = div('conn-page feature-shell');

  const header = div('feature-header');
  const heading = div('feature-heading');
  heading.innerHTML = `
    <span class="feature-kicker">BQC 보조</span>
    <h2 class="feature-title">탭/분기 축 틀어짐 검토</h2>
    <p class="feature-sub">탭 또는 분기 피팅의 연결 축이 연결된 배관/덕트 중심축을 통과하는지 검토합니다.</p>`;

  const runBtn = cardBtn('검토 시작', () => onRun(state));
  const exportBtn = cardBtn('엑셀 내보내기', () => onExport(state));
  runBtn.classList.add('btn-primary');
  exportBtn.classList.add('btn-outline');
  exportBtn.disabled = true;

  const actions = div('feature-actions');
  actions.append(runBtn, exportBtn);
  header.append(heading, actions);
  page.append(header);

  const rowSettings = div('conn-row settings conn-sticky feature-controls');

  const cardSettings = div('conn-card section section-settings');
  const grid = div('conn-grid');
  const tolInput = makeNumber(opts.tol ?? DEFAULTS.tol);
  const unitSelect = makeUnit(state.unit);
  const domainSelect = makeDomain(state.domain);
  const extraParamsInput = makeText('', 'BQC 공통 설정에서 관리');
  const targetFilterInput = makeText('', 'BQC 공통 설정에서 관리');
  const featureTargetFilterInput = makeText(state.featureTargetFilter || '', '예: PM1=값; PM2=값2');
  const includePointXY = makeCheckbox(false);
  const includeLinearMetrics = makeCheckbox(false);
  const excludeEndDummy = makeCheckbox(false);

  tolInput.min = '0';
  tolInput.step = '0.01';
  extraParamsInput.readOnly = true;
  targetFilterInput.readOnly = true;
  extraParamsInput.classList.add('conn-param-input');
  targetFilterInput.classList.add('conn-param-input');
  featureTargetFilterInput.classList.add('conn-param-input');
  [extraParamsInput, targetFilterInput, featureTargetFilterInput].forEach((input) => {
    input.style.width = '100%';
    input.style.minWidth = '0';
    input.style.boxSizing = 'border-box';
    input.style.overflow = 'hidden';
    input.style.textOverflow = 'ellipsis';
    input.style.whiteSpace = 'nowrap';
  });
  includePointXY.disabled = true;
  includeLinearMetrics.disabled = true;
  excludeEndDummy.disabled = true;

  grid.append(
    kv('허용 범위', tolInput),
    kv('거리 단위', unitSelect),
    kv('검토 범위', domainSelect),
    kv('추가 추출 파라미터', extraParamsInput),
    kv('공통 검토 대상 필터', targetFilterInput),
    kv('기능 전용 필터', featureTargetFilterInput),
    kv('좌표 X/Y 추출', includePointXY),
    kv('선형 길이 / 방향 추출', includeLinearMetrics),
    kv('End_ + Dummy 패밀리 제외', excludeEndDummy)
  );
  cardSettings.append(h1('설정'), grid);

  const cardActions = div('conn-card section section-actions');
  cardActions.innerHTML = '<div class="conn-title">결과 검토</div>';
  const guideList = document.createElement('ul');
  guideList.className = 'conn-excel-hint';
  guideList.innerHTML = `
    <li><strong>검토 기준</strong>: 연결 커넥터 축과 연결 라인 중심축 사이의 최단거리</li>
    <li><strong>오류 조건</strong>: 허용 범위를 초과하면 "중심축에서 벗어났습니다."로 표시</li>
    <li><strong>필터 적용</strong>: 공통 검토 대상 필터를 먼저 적용하고, 기능 전용 필터를 추가 AND 조건으로 적용합니다.</li>
    <li><strong>공통 설정</strong>: 추가 추출 파라미터는 분기 객체와 연결 라인 값을 함께 추출하고, End Dummy 제외는 BQC 공통 설정을 그대로 적용합니다.</li>
    <li><strong>엑셀 옵션</strong>: 거리 단위(mm / inch)는 현재 설정을 따르고, 결과 내용 언어는 내보내기 시 선택합니다.</li>`;

  const filterGuide = document.createElement('div');
  filterGuide.className = 'conn-preview-note';
  filterGuide.style.display = 'block';
  filterGuide.textContent = '필터 예시: PM1=값; PM2=값2 또는 and(PM1=값, PM2=값2)';

  const commonSummary = div('conn-summary');
  commonSummary.style.display = 'grid';
  commonSummary.style.gridTemplateColumns = 'repeat(3, minmax(0, 1fr))';
  commonSummary.style.gap = '8px';
  const commonExtra = chip('추가 추출', '-');
  const commonFilter = chip('대상 필터', '-');
  const commonExclude = chip('End Dummy 제외', '미적용');
  [commonExtra, commonFilter, commonExclude].forEach((item) => {
    item.style.minWidth = '0';
    const num = item.querySelector('.num');
    if (num) {
      num.style.display = 'block';
      num.style.minWidth = '0';
      num.style.overflow = 'hidden';
      num.style.textOverflow = 'ellipsis';
      num.style.whiteSpace = 'nowrap';
    }
  });
  commonSummary.append(commonExtra, commonFilter, commonExclude);
  cardActions.append(guideList, filterGuide, commonSummary);

  rowSettings.append(cardSettings, cardActions);

  const cardResults = div('conn-card section section-results conn-sticky feature-results-panel');
  const resultsTitle = h1('검토 결과');
  const summary = div('conn-summary');
  const badgeIssues = chip('오류 수', '0');
  const badgeScope = chip('범위', resolveDomainLabel(state.domain));
  summary.append(badgeIssues, badgeScope);

  const resultHead = div('feature-results-head');
  resultHead.append(resultsTitle, summary);

  const emptyGuide = div('conn-empty');
  emptyGuide.setAttribute('aria-live', 'polite');
  emptyGuide.textContent = '상단 설정을 확인한 뒤 [검토 시작]을 눌러 주세요.';

  const tableWrap = div('conn-tablewrap');
  const table = document.createElement('table');
  table.className = 'conn-table';
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  table.append(thead, tbody);
  tableWrap.append(table);
  tableWrap.style.display = 'none';

  cardResults.append(resultHead, emptyGuide, tableWrap);
  cardResults.style.display = 'none';

  page.append(rowSettings, cardResults);
  target.append(page);

  state.refs = {
    runBtn,
    exportBtn,
    tolInput,
    unitSelect,
    domainSelect,
    extraParamsInput,
    targetFilterInput,
    featureTargetFilterInput,
    includePointXY,
    includeLinearMetrics,
    excludeEndDummy,
    commonExtra,
    commonFilter,
    commonExclude,
    badgeIssues,
    badgeScope,
    emptyGuide,
    tableWrap,
    thead,
    tbody,
    cardResults
  };

  const commit = () => commitOptions(state);
  unitSelect.addEventListener('change', commit);
  domainSelect.addEventListener('change', commit);
  featureTargetFilterInput.addEventListener('change', commit);
  featureTargetFilterInput.addEventListener('blur', commit);
  tolInput.addEventListener('change', commit);
  tolInput.addEventListener('blur', commit);

  loadCommonOptions(state);
  renderCommonOptions(state);
  renderState(state);
}

function bindListeners() {
  onHost('commonoptions:loaded', (payload) => {
    if (!activeState || activeState.commonLoadedFromLocal) return;
    applyCommonOptions(activeState, payload);
  });

  onHost('tapalign:progress', (payload) => {
    if (!activeState) return;
    if (payload?.isExcel) {
      handleTapAlignExcelProgress(activeState, payload || {});
      return;
    }

    const pct = Number(payload?.pct);
    const text = payload?.text || '검토 중';
    const detail = payload?.detail || '';
    ProgressDialog.update(Number.isFinite(pct) ? pct : 0, text, detail);
  });

  onHost('tapalign:done', (payload) => {
    if (!activeState) return;

    activeState.running = false;
    activeState.rows = Array.isArray(payload?.rows) ? payload.rows : [];
    activeState.unit = payload?.unit || activeState.unit || DEFAULTS.unit;
    activeState.domain = payload?.domain || activeState.domain || DEFAULTS.domain;
    activeState.featureTargetFilter = typeof payload?.featureTargetFilter === 'string'
      ? payload.featureTargetFilter
      : (activeState.featureTargetFilter || DEFAULTS.featureTargetFilter);
    activeState.resultUnit = payload?.unit || activeState.resultUnit || activeState.unit || DEFAULTS.unit;
    activeState.resultDomain = payload?.domain || activeState.resultDomain || activeState.domain || DEFAULTS.domain;
    activeState.extraHeaders = Array.isArray(payload?.extraHeaders) ? payload.extraHeaders : [];
    if (payload?.common) applyCommonOptions(activeState, payload.common);

    if (activeState.refs.unitSelect) activeState.refs.unitSelect.value = activeState.unit;
    if (activeState.refs.domainSelect) activeState.refs.domainSelect.value = activeState.domain;
    if (activeState.refs.featureTargetFilterInput) activeState.refs.featureTargetFilterInput.value = activeState.featureTargetFilter || '';

    setBusy(false);
    ProgressDialog.hide();

    if (payload?.ok === false) {
      activeState.refs.emptyGuide.textContent = payload?.message || '검토를 완료하지 못했습니다.';
      activeState.refs.emptyGuide.style.display = 'block';
      activeState.refs.tableWrap.style.display = 'none';
      renderState(activeState);
      return;
    }

    if (!activeState.rows.length) {
      activeState.refs.emptyGuide.textContent = '허용 범위를 초과해 중심축에서 벗어난 탭/분기 피팅이 없습니다.';
    } else {
      activeState.refs.emptyGuide.textContent = `${activeState.rows.length}건의 오류가 확인되었습니다.`;
    }

    saveOpts({
      tol: parseFloat(activeState.refs.tolInput.value || String(DEFAULTS.tol)) || DEFAULTS.tol,
      unit: activeState.unit,
      domain: activeState.domain,
      featureTargetFilter: activeState.featureTargetFilter || ''
    });

    renderCommonOptions(activeState);
    renderState(activeState);
  });

  onHost('tapalign:saved', (payload) => {
    if (!activeState) return;

    activeState.exporting = false;
    ProgressDialog.hide();
    renderState(activeState);

    if (payload?.cancelled) return;

    if (payload?.ok === false) {
      toast(payload?.message || '탭/분기 정렬 결과 엑셀 저장에 실패했습니다. 저장 경로 권한과 파일이 열려 있는지 확인해 주세요. 계속 실패하면 메시지를 관리자에게 전달해 주세요.', 'err');
      return;
    }

    showExcelSavedDialog('탭/분기 정렬 결과 엑셀 저장 완료', payload?.path || '', (savedPath) => {
      post('excel:open', { path: savedPath });
    });
  });
}

function onRun(state) {
  if (!state || state.running) return;

  commitOptions(state);
  state.running = true;
  state.hasRun = true;
  state.rows = [];
  state.extraHeaders = [];
  state.resultUnit = state.refs.unitSelect.value || DEFAULTS.unit;
  state.resultDomain = state.refs.domainSelect.value || DEFAULTS.domain;
  state.refs.emptyGuide.textContent = '검토를 준비하는 중입니다.';
  state.refs.emptyGuide.style.display = 'block';
  state.refs.tableWrap.style.display = 'none';
  state.refs.cardResults.style.display = 'block';

  renderState(state);

  setBusy(true, '탭/분기 축 틀어짐을 검토하는 중입니다.');
  ProgressDialog.show('탭/분기 축 틀어짐 검토', '허용 범위와 공통 설정을 반영해 검토합니다.');
  ProgressDialog.update(0, '검토를 준비하는 중입니다.', '');

  post('tapalign:run', {
    tol: parseFloat(state.refs.tolInput.value || String(DEFAULTS.tol)) || DEFAULTS.tol,
    unit: state.refs.unitSelect.value || DEFAULTS.unit,
    domain: state.refs.domainSelect.value || DEFAULTS.domain,
    featureTargetFilter: state.featureTargetFilter || '',
    commonOptions: buildCommonPayload(state.commonOptions)
  });
}

function onExport(state) {
  if (!state || !state.hasRun) {
    toast('내보낼 결과가 없습니다.', 'warn');
    return;
  }

  chooseExcelMode((excelMode) => {
    state.exporting = true;
    state.lastExcelPct = 0;
    renderState(state);
    ProgressDialog.setActions({});
    ProgressDialog.show('탭/분기 축 검토 엑셀 내보내기', '저장 경로를 선택하는 중입니다.');
    ProgressDialog.update(0, '저장 경로를 선택하는 중입니다.', '엑셀 저장 옵션을 준비하는 중입니다.');
    post('tapalign:save-excel', {
      unit: state.resultUnit || state.unit || DEFAULTS.unit,
      excelMode: excelMode || 'fast',
      locale: getLastExcelExportLocale()
    });
  });
}

function handleTapAlignExcelProgress(state, payload) {
  if (!state) return;

  const phase = normalizeExcelPhase(payload?.phase);
  const current = Number(payload?.current) || 0;
  const total = Number(payload?.total) || 0;
  const percent = computeExcelPercent(state, phase, current, total, payload?.phaseProgress);
  const subtitle = buildExcelSubtitle(phase, current, total);
  const detail = payload?.message || '';

  ProgressDialog.show('탭/분기 축 검토 엑셀 내보내기', subtitle);
  ProgressDialog.update(percent, subtitle, detail);

  if (phase === 'ERROR') {
    state.exporting = false;
    renderState(state);
    setTimeout(() => { ProgressDialog.hide(); state.lastExcelPct = 0; }, 260);
  }
}

function normalizeExcelPhase(phase) {
  return String(phase || '').trim().toUpperCase() || 'EXCEL_WRITE';
}

function computeExcelPercent(state, phase, current, total, phaseProgress) {
  const norm = normalizeExcelPhase(phase);
  if (norm === 'DONE') {
    state.lastExcelPct = 100;
    return 100;
  }
  if (norm === 'ERROR') return state.lastExcelPct || 0;

  const completed = EXCEL_PHASE_ORDER.reduce((acc, key) => {
    if (key === norm) return acc;
    return acc + (EXCEL_PHASE_WEIGHT[key] || 0);
  }, 0);
  const weight = EXCEL_PHASE_WEIGHT[norm] || 0;
  const ratio = total > 0 ? Math.max(0, Math.min(1, current / total)) : 0;
  const staged = Math.max(ratio, clamp01(phaseProgress));
  const pct = Math.max(state.lastExcelPct || 0, Math.min(100, (completed + weight * staged) * 100));
  state.lastExcelPct = pct;
  return pct;
}

function clamp01(value) {
  const n = Number(value);
  return Number.isFinite(n) ? Math.max(0, Math.min(1, n)) : 0;
}

function buildExcelSubtitle(phase, current, total) {
  switch (normalizeExcelPhase(phase)) {
    case 'EXCEL_INIT': return '탭/분기 축 검토 엑셀 워크북을 준비하는 중입니다.';
    case 'EXCEL_WRITE': return `탭/분기 축 검토 엑셀 데이터를 작성하는 중입니다. (${current}/${Math.max(total, current || 1)})`;
    case 'EXCEL_SAVE': return '탭/분기 축 검토 엑셀 파일을 저장하는 중입니다.';
    case 'AUTOFIT': return '열 너비를 자동으로 맞추는 중입니다.';
    case 'DONE': return '탭/분기 축 검토 결과 엑셀 저장 완료';
    case 'ERROR': return '탭/분기 축 검토 엑셀 저장 오류';
    default: return '탭/분기 축 검토 엑셀 내보내기를 진행하는 중입니다.';
  }
}

function commitOptions(state) {
  if (!state || !state.refs) return;

  state.unit = state.refs.unitSelect.value || DEFAULTS.unit;
  state.domain = state.refs.domainSelect.value || DEFAULTS.domain;
  state.featureTargetFilter = (state.refs.featureTargetFilterInput.value || '').trim();
  saveOpts({
    tol: parseFloat(state.refs.tolInput.value || String(DEFAULTS.tol)) || DEFAULTS.tol,
    unit: state.unit,
    domain: state.domain,
    featureTargetFilter: state.featureTargetFilter || ''
  });

  renderState(state);
}

function renderState(state) {
  if (!state || !state.refs) return;

  const hasRows = Array.isArray(state.rows) && state.rows.length > 0;
  const displayUnit = hasRows ? (state.resultUnit || state.unit || DEFAULTS.unit) : (state.unit || DEFAULTS.unit);
  const displayDomain = hasRows ? (state.resultDomain || state.domain || DEFAULTS.domain) : (state.domain || DEFAULTS.domain);
  state.refs.runBtn.disabled = !!state.running || !!state.exporting;
  state.refs.exportBtn.disabled = !state.hasRun || !!state.running || !!state.exporting;
  state.refs.badgeIssues.querySelector('.num').textContent = String(state.rows.length || 0);
  state.refs.badgeScope.querySelector('.num').textContent = resolveDomainLabel(displayDomain);

  if (!state.hasRun) {
    state.refs.cardResults.style.display = 'none';
  } else {
    state.refs.cardResults.style.display = 'block';
  }

  const headers = [
    'ElementId',
    'Category',
    'Family',
    'Type',
    'HostId',
    'HostCategory',
    `중심축으로부터 거리 (${displayUnit})`,
    'XY 평면 기준 각도 (deg)',
    ...state.extraHeaders.flatMap((name) => [`${name} (분기 객체)`, `${name} (연결 라인)`]),
    '메시지'
  ];

  state.refs.thead.innerHTML = '';
  state.refs.tbody.innerHTML = '';

  const headRow = document.createElement('tr');
  headers.forEach((headerText) => {
    const th = document.createElement('th');
    th.textContent = headerText;
    headRow.append(th);
  });
  state.refs.thead.append(headRow);

  if (!hasRows) {
    state.refs.tableWrap.style.display = 'none';
    state.refs.emptyGuide.style.display = state.hasRun ? 'block' : 'none';
    return;
  }

  state.rows.forEach((row) => {
    const tr = document.createElement('tr');
    const values = [
      row.ElementId,
      row.Category,
      row.Family,
      row.Type,
      row.HostId,
      row.HostCategory,
      row.DistanceFromCenter,
      row.ModeledAngle,
      ...state.extraHeaders.flatMap((name) => [row[`BranchParam::${name}`], row[`HostParam::${name}`]]),
      row.Message
    ];

    values.forEach((value, index) => {
      const td = document.createElement('td');
      td.textContent = value == null ? '' : String(value);
      if (index === values.length - 1) td.className = 'status-bad';
      tr.append(td);
    });

    tr.setAttribute('aria-label', [row.File, row.HostType].filter(Boolean).join(' / ') || '탭 정렬 결과');
    state.refs.tbody.append(tr);
  });

  state.refs.emptyGuide.style.display = 'none';
  state.refs.tableWrap.style.display = 'block';
}

function loadCommonOptions(state) {
  let stored = null;
  try {
    const raw = localStorage.getItem(COMMON_OPTIONS_KEY);
    if (raw) stored = JSON.parse(raw);
  } catch {
    stored = null;
  }

  if (stored) {
    state.commonLoadedFromLocal = true;
    applyCommonOptions(state, stored);
    return;
  }

  post('commonoptions:get', { source: 'tapalign' });
}

function applyCommonOptions(state, raw) {
  state.commonOptions = normalizeCommonOptions(raw);
  renderCommonOptions(state);
}

function renderCommonOptions(state) {
  if (!state || !state.refs) return;

  const common = normalizeCommonOptions(state.commonOptions);
  state.refs.extraParamsInput.value = common.extraParamsText || '';
  state.refs.targetFilterInput.value = common.targetFilterText || '';
  state.refs.targetFilterInput.setAttribute('aria-label', common.targetFilterText || '검토 대상 필터');
  state.refs.extraParamsInput.setAttribute('aria-label', common.extraParamsText || '추가 추출 파라미터');
  state.refs.featureTargetFilterInput.value = state.featureTargetFilter || '';
  state.refs.featureTargetFilterInput.setAttribute('aria-label', state.featureTargetFilter || '탭 정렬 대상 필터');
  state.refs.includePointXY.checked = !!common.includePointXY;
  state.refs.includeLinearMetrics.checked = !!common.includeLinearMetrics;
  state.refs.excludeEndDummy.checked = !!common.excludeEndDummy;

  state.refs.commonExtra.querySelector('.num').textContent = common.extraParamsText || '-';
  state.refs.commonFilter.querySelector('.num').textContent = common.targetFilterText || '-';
  state.refs.commonExclude.querySelector('.num').textContent = common.excludeEndDummy ? '적용' : '미적용';
}

function normalizeCommonOptions(raw) {
  return {
    extraParamsText: typeof raw?.extraParamsText === 'string' ? raw.extraParamsText : '',
    targetFilterText: typeof raw?.targetFilterText === 'string' ? raw.targetFilterText : '',
    excludeEndDummy: !!raw?.excludeEndDummy,
    includePointXY: !!raw?.includePointXY,
    includeLinearMetrics: !!raw?.includeLinearMetrics
  };
}

function buildCommonPayload(commonOptions) {
  const normalized = normalizeCommonOptions(commonOptions);
  return {
    extraParamsText: normalized.extraParamsText,
    targetFilterText: normalized.targetFilterText,
    excludeEndDummy: normalized.excludeEndDummy,
    includePointXY: normalized.includePointXY,
    includeLinearMetrics: normalized.includeLinearMetrics
  };
}

function loadOpts() {
  try {
    return Object.assign({}, DEFAULTS, JSON.parse(localStorage.getItem(SKEY) || '{}'));
  } catch {
    return Object.assign({}, DEFAULTS);
  }
}

function saveOpts(next) {
  try {
    localStorage.setItem(SKEY, JSON.stringify(next || {}));
  } catch {
  }
}

function resolveDomainLabel(value) {
  const normalized = String(value || DEFAULTS.domain).toLowerCase();
  if (normalized === 'pipe') return '배관';
  if (normalized === 'duct') return '덕트';
  return '배관 + 덕트';
}

function h1(text) {
  const el = document.createElement('div');
  el.className = 'conn-title';
  el.textContent = text;
  return el;
}

function kv(label, inputEl) {
  const wrap = document.createElement('div');
  wrap.className = 'conn-kv';
  wrap.style.minWidth = '0';
  const cap = document.createElement('label');
  cap.textContent = label;
  if (inputEl?.style) {
    inputEl.style.width = '100%';
    inputEl.style.minWidth = '0';
    inputEl.style.maxWidth = '100%';
    inputEl.style.boxSizing = 'border-box';
  }
  wrap.append(cap, inputEl);
  return wrap;
}

function chip(label, numText) {
  const el = document.createElement('span');
  el.className = 'conn-chip';
  const title = document.createElement('span');
  title.textContent = label;
  const num = document.createElement('span');
  num.className = 'num';
  num.textContent = numText;
  el.append(title, num);
  return el;
}

function cardBtn(text, onClick) {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.textContent = text;
  btn.className = 'card-action-btn';
  if (typeof onClick === 'function') btn.addEventListener('click', onClick);
  return btn;
}

function makeNumber(value) {
  const input = document.createElement('input');
  input.type = 'number';
  input.step = '0.01';
  input.value = String(value ?? DEFAULTS.tol);
  return input;
}

function makeText(value, placeholder) {
  const input = document.createElement('input');
  input.type = 'text';
  input.value = String(value || '');
  input.placeholder = placeholder || '';
  input.style.width = '100%';
  input.style.minWidth = '0';
  input.style.boxSizing = 'border-box';
  return input;
}

function makeCheckbox(checked) {
  const input = document.createElement('input');
  input.type = 'checkbox';
  input.checked = !!checked;
  return input;
}

function makeUnit(value) {
  const select = document.createElement('select');
  select.className = 'kkyt-select';
  select.innerHTML = '<option value="mm">mm</option><option value="inch">inch</option>';
  select.value = String(value || DEFAULTS.unit);
  return select;
}

function makeDomain(value) {
  const select = document.createElement('select');
  select.className = 'kkyt-select';
  select.innerHTML = '<option value="all">배관 + 덕트</option><option value="pipe">배관</option><option value="duct">덕트</option>';
  select.value = String(value || DEFAULTS.domain);
  return select;
}
