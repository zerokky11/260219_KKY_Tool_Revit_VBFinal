// Resources/HubUI/js/views/familylink.js
import { clear, div, toast, debounce, showExcelSavedDialog, chooseExcelMode, getLastExcelExportLocale, showCompletionSummaryDialog } from '../core/dom.js';
import { refreshUiAfterHostDialog } from '../core/hostDialog.js';
import { attachRvtDropZone } from '../core/rvtDrop.js';
import { ProgressDialog } from '../core/progress.js';
import { onHost, post } from '../core/bridge.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js?v=20260504a';

const EXCEL_PHASE_WEIGHT = { EXCEL_INIT: 0.05, EXCEL_WRITE: 0.85, EXCEL_SAVE: 0.08, AUTOFIT: 0.02, DONE: 1, ERROR: 1 };
const DEFAULT_SCHEMA = [
  'FileName',
  'HostFamilyName',
  'HostFamilyCategory',
  'NestedFamilyName',
  'NestedTypeName',
  'NestedInstanceId',
  'NestedPath',
  'NestingLevel',
  'NestedCategory',
  'NestedParamName',
  'TargetParamName',
  'ExpectedGuid',
  'NestedParamGuid',
  'Issue',
  'Notes'
];

export function renderFamilyLink(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  const topbarEl = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar');
  if (topbarEl) topbarEl.classList.add('hub-topbar');

  const state = {
    items: [],
    groups: [],
    selectedGroups: new Set(['(All Groups)']),
    selectedParams: new Set(),
    search: '',
    schema: DEFAULT_SCHEMA.slice(),
    rows: [],
    rvtPaths: [],
    rvtChecked: new Set(),
    includeSingleFamilyCheck: false,
    busy: false
  };
  let lastExcelPct = 0;
  let progressHideTimer = null;

  const page = div('familylink-page feature-shell');

  const header = div('feature-header');
  const heading = div('feature-heading');
  heading.innerHTML = `
    <span class="feature-kicker">네스티드 패밀리 연동</span>
    <h2 class="feature-title">패밀리 공유파라미터 연동 검토</h2>
    <p class="feature-sub">공유 GUID 기준으로 네스티드 패밀리 파라미터 연동 상태를 점검합니다.</p>`;

  const runBtn = cardBtn('검토 시작', onRun);
  const exportBtn = cardBtn('엑셀 내보내기', onExport);
  exportBtn.disabled = true;

  const actions = div('feature-actions');
  actions.append(runBtn, exportBtn);
  header.append(heading, actions);
  page.append(header);

  const body = div('familylink-body');
  const topPanels = div('familylink-top-panels');
  const resultsPanel = div('familylink-results-panel');

  // ----- 공유파라미터 선택 (ParamProp 스타일 재사용) -----
  const paramCard = div('paramprop-card section familylink-card');
  const paramTitle = div('paramprop-title');
  paramTitle.textContent = '공유 파라미터 선택';
  paramCard.append(paramTitle);

  const sourceLine = div('familylink-source');
 


  const searchRow = div('paramprop-row paramprop-search-row');
  const searchBox = document.createElement('input');
  searchBox.type = 'search';
  searchBox.placeholder = '이름 또는 그룹 검색';
  searchBox.className = 'paramprop-search';
  searchBox.addEventListener('input', debounce((e) => {
    state.search = (e.target.value || '').trim();
    renderParamTable();
  }, 120));
  searchRow.append(labelSpan('검색'), searchBox);
  paramCard.append(searchRow);

  const selectGrid = div('paramprop-grid');
  paramCard.append(selectGrid);

  const groupBox = div('paramprop-table-box paramprop-group-box');
  const groupHeader = div('paramprop-subtitle');
  groupHeader.textContent = '그룹 (다중 선택)';
  const groupTable = document.createElement('table');
  groupTable.className = 'paramprop-table paramprop-group-table';
  const groupThead = document.createElement('thead');
  groupThead.innerHTML = '<tr><th>선택</th><th>그룹</th></tr>';
  const groupTbody = document.createElement('tbody');
  groupTable.append(groupThead, groupTbody);
  const groupListWrap = div('paramprop-table-wrap paramprop-group-wrap');
  groupListWrap.append(groupTable);
  groupBox.append(groupHeader, groupListWrap);
  selectGrid.append(groupBox);

  const tableBox = div('paramprop-table-box');
  const tableHead = div('paramprop-subtitle');
  const selectedCount = document.createElement('span');
  selectedCount.className = 'familylink-selected-count';
  tableHead.textContent = '파라미터';
  tableHead.append(selectedCount);
  const table = document.createElement('table');
  table.className = 'paramprop-table';
  const thead = document.createElement('thead');
  thead.innerHTML = '<tr><th>선택</th><th>이름</th></tr>';
  const tbody = document.createElement('tbody');
  table.append(thead, tbody);
  const tableWrap = div('paramprop-table-wrap');
  tableWrap.append(table);
  tableBox.append(tableHead, tableWrap);
  selectGrid.append(tableBox);

  const settingsBox = div('familylink-settings-inline');
  const singleFamilyLabel = document.createElement('label');
  singleFamilyLabel.className = 'familylink-check-row';
  const singleFamilyChk = document.createElement('input');
  singleFamilyChk.type = 'checkbox';
  singleFamilyChk.checked = state.includeSingleFamilyCheck;
  singleFamilyChk.addEventListener('change', () => {
    state.includeSingleFamilyCheck = !!singleFamilyChk.checked;
  });
  const singleFamilyCopy = div('familylink-check-copy');
  const singleFamilyTitle = document.createElement('strong');
  singleFamilyTitle.textContent = '단일 패밀리 파라미터 추가 여부 검토';
  const singleFamilyDesc = document.createElement('span');
  singleFamilyDesc.textContent = '중첩 인스턴스가 없는 단일 패밀리에 선택한 공유파라미터가 추가되어 있는지 함께 확인합니다.';
  singleFamilyCopy.append(singleFamilyTitle, singleFamilyDesc);
  singleFamilyLabel.append(singleFamilyChk, singleFamilyCopy);
  settingsBox.append(singleFamilyLabel);
  paramCard.append(settingsBox);

  // ----- RVT 목록 -----
  const rvtCard = div('paramprop-card section familylink-card');
  const rvtTitle = div('paramprop-title');
  rvtTitle.textContent = '대상 RVT 목록';
  rvtCard.append(rvtTitle);

  const rvtActions = div('familylink-rvt-actions');
  const btnAddRvt = cardBtn('RVT 파일 추가', () => post('familylink:pick-rvts', {}));
  const btnRemoveRvt = cardBtn('선택 제거', () => {
    if (!state.rvtChecked.size) return;
    state.rvtPaths = state.rvtPaths.filter(p => !state.rvtChecked.has(p));
    state.rvtChecked.clear();
    renderRvtList();
    syncRunState();
  });
  const btnClearRvt = cardBtn('등록 목록 비우기', () => {
    state.rvtPaths = [];
    state.rvtChecked.clear();
    renderRvtList();
    syncRunState();
  });
  rvtActions.append(btnAddRvt, btnRemoveRvt, btnClearRvt);
  rvtCard.append(rvtActions);
  const rvtHint = div('rvt-drop-hint');
  rvtHint.textContent = 'RVT 파일 추가 버튼 또는 탐색기 드래그 앤 드롭으로 여러 .rvt를 바로 등록할 수 있습니다.';
  rvtCard.append(rvtHint);

  const rvtTableWrap = div('familylink-rvt-table rvt-drop-zone');
  const { table: rvtTable, tbody: rvtTbody, master: rvtMaster } = createRvtTable();
  rvtTableWrap.append(rvtTable);
  rvtCard.append(rvtTableWrap);
  attachRvtDropZone(rvtTableWrap, {
    onDropPaths: (paths) => {
      const added = appendDroppedRvts(paths);
      if (!added) {
        toast('이미 등록된 RVT입니다.', 'warn');
        return;
      }
      renderRvtList();
      syncRunState();
      toast(`${added}개 RVT를 추가했습니다.`, 'ok');
    },
    onInvalid: () => toast('RVT 파일만 드래그해서 추가할 수 있습니다.', 'warn')
  });

  // ----- 결과 영역 -----
  const resultHead = div('familylink-results-head');
  const resultTitle = div('familylink-results-title');
  resultTitle.textContent = '검토 결과';
  const resultMeta = div('familylink-results-meta');
  resultMeta.textContent = '결과 0개';
  resultHead.append(resultTitle, resultMeta);

  const resultBody = div('familylink-results-body');
  const resultTable = document.createElement('table');
  resultTable.className = 'familylink-table';
  const resultThead = document.createElement('thead');
  const resultTbody = document.createElement('tbody');
  resultTable.append(resultThead, resultTbody);
  resultBody.append(resultTable);

  resultsPanel.append(resultHead, resultBody);

  topPanels.append(paramCard, rvtCard);
  body.append(topPanels, resultsPanel);
  page.append(body);
  target.append(page);

  renderGroups();
  renderParamTable();
  renderRvtList();
  renderResultTable();
  syncRunState();

  onHost('familylink:sharedparams', handleSharedParams);
  onHost('familylink:rvts-picked', handleRvtsPicked);
  onHost('familylink:progress', handleProgress);
  onHost('familylink:result', handleResult);
  onHost('familylink:error', handleError);
  onHost('familylink:exported', handleExported);

  post('familylink:init', {});

  function handleSharedParams(payload) {
    state.items = Array.isArray(payload?.items) ? payload.items : [];
    state.groups = deriveGroups(state.items);
    state.selectedGroups = new Set(['(All Groups)']);
    state.selectedParams = new Set();
    state.search = '';
    searchBox.value = '';

   

    renderGroups();
    renderParamTable();
    syncRunState();
  }

  function handleRvtsPicked(payload) {
    const paths = Array.isArray(payload?.paths) ? payload.paths : [];
    appendDroppedRvts(paths);
    refreshUiAfterHostDialog(() => {
      renderRvtList();
      syncRunState();
    });
  }

  function appendDroppedRvts(paths) {
    const existing = new Set(state.rvtPaths.map((path) => String(path || '').toLowerCase()));
    let added = 0;
    (Array.isArray(paths) ? paths : []).forEach((path) => {
      if (!path) return;
      const key = String(path).toLowerCase();
      if (!existing.has(key)) {
        existing.add(key);
        state.rvtPaths.push(path);
        added += 1;
      }
      state.rvtChecked.add(path);
    });
    return added;
  }

  function handleProgress(payload) {
    if (!payload) return;
    if (payload.phase || payload.current != null || payload.total != null) {
      handleExcelProgress(payload);
      return;
    }
    clearProgressHideTimer();
    const pct = Math.max(0, Math.min(100, Number(payload.percent) || 0));
    const msg = payload.detail || payload.message || '';
    const subtitle = payload.stage || buildRunProgressSubtitle(pct, msg);
    ProgressDialog.show('패밀리 연동 검토', subtitle);
    ProgressDialog.update(pct, subtitle, buildRunProgressDetail(pct, msg));
    if (pct >= 100) {
      scheduleProgressHide(350);
    }
  }

  function handleExcelProgress(payload) {
    const phase = normalizeExcelPhase(payload?.phase);
    const total = Number(payload?.total) || 0;
    const current = Number(payload?.current) || 0;
    const percent = computeExcelPercent(phase, current, total, payload?.phaseProgress, payload?.percent);
    const subtitle = buildExcelSubtitle(phase, current, total);
    const detail = formatExcelDetail(phase, payload?.message);
    const exporting = phase !== 'DONE' && phase !== 'ERROR';

    if (!state.busy && exporting) setBusy(true);
    clearProgressHideTimer();

    ProgressDialog.show('엑셀 내보내기', subtitle || '엑셀 내보내기 진행 중');
    ProgressDialog.update(percent, subtitle, detail);

    if (!exporting) {
      scheduleProgressHide(260, () => {
        lastExcelPct = 0;
        setBusy(false);
      });
    }
  }

  function handleResult(payload) {
    clearProgressHideTimer();
    state.rows = Array.isArray(payload?.rows) ? payload.rows : [];
    state.schema = Array.isArray(payload?.schema) && payload.schema.length ? payload.schema : DEFAULT_SCHEMA.slice();
    renderResultTable();
    exportBtn.disabled = state.rows.length === 0 || state.busy;
    setBusy(false);
    ProgressDialog.hide();
    requestAnimationFrame(() => showFamilyLinkCompletionDialog());
  }

  function handleError(payload) {
    clearProgressHideTimer();
    setBusy(false);
    ProgressDialog.hide();
    lastExcelPct = 0;
    const message = payload?.message || '패밀리 적합성 검토 중 오류가 발생했습니다. 현재 모델과 기준 엑셀 설정을 확인한 뒤 다시 실행해 주세요. 계속 실패하면 메시지를 관리자에게 전달해 주세요.';
    toast(message, 'err', 3200);
  }

  function handleExported(payload) {
    clearProgressHideTimer();
    setBusy(false);
    ProgressDialog.hide();
    lastExcelPct = 0;
    const ok = payload?.ok !== false && payload?.path;
    exportBtn.disabled = state.rows.length === 0 || state.busy;
    if (ok) {
      requestAnimationFrame(() => {
        showExcelSavedDialog('패밀리 적합성 결과 엑셀 저장 완료', payload.path, (p) => post('excel:open', { path: p }));
      });
    } else {
      toast(payload?.message || '패밀리 적합성 결과 엑셀 내보내기에 실패했습니다. 저장 경로 권한과 파일이 열려 있는지 확인해 주세요. 계속 실패하면 메시지를 관리자에게 전달해 주세요.', 'err');
    }
  }

  function clearProgressHideTimer() {
    if (progressHideTimer) {
      clearTimeout(progressHideTimer);
      progressHideTimer = null;
    }
  }

  function scheduleProgressHide(delay = 260, afterHide) {
    clearProgressHideTimer();
    progressHideTimer = setTimeout(() => {
      progressHideTimer = null;
      ProgressDialog.hide();
      if (typeof afterHide === 'function') afterHide();
    }, delay);
  }

  function onRun() {
    if (state.busy) return;
    if (!state.rvtPaths.length) {
      toast('검토할 RVT 파일을 추가해 주세요.', 'warn');
      return;
    }
    const selectedRvts = getCheckedRvtPaths();
    if (!selectedRvts.length) {
      toast('검토할 RVT를 1개 이상 선택해 주세요.', 'warn');
      return;
    }

    const targets = Array.from(state.selectedParams)
      .map(guid => state.items.find(item => item.guid === guid))
      .filter(Boolean)
      .map(item => ({ name: item.name, guid: item.guid }));

    if (!targets.length) {
      toast('검토할 파라미터를 선택해 주세요.', 'warn');
      return;
    }

    setBusy(true);
    exportBtn.disabled = true;
    ProgressDialog.show('패밀리 연동 검토', '검토 구성을 준비하는 중입니다.');
    ProgressDialog.update(0, '검토 구성을 준비하는 중입니다.', '선택한 RVT와 대상 파라미터를 정리하는 중입니다.');

    post('familylink:run', {
      rvtPaths: selectedRvts,
      targets,
      includeSingleFamilyCheck: !!state.includeSingleFamilyCheck
    });
  }

  function onExport() {
    if (state.busy || !state.rows.length) return;
    chooseExcelMode((mode) => {
      exportBtn.disabled = true;
      const selected = mode || 'fast';
      lastExcelPct = 0;
      setBusy(true);
      ProgressDialog.show('엑셀 내보내기', '엑셀 내보내기를 준비하는 중입니다.');
      ProgressDialog.update(0, '엑셀 내보내기를 준비하는 중입니다.', '결과 행과 저장 옵션을 정리하는 중입니다.');
      post('familylink:export', {
        fastExport: selected === 'fast',
        autoFit: selected === 'normal',
        locale: getLastExcelExportLocale()
      });
    });
  }

  function setBusy(on) {
    state.busy = on;
    runBtn.disabled = on;
    runBtn.textContent = on ? '검토 중…' : '검토 시작';
    syncRunState();
  }

  function syncRunState() {
    const hasTargets = state.selectedParams.size > 0;
    const hasRvts = getCheckedRvtPaths().length > 0;
    runBtn.disabled = state.busy || !(hasTargets && hasRvts);
    exportBtn.disabled = state.busy || state.rows.length === 0;
    btnRemoveRvt.disabled = state.rvtChecked.size === 0;
  }

  function getCheckedRvtPaths() {
    return state.rvtPaths.filter((path) => state.rvtChecked.has(path));
  }

  function renderGroups() {
    groupTbody.innerHTML = '';
    const allItem = makeGroupItem('(All Groups)');
    groupTbody.append(allItem);
    state.groups.forEach(g => groupTbody.append(makeGroupItem(g)));
  }

  function makeGroupItem(name) {
    const tr = document.createElement('tr');
    tr.className = state.selectedGroups.has(name) ? 'is-selected' : '';
    tr.dataset.group = name;
    const chk = document.createElement('input');
    chk.type = 'checkbox';
    chk.checked = state.selectedGroups.has(name);
    chk.addEventListener('change', (e) => { e.stopPropagation(); toggleGroup(name, chk.checked); });
    const tdChk = document.createElement('td');
    tdChk.append(chk);
    const nameCell = td(formatGroupDisplayName(name));
    nameCell.setAttribute('aria-label', formatGroupDisplayName(name));
    tr.append(tdChk, nameCell);
    tr.addEventListener('click', () => toggleGroup(name, !state.selectedGroups.has(name)));
    return tr;
  }

  function toggleGroup(name, on) {
    if (name === '(All Groups)') {
      state.selectedGroups = on ? new Set(['(All Groups)']) : new Set();
    } else {
      const sg = new Set(state.selectedGroups);
      sg.delete('(All Groups)');
      if (on) sg.add(name); else sg.delete(name);
      if (sg.size === 0) sg.add('(All Groups)');
      state.selectedGroups = sg;
    }
    renderGroups();
    renderParamTable();
  }

  function renderParamTable() {
    tbody.innerHTML = '';
    updateSelectedCount();

    const filtered = filterDefs();
    if (!filtered.length) {
      const tr = document.createElement('tr');
      const tdEmpty = document.createElement('td');
      tdEmpty.colSpan = 2;
      tdEmpty.textContent = state.items.length
        ? '조건에 맞는 항목이 없습니다.'
        : '공유파라미터 등록이 필요합니다. Revit에서 공유파라미터 TXT를 등록하거나 연결한 뒤 다시 시도해 주세요.';
      tdEmpty.className = 'paramprop-empty';
      tr.append(tdEmpty);
      tbody.append(tr);
      return;
    }

    filtered.forEach(def => {
      const key = def.guid;
      const tr = document.createElement('tr');
      tr.dataset.key = key;
      tr.dataset.group = def.groupName || '';
      tr.className = state.selectedParams.has(key) ? 'is-selected' : '';
      const tdChk = document.createElement('td');
      const chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.checked = state.selectedParams.has(key);
      chk.addEventListener('change', (e) => {
        e.stopPropagation();
        if (chk.checked) state.selectedParams.add(key); else state.selectedParams.delete(key);
        renderParamTable();
        syncRunState();
      });
      tdChk.append(chk);
      const nameCell = td(def.name);
      nameCell.setAttribute('aria-label', `${def.name || ''} / ${def.groupName || ''} / ${def.dataTypeToken || ''}`.trim());
      tr.append(tdChk, nameCell);
      tr.addEventListener('click', () => {
        if (state.selectedParams.has(key)) state.selectedParams.delete(key); else state.selectedParams.add(key);
        renderParamTable();
        syncRunState();
      });
      tbody.append(tr);
    });
  }

  function updateSelectedCount() {
    const count = state.selectedParams.size;
    selectedCount.textContent = count ? ` (선택 ${count}개)` : '';
  }

  function filterDefs() {
    const groups = state.selectedGroups;
    const search = state.search.toLowerCase();
    return state.items.filter(d => {
      const inGroup = groups.has('(All Groups)') || groups.has(d.groupName);
      if (!inGroup) return false;
      if (!search) return true;
      return (d.name || '').toLowerCase().includes(search) || (d.groupName || '').toLowerCase().includes(search);
    });
  }

  function deriveGroups(items) {
    const set = new Set();
    (Array.isArray(items) ? items : []).forEach(d => {
      if (d?.groupName) set.add(d.groupName);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'ko'));
  }

  function normalizeExcelPhase(phase) {
    return String(phase || '').trim().toUpperCase() || 'EXCEL_WRITE';
  }

  function clamp01(value) {
    const n = Number(value);
    return Number.isFinite(n) ? Math.max(0, Math.min(1, n)) : 0;
  }

  function computeExcelPercent(phase, current, total, phaseProgress, percentOverride) {
    const norm = normalizeExcelPhase(phase);
    if (norm === 'DONE') {
      lastExcelPct = 100;
      return 100;
    }
    if (norm === 'ERROR') return lastExcelPct;

    if (typeof percentOverride === 'number' && Number.isFinite(percentOverride) && percentOverride > 0 && percentOverride <= 1) {
      lastExcelPct = Math.max(lastExcelPct, percentOverride * 100);
      return lastExcelPct;
    }

    const completed = ['EXCEL_INIT', 'EXCEL_WRITE', 'EXCEL_SAVE', 'AUTOFIT'].reduce((acc, key) => {
      if (key === norm) return acc;
      return acc + (EXCEL_PHASE_WEIGHT[key] || 0);
    }, 0);
    const weight = EXCEL_PHASE_WEIGHT[norm] || 0;
    const ratio = total > 0 ? Math.max(0, Math.min(1, current / total)) : 0;
    const staged = Math.max(ratio, clamp01(phaseProgress));
    const pct = (completed + weight * staged) * 100;
    lastExcelPct = Math.max(lastExcelPct, Math.min(100, pct));
    return lastExcelPct;
  }

  function buildExcelSubtitle(phase, current, total) {
    const norm = normalizeExcelPhase(phase);
    switch (norm) {
      case 'EXCEL_INIT': return '엑셀 워크북을 준비하는 중입니다.';
      case 'EXCEL_WRITE': return `엑셀 데이터를 작성하는 중입니다. (${current}/${Math.max(total, current || 1)})`;
      case 'EXCEL_SAVE': return '엑셀을 저장하는 중입니다.';
      case 'AUTOFIT': return '열 너비를 자동 조정하는 중입니다.';
      case 'DONE': return '패밀리 적합성 결과 엑셀 내보내기 완료';
      case 'ERROR': return '패밀리 적합성 결과 엑셀 내보내기 오류';
      default: return '패밀리 적합성 결과 엑셀 내보내기를 진행하는 중입니다.';
    }
  }

  function formatExcelDetail(phase, message) {
    if (message) return message;
    return normalizeExcelPhase(phase) === 'DONE' ? '패밀리 적합성 결과 엑셀 내보내기 완료' : '';
  }

  function buildRunProgressSubtitle(percent, message) {
    const raw = String(message || '').trim();
    if (!raw) return '패밀리 연동 검토를 진행하는 중입니다.';
    if (raw.includes('프로젝트 스캔 시작')) return '문서를 준비하는 중입니다.';
    if (raw.includes('패밀리 검사 중')) return `패밀리 검사 중입니다. (${formatRunPercent(percent)})`;
    if (raw.includes('완료')) return '패밀리 연동 검토 완료';
    return `패밀리 연동 검토를 진행하는 중입니다. (${formatRunPercent(percent)})`;
  }

  function buildRunProgressDetail(percent, message) {
    const raw = String(message || '').trim();
    if (raw) return raw;
    return `전체 진행률 ${formatRunPercent(percent)}`;
  }

  function formatRunPercent(percent) {
    const safe = Number(percent);
    if (!Number.isFinite(safe)) return '0%';
    return `${Math.max(0, Math.min(100, Math.round(safe * 10) / 10))}%`;
  }

  function renderRvtList() {
    const rows = state.rvtPaths.map((path, idx) => ({
      index: idx + 1,
      name: getRvtName(path, '—'),
      path,
      checked: state.rvtChecked.has(path),
      onToggle: (checked) => {
        if (checked) state.rvtChecked.add(path);
        else state.rvtChecked.delete(path);
        syncRunState();
      }
    }));

    renderRvtRows(rvtTbody, rows, '등록된 RVT가 없습니다.');
    rvtMaster.checked = rows.length > 0 && rows.every(r => r.checked);
    rvtMaster.indeterminate = rows.some(r => r.checked) && !rvtMaster.checked;
    rvtMaster.onchange = () => {
      state.rvtChecked.clear();
      if (rvtMaster.checked) rows.forEach(r => state.rvtChecked.add(r.path));
      renderRvtList();
      syncRunState();
    };
  }

  function renderResultTable() {
    resultThead.innerHTML = '';
    resultTbody.innerHTML = '';

    const headRow = document.createElement('tr');
    state.schema.forEach(h => {
      const th = document.createElement('th');
      th.textContent = h;
      headRow.append(th);
    });
    resultThead.append(headRow);

    if (!state.rows.length) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = state.schema.length;
      td.className = 'familylink-empty';
      td.textContent = '결과가 없습니다. 검토를 실행해 주세요.';
      tr.append(td);
      resultTbody.append(tr);
    } else {
      state.rows.forEach(row => {
        const tr = document.createElement('tr');
        state.schema.forEach(h => {
          const td = document.createElement('td');
          const val = row?.[h];
          td.textContent = val == null ? '' : String(val);
          tr.append(td);
        });
        resultTbody.append(tr);
      });
    }

    resultMeta.textContent = `결과 ${state.rows.length}개`;
  }

  function showFamilyLinkCompletionDialog() {
    const totalRows = state.rows.length || 0;
    const issueCount = state.rows.filter((row) => isIssueRow(row)).length;
    const okCount = Math.max(totalRows - issueCount, 0);
    const fileCount = countUniqueValues(state.rows, 'FileName');
    const hostFamilyCount = countUniqueValues(state.rows, 'HostFamilyName');

    showCompletionSummaryDialog({
      title: '패밀리 연동 검토 완료',
      message: '네스티드 패밀리 공유파라미터 연동 검토가 끝났습니다. 아래 표에서 상세를 확인하거나 바로 엑셀로 내보낼 수 있습니다.',
      summaryItems: [
        { label: '검토 결과 행', value: String(totalRows) },
        { label: '이슈 행', value: String(issueCount) },
        { label: '정상 행', value: String(okCount) },
        { label: '대상 RVT', value: `${fileCount}개` },
        { label: '호스트 패밀리', value: `${hostFamilyCount}개` }
      ],
      exportDisabled: !!exportBtn.disabled,
      onExport: () => exportBtn.click()
    });
  }

  function isIssueRow(row) {
    const issue = String(row?.Issue || '').trim();
    if (!issue) return false;
    return !/^(ok|정상)$/i.test(issue);
  }

  function countUniqueValues(rows, key) {
    const values = new Set();
    (Array.isArray(rows) ? rows : []).forEach((row) => {
      const value = String(row?.[key] ?? '').trim();
      if (value) values.add(value);
    });
    return values.size;
  }
}

function cardBtn(text, onClick) {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'btn card-btn';
  btn.textContent = text;
  btn.onclick = onClick;
  return btn;
}

function labelSpan(text) {
  const span = document.createElement('span');
  span.className = 'paramprop-label';
  span.textContent = text;
  return span;
}

function td(value) {
  const cell = document.createElement('td');
  cell.textContent = value ?? '';
  return cell;
}

function formatGroupDisplayName(name) {
  return name === '(All Groups)' ? '전체 그룹' : (name ?? '');
}
