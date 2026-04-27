import { clear, div, toast, setBusy, showExcelSavedDialog, chooseExcelMode, getLastExcelExportLocale, showCompletionSummaryDialog } from '../core/dom.js';
import { refreshUiAfterHostDialog } from '../core/hostDialog.js';
import { attachRvtDropZone } from '../core/rvtDrop.js';
import { attachExcelDropZone } from '../core/excelDrop.js';
import { ProgressDialog } from '../core/progress.js';
import { onHost, post } from '../core/bridge.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js';

const EXCEL_PHASE_WEIGHT = {
  EXCEL_INIT: 0.05,
  EXCEL_WRITE: 0.85,
  EXCEL_SAVE: 0.08,
  AUTOFIT: 0.02,
  DONE: 1,
  ERROR: 1
};

const DEFAULT_SCHEMA = [
  'HostFileName',
  'HostFilePath',
  'ReferenceElementId',
  'LinkName',
  'LinkFileName',
  'TypeWorksetNames',
  'InstanceWorksetNames',
  'ApplyTypeWorksetNames',
  'ApplyInstanceWorksetNames',
  'CurrentLinkPath',
  'StoredLinkPath',
  'CurrentPathType',
  'TargetLinkPath',
  'TargetPathType',
  'ApplyStatus',
  'ApplyMessage'
];

const VISIBLE_COLUMNS = [
  ['HostFileName', '호스트 파일'],
  ['LinkName', '링크 이름'],
  ['TypeWorksetNames', '타입 웍셋'],
  ['InstanceWorksetNames', '인스턴스 웍셋'],
  ['ApplyTypeWorksetNames', '적용 타입 웍셋'],
  ['ApplyInstanceWorksetNames', '적용 인스턴스 웍셋'],
  ['CurrentLinkPath', '현재 경로'],
  ['TargetLinkPath', '대상 경로'],
  ['ApplyStatus', '상태'],
  ['ApplyMessage', '메시지']
];

export function renderLinkPath(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  const topbarEl = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar');
  if (topbarEl) topbarEl.classList.add('hub-topbar');

  const state = {
    rvtPaths: [],
    rvtChecked: new Set(),
    rows: [],
    schema: DEFAULT_SCHEMA.slice(),
    workbookPath: '',
    newLinkPlacement: 'origin',
    busy: false,
    lastAction: '',
    actionStartedAt: 0
  };

  let lastExcelPct = 0;
  let progressHideTimer = null;

  const page = div('feature-shell familylink-page');
  const header = div('feature-header');
  const heading = div('feature-heading');
  heading.innerHTML = `
    <span class="feature-kicker">Revit Link Path</span>
    <h2 class="feature-title">Revit 링크 경로 추출 및 변경</h2>
    <p class="feature-sub">RVT에서 링크 현황을 추출하고, 엑셀 기준으로 Reload From 또는 신규 Revit 링크 생성을 일괄 반영합니다.</p>`;
  header.append(heading);
  page.append(header);

  const body = div('familylink-body');
  const topPanels = div('familylink-top-panels');

  const rvtSection = div('paramprop-card section familylink-card');
  const rvtHeader = div('familylink-results-head');
  const rvtTitle = div('paramprop-title');
  rvtTitle.textContent = '1단계 · RVT 등록 및 링크 추출';
  const rvtMeta = div('familylink-results-meta');
  rvtMeta.textContent = '0 files';
  rvtHeader.append(rvtTitle, rvtMeta);
  rvtSection.append(rvtHeader);

  const rvtActions = div('familylink-rvt-actions');
  const rvtManageActions = div('feature-actions');
  const rvtRunActions = div('feature-actions');
  rvtRunActions.style.marginLeft = '16px';

  const btnAdd = cardBtn('RVT 파일 추가', () => post('linkpath:pick-rvts', {}), 'btn--primary');
  const btnRemove = cardBtn('선택 제거', () => {
    if (!state.rvtChecked.size) return;
    state.rvtPaths = state.rvtPaths.filter((path) => !state.rvtChecked.has(path));
    state.rvtChecked.clear();
    renderRvtList();
    syncButtons();
  }, 'btn--secondary');
  const btnClear = cardBtn('등록 목록 비우기', () => {
    state.rvtPaths = [];
    state.rvtChecked.clear();
    renderRvtList();
    syncButtons();
  }, 'btn--secondary');
  const btnExtract = cardBtn('링크 추출', onExtract, 'btn--primary');
  const btnExport = cardBtn('엑셀 내보내기', onExport, 'btn--secondary');

  rvtManageActions.append(btnAdd, btnRemove, btnClear);
  rvtRunActions.append(btnExtract, btnExport);
  rvtActions.append(rvtManageActions, rvtRunActions);
  rvtSection.append(rvtActions);

  const rvtHint = div('rvt-drop-hint');
  rvtHint.textContent = 'RVT 파일 추가 버튼을 누르거나 탐색기에서 여러 .rvt 파일을 바로 끌어 놓을 수 있습니다.';
  rvtSection.append(rvtHint);

  const rvtTableWrap = div('familylink-rvt-table rvt-drop-zone');
  const { table: rvtTable, tbody: rvtTbody, master: rvtMaster } = createRvtTable();
  rvtTableWrap.append(rvtTable);
  rvtSection.append(rvtTableWrap);

  attachRvtDropZone(rvtTableWrap, {
    onDropPaths: (paths) => {
      const added = appendRvts(paths);
      if (!added) {
        toast('이미 등록된 RVT입니다.', 'warn');
        return;
      }
      renderRvtList();
      syncButtons();
      toast(`${added}개 RVT를 추가했습니다.`, 'ok');
    },
    onInvalid: () => toast('RVT 파일만 끌어 놓아 추가할 수 있습니다.', 'warn')
  });

  const applySection = div('paramprop-card section familylink-card');
  const applyHeader = div('familylink-results-head');
  const applyTitle = div('paramprop-title');
  applyTitle.textContent = '2단계 · 링크 수정 / 신규 생성 반영';
  const applyMeta = div('familylink-results-meta');
  applyMeta.textContent = '미선택';
  applyHeader.append(applyTitle, applyMeta);
  applySection.append(applyHeader);

  const applyActions = div('familylink-rvt-actions');
  const applyActionLeft = div('feature-actions');
  const applyActionRight = div('feature-actions');
  applyActionRight.style.marginLeft = '16px';

  const btnPickExcel = cardBtn('엑셀 선택', () => post('linkpath:pick-excel', {}), 'btn--secondary');
  const btnApply = cardBtn('엑셀 기준 적용', onApply, 'btn--primary');
  applyActionLeft.append(btnPickExcel);
  applyActionRight.append(btnApply);
  applyActions.append(applyActionLeft, applyActionRight);
  applySection.append(applyActions);

  const newLinkSettings = div('feature-row__summary');
  newLinkSettings.style.display = 'grid';
  newLinkSettings.style.gap = '8px';
  newLinkSettings.style.border = '1px solid var(--border-subtle)';
  newLinkSettings.style.background = 'var(--surface-elevated)';
  newLinkSettings.style.borderRadius = '8px';
  newLinkSettings.style.padding = '12px 14px';

  const placementLabel = document.createElement('label');
  placementLabel.textContent = '신규 링크 배치 방식';
  placementLabel.style.fontWeight = '700';

  const placementSelect = document.createElement('select');
  placementSelect.className = 'feature-input';
  [
    ['origin', '자동 - 원점 대 원점'],
    ['centered', '자동 - 중심 대 중심'],
    ['shared', '자동 - 공유 좌표 기준'],
    ['site', '자동 - 프로젝트 기준점']
  ].forEach(([value, label]) => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = label;
    placementSelect.append(option);
  });
  placementSelect.value = state.newLinkPlacement;
  placementSelect.addEventListener('change', () => {
    state.newLinkPlacement = placementSelect.value || 'origin';
  });

  const placementHint = div('feature-note');
  placementHint.textContent = '엑셀에서 LinkName과 ReferenceElementId를 비우고 HostFileName, TargetLinkPath, 적용 웍셋을 채운 행은 새 링크로 생성합니다.';
  newLinkSettings.append(placementLabel, placementSelect, placementHint);
  applySection.append(newLinkSettings);

  const excelDrop = div('feature-row__summary');
  excelDrop.style.minHeight = '88px';
  excelDrop.style.display = 'grid';
  excelDrop.style.gap = '8px';
  excelDrop.style.alignContent = 'center';
  excelDrop.style.border = '1px dashed var(--border-strong)';
  excelDrop.style.background = 'var(--surface-elevated)';
  excelDrop.style.borderRadius = '8px';
  excelDrop.style.padding = '14px';

  const excelLead = document.createElement('strong');
  excelLead.textContent = '대상 파일 경로를 적은 엑셀을 선택하면 자동으로 불러옵니다.';
  const excelPath = div('feature-note');
  excelPath.textContent = '선택된 엑셀 파일 없음';
  excelDrop.append(excelLead, excelPath);
  applySection.append(excelDrop);

  attachExcelDropZone(excelDrop, {
    onDropPaths: (paths) => {
      const path = Array.isArray(paths) && paths.length ? paths[0] : '';
      if (!path) return;
      state.workbookPath = path;
      renderWorkbookState();
      onImport();
    },
    onInvalid: () => toast('엑셀 파일만 끌어 놓아 추가할 수 있습니다.', 'warn')
  });

  const resultSection = div('familylink-results-panel');
  const resultHead = div('familylink-results-head');
  const resultTitle = div('familylink-results-title');
  resultTitle.textContent = '링크 현황';
  const resultMeta = div('familylink-results-meta');
  resultMeta.textContent = '0 rows';
  resultHead.append(resultTitle, resultMeta);
  resultSection.append(resultHead);

  const summaryNote = div('feature-note');
  summaryNote.textContent = '링크 추출 후 현재 경로와 대상 경로를 여기에서 확인할 수 있습니다.';
  resultSection.append(summaryNote);

  const resultWrap = div('familylink-results-body');
  const resultTable = document.createElement('table');
  resultTable.className = 'familylink-table';
  const resultThead = document.createElement('thead');
  const resultTbody = document.createElement('tbody');
  resultTable.append(resultThead, resultTbody);
  resultWrap.append(resultTable);
  resultSection.append(resultWrap);

  topPanels.append(rvtSection, applySection);
  body.append(topPanels, resultSection);
  page.append(body);
  target.append(page);

  renderRvtList();
  renderWorkbookState();
  renderResultTable();
  syncButtons();

  onHost('linkpath:rvts-picked', (payload) => {
    const paths = Array.isArray(payload?.paths) ? payload.paths : [];
    if (!paths.length) return;
    appendRvts(paths);
    refreshUiAfterHostDialog(() => {
      renderRvtList();
      syncButtons();
    });
  });

  onHost('linkpath:excel-picked', (payload) => {
    state.workbookPath = String(payload?.path || '').trim();
    refreshUiAfterHostDialog(() => {
      renderWorkbookState();
      syncButtons();
      onImport();
    });
  });

  onHost('linkpath:rows', (payload) => {
    clearProgressHideTimer();
    state.rows = Array.isArray(payload?.rows) ? payload.rows : [];
    state.schema = Array.isArray(payload?.schema) && payload.schema.length ? payload.schema : DEFAULT_SCHEMA.slice();
    state.workbookPath = String(payload?.workbookPath || state.workbookPath || '').trim();
    renderWorkbookState();
    renderResultTable(payload?.summary || null);
    const source = String(payload?.source || '').toLowerCase();
    const summary = payload?.summary || buildSummary(state.rows);
    finishAction(source, summary);
  });

  onHost('linkpath:progress', (payload) => {
    handleProgress(payload || {});
  });

  onHost('linkpath:exported', (payload) => {
    clearProgressHideTimer();
    setBusyState(false);
    ProgressDialog.hide();
    lastExcelPct = 0;
    if (payload?.path) {
      state.workbookPath = payload.path;
      renderWorkbookState();
      requestAnimationFrame(() => {
        showExcelSavedDialog('링크 현황 엑셀을 저장했습니다.', payload.path, (path) => post('excel:open', { path }));
      });
    } else {
      toast(payload?.message || '엑셀 저장에 실패했습니다.', 'err');
    }
    syncButtons();
  });

  onHost('linkpath:applied', (payload) => {
    if (!payload?.ok) return;
    const changed = Number(payload?.summary?.changedCount || 0);
    const deleted = Number(payload?.summary?.deletedCount || 0);
    const errors = Number(payload?.summary?.errorCount || 0);
    toast(`변경 ${changed}건, 삭제 ${deleted}건, 오류 ${errors}건`, errors > 0 ? 'warn' : 'ok');
  });

  onHost('linkpath:error', (payload) => {
    clearProgressHideTimer();
    setBusyState(false);
    ProgressDialog.hide();
    lastExcelPct = 0;
    syncButtons();
    toast(payload?.message || '작업 중 오류가 발생했습니다.', 'err');
  });

  function appendRvts(paths) {
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

  function renderRvtList() {
    rvtMeta.textContent = `${state.rvtPaths.length} files`;
    const rows = state.rvtPaths.map((path, idx) => ({
      index: idx + 1,
      name: getRvtName(path, `RVT ${idx + 1}`),
      path,
      checked: state.rvtChecked.has(path),
      onToggle: (checked) => {
        if (checked) state.rvtChecked.add(path);
        else state.rvtChecked.delete(path);
        syncMasterCheckbox();
      }
    }));
    renderRvtRows(rvtTbody, rows, '등록된 RVT가 없습니다.');
    rvtMaster.checked = rows.length > 0 && rows.every((row) => row.checked);
    rvtMaster.indeterminate = rows.some((row) => row.checked) && !rvtMaster.checked;
  }

  rvtMaster.addEventListener('change', () => {
    if (rvtMaster.checked) state.rvtPaths.forEach((path) => state.rvtChecked.add(path));
    else state.rvtChecked.clear();
    renderRvtList();
    syncButtons();
  });

  function syncMasterCheckbox() {
    const total = state.rvtPaths.length;
    const checked = Array.from(state.rvtChecked).filter((path) => state.rvtPaths.includes(path)).length;
    rvtMaster.checked = total > 0 && checked === total;
    rvtMaster.indeterminate = checked > 0 && checked < total;
  }

  function renderWorkbookState() {
    const path = state.workbookPath || '';
    applyMeta.textContent = path ? '선택됨' : '미선택';
    excelPath.textContent = path || '선택된 엑셀 파일 없음';
    excelPath.title = path || '';
  }

  function renderResultTable(summaryPayload = null) {
    const rows = Array.isArray(state.rows) ? state.rows : [];
    const summaryData = summaryPayload || buildSummary(rows);
    resultMeta.textContent = `${rows.length} rows`;
    summaryNote.textContent =
      `호스트 ${summaryData.hostCount}개, 링크 ${summaryData.rowCount}건, 대상 지정 ${summaryData.targetCount}건, 삭제 후보 ${summaryData.deleteCandidateCount || 0}건, 삭제 ${summaryData.deletedCount || 0}건, 변경 ${summaryData.changedCount}건, 오류 ${summaryData.errorCount}건`;

    resultThead.innerHTML = '';
    const headRow = document.createElement('tr');
    VISIBLE_COLUMNS.forEach(([, label]) => {
      const th = document.createElement('th');
      th.textContent = label;
      headRow.append(th);
    });
    resultThead.append(headRow);

    resultTbody.innerHTML = '';
    if (!rows.length) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = VISIBLE_COLUMNS.length;
      td.className = 'empty-cell';
      td.textContent = '추출된 링크가 없습니다.';
      tr.append(td);
      resultTbody.append(tr);
      return;
    }

    rows.forEach((row) => {
      const tr = document.createElement('tr');
      VISIBLE_COLUMNS.forEach(([key]) => {
        const td = document.createElement('td');
        const text = row && row[key] != null ? String(row[key]) : '';
        td.textContent = text || '-';
        td.title = text || '';
        if (key.toLowerCase().includes('path') || key.toLowerCase().includes('message')) {
          td.className = 'segmentpms-path-cell';
        }
        if (key === 'ApplyStatus') {
          const status = text.toLowerCase();
          if (status === 'changed') td.style.color = 'var(--ok)';
          else if (status === 'error') td.style.color = 'var(--err)';
          else if (status === 'skip' || status === 'info') td.style.color = 'var(--muted)';
        }
        tr.append(td);
      });
      resultTbody.append(tr);
    });
  }

  function buildSummary(rows) {
    const safeRows = Array.isArray(rows) ? rows : [];
    const hostSet = new Set();
    let targetCount = 0;
    let deleteCandidateCount = 0;
    let deletedCount = 0;
    let changedCount = 0;
    let errorCount = 0;
    let skipCount = 0;

    safeRows.forEach((row) => {
      const host = String(row?.HostFilePath || '').trim();
      if (host) hostSet.add(host.toLowerCase());
      const target = String(row?.TargetLinkPath || '').trim();
      if (target) targetCount += 1;
      if (isDeleteCandidate(row)) deleteCandidateCount += 1;
      const status = String(row?.ApplyStatus || '').trim().toLowerCase();
      const message = String(row?.ApplyMessage || '');
      if (status === 'changed' && message.includes('삭제')) deletedCount += 1;
      if (status === 'changed') changedCount += 1;
      else if (status === 'error') errorCount += 1;
      else if (status === 'skip' || status === 'info') skipCount += 1;
    });

    return {
      hostCount: hostSet.size,
      rowCount: safeRows.length,
      targetCount,
      deleteCandidateCount,
      deletedCount,
      changedCount,
      errorCount,
      skipCount
    };
  }

  function isDeleteCandidate(row) {
    const refId = String(row?.ReferenceElementId || '').trim();
    const target = String(row?.TargetLinkPath || '').trim();
    return !!refId && !target;
  }

  function onExtract() {
    if (!state.rvtPaths.length) {
      toast('RVT를 먼저 추가해 주세요.', 'warn');
      return;
    }
    state.lastAction = 'extract';
    state.actionStartedAt = Date.now();
    setBusyState(true);
    lastExcelPct = 0;
    ProgressDialog.show('링크 추출', '링크 정보를 읽는 중입니다.');
    ProgressDialog.update(5, '링크 추출', '등록된 RVT에서 Revit 링크를 확인하고 있습니다.');
    post('linkpath:extract', { rvtPaths: state.rvtPaths });
  }

  function onExport() {
    if (!state.rows.length) {
      toast('먼저 링크를 추출해 주세요.', 'warn');
      return;
    }
    state.lastAction = 'export';
    state.actionStartedAt = Date.now();
    chooseExcelMode((mode) => {
      lastExcelPct = 0;
      setBusyState(true);
      post('linkpath:export', { excelMode: mode || 'fast', locale: getLastExcelExportLocale() });
    });
  }

  function onImport() {
    if (!state.workbookPath) {
      toast('엑셀 파일을 먼저 선택해 주세요.', 'warn');
      return;
    }
    state.lastAction = 'import';
    state.actionStartedAt = Date.now();
    setBusyState(true);
    ProgressDialog.show('엑셀 불러오기', 'TargetLinkPath를 읽는 중입니다.');
    ProgressDialog.update(20, '엑셀 불러오기', state.workbookPath);
    post('linkpath:import', { path: state.workbookPath });
  }

  function onApply() {
    const summary = buildSummary(state.rows);
    const applyCount = Number(summary.targetCount || 0);
    const deleteCount = Number(summary.deleteCandidateCount || 0);
    if (applyCount + deleteCount <= 0) {
      toast('엑셀에서 대상 경로를 채우거나 삭제할 기존 링크의 TargetLinkPath를 비운 뒤 다시 불러와 주세요.', 'warn');
      return;
    }

    const confirmLines = [
      `TargetLinkPath 입력 행 ${applyCount}건은 링크 경로 변경 또는 신규 링크 생성으로 처리됩니다.`,
      `TargetLinkPath가 빈 기존 링크 행 ${deleteCount}건은 링크가 삭제됩니다.`,
      '계속 적용할까요?'
    ];
    const confirmed = window.confirm(confirmLines.join('\n'));
    if (!confirmed) return;

    state.lastAction = 'apply';
    state.actionStartedAt = Date.now();
    setBusyState(true);
    lastExcelPct = 0;
    ProgressDialog.show('엑셀 기준 적용', '등록된 RVT에 링크 경로 변경/신규 생성/삭제를 적용하는 중입니다.');
    ProgressDialog.update(5, '엑셀 기준 적용', '호스트 문서를 열고 링크를 다시 불러오거나 새로 생성/삭제하고 있습니다.');
    post('linkpath:apply', { newLinkPlacement: state.newLinkPlacement, rvtPaths: state.rvtPaths });
  }

  function finishAction(source, summary) {
    const action = source || state.lastAction || '';
    const elapsed = Date.now() - Number(state.actionStartedAt || 0);
    const minVisibleMs = action === 'extract' || action === 'apply' ? 500 : 250;
    const waitMs = Math.max(0, minVisibleMs - elapsed);

    scheduleProgressHide(waitMs, () => {
      setBusyState(false);
      if (action === 'import') {
        toast('엑셀을 불러왔습니다.', 'ok');
        return;
      }
      if (action === 'apply') {
        toast('엑셀 기준 링크 경로 적용이 끝났습니다.', 'ok');
      }
      if (action === 'extract' || action === 'apply') {
        requestAnimationFrame(() => showLinkPathCompletionDialog(action, summary));
      }
    });
  }

  function showLinkPathCompletionDialog(action, summary) {
    const safeSummary = summary || buildSummary(state.rows);
    const isExtract = action === 'extract';
    const title = isExtract ? '링크 추출 완료' : '링크 반영 완료';
    const message = isExtract
      ? '링크 현황을 읽어왔습니다. 결과를 확인하고 엑셀로 내보낼 수 있습니다.'
      : '엑셀에 적은 대상 경로 기준으로 링크 반영이 완료되었습니다.';
    const notes = [];

    if (isExtract) {
      notes.push('현재 결과는 아래 링크 현황 표와 엑셀 내보내기에서 같은 기준으로 확인할 수 있습니다.');
      notes.push('TargetLinkPath를 입력한 기존 링크 행은 Reload From으로 반영됩니다.');
      notes.push('LinkName과 ReferenceElementId가 빈 행은 선택한 배치 방식으로 신규 Revit 링크를 생성합니다.');
      notes.push('기존 링크 행에서 TargetLinkPath를 비우면 해당 링크를 삭제합니다.');
    } else {
      notes.push('센트럴 파일은 로컬로 열어 웍셋을 닫고 반영 후 동기화합니다.');
      notes.push('일반 파일은 웍셋을 닫고 반영 후 저장합니다.');
      if (safeSummary.deletedCount > 0) {
        notes.push(`삭제 ${safeSummary.deletedCount}건이 포함되었습니다.`);
      }
    }

    if (safeSummary.errorCount > 0) {
      notes.push(`오류 ${safeSummary.errorCount}건은 결과 표의 메시지 열에서 바로 확인할 수 있습니다.`);
    }

    showCompletionSummaryDialog({
      title,
      message,
      summaryItems: [
        { label: '호스트 파일', value: `${safeSummary.hostCount || 0}개` },
        { label: '링크 수', value: `${safeSummary.rowCount || 0}건` },
        { label: '대상 지정', value: `${safeSummary.targetCount || 0}건` },
        { label: isExtract ? '삭제 후보' : '삭제', value: `${isExtract ? (safeSummary.deleteCandidateCount || 0) : (safeSummary.deletedCount || 0)}건` },
        { label: isExtract ? '오류' : '변경', value: `${isExtract ? (safeSummary.errorCount || 0) : (safeSummary.changedCount || 0)}건` },
        { label: '건너뜀', value: `${safeSummary.skipCount || 0}건` }
      ],
      notes,
      confirmLabel: '닫기',
      showExport: isExtract,
      exportLabel: '엑셀 내보내기',
      exportDisabled: !state.rows.length,
      onExport: isExtract ? () => onExport() : null
    });
  }

  function handleProgress(payload) {
    if (!payload) return;
    if (payload.phase || payload.current != null || payload.total != null) {
      handleExcelProgress(payload);
      return;
    }

    clearProgressHideTimer();
    const pct = Math.max(0, Math.min(100, Number(payload.percent) || 0));
    const subtitle = payload.stage || payload.message || '링크 작업 진행 중';
    const detail = payload.detail || payload.message || '';
    ProgressDialog.show('링크 작업', subtitle);
    ProgressDialog.update(pct, subtitle, detail);
    if (pct >= 100) scheduleProgressHide(300);
  }

  function handleExcelProgress(payload) {
    const phase = normalizeExcelPhase(payload?.phase);
    const total = Number(payload?.total) || 0;
    const current = Number(payload?.current) || 0;
    const percent = computeExcelPercent(phase, current, total, payload?.phaseProgress, payload?.percent);
    const subtitle = buildExcelSubtitle(phase, current, total);
    const detail = formatExcelDetail(phase, payload?.message);
    const exporting = phase !== 'DONE' && phase !== 'ERROR';

    if (!state.busy && exporting) setBusyState(true);
    clearProgressHideTimer();

    ProgressDialog.show('엑셀 내보내기', subtitle || '엑셀 내보내기 진행 중');
    ProgressDialog.update(percent, subtitle, detail);

    if (!exporting) {
      scheduleProgressHide(260, () => {
        lastExcelPct = 0;
        setBusyState(false);
      });
    }
  }

  function normalizeExcelPhase(phase) {
    return String(phase || '').trim().toUpperCase() || 'EXCEL_WRITE';
  }

  function computeExcelPercent(phase, current, total, phaseProgress, percentOverride) {
    const norm = normalizeExcelPhase(phase);
    if (norm === 'DONE') {
      lastExcelPct = 100;
      return 100;
    }
    if (Number.isFinite(Number(percentOverride))) {
      const pct = Math.max(lastExcelPct, Math.min(100, Number(percentOverride) * 100));
      lastExcelPct = pct;
      return pct;
    }
    if (norm === 'ERROR') return lastExcelPct;

    const completed = ['EXCEL_INIT', 'EXCEL_WRITE', 'EXCEL_SAVE', 'AUTOFIT'].reduce((acc, key) => {
      if (key === norm) return acc;
      return acc + (EXCEL_PHASE_WEIGHT[key] || 0);
    }, 0);
    const weight = EXCEL_PHASE_WEIGHT[norm] || 0;
    const staged = Number.isFinite(Number(phaseProgress))
      ? Math.max(0, Math.min(1, Number(phaseProgress)))
      : (total > 0 ? Math.max(0, Math.min(1, current / total)) : 0);
    const pct = Math.min(100, Math.max(lastExcelPct, (completed + weight * staged) * 100));
    lastExcelPct = pct;
    return pct;
  }

  function buildExcelSubtitle(phase, current, total) {
    switch (normalizeExcelPhase(phase)) {
      case 'EXCEL_INIT':
        return '엑셀 준비 중';
      case 'EXCEL_WRITE':
        return `엑셀 데이터 작성 중 (${current}/${Math.max(total, current || 1)})`;
      case 'EXCEL_SAVE':
        return '엑셀 저장 중';
      case 'AUTOFIT':
        return '열 너비 조정 중';
      case 'DONE':
        return '엑셀 저장 완료';
      case 'ERROR':
        return '엑셀 저장 실패';
      default:
        return '엑셀 작업 진행 중';
    }
  }

  function formatExcelDetail(phase, message) {
    const text = String(message || '').trim();
    if (text) return text;
    return normalizeExcelPhase(phase) === 'DONE' ? '링크 현황 엑셀 저장이 완료되었습니다.' : '';
  }

  function scheduleProgressHide(delay, onDone) {
    clearProgressHideTimer();
    progressHideTimer = window.setTimeout(() => {
      progressHideTimer = null;
      ProgressDialog.hide();
      if (typeof onDone === 'function') onDone();
    }, delay);
  }

  function clearProgressHideTimer() {
    if (!progressHideTimer) return;
    window.clearTimeout(progressHideTimer);
    progressHideTimer = null;
  }

  function setBusyState(nextBusy) {
    state.busy = !!nextBusy;
    setBusy(state.busy);
    syncButtons();
  }

  function syncButtons() {
    const hasRows = state.rows.length > 0;
    const summary = buildSummary(state.rows);
    const hasApplyCandidates = Number(summary.targetCount || 0) + Number(summary.deleteCandidateCount || 0) > 0;
    btnExtract.disabled = state.busy || state.rvtPaths.length === 0;
    btnExport.disabled = state.busy || !hasRows;
    btnApply.disabled = state.busy || !hasApplyCandidates;
    btnPickExcel.disabled = state.busy;
    placementSelect.disabled = state.busy;
    btnAdd.disabled = state.busy;
    btnRemove.disabled = state.busy || state.rvtChecked.size === 0;
    btnClear.disabled = state.busy || state.rvtPaths.length === 0;
  }
}

function cardBtn(label, onClick, variant = 'btn--primary') {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = `btn ${variant}`;
  btn.textContent = label;
  btn.addEventListener('click', onClick);
  return btn;
}
