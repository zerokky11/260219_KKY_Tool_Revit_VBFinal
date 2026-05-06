import { clear, div, toast, setBusy, showCompletionSummaryDialog, showExcelSavedDialog, chooseExcelMode, getLastExcelExportLocale } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { post, onHost } from '../core/bridge.js';
import { attachExcelDropZone } from '../core/excelDrop.js';

export function renderLateralNozzle(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  const state = {
    excelPaths: [],
    checked: new Set(),
    lastResult: null,
    busy: false,
    acceptProgress: false,
    ui: {}
  };

  const page = div('feature-shell deliverycleaner-page lateralnozzle-page');
  page.innerHTML = `
    <div class="feature-header deliverycleaner-header">
      <div class="feature-heading">
        <span class="feature-kicker">유틸리티 · 엑셀 워크플로</span>
        <h2 class="feature-title">\ub178\uc990\ucf54\ub4dc KTA \ub2e8\uc77c\ud654</h2>
        <p class="feature-sub">\uc811\uc218\ubc1b\uc740 KTA \uc591\uc2dd\uc744 \uc815\ud574\uc9c4 \ud558\ub098\uc758 \uc2dc\ud2b8\uc591\uc2dd\uc73c\ub85c \ucd94\ucd9c\ud569\ub2c8\ub2e4.</p>
      </div>
    </div>
  `;

  const layout = div('deliverycleaner-top-grid lateralnozzle-layout');
  layout.append(buildControlCard(state), buildFileCard(state));
  page.append(layout);
  target.append(page);

  onHost('lateralnozzle:init', (payload) => applyHostState(state, payload));
  onHost('lateralnozzle:excels-picked', (payload) => {
    const paths = Array.isArray(payload?.paths) ? payload.paths : [];
    if (!paths.length) return;

    const added = appendExcelPaths(state, paths);
    renderExcelTable(state);
    renderRunSummary(state);
    renderResultSummary(state);
    updateActionState(state);

    if (added > 0) toast(`${added}\uac1c \uc5d1\uc140 \ud30c\uc77c\uc744 \ucd94\uac00\ud588\uc2b5\ub2c8\ub2e4.`, 'ok');
  });
  onHost('lateralnozzle:progress', (payload) => {
    if (!state.acceptProgress || !state.busy) return;
    ProgressDialog.setActions({});
    ProgressDialog.show(payload?.title || '\ub178\uc990\ucf54\ub4dc KTA \ub2e8\uc77c\ud654', payload?.message || '');
    ProgressDialog.update(Number(payload?.percent) || 0, payload?.message || '', payload?.detail || '');
  });
  onHost('lateralnozzle:done', (payload) => {
    state.acceptProgress = false;
    ProgressDialog.hide();
    setPageBusy(state, false);

    state.lastResult = {
      ok: payload?.ok !== false,
      message: payload?.message || '',
      outputFolder: payload?.outputFolder || '',
      resultWorkbookPath: payload?.resultWorkbookPath || '',
      summary: payload?.summary || {},
      fileCount: Number(payload?.fileCount) || 0,
      canExport: payload?.canExport === true
    };

    renderRunSummary(state);
    renderResultSummary(state);
    updateActionState(state);
    toast(payload?.message || '\ub178\uc990\ucf54\ub4dc KTA \ub2e8\uc77c\ud654\uac00 \uc644\ub8cc\ub418\uc5c8\uc2b5\ub2c8\ub2e4.', payload?.ok === false ? 'err' : 'ok', 3200);
    showLateralNozzleResultDialog(state, payload || {});
  });
  onHost('lateralnozzle:exported', (payload) => {
    ProgressDialog.hide();
    setPageBusy(state, false);
    if (payload?.cancelled) return;
    if (payload?.ok === false) {
      toast(payload?.message || '노즐코드 KTA 단일화 결과 엑셀 저장에 실패했습니다. 저장 경로 권한과 파일이 열려 있는지 확인해 주세요. 계속 실패하면 메시지를 관리자에게 전달해 주세요.', 'err', 3600);
      return;
    }
    if (!payload?.path) return;

    if (state.lastResult) state.lastResult.resultWorkbookPath = payload.path;
    window.setTimeout(() => {
      showExcelSavedDialog('노즐코드 KTA 단일화 엑셀을 저장했습니다.', payload.path, (path) => post('excel:open', { path }));
    }, 120);
  });
  onHost('lateralnozzle:error', (payload) => {
    state.acceptProgress = false;
    ProgressDialog.hide();
    setPageBusy(state, false);
    renderRunSummary(state);
    renderResultSummary(state);
    updateActionState(state);
    toast(payload?.message || '노즐코드 KTA 단일화 중 오류가 발생했습니다. 선택한 엑셀 파일과 KTA 양식 상태를 확인한 뒤 다시 실행해 주세요. 계속 실패하면 메시지를 관리자에게 전달해 주세요.', 'err', 3600);
  });

  renderExcelTable(state);
  renderRunSummary(state);
  renderResultSummary(state);
  updateActionState(state);
  post('lateralnozzle:init', {});
}

function buildControlCard(state) {
  const card = div('deliverycleaner-card lateralnozzle-card lateralnozzle-card--control');
  card.innerHTML = `
    <div class="deliverycleaner-card__head">
      <div>
        <h3>\uc5d1\uc140 \ub4f1\ub85d \ubc0f \uc2e4\ud589</h3>
        <p>\ud30c\uc77c \ucd94\uac00\uc640 \ucd94\ucd9c \uc2e4\ud589\uc744 \ud55c \uc601\uc5ed\uc5d0\uc11c \ubc14\ub85c \uc9c4\ud589\ud569\ub2c8\ub2e4.</p>
      </div>
    </div>
  `;

  const actions = div('deliverycleaner-inline-actions lateralnozzle-actions');
  const addBtn = actionButton('\uc5d1\uc140 \ucd94\uac00', () => post('lateralnozzle:pick-excels', {}), 'primary');
  const runBtn = actionButton('\ucd94\ucd9c \uc2dc\uc791', () => runExtraction(state), 'primary');
  const removeBtn = actionButton('\uc120\ud0dd \uc81c\uac70', () => {
    state.excelPaths = state.excelPaths.filter((path) => !state.checked.has(path));
    state.checked = new Set(state.excelPaths);
    renderExcelTable(state);
    renderRunSummary(state);
    renderResultSummary(state);
    updateActionState(state);
  });
  const clearBtn = actionButton('\ubaa9\ub85d \uc9c0\uc6b0\uae30', () => {
    state.excelPaths = [];
    state.checked.clear();
    renderExcelTable(state);
    renderRunSummary(state);
    renderResultSummary(state);
    updateActionState(state);
  });
  actions.append(addBtn, runBtn, removeBtn, clearBtn);

  const runSummary = div('deliverycleaner-summary-box lateralnozzle-summary-box');
  const validation = div('deliverycleaner-note lateralnozzle-rule-note');
  const resultSummary = div('deliverycleaner-summary-box lateralnozzle-result-box');

  card.append(actions, runSummary, validation, resultSummary);

  state.ui.addBtn = addBtn;
  state.ui.runBtn = runBtn;
  state.ui.removeBtn = removeBtn;
  state.ui.clearBtn = clearBtn;
  state.ui.runSummary = runSummary;
  state.ui.validation = validation;
  state.ui.resultSummary = resultSummary;
  return card;
}

function buildFileCard(state) {
  const card = div('deliverycleaner-card lateralnozzle-card lateralnozzle-card--list');
  card.innerHTML = `
    <div class="deliverycleaner-card__head">
      <div>
        <h3>\uc5d1\uc140 \ud30c\uc77c \ubaa9\ub85d</h3>
        <p>\ud30c\uc77c\uba85\uacfc \uacbd\ub85c\ub97c \ub113\uac8c \ud655\uc778\ud558\uba74\uc11c \uc120\ud0dd \uc0c1\ud0dc\ub97c \uad00\ub9ac\ud569\ub2c8\ub2e4.</p>
      </div>
    </div>
  `;

  const hint = div('rvt-drop-hint lateralnozzle-drop-hint');
  hint.textContent = '\ud0d0\uc0c9\uae30\uc5d0\uc11c .xlsx \ub610\ub294 .xls \ud30c\uc77c\uc744 \uc774 \uc601\uc5ed\uc73c\ub85c \ub04c\uc5b4\uc624\uba74 \ubc14\ub85c \ucd94\uac00\ub429\ub2c8\ub2e4. \uc62c\ubc14\ub978 \uc5d1\uc140 \ud30c\uc77c\uc740 \uc798\ubabb\ub41c \ud615\uc2dd \uacbd\uace0\uac00 \ub728\uc9c0 \uc54a\ub3c4\ub85d \ubcc4\ub3c4 \ub4dc\ub86d \uacbd\ub85c\ub85c \ucc98\ub9ac\ud569\ub2c8\ub2e4.';

  const tableWrap = div('rvt-expand-table rvt-drop-zone lateralnozzle-drop-zone');
  const table = document.createElement('table');
  table.className = 'segmentpms-table rvt-register-table lateralnozzle-table';
  table.innerHTML = `
    <colgroup>
      <col style="width:44px">
      <col style="width:56px">
      <col style="width:260px">
      <col>
    </colgroup>
    <thead>
      <tr>
        <th style="text-align:center;"><input type="checkbox"></th>
        <th style="text-align:center;">#</th>
        <th>\ud30c\uc77c\uba85</th>
        <th>\ud30c\uc77c \uacbd\ub85c</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;
  tableWrap.append(table);

  attachExcelDropZone(tableWrap, {
    onDropPaths: (paths) => {
      const added = appendExcelPaths(state, paths);
      renderExcelTable(state);
      renderRunSummary(state);
      renderResultSummary(state);
      updateActionState(state);

      if (added > 0) toast(`${added}\uac1c \uc5d1\uc140 \ud30c\uc77c\uc744 \ucd94\uac00\ud588\uc2b5\ub2c8\ub2e4.`, 'ok');
      else toast('\uc774\ubbf8 \ub4f1\ub85d\ub41c \uc5d1\uc140 \ud30c\uc77c\uc785\ub2c8\ub2e4.', 'warn');
    },
    onInvalid: (payload) => {
      toast(payload?.message || '\uc5d1\uc140 \ud30c\uc77c(.xlsx, .xls)\ub9cc \ucd94\uac00\ud560 \uc218 \uc788\uc2b5\ub2c8\ub2e4.', 'warn', 2600);
    }
  });

  const note = div('deliverycleaner-note lateralnozzle-file-note');
  note.textContent = '\uac01 \ud30c\uc77c\uc758 \ubaa8\ub4e0 \uc2dc\ud2b8\ub97c \uac80\uc0ac\ud574 UT\uba85 / \ubc30\uad00No / Nozzle Code /  No \ud5e4\ub354 \ube14\ub85d\uc744 \ucc3e\uace0, Nozzle Code\uc640  No\ub97c \uc774\uc5b4 \ud558\ub098\uc758 \uacb0\uacfc \uac12\uc73c\ub85c \uc815\ub9ac\ud569\ub2c8\ub2e4.';

  card.append(hint, tableWrap, note);

  state.ui.tableBody = table.querySelector('tbody');
  state.ui.masterCheck = table.querySelector('thead input[type="checkbox"]');
  return card;
}

function applyHostState(state, payload) {
  const settings = payload?.settings || {};
  state.excelPaths = Array.isArray(settings.excelPaths) ? [...settings.excelPaths] : [];
  state.checked = new Set(state.excelPaths);
  state.lastResult = payload?.result
    ? {
      ...payload.result,
      fileCount: Number(payload?.result?.fileCount) || 0,
      canExport: payload?.result?.canExport === true
    }
    : null;

  renderExcelTable(state);
  renderRunSummary(state);
  renderResultSummary(state);
  updateActionState(state);
}

function renderExcelTable(state) {
  const tbody = state.ui.tableBody;
  if (!tbody) return;

  tbody.innerHTML = '';
  if (!state.excelPaths.length) {
    const tr = document.createElement('tr');
    tr.className = 'empty-row';
    const td = document.createElement('td');
    td.colSpan = 4;
    td.className = 'empty-cell';
    td.textContent = '\ub4f1\ub85d\ub41c \uc5d1\uc140 \ud30c\uc77c\uc774 \uc5c6\uc2b5\ub2c8\ub2e4.';
    tr.append(td);
    tbody.append(tr);
  } else {
    state.excelPaths.forEach((path, index) => {
      const tr = document.createElement('tr');

      const checkCell = document.createElement('td');
      checkCell.style.textAlign = 'center';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = state.checked.has(path);
      checkbox.addEventListener('change', () => {
        if (checkbox.checked) state.checked.add(path);
        else state.checked.delete(path);
        renderExcelTable(state);
        renderRunSummary(state);
        updateActionState(state);
      });
      checkCell.append(checkbox);

      const indexCell = document.createElement('td');
      indexCell.style.textAlign = 'center';
      indexCell.textContent = String(index + 1);

      const nameCell = document.createElement('td');
      nameCell.className = 'segmentpms-path-cell';
      nameCell.textContent = getFileName(path);
      nameCell.setAttribute('aria-label', getFileName(path));

      const pathCell = document.createElement('td');
      pathCell.className = 'segmentpms-path-cell';
      pathCell.textContent = path;
      pathCell.setAttribute('aria-label', path || '-');

      tr.append(checkCell, indexCell, nameCell, pathCell);
      tbody.append(tr);
    });
  }

  if (state.ui.masterCheck) {
    state.ui.masterCheck.disabled = state.excelPaths.length === 0;
    state.ui.masterCheck.checked = state.excelPaths.length > 0 && state.excelPaths.every((path) => state.checked.has(path));
    state.ui.masterCheck.onchange = () => {
      state.checked = state.ui.masterCheck.checked ? new Set(state.excelPaths) : new Set();
      renderExcelTable(state);
      renderRunSummary(state);
      updateActionState(state);
    };
  }
}

function renderRunSummary(state) {
  if (!state.ui.runSummary || !state.ui.validation) return;

  const selectedCount = getCheckedExcelPaths(state).length;
  state.ui.runSummary.innerHTML = `
    <strong>\ub4f1\ub85d \ud30c\uc77c ${state.excelPaths.length}\uac1c</strong>
    <span>\uc2e4\ud589 \ub300\uc0c1 ${selectedCount}\uac1c</span>
  `;
  state.ui.validation.textContent = '\uac80\uc0ac \uaddc\uce59: 1) UTILITY / LATERAL NO / Nozzle Code \uc911 \ud558\ub098\ub77c\ub3c4 \ube44\uba74 \ube44\uace0\uc5d0 \ub204\ub77d \ud56d\ubaa9 \ud45c\uc2dc 2) Nozzle Code\ub294 Nozzle Code +  No \uac12\uc744 "_" \ub85c \uc774\uc5b4 \ub9cc\ub4e4\uba70 3) \ucd5c\uc885 Nozzle Code\ub294 \ubc18\ub4dc\uc2dc "_000" \ud615\ud0dc\uc758 \uc22b\uc790 3\uc790\ub9ac\ub85c \ub05d\ub098\uc57c \ud558\uba70, \uc544\ub2c8\uba74 \ube44\uace0\uc5d0 \ud615\uc2dd \ubd88\uc77c\uce58 \ud45c\uc2dc';
}

function renderResultSummary(state) {
  if (!state.ui.resultSummary) return;

  const result = state.lastResult;
  if (!result) {
    state.ui.resultSummary.innerHTML = `
      <strong>\ucd5c\uadfc \uacb0\uacfc \uc5c6\uc74c</strong>
      <span>\uc2e4\ud589\uc774 \ub05d\ub098\uba74 \ucc98\ub9ac \ud30c\uc77c \uc218\uc640 \ucd94\ucd9c \uac74\uc218, \ube44\uace0 \uac74\uc218\ub97c \uc5ec\uae30\uc5d0 \ud45c\uc2dc\ud569\ub2c8\ub2e4.</span>
    `;
    return;
  }

  const summary = result.summary || {};
  state.ui.resultSummary.innerHTML = `
    <strong>${result.message || '\ucc98\ub9ac\uac00 \uc644\ub8cc\ub418\uc5c8\uc2b5\ub2c8\ub2e4.'}</strong>
    <span>\ucc98\ub9ac \ud30c\uc77c ${Number(result.fileCount) || 0}\uac1c</span>
    <span>\ucd94\ucd9c ${Number(summary.extractedRowCount) || 0}\uac74 / \ube44\uace0 ${Number(summary.remarkRowCount) || 0}\uac74</span>
  `;
}

function updateActionState(state) {
  const hasFiles = state.excelPaths.length > 0;
  const hasChecked = getCheckedExcelPaths(state).length > 0;
  if (state.ui.addBtn) state.ui.addBtn.disabled = !!state.busy;
  if (state.ui.runBtn) state.ui.runBtn.disabled = !hasChecked || !!state.busy;
  if (state.ui.removeBtn) state.ui.removeBtn.disabled = !hasChecked || !!state.busy;
  if (state.ui.clearBtn) state.ui.clearBtn.disabled = !hasFiles || !!state.busy;
}

function runExtraction(state) {
  const paths = getCheckedExcelPaths(state);
  if (!paths.length) {
    toast('실행할 엑셀 파일을 1개 이상 선택해 주세요.', 'warn');
    return;
  }

  state.acceptProgress = true;
  setPageBusy(state, true);
  ProgressDialog.setActions({});
  ProgressDialog.show('노즐코드 KTA 단일화', '엑셀 파일을 읽는 중입니다.');
  ProgressDialog.update(0, '엑셀 파일을 읽는 중입니다.', '선택한 엑셀 파일 목록과 변환 대상을 정리하는 중입니다.');
  post('lateralnozzle:run', {
    excelPaths: paths,
    outputFolder: ''
  });
}

async function promptLateralNozzleExcelExport(state) {
  const excelMode = await chooseExcelMode();
  if (!excelMode) return;

  setPageBusy(state, true);
  ProgressDialog.setActions({});
  ProgressDialog.show('엑셀 내보내기', '엑셀 저장을 준비하는 중입니다.');
  ProgressDialog.update(0, '엑셀 저장을 준비하는 중입니다.', '결과 엑셀 저장 옵션을 정리하는 중입니다.');
  post('lateralnozzle:export', { excelMode: excelMode || 'fast', locale: getLastExcelExportLocale() });
}

function showLateralNozzleResultDialog(state, payload = {}) {
  const summary = payload?.summary || state?.lastResult?.summary || {};
  const fileCount = Number(payload?.fileCount ?? state?.lastResult?.fileCount) || 0;
  const successCount = Number(summary.successCount) || 0;
  const failCount = Number(summary.failCount) || 0;
  const noDataCount = Number(summary.noDataCount) || 0;
  const extractedRowCount = Number(summary.extractedRowCount) || 0;
  const remarkRowCount = Number(summary.remarkRowCount) || 0;

  const notes = [
    `\ucd94\ucd9c \uc5c6\uc74c ${noDataCount}\uac1c`,
    '\uacb0\uacfc \uc2dc\ud2b8\ub294 UTILITY / LATERAL NO / Nozzle Code / \ube44\uace0 \uad6c\uc131\uc73c\ub85c \uc800\uc7a5\ub429\ub2c8\ub2e4.',
    'Nozzle Code \uc5f4\uc740 \uc6d0\ubcf8 \uc591\uc2dd\uc758 Nozzle Code\uc640  No \uac12\uc744 "_" \ub85c \uc774\uc5b4 \uc0dd\uc131\ud569\ub2c8\ub2e4.',
    '\ube44\uace0\uc5d0\ub294 \ud544\uc218 \ud56d\ubaa9 \ub204\ub77d\uacfc Nozzle Code \ud615\uc2dd \ubd88\uc77c\uce58\uac00 \ud45c\uc2dc\ub429\ub2c8\ub2e4.'
  ];

  showCompletionSummaryDialog({
    title: '\ub178\uc990\ucf54\ub4dc KTA \ub2e8\uc77c\ud654 \uc644\ub8cc',
    message: payload?.message || '\ucd94\ucd9c\uc774 \uc644\ub8cc\ub418\uc5c8\uc2b5\ub2c8\ub2e4. \ud544\uc694\ud558\uba74 \uacb0\uacfc \uc5d1\uc140\uc744 \uc800\uc7a5\ud558\uc138\uc694.',
    summaryItems: [
      { label: '\ub300\uc0c1 \ud30c\uc77c', value: `${fileCount}\uac1c` },
      { label: '\uc131\uacf5', value: `${successCount}\uac1c` },
      { label: '\uc2e4\ud328', value: `${failCount}\uac1c` },
      { label: '\ucd94\ucd9c \ud589', value: `${extractedRowCount}\uac74` },
      { label: '\ube44\uace0 \ud589', value: `${remarkRowCount}\uac74` }
    ],
    notes,
    exportLabel: '\ub178\uc990\ucf54\ub4dc KTA \ub2e8\uc77c\ud654 \uc5d1\uc140',
    showExport: payload?.canExport === true || state?.lastResult?.canExport === true,
    onExport: () => promptLateralNozzleExcelExport(state)
  });
}

function appendExcelPaths(state, paths) {
  let added = 0;
  (Array.isArray(paths) ? paths : []).forEach((path) => {
    const normalized = String(path || '').trim();
    if (!normalized) return;
    if (!state.excelPaths.some((item) => samePath(item, normalized))) {
      state.excelPaths.push(normalized);
      added += 1;
    }
    state.checked.add(normalized);
  });
  return added;
}

function getCheckedExcelPaths(state) {
  return state.excelPaths.filter((path) => state.checked.has(path));
}

function samePath(left, right) {
  return String(left || '').toLowerCase() === String(right || '').toLowerCase();
}

function getFileName(path) {
  const parts = String(path || '').split(/[/\\]/);
  return parts[parts.length - 1] || '';
}

function actionButton(label, onClick, variant = 'secondary') {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = `btn ${variant === 'primary' ? 'btn--primary' : 'btn--secondary'}`;
  btn.textContent = label;
  btn.addEventListener('click', onClick);
  return btn;
}

function setPageBusy(state, on) {
  state.busy = !!on;
  if (!state.busy) state.acceptProgress = false;
  setBusy(on, on ? '노즐코드 KTA 단일화 작업을 처리하는 중입니다.' : '');
  updateActionState(state);
}
