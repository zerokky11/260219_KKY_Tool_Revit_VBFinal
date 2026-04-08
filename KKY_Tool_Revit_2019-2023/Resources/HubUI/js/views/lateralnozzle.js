import { clear, div, toast, setBusy } from '../core/dom.js';
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
        <span class="feature-kicker">Utility · Excel Workflow</span>
        <h2 class="feature-title">노즐코드 KTA 단일화</h2>
        <p class="feature-sub">접수받은 KTA 양식을 정해진 하나의 시트양식으로 추출합니다.</p>
      </div>
    </div>
  `;

  const layout = div('deliverycleaner-top-grid');
  const leftColumn = div('deliverycleaner-field-stack');
  leftColumn.append(buildFileCard(state));
  layout.append(leftColumn, buildRunCard(state));
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
    if (added > 0) toast(`${added}개 엑셀 파일을 추가했습니다.`, 'ok');
  });
  onHost('lateralnozzle:progress', (payload) => {
    if (!state.acceptProgress || !state.busy) return;
    ProgressDialog.setActions({});
    ProgressDialog.show(payload?.title || '노즐코드 KTA 단일화', payload?.message || '');
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
      fileCount: Number(payload?.fileCount) || 0
    };
    renderRunSummary(state);
    renderResultSummary(state);
    updateActionState(state);
    toast(payload?.message || '노즐코드 KTA 단일화가 완료되었습니다.', payload?.ok === false ? 'err' : 'ok', 3200);
  });
  onHost('lateralnozzle:error', (payload) => {
    state.acceptProgress = false;
    ProgressDialog.hide();
    setPageBusy(state, false);
    renderRunSummary(state);
    renderResultSummary(state);
    updateActionState(state);
    toast(payload?.message || '오류가 발생했습니다.', 'err', 3600);
  });

  renderExcelTable(state);
  renderRunSummary(state);
  renderResultSummary(state);
  updateActionState(state);
  post('lateralnozzle:init', {});
}

function buildFileCard(state) {
  const card = div('deliverycleaner-card lateralnozzle-card');
  card.innerHTML = `
    <div class="deliverycleaner-card__head">
      <div>
        <h3>엑셀 파일 등록</h3>
        <p>xlsx/xls 파일을 추가하고 드래그해서 목록에 바로 등록할 수 있습니다.</p>
      </div>
    </div>
  `;

  const actions = div('deliverycleaner-inline-actions multi-action-card__actions');
  const addBtn = actionButton('엑셀 추가', () => post('lateralnozzle:pick-excels', {}), 'primary');
  const removeBtn = actionButton('선택 제거', () => {
    state.excelPaths = state.excelPaths.filter((path) => !state.checked.has(path));
    state.checked = new Set(state.excelPaths);
    renderExcelTable(state);
    renderRunSummary(state);
    renderResultSummary(state);
    updateActionState(state);
  });
  const clearBtn = actionButton('목록 지우기', () => {
    state.excelPaths = [];
    state.checked.clear();
    renderExcelTable(state);
    renderRunSummary(state);
    renderResultSummary(state);
    updateActionState(state);
  });
  actions.append(addBtn, removeBtn, clearBtn);

  const hint = div('rvt-drop-hint');
  hint.textContent = '탐색기에서 .xlsx 또는 .xls 파일을 이 영역으로 끌어오면 바로 추가됩니다. 올바른 엑셀 파일은 잘못된 형식 경고가 뜨지 않도록 별도 드롭 경로로 처리합니다.';

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
        <th>파일명</th>
        <th>파일 경로</th>
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
      if (added > 0) {
        toast(`${added}개 엑셀 파일을 추가했습니다.`, 'ok');
      } else {
        toast('이미 등록된 엑셀 파일입니다.', 'warn');
      }
    },
    onInvalid: (payload) => {
      const message = payload?.message || '엑셀 파일(.xlsx, .xls)만 추가할 수 있습니다.';
      toast(message, 'warn', 2600);
    }
  });

  const note = div('deliverycleaner-note');
  note.textContent = '각 파일의 모든 시트를 검사해 UT명 / 배관No / 연결호기 블록을 찾고, 결과는 하나의 시트 양식으로 정리합니다.';

  card.append(actions, hint, tableWrap, note);

  state.ui.addBtn = addBtn;
  state.ui.removeBtn = removeBtn;
  state.ui.clearBtn = clearBtn;
  state.ui.tableBody = table.querySelector('tbody');
  state.ui.masterCheck = table.querySelector('thead input[type="checkbox"]');
  return card;
}

function buildRunCard(state) {
  const card = div('deliverycleaner-card lateralnozzle-card lateralnozzle-card--run');
  card.innerHTML = `
    <div class="deliverycleaner-card__head">
      <div>
        <h3>실행 및 결과</h3>
        <p>추출 결과는 바탕화면에 엑셀 파일로 저장되고, 누락/형식 오류는 비고 컬럼에 표시됩니다.</p>
      </div>
    </div>
  `;

  const actionRow = div('deliverycleaner-inline-actions multi-action-card__actions');
  const runBtn = actionButton('추출 시작', () => runExtraction(state), 'primary');
  const openResultBtn = actionButton('결과 파일 열기', () => {
    if (!state.lastResult?.resultWorkbookPath) return;
    post('excel:open', { path: state.lastResult.resultWorkbookPath });
  });
  const openFolderBtn = actionButton('결과 폴더 열기', () => {
    if (!state.lastResult?.outputFolder) return;
    post('lateralnozzle:open-folder', { path: state.lastResult.outputFolder });
  });
  actionRow.append(runBtn, openResultBtn, openFolderBtn);

  const runSummary = div('deliverycleaner-note');
  const validation = div('deliverycleaner-note');
  const resultSummary = div('deliverycleaner-summary-box lateralnozzle-result-box');
  card.append(actionRow, runSummary, validation, resultSummary);

  state.ui.runBtn = runBtn;
  state.ui.openResultBtn = openResultBtn;
  state.ui.openFolderBtn = openFolderBtn;
  state.ui.runSummary = runSummary;
  state.ui.validation = validation;
  state.ui.resultSummary = resultSummary;
  return card;
}

function applyHostState(state, payload) {
  const settings = payload?.settings || {};
  state.excelPaths = Array.isArray(settings.excelPaths) ? [...settings.excelPaths] : [];
  state.checked = new Set(state.excelPaths);
  state.lastResult = payload?.result || null;

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
    td.textContent = '등록된 엑셀 파일이 없습니다.';
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
        updateActionState(state);
      });
      checkCell.append(checkbox);

      const indexCell = document.createElement('td');
      indexCell.style.textAlign = 'center';
      indexCell.textContent = String(index + 1);

      const nameCell = document.createElement('td');
      nameCell.className = 'segmentpms-path-cell';
      nameCell.textContent = getFileName(path);
      nameCell.title = getFileName(path);

      const pathCell = document.createElement('td');
      pathCell.className = 'segmentpms-path-cell';
      pathCell.textContent = path;
      pathCell.title = path;

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
      updateActionState(state);
    };
  }
}

function renderRunSummary(state) {
  if (!state.ui.runSummary || !state.ui.validation) return;

  const selectedCount = getCheckedExcelPaths(state).length;
  state.ui.runSummary.textContent = `등록 파일 ${state.excelPaths.length}개 / 실행 대상 ${selectedCount}개`;
  state.ui.validation.textContent = '검사 규칙: 1) UTILITY / LATERAL NO / Nozzle Code 중 하나라도 비면 비고에 누락 항목 표시 2) Nozzle Code는 반드시 "_000" 형태의 숫자 3자리로 끝나야 하며, 아니면 비고에 형식 불일치 표시';
}

function renderResultSummary(state) {
  if (!state.ui.resultSummary) return;

  const result = state.lastResult;
  if (!result) {
    state.ui.resultSummary.innerHTML = `
      <strong>최근 결과 없음</strong>
      <span>실행이 끝나면 추출 건수와 비고 건수, 결과 파일 경로를 여기에 표시합니다.</span>
    `;
    return;
  }

  const summary = result.summary || {};
  state.ui.resultSummary.innerHTML = `
    <strong>${result.message || '처리가 완료되었습니다.'}</strong>
    <span>처리 파일 ${Number(result.fileCount) || 0}개 / 추출 ${Number(summary.extractedRowCount) || 0}건 / 비고 ${Number(summary.remarkRowCount) || 0}건</span>
    <span class="lateralnozzle-path" title="${result.resultWorkbookPath || ''}">${result.resultWorkbookPath || '결과 파일 경로 없음'}</span>
  `;
}

function updateActionState(state) {
  const hasFiles = state.excelPaths.length > 0;
  const hasChecked = getCheckedExcelPaths(state).length > 0;
  if (state.ui.addBtn) state.ui.addBtn.disabled = !!state.busy;
  if (state.ui.removeBtn) state.ui.removeBtn.disabled = !hasChecked || !!state.busy;
  if (state.ui.clearBtn) state.ui.clearBtn.disabled = !hasFiles || !!state.busy;
  if (state.ui.runBtn) state.ui.runBtn.disabled = !hasChecked || !!state.busy;
  if (state.ui.openResultBtn) state.ui.openResultBtn.disabled = !state.lastResult?.resultWorkbookPath || !!state.busy;
  if (state.ui.openFolderBtn) state.ui.openFolderBtn.disabled = !state.lastResult?.outputFolder || !!state.busy;
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
  ProgressDialog.update(0, '엑셀 파일을 읽는 중입니다.', '');
  post('lateralnozzle:run', {
    excelPaths: paths,
    outputFolder: ''
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
