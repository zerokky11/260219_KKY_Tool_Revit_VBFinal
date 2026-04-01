import { clear, div, toast, setBusy, showCompletionSummaryDialog } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { refreshUiAfterHostDialog } from '../core/hostDialog.js';
import { attachRvtDropZone } from '../core/rvtDrop.js';
import { post, onHost } from '../core/bridge.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js';

const FILTER_OPERATORS = [
  'Equals', 'NotEquals', 'Contains', 'NotContains', 'BeginsWith', 'NotBeginsWith',
  'EndsWith', 'NotEndsWith', 'Greater', 'GreaterOrEqual', 'Less', 'LessOrEqual',
  'HasValue', 'HasNoValue'
];

export function renderParamModifier(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  const state = {
    rvtPaths: [],
    checked: new Set(),
    synchronizeAfterProcessing: false,
    syncComment: '',
    elementParameterUpdate: createElementUpdateState(),
    lastResult: null,
    activeDocument: null,
    busy: false,
    acceptProgress: false,
    ui: {}
  };

  const page = div('feature-shell deliverycleaner-page');
  page.innerHTML = `
    <div class="feature-header deliverycleaner-header">
      <div class="feature-heading">
        <span class="feature-kicker">UTILITIES · 별도 워크플로우</span>
        <h2 class="feature-title">파라미터 수정기</h2>
        <p class="feature-sub">입력 조건을 기준으로 활성 문서 또는 여러 RVT에 파라미터 값을 일괄 적용합니다.</p>
      </div>
    </div>
  `;

  const layout = div('deliverycleaner-top-grid');
  const leftColumn = div('deliverycleaner-field-stack');
  leftColumn.append(buildRunCard(state));
  layout.append(leftColumn, buildSettingsCard(state));
  page.append(layout, buildRvtModal(state));
  target.append(page);

  onHost('parammodifier:init', (payload) => applyHostState(state, payload));
  onHost('parammodifier:rvts-picked', (payload) => {
    const paths = Array.isArray(payload?.paths) ? payload.paths : [];
    if (!paths.length) return;

    const added = appendDroppedRvts(state, paths);
    refreshUiAfterHostDialog(() => {
      renderRvtModal(state);
      renderRunSummary(state);
      updateActionState(state);
      if (added) toast(`${added}개 RVT를 추가했습니다.`, 'ok');
      openRvtModal(state);
    });
  });
  onHost('parammodifier:progress', (payload) => {
    if (!state.acceptProgress || !state.busy) return;
    ProgressDialog.setActions({});
    ProgressDialog.show(payload?.title || '파라미터 수정기', payload?.message || '');
    ProgressDialog.update(Number(payload?.percent) || 0, payload?.message || '', payload?.detail || '');
  });
  onHost('parammodifier:done', (payload) => {
    state.acceptProgress = false;
    ProgressDialog.hide();
    setPageBusy(state, false);
    state.lastResult = {
      ok: payload?.ok,
      message: payload?.message || '',
      outputFolder: payload?.outputFolder || '',
      resultWorkbookPath: payload?.resultWorkbookPath || '',
      logTextPath: payload?.logTextPath || '',
      summary: payload?.summary || null,
      fileCount: Number(payload?.fileCount) || 0
    };
    if (payload?.activeDocument) state.activeDocument = payload.activeDocument;
    renderRunSummary(state);
    renderRvtModal(state);
    updateActionState(state);
    requestAnimationFrame(() => {
      ProgressDialog.hide();
      openCompletionResultDialog(state, payload || {});
    });
    toast(payload?.message || '파라미터 일괄 입력이 완료되었습니다.', payload?.ok === false ? 'err' : 'ok', 3200);
  });
  onHost('parammodifier:artifacts-exported', (payload) => {
    state.acceptProgress = false;
    ProgressDialog.hide();
    if (!state.lastResult) state.lastResult = {};
    if (payload?.outputFolder) state.lastResult.outputFolder = payload.outputFolder;
    if (payload?.resultWorkbookPath) state.lastResult.resultWorkbookPath = payload.resultWorkbookPath;
    if (payload?.logTextPath) state.lastResult.logTextPath = payload.logTextPath;
    renderRunSummary(state);
    renderRvtModal(state);
    updateActionState(state);
    requestAnimationFrame(() => {
      ProgressDialog.hide();
      openCompletionResultDialog(state, payload || {});
    });
    toast(payload?.message || '결과 파일을 저장했습니다.', payload?.ok === false ? 'err' : 'ok', 3200);
  });
  onHost('parammodifier:error', (payload) => {
    state.acceptProgress = false;
    ProgressDialog.hide();
    setPageBusy(state, false);
    toast(payload?.message || '오류가 발생했습니다.', 'err', 3600);
  });

  renderElementUpdateRows(state);
  renderRunSummary(state);
  renderRvtModal(state);
  updateActionState(state);
  post('parammodifier:init', {});
}

function openCompletionResultDialog(state, payload) {
  const summary = payload?.summary || state.lastResult?.summary || {};
  const fileCount = Number(payload?.fileCount ?? state.lastResult?.fileCount) || 0;
  const workbookPath = payload?.resultWorkbookPath || state.lastResult?.resultWorkbookPath || '';
  const logPath = payload?.logTextPath || state.lastResult?.logTextPath || '';
  const outputFolder = payload?.outputFolder || state.lastResult?.outputFolder || '';

  const successCount = summary?.SuccessCount ?? summary?.successCount ?? 0;
  const failCount = summary?.FailCount ?? summary?.failCount ?? 0;
  const noChangeCount = summary?.NoChangeCount ?? summary?.noChangeCount ?? 0;
  const updatedElementCount = summary?.TotalUpdatedElementCount ?? summary?.totalUpdatedElementCount ?? 0;
  const updatedParameterCount = summary?.TotalUpdatedParameterCount ?? summary?.totalUpdatedParameterCount ?? 0;
  const targetCount = fileCount || successCount + failCount + noChangeCount || (state.rvtPaths?.length || 1);

  const notes = [];
  if (outputFolder) notes.push(`결과 폴더: ${outputFolder}`);
  if (workbookPath) notes.push(`결과 엑셀: ${workbookPath}`);
  if (logPath) notes.push(`로그 파일: ${logPath}`);
  if (!workbookPath) notes.push('엑셀과 로그는 완료 후 이 창에서 필요할 때 추출합니다.');

  const actions = [];
  if (workbookPath) {
    actions.push({
      label: '결과 엑셀 열기',
      onClick: () => post('excel:open', { path: workbookPath })
    });
  } else {
    actions.push({
      label: '결과 엑셀 추출',
      variant: 'primary',
      onClick: () => post('parammodifier:export-results', {})
    });
  }
  if (logPath) {
    actions.push({
      label: '로그 열기',
      onClick: () => post('excel:open', { path: logPath })
    });
  }
  if (outputFolder) {
    actions.push({
      label: '결과 폴더 열기',
      onClick: () => post('parammodifier:open-folder', { path: outputFolder })
    });
  }

  showCompletionSummaryDialog({
    title: '파라미터 수정 완료',
    message: '작업이 완료되었습니다. 결과를 확인하고 필요하면 엑셀을 추출하세요.',
    summaryItems: [
      { label: '대상 파일', value: `${targetCount}개` },
      { label: '성공', value: `${successCount}개` },
      { label: '실패', value: `${failCount}개` },
      { label: '변경 없음', value: `${noChangeCount}개` },
      { label: '업데이트 요소', value: `${updatedElementCount}개` },
      { label: '업데이트 파라미터', value: `${updatedParameterCount}개` }
    ],
    notes,
    actions,
    confirmLabel: '닫기',
    showExport: false
  });
}
function buildRunCard(state) {
  const card = div('deliverycleaner-card deliverycleaner-card--files');
  card.innerHTML = `
    <div class="deliverycleaner-card__head">
      <div>
        <h3>실행</h3>
        <p>활성 문서에는 바로 적용하고, 여러 RVT는 목록 창에 등록한 뒤 일괄 적용합니다.</p>
      </div>
    </div>
  `;

  const actionRow = div('deliverycleaner-inline-actions multi-action-card__actions');
  const applyActiveBtn = actionButton('활성문서 적용', () => runActiveDocument(state), 'primary');
  const openBatchBtn = actionButton('여러 RVT 적용', () => openRvtModal(state), 'secondary');
  openBatchBtn.classList.add('btn--multi');
  actionRow.append(applyActiveBtn, openBatchBtn);

  const summary = div('deliverycleaner-note');

  const syncSection = createSubsection('동기화', '활성 문서 적용 후 동기화 여부를 선택합니다.');
  const syncCheckLine = document.createElement('label');
  syncCheckLine.className = 'deliverycleaner-checkline';
  const syncCheck = document.createElement('input');
  syncCheck.type = 'checkbox';
  syncCheck.addEventListener('change', () => {
    state.synchronizeAfterProcessing = !!syncCheck.checked;
    renderRunSummary(state);
  });
  const syncText = document.createElement('span');
  syncText.textContent = '활성 문서 작업 후 동기화';
  syncCheckLine.append(syncCheck, syncText);

  const syncHint = div('deliverycleaner-note');
  const activeSyncCommentField = fieldBlock('동기화 Comment');
  const activeSyncCommentInput = document.createElement('input');
  activeSyncCommentInput.type = 'text';
  activeSyncCommentInput.placeholder = '활성 문서 동기화 Comment';
  activeSyncCommentInput.addEventListener('input', () => {
    updateSyncComment(state, activeSyncCommentInput.value);
  });
  activeSyncCommentField.append(activeSyncCommentInput);
  syncSection.append(syncCheckLine, activeSyncCommentField, syncHint);

  const guide = div('deliverycleaner-note');
  card.append(actionRow, summary, syncSection, guide);

  state.ui.applyActiveBtn = applyActiveBtn;
  state.ui.openBatchBtn = openBatchBtn;
  state.ui.runSummary = summary;
  state.ui.syncCheck = syncCheck;
  state.ui.activeSyncCommentInput = activeSyncCommentInput;
  state.ui.syncHint = syncHint;
  state.ui.runGuide = guide;
  return card;
}

function buildSettingsCard(state) {
  const card = div('deliverycleaner-card deliverycleaner-card--workspace');
  card.append(buildElementPanel(state));
  return card;
}

function buildElementPanel(state) {
  const panel = createSubsection('입력 설정', '조건과 입력 파라미터를 정의해 어떤 객체에 어떤 값을 넣을지 결정합니다.');

  const comboRow = div('deliverycleaner-inline-select');
  const comboLabel = document.createElement('span');
  comboLabel.textContent = '조건 결합';
  const comboSelect = document.createElement('select');
  ['And', 'Or'].forEach((value) => {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = value;
    comboSelect.append(opt);
  });
  comboSelect.addEventListener('change', () => {
    state.elementParameterUpdate.combinationMode = comboSelect.value;
    renderElementUpdateSummary(state);
  });
  comboRow.append(comboLabel, comboSelect);

  const conditionSection = createCompactSection('조건', '대상 객체를 추려내는 조건입니다.');
  const conditionWrap = div('deliverycleaner-table-scroll');
  const conditionTable = document.createElement('table');
  conditionTable.className = 'deliverycleaner-grid-table';
  conditionTable.innerHTML = `
    <thead>
      <tr>
        <th>조건 파라미터</th>
        <th>Operator</th>
        <th>Value</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;
  conditionWrap.append(conditionTable);
  conditionSection.append(conditionWrap);

  const assignmentSection = createCompactSection('입력 파라미터', '선택된 객체에 입력할 파라미터와 값을 지정합니다.');
  const assignmentWrap = div('deliverycleaner-table-scroll');
  const assignmentTable = document.createElement('table');
  assignmentTable.className = 'deliverycleaner-grid-table';
  assignmentTable.innerHTML = `
    <thead>
      <tr>
        <th>입력 파라미터</th>
        <th>Value</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;
  assignmentWrap.append(assignmentTable);
  assignmentSection.append(assignmentWrap);

  const splitGrid = div('deliverycleaner-split-grid');
  splitGrid.append(conditionSection, assignmentSection);

  const summary = div('deliverycleaner-summary-box');
  panel.append(comboRow, splitGrid, summary);

  state.ui.combinationMode = comboSelect;
  state.ui.conditionBody = conditionTable.querySelector('tbody');
  state.ui.assignmentBody = assignmentTable.querySelector('tbody');
  state.ui.elementSummary = summary;
  return panel;
}

function buildRvtModal(state) {
  const overlay = div('rvt-expand-overlay');
  overlay.classList.add('is-hidden');

  const modal = div('rvt-expand-modal');
  const toolbar = div('rvt-expand-toolbar');
  const titleWrap = div('rvt-expand-title');
  const title = document.createElement('h3');
  title.textContent = '여러 RVT 적용';
  const badge = document.createElement('span');
  badge.className = 'chip chip--info';
  titleWrap.append(title, badge);

  const toolbarActions = div('rvt-expand-actions');
  const btnAdd = actionButton('RVT 추가', () => post('parammodifier:pick-rvts', {}), 'primary');
  const btnRemove = actionButton('선택 제거', () => {
    state.rvtPaths = state.rvtPaths.filter((path) => !state.checked.has(path));
    state.checked.clear();
    renderRvtModal(state);
    renderRunSummary(state);
    updateActionState(state);
  });
  const btnClear = actionButton('목록 지우기', () => {
    state.rvtPaths = [];
    state.checked.clear();
    renderRvtModal(state);
    renderRunSummary(state);
    updateActionState(state);
  });
  const btnClose = actionButton('닫기', () => closeRvtModal(state), 'secondary');
  toolbarActions.append(btnAdd, btnRemove, btnClear, btnClose);
  toolbar.append(titleWrap, toolbarActions);

  const body = div('rvt-expand-body');
  const panel = div('rvt-expand-panel');

  const listSection = div('rvt-expand-section rvt-expand-section--list');
  const listHead = div('rvt-expand-section__head');
  const listTitle = document.createElement('h4');
  listTitle.textContent = '선택된 RVT 목록';
  const listSub = document.createElement('span');
  listSub.textContent = 'Central File은 로컬로 열고, 워크셋은 모두 닫힌 상태로 처리합니다. 탐색기에서 .rvt 파일을 바로 끌어와 추가할 수도 있습니다.';
  listHead.append(listTitle, listSub);

  const tableWrap = div('rvt-expand-table rvt-drop-zone');
  const { table, tbody, master } = createRvtTable();
  table.classList.add('rvt-expand-table__grid');
  const listEmpty = div('rvt-expand-list-empty');
  const emptyTitle = document.createElement('strong');
  emptyTitle.textContent = '등록된 RVT가 없습니다.';
  const emptyText = document.createElement('span');
  emptyText.textContent = 'RVT 추가 버튼을 누르거나 탐색기에서 .rvt 파일을 이 영역으로 끌어오면 바로 목록에 추가됩니다.';
  const emptyBtn = actionButton('RVT 추가', () => post('parammodifier:pick-rvts', {}), 'primary');
  listEmpty.append(emptyTitle, emptyText, emptyBtn);
  tableWrap.append(table, listEmpty);
  listSection.append(listHead, tableWrap);

  const sideSection = div('rvt-expand-section rvt-expand-section--side');

  const syncSection = div('rvt-expand-subsection');
  const syncHead = div('rvt-expand-subsection__head');
  const syncTitle = document.createElement('h4');
  syncTitle.textContent = '동기화';
  const syncDescription = document.createElement('span');
  syncDescription.textContent = '여러 RVT 작업 후 자동 동기화할 때 아래 Comment를 사용합니다.';
  syncHead.append(syncTitle, syncDescription);
  const syncCommentField = fieldBlock('동기화 Comment');
  const syncCommentInput = document.createElement('input');
  syncCommentInput.type = 'text';
  syncCommentInput.placeholder = 'Parameter Modifier batch update';
  syncCommentInput.addEventListener('input', () => {
    updateSyncComment(state, syncCommentInput.value);
  });
  syncCommentField.append(syncCommentInput);
  syncSection.append(syncHead, syncCommentField);

  const readySection = div('rvt-expand-subsection rvt-expand-subsection--recent');
  const readyHead = div('rvt-expand-subsection__head');
  const readyTitle = document.createElement('h4');
  readyTitle.textContent = '적용 준비';
  const readyCaption = document.createElement('div');
  readyCaption.className = 'multi-recent-caption';
  const readyHint = document.createElement('div');
  readyHint.className = 'multi-recent-hint';
  readyHead.append(readyTitle, readyCaption, readyHint);

  const readyCard = div('review-feature-card');
  const readyCardHead = div('review-feature-card__head');
  const readyCardLabel = document.createElement('strong');
  readyCardLabel.textContent = '파라미터 수정 적용';
  const readyCardBadge = document.createElement('span');
  readyCardBadge.className = 'review-feature-card__badge';
  readyCardHead.append(readyCardLabel, readyCardBadge);
  const readyList = document.createElement('ul');
  readyList.className = 'review-feature-card__list';
  const readyRunBtn = actionButton('선택된 RVT 적용 시작', () => runBatchDocuments(state), 'primary');
  readyRunBtn.classList.add('review-feature-card__action');
  readyCard.append(readyCardHead, readyList, readyRunBtn);
  readySection.append(readyHead, readyCard);

  sideSection.append(syncSection, readySection);
  panel.append(listSection, sideSection);
  body.append(panel);
  modal.append(toolbar, body);
  overlay.append(modal);

  attachRvtDropZone(tableWrap, {
    onDropPaths: (paths) => {
      const added = appendDroppedRvts(state, paths);
      if (!added) {
        toast('이미 등록된 RVT입니다.', 'warn');
        return;
      }
      renderRvtModal(state);
      renderRunSummary(state);
      updateActionState(state);
      toast(`${added}개 RVT를 추가했습니다.`, 'ok');
    },
    onInvalid: () => toast('RVT 파일만 등록할 수 있습니다.', 'warn')
  });

  overlay.addEventListener('click', (ev) => {
    if (ev.target === overlay) closeRvtModal(state);
  });
  modal.addEventListener('click', (ev) => ev.stopPropagation());

  state.ui.rvtOverlay = overlay;
  state.ui.rvtBody = tbody;
  state.ui.rvtMaster = master;
  state.ui.rvtTable = table;
  state.ui.rvtEmpty = listEmpty;
  state.ui.rvtModalCount = badge;
  state.ui.rvtModalRemoveBtn = btnRemove;
  state.ui.rvtModalClearBtn = btnClear;
  state.ui.batchSyncCommentInput = syncCommentInput;
  state.ui.batchReadyCaption = readyCaption;
  state.ui.batchReadyHint = readyHint;
  state.ui.batchReadyList = readyList;
  state.ui.batchReadyBadge = readyCardBadge;
  state.ui.batchReadyRunBtn = readyRunBtn;
  return overlay;
}

function runActiveDocument(state) {
  if (getActiveBlockingReason(state)) {
    updateActionState(state);
    return;
  }

  state.acceptProgress = true;
  setPageBusy(state, true);
  post('parammodifier:run', buildPayload(state, true));
}

function runBatchDocuments(state) {
  if (getBatchBlockingReason(state)) {
    updateActionState(state);
    return;
  }

  state.acceptProgress = true;
  setPageBusy(state, true);
  closeRvtModal(state);
  post('parammodifier:run', buildPayload(state, false));
}

function openRvtModal(state) {
  if (getBatchOpenBlockingReason(state)) {
    updateActionState(state);
    return;
  }
  if (state.ui.rvtOverlay) state.ui.rvtOverlay.classList.remove('is-hidden');
  renderRvtModal(state);
  updateActionState(state);
}

function closeRvtModal(state) {
  if (state.ui.rvtOverlay) state.ui.rvtOverlay.classList.add('is-hidden');
}

function updateSyncComment(state, value) {
  state.syncComment = String(value || '').trim();
  syncInputs(state);
  renderRunSummary(state);
  renderBatchReadyCard(state);
}

function renderRunSummary(state) {
  if (!state.ui.runSummary || !state.ui.syncHint) return;

  const title = state.activeDocument?.title || '활성 문서를 찾을 수 없습니다.';
  const docType = state.activeDocument?.isWorkshared ? '워크셋 문서' : '일반 문서';
  const rvtCount = state.rvtPaths.length;
  const syncCommentState = state.syncComment ? `입력됨 (${state.syncComment})` : '미입력';
  state.ui.runSummary.textContent = `현재 문서: ${title}\n문서 유형: ${docType}\n등록된 RVT: ${rvtCount}개`;
  state.ui.syncHint.textContent = state.synchronizeAfterProcessing
    ? `활성 문서 적용 후 동기화를 수행합니다. 현재 Comment: ${syncCommentState}`
    : '활성 문서는 동기화 없이 적용합니다. 여러 RVT는 작업 후 자동 동기화됩니다.';
}

function renderRvtModal(state) {
  if (!state.ui.rvtBody || !state.ui.rvtMaster) return;

  const rows = state.rvtPaths.map((path, index) => ({
    checked: state.checked.has(path),
    index: index + 1,
    name: getRvtName(path),
    path,
    title: path,
    onToggle: (checked) => {
      if (checked) state.checked.add(path);
      else state.checked.delete(path);
      renderRvtModal(state);
      updateActionState(state);
    }
  }));

  renderRvtRows(state.ui.rvtBody, rows, '등록된 RVT가 없습니다.');

  const count = state.rvtPaths.length;
  state.ui.rvtModalCount.textContent = `${count}개`;
  state.ui.rvtEmpty.style.display = count ? 'none' : 'flex';
  state.ui.rvtTable.style.display = count ? 'table' : 'none';
  state.ui.rvtMaster.checked = count > 0 && state.rvtPaths.every((path) => state.checked.has(path));
  state.ui.rvtMaster.disabled = count === 0;
  state.ui.rvtMaster.onchange = () => {
    if (state.ui.rvtMaster.checked) state.checked = new Set(state.rvtPaths);
    else state.checked.clear();
    renderRvtModal(state);
    updateActionState(state);
  };

  if (state.ui.rvtModalRemoveBtn) state.ui.rvtModalRemoveBtn.disabled = state.checked.size === 0 || state.busy;
  if (state.ui.rvtModalClearBtn) state.ui.rvtModalClearBtn.disabled = count === 0 || state.busy;

  renderBatchReadyCard(state);
}

function renderBatchReadyCard(state) {
  if (!state.ui.batchReadyCaption || !state.ui.batchReadyHint || !state.ui.batchReadyList || !state.ui.batchReadyBadge) return;

  const assignments = getAssignmentRows(state);
  const conditions = getConditionRows(state);
  const ready = assignments.length > 0;

  state.ui.batchReadyCaption.textContent = ready ? '입력 설정과 적용 준비가 완료되었습니다.' : '입력 설정이 필요합니다.';
  state.ui.batchReadyHint.textContent = state.rvtPaths.length
    ? `${state.rvtPaths.length}개 RVT가 등록되어 있습니다.`
    : 'RVT를 추가하면 바로 배치 적용을 시작할 수 있습니다.';
  state.ui.batchReadyBadge.textContent = ready ? '적용 준비됨' : '입력 설정 필요';

  state.ui.batchReadyList.innerHTML = '';
  [
    `조건 ${conditions.length}개`,
    `입력 파라미터 ${assignments.length}개`,
    `동기화 Comment ${state.syncComment ? '입력됨' : '미입력'}`,
    'Central File은 로컬로 열고 워크셋은 모두 닫은 채 처리'
  ].forEach((line) => {
    const item = document.createElement('li');
    item.textContent = line;
    state.ui.batchReadyList.append(item);
  });
}

function renderElementUpdateRows(state) {
  const conditionBody = state.ui.conditionBody;
  const assignmentBody = state.ui.assignmentBody;
  if (!conditionBody || !assignmentBody) return;

  conditionBody.innerHTML = '';
  state.elementParameterUpdate.conditions.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.append(
      tdWithInput(row.parameterName, '파라미터 이름', (value) => {
        state.elementParameterUpdate.conditions[index].parameterName = value;
        renderElementUpdateSummary(state);
      }),
      tdWithSelect(row.operatorName, FILTER_OPERATORS, (value) => {
        state.elementParameterUpdate.conditions[index].operatorName = value;
        renderElementUpdateSummary(state);
      }),
      tdWithInput(row.value, '값', (value) => {
        state.elementParameterUpdate.conditions[index].value = value;
        renderElementUpdateSummary(state);
      })
    );
    conditionBody.append(tr);
  });

  assignmentBody.innerHTML = '';
  state.elementParameterUpdate.assignments.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.append(
      tdWithInput(row.parameterName, '파라미터 이름', (value) => {
        state.elementParameterUpdate.assignments[index].parameterName = value;
        renderElementUpdateSummary(state);
      }),
      tdWithInput(row.value, '값', (value) => {
        state.elementParameterUpdate.assignments[index].value = value;
        renderElementUpdateSummary(state);
      })
    );
    assignmentBody.append(tr);
  });

  renderElementUpdateSummary(state);
}

function renderElementUpdateSummary(state) {
  if (!state.ui.elementSummary) return;

  const conditions = getConditionRows(state);
  const assignments = getAssignmentRows(state);
  const joiner = state.elementParameterUpdate.combinationMode === 'Or' ? ' OR ' : ' AND ';

  const conditionText = conditions.length
    ? conditions.map((row) => (row.operatorName === 'HasValue' || row.operatorName === 'HasNoValue')
      ? `${row.parameterName} ${row.operatorName}`
      : `${row.parameterName} ${row.operatorName} ${row.value || ''}`).join(joiner)
    : '추가 조건 없음';

  const assignmentText = assignments.length
    ? assignments.map((row) => `${row.parameterName} = ${row.value || ''}`).join(' / ')
    : '입력 파라미터 없음';

  state.ui.elementSummary.textContent = `조건: ${conditionText}\n입력: ${assignmentText}`;
  renderRvtModal(state);
  updateActionState(state);
}

function applyHostState(state, payload) {
  const settings = payload?.settings || {};
  const result = payload?.result || null;
  const activeDocument = payload?.activeDocument || null;

  if (Array.isArray(settings.rvtPaths)) {
    state.rvtPaths = [...settings.rvtPaths];
    state.checked = new Set(state.rvtPaths);
  }
  if (typeof settings.synchronizeAfterProcessing === 'boolean') state.synchronizeAfterProcessing = settings.synchronizeAfterProcessing;
  if (typeof settings.syncComment === 'string') state.syncComment = settings.syncComment;
  if (settings.elementParameterUpdate) state.elementParameterUpdate = normalizeElementUpdate(settings.elementParameterUpdate);

  state.lastResult = result;
  state.activeDocument = activeDocument;

  syncInputs(state);
  renderElementUpdateRows(state);
  renderRunSummary(state);
  renderRvtModal(state);
  updateActionState(state);
}

function syncInputs(state) {
  if (state.ui.syncCheck) state.ui.syncCheck.checked = !!state.synchronizeAfterProcessing;
  if (state.ui.activeSyncCommentInput) state.ui.activeSyncCommentInput.value = state.syncComment || '';
  if (state.ui.batchSyncCommentInput) state.ui.batchSyncCommentInput.value = state.syncComment || '';
  if (state.ui.combinationMode) state.ui.combinationMode.value = state.elementParameterUpdate.combinationMode || 'And';
}

function updateActionState(state) {
  const activeReason = getActiveBlockingReason(state);
  const batchOpenReason = getBatchOpenBlockingReason(state);
  const batchRunReason = getBatchBlockingReason(state);

  if (state.ui.applyActiveBtn) state.ui.applyActiveBtn.disabled = !!activeReason || state.busy;
  if (state.ui.openBatchBtn) state.ui.openBatchBtn.disabled = !!batchOpenReason || state.busy;
  if (state.ui.batchReadyRunBtn) state.ui.batchReadyRunBtn.disabled = !!batchRunReason || state.busy;

  if (state.ui.runGuide) {
    if (state.busy) {
      state.ui.runGuide.textContent = '작업 진행 중입니다. 완료되면 결과 알림을 표시합니다.';
    } else if (batchOpenReason) {
      state.ui.runGuide.textContent = '입력 파라미터를 1개 이상 설정하면 활성문서 적용과 여러 RVT 적용 버튼이 활성화됩니다.';
    } else if (!state.activeDocument) {
      state.ui.runGuide.textContent = '현재 활성 문서를 찾을 수 없습니다.';
    } else {
      state.ui.runGuide.textContent = '입력 설정이 준비되었습니다. 활성문서 적용은 바로 실행하고, 여러 RVT 적용은 목록 창에서 시작합니다.';
    }
  }
}

function getActiveBlockingReason(state) {
  if (!getAssignmentRows(state).length) return '입력 파라미터를 1개 이상 설정해 주세요.';
  if (!state.activeDocument) return '활성 문서를 찾을 수 없습니다.';
  return '';
}

function getBatchOpenBlockingReason(state) {
  if (!getAssignmentRows(state).length) return '입력 파라미터를 1개 이상 설정해 주세요.';
  return '';
}

function getBatchBlockingReason(state) {
  if (!getAssignmentRows(state).length) return '입력 파라미터를 1개 이상 설정해 주세요.';
  if (!state.rvtPaths.length) return '적용할 RVT를 1개 이상 등록해 주세요.';
  return '';
}

function buildPayload(state, useActiveDocument) {
  return {
    useActiveDocument: !!useActiveDocument,
    rvtPaths: useActiveDocument ? [] : [...state.rvtPaths],
    outputFolder: '',
    closeAllWorksetsOnOpen: true,
    synchronizeAfterProcessing: useActiveDocument ? !!state.synchronizeAfterProcessing : true,
    syncComment: state.syncComment,
    filterProfile: null,
    elementParameterUpdate: {
      enabled: true,
      combinationMode: state.elementParameterUpdate.combinationMode,
      conditions: state.elementParameterUpdate.conditions.map((row) => ({
        ...row,
        enabled: !!String(row.parameterName || '').trim()
      })),
      assignments: state.elementParameterUpdate.assignments.map((row) => ({
        ...row,
        enabled: !!String(row.parameterName || '').trim()
      }))
    }
  };
}

function createElementUpdateState() {
  return {
    enabled: true,
    combinationMode: 'And',
    conditions: Array.from({ length: 4 }, () => ({ enabled: false, parameterName: '', operatorName: 'Equals', value: '' })),
    assignments: Array.from({ length: 4 }, () => ({ enabled: false, parameterName: '', value: '' }))
  };
}

function normalizeElementUpdate(source) {
  const base = createElementUpdateState();
  base.combinationMode = source.combinationMode || 'And';
  base.conditions = base.conditions.map((row, index) => ({ ...row, ...(Array.isArray(source.conditions) ? source.conditions[index] || {} : {}) }));
  base.assignments = base.assignments.map((row, index) => ({ ...row, ...(Array.isArray(source.assignments) ? source.assignments[index] || {} : {}) }));
  return base;
}

function getConditionRows(state) {
  return state.elementParameterUpdate.conditions.filter((row) => String(row.parameterName || '').trim());
}

function getAssignmentRows(state) {
  return state.elementParameterUpdate.assignments.filter((row) => String(row.parameterName || '').trim());
}

function appendDroppedRvts(state, paths) {
  let added = 0;
  (Array.isArray(paths) ? paths : []).forEach((path) => {
    if (!path) return;
    if (!state.rvtPaths.some((item) => samePath(item, path))) {
      state.rvtPaths.push(path);
      added += 1;
    }
    state.checked.add(path);
  });
  return added;
}

function samePath(left, right) {
  return String(left || '').toLowerCase() === String(right || '').toLowerCase();
}

function actionButton(label, onClick, variant = 'secondary') {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = `btn ${variant === 'primary' ? 'btn--primary' : 'btn--secondary'}`;
  btn.textContent = label;
  btn.addEventListener('click', onClick);
  return btn;
}

function fieldBlock(labelText) {
  const wrap = div('deliverycleaner-field');
  const label = document.createElement('label');
  label.textContent = labelText;
  wrap.append(label);
  return wrap;
}

function createSubsection(title, description) {
  const section = div('deliverycleaner-subsection');
  const head = div('deliverycleaner-subsection__head');
  const h4 = document.createElement('h4');
  h4.textContent = title;
  const p = document.createElement('p');
  p.textContent = description;
  head.append(h4, p);
  section.append(head);
  return section;
}

function createCompactSection(title, description) {
  const section = div('deliverycleaner-subsection deliverycleaner-subsection--compact');
  const head = div('deliverycleaner-subsection__head');
  const h4 = document.createElement('h4');
  h4.textContent = title;
  const p = document.createElement('p');
  p.textContent = description;
  head.append(h4, p);
  section.append(head);
  return section;
}

function tdWithInput(value, placeholder, onChange) {
  const td = document.createElement('td');
  const input = document.createElement('input');
  input.type = 'text';
  input.value = value || '';
  input.placeholder = placeholder || '';
  input.addEventListener('input', () => onChange(input.value.trim()));
  td.append(input);
  return td;
}

function tdWithSelect(value, options, onChange) {
  const td = document.createElement('td');
  const select = document.createElement('select');
  options.forEach((optionValue) => {
    const opt = document.createElement('option');
    opt.value = optionValue;
    opt.textContent = optionValue;
    select.append(opt);
  });
  select.value = value || options[0];
  select.addEventListener('change', () => onChange(select.value));
  td.append(select);
  return td;
}

function setPageBusy(state, on) {
  state.busy = !!on;
  if (!state.busy) state.acceptProgress = false;
  setBusy(on, on ? '파라미터 수정 작업을 처리하는 중입니다.' : '');
  updateActionState(state);
}

