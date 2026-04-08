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

export function renderConditionExtract(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  const state = {
    rvtPaths: [],
    checked: new Set(),
    includeCoordinates: false,
    includeLinearMetrics: false,
    lengthUnit: 'mm',
    areaUnit: 'mm2',
    volumeUnit: 'mm3',
    extractParameterNamesCsv: '',
    elementParameterUpdate: createConditionState(),
    lastResult: null,
    activeDocument: null,
    activeSettingsTab: 'filters',
    busy: false,
    acceptProgress: false,
    ui: {}
  };

  const page = div('feature-shell deliverycleaner-page');
  page.innerHTML = `
    <div class="feature-header deliverycleaner-header">
      <div class="feature-heading">
        <span class="feature-kicker">BQC · 객체 조건 검토 워크플로</span>
        <h2 class="feature-title">조건별 객체 속성 추출</h2>
        <p class="feature-sub">조건식으로 객체를 추려 지정한 속성과 좌표, 선형 정보를 한 번에 추출합니다.</p>
      </div>
    </div>
  `;

  const layout = div('deliverycleaner-top-grid');
  const leftColumn = div('deliverycleaner-field-stack');
  leftColumn.append(buildRunCard(state));
  layout.append(leftColumn, buildSettingsCard(state));
  page.append(layout, buildRvtModal(state));
  target.append(page);

  onHost('conditionextract:init', (payload) => applyHostState(state, payload));
  onHost('conditionextract:rvts-picked', (payload) => {
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
  onHost('conditionextract:progress', (payload) => {
    if (!state.acceptProgress || !state.busy) return;
    ProgressDialog.setActions({});
    ProgressDialog.show(payload?.title || '조건별 객체 속성 추출', payload?.message || '');
    ProgressDialog.update(Number(payload?.percent) || 0, payload?.message || '', payload?.detail || '');
  });
  onHost('conditionextract:done', (payload) => {
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
    requestAnimationFrame(() => openCompletionResultDialog(state, payload || {}));
    toast(payload?.message || '조건별 객체 속성 추출이 완료되었습니다.', payload?.ok === false ? 'err' : 'ok', 3200);
  });
  onHost('conditionextract:artifacts-exported', (payload) => {
    state.acceptProgress = false;
    ProgressDialog.hide();
    if (!state.lastResult) state.lastResult = {};
    if (payload?.outputFolder) state.lastResult.outputFolder = payload.outputFolder;
    if (payload?.resultWorkbookPath) state.lastResult.resultWorkbookPath = payload.resultWorkbookPath;
    if (payload?.logTextPath) state.lastResult.logTextPath = payload.logTextPath;
    renderRunSummary(state);
    renderRvtModal(state);
    updateActionState(state);
    requestAnimationFrame(() => openCompletionResultDialog(state, payload || {}));
    toast(payload?.message || '결과 파일을 저장했습니다.', payload?.ok === false ? 'err' : 'ok', 3200);
  });
  onHost('conditionextract:error', (payload) => {
    state.acceptProgress = false;
    ProgressDialog.hide();
    setPageBusy(state, false);
    state.lastResult = null;
    renderRunSummary(state);
    renderRvtModal(state);
    updateActionState(state);
    toast(payload?.message || '오류가 발생했습니다.', 'err', 3600);
  });

  renderConditionRows(state);
  renderSettingSummary(state);
  renderRunSummary(state);
  renderRvtModal(state);
  updateActionState(state);
  post('conditionextract:init', {});
}

function buildRunCard(state) {
  const card = div('deliverycleaner-card deliverycleaner-card--files');
  card.innerHTML = `
    <div class="deliverycleaner-card__head">
      <div>
        <h3>실행</h3>
        <p>활성 문서를 바로 검토하거나 여러 RVT를 목록에 등록해 한 번에 추출합니다.</p>
      </div>
    </div>
  `;

  const actionRow = div('deliverycleaner-inline-actions multi-action-card__actions');
  const activeBtn = actionButton('활성문서 검토', () => runActiveDocument(state), 'primary');
  const batchBtn = actionButton('여러 RVT 검토', () => openRvtModal(state), 'secondary');
  batchBtn.classList.add('btn--multi');
  actionRow.append(activeBtn, batchBtn);

  const summary = div('deliverycleaner-note');
  const guide = div('deliverycleaner-note');
  card.append(actionRow, summary, guide);

  state.ui.applyActiveBtn = activeBtn;
  state.ui.openBatchBtn = batchBtn;
  state.ui.runSummary = summary;
  state.ui.runGuide = guide;
  return card;
}

function buildSettingsCard(state) {
  const card = div('deliverycleaner-card deliverycleaner-card--workspace');
  card.append(buildConditionPanel(state));
  return card;
}

function buildConditionPanel(state) {
  const panel = createSubsection('추출 설정', '필터, 추출 파라미터, 좌표/선형 옵션을 한 곳에서 설정합니다.');

  const tabs = div('deliverycleaner-tabs');
  const panels = div('deliverycleaner-panels conditionextract-panels');

  const conditionTab = createSettingsTab(state, 'filters', '필터');
  const extractTab = createSettingsTab(state, 'extract', '추출');
  const unitsTab = createSettingsTab(state, 'units', '단위');
  const extrasTab = createSettingsTab(state, 'extras', '옵션');
  tabs.append(conditionTab, extractTab, unitsTab, extrasTab);

  const conditionPanel = div('deliverycleaner-panel');
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
    renderSettingSummary(state);
  });
  comboRow.append(comboLabel, comboSelect);

  const conditionSection = createCompactSection('필터', '필터 파라미터는 인스턴스와 타입 파라미터를 모두 대상으로 찾고, 필요한 만큼 추가할 수 있습니다.');
  const conditionActions = div('conditionextract-section-actions');
  const addConditionBtn = actionButton('필터 추가', () => {
    state.elementParameterUpdate.conditions.push(createEmptyConditionRow());
    renderConditionRows(state);
    renderSettingSummary(state);
  }, 'secondary');
  conditionActions.append(addConditionBtn);
  const conditionWrap = div('deliverycleaner-table-scroll deliverycleaner-table-scroll--filter');
  const conditionTable = document.createElement('table');
  conditionTable.className = 'deliverycleaner-grid-table deliverycleaner-grid-table--conditions';
  conditionTable.innerHTML = `
    <thead>
      <tr>
        <th>필터 파라미터</th>
        <th>Operator</th>
        <th>Value</th>
        <th>관리</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;
  conditionWrap.append(conditionTable);
  conditionSection.append(conditionActions, conditionWrap);
  conditionPanel.append(comboRow, conditionSection);

  const extractPanel = div('deliverycleaner-panel');
  const extractSection = createCompactSection('추출 파라미터', '쉼표로 구분해 입력한 모든 파라미터를 추출하고, 필요하면 좌표와 선형 정보도 함께 포함합니다.');
  const extractField = div('deliverycleaner-field');
  extractField.innerHTML = `
    <label>파라미터 목록</label>
    <textarea rows="4" placeholder="예: Comments, Mark, Type Comments, Area, Volume"></textarea>
  `;
  const extractInput = extractField.querySelector('textarea');
  extractInput.addEventListener('input', () => {
    state.extractParameterNamesCsv = String(extractInput.value || '').trim();
    renderSettingSummary(state);
  });
  const coordLine = document.createElement('label');
  coordLine.className = 'deliverycleaner-checkline';
  coordLine.innerHTML = '<input type="checkbox"> 좌표 추출 (X / Y / Z)';
  const coordInput = coordLine.querySelector('input');
  coordInput.addEventListener('change', () => {
    state.includeCoordinates = !!coordInput.checked;
    renderSettingSummary(state);
  });

  const linearLine = document.createElement('label');
  linearLine.className = 'deliverycleaner-checkline';
  linearLine.innerHTML = '<input type="checkbox"> 선형 객체 방향 / 길이 추출';
  const linearInput = linearLine.querySelector('input');
  linearInput.addEventListener('change', () => {
    state.includeLinearMetrics = !!linearInput.checked;
    renderSettingSummary(state);
  });

  const extractOptions = div('deliverycleaner-section-callout deliverycleaner-section-callout--compact');
  const extractOptionsText = div('deliverycleaner-vg-rulebar__text');
  extractOptionsText.innerHTML = `
    <strong>추출 옵션</strong>
    <span>좌표는 X / Y / Z 열로 저장되고, 선형 객체는 DirectionX / DirectionY / DirectionZ 와 Length 열을 함께 기록합니다.</span>
  `;
  const extractOptionsControls = div('deliverycleaner-inline-controls');
  const lengthSelectWrap = createUnitField('좌표 / 길이 단위', [
    { value: 'mm', label: 'mm' },
    { value: 'inch', label: 'inch' },
    { value: 'ft', label: 'ft' }
  ], (value) => {
    state.lengthUnit = value;
    renderSettingSummary(state);
  });
  lengthSelectWrap.wrap.classList.add('deliverycleaner-inline-select', 'deliverycleaner-inline-select--stacked');
  extractOptionsControls.append(coordLine, linearLine, lengthSelectWrap.wrap);
  extractOptions.append(extractOptionsText, extractOptionsControls);

  const extractNote = div('deliverycleaner-note');
  extractNote.textContent = '값은 인스턴스 우선으로 찾고, 없으면 타입 파라미터에서 찾습니다. 좌표와 길이, 길이 파라미터는 위 단위를 따르고, 면적/체적은 단위 탭 설정으로 변환됩니다.';
  extractSection.append(extractField, extractOptions, extractNote);
  extractPanel.append(extractSection);

  const unitsPanel = div('deliverycleaner-panel');
  const unitsSection = createCompactSection('단위 설정', '면적과 체적 파라미터를 선택한 단위로 출력합니다. 좌표와 길이는 추출 탭의 단위를 따릅니다.');
  const unitGrid = div('conditionextract-unit-grid');
  const areaSelect = createUnitField('면적', [
    { value: 'mm2', label: 'mm^2' },
    { value: 'in2', label: 'inch^2' },
    { value: 'ft2', label: 'ft^2' }
  ], (value) => {
    state.areaUnit = value;
    renderSettingSummary(state);
  });
  const volumeSelect = createUnitField('체적', [
    { value: 'mm3', label: 'mm^3' },
    { value: 'in3', label: 'inch^3' },
    { value: 'ft3', label: 'ft^3' }
  ], (value) => {
    state.volumeUnit = value;
    renderSettingSummary(state);
  });
  unitGrid.append(areaSelect.wrap, volumeSelect.wrap);
  const unitNote = div('deliverycleaner-note');
  unitNote.textContent = '파라미터 타입이 면적, 체적으로 판별되면 선택 단위로 변환하고, 그 외 값은 원래 표시값을 유지합니다.';
  unitsSection.append(unitGrid, unitNote);
  unitsPanel.append(unitsSection);

  const optionPanel = div('deliverycleaner-panel');
  const optionSection = createCompactSection('추가 옵션 안내', '필터와 추출 대상의 동작 기준을 확인합니다.');
  const optionNote = div('deliverycleaner-note');
  optionNote.textContent = '필터는 인스턴스와 타입 파라미터를 모두 검사합니다. 좌표는 X / Y / Z, 선형 정보는 DirectionX / DirectionY / DirectionZ 와 Length 열에 기록됩니다.';
  optionSection.append(optionNote);
  optionPanel.append(optionSection);

  const summary = div('deliverycleaner-summary-box');
  panels.append(conditionPanel, extractPanel, unitsPanel, optionPanel);
  panel.append(tabs, panels, summary);

  state.ui.settingsTabs = {
    filters: conditionTab,
    extract: extractTab,
    units: unitsTab,
    extras: extrasTab
  };
  state.ui.settingsPanels = {
    filters: conditionPanel,
    extract: extractPanel,
    units: unitsPanel,
    extras: optionPanel
  };
  state.ui.combinationMode = comboSelect;
  state.ui.conditionBody = conditionTable.querySelector('tbody');
  state.ui.addConditionBtn = addConditionBtn;
  state.ui.extractInput = extractInput;
  state.ui.includeCoordinates = coordInput;
  state.ui.includeLinearMetrics = linearInput;
  state.ui.lengthUnit = lengthSelectWrap.select;
  state.ui.areaUnit = areaSelect.select;
  state.ui.volumeUnit = volumeSelect.select;
  state.ui.settingSummary = summary;
  syncSettingsTab(state);
  return panel;
}

function buildRvtModal(state) {
  const overlay = div('rvt-expand-overlay');
  overlay.classList.add('is-hidden');

  const modal = div('rvt-expand-modal');
  const toolbar = div('rvt-expand-toolbar');
  const titleWrap = div('rvt-expand-title');
  const title = document.createElement('h3');
  title.textContent = '여러 RVT 검토';
  const badge = document.createElement('span');
  badge.className = 'chip chip--info';
  titleWrap.append(title, badge);

  const toolbarActions = div('rvt-expand-actions');
  const btnAdd = actionButton('RVT 추가', () => post('conditionextract:pick-rvts', {}), 'primary');
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
  listSub.textContent = 'Central File은 Detach로 열고, 워크셋은 모두 닫은 상태로 검토합니다. 탐색기에서 .rvt 파일을 끌어와 바로 추가할 수도 있습니다.';
  listHead.append(listTitle, listSub);

  const tableWrap = div('rvt-expand-table rvt-drop-zone');
  const { table, tbody, master } = createRvtTable();
  table.classList.add('rvt-expand-table__grid');
  const listEmpty = div('rvt-expand-list-empty');
  const emptyTitle = document.createElement('strong');
  emptyTitle.textContent = '등록된 RVT가 없습니다.';
  const emptyText = document.createElement('span');
  emptyText.textContent = 'RVT 추가 버튼을 누르거나 탐색기에서 .rvt 파일을 끌어오면 바로 목록에 추가됩니다.';
  const emptyBtn = actionButton('RVT 추가', () => post('conditionextract:pick-rvts', {}), 'primary');
  listEmpty.append(emptyTitle, emptyText, emptyBtn);
  tableWrap.append(table, listEmpty);
  listSection.append(listHead, tableWrap);

  const sideSection = div('rvt-expand-section rvt-expand-section--side');
  const readySection = div('rvt-expand-subsection rvt-expand-subsection--recent');
  const readyHead = div('rvt-expand-subsection__head');
  const readyTitle = document.createElement('h4');
  readyTitle.textContent = '검토 준비';
  const readyCaption = document.createElement('div');
  readyCaption.className = 'multi-recent-caption';
  const readyHint = document.createElement('div');
  readyHint.className = 'multi-recent-hint';
  readyHead.append(readyTitle, readyCaption, readyHint);

  const readyCard = div('review-feature-card');
  const readyCardHead = div('review-feature-card__head');
  const readyCardLabel = document.createElement('strong');
  readyCardLabel.textContent = '조건별 객체 속성 추출';
  const readyCardBadge = document.createElement('span');
  readyCardBadge.className = 'review-feature-card__badge';
  readyCardHead.append(readyCardLabel, readyCardBadge);
  const readyList = document.createElement('ul');
  readyList.className = 'review-feature-card__list';
  const readyRunBtn = actionButton('선택한 RVT 검토 시작', () => runBatchDocuments(state), 'primary');
  readyRunBtn.classList.add('review-feature-card__action');
  readyCard.append(readyCardHead, readyList, readyRunBtn);
  readySection.append(readyHead, readyCard);

  sideSection.append(readySection);
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
  state.ui.batchReadyCaption = readyCaption;
  state.ui.batchReadyHint = readyHint;
  state.ui.batchReadyList = readyList;
  state.ui.batchReadyBadge = readyCardBadge;
  state.ui.batchReadyRunBtn = readyRunBtn;
  return overlay;
}

function openCompletionResultDialog(state, payload) {
  const summary = payload?.summary || state.lastResult?.summary || {};
  const fileCount = Number(payload?.fileCount ?? state.lastResult?.fileCount) || 0;
  const workbookPath = payload?.resultWorkbookPath || state.lastResult?.resultWorkbookPath || '';
  const logPath = payload?.logTextPath || state.lastResult?.logTextPath || '';
  const outputFolder = payload?.outputFolder || state.lastResult?.outputFolder || '';

  const successCount = summary?.SuccessCount ?? summary?.successCount ?? 0;
  const failCount = summary?.FailCount ?? summary?.failCount ?? 0;
  const noDataCount = summary?.NoDataCount ?? summary?.noDataCount ?? 0;
  const matchedCount = summary?.TotalMatchedCount ?? summary?.totalMatchedCount ?? 0;
  const extractedCount = summary?.TotalExtractedElementCount ?? summary?.totalExtractedElementCount ?? 0;
  const lengthUnitLabel = formatLengthUnit(state.lengthUnit);
  const areaUnitLabel = formatAreaUnit(state.areaUnit);
  const volumeUnitLabel = formatVolumeUnit(state.volumeUnit);

  const notes = [];
  if (outputFolder) notes.push(`결과 폴더: ${outputFolder}`);
  if (workbookPath) notes.push(`결과 파일: ${workbookPath}`);
  if (logPath) notes.push(`로그 파일: ${logPath}`);
  notes.push(`좌표와 길이는 ${lengthUnitLabel} 기준으로 저장됩니다.`);
  notes.push(`면적 파라미터는 ${areaUnitLabel}, 체적 파라미터는 ${volumeUnitLabel} 기준으로 저장됩니다.`);
  notes.push('선형 방향은 DirectionX / DirectionY / DirectionZ 열에 기록됩니다.');

  const actions = [];
  if (workbookPath) {
    actions.push({ label: '결과 파일 열기', onClick: () => post('excel:open', { path: workbookPath }) });
  } else {
    actions.push({
      label: '결과 파일 추출',
      variant: 'primary',
      onClick: () => {
        beginConditionExtractProgress('조건별 객체 속성 추출', '결과 파일을 준비하고 있습니다.');
        post('conditionextract:export-results', {});
      }
    });
  }
  if (logPath) actions.push({ label: '로그 열기', onClick: () => post('excel:open', { path: logPath }) });
  if (outputFolder) actions.push({ label: '결과 폴더 열기', onClick: () => post('conditionextract:open-folder', { path: outputFolder }) });

  showCompletionSummaryDialog({
    dialogClassName: 'completion-summary-dialog--wide',
    title: '조건별 객체 속성 추출 완료',
    message: '검토와 추출이 완료되었습니다. 결과를 이어서 확인하세요.',
    summaryItems: [
      { label: '대상 파일', value: `${fileCount || successCount + failCount + noDataCount || (state.rvtPaths.length || 1)}개` },
      { label: '성공', value: `${successCount}개` },
      { label: '실패', value: `${failCount}개` },
      { label: '조건 불일치', value: `${noDataCount}개` },
      { label: '조건 일치 객체', value: `${matchedCount}개` },
      { label: '추출 객체', value: `${extractedCount}개` }
    ],
    notes,
    actions,
    confirmLabel: '닫기',
    showExport: false
  });
}

function runActiveDocument(state) {
  if (getActiveBlockingReason(state)) {
    updateActionState(state);
    return;
  }

  state.acceptProgress = true;
  setPageBusy(state, true);
  beginConditionExtractProgress('조건별 객체 속성 추출', '활성 문서 검토를 준비 중입니다.');
  post('conditionextract:run', buildPayload(state, true));
}

function runBatchDocuments(state) {
  if (getBatchBlockingReason(state)) {
    updateActionState(state);
    return;
  }

  state.acceptProgress = true;
  setPageBusy(state, true);
  closeRvtModal(state);
  beginConditionExtractProgress('조건별 객체 속성 추출', '여러 RVT 검토를 준비 중입니다.');
  post('conditionextract:run', buildPayload(state, false));
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

function beginConditionExtractProgress(title, message, detail = '', percent = 0) {
  ProgressDialog.setActions({});
  ProgressDialog.show(title || '조건별 객체 속성 추출', message || '준비 중...');
  ProgressDialog.update(percent, message || '준비 중...', detail || '');
}

function applyHostState(state, payload) {
  const settings = payload?.settings || {};
  state.rvtPaths = Array.isArray(settings.rvtPaths) ? [...settings.rvtPaths] : [];
  state.checked = new Set(state.rvtPaths);
  state.elementParameterUpdate = normalizeConditionState(settings.elementParameterUpdate);
  state.includeCoordinates = !!settings.includeCoordinates;
  state.includeLinearMetrics = !!settings.includeLinearMetrics;
  state.lengthUnit = settings.lengthUnit || 'mm';
  state.areaUnit = settings.areaUnit || 'mm2';
  state.volumeUnit = settings.volumeUnit || 'mm3';
  state.extractParameterNamesCsv = String(settings.extractParameterNamesCsv || '').trim();
  state.lastResult = payload?.result || null;
  state.activeDocument = payload?.activeDocument || null;

  syncInputs(state);
  renderConditionRows(state);
  renderSettingSummary(state);
  renderRunSummary(state);
  renderRvtModal(state);
  updateActionState(state);
}

function syncInputs(state) {
  if (state.ui.combinationMode) state.ui.combinationMode.value = state.elementParameterUpdate.combinationMode || 'And';
  if (state.ui.extractInput) state.ui.extractInput.value = state.extractParameterNamesCsv || '';
  if (state.ui.includeCoordinates) state.ui.includeCoordinates.checked = !!state.includeCoordinates;
  if (state.ui.includeLinearMetrics) state.ui.includeLinearMetrics.checked = !!state.includeLinearMetrics;
  if (state.ui.lengthUnit) state.ui.lengthUnit.value = state.lengthUnit || 'mm';
  if (state.ui.areaUnit) state.ui.areaUnit.value = state.areaUnit || 'mm2';
  if (state.ui.volumeUnit) state.ui.volumeUnit.value = state.volumeUnit || 'mm3';
  syncSettingsTab(state);
}

function renderConditionRows(state) {
  const conditionBody = state.ui.conditionBody;
  if (!conditionBody) return;

  conditionBody.innerHTML = '';
  state.elementParameterUpdate.conditions.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.append(
      tdWithInput(row.parameterName, '파라미터 이름', (value) => {
        state.elementParameterUpdate.conditions[index].parameterName = value;
        renderSettingSummary(state);
      }),
      tdWithSelect(row.operatorName, FILTER_OPERATORS, (value) => {
        state.elementParameterUpdate.conditions[index].operatorName = value;
        renderSettingSummary(state);
      }),
      tdWithInput(row.value, '값', (value) => {
        state.elementParameterUpdate.conditions[index].value = value;
        renderSettingSummary(state);
      }),
      tdWithAction(() => {
        removeConditionRow(state, index);
      }, state.elementParameterUpdate.conditions.length <= 1)
    );
    conditionBody.append(tr);
  });

  if (state.ui.addConditionBtn) state.ui.addConditionBtn.disabled = !!state.busy;
}

function renderSettingSummary(state) {
  if (!state.ui.settingSummary) return;

  const conditions = getConditionRows(state);
  const extractNames = getExtractParameterRows(state);
  const joiner = state.elementParameterUpdate.combinationMode === 'Or' ? ' OR ' : ' AND ';
  const conditionText = conditions.length
    ? conditions.map((row) => (row.operatorName === 'HasValue' || row.operatorName === 'HasNoValue')
      ? `${row.parameterName} ${row.operatorName}`
      : `${row.parameterName} ${row.operatorName} ${row.value || ''}`).join(joiner)
    : '추가 필터 없음';
  const extractText = extractNames.length ? extractNames.join(', ') : '추출 파라미터 없음';
  const extras = [];
  if (state.includeCoordinates) extras.push('좌표 X/Y/Z');
  if (state.includeLinearMetrics) extras.push('선형 방향/길이');
  const unitsText = `길이 ${formatLengthUnit(state.lengthUnit)} / 면적 ${formatAreaUnit(state.areaUnit)} / 체적 ${formatVolumeUnit(state.volumeUnit)}`;

  state.ui.settingSummary.textContent = `필터: ${conditionText}\n추출: ${extractText}\n단위: ${unitsText}\n추가 옵션: ${extras.length ? extras.join(' / ') : '없음'}`;
  renderRunSummary(state);
  renderBatchReadyCard(state);
  updateActionState(state);
}

function renderRunSummary(state) {
  if (!state.ui.runSummary || !state.ui.runGuide) return;

  const title = state.activeDocument?.title || '활성 문서를 찾을 수 없습니다.';
  const docType = state.activeDocument?.isWorkshared ? '워크셋 문서' : '일반 문서';
  const extractCount = getExtractParameterRows(state).length;
  const extraCount = Number(!!state.includeCoordinates) + Number(!!state.includeLinearMetrics);
  state.ui.runSummary.textContent = `현재 문서: ${title}\n문서 유형: ${docType}\n등록된 RVT: ${state.rvtPaths.length}개\n추출 항목: ${extractCount + extraCount}개`;

  if (state.busy) {
    state.ui.runGuide.textContent = '작업이 진행 중입니다. 완료되면 결과 요약과 파일 경로를 보여줍니다.';
  } else if (getBatchOpenBlockingReason(state)) {
    state.ui.runGuide.textContent = '추출 파라미터 또는 좌표/선형 옵션을 1개 이상 설정하면 검토 버튼이 활성화됩니다.';
  } else {
  state.ui.runGuide.textContent = '필터 파라미터는 인스턴스/타입을 모두 검사합니다. 활성문서는 즉시 검토하고, 여러 RVT는 목록 창에서 시작합니다.';
  }
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

  const extractNames = getExtractParameterRows(state);
  const conditions = getConditionRows(state);
  const ready = !getBatchBlockingReason(state);

  state.ui.batchReadyCaption.textContent = ready ? '필터 검토와 추출 준비가 완료되었습니다.' : '추출 설정이 더 필요합니다.';
  state.ui.batchReadyHint.textContent = state.rvtPaths.length
    ? `${state.rvtPaths.length}개 RVT가 등록되어 있습니다.`
    : 'RVT를 추가하면 바로 배치 검토를 시작할 수 있습니다.';
  state.ui.batchReadyBadge.textContent = ready ? '검토 준비됨' : '설정 필요';

  state.ui.batchReadyList.innerHTML = '';
  [
    `필터 ${conditions.length}개`,
    `추출 파라미터 ${extractNames.length}개`,
    `단위 ${formatLengthUnit(state.lengthUnit)} / ${formatAreaUnit(state.areaUnit)} / ${formatVolumeUnit(state.volumeUnit)}`,
    `추가 옵션 ${[state.includeCoordinates ? '좌표' : null, state.includeLinearMetrics ? '선형' : null].filter(Boolean).join(' / ') || '없음'}`,
    '필터 파라미터는 인스턴스와 타입을 모두 검사합니다.',
    'Central File은 Detach로 열고, 워크셋은 모두 닫은 상태로 처리합니다.'
  ].forEach((line) => {
    const item = document.createElement('li');
    item.textContent = line;
    state.ui.batchReadyList.append(item);
  });
}

function updateActionState(state) {
  const activeReason = getActiveBlockingReason(state);
  const batchOpenReason = getBatchOpenBlockingReason(state);
  const batchRunReason = getBatchBlockingReason(state);

  if (state.ui.applyActiveBtn) state.ui.applyActiveBtn.disabled = !!activeReason || state.busy;
  if (state.ui.openBatchBtn) state.ui.openBatchBtn.disabled = !!batchOpenReason || state.busy;
  if (state.ui.batchReadyRunBtn) state.ui.batchReadyRunBtn.disabled = !!batchRunReason || state.busy;
}

function getActiveBlockingReason(state) {
  if (!hasExtractionTarget(state)) return '추출 항목을 1개 이상 설정해 주세요.';
  if (!state.activeDocument) return '활성 문서를 찾을 수 없습니다.';
  return '';
}

function getBatchOpenBlockingReason(state) {
  if (!hasExtractionTarget(state)) return '추출 항목을 1개 이상 설정해 주세요.';
  return '';
}


function buildPayload(state, useActiveDocument) {
  return {
    useActiveDocument: !!useActiveDocument,
    rvtPaths: useActiveDocument ? [] : getCheckedRvtPaths(state),
    outputFolder: '',
    closeAllWorksetsOnOpen: true,
    includeCoordinates: !!state.includeCoordinates,
    includeLinearMetrics: !!state.includeLinearMetrics,
    lengthUnit: state.lengthUnit,
    areaUnit: state.areaUnit,
    volumeUnit: state.volumeUnit,
    extractParameterNamesCsv: state.extractParameterNamesCsv,
    elementParameterUpdate: {
      enabled: true,
      combinationMode: state.elementParameterUpdate.combinationMode,
      conditions: state.elementParameterUpdate.conditions.map((row) => ({
        ...row,
        enabled: !!String(row.parameterName || '').trim()
      })),
      assignments: []
    }
  };
}

function createConditionState() {
  return {
    combinationMode: 'And',
    conditions: [createEmptyConditionRow()]
  };
}

function normalizeConditionState(source) {
  const base = createConditionState();
  if (!source) return base;
  base.combinationMode = source.combinationMode || 'And';
  if (Array.isArray(source.conditions) && source.conditions.length) {
    base.conditions = source.conditions.map((row) => ({
      ...createEmptyConditionRow(),
      ...(row || {})
    }));
  }
  if (!base.conditions.length) base.conditions = [createEmptyConditionRow()];
  return base;
}

function getConditionRows(state) {
  return state.elementParameterUpdate.conditions.filter((row) => String(row.parameterName || '').trim());
}

function getExtractParameterRows(state) {
  return String(state.extractParameterNamesCsv || '')
    .split(/[,\n\r;]+/)
    .map((value) => value.trim())
    .filter(Boolean)
    .filter((value, index, arr) => arr.indexOf(value) === index);
}

function hasExtractionTarget(state) {
  return getExtractParameterRows(state).length > 0 || state.includeCoordinates || state.includeLinearMetrics;
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

function getCheckedRvtPaths(state) {
  return state.rvtPaths.filter((path) => state.checked.has(path));
}

function getBatchBlockingReason(state) {
  if (!hasExtractionTarget(state)) return '추출 항목을 1개 이상 설정해 주세요.';
  if (!state.rvtPaths.length) return '검토할 RVT를 1개 이상 등록해 주세요.';
  if (!getCheckedRvtPaths(state).length) return '실행할 RVT를 1개 이상 선택해 주세요.';
  return '';
}

function actionButton(label, onClick, variant = 'secondary') {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = `btn ${variant === 'primary' ? 'btn--primary' : 'btn--secondary'}`;
  btn.textContent = label;
  btn.addEventListener('click', onClick);
  return btn;
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

function createSettingsTab(state, key, label) {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'deliverycleaner-tab';
  btn.textContent = label;
  btn.addEventListener('click', () => {
    state.activeSettingsTab = key;
    syncSettingsTab(state);
  });
  return btn;
}

function syncSettingsTab(state) {
  const activeKey = state.activeSettingsTab || 'filters';
  Object.entries(state.ui.settingsTabs || {}).forEach(([key, button]) => {
    if (!button) return;
    button.classList.toggle('is-active', key === activeKey);
  });
  Object.entries(state.ui.settingsPanels || {}).forEach(([key, panel]) => {
    if (!panel) return;
    panel.classList.toggle('is-active', key === activeKey);
  });
}

function createUnitField(label, options, onChange) {
  const wrap = div('deliverycleaner-field');
  const labelEl = document.createElement('label');
  labelEl.textContent = label;
  const select = document.createElement('select');
  options.forEach(({ value, label: text }) => {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = text;
    select.append(opt);
  });
  select.addEventListener('change', () => onChange(select.value));
  wrap.append(labelEl, select);
  return { wrap, select };
}

function createEmptyConditionRow() {
  return { parameterName: '', operatorName: 'Equals', value: '' };
}

function removeConditionRow(state, index) {
  if (!Array.isArray(state.elementParameterUpdate.conditions)) {
    state.elementParameterUpdate.conditions = [createEmptyConditionRow()];
  }

  if (state.elementParameterUpdate.conditions.length <= 1) {
    state.elementParameterUpdate.conditions = [createEmptyConditionRow()];
  } else {
    state.elementParameterUpdate.conditions.splice(index, 1);
  }

  renderConditionRows(state);
  renderSettingSummary(state);
}

function tdWithAction(onRemove, disabled = false) {
  const td = document.createElement('td');
  td.className = 'conditionextract-action-cell';
  const btn = actionButton('삭제', onRemove, 'secondary');
  btn.classList.add('conditionextract-row-btn');
  btn.disabled = !!disabled;
  td.append(btn);
  return td;
}

function formatLengthUnit(value) {
  return value === 'ft' ? 'ft' : value === 'inch' ? 'inch' : 'mm';
}

function formatAreaUnit(value) {
  return value === 'ft2' ? 'ft^2' : value === 'in2' ? 'inch^2' : 'mm^2';
}

function formatVolumeUnit(value) {
  return value === 'ft3' ? 'ft^3' : value === 'in3' ? 'inch^3' : 'mm^3';
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
  setBusy(on, on ? '조건별 객체 속성 추출 작업을 처리하는 중입니다.' : '');
  updateActionState(state);
}
