import { clear, div, toast, setBusy } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { post, onHost } from '../core/bridge.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js';

const FILTER_OPERATORS = [
  'Equals', 'NotEquals', 'Contains', 'NotContains', 'BeginsWith', 'NotBeginsWith',
  'EndsWith', 'NotEndsWith', 'Greater', 'GreaterOrEqual', 'Less', 'LessOrEqual',
  'HasValue', 'HasNoValue'
];

const TAB_KEYS = ['view', 'element', 'filter'];
const EXCEL_PHASE_WEIGHT = { EXCEL_INIT: 0.05, EXCEL_WRITE: 0.85, EXCEL_SAVE: 0.08, AUTOFIT: 0.02, DONE: 1, ERROR: 1 };

export function renderDeliveryCleaner(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  const state = {
    activeTab: 'view',
    filePaths: [],
    checked: new Set(),
    outputFolder: '',
    target3DViewName: 'KKY_CLEAN_3D',
    extractParameterNamesCsv: '',
    viewParameters: createViewParameterRows(),
    useFilter: false,
    applyFilterInitially: true,
    autoEnableFilterIfEmpty: false,
    filterProfile: null,
    elementParameterUpdate: createElementUpdateState(),
    logs: [],
    lastLogExportPath: '',
    session: null,
    busy: false,
    purgeSnapshot: null,
    filterDocItems: [],
    progressPercent: 0,
    ui: {}
  };

  const page = div('feature-shell deliverycleaner-page');
  page.innerHTML = `
    <div class="feature-header deliverycleaner-header">
      <div class="feature-heading">
        <span class="feature-kicker">납품시 BQC검토 · 별도 워크플로우</span>
        <h2 class="feature-title">RVT 정리 (납품용)</h2>
        <p class="feature-sub">링크 정리, 뷰·객체 파라미터 입력, 뷰 필터 적용, 검토, 속성 추출, Purge를 허브 안에서 이어서 실행합니다.</p>
      </div>
    </div>
  `;

  const controlCard = buildControlCard(state);
  const filesCard = buildFilesCard(state);
  const settingsModal = buildSettingsModal(state);
  const extractModal = buildExtractModal(state);
  const filterDocModal = buildFilterDocModal(state);
  const topGrid = div('deliverycleaner-top-grid');
  topGrid.append(controlCard, filesCard);

  page.append(topGrid, settingsModal, extractModal, filterDocModal);
  target.append(page);

  onHost('deliverycleaner:init', (payload) => applyHostState(state, payload));
  onHost('deliverycleaner:rvts-picked', (payload) => {
    const paths = Array.isArray(payload?.paths) ? payload.paths : [];
    if (!paths.length) return;

    paths.forEach((path) => {
      if (!state.filePaths.includes(path)) state.filePaths.push(path);
      state.checked.add(path);
    });

    syncStateToInputs(state);
    renderRvtList(state);
    updateActionState(state);
  });
  onHost('deliverycleaner:output-folder-picked', (payload) => {
    if (!payload?.path) return;
    state.outputFolder = payload.path;
    syncStateToInputs(state);
    updateActionState(state);
  });
  onHost('deliverycleaner:filter-loaded', (payload) => {
    if (!payload?.profile) return;
    state.filterProfile = normalizeFilterProfile(payload.profile);
    state.useFilter = true;
    syncStateToInputs(state);
    renderFilterPreview(state);
    updateActionState(state);
    if (payload?.source) toast(`필터를 불러왔습니다: ${payload.source}`, 'ok');
  });
  onHost('deliverycleaner:filter-saved', (payload) => {
    if (payload?.path) toast(`필터 XML 저장 완료: ${payload.path}`, 'ok');
  });
  onHost('deliverycleaner:filter-doc-list', (payload) => {
    state.filterDocItems = Array.isArray(payload?.items) ? payload.items : [];
    openFilterDocModal(state, payload?.docTitle || '');
  });
  onHost('deliverycleaner:progress', (payload) => {
    handleDeliveryCleanerProgress(state, payload);
  });
  onHost('deliverycleaner:run-done', (payload) => {
    resetDeliveryCleanerProgress(state);
    setPageBusy(state, false);
    applyHostState(state, payload?.state || {});
    const summary = payload?.summary || {};
    toast(`정리 완료 · 성공 ${summary.successCount ?? 0} / 실패 ${summary.failCount ?? 0}`, summary.failCount ? 'err' : 'ok', 3200);
  });
  onHost('deliverycleaner:verify-done', (payload) => {
    resetDeliveryCleanerProgress(state);
    setPageBusy(state, false);
    applyHostState(state, payload?.state || {});
    toast(payload?.path ? `정리 결과 검토 엑셀 생성: ${payload.path}` : '정리 결과 검토가 완료되었습니다.', 'ok', 3200);
  });
  onHost('deliverycleaner:extract-done', (payload) => {
    resetDeliveryCleanerProgress(state);
    setPageBusy(state, false);
    applyHostState(state, payload?.state || {});
    closeExtractModal(state);
    if (typeof payload?.parameterNamesCsv === 'string') {
      state.extractParameterNamesCsv = payload.parameterNamesCsv;
      syncStateToInputs(state);
    }
    toast(payload?.path ? `속성값 추출 엑셀 생성: ${payload.path}` : '속성값 추출이 완료되었습니다.', 'ok', 3200);
  });
  onHost('deliverycleaner:purge-started', (payload) => {
    applyHostState(state, payload?.state || {});
    state.purgeSnapshot = payload?.snapshot || state.purgeSnapshot;
    renderPurgeStatus(state);
    startPurgePolling(state);
    updateActionState(state);
    toast('Purge 일괄처리를 시작했습니다.', 'ok');
  });
  onHost('deliverycleaner:purge-status', (payload) => {
    state.purgeSnapshot = payload?.snapshot || null;
    renderPurgeStatus(state);
    updateActionState(state);

    const snapshot = state.purgeSnapshot || {};
    if (!snapshot.isRunning && (snapshot.isCompleted || snapshot.isFaulted)) {
      stopPurgePolling(state);
      resetDeliveryCleanerProgress(state);
      setPageBusy(state, false);
      if (snapshot.isCompleted) toast('Purge 일괄처리가 완료되었습니다.', 'ok', 3200);
      if (snapshot.isFaulted) toast(snapshot.message || 'Purge 진행 중 오류가 발생했습니다.', 'err', 3600);
    }
  });
  onHost('deliverycleaner:log-exported', (payload) => {
    resetDeliveryCleanerProgress(state);
    setPageBusy(state, false);
    applyHostState(state, payload?.state || {});
    if (payload?.path) toast(`로그 엑셀 저장 완료: ${payload.path}`, 'ok', 3200);
  });
  onHost('deliverycleaner:folder-opened', (payload) => {
    if (payload?.ok) toast('결과 폴더를 열었습니다.', 'ok');
  });
  onHost('deliverycleaner:log', (payload) => {
    appendLog(state, payload?.message || '');
  });
  onHost('deliverycleaner:error', (payload) => {
    resetDeliveryCleanerProgress(state);
    setPageBusy(state, false);
    stopPurgePolling(state);
    if (payload?.message) {
      appendLog(state, `[오류] ${payload.message}`);
      toast(payload.message, 'err', 3600);
    }
  });

  renderTabs(state);
  renderViewParameterRows(state);
  renderElementUpdateRows(state);
  renderFilterPreview(state);
  renderRvtList(state);
  renderPurgeStatus(state);
  updateActionState(state);
  post('deliverycleaner:init', {});
}

function buildControlCard(state) {
  const card = div('deliverycleaner-card deliverycleaner-card--control');
  card.innerHTML = `
    <div class="deliverycleaner-card__head">
      <div>
        <h3>실행 및 결과</h3>
        <p>작업 버튼, 현재 상태, Purge 진행, 로그 저장을 한 자리에서 확인합니다.</p>
      </div>
    </div>
  `;

  const buttonGrid = div('deliverycleaner-action-grid');
  const runBtn = actionButton('정리 시작', () => {
    setPageBusy(state, true);
    post('deliverycleaner:run', buildPayload(state));
  }, 'primary');
  const verifyBtn = actionButton('정리 결과 검토', () => {
    setPageBusy(state, true);
    post('deliverycleaner:verify', buildPayload(state));
  });
  const extractBtn = actionButton('속성값 추출', () => {
    openExtractModal(state);
  });
  const purgeBtn = actionButton('Purge 일괄처리', () => {
    setPageBusy(state, true);
    post('deliverycleaner:purge', buildPayload(state));
  });
  const folderBtn = actionButton('결과 폴더 열기', () => {
    post('deliverycleaner:open-folder', { path: state.outputFolder || state.session?.outputFolder || '' });
  });
  const exportLogBtn = actionButton('로그 엑셀 저장', () => {
    setPageBusy(state, true);
    post('deliverycleaner:export-log', { outputFolder: state.outputFolder || state.session?.outputFolder || '' });
  });

  buttonGrid.append(runBtn, verifyBtn, extractBtn, purgeBtn, folderBtn, exportLogBtn);

  const settingsBtn = actionButton('설정하기 (기본/세부)', () => openSettingsModal(state));
  settingsBtn.classList.add('deliverycleaner-settings-trigger');

  const statusBox = div('deliverycleaner-status');
  const purgeStatus = div('deliverycleaner-purge');
  const resultBox = div('deliverycleaner-summary-box');
  card.append(buttonGrid, settingsBtn, statusBox, purgeStatus, resultBox);

  state.ui.runBtn = runBtn;
  state.ui.verifyBtn = verifyBtn;
  state.ui.extractBtn = extractBtn;
  state.ui.purgeBtn = purgeBtn;
  state.ui.folderBtn = folderBtn;
  state.ui.exportLogBtn = exportLogBtn;
  state.ui.settingsBtn = settingsBtn;
  state.ui.status = statusBox;
  state.ui.purgeStatus = purgeStatus;
  state.ui.resultSummary = resultBox;

  return card;
}

function buildFilesCard(state) {
  const card = div('deliverycleaner-card deliverycleaner-card--files');
  card.innerHTML = `
    <div class="deliverycleaner-card__head">
      <div>
        <h3>대상 RVT</h3>
        <p>정리할 납품 대상 모델을 선택합니다.</p>
      </div>
      <div class="deliverycleaner-chip">필수</div>
    </div>
  `;

  const actions = div('deliverycleaner-inline-actions');
  const addBtn = actionButton('RVT 추가', () => post('deliverycleaner:pick-rvts', {}), 'primary');
  const removeBtn = actionButton('선택 제거', () => {
    state.filePaths = state.filePaths.filter((path) => !state.checked.has(path));
    state.checked.clear();
    renderRvtList(state);
    updateActionState(state);
  });
  const clearBtn = actionButton('목록 비우기', () => {
    state.filePaths = [];
    state.checked.clear();
    renderRvtList(state);
    updateActionState(state);
  });
  actions.append(addBtn, removeBtn, clearBtn);

  const tableWrap = div('deliverycleaner-table-wrap');
  const { table, tbody, master } = createRvtTable();
  table.classList.add('deliverycleaner-rvt-table');
  tableWrap.append(table);

  card.append(actions, tableWrap);
  state.ui.rvtBody = tbody;
  state.ui.rvtMaster = master;
  return card;
}

function buildSettingsModal(state) {
  const overlay = div('deliverycleaner-modal deliverycleaner-settings-modal is-hidden');
  overlay.innerHTML = `
    <div class="deliverycleaner-modal__dialog deliverycleaner-settings-modal__dialog">
      <div class="deliverycleaner-modal__head">
        <div>
          <h3>설정하기 (기본/세부)</h3>
          <p>기본 설정과 세부 설정을 한 창에서 정리합니다.</p>
        </div>
        <button type="button" class="deliverycleaner-modal__close" data-close>&times;</button>
      </div>
      <div class="deliverycleaner-modal__body" data-settings-body></div>
      <div class="deliverycleaner-modal__foot">
        <button type="button" class="btn btn--secondary" data-cancel>닫기</button>
      </div>
    </div>
  `;

  const settingsBody = overlay.querySelector('[data-settings-body]');
  const settingsLayout = div('deliverycleaner-settings-layout');
  const basicsCard = buildBasicsCard(state);
  const workspaceCard = buildWorkspaceCard(state);
  settingsLayout.append(basicsCard, workspaceCard);
  settingsBody.append(settingsLayout);

  overlay.querySelector('[data-close]').addEventListener('click', () => closeSettingsModal(state));
  overlay.querySelector('[data-cancel]').addEventListener('click', () => closeSettingsModal(state));
  overlay.addEventListener('click', (ev) => {
    if (ev.target === overlay) closeSettingsModal(state);
  });

  state.ui.settingsOverlay = overlay;
  return overlay;
}

function buildBasicsCard(state) {
  const card = div('deliverycleaner-card deliverycleaner-card--basics');
  card.innerHTML = `
    <div class="deliverycleaner-card__head">
      <div>
        <h3>기본 설정</h3>
        <p>결과 폴더와 납품 정리 기준이 되는 기본 항목을 지정합니다.</p>
      </div>
    </div>
  `;

  const fields = div('deliverycleaner-field-stack');

  const outputField = fieldBlock('정리 결과 폴더');
  const outputRow = div('deliverycleaner-input-row');
  const outputInput = document.createElement('input');
  outputInput.type = 'text';
  outputInput.placeholder = '정리 결과 폴더';
  outputInput.addEventListener('input', () => {
    state.outputFolder = outputInput.value.trim();
    updateActionState(state);
  });
  const browseBtn = actionButton('찾아보기', () => post('deliverycleaner:browse-output-folder', {}));
  outputRow.append(outputInput, browseBtn);
  outputField.append(outputRow);

  const viewNameField = fieldBlock('정리용 3D 뷰 이름');
  const viewNameInput = document.createElement('input');
  viewNameInput.type = 'text';
  viewNameInput.placeholder = '예: KKY_CLEAN_3D';
  viewNameInput.addEventListener('input', () => {
    state.target3DViewName = viewNameInput.value.trim();
  });
  viewNameField.append(viewNameInput);

  const extractField = fieldBlock('속성값 추출 파라미터');
  const extractInput = document.createElement('textarea');
  extractInput.rows = 2;
  extractInput.placeholder = '예: Comments, Mark, Type Comments';
  extractInput.addEventListener('input', () => {
    state.extractParameterNamesCsv = extractInput.value.trim();
  });
  extractField.append(extractInput);

  fields.append(outputField, viewNameField);
  card.append(fields);

  state.ui.outputFolderInput = outputInput;
  state.ui.viewNameInput = viewNameInput;
  return card;
}

function buildWorkspaceCard(state) {
  const card = div('deliverycleaner-card deliverycleaner-card--workspace');
  card.innerHTML = `
    <div class="deliverycleaner-card__head">
      <div>
        <h3>세부 설정</h3>
        <p>탭을 전환하면서 뷰 파라미터, 객체 파라미터, 뷰 필터를 정리합니다.</p>
      </div>
    </div>
  `;

  const tabBar = div('deliverycleaner-tabs');
  TAB_KEYS.forEach((key) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'deliverycleaner-tab';
    btn.textContent = key === 'view' ? '뷰 파라미터' : key === 'element' ? '객체 파라미터' : '뷰 필터';
    btn.addEventListener('click', () => {
      state.activeTab = key;
      renderTabs(state);
    });
    tabBar.append(btn);
  });

  const panelWrap = div('deliverycleaner-panels');

  const viewPanel = div('deliverycleaner-panel');
  viewPanel.innerHTML = `<div class="deliverycleaner-note">최대 5개까지 뷰 파라미터를 지정해 정리용 3D 뷰에 입력합니다.</div>`;
  const viewScroll = div('deliverycleaner-table-scroll deliverycleaner-table-scroll--grid');
  const viewTable = document.createElement('table');
  viewTable.className = 'deliverycleaner-grid-table deliverycleaner-grid-table--view';
  viewTable.innerHTML = `
    <thead><tr><th>사용</th><th>파라미터 이름</th><th>입력값</th></tr></thead>
    <tbody></tbody>
  `;
  viewScroll.append(viewTable);
  viewPanel.append(viewScroll);

  const elementPanel = div('deliverycleaner-panel');
  elementPanel.innerHTML = `
    <div class="deliverycleaner-section-callout">
      <label class="deliverycleaner-checkline"><input type="checkbox" data-element-enabled> 객체 파라미터 일괄 입력 사용</label>
      <div class="deliverycleaner-inline-select">
        <span>조건 결합</span>
        <select data-combination-mode>
          <option value="And">AND</option>
          <option value="Or">OR</option>
        </select>
      </div>
    </div>
    <div class="deliverycleaner-split-grid">
      <section class="deliverycleaner-subsection">
        <div class="deliverycleaner-subsection__head">
          <h4>조건</h4>
          <p>최대 4개</p>
        </div>
        <div class="deliverycleaner-table-scroll deliverycleaner-table-scroll--grid">
          <table class="deliverycleaner-grid-table deliverycleaner-grid-table--conditions">
            <thead><tr><th>사용</th><th>파라미터</th><th>연산자</th><th>값</th></tr></thead>
            <tbody data-condition-body></tbody>
          </table>
        </div>
      </section>
      <section class="deliverycleaner-subsection">
        <div class="deliverycleaner-subsection__head">
          <h4>입력</h4>
          <p>최대 4개</p>
        </div>
        <div class="deliverycleaner-table-scroll deliverycleaner-table-scroll--grid">
          <table class="deliverycleaner-grid-table deliverycleaner-grid-table--assignments">
            <thead><tr><th>사용</th><th>파라미터</th><th>값</th></tr></thead>
            <tbody data-assignment-body></tbody>
          </table>
        </div>
      </section>
    </div>
    <div class="deliverycleaner-summary-box" data-element-summary></div>
  `;

  const filterPanel = div('deliverycleaner-panel');
  filterPanel.innerHTML = `
    <div class="deliverycleaner-filter-top">
      <label class="deliverycleaner-checkline"><input type="checkbox" data-use-filter> 필터 사용</label>
      <label class="deliverycleaner-checkline"><input type="checkbox" data-apply-filter> 최초 열기 시 적용</label>
      <label class="deliverycleaner-checkline"><input type="checkbox" data-auto-enable-filter> 뷰가 비면 자동 활성화</label>
    </div>
    <div class="deliverycleaner-inline-actions">
      <button type="button" class="btn btn--secondary" data-filter-import>XML 가져오기</button>
      <button type="button" class="btn btn--secondary" data-filter-save>XML 저장</button>
      <button type="button" class="btn btn--secondary" data-filter-doc>문서 필터 추출</button>
    </div>
    <div class="deliverycleaner-filter-preview">
      <section class="deliverycleaner-subsection deliverycleaner-subsection--compact">
        <div class="deliverycleaner-subsection__head">
          <h4>카테고리</h4>
          <p>필터에 포함된 카테고리</p>
        </div>
        <div class="deliverycleaner-category-list" data-category-list></div>
      </section>
      <section class="deliverycleaner-subsection">
        <div class="deliverycleaner-subsection__head">
          <h4>조건</h4>
          <p>Revit 필터 설정창과 비슷한 구조로 표시합니다.</p>
        </div>
        <div class="deliverycleaner-table-scroll deliverycleaner-table-scroll--grid deliverycleaner-table-scroll--filter">
          <table class="deliverycleaner-grid-table deliverycleaner-grid-table--filter-preview">
            <thead><tr><th>Join</th><th>Group</th><th>Parameter</th><th>Operator</th><th>Value</th></tr></thead>
            <tbody data-filter-condition-body></tbody>
          </table>
        </div>
      </section>
    </div>
    <div class="deliverycleaner-summary-box" data-filter-summary></div>
  `;

  panelWrap.append(viewPanel, elementPanel, filterPanel);
  card.append(tabBar, panelWrap);

  state.ui.tabButtons = Array.from(tabBar.querySelectorAll('.deliverycleaner-tab'));
  state.ui.panels = { view: viewPanel, element: elementPanel, filter: filterPanel };
  state.ui.viewParamBody = viewTable.querySelector('tbody');
  state.ui.elementEnabled = elementPanel.querySelector('[data-element-enabled]');
  state.ui.combinationMode = elementPanel.querySelector('[data-combination-mode]');
  state.ui.conditionBody = elementPanel.querySelector('[data-condition-body]');
  state.ui.assignmentBody = elementPanel.querySelector('[data-assignment-body]');
  state.ui.elementSummary = elementPanel.querySelector('[data-element-summary]');
  state.ui.useFilter = filterPanel.querySelector('[data-use-filter]');
  state.ui.applyFilter = filterPanel.querySelector('[data-apply-filter]');
  state.ui.autoEnableFilter = filterPanel.querySelector('[data-auto-enable-filter]');
  state.ui.filterImportBtn = filterPanel.querySelector('[data-filter-import]');
  state.ui.filterSaveBtn = filterPanel.querySelector('[data-filter-save]');
  state.ui.filterDocBtn = filterPanel.querySelector('[data-filter-doc]');
  state.ui.categoryList = filterPanel.querySelector('[data-category-list]');
  state.ui.filterConditionBody = filterPanel.querySelector('[data-filter-condition-body]');
  state.ui.filterSummary = filterPanel.querySelector('[data-filter-summary]');

  state.ui.elementEnabled.addEventListener('change', () => {
    state.elementParameterUpdate.enabled = state.ui.elementEnabled.checked;
    renderElementUpdateSummary(state);
  });
  state.ui.combinationMode.addEventListener('change', () => {
    state.elementParameterUpdate.combinationMode = state.ui.combinationMode.value;
    renderElementUpdateSummary(state);
  });
  state.ui.useFilter.addEventListener('change', () => {
    state.useFilter = state.ui.useFilter.checked;
    updateActionState(state);
  });
  state.ui.applyFilter.addEventListener('change', () => {
    state.applyFilterInitially = state.ui.applyFilter.checked;
  });
  state.ui.autoEnableFilter.addEventListener('change', () => {
    state.autoEnableFilterIfEmpty = state.ui.autoEnableFilter.checked;
  });
  state.ui.filterImportBtn.addEventListener('click', () => post('deliverycleaner:filter-import', {}));
  state.ui.filterSaveBtn.addEventListener('click', () => post('deliverycleaner:filter-save', { filterProfile: state.filterProfile }));
  state.ui.filterDocBtn.addEventListener('click', () => post('deliverycleaner:filter-doc-list', {}));

  return card;
}

function buildFilterDocModal(state) {
  const overlay = div('deliverycleaner-modal is-hidden');
  overlay.innerHTML = `
    <div class="deliverycleaner-modal__dialog">
      <div class="deliverycleaner-modal__head">
        <div>
          <h3>현재 문서 필터 추출</h3>
          <p data-doc-title></p>
        </div>
        <button type="button" class="deliverycleaner-modal__close" data-close>&times;</button>
      </div>
      <div class="deliverycleaner-modal__body">
        <div class="deliverycleaner-doclist" data-doc-list></div>
      </div>
      <div class="deliverycleaner-modal__foot">
        <button type="button" class="btn btn--secondary" data-cancel>닫기</button>
      </div>
    </div>
  `;

  overlay.querySelector('[data-close]').addEventListener('click', () => closeFilterDocModal(state));
  overlay.querySelector('[data-cancel]').addEventListener('click', () => closeFilterDocModal(state));
  overlay.addEventListener('click', (ev) => {
    if (ev.target === overlay) closeFilterDocModal(state);
  });

  state.ui.filterDocOverlay = overlay;
  state.ui.filterDocTitle = overlay.querySelector('[data-doc-title]');
  state.ui.filterDocList = overlay.querySelector('[data-doc-list]');
  return overlay;
}

function buildExtractModal(state) {
  const overlay = div('deliverycleaner-modal deliverycleaner-extract-modal is-hidden');
  overlay.innerHTML = `
    <div class="deliverycleaner-modal__dialog deliverycleaner-extract-modal__dialog">
      <div class="deliverycleaner-modal__head">
        <div>
          <h3>속성값 추출</h3>
          <p>추출할 파라미터를 입력한 뒤 이 창에서 바로 엑셀 추출을 실행합니다.</p>
        </div>
        <button type="button" class="deliverycleaner-modal__close" data-close>&times;</button>
      </div>
      <div class="deliverycleaner-modal__body">
        <div class="deliverycleaner-field-stack">
          <div class="deliverycleaner-note">
            New Schedule/Quantities에서 리스트업 가능한 카테고리의 객체만 추출 대상으로 사용합니다.
            Centerline, 주석, 일반 선처럼 스케줄 대상이 아닌 객체는 제외됩니다.
          </div>
          <div class="deliverycleaner-field">
            <label>추출 파라미터</label>
            <textarea rows="5" data-extract-input placeholder="예: Comments, Mark, Type Comments"></textarea>
          </div>
          <div class="deliverycleaner-summary-box" data-extract-summary></div>
        </div>
      </div>
      <div class="deliverycleaner-modal__foot">
        <button type="button" class="btn btn--secondary" data-cancel>닫기</button>
        <button type="button" class="btn btn--primary" data-run>엑셀로 추출</button>
      </div>
    </div>
  `;

  overlay.querySelector('[data-close]').addEventListener('click', () => closeExtractModal(state));
  overlay.querySelector('[data-cancel]').addEventListener('click', () => closeExtractModal(state));
  overlay.addEventListener('click', (ev) => {
    if (ev.target === overlay) closeExtractModal(state);
  });

  const extractInput = overlay.querySelector('[data-extract-input]');
  const runBtn = overlay.querySelector('[data-run]');

  extractInput.addEventListener('input', () => {
    state.extractParameterNamesCsv = extractInput.value.trim();
    renderExtractModalSummary(state);
  });

  runBtn.addEventListener('click', () => {
    if (!getDeliveryCleanerExtractionTargetCount(state)) {
      toast('추출 대상 RVT가 없습니다. 먼저 RVT를 추가하거나 정리 결과를 준비해주세요.', 'err', 3200);
      return;
    }

    if (!state.extractParameterNamesCsv.trim()) {
      toast('추출할 파라미터 이름을 입력해주세요.', 'err', 3200);
      return;
    }

    setPageBusy(state, true);
    post('deliverycleaner:extract', buildPayload(state));
  });

  state.ui.extractOverlay = overlay;
  state.ui.extractInput = extractInput;
  state.ui.extractRunBtn = runBtn;
  state.ui.extractSummary = overlay.querySelector('[data-extract-summary]');
  return overlay;
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

function createViewParameterRows() {
  return Array.from({ length: 5 }, () => ({ enabled: false, parameterName: '', parameterValue: '' }));
}

function createElementUpdateState() {
  return {
    enabled: false,
    combinationMode: 'And',
    conditions: Array.from({ length: 4 }, () => ({ enabled: false, parameterName: '', operatorName: 'Equals', value: '' })),
    assignments: Array.from({ length: 4 }, () => ({ enabled: false, parameterName: '', value: '' }))
  };
}

function renderTabs(state) {
  state.ui.tabButtons.forEach((btn, index) => {
    const key = TAB_KEYS[index];
    btn.classList.toggle('is-active', state.activeTab === key);
  });

  Object.entries(state.ui.panels).forEach(([key, panel]) => {
    panel.classList.toggle('is-active', state.activeTab === key);
  });
}

function renderRvtList(state) {
  const master = state.ui.rvtMaster;
  master.checked = state.filePaths.length > 0 && state.filePaths.every((path) => state.checked.has(path));
  master.disabled = state.filePaths.length === 0;
  master.onchange = () => {
    if (master.checked) state.checked = new Set(state.filePaths);
    else state.checked.clear();
    renderRvtList(state);
  };

  const rows = state.filePaths.map((path, index) => ({
    checked: state.checked.has(path),
    index: index + 1,
    name: getRvtName(path),
    path,
    title: path,
    onToggle: (checked) => {
      if (checked) state.checked.add(path);
      else state.checked.delete(path);
    }
  }));

  renderRvtRows(state.ui.rvtBody, rows, '등록된 RVT가 없습니다.');
}

function renderViewParameterRows(state) {
  const body = state.ui.viewParamBody;
  body.innerHTML = '';

  state.viewParameters.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.append(
      tdWithCheck(row.enabled, (checked) => { state.viewParameters[index].enabled = checked; }),
      tdWithInput(row.parameterName, '파라미터 이름', (value) => { state.viewParameters[index].parameterName = value; }),
      tdWithInput(row.parameterValue, '입력값', (value) => { state.viewParameters[index].parameterValue = value; })
    );
    body.append(tr);
  });
}

function renderElementUpdateRows(state) {
  const conditionBody = state.ui.conditionBody;
  conditionBody.innerHTML = '';
  state.elementParameterUpdate.conditions.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.append(
      tdWithCheck(row.enabled, (checked) => { state.elementParameterUpdate.conditions[index].enabled = checked; renderElementUpdateSummary(state); }),
      tdWithInput(row.parameterName, '파라미터 이름', (value) => { state.elementParameterUpdate.conditions[index].parameterName = value; renderElementUpdateSummary(state); }),
      tdWithSelect(row.operatorName, FILTER_OPERATORS, (value) => { state.elementParameterUpdate.conditions[index].operatorName = value; renderElementUpdateSummary(state); }),
      tdWithInput(row.value, '값', (value) => { state.elementParameterUpdate.conditions[index].value = value; renderElementUpdateSummary(state); })
    );
    conditionBody.append(tr);
  });

  const assignmentBody = state.ui.assignmentBody;
  assignmentBody.innerHTML = '';
  state.elementParameterUpdate.assignments.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.append(
      tdWithCheck(row.enabled, (checked) => { state.elementParameterUpdate.assignments[index].enabled = checked; renderElementUpdateSummary(state); }),
      tdWithInput(row.parameterName, '파라미터 이름', (value) => { state.elementParameterUpdate.assignments[index].parameterName = value; renderElementUpdateSummary(state); }),
      tdWithInput(row.value, '값', (value) => { state.elementParameterUpdate.assignments[index].value = value; renderElementUpdateSummary(state); })
    );
    assignmentBody.append(tr);
  });

  renderElementUpdateSummary(state);
}

function renderElementUpdateSummary(state) {
  const enabled = state.elementParameterUpdate.enabled;
  const conds = state.elementParameterUpdate.conditions.filter((row) => row.enabled && row.parameterName.trim());
  const assigns = state.elementParameterUpdate.assignments.filter((row) => row.enabled && row.parameterName.trim());
  const joiner = state.elementParameterUpdate.combinationMode === 'Or' ? ' OR ' : ' AND ';
  const conditionText = conds.length
    ? conds.map((row) => (row.operatorName === 'HasValue' || row.operatorName === 'HasNoValue')
      ? `${row.parameterName} ${row.operatorName}`
      : `${row.parameterName} ${row.operatorName} ${row.value || ''}`).join(joiner)
    : '조건 없음';
  const assignmentText = assigns.length
    ? assigns.map((row) => `${row.parameterName} = ${row.value || ''}`).join(' / ')
    : '입력 없음';

  state.ui.elementSummary.textContent = enabled
    ? `조건: ${conditionText}\n입력: ${assignmentText}`
    : '사용 안 함';
}

function renderFilterPreview(state) {
  state.ui.useFilter.checked = !!state.useFilter;
  state.ui.applyFilter.checked = !!state.applyFilterInitially;
  state.ui.autoEnableFilter.checked = !!state.autoEnableFilterIfEmpty;

  const categoryList = state.ui.categoryList;
  const conditionBody = state.ui.filterConditionBody;
  const summaryBox = state.ui.filterSummary;
  categoryList.innerHTML = '';
  conditionBody.innerHTML = '';

  if (!state.filterProfile || !isFilterConfigured(state.filterProfile)) {
    summaryBox.textContent = '필터가 아직 준비되지 않았습니다. XML 가져오기 또는 현재 문서 필터 추출을 사용하세요.';
    return;
  }

  getCategoryTokens(state.filterProfile.categoriesCsv).forEach((name) => {
    const item = div('deliverycleaner-category-chip');
    item.textContent = name;
    categoryList.append(item);
  });

  const rows = buildFilterConditionRows(state.filterProfile);
  rows.forEach((row) => {
    const tr = document.createElement('tr');
    ['join', 'group', 'parameter', 'operator', 'value'].forEach((key) => {
      const td = document.createElement('td');
      td.textContent = row[key] || '';
      tr.append(td);
    });
    conditionBody.append(tr);
  });

  const parts = [
    `Filter: ${state.filterProfile.filterName || ''}`,
    `Categories: ${state.filterProfile.categoriesCsv || ''}`
  ];
  if (state.filterProfile.structureSummary) parts.push('', 'Structure:', state.filterProfile.structureSummary);
  summaryBox.textContent = parts.join('\n');
}

function buildFilterConditionRows(profile) {
  const rows = [];
  if (!profile) return rows;

  let root = null;
  if (profile.filterDefinitionXml) {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(profile.filterDefinitionXml, 'application/xml');
      root = doc.documentElement;
    } catch {
      root = null;
    }
  }

  if (root) collectConditionRows(root, rows, '', '');
  if (!rows.length && profile.parameterToken) {
    rows.push({
      join: '',
      group: 'Rule 1',
      parameter: profile.parameterToken,
      operator: profile.operatorName || '',
      value: profile.ruleValue || ''
    });
  }
  return rows;
}

function collectConditionRows(node, rows, parentJoin, inheritedPath) {
  if (!node) return;
  if (node.nodeName === 'Logical') {
    const logicalType = String(node.getAttribute('Type') || 'And').toUpperCase();
    Array.from(node.children).forEach((child, index) => {
      const group = inheritedPath ? `${inheritedPath} > ${logicalType} ${index + 1}` : `${logicalType} ${index + 1}`;
      const join = index === 0 ? parentJoin : logicalType;
      collectConditionRows(child, rows, join, group);
    });
    return;
  }

  rows.push({
    join: parentJoin || '',
    group: inheritedPath || 'Rule 1',
    parameter: node.getAttribute('Parameter') || node.getAttribute('Param') || node.getAttribute('ParameterToken') || '',
    operator: node.getAttribute('Operator') || node.nodeName || '',
    value: node.getAttribute('Value') || node.textContent || ''
  });
}

function appendLog(state, message) {
  if (!message) return;
  state.logs.push(message);
  if (state.logs.length > 2000) state.logs.splice(0, state.logs.length - 2000);
  updateActionState(state);
}

function renderPurgeStatus(state) {
  const box = state.ui.purgeStatus;
  const snap = state.purgeSnapshot || {};

  if (!snap || (!snap.isRunning && !snap.isCompleted && !snap.isFaulted)) {
    box.textContent = 'Purge 대기 중\n정리된 결과 파일을 대상으로 일괄 처리합니다.';
    return;
  }

  const fileText = snap.totalFiles ? `${snap.currentFileIndex || 0}/${snap.totalFiles}` : '-';
  const iterText = snap.totalIterations ? `${snap.currentIteration || 0}/${snap.totalIterations}` : '-';
  const chunks = [
    `Purge 상태: ${snap.stateName || '대기'}`,
    `파일 ${fileText}`,
    `반복 ${iterText}`
  ];
  if (snap.currentFileName) chunks.push(snap.currentFileName);
  if (snap.message) chunks.push(snap.message);
  box.textContent = chunks.join('\n');
}

function renderResultSummary(state) {
  const lines = [];
  const session = state.session || {};
  const cleanedCount = Array.isArray(session.cleanedOutputPaths) ? session.cleanedOutputPaths.length : 0;
  const successCount = Array.isArray(session.results) ? session.results.filter((item) => item?.success).length : 0;
  const failCount = Array.isArray(session.results) ? session.results.filter((item) => item && item.success === false).length : 0;

  if (!cleanedCount && !session.verificationCsvPath && !state.lastLogExportPath) {
    lines.push('정리 시작');
    lines.push('선택한 RVT를 정리해 결과 폴더에 저장합니다.');
    lines.push('');
    lines.push('정리 결과 검토');
    lines.push('정리된 결과를 검토해서 엑셀 파일로 대상 폴더에 저장합니다.');
    lines.push('Design Option 유무 검토 결과도 같은 대상 폴더에 함께 저장됩니다.');
    lines.push('');
    lines.push('속성값 추출');
    lines.push('선택한 RVT에 존재하는 모든 모델 객체의 속성 정보를 추출합니다.');
    lines.push('추출 결과 엑셀도 대상 폴더에 저장됩니다.');
    lines.push('');
    lines.push('로그 엑셀 저장');
    lines.push('실행 로그를 결과 폴더에 엑셀 파일로 저장합니다.');
  } else {
    lines.push(`정리 결과: ${cleanedCount ? `${cleanedCount}개 파일 준비` : '아직 없음'}`);
    if (Array.isArray(session.results) && session.results.length) {
      lines.push(`최근 실행 요약: 성공 ${successCount} / 실패 ${failCount}`);
    }
    lines.push(`정리 결과 검토 엑셀: ${session.verificationCsvPath || '없음'}`);
    lines.push(`Design Option 검토 엑셀: ${session.designOptionAuditCsvPath || '없음'}`);
    lines.push('Design Option 유무 검토 결과는 대상 폴더에 함께 저장됩니다.');
    lines.push('속성값 추출은 선택한 RVT에 존재하는 모든 모델 객체의 속성 정보를 엑셀로 저장합니다.');
    lines.push(`로그 엑셀: ${state.lastLogExportPath || '아직 저장 안 함'}`);
  }

  state.ui.resultSummary.textContent = lines.join('\n');
}

function isDeliveryCleanerConfigured(state) {
  return !!state.outputFolder;
}

function getDeliveryCleanerExtractionTargetCount(state) {
  if (state.filePaths.length) return state.filePaths.length;
  if (Array.isArray(state.session?.cleanedOutputPaths)) return state.session.cleanedOutputPaths.length;
  return 0;
}

function stripLogStamp(line) {
  if (!line) return '';
  return String(line).replace(/^\[[^\]]+\]\s*/, '').trim();
}

function extractValueAfterColonText(line) {
  const text = stripLogStamp(line);
  const index = text.indexOf(': ');
  if (index >= 0 && index < text.length - 2) return text.substring(index + 2).trim();
  return text;
}

function getFileNameOnly(path) {
  if (!path) return '';
  const text = String(path).trim();
  const tokens = text.split(/[\\/]/);
  return tokens[tokens.length - 1] || text;
}

function buildProgressSnapshot(state) {
  const snapshot = {
    mode: '',
    currentFile: '',
    currentTask: '',
    completedCount: 0,
    totalCount: state.filePaths.length || (Array.isArray(state.session?.cleanedOutputPaths) ? state.session.cleanedOutputPaths.length : 0)
  };

  if (state.purgeSnapshot?.isRunning) {
    snapshot.mode = 'purge';
    snapshot.currentFile = state.purgeSnapshot.currentFileName || '';
    snapshot.currentTask = state.purgeSnapshot.message || state.purgeSnapshot.stateName || 'Purge 진행 중';
    snapshot.completedCount = Math.max(0, (state.purgeSnapshot.currentFileIndex || 1) - 1);
    snapshot.totalCount = state.purgeSnapshot.totalFiles || snapshot.totalCount;
    return snapshot;
  }

  state.logs.forEach((rawLine) => {
    const line = stripLogStamp(rawLine);
    if (!line) return;

    if (line.startsWith('정리 시작: ') || line.startsWith('대상 시작: ')) {
      snapshot.mode = 'clean';
      snapshot.currentFile = getFileNameOnly(extractValueAfterColonText(line));
      snapshot.currentTask = '정리 준비 중';
      return;
    }

    if (line.startsWith('[STEP] ')) {
      snapshot.mode = 'clean';
      snapshot.currentTask = line.substring(7).trim() || '정리 진행 중';
      return;
    }

    if (line.startsWith('정리 및 저장 완료: ') || line.startsWith('저장 완료: ')) {
      snapshot.mode = 'clean';
      snapshot.currentFile = getFileNameOnly(extractValueAfterColonText(line));
      snapshot.currentTask = '정리 저장 완료';
      snapshot.completedCount += 1;
      return;
    }

    if (line.startsWith('실패: ')) {
      snapshot.currentTask = '오류 확인 필요';
      return;
    }

    if (line.startsWith('검토 파일 열기: ')) {
      snapshot.mode = 'verify';
      snapshot.currentFile = getFileNameOnly(extractValueAfterColonText(line));
      snapshot.currentTask = '정리 결과 검토 중';
      return;
    }

    if (line.startsWith('검토 완료: ')) {
      snapshot.mode = 'verify';
      snapshot.currentTask = '정리 결과 검토 완료';
      snapshot.completedCount += 1;
      return;
    }

    if (line.startsWith('속성값 추출 파일 열기: ')) {
      snapshot.mode = 'extract';
      snapshot.currentFile = getFileNameOnly(extractValueAfterColonText(line));
      snapshot.currentTask = '속성값 추출 중';
      return;
    }

    if (line.startsWith('속성값 추출 완료: ')) {
      snapshot.mode = 'extract';
      snapshot.currentTask = '속성값 추출 완료';
      snapshot.completedCount += 1;
      return;
    }

    if (line.startsWith('Purge 실행: ')) {
      snapshot.mode = 'purge';
      snapshot.currentFile = getFileNameOnly(extractValueAfterColonText(line));
      snapshot.currentTask = 'Purge 일괄처리 중';
      return;
    }

    if (line.startsWith('Purge 후 저장 완료: ')) {
      snapshot.mode = 'purge';
      snapshot.currentFile = getFileNameOnly(extractValueAfterColonText(line));
      snapshot.currentTask = 'Purge 후 저장 완료';
      snapshot.completedCount += 1;
    }
  });

  return snapshot;
}

function renderStatusBox(state, hasFiles, isConfigured, hasSessionTargets, isPurging) {
  const progress = buildProgressSnapshot(state);
  const lines = [];
  let tone = 'idle';

  if (state.busy || isPurging) {
    tone = 'active';
    lines.push('작업 진행 중');
    lines.push(progress.currentTask ? `현재 작업: ${progress.currentTask}` : '현재 작업: 진행 정보 확인 중');
    lines.push(progress.currentFile ? `현재 파일: ${progress.currentFile}` : '현재 파일: 대기 중');
    if (progress.totalCount > 0) lines.push(`진행 파일: ${Math.min(progress.completedCount, progress.totalCount)}/${progress.totalCount}`);
  } else if (!hasFiles) {
    tone = 'idle';
    lines.push('대상 RVT를 먼저 추가해 주세요.');
    lines.push('RVT를 추가하면 설정 버튼이 강조되고, 설정 완료 후 정리 시작이 활성화됩니다.');
  } else if (!isConfigured) {
    tone = 'required';
    lines.push(`대상 RVT ${state.filePaths.length}개가 준비되었습니다.`);
    lines.push('다음 단계: 설정하기 (기본/세부)에서 결과 폴더를 지정해 주세요.');
    lines.push('설정이 완료되면 정리 시작 버튼이 활성화됩니다.');
  } else {
    tone = 'ready';
    lines.push(`대상 RVT ${state.filePaths.length}개와 결과 폴더 설정이 준비되었습니다.`);
    lines.push('정리 시작을 눌러 납품용 정리를 실행할 수 있습니다.');
    if (hasSessionTargets) lines.push(`최근 정리 결과 ${state.session.cleanedOutputPaths.length}개 파일이 준비되어 있습니다.`);
  }

  if (state.useFilter && state.filterProfile && isFilterConfigured(state.filterProfile) && !(state.busy || isPurging)) {
    lines.push('뷰 필터 설정도 적용 가능한 상태입니다.');
  }

  state.ui.status.classList.remove('is-idle', 'is-required', 'is-ready', 'is-active');
  state.ui.status.classList.add(`is-${tone}`);
  state.ui.status.textContent = lines.join('\n');
}

function updateActionState(state) {
  const hasFiles = state.filePaths.length > 0;
  const hasOutput = !!state.outputFolder;
  const isConfigured = isDeliveryCleanerConfigured(state);
  const hasSessionTargets = Array.isArray(state.session?.cleanedOutputPaths) && state.session.cleanedOutputPaths.length > 0;
  const isPurging = !!state.purgeSnapshot?.isRunning;
  const canRun = !state.busy && hasFiles && isConfigured;

  state.ui.runBtn.disabled = !canRun;
  state.ui.verifyBtn.disabled = state.busy || !(hasFiles || hasSessionTargets);
  state.ui.extractBtn.disabled = state.busy || !(hasFiles || hasSessionTargets);
  state.ui.purgeBtn.disabled = state.busy || isPurging || !(hasFiles || hasSessionTargets);
  state.ui.folderBtn.disabled = state.busy || !(state.outputFolder || state.session?.outputFolder);
  state.ui.exportLogBtn.disabled = state.busy || !state.logs.length;

  state.ui.settingsBtn.classList.toggle('is-required', hasFiles && !isConfigured && !state.busy);
  state.ui.settingsBtn.classList.toggle('is-complete', hasFiles && isConfigured && !state.busy);
  state.ui.runBtn.classList.toggle('deliverycleaner-run-ready', canRun);
  state.ui.runBtn.classList.toggle('deliverycleaner-run-blocked', hasFiles && !isConfigured && !state.busy);

  renderStatusBox(state, hasFiles, isConfigured, hasSessionTargets, isPurging);
  renderResultSummary(state);
  renderExtractModalSummary(state);
}

function applyHostState(state, payload) {
  const settings = payload?.settings || {};
  const session = payload?.session || null;

  if (Array.isArray(settings.filePaths)) {
    state.filePaths = [...settings.filePaths];
    state.checked = new Set(state.filePaths);
  }
  if (typeof settings.outputFolder === 'string') state.outputFolder = settings.outputFolder;
  if (typeof settings.target3DViewName === 'string' && settings.target3DViewName) state.target3DViewName = settings.target3DViewName;
  if (Array.isArray(settings.viewParameters) && settings.viewParameters.length) {
    state.viewParameters = createViewParameterRows().map((row, index) => ({ ...row, ...(settings.viewParameters[index] || {}) }));
  }
  if (typeof settings.useFilter === 'boolean') state.useFilter = settings.useFilter;
  if (typeof settings.applyFilterInitially === 'boolean') state.applyFilterInitially = settings.applyFilterInitially;
  if (typeof settings.autoEnableFilterIfEmpty === 'boolean') state.autoEnableFilterIfEmpty = settings.autoEnableFilterIfEmpty;
  if (settings.filterProfile) state.filterProfile = normalizeFilterProfile(settings.filterProfile);
  if (settings.elementParameterUpdate) state.elementParameterUpdate = normalizeElementUpdate(settings.elementParameterUpdate);
  if (typeof payload?.extractParameterNamesCsv === 'string') state.extractParameterNamesCsv = payload.extractParameterNamesCsv;
  if (Array.isArray(payload?.logs)) state.logs = [...payload.logs];
  if (typeof payload?.lastLogExportPath === 'string') state.lastLogExportPath = payload.lastLogExportPath;
  if (session) state.session = session;
  if (payload?.purge) state.purgeSnapshot = payload.purge;

  syncStateToInputs(state);
  renderViewParameterRows(state);
  renderElementUpdateRows(state);
  renderFilterPreview(state);
  renderRvtList(state);
  renderPurgeStatus(state);
  updateActionState(state);
}

function syncStateToInputs(state) {
  state.ui.outputFolderInput.value = state.outputFolder || '';
  state.ui.viewNameInput.value = state.target3DViewName || '';
  if (state.ui.extractInput) state.ui.extractInput.value = state.extractParameterNamesCsv || '';
  state.ui.elementEnabled.checked = !!state.elementParameterUpdate.enabled;
  state.ui.combinationMode.value = state.elementParameterUpdate.combinationMode || 'And';
}

function normalizeFilterProfile(profile) {
  return {
    filterName: profile.filterName || '',
    categoriesCsv: profile.categoriesCsv || '',
    parameterToken: profile.parameterToken || '',
    operatorName: profile.operatorName || 'Equals',
    ruleValue: profile.ruleValue || '',
    filterDefinitionXml: profile.filterDefinitionXml || '',
    structureSummary: profile.structureSummary || ''
  };
}

function normalizeElementUpdate(source) {
  const base = createElementUpdateState();
  base.enabled = !!source.enabled;
  base.combinationMode = source.combinationMode || 'And';
  base.conditions = base.conditions.map((row, index) => ({ ...row, ...(Array.isArray(source.conditions) ? source.conditions[index] || {} : {}) }));
  base.assignments = base.assignments.map((row, index) => ({ ...row, ...(Array.isArray(source.assignments) ? source.assignments[index] || {} : {}) }));
  return base;
}

function buildPayload(state) {
  return {
    filePaths: [...state.filePaths],
    outputFolder: state.outputFolder,
    target3DViewName: state.target3DViewName,
    extractParameterNamesCsv: state.extractParameterNamesCsv,
    viewParameters: state.viewParameters.map((row) => ({ ...row })),
    useFilter: state.useFilter,
    applyFilterInitially: state.applyFilterInitially,
    autoEnableFilterIfEmpty: state.autoEnableFilterIfEmpty,
    filterProfile: state.filterProfile ? { ...state.filterProfile } : null,
    elementParameterUpdate: {
      enabled: state.elementParameterUpdate.enabled,
      combinationMode: state.elementParameterUpdate.combinationMode,
      conditions: state.elementParameterUpdate.conditions.map((row) => ({ ...row })),
      assignments: state.elementParameterUpdate.assignments.map((row) => ({ ...row }))
    }
  };
}

function handleDeliveryCleanerProgress(state, payload) {
  if (!payload) return;

  if (payload.phase || payload.current != null || payload.total != null) {
    const phase = normalizeDeliveryCleanerExcelPhase(payload?.phase);
    const total = Number(payload?.total) || 0;
    const current = Number(payload?.current) || 0;
    const percent = computeDeliveryCleanerExcelPercent(state, phase, current, total, payload?.phaseProgress, payload?.percent);
    const subtitle = buildDeliveryCleanerExcelSubtitle(phase, current, total);
    const detail = payload?.message || '';

    ProgressDialog.show(payload?.title || 'RVT 정리 (납품용)', subtitle || '진행 중...');
    ProgressDialog.update(percent, subtitle || '진행 중...', detail);

    if (phase === 'DONE' || phase === 'ERROR') {
      window.setTimeout(() => resetDeliveryCleanerProgress(state), 260);
    }
    return;
  }

  const title = payload?.title || 'RVT 정리 (납품용)';
  const message = payload?.message || '진행 중...';
  const detail = payload?.detail || '';
  const percent = clampDeliveryCleanerPercent(payload?.percent);
  state.progressPercent = Math.max(state.progressPercent || 0, percent);
  ProgressDialog.show(title, message);
  ProgressDialog.update(state.progressPercent, message, detail);

  if (payload?.complete) {
    window.setTimeout(() => resetDeliveryCleanerProgress(state), 260);
  }
}

function resetDeliveryCleanerProgress(state) {
  state.progressPercent = 0;
  ProgressDialog.hide();
}

function normalizeDeliveryCleanerExcelPhase(phase) {
  return String(phase || '').trim().toUpperCase() || 'EXCEL_WRITE';
}

function clampDeliveryCleanerPercent(value) {
  const n = Number(value);
  return Number.isFinite(n) ? Math.max(0, Math.min(100, n)) : 0;
}

function clampDeliveryCleanerRatio(value) {
  const n = Number(value);
  return Number.isFinite(n) ? Math.max(0, Math.min(1, n)) : 0;
}

function computeDeliveryCleanerExcelPercent(state, phase, current, total, phaseProgress, percentOverride) {
  const norm = normalizeDeliveryCleanerExcelPhase(phase);
  if (norm === 'DONE') {
    state.progressPercent = 100;
    return 100;
  }
  if (norm === 'ERROR') return state.progressPercent || 0;

  if (typeof percentOverride === 'number' && Number.isFinite(percentOverride) && percentOverride > 0 && percentOverride <= 1) {
    state.progressPercent = Math.max(state.progressPercent || 0, percentOverride * 100);
    return state.progressPercent;
  }

  const completed = ['EXCEL_INIT', 'EXCEL_WRITE', 'EXCEL_SAVE', 'AUTOFIT'].reduce((acc, key) => {
    if (key === norm) return acc;
    return acc + (EXCEL_PHASE_WEIGHT[key] || 0);
  }, 0);
  const weight = EXCEL_PHASE_WEIGHT[norm] || 0;
  const ratio = total > 0 ? Math.max(0, Math.min(1, current / total)) : 0;
  const staged = Math.max(ratio, clampDeliveryCleanerRatio(phaseProgress));
  const denominator = completed + weight || 1;
  const percent = (completed + weight * staged) / denominator * 100;
  state.progressPercent = Math.max(state.progressPercent || 0, Math.min(100, percent));
  return state.progressPercent;
}

function buildDeliveryCleanerExcelSubtitle(phase, current, total) {
  switch (normalizeDeliveryCleanerExcelPhase(phase)) {
    case 'EXCEL_INIT': return '엑셀 저장 준비 중...';
    case 'EXCEL_WRITE': return `엑셀 데이터 작성 중... (${current}/${Math.max(total, current || 1)})`;
    case 'EXCEL_SAVE': return '엑셀 파일 저장 중...';
    case 'AUTOFIT': return '열 너비 자동 조정 중...';
    case 'DONE': return '내보내기 완료';
    case 'ERROR': return '내보내기 오류';
    default: return '진행 중...';
  }
}

function setPageBusy(state, on) {
  state.busy = !!on;
  setBusy(on, on ? 'RVT 정리 (납품용) 작업 중' : '');
  updateActionState(state);
}

function startPurgePolling(state) {
  stopPurgePolling(state);
  state.ui.purgeTimer = window.setInterval(() => {
    post('deliverycleaner:purge-status', {});
  }, 2000);
}

function stopPurgePolling(state) {
  if (state.ui.purgeTimer) {
    window.clearInterval(state.ui.purgeTimer);
    state.ui.purgeTimer = null;
  }
}

function openSettingsModal(state) {
  state.ui.settingsOverlay?.classList.remove('is-hidden');
}

function closeSettingsModal(state) {
  state.ui.settingsOverlay?.classList.add('is-hidden');
}

function openExtractModal(state) {
  renderExtractModalSummary(state);
  state.ui.extractOverlay?.classList.remove('is-hidden');
  state.ui.extractInput?.focus();
}

function closeExtractModal(state) {
  state.ui.extractOverlay?.classList.add('is-hidden');
}

function openFilterDocModal(state, docTitle) {
  state.ui.filterDocTitle.textContent = docTitle ? `현재 문서: ${docTitle}` : '현재 문서 필터 목록';
  state.ui.filterDocList.innerHTML = '';

  if (!state.filterDocItems.length) {
    const empty = div('deliverycleaner-empty');
    empty.textContent = '추출할 문서 필터가 없습니다.';
    state.ui.filterDocList.append(empty);
  } else {
    state.filterDocItems.forEach((item) => {
      const row = document.createElement('button');
      row.type = 'button';
      row.className = 'deliverycleaner-doclist__item';
      row.textContent = item.name || '이름 없음';
      row.addEventListener('click', () => {
        closeFilterDocModal(state);
        post('deliverycleaner:filter-doc-extract', { filterId: item.id });
      });
      state.ui.filterDocList.append(row);
    });
  }

  state.ui.filterDocOverlay.classList.remove('is-hidden');
}

function closeFilterDocModal(state) {
  state.ui.filterDocOverlay.classList.add('is-hidden');
}

function renderExtractModalSummary(state) {
  if (!state.ui.extractSummary) return;

  const targetCount = getDeliveryCleanerExtractionTargetCount(state);
  const outputFolder = state.outputFolder || state.session?.outputFolder || '';
  const parameterCount = (state.extractParameterNamesCsv || '')
    .split(/[,\n;\r]+/)
    .map((item) => item.trim())
    .filter(Boolean)
    .length;

  const lines = [
    `대상 RVT: ${targetCount ? `${targetCount}개` : '없음'}`,
    `추출 파라미터: ${parameterCount ? `${parameterCount}개` : '입력 필요'}`,
    `결과 저장 폴더: ${outputFolder || '설정 필요'}`,
    '',
    '추출 결과는 대상 폴더에 엑셀로 저장됩니다.'
  ];

  state.ui.extractSummary.textContent = lines.join('\n');

  if (state.ui.extractRunBtn) {
    state.ui.extractRunBtn.disabled = state.busy || !targetCount || !parameterCount;
  }
}

function isFilterConfigured(profile) {
  return !!(profile?.filterName && profile?.categoriesCsv && (profile?.filterDefinitionXml || (profile?.parameterToken && profile?.ruleValue != null)));
}

function getCategoryTokens(csv) {
  return String(csv || '')
    .split(/[,\n\r;]+/)
    .map((token) => token.trim())
    .filter(Boolean)
    .filter((token, index, arr) => arr.indexOf(token) === index);
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

function tdWithCheck(value, onChange) {
  const td = document.createElement('td');
  td.className = 'deliverycleaner-cell-center';
  const input = document.createElement('input');
  input.type = 'checkbox';
  input.checked = !!value;
  input.addEventListener('change', () => onChange(input.checked));
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
