import { clear, div, toast, setBusy, showCompletionSummaryDialog, showExcelSavedDialog } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { chooseExcelMode } from '../core/dom.js';
import { attachRvtDropZone } from '../core/rvtDrop.js';
import { post, onHost } from '../core/bridge.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js';

const FILTER_OPERATORS = [
  'Equals', 'NotEquals', 'Contains', 'NotContains', 'BeginsWith', 'NotBeginsWith',
  'EndsWith', 'NotEndsWith', 'Greater', 'GreaterOrEqual', 'Less', 'LessOrEqual',
  'HasValue', 'HasNoValue'
];
const TEXT_FILTER_OPERATORS = [
  'Equals', 'NotEquals', 'Contains', 'NotContains', 'BeginsWith', 'NotBeginsWith',
  'EndsWith', 'NotEndsWith', 'HasValue', 'HasNoValue'
];

const TAB_KEYS = ['view', 'element', 'filter', 'vg'];
const EXCEL_PHASE_WEIGHT = { EXCEL_INIT: 0.05, EXCEL_WRITE: 0.85, EXCEL_SAVE: 0.08, AUTOFIT: 0.02, DONE: 1, ERROR: 1 };

function refreshUiAfterHostDialog(render, delay = 120) {
  if (typeof render !== 'function') return;

  const run = () => {
    try { render(); } catch { }
  };

  run();

  if (typeof window === 'undefined') return;

  let released = false;
  let raf1 = 0;
  let raf2 = 0;
  let timerRun = 0;
  let timerFinalize = 0;

  const cleanup = () => {
    if (released) return;
    released = true;
    window.removeEventListener('focus', onFocus, true);
    if (typeof document !== 'undefined') {
      document.removeEventListener('visibilitychange', onVisible, true);
    }
    if (raf1 && typeof window.cancelAnimationFrame === 'function') window.cancelAnimationFrame(raf1);
    if (raf2 && typeof window.cancelAnimationFrame === 'function') window.cancelAnimationFrame(raf2);
    if (timerRun) window.clearTimeout(timerRun);
    if (timerFinalize) window.clearTimeout(timerFinalize);
  };

  const rerender = () => {
    cleanup();
    run();
  };

  const onFocus = () => rerender();
  const onVisible = () => {
    if (document.visibilityState === 'visible') rerender();
  };

  window.addEventListener('focus', onFocus, true);
  if (typeof document !== 'undefined') {
    document.addEventListener('visibilitychange', onVisible, true);
  }

  if (typeof window.requestAnimationFrame === 'function') {
    raf1 = window.requestAnimationFrame(() => {
      run();
      raf2 = window.requestAnimationFrame(run);
    });
  } else {
    timerRun = window.setTimeout(run, 0);
  }

  timerFinalize = window.setTimeout(rerender, Math.max(0, Number(delay) || 0));
}

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
    visibilityRules: createVisibilityRuleState(),
    availableVisibilityCategories: [],
    visibilityCategoryPicker: {
      rowIndex: -1,
      searchText: '',
      selectedNames: []
    },
    elementParameterUpdate: createElementUpdateState(),
    logs: [],
    lastLogExportPath: '',
    session: null,
    busy: false,
    purgeSnapshot: null,
    purgeResultShown: false,
    filterDocItems: [],
    progressPercent: 0,
    acceptProgress: false,
    ui: {}
  };

  const page = div('feature-shell deliverycleaner-page');
  page.innerHTML = `
    <div class="feature-header deliverycleaner-header">
      <div class="feature-heading">
        <span class="feature-kicker">납품시 BQC검토 · 파일 정리</span>
        <h2 class="feature-title">RVT 정리 (납품용)</h2>
        <p class="feature-sub">링크 정리, 뷰/객체 파라미터 입력, 뷰 필터 적용, 검토, 속성 추출, Purge를 허브 안에서 이어서 실행합니다.</p>
      </div>
    </div>
  `;

  const controlCard = buildControlCard(state);
  const filesCard = buildFilesCard(state);
  const settingsModal = buildSettingsModal(state);
  const extractModal = buildExtractModal(state);
  const filterDocModal = buildFilterDocModal(state);
  const visibilityCategoryModal = buildVisibilityCategoryModal(state);
  const topGrid = div('deliverycleaner-top-grid');
  topGrid.append(controlCard, filesCard);

  page.append(topGrid, settingsModal, extractModal, filterDocModal, visibilityCategoryModal);
  target.append(page);

  onHost('deliverycleaner:init', (payload) => applyHostState(state, payload));
  onHost('deliverycleaner:rvts-picked', (payload) => {
    const paths = Array.isArray(payload?.paths) ? payload.paths : [];
    if (!paths.length) return;

    paths.forEach((path) => {
      if (!state.filePaths.includes(path)) state.filePaths.push(path);
      state.checked.add(path);
    });

    refreshUiAfterHostDialog(() => {
      syncStateToInputs(state);
      renderRvtList(state);
      updateActionState(state)
    });
  });
  onHost('deliverycleaner:output-folder-picked', (payload) => {
    if (!payload?.path) return;
    state.outputFolder = payload.path;
    refreshUiAfterHostDialog(() => {
      syncStateToInputs(state);
      updateActionState(state);
    });
  });
  onHost('deliverycleaner:filter-loaded', (payload) => {
    if (!payload?.profile) return;
    state.filterProfile = normalizeFilterProfile(payload.profile);
    state.useFilter = true;
    refreshUiAfterHostDialog(() => {
      syncStateToInputs(state);
      renderFilterPreview(state);
      updateActionState(state);
    });
    if (payload?.source) toast(`필터 설정을 불러왔습니다: ${payload.source}`, 'ok');
  });
  onHost('deliverycleaner:filter-saved', (payload) => {
    if (payload?.path) toast(`필터 XML을 저장했습니다: ${payload.path}`, 'ok');
  });
  onHost('deliverycleaner:visibility-rules-loaded', (payload) => {
    if (!payload?.visibilityRules) return;
    state.visibilityRules = normalizeVisibilityRules(payload.visibilityRules);
    refreshUiAfterHostDialog(() => {
      syncStateToInputs(state);
      renderVisibilityRuleRows(state);
      renderVisibilityCategoryModal(state);
    });
    if (payload?.source) toast(`VV 규칙을 불러왔습니다: ${payload.source}`, 'ok');
  });
  onHost('deliverycleaner:visibility-rules-saved', (payload) => {
    if (payload?.path) toast(`VV 규칙 XML을 저장했습니다: ${payload.path}`, 'ok');
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
    toast(`정리 완료: 성공 ${summary.successCount ?? 0} / 실패 ${summary.failCount ?? 0}`, summary.failCount ? 'err' : 'ok', 3200);
    showDeliveryCleanerRunDialog(state, payload || {});
  });
  onHost('deliverycleaner:verify-done', (payload) => {
    resetDeliveryCleanerProgress(state);
    setPageBusy(state, false);
    applyHostState(state, payload?.state || {});
    toast(payload?.path ? `정리 결과 검토 파일 생성: ${payload.path}` : '정리 결과 검토가 완료되었습니다.', 'ok', 3200);
    showDeliveryCleanerVerifyDialog(state, payload || {});
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
    toast(payload?.path ? `속성값 추출 파일 생성: ${payload.path}` : '속성값 추출이 완료되었습니다.', 'ok', 3200);
    showDeliveryCleanerExtractDialog(state, payload || {});
  });
  onHost('deliverycleaner:purge-started', (payload) => {
    applyHostState(state, payload?.state || {});
    state.purgeSnapshot = payload?.snapshot || state.purgeSnapshot;
    state.purgeResultShown = false;
    renderPurgeStatus(state);
    startPurgePolling(state);
    updateActionState(state);
    toast('Purge 일괄처리를 시작했습니다.', 'ok');
  });
  onHost('deliverycleaner:purge-status', (payload) => {
    applyHostState(state, payload?.state || {});
    state.purgeSnapshot = payload?.snapshot || null;
    renderPurgeStatus(state);
    updateActionState(state);

    const snapshot = state.purgeSnapshot || {};
    if (!snapshot.isRunning && (snapshot.isCompleted || snapshot.isFaulted)) {
      stopPurgePolling(state);
      resetDeliveryCleanerProgress(state);
      setPageBusy(state, false);
      if (snapshot.isCompleted) toast('Purge 일괄처리가 완료되었습니다.', 'ok', 3200);
      if (snapshot.isFaulted) toast(snapshot.message || 'Purge 처리 중 오류가 발생했습니다.', 'err', 3600);
      if (snapshot.isCompleted && !state.purgeResultShown) {
        state.purgeResultShown = true;
        showDeliveryCleanerPurgeDialog(state, payload || {});
      }
    }
  });
  onHost('deliverycleaner:log-exported', (payload) => {
    resetDeliveryCleanerProgress(state);
    setPageBusy(state, false);
    applyHostState(state, payload?.state || {});
    if (payload?.path) toast(`로그 엑셀을 저장했습니다: ${payload.path}`, 'ok', 3200);
  });
  onHost('deliverycleaner:verify-exported', (payload) => {
    handleDeliveryCleanerWorkbookExported(state, '정리 결과 검토 엑셀을 저장했습니다.', payload);
  });
  onHost('deliverycleaner:extract-exported', (payload) => {
    handleDeliveryCleanerWorkbookExported(state, '속성값 추출 엑셀을 저장했습니다.', payload);
  });
  onHost('deliverycleaner:designoption-exported', (payload) => {
    handleDeliveryCleanerWorkbookExported(state, 'Design Option 검토 엑셀을 저장했습니다.', payload);
  });
  onHost('deliverycleaner:purge-exported', (payload) => {
    handleDeliveryCleanerWorkbookExported(state, 'Purge 객체수 비교 엑셀을 저장했습니다.', payload);
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

function beginDeliveryCleanerProgress(title, message, detail = '', percent = 0) {
  ProgressDialog.show(title || 'RVT 정리 (납품용)', message || '준비 중...');
  ProgressDialog.update(percent, message || '준비 중...', detail || '');
}

function buildControlCard(state) {
  const card = div('deliverycleaner-card deliverycleaner-card--control');
  card.innerHTML = `
    <div class="deliverycleaner-card__head">
      <div>
        <h3>실행 및 결과</h3>
        <p>정리 실행, 결과 검토, 속성값 추출, Purge 진행과 엑셀 내보내기를 한 자리에서 확인합니다.</p>
      </div>
    </div>
  `;

  const buttonGrid = div('deliverycleaner-action-grid');
  const runBtn = actionButton('정리 시작', () => {
    if (!canUseSelectedDeliveryCleanerFiles(state)) {
      toast('정리할 RVT를 1개 이상 선택하세요.', 'warn');
      updateActionState(state);
      return;
    }
    setPageBusy(state, true);
    beginDeliveryCleanerProgress('RVT 정리 (납품용)', '정리 작업을 준비 중입니다.');
    post('deliverycleaner:run', buildPayload(state));
  }, 'primary');
  const verifyBtn = actionButton('정리 결과 검토', () => {
    if (!canUseSelectedDeliveryCleanerFiles(state, { allowSessionFallback: true })) {
      toast('검토할 RVT를 1개 이상 선택하세요.', 'warn');
      updateActionState(state);
      return;
    }
    setPageBusy(state, true);
    beginDeliveryCleanerProgress('정리 결과 검토', '검토 작업을 준비 중입니다.');
    post('deliverycleaner:verify', buildPayload(state));
  });
  const extractBtn = actionButton('속성값 추출', () => {
    openExtractModal(state);
  });
  const purgeBtn = actionButton('Purge 일괄처리', () => {
    if (!canUseSelectedDeliveryCleanerFiles(state, { allowSessionFallback: true })) {
      toast('Purge 대상 RVT를 1개 이상 선택하세요.', 'warn');
      updateActionState(state);
      return;
    }
    setPageBusy(state, true);
    beginDeliveryCleanerProgress('Purge 일괄처리', 'Purge 작업을 준비 중입니다.');
    post('deliverycleaner:purge', buildPayload(state));
  });
  const folderBtn = actionButton('결과 폴더 열기', () => {
    post('deliverycleaner:open-folder', { path: state.outputFolder || state.session?.outputFolder || '' });
  });
  const exportLogBtn = actionButton('로그 엑셀 저장', () => {
    setPageBusy(state, true);
    beginDeliveryCleanerProgress('로그 엑셀 저장', '로그 엑셀 저장을 준비 중입니다.');
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
        <p>정리와 검토를 수행할 납품 대상 RVT를 선택합니다.</p>
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
  const hint = div('rvt-drop-hint');
  hint.textContent = 'RVT 추가 버튼을 누르거나 탐색기에서 .rvt 파일을 이 목록으로 끌어다 놓으면 바로 등록됩니다.';

  const tableWrap = div('deliverycleaner-table-wrap rvt-drop-zone');
  const { table, tbody, master } = createRvtTable();
  table.classList.add('deliverycleaner-rvt-table');
  tableWrap.append(table);
  attachRvtDropZone(tableWrap, {
    onDropPaths: (paths) => {
      const added = appendDroppedRvts(state, paths);
      if (!added) {
        toast('이미 등록된 RVT입니다.', 'warn');
        return;
      }
      renderRvtList(state);
      toast(`${added}개 RVT를 추가했습니다.`, 'ok');
    },
    onInvalid: () => toast('RVT 파일만 드래그해서 추가할 수 있습니다.', 'warn')
  });

  card.append(actions, hint, tableWrap);
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
          <p>기본 설정과 세부 설정을 한 화면에서 정리합니다.</p>
        </div>
        <button type="button" class="deliverycleaner-modal__close" data-close>&times;</button>
      </div>
      <div class="deliverycleaner-modal__body" data-settings-body></div>
      <div class="deliverycleaner-modal__foot">
        <button type="button" class="btn btn--primary" data-cancel>적용</button>
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
        <p>결과 폴더와 정리용 3D 뷰 이름 등 기본 항목을 지정합니다.</p>
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

  const extractField = fieldBlock('속성값 추출 기본 파라미터');
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
        <p>뷰 파라미터, 객체 파라미터, 뷰 필터, V/G 설정을 탭으로 전환하며 설정합니다.</p>
      </div>
    </div>
  `;

  const tabBar = div('deliverycleaner-tabs');
  TAB_KEYS.forEach((key) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'deliverycleaner-tab';
    btn.textContent = key === 'view'
      ? '뷰 파라미터'
      : key === 'element'
        ? '객체 파라미터'
        : key === 'filter'
          ? '뷰 필터'
          : 'V/G 설정';
    btn.addEventListener('click', () => {
      state.activeTab = key;
      renderTabs(state);
    });
    tabBar.append(btn);
  });

  const panelWrap = div('deliverycleaner-panels');

  const viewPanel = div('deliverycleaner-panel');
  viewPanel.innerHTML = `
    <div class="deliverycleaner-note">최대 5개까지 뷰 파라미터를 지정해 정리용 3D 뷰에 입력합니다.</div>
  `;
  const viewScroll = div('deliverycleaner-table-scroll deliverycleaner-table-scroll--grid');
  const viewTable = document.createElement('table');
  viewTable.className = 'deliverycleaner-grid-table deliverycleaner-grid-table--view';
  viewTable.innerHTML = `
    <thead><tr><th>파라미터 이름</th><th>값</th></tr></thead>
    <tbody></tbody>
  `;
  viewScroll.append(viewTable);
  viewPanel.append(viewScroll);

  const elementPanel = div('deliverycleaner-panel');
  elementPanel.innerHTML = `
    <div class="deliverycleaner-section-callout">
      <div class="deliverycleaner-note">조건과 입력을 채우면 자동으로 객체 파라미터 입력에 반영됩니다.</div>
      <div class="deliverycleaner-inline-controls">
        <div class="deliverycleaner-inline-select">
          <span>조건 결합</span>
          <select data-combination-mode>
            <option value="And">AND</option>
            <option value="Or">OR</option>
          </select>
        </div>
        <div class="deliverycleaner-inline-select">
          <span>중복 파라미터 처리</span>
          <select data-duplicate-mode>
            <option value="single">하나만 입력</option>
            <option value="all">중복 전체 입력</option>
          </select>
        </div>
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
            <thead><tr><th>파라미터</th><th>연산자</th><th>값</th></tr></thead>
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
            <thead><tr><th>파라미터</th><th>값</th></tr></thead>
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
      <label class="deliverycleaner-checkline"><input type="checkbox" data-auto-enable-filter>  뷰가 비면 자동 활성화</label>
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
          <p>필터에 포함할 카테고리</p>
        </div>
        <div class="deliverycleaner-category-list" data-category-list></div>
      </section>
      <section class="deliverycleaner-subsection">
        <div class="deliverycleaner-subsection__head">
          <h4>조건</h4>
          <p>Revit 필터 설정창과 비슷한 구조로 조건을 표시합니다.</p>
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

  const vgPanel = div('deliverycleaner-panel');
  vgPanel.innerHTML = `
    <section class="deliverycleaner-subsection">
      <div class="deliverycleaner-subsection__head">
        <h4>V/G 설정</h4>
        <p>Imported 카테고리 표시 상태와 커스텀 서브카테고리 숨김 규칙을 정리용 3D 뷰에 적용합니다.</p>
      </div>
      <div class="deliverycleaner-vg-import-row">
        <span class="deliverycleaner-vg-import-row__label">Imported 표시</span>
        <div class="deliverycleaner-vg-import-row__controls">
          <label class="deliverycleaner-inline-select deliverycleaner-inline-select--compact">
            <span>Categories</span>
            <select data-visibility-imported-toggle>
              <option value="">적용 안 함</option>
              <option value="show">체크</option>
              <option value="hide">해제</option>
            </select>
          </label>
          <label class="deliverycleaner-inline-select deliverycleaner-inline-select--compact">
            <span>Families</span>
            <select data-visibility-imports-family-toggle>
              <option value="">적용 안 함</option>
              <option value="show">체크</option>
              <option value="hide">해제</option>
            </select>
          </label>
        </div>
      </div>
      <div class="deliverycleaner-vg-rulebar">
        <div class="deliverycleaner-vg-rulebar__text">
          <strong>서브카테고리 숨김 규칙</strong>
          <span>주 카테고리, 연산자, 서브카테고리 텍스트를 입력해 정리용 3D 뷰의 V/G 기준을 맞춥니다.</span>
        </div>
        <div class="deliverycleaner-vg-rulebar__controls">
          <label class="deliverycleaner-inline-select deliverycleaner-inline-select--compact">
            <span>규칙 결합</span>
            <select data-visibility-combination-mode>
              <option value="Or">OR</option>
              <option value="And">AND</option>
            </select>
          </label>
          <button type="button" class="btn btn--secondary" data-visibility-import>규칙 불러오기</button>
          <button type="button" class="btn btn--secondary" data-visibility-save>규칙 저장</button>
        </div>
      </div>
      <div class="deliverycleaner-table-scroll deliverycleaner-table-scroll--grid deliverycleaner-table-scroll--visibility-rules">
        <table class="deliverycleaner-grid-table deliverycleaner-grid-table--visibility-rules">
          <thead><tr><th>주 카테고리들</th><th>연산자</th><th>서브카테고리 텍스트</th></tr></thead>
          <tbody data-visibility-rule-body></tbody>
        </table>
      </div>
      <div class="deliverycleaner-summary-box" data-visibility-summary></div>
    </section>
  `;

  panelWrap.append(viewPanel, elementPanel, filterPanel, vgPanel);
  card.append(tabBar, panelWrap);

  state.ui.tabButtons = Array.from(tabBar.querySelectorAll('.deliverycleaner-tab'));
  state.ui.panels = { view: viewPanel, element: elementPanel, filter: filterPanel, vg: vgPanel };
  state.ui.viewParamBody = viewTable.querySelector('tbody');
  state.ui.visibilityCombinationMode = vgPanel.querySelector('[data-visibility-combination-mode]');
  state.ui.visibilityImportBtn = vgPanel.querySelector('[data-visibility-import]');
  state.ui.visibilitySaveBtn = vgPanel.querySelector('[data-visibility-save]');
  state.ui.visibilityImportedToggle = vgPanel.querySelector('[data-visibility-imported-toggle]');
  state.ui.visibilityImportsFamilyToggle = vgPanel.querySelector('[data-visibility-imports-family-toggle]');
  state.ui.visibilityRuleBody = vgPanel.querySelector('[data-visibility-rule-body]');
  state.ui.visibilitySummary = vgPanel.querySelector('[data-visibility-summary]');
  state.ui.combinationMode = elementPanel.querySelector('[data-combination-mode]');
  state.ui.duplicateMode = elementPanel.querySelector('[data-duplicate-mode]');
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

  state.ui.visibilityCombinationMode.addEventListener('change', () => {
    state.visibilityRules.combinationMode = state.ui.visibilityCombinationMode.value;
    renderVisibilityRuleSummary(state);
  });
  state.ui.visibilityImportedToggle.addEventListener('change', () => {
    state.visibilityRules.showImportedCategoriesInView = parseVisibilityToggleValue(state.ui.visibilityImportedToggle.value);
    renderVisibilityRuleSummary(state);
  });
  state.ui.visibilityImportsFamilyToggle.addEventListener('change', () => {
    state.visibilityRules.showImportsInFamilies = parseVisibilityToggleValue(state.ui.visibilityImportsFamilyToggle.value);
    renderVisibilityRuleSummary(state);
  });
  state.ui.visibilityImportBtn.addEventListener('click', () => post('deliverycleaner:visibility-rules-import', {}));
  state.ui.visibilitySaveBtn.addEventListener('click', () => {
    post('deliverycleaner:visibility-rules-save', { visibilityRules: buildPayload(state).visibilityRules });
  });
  state.ui.combinationMode.addEventListener('change', () => {
    state.elementParameterUpdate.combinationMode = state.ui.combinationMode.value;
    renderElementUpdateSummary(state);
  });
  state.ui.duplicateMode.addEventListener('change', () => {
    state.elementParameterUpdate.applyToAllMatchingParameters = state.ui.duplicateMode.value === 'all';
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
          <h3>문서 필터 추출</h3>
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

function buildVisibilityCategoryModal(state) {
  const overlay = div('deliverycleaner-modal deliverycleaner-category-modal is-hidden');
  overlay.innerHTML = `
    <div class="deliverycleaner-modal__dialog deliverycleaner-category-modal__dialog">
      <div class="deliverycleaner-modal__head">
        <div>
          <h3>카테고리 선택</h3>
          <p>현재 문서의 V/G 카테고리 중에서 규칙에 적용할 주 카테고리를 여러 개 선택합니다.</p>
        </div>
        <button type="button" class="deliverycleaner-modal__close" data-close>&times;</button>
      </div>
      <div class="deliverycleaner-modal__body">
        <div class="deliverycleaner-field">
          <label>카테고리 검색</label>
          <input type="text" data-category-search placeholder="카테고리 이름 검색">
        </div>
        <div class="deliverycleaner-note" data-category-selection-summary></div>
        <div class="deliverycleaner-category-picker-list" data-category-picker-list></div>
      </div>
      <div class="deliverycleaner-modal__foot">
        <button type="button" class="btn btn--secondary" data-cancel>취소</button>
        <button type="button" class="btn btn--primary" data-apply>적용</button>
      </div>
    </div>
  `;

  overlay.querySelector('[data-close]').addEventListener('click', () => closeVisibilityCategoryModal(state));
  overlay.querySelector('[data-cancel]').addEventListener('click', () => closeVisibilityCategoryModal(state));
  overlay.addEventListener('click', (ev) => {
    if (ev.target === overlay) closeVisibilityCategoryModal(state);
  });

  const searchInput = overlay.querySelector('[data-category-search]');
  searchInput.addEventListener('input', () => {
    state.visibilityCategoryPicker.searchText = String(searchInput.value || '').trim();
    renderVisibilityCategoryModal(state);
  });

  overlay.querySelector('[data-apply]').addEventListener('click', () => {
    applyVisibilityCategorySelection(state);
  });

  state.ui.visibilityCategoryOverlay = overlay;
  state.ui.visibilityCategorySearch = searchInput;
  state.ui.visibilityCategoryList = overlay.querySelector('[data-category-picker-list]');
  state.ui.visibilityCategorySummary = overlay.querySelector('[data-category-selection-summary]');
  return overlay;
}

function buildExtractModal(state) {
  const overlay = div('deliverycleaner-modal deliverycleaner-extract-modal is-hidden');
  overlay.innerHTML = `
    <div class="deliverycleaner-modal__dialog deliverycleaner-extract-modal__dialog">
      <div class="deliverycleaner-modal__head">
        <div>
          <h3>속성값 추출</h3>
          <p>추출할 파라미터를 입력한 뒤 선택한 RVT의 객체 속성 정보를 추출합니다.</p>
        </div>
        <button type="button" class="deliverycleaner-modal__close" data-close>&times;</button>
      </div>
      <div class="deliverycleaner-modal__body">
        <div class="deliverycleaner-field-stack">
          <div class="deliverycleaner-note">
            New Schedule/Quantities에서 리스트업 가능한 실제 시공 객체 중심으로만 추출합니다.
            Centerline, 주석, 일반 선, 분석용 객체 등 스케줄 대상이 아닌 요소는 제외됩니다.
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
      toast('속성값 추출 대상 RVT가 없습니다. 먼저 RVT를 추가하거나 정리 결과 파일을 준비해주세요.', 'err', 3200);
      return;
    }

    if (!state.extractParameterNamesCsv.trim()) {
      toast('추출할 파라미터를 하나 이상 입력해주세요.', 'err', 3200);
      return;
    }
    if (!canUseSelectedDeliveryCleanerFiles(state, { allowSessionFallback: true })) {
      toast('추출할 RVT를 1개 이상 선택하세요.', 'err', 3200);
      updateActionState(state);
      return;
    }

    setPageBusy(state, true);
    beginDeliveryCleanerProgress('속성값 추출', '속성값 추출을 준비 중입니다.');
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

function createVisibilityRuleState() {
  return {
    combinationMode: 'Or',
    showImportedCategoriesInView: null,
    showImportsInFamilies: null,
    rules: Array.from({ length: 6 }, () => ({ enabled: false, parentCategoryNames: [], operatorName: 'Contains', subCategoryText: '' }))
  };
}

function createElementUpdateState() {
  return {
    enabled: false,
    combinationMode: 'And',
    applyToAllMatchingParameters: false,
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

function appendDroppedRvts(state, paths) {
  let added = 0;
  (Array.isArray(paths) ? paths : []).forEach((path) => {
    if (!path) return;
    const exists = state.filePaths.some((item) => String(item).toLowerCase() === String(path).toLowerCase());
    if (!exists) {
      state.filePaths.push(path);
      added += 1;
    }
    state.checked.add(path);
  });
  return added;
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
  renderExtractModalSummary(state);
  updateActionState(state);
}

function renderViewParameterRows(state) {
  const body = state.ui.viewParamBody;
  body.innerHTML = '';

  state.viewParameters.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.append(
      tdWithInput(row.parameterName, '파라미터 이름', (value) => { state.viewParameters[index].parameterName = value; }),
      tdWithInput(row.parameterValue, '값', (value) => { state.viewParameters[index].parameterValue = value; })
    );
    body.append(tr);
  });

  renderVisibilityRuleRows(state);
}

function renderVisibilityRuleRows(state) {
  const body = state.ui.visibilityRuleBody;
  if (!body) return;
  body.innerHTML = '';

  state.visibilityRules.rules.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.append(
      tdWithCategoryPicker(state, row, index),
      tdWithSelect(row.operatorName, TEXT_FILTER_OPERATORS, (value) => {
        state.visibilityRules.rules[index].operatorName = value;
        renderVisibilityRuleSummary(state);
      }),
      tdWithInput(row.subCategoryText, '예: End Cut', (value) => {
        state.visibilityRules.rules[index].subCategoryText = value;
        renderVisibilityRuleSummary(state);
      })
    );
    body.append(tr);
  });

  renderVisibilityRuleSummary(state);
}

function renderVisibilityRuleSummary(state) {
  if (!state.ui.visibilitySummary) return;
  const rules = state.visibilityRules.rules.filter((row) => (row.parentCategoryNames || []).length);
  const joiner = state.visibilityRules.combinationMode === 'And' ? ' AND ' : ' OR ';
  const importedLines = [
    formatVisibilityToggleSummary('Imported Categories 표시', state.visibilityRules.showImportedCategoriesInView),
    formatVisibilityToggleSummary('Imports in Families 표시', state.visibilityRules.showImportsInFamilies)
  ].filter(Boolean);
  const ruleText = rules.length
    ? rules.map((row) => {
      const categories = (row.parentCategoryNames || []).join(', ');
      if (row.operatorName === 'HasValue' || row.operatorName === 'HasNoValue') {
        return `${categories} / ${row.operatorName}`;
      }
      return `${categories} / ${row.operatorName} / ${row.subCategoryText || ''}`;
    }).join(joiner)
    : '커스텀 VV 서브카테고리 규칙이 없습니다.';
  if (rules.length || importedLines.length) {
    const lines = [`규칙 결합: ${state.visibilityRules.combinationMode === 'And' ? 'AND' : 'OR'}`];
    if (importedLines.length) lines.push(...importedLines);
    if (rules.length) lines.push(ruleText);
    state.ui.visibilitySummary.textContent = lines.join('\n');
    return;
  }

  state.ui.visibilitySummary.textContent = '카테고리 선택 버튼으로 주 카테고리를 고르고, 서브카테고리 텍스트 조건을 지정하면 커스텀 VV 숨김 규칙으로 반영됩니다.';
}

function renderElementUpdateRows(state) {
  const conditionBody = state.ui.conditionBody;
  conditionBody.innerHTML = '';
  state.elementParameterUpdate.conditions.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.append(
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
      tdWithInput(row.parameterName, '파라미터 이름', (value) => { state.elementParameterUpdate.assignments[index].parameterName = value; renderElementUpdateSummary(state); }),
      tdWithInput(row.value, '값', (value) => { state.elementParameterUpdate.assignments[index].value = value; renderElementUpdateSummary(state); })
    );
    assignmentBody.append(tr);
  });

  renderElementUpdateSummary(state);
}

function renderElementUpdateSummary(state) {
  const conds = state.elementParameterUpdate.conditions.filter((row) => row.parameterName.trim());
  const assigns = state.elementParameterUpdate.assignments.filter((row) => row.parameterName.trim());
  const joiner = state.elementParameterUpdate.combinationMode === 'Or' ? ' OR ' : ' AND ';
  const conditionText = conds.length
    ? conds.map((row) => (row.operatorName === 'HasValue' || row.operatorName === 'HasNoValue')
      ? `${row.parameterName} ${row.operatorName}`
      : `${row.parameterName} ${row.operatorName} ${row.value || ''}`).join(joiner)
    : '조건 미입력';
  const assignmentText = assigns.length
    ? assigns.map((row) => `${row.parameterName} = ${row.value || ''}`).join(' / ')
    : '입력값 미지정';
  const duplicateText = state.elementParameterUpdate.applyToAllMatchingParameters ? '중복 파라미터 전체 입력' : '중복 파라미터 하나만 입력';

  state.ui.elementSummary.textContent = (conds.length || assigns.length)
    ? `조건: ${conditionText}\n입력: ${assignmentText}\n중복 처리: ${duplicateText}`
    : '조건과 입력을 작성하면 자동으로 객체 파라미터 입력에 반영됩니다.';
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
    box.textContent = 'Purge 대기 중입니다. 정리 결과 또는 선택한 RVT를 기준으로 Purge를 실행하면 진행 상태가 여기에 표시됩니다.';
    return;
  }

  const fileText = snap.totalFiles ? `${snap.currentFileIndex || 0}/${snap.totalFiles}` : '-';
  const iterText = snap.totalIterations ? `${snap.currentIteration || 0}/${snap.totalIterations}` : '-';
  const chunks = [
    `Purge 상태: ${snap.stateName || '대기'}`,
    `파일 진행: ${fileText}`,
    `반복 진행: ${iterText}`
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

  if (!cleanedCount && !session.verificationCsvPath && !session.designOptionAuditCsvPath && !state.lastLogExportPath) {
    lines.push('아직 실행 결과가 없습니다.');
    lines.push('정리 시작, 정리 결과 검토, Design Option 검토, 로그 엑셀 저장 결과가 여기에 정리됩니다.');
    lines.push('');
    lines.push('정리 실행');
    lines.push('정리 완료 후에는 결과 파일 수와 성공/실패 파일 수가 표시되고, Design Option 검토 결과도 함께 확인할 수 있습니다.');
    lines.push('Design Option 검토 결과는 결과창에서 원하는 경로로 엑셀 저장할 수 있습니다.');
    lines.push('');
    lines.push('정리 결과 검토');
    lines.push('정리 결과 검토를 실행하면 파일별 검토 결과를 확인할 수 있고, 결과창에서 엑셀로 저장할 수 있습니다.');
    lines.push('속성값 추출은 별도 설정창에서 파라미터를 지정한 뒤 실행하며, 완료 후 결과창에서 원하는 경로로 엑셀 저장할 수 있습니다.');
    lines.push('');
    lines.push('로그 엑셀');
    lines.push('로그 엑셀 저장은 필요할 때만 수동으로 저장하며, 작업별 성공/실패가 요약되어 기록됩니다.');
  } else {
    lines.push(`정리 결과 파일: ${cleanedCount ? `${cleanedCount}개 생성` : '아직 없음'}`);
    if (Array.isArray(session.results) && session.results.length) {
      lines.push(`정리 결과: 성공 ${successCount} / 실패 ${failCount}`);
    }
    lines.push(`정리 결과 검토 파일: ${session.verificationCsvPath || '아직 없음'}`);
    lines.push(`Design Option 검토 파일: ${session.designOptionAuditCsvPath || '아직 없음'}`);
    lines.push('정리 완료 후에는 Design Option 검토와 객체 수 비교 결과를 결과창에서 바로 내보낼 수 있습니다.');
    lines.push('정리 결과 검토와 속성값 추출도 각각 완료 후 결과창에서 엑셀 저장이 가능합니다.');
    lines.push(`로그 엑셀: ${state.lastLogExportPath || '아직 저장 안 함'}`);
  }

  state.ui.resultSummary.textContent = lines.join('\n');
}

function isDeliveryCleanerConfigured(state) {
  return !!state.outputFolder;
}

function getDeliveryCleanerExtractionTargetCount(state) {
  if (state.filePaths.length) return getCheckedDeliveryCleanerFilePaths(state).length;
  if (Array.isArray(state.session?.cleanedOutputPaths)) return state.session.cleanedOutputPaths.length;
  return 0;
}

function getCheckedDeliveryCleanerFilePaths(state) {
  return state.filePaths.filter((path) => state.checked.has(path));
}

function canUseSelectedDeliveryCleanerFiles(state, options = {}) {
  const allowSessionFallback = !!options.allowSessionFallback;
  if (state.filePaths.length) return getCheckedDeliveryCleanerFilePaths(state).length > 0;
  if (!allowSessionFallback) return false;
  return Array.isArray(state.session?.cleanedOutputPaths) && state.session.cleanedOutputPaths.length > 0;
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

    const fileMatch = line.match(/([^\\\/\s]+\.rvt)\b/i);
    if (fileMatch) snapshot.currentFile = fileMatch[1];

    if (line.includes('정리 시작') || line.includes('Prepare start')) {
      snapshot.mode = 'clean';
      snapshot.currentTask = '정리 시작';
      return;
    }

    if (line.startsWith('[STEP] ')) {
      snapshot.mode = 'clean';
      snapshot.currentTask = line.substring(7).trim() || '정리 진행 중';
      return;
    }

    if (line.includes('정리 완료') || line.includes('Prepare completed')) {
      snapshot.mode = 'clean';
      snapshot.currentTask = '정리 완료';
      snapshot.completedCount += 1;
      return;
    }

    if (line.includes('오류') || line.includes('실패')) {
      snapshot.currentTask = '오류 확인 필요';
      return;
    }

    if (line.includes('검토 시작') || line.includes('Verify start')) {
      snapshot.mode = 'verify';
      snapshot.currentTask = '정리 결과 검토 시작';
      return;
    }

    if (line.includes('검토 완료') || line.includes('Verify completed')) {
      snapshot.mode = 'verify';
      snapshot.currentTask = '정리 결과 검토 완료';
      snapshot.completedCount += 1;
      return;
    }

    if (line.includes('속성값 추출 시작') || line.includes('Extraction start')) {
      snapshot.mode = 'extract';
      snapshot.currentTask = '속성값 추출 시작';
      return;
    }

    if (line.includes('속성값 추출 완료') || line.includes('Extraction completed')) {
      snapshot.mode = 'extract';
      snapshot.currentTask = '속성값 추출 완료';
      snapshot.completedCount += 1;
      return;
    }

    if (line.includes('Purge 시작') || line.includes('Purge file start')) {
      snapshot.mode = 'purge';
      snapshot.currentTask = 'Purge 시작';
      return;
    }

    if (line.includes('Purge 완료') || line.includes('Purge file completed')) {
      snapshot.mode = 'purge';
      snapshot.currentTask = 'Purge 완료';
      snapshot.completedCount += 1;
    }
  });

  return snapshot;
}

function renderStatusBox(state, hasFiles, isConfigured, hasSessionTargets, isPurging) {
  const progress = buildProgressSnapshot(state);
  const checkedFileCount = getCheckedDeliveryCleanerFilePaths(state).length;
  const lines = [];
  let tone = 'idle';

  if (state.busy || isPurging) {
    tone = 'active';
    lines.push('작업 진행 중');
    lines.push(progress.currentTask ? `현재 작업: ${progress.currentTask}` : '현재 작업: 진행 정보를 수집하는 중입니다.');
    lines.push(progress.currentFile ? `현재 파일: ${progress.currentFile}` : '현재 파일: 확인 중');
    if (progress.totalCount > 0) lines.push(`진행 파일 수: ${Math.min(progress.completedCount, progress.totalCount)}/${progress.totalCount}`);
  } else if (!hasFiles) {
    tone = 'idle';
    lines.push('대상 RVT가 아직 없습니다.');
    lines.push('RVT를 추가한 뒤 설정하기에서 기본 설정을 완료하면 정리 시작이 활성화됩니다.');
  } else if (!isConfigured) {
    tone = 'required';
    lines.push(`대상 RVT ${state.filePaths.length}개가 등록되었고 ${checkedFileCount}개가 선택되었습니다.`);
    lines.push('설정하기에서 결과 폴더를 지정하면 정리 시작 버튼이 활성화됩니다.');
    lines.push('설정 버튼이 강조되어 있으면 아직 필수 설정이 남아 있다는 뜻입니다.');
  } else {
    tone = 'ready';
    lines.push(`대상 RVT ${state.filePaths.length}개가 등록되었고 ${checkedFileCount}개가 실행 대상으로 준비되었습니다.`);
    lines.push('정리 시작을 눌러 정리 작업을 실행할 수 있습니다.');
    if (hasSessionTargets) lines.push(`최근 정리 결과 파일 ${state.session.cleanedOutputPaths.length}개가 세션에 연결되어 있습니다.`);
  }

  if (state.useFilter && state.filterProfile && isFilterConfigured(state.filterProfile) && !(state.busy || isPurging)) {
    lines.push('현재 뷰 필터 설정이 준비되어 있어 정리 시 함께 적용됩니다.');
  }

  state.ui.status.classList.remove('is-idle', 'is-required', 'is-ready', 'is-active');
  state.ui.status.classList.add(`is-${tone}`);
  state.ui.status.textContent = lines.join('\n');
}

function updateActionState(state) {
  const hasFiles = state.filePaths.length > 0;
  const checkedFileCount = getCheckedDeliveryCleanerFilePaths(state).length;
  const hasCheckedFiles = checkedFileCount > 0;
  const isConfigured = isDeliveryCleanerConfigured(state);
  const hasSessionTargets = Array.isArray(state.session?.cleanedOutputPaths) && state.session.cleanedOutputPaths.length > 0;
  const isPurging = !!state.purgeSnapshot?.isRunning;
  const canRun = !state.busy && hasCheckedFiles && isConfigured;
  const canReuseSessionTargets = !hasFiles && hasSessionTargets;
  const canVerifyLikeAction = !state.busy && (hasCheckedFiles || canReuseSessionTargets);

  state.ui.runBtn.disabled = !canRun;
  state.ui.verifyBtn.disabled = !canVerifyLikeAction;
  state.ui.extractBtn.disabled = !canVerifyLikeAction;
  state.ui.purgeBtn.disabled = state.busy || isPurging || !(hasCheckedFiles || canReuseSessionTargets);
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
  if (settings.visibilityRules) state.visibilityRules = normalizeVisibilityRules(settings.visibilityRules);
  if (Array.isArray(payload?.availableVisibilityCategories)) state.availableVisibilityCategories = normalizeVisibilityCategoryOptions(payload.availableVisibilityCategories);
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
  renderVisibilityCategoryModal(state);
  updateActionState(state);
}

function syncStateToInputs(state) {
  state.ui.outputFolderInput.value = state.outputFolder || '';
  state.ui.viewNameInput.value = state.target3DViewName || '';
  if (state.ui.extractInput) state.ui.extractInput.value = state.extractParameterNamesCsv || '';
  if (state.ui.visibilityCombinationMode) {
    state.ui.visibilityCombinationMode.value = state.visibilityRules.combinationMode || 'Or';
  }
  if (state.ui.visibilityImportedToggle) {
    state.ui.visibilityImportedToggle.value = formatVisibilityToggleValue(state.visibilityRules.showImportedCategoriesInView);
  }
  if (state.ui.visibilityImportsFamilyToggle) {
    state.ui.visibilityImportsFamilyToggle.value = formatVisibilityToggleValue(state.visibilityRules.showImportsInFamilies);
  }
  if (state.ui.visibilityCategorySearch) {
    state.ui.visibilityCategorySearch.value = state.visibilityCategoryPicker.searchText || '';
  }
  state.ui.combinationMode.value = state.elementParameterUpdate.combinationMode || 'And';
  if (state.ui.duplicateMode) {
    state.ui.duplicateMode.value = state.elementParameterUpdate.applyToAllMatchingParameters ? 'all' : 'single';
  }
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

function normalizeVisibilityRules(source) {
  const base = createVisibilityRuleState();
  base.combinationMode = source.combinationMode || 'Or';
  base.showImportedCategoriesInView = normalizeVisibilityToggleValue(source.showImportedCategoriesInView);
  base.showImportsInFamilies = normalizeVisibilityToggleValue(source.showImportsInFamilies);
  base.rules = base.rules.map((row, index) => {
    const raw = Array.isArray(source.rules) ? source.rules[index] || {} : {};
    return {
      ...row,
      ...raw,
      parentCategoryNames: splitVisibilityCategoryNames(raw.parentCategoryNames || raw.parentCategoryName || '')
    };
  });
  return base;
}

function normalizeElementUpdate(source) {
  const base = createElementUpdateState();
  base.enabled = !!source.enabled;
  base.combinationMode = source.combinationMode || 'And';
  base.applyToAllMatchingParameters = !!source.applyToAllMatchingParameters;
  base.conditions = base.conditions.map((row, index) => ({ ...row, ...(Array.isArray(source.conditions) ? source.conditions[index] || {} : {}) }));
  base.assignments = base.assignments.map((row, index) => ({ ...row, ...(Array.isArray(source.assignments) ? source.assignments[index] || {} : {}) }));
  return base;
}

function buildPayload(state) {
  const normalizedViewParameters = state.viewParameters.map((row) => ({
    ...row,
    enabled: !!String(row.parameterName || '').trim()
  }));
  const normalizedConditions = state.elementParameterUpdate.conditions.map((row) => ({
    ...row,
    enabled: !!String(row.parameterName || '').trim()
  }));
  const normalizedVisibilityRules = state.visibilityRules.rules.map((row) => ({
    ...row,
    parentCategoryNames: splitVisibilityCategoryNames(row.parentCategoryNames || ''),
    enabled: splitVisibilityCategoryNames(row.parentCategoryNames || '').length > 0 && ((row.operatorName === 'HasValue' || row.operatorName === 'HasNoValue') || !!String(row.subCategoryText || '').trim())
  }));
  const normalizedAssignments = state.elementParameterUpdate.assignments.map((row) => ({
    ...row,
    enabled: !!String(row.parameterName || '').trim()
  }));
  const hasConditions = normalizedConditions.some((row) => row.enabled);
  const hasAssignments = normalizedAssignments.some((row) => row.enabled);

  return {
    filePaths: [...state.filePaths],
    selectedFilePaths: state.filePaths.length ? getCheckedDeliveryCleanerFilePaths(state) : [],
    outputFolder: state.outputFolder,
    target3DViewName: state.target3DViewName,
    extractParameterNamesCsv: state.extractParameterNamesCsv,
    viewParameters: normalizedViewParameters,
    useFilter: state.useFilter,
    applyFilterInitially: state.applyFilterInitially,
    autoEnableFilterIfEmpty: state.autoEnableFilterIfEmpty,
    filterProfile: state.filterProfile ? { ...state.filterProfile } : null,
    visibilityRules: {
      combinationMode: state.visibilityRules.combinationMode,
      showImportedCategoriesInView: state.visibilityRules.showImportedCategoriesInView,
      showImportsInFamilies: state.visibilityRules.showImportsInFamilies,
      rules: normalizedVisibilityRules
    },
    elementParameterUpdate: {
      enabled: hasConditions && hasAssignments,
      combinationMode: state.elementParameterUpdate.combinationMode,
      applyToAllMatchingParameters: !!state.elementParameterUpdate.applyToAllMatchingParameters,
      conditions: normalizedConditions,
      assignments: normalizedAssignments
    }
  };
}

function countDeliveryCleanerExtractParameters(state) {
  return (state.extractParameterNamesCsv || '')
    .split(/[\,\n;\r]+/)
    .map((item) => item.trim())
    .filter(Boolean)
    .length;
}

function buildDeliveryCleanerCountNotes(items = [], emptyText = '객체수 비교 결과가 없습니다.') {
  if (!Array.isArray(items) || !items.length) return [emptyText];
  return items.map((item) => {
    const fileName = item?.fileName || getFileNameOnly(item?.outputPath || item?.sourcePath || '이름 없음');
    const beforeText = Number.isFinite(Number(item?.beforeCount)) ? `${Number(item.beforeCount)}개` : '-';
    const afterText = Number.isFinite(Number(item?.afterCount)) ? `${Number(item.afterCount)}개` : '-';
    const status = item?.status || '';
    const note = item?.note ? ` · ${item.note}` : '';
    if (beforeText !== '-' && afterText !== '-') {
      const delta = Number(item.afterCount) - Number(item.beforeCount);
      const deltaText = delta === 0 ? '변화 없음' : `${delta > 0 ? '+' : ''}${delta}개`;
      return `${fileName} · ${beforeText} -> ${afterText} (${deltaText})${status ? ` · ${status}` : ''}${note}`;
    }
    return `${fileName}${status ? ` · ${status}` : ''}${note}`;
  });
}

function handleDeliveryCleanerWorkbookExported(state, message, payload) {
  resetDeliveryCleanerProgress(state);
  setPageBusy(state, false);
  if (!payload?.path) return;
  window.setTimeout(() => {
    showExcelSavedDialog(message, payload.path, (path) => post('excel:open', { path }));
  }, 120);
}

async function promptDeliveryCleanerExcelExport(state, eventName) {
  const excelMode = await chooseExcelMode();
  if (!excelMode) return;
  setPageBusy(state, true);
  beginDeliveryCleanerProgress('엑셀 내보내기', '엑셀 저장을 준비 중입니다.');
  post(eventName, { excelMode: excelMode || 'fast' });
}

function showDeliveryCleanerRunDialog(state, payload = {}) {
  const summary = payload?.summary || {};
  const cleanedCount = summary.cleanedCount ?? (Array.isArray(state.session?.cleanedOutputPaths) ? state.session.cleanedOutputPaths.length : 0);
  const targetCount = state.filePaths.length || cleanedCount || (Array.isArray(state.session?.results) ? state.session.results.length : 0);
  const comparisons = Array.isArray(state.session?.cleanCountComparisons) ? state.session.cleanCountComparisons : [];
  const notes = [
    `정리 전 객체수 합계 ${summary.beforeObjectCount ?? 0}개 · 정리 후 객체수 합계 ${summary.afterObjectCount ?? 0}개`,
    ...buildDeliveryCleanerCountNotes(comparisons, '정리 객체수 비교 결과가 없습니다.'),
    'Design Option 검토와 정리 전후 객체수 비교 결과가 같은 엑셀에 함께 저장됩니다.'
  ];

  showCompletionSummaryDialog({
    title: 'RVT 정리 완료',
    message: '정리 작업이 완료되었습니다. 아래 요약을 확인하고 필요하면 결과 엑셀을 저장하세요.',
    summaryItems: [
      { label: '대상 파일', value: `${targetCount}개` },
      { label: '성공', value: `${summary.successCount ?? 0}개` },
      { label: '실패', value: `${summary.failCount ?? 0}개` },
      { label: '정리 결과', value: `${cleanedCount}개` }
    ],
    notes,
    exportLabel: 'Design Option + 객체수 비교 엑셀',
    showExport: payload?.canExportDesignOption === true,
    onExport: () => promptDeliveryCleanerExcelExport(state, 'deliverycleaner:export-designoption')
  });
}

function showDeliveryCleanerVerifyDialog(state, payload = {}) {
  const verifiedCount = Number(payload?.rowCount) || 0;
  const targetCount = (Array.isArray(state.session?.cleanedOutputPaths) ? state.session.cleanedOutputPaths.length : 0) || state.filePaths.length || 0;
  showCompletionSummaryDialog({
    title: '정리 결과 검토 완료',
    message: '정리 결과 검토가 완료되었습니다. 필요하면 검토 결과 엑셀을 저장하세요.',
    summaryItems: [
      { label: '검토 대상', value: `${targetCount}개 파일` },
      { label: '검토 행 수', value: `${verifiedCount}행` }
    ],
    notes: [
      '정리 결과 검토 엑셀은 저장 경로를 직접 지정해서 보관할 수 있습니다.',
      '검토 결과에서 CHECK 항목이 있으면 사용자가 직접 파일을 확인해 수정하면 됩니다.'
    ],
    exportLabel: '정리 결과 검토 엑셀',
    showExport: payload?.canExport === true,
    onExport: () => promptDeliveryCleanerExcelExport(state, 'deliverycleaner:export-verify')
  });
}

function showDeliveryCleanerExtractDialog(state, payload = {}) {
  const rowCount = Number(payload?.rowCount) || 0;
  const targetCount = getDeliveryCleanerExtractionTargetCount(state);
  const parameterCount = countDeliveryCleanerExtractParameters(state);
  showCompletionSummaryDialog({
    title: '속성값 추출 완료',
    message: '속성값 추출이 완료되었습니다. 필요하면 결과 엑셀을 저장하세요.',
    summaryItems: [
      { label: '대상 RVT', value: `${targetCount}개` },
      { label: '추출 파라미터', value: `${parameterCount}개` },
      { label: '추출 행 수', value: `${rowCount}행` }
    ],
    notes: [
      '속성값 추출은 스케줄 가능한 실제 시공 객체 기준으로 집계됩니다.',
      '결과 엑셀은 저장 경로를 직접 지정해서 보관할 수 있습니다.'
    ],
    exportLabel: '속성값 추출 엑셀',
    showExport: payload?.canExport === true,
    onExport: () => promptDeliveryCleanerExcelExport(state, 'deliverycleaner:export-extract')
  });
}

function showDeliveryCleanerPurgeDialog(state, payload = {}) {
  const comparisons = Array.isArray(state.session?.purgeCountComparisons) ? state.session.purgeCountComparisons : [];
  const targetCount = comparisons.length || (Array.isArray(state.session?.cleanedOutputPaths) ? state.session.cleanedOutputPaths.length : 0);
  const rowCount = Number(payload?.rowCount) || comparisons.filter((item) => Number.isFinite(Number(item?.beforeCount)) || Number.isFinite(Number(item?.afterCount))).length;
  showCompletionSummaryDialog({
    title: 'Purge 완료',
    message: 'Purge가 완료되었습니다. 파일별 객체수 비교를 확인하고 필요하면 엑셀을 저장하세요.',
    summaryItems: [
      { label: '대상 파일', value: `${targetCount}개` },
      { label: '비교 완료', value: `${rowCount}개` }
    ],
    notes: buildDeliveryCleanerCountNotes(comparisons, 'Purge 객체수 비교 결과가 없습니다.'),
    exportLabel: 'Purge 객체수 비교 엑셀',
    showExport: payload?.canExport === true,
    onExport: () => promptDeliveryCleanerExcelExport(state, 'deliverycleaner:export-purge')
  });
}

function handleDeliveryCleanerProgress(state, payload) {
  if (!state.acceptProgress || !state.busy) return;
  if (!payload) return;

  if (payload.phase || payload.current != null || payload.total != null) {
    const phase = normalizeDeliveryCleanerExcelPhase(payload?.phase);
    const total = Number(payload?.total) || 0;
    const current = Number(payload?.current) || 0;
    const percent = computeDeliveryCleanerExcelPercent(state, phase, current, total, payload?.phaseProgress, payload?.percent);
    const subtitle = buildDeliveryCleanerExcelSubtitle(phase, current, total);
    const detail = payload?.message || '';

    ProgressDialog.show(payload?.title || 'RVT 정리 (납품용)', subtitle || '작업을 처리하는 중입니다.');
    ProgressDialog.update(percent, subtitle || '작업을 처리하는 중입니다.', detail);

    if (phase === 'DONE' || phase === 'ERROR') {
      window.setTimeout(() => resetDeliveryCleanerProgress(state), 260);
    }
    return;
  }

  const title = payload?.title || 'RVT 정리 (납품용)';
  const message = payload?.message || '작업을 처리하는 중입니다.';
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
  state.acceptProgress = false;
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
    case 'EXCEL_INIT': return '엑셀 저장을 준비하는 중입니다.';
    case 'EXCEL_WRITE': return `엑셀 데이터를 작성하는 중입니다. (${current}/${Math.max(total, current || 1)})`;
    case 'EXCEL_SAVE': return '엑셀 파일을 저장하는 중입니다.';
    case 'AUTOFIT': return '열 너비와 스타일을 정리하는 중입니다.';
    case 'DONE': return '엑셀 저장이 완료되었습니다.';
    case 'ERROR': return '엑셀 저장 중 오류가 발생했습니다.';
    default: return '작업을 처리하는 중입니다.';
  }
}

function setPageBusy(state, on) {
  state.busy = !!on;
  state.acceptProgress = !!on;
  setBusy(on, on ? 'RVT 정리 작업을 처리하는 중입니다.' : '');
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
  state.ui.filterDocTitle.textContent = docTitle ? `현재 문서 필터: ${docTitle}` : '현재 문서의 필터를 선택하세요.';
  state.ui.filterDocList.innerHTML = '';

  if (!state.filterDocItems.length) {
    const empty = div('deliverycleaner-empty');
    empty.textContent = '현재 문서에서 추출 가능한 필터가 없습니다.';
    state.ui.filterDocList.append(empty);
  } else {
    state.filterDocItems.forEach((item) => {
      const row = document.createElement('button');
      row.type = 'button';
      row.className = 'deliverycleaner-doclist__item';
      row.textContent = item.name || '이름 없는 필터';
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
  const parameterCount = countDeliveryCleanerExtractParameters(state);

  const lines = [
    `대상 RVT: ${targetCount ? `${targetCount}개` : '아직 없음'}`,
    `추출 파라미터 수: ${parameterCount ? `${parameterCount}개` : '입력 필요'}`,
    '',
    '추출 대상은 스케줄로 리스트업 가능한 실제 시공 객체 중심으로 제한됩니다.',
    '추출이 완료되면 결과창에서 행 수를 확인하고 원하는 경로로 엑셀 저장할 수 있습니다.'
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

function splitVisibilityCategoryNames(value) {
  if (Array.isArray(value)) {
    return value
      .map((item) => String(item || '').trim())
      .filter(Boolean)
      .filter((item, index, arr) => arr.findIndex((x) => x.toLowerCase() === item.toLowerCase()) === index);
  }

  return String(value || '')
    .split(/[,\n\r;]+/)
    .map((item) => item.trim())
    .filter(Boolean)
      .filter((item, index, arr) => arr.findIndex((x) => x.toLowerCase() === item.toLowerCase()) === index);
}

function normalizeVisibilityCategoryOptions(items) {
  return (Array.isArray(items) ? items : [])
    .map((item) => String(item || '').trim())
    .filter(Boolean)
    .filter((item, index, arr) => arr.findIndex((x) => x.toLowerCase() === item.toLowerCase()) === index)
    .sort((a, b) => a.localeCompare(b, 'ko'));
}

function tdWithCategoryPicker(state, row, index) {
  const td = document.createElement('td');
  const wrap = div('deliverycleaner-category-picker-cell');
  const preview = document.createElement('textarea');
  preview.rows = 2;
  preview.readOnly = true;
  preview.className = 'deliverycleaner-cell-textarea deliverycleaner-cell-textarea--readonly';
  preview.placeholder = '카테고리 선택';
  preview.value = (row.parentCategoryNames || []).join(', ');

  const meta = div('deliverycleaner-category-picker-meta');
  const count = document.createElement('span');
  count.className = 'deliverycleaner-category-picker-count';
  count.textContent = row.parentCategoryNames?.length ? `${row.parentCategoryNames.length}개 선택` : '선택 없음';

  const actions = div('deliverycleaner-category-picker-actions');
  const pickBtn = actionButton('카테고리 선택', () => openVisibilityCategoryModal(state, index), 'secondary');
  const clearBtn = actionButton('초기화', () => {
    state.visibilityRules.rules[index].parentCategoryNames = [];
    renderVisibilityRuleRows(state);
  }, 'secondary');
  clearBtn.disabled = !(row.parentCategoryNames || []).length;

  actions.append(pickBtn, clearBtn);
  meta.append(count);
  wrap.append(preview, meta, actions);
  td.append(wrap);
  return td;
}

function openVisibilityCategoryModal(state, rowIndex) {
  state.visibilityCategoryPicker.rowIndex = rowIndex;
  state.visibilityCategoryPicker.searchText = '';
  state.visibilityCategoryPicker.selectedNames = splitVisibilityCategoryNames(state.visibilityRules.rules[rowIndex]?.parentCategoryNames || []);
  renderVisibilityCategoryModal(state);
  state.ui.visibilityCategoryOverlay?.classList.remove('is-hidden');
  state.ui.visibilityCategorySearch?.focus();
}

function closeVisibilityCategoryModal(state) {
  state.visibilityCategoryPicker.rowIndex = -1;
  state.visibilityCategoryPicker.searchText = '';
  state.visibilityCategoryPicker.selectedNames = [];
  state.ui.visibilityCategoryOverlay?.classList.add('is-hidden');
}

function renderVisibilityCategoryModal(state) {
  const listWrap = state.ui.visibilityCategoryList;
  const summary = state.ui.visibilityCategorySummary;
  if (!listWrap || !summary) return;

  if (state.ui.visibilityCategorySearch) {
    state.ui.visibilityCategorySearch.value = state.visibilityCategoryPicker.searchText || '';
  }

  const search = String(state.visibilityCategoryPicker.searchText || '').trim().toLowerCase();
  const selected = new Set(splitVisibilityCategoryNames(state.visibilityCategoryPicker.selectedNames || []));
  const options = normalizeVisibilityCategoryOptions(state.availableVisibilityCategories);
  const filtered = search
    ? options.filter((name) => name.toLowerCase().includes(search))
    : options;

  listWrap.innerHTML = '';
  if (!filtered.length) {
    const empty = div('deliverycleaner-note');
    empty.textContent = options.length
      ? '검색 조건에 맞는 카테고리가 없습니다.'
      : '현재 문서에서 선택할 수 있는 V/G 카테고리를 찾지 못했습니다.';
    listWrap.append(empty);
  } else {
    filtered.forEach((name) => {
      const label = document.createElement('label');
      label.className = 'deliverycleaner-category-option';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = selected.has(name);
      checkbox.addEventListener('change', () => {
        const next = new Set(splitVisibilityCategoryNames(state.visibilityCategoryPicker.selectedNames || []));
        if (checkbox.checked) next.add(name);
        else next.forEach((item) => {
          if (item.toLowerCase() === name.toLowerCase()) next.delete(item);
        });
        state.visibilityCategoryPicker.selectedNames = Array.from(next).sort((a, b) => a.localeCompare(b, 'ko'));
        renderVisibilityCategoryModal(state);
      });
      const text = document.createElement('span');
      text.textContent = name;
      label.append(checkbox, text);
      listWrap.append(label);
    });
  }

  const count = splitVisibilityCategoryNames(state.visibilityCategoryPicker.selectedNames || []).length;
  summary.textContent = count ? `${count}개 카테고리를 선택했습니다.` : '선택한 카테고리가 없습니다.';
}

function applyVisibilityCategorySelection(state) {
  const rowIndex = state.visibilityCategoryPicker.rowIndex;
  if (rowIndex < 0 || rowIndex >= state.visibilityRules.rules.length) {
    closeVisibilityCategoryModal(state);
    return;
  }

  state.visibilityRules.rules[rowIndex].parentCategoryNames = splitVisibilityCategoryNames(state.visibilityCategoryPicker.selectedNames || []);
  closeVisibilityCategoryModal(state);
  renderVisibilityRuleRows(state);
}

function parseVisibilityToggleValue(value) {
  if (value === 'show') return true;
  if (value === 'hide') return false;
  return null;
}

function normalizeVisibilityToggleValue(value) {
  if (typeof value === 'boolean') return value;
  if (value == null) return null;

  const text = String(value).trim().toLowerCase();
  if (!text) return null;
  if (['show', 'shown', 'checked', 'check', 'true', 'on', 'visible', 'yes'].includes(text)) return true;
  if (['hide', 'hidden', 'unchecked', 'uncheck', 'false', 'off', 'invisible', 'no'].includes(text)) return false;
  return null;
}

function formatVisibilityToggleValue(value) {
  if (value === true) return 'show';
  if (value === false) return 'hide';
  return '';
}

function formatVisibilityToggleSummary(label, value) {
  if (value === true) return `${label}: 체크`;
  if (value === false) return `${label}: 해제`;
  return '';
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

function tdWithTextarea(value, placeholder, onChange) {
  const td = document.createElement('td');
  const input = document.createElement('textarea');
  input.rows = 2;
  input.value = value || '';
  input.placeholder = placeholder || '';
  input.className = 'deliverycleaner-cell-textarea';
  input.addEventListener('input', () => onChange(input.value));
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


