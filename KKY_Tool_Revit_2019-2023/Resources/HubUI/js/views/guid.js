import { clear, div, toast, showExcelSavedDialog, chooseExcelMode, showCompletionSummaryDialog } from '../core/dom.js';
import { refreshUiAfterHostDialog } from '../core/hostDialog.js';
import { getLastExcelExportLocale } from '../core/dom.js';
import { attachRvtDropZone } from '../core/rvtDrop.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js?v=20260504a';
import { ProgressDialog } from '../core/progress.js';
import { post, onHost } from '../core/bridge.js';

const LS_RVTS = 'kky_guid_rvts';
const LS_SETTINGS = 'kky_guid_settings';

const HIDDEN_DETAIL_COLS = new Set(['RvtPath']);
const EXCEL_PHASE_WEIGHT = { EXCEL_INIT: 0.05, EXCEL_WRITE: 0.85, EXCEL_SAVE: 0.08, AUTOFIT: 0.02, DONE: 1, ERROR: 1 };

export function renderGuid(root) {
    const target = root || document.getElementById('view-root') || document.getElementById('app');
    clear(target);
    const top = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar'); if (top) top.classList.add('hub-topbar');

    const initialRvtList = loadRvtList();
    const initialSettings = loadGuidSettings();
    const state = {
        rvtList: initialRvtList,
        rvtChecked: new Set(initialRvtList),
        project: { columns: [], rows: [] },
        familyNav: { columns: [], rows: [] },
        familyDetail: { columns: [], rows: [] },
        activeProjectDoc: '',
        activeFamilyDoc: '',
        activeFamily: '',
        activeTab: 'project',
        busy: false,
        acceptRunProgress: false,
        acceptCleanupProgress: false,
        acceptExcelProgress: false,
        runId: '',
        hasRun: false,
        cleanupSourceExcelPath: '',
        includeFamily: !!initialSettings.includeFamily,
        includeAnnotation: !!initialSettings.includeAnnotation,
        closeAllWorksetsOnOpen: initialSettings.closeAllWorksetsOnOpen !== false,
        useSyncComment: !!initialSettings.useSyncComment,
        syncComment: initialSettings.syncComment || 'KKY Tools - 파라미터 GUID 정리',
        familyFilter: 'all',
        sharedParamStatus: null
    };
    const auditUi = {};
    const processButtons = {};
    const processSync = {};
    const processUi = {};
    let lastExcelPct = 0;

    const page = div('feature-shell familylink-page guid-page');

    // Header
    const header = div('feature-header guid-header');
    const headerLeft = div('feature-heading');
    headerLeft.innerHTML = `
      <span class="feature-kicker">파라미터 정리</span>
      <h2 class="feature-title">파라미터 GUID 검토 및 정리</h2>
      <p class="feature-sub">프로젝트/패밀리 파라미터 GUID를 검토하고 삭제용 엑셀 기준으로 정리합니다.</p>`;

    const sharedStatus = buildSharedParamStatus();
    header.append(headerLeft);
    page.append(header);

    const body = div('familylink-body guid-body');
    const topPanels = div('familylink-top-panels guid-top-panels');
    const leftColumn = div('guid-left-column');
    const rightColumn = div('guid-right-column');

    // GUID 검토
    const reviewSection = div('feature-results-panel guid-panel familylink-card segmentpms-extract');
    const reviewHeader = document.createElement('div');
    reviewHeader.className = 'feature-results-head familylink-results-head guid-rvt-head';
    const reviewTitle = document.createElement('div');
    reviewTitle.className = 'guid-title';
    reviewTitle.innerHTML = '<h3>GUID 검토</h3><p class="feature-note">검토 대상 RVT를 등록한 뒤 GUID 검토를 실행합니다. 목록이 비어 있으면 현재 활성 문서를 기준으로 검토합니다.</p>';
    const rvtMeta = div('familylink-results-meta');
    rvtMeta.textContent = '파일 0개';
    reviewHeader.append(reviewTitle, rvtMeta);

    const reviewActions = div('familylink-rvt-actions');
    const reviewManageActions = div('feature-actions guid-action-grid guid-action-grid--primary');
    const reviewRunActions = div('feature-actions guid-action-grid guid-action-grid--secondary');
    let btnRemove = null;
    let btnClear = null;
    const btnAdd = cardBtn('RVT 파일 추가', () => post('guid:add-files', { pick: 'files' }));
    btnAdd.classList.add('guid-btn-success');
    const btnSettings = cardBtn('설정', openSettingsDialog);
    btnSettings.classList.add('guid-btn-subtle');
    btnRemove = cardBtn('선택 제거', onRemoveSelected);
    btnRemove.disabled = true;
    btnClear = cardBtn('등록 목록 비우기', () => { state.rvtList = []; state.rvtChecked.clear(); persistRvts(); renderRvtList(); syncRvtActionState(); });
    const runBtn = cardBtn('검토 시작', onRun);
    const exportBtn = cardBtn('삭제용 엑셀 내보내기', onExport);
    const pickCleanupExcelBtn = cardBtn('삭제용 엑셀 불러오기', onPickCleanupExcel);
    const cleanupBtn = cardBtn('정리 시작', onCleanupStart);
    runBtn.classList.add('btn-primary');
    cleanupBtn.classList.add('guid-btn-warning');
    exportBtn.disabled = true;
    cleanupBtn.disabled = true;
    reviewManageActions.append(btnAdd, btnSettings, runBtn);
    reviewRunActions.append(btnRemove, btnClear);
    reviewActions.append(reviewManageActions, reviewRunActions);

    const reviewHint = div('rvt-drop-hint');
    reviewHint.textContent = 'RVT 파일 추가 버튼을 누르거나 탐색기에서 .rvt 파일을 드래그하면 아래 목록에 바로 등록됩니다.';
    const reviewSummary = div('feature-note guid-config-summary');
    auditUi.summary = reviewSummary;

    const rvtTableWrap = div('segmentpms-rvtlist guid-rvt-wrap rvt-drop-zone');
    const { table: rvtTable, tbody: rvtBody, master: rvtMaster } = createRvtTable();
    rvtTableWrap.append(rvtTable);
    const rvtSummary = div('segmentpms-summary');
    rvtSummary.textContent = '파일 0개';
    const reviewStatus = div('guid-audit-status');
    reviewStatus.append(sharedStatus);
    reviewSection.append(reviewHeader, reviewActions, reviewHint, reviewSummary, rvtTableWrap, rvtSummary, reviewStatus);
    attachRvtDropZone(rvtTableWrap, {
        onDropPaths: (paths) => {
            const added = appendDroppedRvts(paths);
            if (!added) {
                toast('이미 등록된 RVT입니다.', 'warn');
                return;
            }
            persistRvts();
            renderRvtList();
            syncRvtActionState();
            toast(`${added}개 RVT를 추가했습니다.`, 'ok');
        },
        onInvalid: () => toast('RVT 파일만 드래그해서 추가할 수 있습니다.', 'warn')
    });

    // 2단계 정리용 엑셀 반영
    const processSection = buildProcessSection(exportBtn, pickCleanupExcelBtn, cleanupBtn);
    leftColumn.append(reviewSection);
    rightColumn.append(processSection);
    topPanels.append(leftColumn, rightColumn);
    body.append(topPanels);

    // Result tabs
    const tabs = div('feature-results-panel familylink-results-panel guid-results feature-tabs');
    const tabHead = div('feature-results-head familylink-results-head guid-results-head');
    const tabHeadLeft = div('guid-results-tabs');
    const tabHeadRight = div('guid-results-filters');
    const tabBtns = div('pill-tabs');
    const btnTabProject = document.createElement('button'); btnTabProject.type = 'button'; btnTabProject.className = 'pill-tab is-active'; btnTabProject.innerHTML = `<span class="pill-label">RVT 검토결과</span><span class="pill-count">0</span>`;
    const btnTabFamily = document.createElement('button'); btnTabFamily.type = 'button'; btnTabFamily.className = 'pill-tab'; btnTabFamily.innerHTML = `<span class="pill-label">패밀리 검토 결과</span><span class="pill-count">0</span>`;
    tabBtns.append(btnTabProject, btnTabFamily);
    tabHeadLeft.append(tabBtns);
    const filterBar = buildFamilyFilter();
    filterBar.classList.add('is-hidden');
    tabHeadRight.append(filterBar);
    tabHead.append(tabHeadLeft, tabHeadRight);
    tabs.append(tabHead);

    const emptyState = div('guid-empty-state');
    emptyState.textContent = 'RVT를 등록한 뒤 검토를 시작해 주세요.';
    tabs.append(emptyState);

    const tabPanels = div('guid-tab-panels');
    const tabPanelProject = div('guid-tab-panel');
    const projWrap = div('guid-detail-wrap');
    const projNavPane = div('guid-detail-nav feature-results-panel guid-scroll-box');
    const projNavList = document.createElement('ul'); projNavList.className = 'guid-nav-list';
    projNavPane.append(projNavList);
    const projDetailPane = div('guid-detail-pane');
    const projTableWrap = div('guid-table-wrap guid-scroll-box');
    const projTable = document.createElement('table'); projTable.className = 'guid-table';
    const projHead = document.createElement('thead');
    const projBody = document.createElement('tbody');
    projTable.append(projHead, projBody);
    projTableWrap.append(projTable);
    projDetailPane.append(projTableWrap);
    projWrap.append(projNavPane, projDetailPane);
    tabPanelProject.append(projWrap);

    const tabPanelFamily = div('guid-tab-panel is-hidden');
    const detailWrap = div('guid-detail-wrap');
    const navPane = div('guid-detail-nav feature-results-panel guid-scroll-box');
    const navList = document.createElement('ul'); navList.className = 'guid-nav-list';
    navPane.append(navList);
    const detailPane = div('guid-detail-pane');
    const familyUserSection = div('guid-section');
    const familyUserHeader = div('guid-section-header');
    const familyUserTitle = document.createElement('div'); familyUserTitle.className = 'guid-section-title'; familyUserTitle.textContent = '사용자 파라미터';
    const familyUserCount = document.createElement('span'); familyUserCount.className = 'guid-section-count'; familyUserCount.textContent = '0';
    familyUserHeader.append(familyUserTitle, familyUserCount);
    const familyEmpty = div('guid-empty'); familyEmpty.textContent = '파라미터가 없습니다.';
    const detailTableWrap = div('guid-table-wrap guid-scroll-box');
    const detailTable = document.createElement('table'); detailTable.className = 'guid-table';
    const detailHead = document.createElement('thead');
    const detailBody = document.createElement('tbody');
    detailTable.append(detailHead, detailBody);
    detailTableWrap.append(detailTable);
    familyUserSection.append(familyUserHeader, familyEmpty, detailTableWrap);

    const builtSection = div('guid-section');
    const builtHeader = div('guid-section-header');
    const builtTitle = document.createElement('div'); builtTitle.className = 'guid-section-title'; builtTitle.textContent = '기본 제공 파라미터';
    const builtCount = document.createElement('span'); builtCount.className = 'guid-section-count'; builtCount.textContent = '0';
    const builtToggle = document.createElement('button'); builtToggle.type = 'button'; builtToggle.className = 'pill-tab'; builtToggle.textContent = '펼치기';
    builtHeader.append(builtTitle, builtCount, builtToggle);
    const builtTableWrap = div('guid-table-wrap guid-scroll-box');
    const builtTable = document.createElement('table'); builtTable.className = 'guid-table';
    const builtHead = document.createElement('thead');
    const builtBody = document.createElement('tbody');
    builtTable.append(builtHead, builtBody);
    builtTableWrap.append(builtTable);
    builtSection.append(builtHeader, builtTableWrap);
    detailPane.append(familyUserSection, builtSection);

    builtToggle.onclick = () => {
        builtSection.classList.toggle('is-open');
        builtToggle.textContent = builtSection.classList.contains('is-open') ? '접기' : '펼치기';
        paintFamily();
    };
    detailWrap.append(navPane, detailPane);
    tabPanelFamily.append(detailWrap);

    tabPanels.append(tabPanelProject, tabPanelFamily);
    tabs.append(tabPanels);
    tabs.classList.add('guid-results-hidden');
    body.append(tabs);

    page.append(body);
    target.append(page);

    renderRvtList();
    syncTabState();
    syncRvtActionState();
    syncAnnotationToggle();
    syncAuditConfigSummary();
    syncResultState();
    requestSharedParamStatus();

    // Host events
    onHost('guid:files', ({ paths }) => {
        const added = appendDroppedRvts(paths);
        if (added) persistRvts();
        refreshUiAfterHostDialog(() => {
            renderRvtList();
            syncRvtActionState();
        });
    });

    onHost('guid:progress', (payload) => {
        if (payload && payload.phase) {
            handleExcelProgress(payload);
        } else {
            handleRunProgress(payload);
        }
    });
    onHost('guid:cleanup-progress', handleCleanupProgress);

    onHost('guid:done', (payload) => {
        state.acceptRunProgress = false;
        ProgressDialog.hide();
        setBusy(false);
        lastExcelPct = 0;
        const proj = payload?.project || {};
        const famNav = payload?.family || payload?.familyIndex || {};
        state.runId = payload?.runId || '';
        state.hasRun = true;
        state.includeFamily = !!payload?.includeFamily;
        state.includeAnnotation = !!payload?.includeAnnotation;
        state.project = {
            columns: Array.isArray(proj.columns) ? proj.columns : [],
            rows: Array.isArray(proj.rows) ? proj.rows : []
        };
        state.familyNav = {
            columns: Array.isArray(famNav.columns) ? famNav.columns : [],
            rows: Array.isArray(famNav.rows) ? famNav.rows : []
        };
        state.familyDetail = { columns: [], rows: [] };
        state.activeProjectDoc = '';
        state.activeFamilyDoc = '';
        state.activeFamily = '';
        state.familyFilter = 'all';
        syncAnnotationToggle();
        syncAuditConfigSummary();
        exportBtn.disabled = !hasRowsForExport();
        updateTabCounts();
        paintProject();
        paintFamily();
        syncTabState();
        syncResultState();
        requestAnimationFrame(() => showGuidCompletionDialog(payload));
        toast('검토 완료', 'ok');
    });

    onHost('guid:cleanup-done', (payload) => {
        state.acceptCleanupProgress = false;
        ProgressDialog.hide();
        setBusy(false);
        lastExcelPct = 0;
        state.cleanupSourceExcelPath = String(payload?.sourceExcelPath || state.cleanupSourceExcelPath || '').trim();
        syncCleanupPanel();
        if (payload?.message) toast(payload.message, payload?.ok ? 'ok' : 'warn');
        requestAnimationFrame(() => showGuidCleanupDialog(payload));
    });
    onHost('guid:cleanup-cancelled', () => {
        refreshUiAfterHostDialog(() => {
            state.acceptCleanupProgress = false;
            ProgressDialog.hide();
            setBusy(false);
        });
    });
    onHost('guid:cleanup-excel-picked', ({ path }) => {
        refreshUiAfterHostDialog(() => {
            state.cleanupSourceExcelPath = String(path || '').trim();
            syncCleanupPanel();
        });
    });

    onHost('guid:family-detail', (payload) => {
        if (payload?.runId && state.runId && payload.runId !== state.runId) return;
        state.familyDetail = {
            columns: Array.isArray(payload?.columns) ? payload.columns : [],
            rows: Array.isArray(payload?.rows) ? payload.rows : []
        };
        state.activeFamilyDoc = payload?.rvtPath || state.activeFamilyDoc || '';
        state.activeFamily = payload?.familyName || state.activeFamily || '';
        updateTabCounts();
        paintFamily();
        syncTabState();
        syncResultState();
    });

    onHost('guid:warn', ({ message }) => {
        if (message) toast(message, 'warn');
    });

    onHost('guid:exported', ({ path }) => {
        state.acceptExcelProgress = false;
        ProgressDialog.hide();
        setBusy(false);
        lastExcelPct = 0;
        state.cleanupSourceExcelPath = String(path || state.cleanupSourceExcelPath || '').trim();
        syncCleanupPanel();
        if (path) {
            showExcelSavedDialog('파라미터 GUID 삭제용 엑셀을 저장했습니다.', path, (p) => post('excel:open', { path: p }));
        } else {
            toast('파라미터 GUID 삭제용 엑셀 내보내기를 완료했습니다.', 'ok');
        }
    });

    const handleError = ({ message }) => {
        state.acceptRunProgress = false;
        state.acceptCleanupProgress = false;
        state.acceptExcelProgress = false;
        ProgressDialog.hide();
        setBusy(false);
        lastExcelPct = 0;
        if (message) toast(message, 'err');
    };
    onHost('guid:error', handleError);
    onHost('revit:error', handleError);
    onHost('host:error', handleError);
    onHost('sharedparam:status', (payload) => {
        state.sharedParamStatus = payload || {};
        updateSharedParamStatus(sharedStatus, state.sharedParamStatus);
    });

    // UI handlers
    btnTabProject.onclick = () => { state.activeTab = 'project'; syncTabState(); };
    btnTabFamily.onclick = () => {
        if (!state.includeFamily) return;
        state.activeTab = 'family';
        syncTabState();
    };

    function onRun() {
        if (state.busy) return;
        if (!canRunWithSharedParam()) return;
        state.rvtList = dedupPaths(state.rvtList);
        const targets = dedupPaths(state.rvtList.filter(p => state.rvtChecked.has(p)));
        if (state.rvtList.length > 0 && targets.length === 0) {
            toast('선택된 RVT가 없습니다.', 'warn');
            return;
        }

        const includeFamily = !!state.includeFamily;
        const payload = {
            mode: includeFamily ? 2 : 1,
            rvtPaths: state.rvtList.length === 0 ? [] : targets,
            includeFamily,
            includeAnnotation: includeFamily ? !!state.includeAnnotation : false
        };
        persistRvts();
        persistGuidSettings();

        state.familyNav = { columns: [], rows: [] };
        state.familyDetail = { columns: [], rows: [] };
        state.project = { columns: [], rows: [] };
        state.activeProjectDoc = '';
        state.activeFamilyDoc = '';
        state.activeFamily = '';
        setBusy(true);
        state.acceptRunProgress = true;
        state.acceptCleanupProgress = false;
        state.acceptExcelProgress = false;
        state.runId = '';
        state.activeTab = 'project';
        ProgressDialog.show('파라미터 GUID 검토', '검토를 준비하는 중입니다.');
        post('guid:run', payload);
    }

    function requestSharedParamStatus() {
        post('sharedparam:status', { source: 'guid' });
    }

    function canRunWithSharedParam() {
        const status = state.sharedParamStatus || {};
        if (!status.status || status.status === 'ok') return true;
        const msg = status.warning || '공유파라미터 파일 상태가 올바르지 않습니다.';
        toast(msg, 'err');
        return false;
    }

    function onExport() {
        if (state.busy) return;
        if (!hasRowsForExport()) { toast('저장할 결과가 없습니다.', 'warn'); return; }
        let which = 'project';
        if (state.includeFamily) which = 'all';
        chooseExcelMode((mode) => {
            const excelMode = mode || 'fast';
            lastExcelPct = 0;
            setBusy(true);
            state.acceptExcelProgress = true;
            state.acceptRunProgress = false;
            state.acceptCleanupProgress = false;
            ProgressDialog.show('삭제용 엑셀 내보내기', '엑셀 파일을 만드는 중입니다.');
            post('guid:export', { which, excelMode, locale: getLastExcelExportLocale() });
        });
    }

    function onPickCleanupExcel() {
        if (state.busy) return;
        post('guid:pick-cleanup-excel', {});
    }

    function onCleanupStart() {
        if (state.busy) return;
        if (!state.cleanupSourceExcelPath) {
            toast('먼저 삭제용 엑셀을 불러와 주세요.', 'warn');
            return;
        }
        state.closeAllWorksetsOnOpen = true;
        persistGuidSettings();
        lastExcelPct = 0;
        setBusy(true);
        state.acceptRunProgress = false;
        state.acceptExcelProgress = false;
        state.acceptCleanupProgress = true;
        ProgressDialog.show('파라미터 GUID 정리', '삭제용 엑셀 기준으로 정리를 준비하는 중입니다.');
        post('guid:cleanup', {
            excelPath: state.cleanupSourceExcelPath,
            closeAllWorksetsOnOpen: true,
            useSyncComment: !!state.useSyncComment,
            syncComment: state.syncComment || ''
        });
    }

    function openSettingsDialog() {
        const existing = document.querySelector('.guid-settings-backdrop');
        if (existing) existing.remove();

        let includeFamily = !!state.includeFamily;
        let includeAnnotation = !!state.includeAnnotation;

        const backdrop = document.createElement('div');
        backdrop.className = 'guid-settings-backdrop';

        const dialog = document.createElement('section');
        dialog.className = 'guid-settings-dialog';
        dialog.setAttribute('role', 'dialog');
        dialog.setAttribute('aria-modal', 'true');
        dialog.setAttribute('aria-label', 'GUID 검토 설정');

        const title = document.createElement('div');
        title.className = 'guid-settings-dialog__title';
        title.textContent = 'GUID 검토 설정';

        const desc = document.createElement('div');
        desc.className = 'guid-settings-dialog__desc';
        desc.textContent = '검토 시 포함할 패밀리 범위를 선택한 뒤 적용해 주세요.';

        const options = div('guid-setting-checks');
        const familyLabel = document.createElement('label');
        familyLabel.className = 'guid-checkbox-row';
        const familyCk = document.createElement('input');
        familyCk.type = 'checkbox';
        familyCk.checked = includeFamily;
        const familyText = document.createElement('span');
        familyText.textContent = '패밀리 포함';
        familyLabel.append(familyCk, familyText);

        const annotationLabel = document.createElement('label');
        annotationLabel.className = 'guid-checkbox-row';
        const annotationCk = document.createElement('input');
        annotationCk.type = 'checkbox';
        annotationCk.checked = includeAnnotation;
        const annotationText = document.createElement('span');
        annotationText.textContent = '주석 패밀리 포함';
        annotationLabel.append(annotationCk, annotationText);
        options.append(familyLabel, annotationLabel);

        const syncDraft = () => {
            includeFamily = !!familyCk.checked;
            includeAnnotation = !!annotationCk.checked;
            annotationCk.disabled = !includeFamily;
            annotationLabel.classList.toggle('is-disabled', !includeFamily);
            if (!includeFamily) {
                includeAnnotation = false;
                annotationCk.checked = false;
            }
        };

        familyCk.addEventListener('change', syncDraft);
        annotationCk.addEventListener('change', syncDraft);
        syncDraft();

        const actions = document.createElement('div');
        actions.className = 'guid-settings-dialog__actions';
        const cancelBtn = cardBtn('취소', closeSettingsDialog);
        const applyBtn = cardBtn('적용', () => {
            state.includeFamily = includeFamily;
            state.includeAnnotation = includeAnnotation;
            persistGuidSettings();
            syncAnnotationToggle();
            syncAuditConfigSummary();
            syncTabState();
            closeSettingsDialog();
        });
        applyBtn.classList.add('btn-primary');
        actions.append(cancelBtn, applyBtn);

        dialog.append(title, desc, options, actions);
        backdrop.append(dialog);
        backdrop.addEventListener('click', (event) => {
            if (event.target === backdrop) closeSettingsDialog();
        });
        document.body.append(backdrop);
    }

    function closeSettingsDialog() {
        const existing = document.querySelector('.guid-settings-backdrop');
        if (existing) existing.remove();
    }

    function buildProcessSection(exportBtn, pickCleanupExcelBtn, cleanupBtn) {
        const panel = div('feature-results-panel guid-panel familylink-card guid-process-panel');
        const head = div('feature-results-head familylink-results-head guid-process-head');
        const title = document.createElement('div');
        title.className = 'guid-title';
        title.innerHTML = '<h3>파라미터 정리용 엑셀</h3><p class="feature-note">검토 결과를 삭제용 엑셀로 내보낸 뒤, 삭제 표시한 엑셀을 다시 불러와 정리를 적용합니다.</p>';
        const meta = div('familylink-results-meta');
        meta.textContent = '삭제용 엑셀 선택 필요';
        processUi.meta = meta;
        head.append(title, meta);

        const actions = div('familylink-rvt-actions guid-process-actions');
        const actionGrid = div('feature-actions guid-action-grid guid-action-grid--process');
        actionGrid.append(exportBtn, pickCleanupExcelBtn, cleanupBtn);
        actions.append(actionGrid);

        processButtons.export = exportBtn;
        processButtons.pickCleanupExcel = pickCleanupExcelBtn;
        processButtons.cleanup = cleanupBtn;

        const excelDrop = div('feature-row__summary guid-apply-drop');
        const excelLead = document.createElement('strong');
        excelLead.textContent = "선택한 삭제용 엑셀의 '삭제여부' 열에 '삭제'가 입력된 행만 정리 대상으로 적용합니다.";
        const excelPath = div('feature-note');
        excelPath.textContent = '아직 선택한 삭제용 엑셀이 없습니다.';
        processUi.path = excelPath;
        excelDrop.append(excelLead, excelPath);

        const ruleWrap = div('guid-rule-box');
        const processTitle = document.createElement('div');
        processTitle.className = 'guid-setting-card__title';
        processTitle.textContent = '간단한 진행 순서';
        const processList = div('guid-setting-process');
        processList.append(
            buildProcessRow('1', '검토 시작으로 GUID 결과를 확인합니다.'),
            buildProcessRow('2', '삭제용 엑셀 내보내기로 작업용 엑셀을 저장합니다.'),
            buildProcessRow('3', "엑셀의 '삭제여부' 열에 삭제라고 입력한 뒤 다시 불러옵니다."),
            buildProcessRow('4', '정리 시작을 누르면 센트럴 파일도 로컬로, 모든 웍셋을 닫고 적용합니다.')
        );
        const rule = div('guid-setting-code');
        rule.innerHTML = '<span>입력 예시</span><strong>삭제여부 = 삭제</strong>';
        const ruleNote = div('feature-note');
        ruleNote.textContent = '삭제하지 않을 행은 비워 두고, 마지막 숨김 키 행은 수정하지 않습니다.';
        ruleWrap.append(processTitle, processList, rule, ruleNote);

        const syncBlock = div('guid-process-sync');
        const syncLabel = document.createElement('label');
        syncLabel.className = 'guid-checkbox-row';
        const syncCk = document.createElement('input');
        syncCk.type = 'checkbox';
        syncCk.checked = !!state.useSyncComment;
        const syncText = document.createElement('span');
        syncText.textContent = '동기화 시 코멘트 작성';
        syncLabel.append(syncCk, syncText);

        const syncInput = document.createElement('input');
        syncInput.type = 'text';
        syncInput.className = 'guid-text-input';
        syncInput.placeholder = '예) KKY Tools - 파라미터 GUID 정리';
        syncInput.value = state.syncComment || '';

        const syncState = () => {
            state.useSyncComment = !!syncCk.checked;
            syncInput.disabled = !state.useSyncComment;
            syncLabel.classList.toggle('is-disabled', !state.useSyncComment);
            persistGuidSettings();
        };

        syncCk.onchange = syncState;
        syncInput.addEventListener('input', () => {
            state.syncComment = syncInput.value || '';
            persistGuidSettings();
        });
        syncState();

        processSync.checkbox = syncCk;
        processSync.input = syncInput;
        processSync.sync = () => {
            syncCk.checked = !!state.useSyncComment;
            syncInput.value = state.syncComment || '';
            syncState();
        };

        const syncTitle = document.createElement('div');
        syncTitle.className = 'guid-setting-card__title';
        syncTitle.textContent = '동기화 코멘트';
        syncBlock.append(syncLabel, syncInput);

        panel.append(head, actions, excelDrop, ruleWrap, syncTitle, syncBlock);
        return panel;
    }

    function buildSharedParamStatus() {
        const wrap = div('sharedparam-status');
        wrap.innerHTML = `
          <div class="sharedparam-status__head">
            <span class="sharedparam-status__title">공유파라미터 상태</span>
            <span class="sharedparam-status__badge chip">조회 중</span>
          </div>
          <div class="sharedparam-status__body">
            <div class="sharedparam-status__row">
              <span class="sharedparam-status__label">경로</span>
              <span class="sharedparam-status__value" data-sp-path>조회 중</span>
            </div>
            <div class="sharedparam-status__row">
              <span class="sharedparam-status__label">파일 존재</span>
              <span class="sharedparam-status__value" data-sp-exists>—</span>
            </div>
            <div class="sharedparam-status__row">
              <span class="sharedparam-status__label">파일 열기</span>
              <span class="sharedparam-status__value" data-sp-open>—</span>
            </div>
          </div>
          <div class="sharedparam-status__hint" data-sp-hint></div>`;
        return wrap;
    }

    function updateSharedParamStatus(container, payload) {
        if (!container) return;
        const badge = container.querySelector('.sharedparam-status__badge');
        const pathEl = container.querySelector('[data-sp-path]');
        const existsEl = container.querySelector('[data-sp-exists]');
        const openEl = container.querySelector('[data-sp-open]');
        const hintEl = container.querySelector('[data-sp-hint]');

        const status = payload?.status || 'unknown';
        const label = payload?.statusLabel || '상태를 확인할 수 없습니다.';
        const path = payload?.path || '미설정';
        const exists = payload?.existsOnDisk ? '파일 있음' : '파일 없음';
        const canOpen = payload?.canOpen ? '열기 가능' : '열기 실패';
        const warning = payload?.warning || payload?.errorMessage || '';

        if (pathEl) pathEl.textContent = path;
        if (existsEl) existsEl.textContent = payload?.isSet ? exists : '미설정';
        if (openEl) openEl.textContent = payload?.isSet ? canOpen : '미설정';
        if (badge) {
            badge.textContent = label;
            badge.classList.remove('sharedparam-badge--ok', 'sharedparam-badge--warn', 'sharedparam-badge--error');
            if (status === 'ok') badge.classList.add('sharedparam-badge--ok');
            else if (status === 'unset' || status === 'missing') badge.classList.add('sharedparam-badge--warn');
            else badge.classList.add('sharedparam-badge--error');
        }
        if (hintEl) {
            hintEl.textContent = warning ? `이 상태에서는 검토가 실패할 수 있습니다. ${warning}` : '';
            hintEl.style.display = warning ? 'block' : 'none';
        }
    }

    function buildFamilyFilter() {
        const bar = div('guid-filter-bar');
        const btnAll = document.createElement('button'); btnAll.type = 'button'; btnAll.className = 'pill-tab is-active'; btnAll.textContent = '전체';
        const btnShared = document.createElement('button'); btnShared.type = 'button'; btnShared.className = 'pill-tab'; btnShared.textContent = '공유';
        const btnFamily = document.createElement('button'); btnFamily.type = 'button'; btnFamily.className = 'pill-tab'; btnFamily.textContent = '패밀리';
        const sync = () => {
            btnAll.classList.toggle('is-active', state.familyFilter === 'all');
            btnShared.classList.toggle('is-active', state.familyFilter === 'shared');
            btnFamily.classList.toggle('is-active', state.familyFilter === 'family');
        };
        btnAll.onclick = () => { state.familyFilter = 'all'; sync(); paintFamily(); };
        btnShared.onclick = () => { state.familyFilter = 'shared'; sync(); paintFamily(); };
        btnFamily.onclick = () => { state.familyFilter = 'family'; sync(); paintFamily(); };
        bar.append(btnAll, btnShared, btnFamily);
        bar.sync = sync;
        return bar;
    }

    function renderRvtList() {
        state.rvtList = dedupPaths(state.rvtList);
        state.rvtChecked = new Set(state.rvtList.filter(p => state.rvtChecked.has(p)));
        const allChecked = state.rvtList.length > 0 && state.rvtList.every(p => state.rvtChecked.has(p));
        rvtMaster.checked = allChecked;
        rvtMaster.indeterminate = state.rvtList.length > 0 && !allChecked && state.rvtChecked.size > 0;
        rvtMaster.disabled = state.rvtList.length === 0;
        rvtMaster.onchange = () => {
            if (rvtMaster.checked) state.rvtChecked = new Set(state.rvtList);
            else state.rvtChecked.clear();
            persistRvts();
            renderRvtList();
        };
        const rows = state.rvtList.map((p, i) => ({
            checked: state.rvtChecked.has(p),
            index: i + 1,
            name: getRvtName(p),
            path: p,
            title: p,
            onToggle: (checked) => {
                if (checked) state.rvtChecked.add(p); else state.rvtChecked.delete(p);
                persistRvts();
                renderRvtList();
            }
        }));
        renderRvtRows(rvtBody, rows);
        rvtSummary.textContent = state.rvtList.length ? `파일 ${state.rvtList.length}개` : '파일 0개';
        syncRvtActionState();
    }

    function paintProject() {
        projNavList.innerHTML = '';
        if (!state.project.rows.length) {
            const empty = document.createElement('li');
            empty.className = 'guid-nav-empty';
            empty.textContent = '결과가 없습니다.';
            projNavList.append(empty);
            projHead.innerHTML = '';
            projBody.innerHTML = '';
            return;
        }
        const idxPath = projectCol('RvtPath');
        const idxName = projectCol('RvtName');
        const docMap = new Map();
        state.project.rows.forEach(row => {
            const path = idxPath >= 0 ? safe(row[idxPath]) : '';
            const name = idxName >= 0 ? safe(row[idxName]) : (path || '(Doc)');
            const key = path || name;
            if (!docMap.has(key)) docMap.set(key, name);
        });
        const docs = Array.from(docMap.entries()).sort((a, b) => a[1].localeCompare(b[1]));
        if (!state.activeProjectDoc && docs.length) state.activeProjectDoc = docs[0][0];
        docs.forEach(([key, name]) => {
            const li = document.createElement('li');
            const btn = document.createElement('button'); btn.type = 'button'; btn.className = 'nav-fam-item'; btn.textContent = name;
            btn.onclick = () => { state.activeProjectDoc = key; paintProject(); };
            if (state.activeProjectDoc === key) btn.classList.add('is-active');
            li.append(btn);
            projNavList.append(li);
        });
        buildHead(projHead, state.project.columns, new Set(['RvtPath']));
        paintVirtualRows(projBody, state.project.columns, filteredProjectRows(), new Set(['RvtPath']));
    }

    function paintFamily() {
        navList.innerHTML = '';
        if (!state.includeFamily) {
            navList.innerHTML = '<li class="guid-nav-empty">패밀리 검토 결과 추가 검토를 선택한 뒤 실행해 주세요.</li>';
            detailHead.innerHTML = '';
            detailBody.innerHTML = '';
            builtHead.innerHTML = '';
            builtBody.innerHTML = '';
            familyEmpty.style.display = 'none';
            detailTableWrap.style.display = 'none';
            builtTableWrap.style.display = 'none';
            if (filterBar && typeof filterBar.sync === 'function') filterBar.sync();
            return;
        }
        if (!state.familyNav.rows.length) {
            const empty = document.createElement('li');
            empty.className = 'guid-nav-empty';
            empty.textContent = '패밀리 결과가 없습니다.';
            navList.append(empty);
            detailHead.innerHTML = '';
            detailBody.innerHTML = '';
            builtHead.innerHTML = '';
            builtBody.innerHTML = '';
            familyEmpty.style.display = 'none';
            detailTableWrap.style.display = 'none';
            builtTableWrap.style.display = 'none';
            if (filterBar && typeof filterBar.sync === 'function') filterBar.sync();
            return;
        }
        if (filterBar && typeof filterBar.sync === 'function') filterBar.sync();
        const idxPath = indexCol(state.familyNav, 'RvtPath');
        const idxName = indexCol(state.familyNav, 'RvtName');
        const idxFam = indexCol(state.familyNav, 'FamilyName');
        const map = new Map();
        state.familyNav.rows.forEach(row => {
            const path = idxPath >= 0 ? safe(row[idxPath]) : '';
            const rname = idxName >= 0 ? safe(row[idxName]) : (path || '(Doc)');
            const fam = idxFam >= 0 ? safe(row[idxFam]) : '';
            const key = path || rname;
            if (!map.has(key)) map.set(key, { name: rname, families: new Set() });
            if (fam) map.get(key).families.add(fam);
        });
        const entries = Array.from(map.entries()).sort((a, b) => a[1].name.localeCompare(b[1].name));
        if (!state.activeFamilyDoc && entries.length) state.activeFamilyDoc = entries[0][0];
        entries.forEach(([key, info]) => {
            const docItem = document.createElement('li');
            docItem.className = 'guid-nav-doc';
            const docTitle = document.createElement('div');
            docTitle.className = 'nav-doc-title';
            docTitle.textContent = info.name;
            docItem.append(docTitle);

            const famList = document.createElement('ul');
            famList.className = 'guid-nav-fams';

            Array.from(info.families).sort((a, b) => a.localeCompare(b)).forEach(f => {
                const li = document.createElement('li');
                const btn = document.createElement('button'); btn.type = 'button'; btn.className = 'nav-fam-item'; btn.textContent = f;
                btn.setAttribute('aria-label', f);
                btn.onclick = () => { onRequestFamilyDetail(key, f); };
                if (state.activeFamilyDoc === key && state.activeFamily === f) btn.classList.add('is-active');
                li.append(btn);
                famList.append(li);
            });

            docItem.append(famList);
            navList.append(docItem);
        });
        const { userRows, builtRows } = splitFamilyRows();
        familyUserCount.textContent = String(userRows.length || 0);
        builtCount.textContent = String(builtRows.length || 0);
        familyEmpty.style.display = userRows.length ? 'none' : 'block';
        detailTableWrap.style.display = userRows.length ? 'block' : 'none';
        buildHead(detailHead, state.familyDetail.columns, HIDDEN_DETAIL_COLS);
        paintVirtualRows(detailBody, state.familyDetail.columns, userRows, HIDDEN_DETAIL_COLS);

        const builtOpen = builtSection.classList.contains('is-open');
        builtTableWrap.style.display = builtOpen ? 'block' : 'none';
        if (builtOpen) {
            buildHead(builtHead, state.familyDetail.columns, HIDDEN_DETAIL_COLS);
            paintVirtualRows(builtBody, state.familyDetail.columns, builtRows, HIDDEN_DETAIL_COLS);
        } else {
            builtHead.innerHTML = '';
            builtBody.innerHTML = '';
        }
    }

    function onRequestFamilyDetail(rvtPath, familyName) {
        if (!state.runId) {
            toast('먼저 검토를 실행해 주세요.', 'warn');
            return;
        }
        state.activeFamilyDoc = rvtPath || familyName || '';
        state.activeFamily = familyName || '';
        state.familyDetail = { columns: [], rows: [] };
        paintFamily();
        post('guid:request-family-detail', { runId: state.runId, rvtPath, familyName });
    }

    function filteredProjectRows() {
        if (!state.project.rows.length) return [];
        const idxPath = projectCol('RvtPath');
        const idxName = projectCol('RvtName');
        if (!state.activeProjectDoc) return state.project.rows;
        return state.project.rows.filter(row => {
            const path = idxPath >= 0 ? safe(row[idxPath]) : '';
            const name = idxName >= 0 ? safe(row[idxName]) : '';
            return (path || name) === state.activeProjectDoc;
        });
    }

    function filteredFamilyRows() {
        if (!state.familyDetail.rows.length) return [];
        const idxPath = colIndex('RvtPath');
        const idxFam = colIndex('FamilyName');
        const idxKind = colIndex('ParamKind');
        const idxShared = colIndex('IsShared');
        return state.familyDetail.rows.filter(row => {
            const path = idxPath >= 0 ? safe(row[idxPath]) : '';
            const famName = idxFam >= 0 ? safe(row[idxFam]) : '';
            const kind = idxKind >= 0 ? safe(row[idxKind]) : '';
            const shared = idxShared >= 0 ? safe(row[idxShared]) : '';
            const docMatch = idxPath < 0 || !state.activeFamilyDoc || ((path || '') === state.activeFamilyDoc);
            const famMatch = !state.activeFamily || (famName === state.activeFamily);
            const filterMatch = state.familyFilter === 'all' ||
                (state.familyFilter === 'shared' && (kind === 'Shared' || shared === 'Y')) ||
                (state.familyFilter === 'family' && (kind === 'Family' || shared === 'N'));
            return docMatch && famMatch && filterMatch;
        });
    }

    function splitFamilyRows() {
        const rows = filteredFamilyRows();
        const idxKind = colIndex('ParamKind');
        const userRows = [];
        const builtRows = [];
        rows.forEach(row => {
            const kind = idxKind >= 0 ? safe(row[idxKind]) : '';
            if (kind === 'BuiltIn') {
                builtRows.push(row);
            } else if (kind === 'Shared' || kind === 'Family') {
                userRows.push(row);
            } else {
                userRows.push(row);
            }
        });
        return { userRows, builtRows };
    }

    function colIndex(name) {
        return state.familyDetail.columns.findIndex(c => c === name);
    }

    function projectCol(name) {
        return state.project.columns.findIndex(c => c === name);
    }

    function indexCol(table, name) {
        return (table.columns || []).findIndex(c => c === name);
    }

    function buildHead(thead, columns, hidden) {
        thead.innerHTML = '';
        const tr = document.createElement('tr');
        columns.forEach(c => {
            if (hidden.has(c)) return;
            const th = document.createElement('th');
            th.textContent = c;
            tr.append(th);
        });
        thead.append(tr);
    }

    function paintVirtualRows(tbody, columns, rows, hidden) {
        tbody.innerHTML = '';
        let idx = 0;
        const chunk = () => {
            const frag = document.createDocumentFragment();
            for (let i = 0; i < 200 && idx < rows.length; i++, idx++) {
                const row = rows[idx];
                const tr = document.createElement('tr');
                columns.forEach((c, ci) => {
                    if (hidden.has(c)) return;
                    const td = document.createElement('td');
                    const text = safe(row[ci]);
                    td.textContent = text;
                    td.setAttribute('aria-label', text || '-');
                    tr.append(td);
                });
                frag.append(tr);
            }
            tbody.append(frag);
            if (idx < rows.length) setTimeout(chunk, 0);
        };
        chunk();
    }

    function hasRowsForExport() {
        const hasProject = (state.project.columns || []).length > 0;
        const hasFamily = state.includeFamily && ((state.familyDetail.columns || []).length > 0 || (state.familyNav.columns || []).length > 0);
        return hasProject || hasFamily;
    }

    function syncTabState() {
        btnTabProject.classList.toggle('is-active', state.activeTab === 'project');
        btnTabFamily.classList.toggle('is-active', state.activeTab === 'family');
        btnTabFamily.disabled = !state.includeFamily;
        tabPanelProject.classList.toggle('is-hidden', state.activeTab !== 'project');
        tabPanelFamily.classList.toggle('is-hidden', state.activeTab !== 'family');
        filterBar.classList.toggle('is-hidden', state.activeTab !== 'family');
        if (state.activeTab === 'family' && !state.includeFamily) {
            state.activeTab = 'project';
            tabPanelProject.classList.remove('is-hidden');
            tabPanelFamily.classList.add('is-hidden');
            filterBar.classList.add('is-hidden');
        }
        exportBtn.disabled = !hasRowsForExport();
        if (processButtons.export) processButtons.export.disabled = !!exportBtn.disabled;
        updateTabCounts();
    }

    function syncResultState() {
        const showResults = state.hasRun;
        tabPanels.style.display = showResults ? '' : 'none';
        emptyState.style.display = showResults ? 'none' : 'flex';
    }

    function syncAnnotationToggle() {
        if (!state.includeFamily) state.includeAnnotation = false;
    }

    function syncAuditConfigSummary() {
        if (!auditUi.summary) return;
        const familyText = state.includeFamily ? '포함' : '미포함';
        const annotationText = state.includeFamily && state.includeAnnotation ? '포함' : '미포함';
        auditUi.summary.textContent = `현재 설정: 패밀리 ${familyText} / 주석 패밀리 ${annotationText}`;
    }

    function setBusy(on) {
        state.busy = on;
        runBtn.disabled = on;
        exportBtn.disabled = on || !hasRowsForExport();
        pickCleanupExcelBtn.disabled = on;
        cleanupBtn.disabled = on || !state.cleanupSourceExcelPath;
        if (processButtons.run) processButtons.run.disabled = !!runBtn.disabled;
        if (processButtons.export) processButtons.export.disabled = !!exportBtn.disabled;
        if (processButtons.pickCleanupExcel) processButtons.pickCleanupExcel.disabled = !!pickCleanupExcelBtn.disabled;
        if (processButtons.cleanup) processButtons.cleanup.disabled = !!cleanupBtn.disabled;
    }

    function persistRvts() {
        state.rvtList = dedupPaths(state.rvtList);
        try { localStorage.setItem(LS_RVTS, JSON.stringify(state.rvtList || [])); } catch { }
    }

    function loadRvtList() {
        try {
            const raw = localStorage.getItem(LS_RVTS);
            const arr = JSON.parse(raw || '[]');
            if (Array.isArray(arr)) return dedupPaths(arr);
        } catch { }
        return [];
    }

    function persistGuidSettings() {
        const payload = {
            includeFamily: !!state.includeFamily,
            includeAnnotation: !!state.includeAnnotation,
            closeAllWorksetsOnOpen: !!state.closeAllWorksetsOnOpen,
            useSyncComment: !!state.useSyncComment,
            syncComment: state.syncComment || ''
        };
        try { localStorage.setItem(LS_SETTINGS, JSON.stringify(payload)); } catch { }
    }

    function loadGuidSettings() {
        try {
            const raw = localStorage.getItem(LS_SETTINGS);
            const parsed = JSON.parse(raw || '{}');
            if (parsed && typeof parsed === 'object') return parsed;
        } catch { }
        return {};
    }

    function handleRunProgress(payload) {
        if (!state.acceptRunProgress) return;
        const percent = typeof payload?.pct === 'number' ? payload.pct : 0;
        const message = payload?.detail || payload?.text || '';
        if (!state.busy && percent <= 0) return;
        if (!state.busy) setBusy(true);
        const subtitle = payload?.stage || buildRunProgressSubtitle(percent, message);
        ProgressDialog.show('파라미터 GUID 검토', subtitle);
        ProgressDialog.update(percent, subtitle, buildRunProgressDetail(percent, message));
    }

    function handleCleanupProgress(payload) {
        if (!state.acceptCleanupProgress) return;
        const percent = typeof payload?.pct === 'number' ? payload.pct : 0;
        const message = payload?.detail || payload?.text || '';
        if (!state.busy && percent <= 0) return;
        if (!state.busy) setBusy(true);
        const subtitle = payload?.stage || buildCleanupProgressSubtitle(percent, message);
        ProgressDialog.show('파라미터 GUID 정리', subtitle);
        ProgressDialog.update(percent, subtitle, message || `전체 진행률 ${formatRunPercent(percent)}`);
    }

    function handleExcelProgress(payload) {
        if (!state.acceptExcelProgress) return;
        const phase = normalizeExcelPhase(payload?.phase);
        const total = Number(payload?.total) || 0;
        const current = Number(payload?.current) || 0;
        const percent = computeExcelPercent(phase, current, total, payload?.phaseProgress);
        const subtitle = buildExcelSubtitle(phase, current, total);
        const detail = formatExcelDetail(phase, payload?.message);

        const exporting = phase !== 'DONE' && phase !== 'ERROR';
        if (!state.busy && exporting) setBusy(true);

        ProgressDialog.show('삭제용 엑셀 내보내기', subtitle || '파라미터 GUID 삭제용 엑셀을 내보내는 중입니다.');
        ProgressDialog.update(percent, subtitle, detail);

        if (!exporting) {
            setTimeout(() => { ProgressDialog.hide(); lastExcelPct = 0; setBusy(false); }, 260);
        }
    }

    function normalizeExcelPhase(phase) {
        return String(phase || '').trim().toUpperCase() || 'EXCEL_WRITE';
    }

    function computeExcelPercent(phase, current, total, phaseProgress) {
        const norm = normalizeExcelPhase(phase);
        if (norm === 'DONE') { lastExcelPct = 100; return 100; }
        if (norm === 'ERROR') return lastExcelPct;

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

    function clamp01(v) { const n = Number(v); if (Number.isFinite(n)) return Math.max(0, Math.min(1, n)); return 0; }

    function buildExcelSubtitle(phase, current, total) {
        const norm = normalizeExcelPhase(phase);
        switch (norm) {
            case 'EXCEL_INIT': return '파라미터 GUID 삭제용 엑셀 워크북을 준비하는 중입니다.';
            case 'EXCEL_WRITE': return `파라미터 GUID 삭제용 엑셀 데이터를 작성하는 중입니다. (${current}/${Math.max(total, current || 1)})`;
            case 'EXCEL_SAVE': return '파라미터 GUID 삭제용 엑셀을 내보내는 중입니다.';
            case 'AUTOFIT': return '파라미터 GUID 삭제용 엑셀 열 너비를 자동 조정하는 중입니다.';
            case 'DONE': return '파라미터 GUID 삭제용 엑셀 내보내기 완료';
            case 'ERROR': return '파라미터 GUID 삭제용 엑셀 내보내기 오류';
            default: return '파라미터 GUID 삭제용 엑셀을 내보내는 중입니다.';
        }
    }

    function formatExcelDetail(phase, message) {
        const norm = normalizeExcelPhase(phase);
        if (norm === 'AUTOFIT') return '파라미터 GUID 삭제용 엑셀 열 너비를 자동 조정하는 중입니다.';
        return message || '';
    }

    function buildRunProgressSubtitle(percent, message) {
        const raw = String(message || '').trim();
        if (!raw) return '파라미터 GUID 검토를 진행하는 중입니다.';
        if (raw.includes('문서 여는 중')) return '파라미터 GUID 검토 대상 문서를 여는 중입니다.';
        if (raw.includes('프로젝트 파라미터')) return `프로젝트 파라미터 GUID를 검토하는 중입니다. (${formatRunPercent(percent)})`;
        if (raw.includes('패밀리 처리 중')) return `패밀리 파라미터 GUID를 검토하는 중입니다. (${formatRunPercent(percent)})`;
        if (raw.includes('실패')) return '문서 처리 실패';
        if (raw.includes('완료')) return '파라미터 GUID 검토 완료';
        return `파라미터 GUID 검토를 진행하는 중입니다. (${formatRunPercent(percent)})`;
    }

    function buildCleanupProgressSubtitle(percent, message) {
        const raw = String(message || '').trim();
        if (!raw) return '파라미터 GUID 정리를 진행하는 중입니다.';
        if (raw.includes('문서 여는 중')) return '삭제 대상 문서를 여는 중입니다.';
        if (raw.includes('프로젝트 파라미터 정리 중')) return `프로젝트 파라미터를 정리하는 중입니다. (${formatRunPercent(percent)})`;
        if (raw.includes('로드 패밀리 파라미터 정리 중')) return `로드 패밀리를 정리하는 중입니다. (${formatRunPercent(percent)})`;
        if (raw.includes('완료')) return '파라미터 GUID 정리 완료';
        return `파라미터 GUID 정리를 진행하는 중입니다. (${formatRunPercent(percent)})`;
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

    function onRemoveSelected() {
        if (!state.rvtChecked.size) { toast('제거할 RVT를 선택해 주세요.', 'warn'); return; }
        state.rvtList = state.rvtList.filter(p => !state.rvtChecked.has(p));
        state.rvtChecked.clear();
        persistRvts();
        renderRvtList();
    }

    function appendDroppedRvts(paths) {
        let added = 0;
        (Array.isArray(paths) ? paths : []).forEach((entry) => {
            const path = normalizeRvtPath(entry);
            if (!path) return;
            const exists = state.rvtList.some((item) => samePath(item, path));
            if (!exists) {
                state.rvtList.push(path);
                added += 1;
            }
            state.rvtChecked.add(path);
        });
        state.rvtList = dedupPaths(state.rvtList);
        return added;
    }

    function syncRvtActionState() {
        if (btnRemove) btnRemove.disabled = state.rvtChecked.size === 0;
        if (btnClear) btnClear.disabled = state.rvtList.length === 0;
        rvtMeta.textContent = `파일 ${state.rvtList.length}개`;
        rvtSummary.textContent = `파일 ${state.rvtList.length}개`;
        if (processButtons.run) processButtons.run.disabled = !!runBtn.disabled;
        if (processButtons.export) processButtons.export.disabled = !!exportBtn.disabled;
        if (processButtons.cleanup) processButtons.cleanup.disabled = !!cleanupBtn.disabled;
    }

    function syncCleanupPanel() {
        if (processUi.meta) {
            processUi.meta.textContent = state.cleanupSourceExcelPath ? '삭제용 엑셀 선택 완료' : '삭제용 엑셀 선택 필요';
        }
        if (processUi.path) {
            processUi.path.textContent = state.cleanupSourceExcelPath || '아직 선택한 삭제용 엑셀이 없습니다.';
        }
        cleanupBtn.disabled = !!state.busy || !state.cleanupSourceExcelPath;
        if (processButtons.cleanup) processButtons.cleanup.disabled = !!cleanupBtn.disabled;
        if (processSync.sync) processSync.sync();
    }

    function showGuidCompletionDialog(payload) {
        const projectCount = state.project.rows.length || 0;
        const familyCount = state.includeFamily ? (state.familyNav.rows.length || 0) : 0;
        const projectDocCount = countUniqueDocEntries(state.project.columns, state.project.rows);
        const familyDocCount = state.includeFamily ? countUniqueDocEntries(state.familyNav.columns, state.familyNav.rows) : 0;
        const notes = [];

        if (state.includeFamily) notes.push('패밀리 탭에서 RVT별 패밀리 상세를 이어서 확인할 수 있습니다.');
        if (payload?.includeAnnotation) notes.push('주석 패밀리까지 포함해 GUID를 비교했습니다.');
        notes.push("삭제할 행은 엑셀의 '삭제여부' 열에 '삭제'를 입력한 뒤, 삭제용 엑셀 불러오기로 다시 실행하면 됩니다.");

        showCompletionSummaryDialog({
            title: '파라미터 GUID 검토 완료',
            message: '프로젝트와 패밀리 GUID 검토가 끝났습니다. 아래 탭에서 상세 결과를 확인하거나 삭제용 엑셀로 바로 내보낼 수 있습니다.',
            summaryItems: [
                { label: '프로젝트 결과 건수', value: String(projectCount) },
                { label: '프로젝트 대상 RVT', value: `${projectDocCount}개` },
                { label: '패밀리 결과 건수', value: String(familyCount) },
                { label: '패밀리 대상 RVT', value: `${familyDocCount}개` },
                { label: '패밀리 검토 포함', value: state.includeFamily ? '예' : '아니오' }
            ],
            notes,
            exportDisabled: !!exportBtn.disabled,
            onExport: () => exportBtn.click()
        });
    }

    function showGuidCleanupDialog(payload) {
        const successCount = Number(payload?.successCount) || 0;
        const failCount = Number(payload?.failCount) || 0;
        const noChangeCount = Number(payload?.noChangeCount) || 0;
        const deletedCount = Number(payload?.deletedCount) || 0;
        const instructionCount = Number(payload?.instructionCount) || 0;
        const notes = [];

        if (payload?.sourceExcelPath) notes.push(`삭제 기준 엑셀: ${payload.sourceExcelPath}`);
        if (payload?.settings?.closeAllWorksetsOnOpen) notes.push('워크셰어링 문서는 모든 웍셋을 닫은 상태로 열었습니다.');
        if (payload?.settings?.useSyncComment) notes.push(`동기화 코멘트 사용: ${payload?.settings?.syncComment || '(빈 코멘트)'}`);

        showCompletionSummaryDialog({
            title: '파라미터 GUID 정리 완료',
            message: payload?.message || '삭제용 엑셀 기준 정리가 완료되었습니다.',
            summaryItems: [
                { label: '삭제 요청 건수', value: String(instructionCount) },
                { label: '실제 삭제 건수', value: String(deletedCount) },
                { label: '성공 파일', value: `${successCount}개` },
                { label: '변경 없는 파일', value: `${noChangeCount}개` },
                { label: '실패 파일', value: `${failCount}개` }
            ],
            notes,
            showExport: false
        });
    }

    function countUniqueDocEntries(columns, rows) {
        const list = Array.isArray(rows) ? rows : [];
        if (!list.length) return 0;
        const idxPath = Array.isArray(columns) ? columns.findIndex((col) => col === 'RvtPath') : -1;
        const idxName = Array.isArray(columns) ? columns.findIndex((col) => col === 'RvtName') : -1;
        const docs = new Set();
        list.forEach((row) => {
            const path = idxPath >= 0 ? safe(row?.[idxPath]) : '';
            const name = idxName >= 0 ? safe(row?.[idxName]) : '';
            const key = path || name;
            if (key) docs.add(key);
        });
        return docs.size;
    }

    function updateTabCounts() {
        setTabCount(btnTabProject, state.project.rows.length || 0);
        const detCount = state.includeFamily ? (state.familyNav.rows.length || 0) : 0;
        setTabCount(btnTabFamily, detCount);
    }
}

function normalizeRvtPath(entry) {
    if (!entry) return '';
    if (typeof entry === 'string') return entry.trim();
    if (typeof entry === 'object') {
        if (typeof entry.path === 'string') return entry.path.trim();
        if (typeof entry.fullPath === 'string') return entry.fullPath.trim();
    }
    return '';
}

function dedupPaths(list) {
    const seen = new Set();
    const clean = [];
    (list || []).forEach(item => {
        const path = normalizeRvtPath(item);
        if (!path) return;
        const key = path.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        clean.push(path);
    });
    return clean;
}

function setTabCount(btn, count) {
    if (!btn) return;
    const badge = btn.querySelector('.pill-count');
    if (!badge) return;
    badge.textContent = Number.isFinite(count) ? count : 0;
}

function safe(v) {
    if (v === null || v === undefined) return '';
    return String(v);
}

function samePath(a, b) {
    if (!a || !b) return false;
    return a.toLowerCase() === b.toLowerCase();
}

function cardBtn(text, onClick) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'btn card-btn';
    btn.textContent = text;
    btn.onclick = onClick;
    return btn;
}

function buildProcessRow(step, text) {
    const row = div('guid-setting-process__row');
    const num = div('guid-setting-process__num');
    num.textContent = step;
    const copy = document.createElement('div');
    copy.textContent = text;
    row.append(num, copy);
    return row;
}
