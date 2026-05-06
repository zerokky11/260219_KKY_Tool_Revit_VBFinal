import { clear, div, toast, showExcelSavedDialog, chooseExcelMode } from '../core/dom.js';
import { refreshUiAfterHostDialog } from '../core/hostDialog.js';
import { getLastExcelExportLocale } from '../core/dom.js';
import { attachRvtDropZone } from '../core/rvtDrop.js';
import { ProgressDialog } from '../core/progress.js';
import { post, onHost } from '../core/bridge.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js?v=20260504a';

const DEFAULT_SUMMARY = { okCount: 0, failCount: 0, skipCount: 0 };
const CATEGORY_PRESET_KEY = "kky_spb_cat_presets";
const EXCEL_PHASE_WEIGHT = { EXCEL_INIT: 0.05, EXCEL_WRITE: 0.85, EXCEL_SAVE: 0.08, AUTOFIT: 0.02, DONE: 1, ERROR: 1 };

function formatRevitParamGroupLabel(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  if (raw.toUpperCase() === 'INVALID') return 'Other';
  const source = raw.toUpperCase().startsWith('PG_') ? raw.substring(3) : raw;
  return source
    .split('_')
    .filter(Boolean)
    .map((part) => {
      const lower = String(part || '').toLowerCase();
      return lower ? lower.charAt(0).toUpperCase() + lower.slice(1) : '';
    })
    .join(' ');
}

export function renderSharedParamBatch(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);
  const top = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar');
  if (top) top.classList.add('hub-topbar');

  const state = {
    spFilePath: '',
    groups: [],
    defsByGroup: {},
    categoryTree: [],
    paramGroups: [],
    selectedGroup: '',
    selectedParams: [],
    defSelection: new Set(),
    defSearch: '',
    rvtList: [],
    rvtChecked: new Set(),
    pendingFolderBrowse: false,
    options: {
      closeAllWorksetsOnOpen: true,
      syncComment: ''
    },
    summary: { ...DEFAULT_SUMMARY },
    logs: [],
    running: false,
    lastProgressPct: 0,
    categoryFilterText: ''
  };
  let progressHideTimer = null;

  const page = div('feature-shell sharedparambatch-page');
  const header = div('feature-header');
  const heading = div('feature-heading');
  heading.innerHTML = `
    <span class="feature-kicker">프로젝트 파라미터</span>
    <h2 class="feature-title">프로젝트/공유 파라미터 일괄 추가</h2>
    <p class="feature-sub">프로젝트 또는 공유 파라미터를 여러 RVT에 일괄 추가하고 바인딩합니다.</p>`;

  const btnRun = cardBtn('실행', onRun);
  const btnExport = cardBtn('엑셀 내보내기', onExport, 'btn--secondary');
  btnExport.disabled = true;
  const btnViewResult = cardBtn('결과 보기', openResultModal, 'btn--secondary');
  btnViewResult.disabled = true;

  header.append(heading);
  page.append(header);

  const layout = div('sharedparambatch-layout');
  page.append(layout);

  const warningSection = div('section sharedparambatch-section spb-warning');
  const warningText = document.createElement('div');
  warningText.textContent = '공유파라미터 TXT가 Revit에 설정되어 있지 않습니다. Revit의 관리 > 공유 매개변수에서 등록한 뒤 새로고침해 주세요.';
  const warningActions = div('section-actions');
  warningActions.append(cardBtn('새로고침', () => post('sharedparambatch:init', {}), 'btn--secondary'));
  warningSection.append(warningText, warningActions);
  warningSection.style.display = 'none';

  const selectSection = div('section sharedparambatch-section');
  const paramPickerBtn = cardBtn('파라미터 선택', openParamPicker, 'btn--secondary');
  const selectHeader = buildSelectHeader('공유파라미터 선택', paramPickerBtn, btnRun, btnExport);
  selectSection.append(selectHeader);
  const selectHeaderActions = selectHeader.querySelector('.spb-headerRight');
  if (selectHeaderActions) {
    selectHeaderActions.insertBefore(btnViewResult, btnExport);
  }

  const selectedSection = div('section sharedparambatch-section');
  const bulkApplyBtn = cardBtn('설정 일괄 적용', applyBulkSettings, 'btn--secondary');
  selectedSection.append(sectionHeader('선택한 파라미터', [bulkApplyBtn]));
  const paramTable = document.createElement('table');
  paramTable.className = 'sharedparambatch-table';
  paramTable.innerHTML = '<thead><tr><th>이름</th><th>GUID</th><th>바인딩</th><th>그룹</th><th>카테고리</th><th>작업</th></tr></thead><tbody></tbody>';
  const paramBody = paramTable.querySelector('tbody');
  const paramWrap = div('spb-tableWrap');
  paramWrap.append(paramTable);
  selectedSection.append(paramWrap);

  const rvtSection = div('section sharedparambatch-section');
  const rvtHeader = div('spb-rvtHeader');
  const rvtHeaderLeft = div('spb-rvtHeaderLeft');
  const rvtTitle = document.createElement('h3');
  rvtTitle.textContent = 'RVT 파일';
  const rvtActions = div('section-actions');
  rvtActions.append(
    cardBtn('추가', () => post('sharedparambatch:browse-rvts', {})),
    cardBtn('폴더 선택', onBrowseFolder),
    cardBtn('선택 삭제', removeSelectedRvts, 'btn--secondary'),
    cardBtn('전체 삭제', clearRvts, 'btn--secondary'),
    cardBtn('리스트 크게보기', openRvtModal, 'btn--secondary')
  );
  rvtHeaderLeft.append(rvtTitle, rvtActions);

  const rvtHeaderRight = div('spb-rvtHeaderRight');
  const closeWrap = document.createElement('label');
  closeWrap.className = 'spb-inlineCheck';
  const closeChk = document.createElement('input');
  closeChk.type = 'checkbox';
  closeChk.checked = true;
  closeChk.id = 'spb-close-worksets';
  closeChk.addEventListener('change', () => { state.options.closeAllWorksetsOnOpen = !!closeChk.checked; });
  const closeLbl = document.createElement('span');
  closeLbl.textContent = '워크셰어링 파일: 모든 웍셋 닫고 열기';
  closeWrap.append(closeChk, closeLbl);

  const syncInput = document.createElement('input');
  syncInput.type = 'text';
  syncInput.className = 'sharedparambatch-input spb-syncInput';
  syncInput.placeholder = '동기화 코멘트';
  syncInput.addEventListener('input', () => { state.options.syncComment = syncInput.value || ''; });
  rvtHeaderRight.append(closeWrap, labelSpan('동기화 코멘트'), syncInput);

  rvtHeader.append(rvtHeaderLeft, rvtHeaderRight);
  rvtSection.append(rvtHeader);
  const rvtHint = div('rvt-drop-hint');
  rvtHint.textContent = 'RVT 추가 버튼을 누르거나 탐색기에서 .rvt 파일을 이 목록으로 끌어다 놓으면 바로 등록됩니다.';
  rvtSection.append(rvtHint);
  const { table: rvtTable, tbody: rvtBody, master: rvtMaster } = createRvtTable();
  const rvtListWrap = div('spb-rvtTableWrap rvt-drop-zone');
  rvtListWrap.append(rvtTable);
  rvtSection.append(rvtListWrap);
  attachRvtDropZone(rvtListWrap, {
    onDropPaths: (paths) => {
      const added = appendDroppedRvts(paths);
      if (!added) {
        toast('이미 등록된 RVT입니다.', 'warn');
        return;
      }
      renderRvtList();
      toast(`${added}개 RVT를 추가했습니다.`, 'ok');
    },
    onInvalid: () => toast('RVT 파일만 드래그해서 추가할 수 있습니다.', 'warn')
  });

  const rvtModal = buildRvtModal();
  const paramPickerModal = buildParamPickerModal();
  const resultModal = buildResultModal();

  const resultSection = div('sharedparambatch-result');
  const resultHeader = sectionHeader('최근 실행 결과', []);
  resultSection.append(resultHeader);
  const summaryRow = div('sharedparambatch-summary');
  const badgeOk = summaryBadge('정상', '0');
  const badgeFail = summaryBadge('실패', '0');
  const badgeSkip = summaryBadge('건너뜀', '0');
  summaryRow.append(badgeOk.wrap, badgeFail.wrap, badgeSkip.wrap);

  const logPathRow = div('sharedparambatch-logpath');
  const logPathText = document.createElement('span');
  logPathText.textContent = '로그 파일: -';
  const logOpenBtn = cardBtn('폴더 열기', () => {
    if (!state.logTextPath) { toast('로그 파일이 없습니다.', 'err'); return; }
    post('sharedparambatch:open-folder', { path: state.logTextPath });
  }, 'btn--secondary');
  logOpenBtn.disabled = true;
  logPathRow.append(logPathText, logOpenBtn);

  const logTable = document.createElement('table');
  logTable.className = 'sharedparambatch-table';
  logTable.innerHTML = '<thead><tr><th>단계</th><th>RVT</th><th>메시지</th></tr></thead><tbody></tbody>';
  const logBody = logTable.querySelector('tbody');
  resultSection.append(summaryRow, logPathRow, logTable);
  resultSection.style.display = 'none';
  resultModal.body.append(summaryRow, logTable);

  layout.append(warningSection, selectSection, selectedSection, rvtSection);
  page.append(layout);
  page.append(buildSettingsModal(), paramPickerModal, rvtModal, resultModal.overlay);

  target.append(page);

  onHost('sharedparambatch:init', handleInit);
  onHost('sharedparambatch:rvts-picked', handleRvtsPicked);
  onHost('sharedparambatch:progress', handleProgress);
  onHost('sharedparambatch:done', handleDone);
  onHost('sharedparambatch:exported', handleExported);
  onHost('revit:error', ({ message }) => {
    state.pendingFolderBrowse = false;
    finishRunning(true);
    toast(message || '프로젝트 파라미터 추가 중 Revit 오류가 발생했습니다. 대상 RVT와 선택한 파라미터 설정을 확인한 뒤 다시 실행해 주세요. 계속 실패하면 메시지를 관리자에게 전달해 주세요.', 'err', 3200);
  });
  onHost('host:error', ({ message }) => {
    state.pendingFolderBrowse = false;
    finishRunning(true);
    toast(message || '프로젝트 파라미터 추가 중 호스트 오류가 발생했습니다. Hub 화면을 다시 열고 Revit 연결 상태를 확인한 뒤 다시 실행해 주세요. 계속 실패하면 메시지를 관리자에게 전달해 주세요.', 'err', 3200);
  });

  post('sharedparambatch:init', {});
  renderRvtList();

  function handleInit(payload) {
    if (!payload || !payload.ok) {
      warningSection.style.display = 'flex';
      toast(payload?.message || '프로젝트 파라미터 추가 화면을 초기화하지 못했습니다. 공유파라미터 파일 연결 상태를 확인한 뒤 Hub 화면을 다시 열어 주세요. 계속 실패하면 메시지를 관리자에게 전달해 주세요.', 'err');
      return;
    }
    warningSection.style.display = 'none';
    state.spFilePath = payload.spFilePath || '';
    state.groups = Array.isArray(payload.groups) ? payload.groups : [];
    state.defsByGroup = payload.defsByGroup || {};
    state.categoryTree = Array.isArray(payload.categoryTree) ? payload.categoryTree : [];
    state.paramGroups = Array.isArray(payload.paramGroups) ? payload.paramGroups : [];
    if (syncInput) syncInput.value = state.options.syncComment || '';
    closeChk.checked = !!state.options.closeAllWorksetsOnOpen;
    renderGroupOptions();
    renderDefinitionList();
    renderSelectedParams();
  }

  function renderGroupOptions() {
    const modal = buildParamPickerModal;
    const groupSelect = modal.groupSelect;
    if (!groupSelect) return;
    groupSelect.innerHTML = '';
    if (!state.groups.length) {
      const opt = document.createElement('option');
      opt.value = '';
      opt.textContent = '그룹이 없습니다.';
      groupSelect.append(opt);
      state.selectedGroup = '';
      return;
    }
    state.groups.forEach((g, idx) => {
      const opt = document.createElement('option');
      opt.value = g;
      opt.textContent = g;
      groupSelect.append(opt);
      if (idx === 0) state.selectedGroup = g;
    });
    groupSelect.value = state.selectedGroup;
  }

  function renderDefinitionList() {
    const modal = buildParamPickerModal;
    const defList = modal.defList;
    if (!defList) return;
    defList.innerHTML = '';
    const defs = state.defsByGroup[state.selectedGroup] || [];
    const search = (state.defSearch || '').toLowerCase();
    const filtered = search
      ? defs.filter(d => (d.name || '').toLowerCase().includes(search))
      : defs;
    if (!filtered.length) {
      const opt = document.createElement('option');
      opt.value = '';
      opt.textContent = '정의가 없습니다.';
      defList.append(opt);
      return;
    }
    filtered.forEach((d) => {
      const opt = document.createElement('option');
      opt.value = d.guid;
      opt.textContent = `${d.name} (${d.paramTypeLabel || ''})`;
      opt.dataset.name = d.name;
      opt.dataset.group = state.selectedGroup;
      opt.dataset.guid = d.guid;
      opt.dataset.paramTypeLabel = d.paramTypeLabel || '';
      opt.dataset.desc = d.desc || '';
      if (state.defSelection.has(d.guid)) opt.selected = true;
      defList.append(opt);
    });
  }

  function openParamPicker() {
    const modal = buildParamPickerModal;
    if (modal.searchInput) modal.searchInput.value = state.defSearch || '';
    if (modal.groupSelect) modal.groupSelect.value = state.selectedGroup || '';
    renderDefinitionList();
    if (modal.overlay) modal.overlay.classList.add('is-open');
  }

  function onAddSelectedDefs() {
    const modal = buildParamPickerModal;
    const defList = modal.defList;
    if (!defList || !defList.options.length) return;
    const selected = Array.from(defList.selectedOptions || []);
    if (!selected.length) { toast('추가할 파라미터를 선택해 주세요.', 'err'); return; }

    const nextSelection = new Set(state.defSelection);
    selected.forEach((opt) => {
      const guid = opt.dataset.guid;
      if (!guid || state.selectedParams.some(p => p.guid === guid)) return;
      nextSelection.add(guid);
      state.selectedParams.push({
        groupName: opt.dataset.group || state.selectedGroup,
        name: opt.dataset.name || opt.textContent,
        guid,
        paramTypeLabel: opt.dataset.paramTypeLabel || '',
        desc: opt.dataset.desc || '',
        settings: {
          isInstanceBinding: true,
          paramGroup: getDefaultParamGroupValue(),
          allowVaryBetweenGroups: false,
          categories: []
        }
      });
    });
    state.defSelection = nextSelection;
    renderSelectedParams();
  }

  function renderSelectedParams() {
    paramBody.innerHTML = '';
    if (!state.selectedParams.length) {
      const tr = document.createElement('tr');
      tr.className = 'empty-row';
      const td = document.createElement('td');
      td.className = 'empty-cell';
      td.colSpan = 6;
      td.textContent = '선택된 파라미터가 없습니다.';
      tr.append(td);
      paramBody.append(tr);
      updateButtons();
      return;
    }

    state.selectedParams.forEach((p, idx) => {
      const tr = document.createElement('tr');
      tr.append(tdText(p.name));
      tr.append(tdText(p.guid));
      tr.append(tdText(p.settings.isInstanceBinding ? '인스턴스' : '타입'));
      tr.append(tdText(formatParamGroup(p.settings.paramGroup)));
      tr.append(tdText(String(p.settings.categories.length)));

      const actionTd = document.createElement('td');
      const btnSetting = actionBtn('설정', () => openSettingsModal(idx));
      const btnRemove = actionBtn('삭제', () => {
        state.selectedParams.splice(idx, 1);
        renderSelectedParams();
      }, 'btn--ghost');
      actionTd.append(btnSetting, btnRemove);
      tr.append(actionTd);
      paramBody.append(tr);
    });
    updateButtons();
  }

  function renderRvtList() {
    const allChecked = state.rvtList.length > 0 && state.rvtList.every(f => state.rvtChecked.has(f));
    [rvtMaster, buildRvtModal.master].forEach((master) => {
      if (!master) return;
      master.checked = allChecked;
      master.disabled = state.rvtList.length === 0;
      master.onchange = () => {
        if (master.checked) state.rvtChecked = new Set(state.rvtList);
        else state.rvtChecked.clear();
        renderRvtList();
        updateButtons();
      };
    });
    const rows = state.rvtList.map((p, idx) => ({
      checked: state.rvtChecked.has(p),
      index: idx + 1,
      name: getRvtName(p),
      path: p,
      title: p,
      onToggle: (checked) => {
        if (checked) state.rvtChecked.add(p);
        else state.rvtChecked.delete(p);
        updateButtons();
      }
    }));
    renderRvtRows(rvtBody, rows);
    if (buildRvtModal.tbody) {
      renderRvtRows(buildRvtModal.tbody, rows);
    }
    updateButtons();
  }

  function handleRvtsPicked(payload) {
    if (state.pendingFolderBrowse) {
      ProgressDialog.hide();
      state.pendingFolderBrowse = false;
    }
    if (!payload || !payload.ok) {
      if (payload?.message) toast(payload.message, 'err');
      return;
    }
    const droppedPaths = Array.isArray(payload?.rvtPaths)
      ? payload.rvtPaths
      : (Array.isArray(payload?.paths) ? payload.paths : []);
    const added = appendDroppedRvts(droppedPaths);
    const changed = added > 0;
    if (changed) {
      refreshUiAfterHostDialog(() => renderRvtList());
    }
    if (payload?.fromFolder) {
      toast(added ? `${added}개 추가 완료` : '추가된 RVT가 없습니다.', added ? 'ok' : 'err');
    }
  }

  function removeSelectedRvts() {
    if (!state.rvtChecked.size) { toast('삭제할 RVT를 선택해 주세요.', 'err'); return; }
    state.rvtList = state.rvtList.filter(p => !state.rvtChecked.has(p));
    state.rvtChecked.clear();
    renderRvtList();
  }

  function clearRvts() {
    state.rvtList = [];
    state.rvtChecked.clear();
    renderRvtList();
  }

  function appendDroppedRvts(paths) {
    let added = 0;
    (Array.isArray(paths) ? paths : []).forEach((path) => {
      if (!path) return;
      if (!state.rvtList.includes(path)) {
        state.rvtList.push(path);
        added += 1;
      }
      state.rvtChecked.add(path);
    });
    return added;
  }

  function onBrowseFolder() {
    if (state.running) return;
    state.pendingFolderBrowse = true;
    ProgressDialog.show('RVT 폴더 선택', '폴더 내 RVT 파일을 찾는 중입니다.');
    ProgressDialog.update(20, '폴더 내 RVT 파일을 찾는 중입니다.', '하위 폴더를 포함해 RVT 파일을 검색하는 중입니다.');
    post('sharedparambatch:browse-folder', {});
  }

  function onRun() {
    if (state.running) return;
    if (!state.selectedParams.length) { toast('파라미터를 선택해 주세요.', 'err'); return; }
    if (!state.rvtList.length) { toast('RVT 파일을 추가해 주세요.', 'err'); return; }
    const selectedRvts = getCheckedRvtPaths();
    if (!selectedRvts.length) { toast('적용할 RVT를 1개 이상 선택해 주세요.', 'err'); return; }
    const missingCats = state.selectedParams.filter(p => !p.settings.categories.length);
    if (missingCats.length) {
      toast('카테고리를 지정하지 않은 파라미터가 있습니다.', 'err');
      return;
    }

    const payload = {
      spFilePath: state.spFilePath,
      rvtPaths: selectedRvts,
      parameters: state.selectedParams.map(p => ({
        groupName: p.groupName,
        paramName: p.name,
        guid: p.guid,
        paramTypeLabel: p.paramTypeLabel,
        description: p.desc,
        settings: {
          isInstanceBinding: p.settings.isInstanceBinding,
          paramGroup: p.settings.paramGroup,
          allowVaryBetweenGroups: p.settings.allowVaryBetweenGroups,
          categories: p.settings.categories
        }
      })),
      closeAllWorksetsOnOpen: state.options.closeAllWorksetsOnOpen,
      syncComment: state.options.syncComment
    };

    state.running = true;
    updateButtons();
    ProgressDialog.show('프로젝트 파라미터 추가', '작업을 준비하는 중입니다.');
    ProgressDialog.update(0, '작업을 준비하는 중입니다.', '선택한 파라미터와 RVT 목록을 정리하는 중입니다.');
    post('sharedparambatch:run', payload);
  }

  function handleProgress(payload) {
    if (!payload) return;
    if (payload.phase || payload.current != null || payload.phaseProgress != null) {
      handleExcelProgress(payload);
      return;
    }
    handleRunProgress(payload);
  }

  function handleRunProgress(payload) {
    clearProgressHideTimer();
    const total = Number(payload.total || payload.Total || 0);
    const step = Number(payload.step || payload.Step || 0);
    const text = payload.text || payload.message || '';
    const pct = Math.max(0, Math.min(100, total > 0 ? (step / total) * 100 : (Number(payload.percent) || 0)));
    state.lastProgressPct = pct;
    const subtitle = buildRunProgressSubtitle(pct, text, step, total);
    const detail = buildRunProgressDetail(text, step, total);
    ProgressDialog.show('프로젝트 파라미터 추가', subtitle);
    ProgressDialog.update(pct, subtitle, detail);
    if (pct >= 100) scheduleProgressHide();
  }

  function handleExcelProgress(payload) {
    clearProgressHideTimer();
    const phase = normalizeExcelPhase(payload?.phase);
    const total = Number(payload?.total) || 0;
    const current = Number(payload?.current) || 0;
    const percent = computeExcelPercent(phase, current, total, payload?.phaseProgress, payload?.percent);
    const subtitle = buildExcelSubtitle(phase, current, total);
    const detail = formatExcelDetail(phase, payload?.message, current, total);

    state.lastProgressPct = percent;
    ProgressDialog.show('프로젝트 파라미터 결과 내보내기', subtitle);
    ProgressDialog.update(percent, subtitle, detail);

    if (phase === 'DONE' || phase === 'ERROR') {
      scheduleProgressHide(260);
    }
  }

  function handleDone(payload) {
    finishRunning(true);
    if (!payload || !payload.ok) {
      toast(payload?.message || '프로젝트 파라미터 추가 실행에 실패했습니다. 대상 RVT와 선택한 파라미터 설정을 확인한 뒤 다시 실행해 주세요. 계속 실패하면 메시지를 관리자에게 전달해 주세요.', 'err');
      return;
    }
    state.summary = payload.summary || { ...DEFAULT_SUMMARY };
    state.logs = Array.isArray(payload.logs) ? payload.logs : [];
    renderSummary();
    renderLogs();
    if (hasRunResults()) {
      requestAnimationFrame(() => openResultModal());
    }
    if (state.logs.length) toast('완료되었습니다.', 'ok');
  }

  function finishRunning(resetProgress) {
    state.running = false;
    updateButtons();
    if (resetProgress) {
      clearProgressHideTimer();
      state.lastProgressPct = 0;
      ProgressDialog.hide();
    }
  }

  async function onExport() {
    if (state.running) return;
    if (!state.logs.length) { toast('내보낼 로그가 없습니다.', 'err'); return; }
    const excelMode = await chooseExcelMode();
    if (!excelMode) return;
    state.running = true;
    updateButtons();
    ProgressDialog.show('프로젝트 파라미터 결과 내보내기', '엑셀 파일을 준비하는 중입니다.');
    ProgressDialog.update(0, '엑셀 파일을 준비하는 중입니다.', '로그 결과와 저장 옵션을 정리하는 중입니다.');
    post('sharedparambatch:export-excel', { excelMode, locale: getLastExcelExportLocale() });
  }

  function handleExported(payload) {
    finishRunning(true);
    if (!payload || !payload.ok) {
      toast(payload?.message || '프로젝트 파라미터 추가 결과 엑셀 내보내기에 실패했습니다. 저장 경로 권한과 파일이 열려 있는지 확인해 주세요. 계속 실패하면 메시지를 관리자에게 전달해 주세요.', 'err');
      return;
    }
    const path = payload.filePath || payload.path;
    requestAnimationFrame(() => {
      showExcelSavedDialog('프로젝트/공유 파라미터 일괄 추가 결과 엑셀을 저장했습니다.', path, (p) => post('excel:open', { path: p }));
    });
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

  function buildRunProgressSubtitle(percent, message, step, total) {
    const raw = String(message || '').trim();
    if (raw.includes('정의 로드')) return '공유파라미터 정의를 읽는 중입니다.';
    if (raw.includes('완료')) return '프로젝트 파라미터 추가 완료';
    if (total > 0) return `RVT를 처리하는 중입니다. (${Math.max(step, 0)}/${total})`;
    return `프로젝트 파라미터를 추가하는 중입니다. (${formatProgressPercent(percent)})`;
  }

  function buildRunProgressDetail(message, step, total) {
    const raw = String(message || '').trim();
    if (raw) return raw;
    if (total > 0) return `${Math.max(step, 0)} / ${total}`;
    return '';
  }

  function normalizeExcelPhase(phase) {
    return String(phase || '').trim().toUpperCase() || 'EXCEL_WRITE';
  }

  function computeExcelPercent(phase, current, total, phaseProgress, fallbackPercent) {
    const norm = normalizeExcelPhase(phase);
    if (norm === 'DONE') { state.lastProgressPct = 100; return 100; }
    if (norm === 'ERROR') return state.lastProgressPct;

    if (Number.isFinite(fallbackPercent)) {
      const direct = Math.max(state.lastProgressPct, Math.min(100, Number(fallbackPercent)));
      state.lastProgressPct = direct;
      return direct;
    }

    const completed = ['EXCEL_INIT', 'EXCEL_WRITE', 'EXCEL_SAVE', 'AUTOFIT'].reduce((acc, key) => {
      if (key === norm) return acc;
      return acc + (EXCEL_PHASE_WEIGHT[key] || 0);
    }, 0);
    const weight = EXCEL_PHASE_WEIGHT[norm] || 0;
    const ratio = total > 0 ? Math.max(0, Math.min(1, current / total)) : 0;
    const staged = Math.max(ratio, clamp01(phaseProgress));
    const pct = Math.min(100, Math.max(state.lastProgressPct, (completed + weight * staged) * 100));
    state.lastProgressPct = pct;
    return pct;
  }

  function buildExcelSubtitle(phase, current, total) {
    const norm = normalizeExcelPhase(phase);
    switch (norm) {
      case 'EXCEL_INIT': return '엑셀 워크북을 준비하는 중입니다.';
      case 'EXCEL_WRITE': return `엑셀 로그를 작성하는 중입니다. (${current}/${Math.max(total, current || 1)})`;
      case 'EXCEL_SAVE': return '엑셀 파일을 저장하는 중입니다.';
      case 'AUTOFIT': return '엑셀 스타일을 적용하는 중입니다.';
      case 'DONE': return '프로젝트/공유 파라미터 일괄 추가 결과 엑셀 내보내기 완료';
      case 'ERROR': return '프로젝트/공유 파라미터 일괄 추가 결과 엑셀 내보내기 오류';
      default: return '프로젝트/공유 파라미터 일괄 추가 결과 엑셀 내보내기를 진행하는 중입니다.';
    }
  }

  function formatExcelDetail(phase, message, current, total) {
    const raw = String(message || '').trim();
    if (raw) return raw;
    const norm = normalizeExcelPhase(phase);
    if (norm === 'EXCEL_WRITE' && total > 0) return `${current} / ${total} 로그를 기록하는 중입니다.`;
    if (norm === 'AUTOFIT') return '셀 스타일과 열 너비를 정리하는 중입니다.';
    if (norm === 'DONE') return '프로젝트/공유 파라미터 일괄 추가 결과 엑셀 파일이 저장되었습니다.';
    return '';
  }

  function clamp01(value) {
    const num = Number(value);
    if (!Number.isFinite(num)) return 0;
    return Math.max(0, Math.min(1, num));
  }

  function formatProgressPercent(percent) {
    const safe = Number(percent);
    if (!Number.isFinite(safe)) return '0%';
    return `${Math.max(0, Math.min(100, Math.round(safe * 10) / 10))}%`;
  }

  function renderSummary() {
    badgeOk.value.textContent = String(state.summary.okCount ?? 0);
    badgeFail.value.textContent = String(state.summary.failCount ?? 0);
    badgeSkip.value.textContent = String(state.summary.skipCount ?? 0);
    logPathText.textContent = state.logTextPath ? `로그 파일: ${state.logTextPath}` : '로그 파일: -';
    logOpenBtn.disabled = !state.logTextPath;
    btnViewResult.disabled = !hasRunResults();
    btnExport.disabled = !hasRunResults();
    resultSection.style.display = 'none';
  }

  function renderLogs() {
    logBody.innerHTML = '';
    if (!state.logs.length) {
      const tr = document.createElement('tr');
      tr.className = 'empty-row';
      const td = document.createElement('td');
      td.className = 'empty-cell';
      td.colSpan = 3;
      td.textContent = '로그가 없습니다.';
      tr.append(td);
      logBody.append(tr);
      return;
    }
    state.logs.forEach((log) => {
      const tr = document.createElement('tr');
      tr.append(tdText(log.level || ''));
      tr.append(tdText(log.file || ''));
      tr.append(tdText(log.msg || ''));
      logBody.append(tr);
    });
  }

  function getCheckedRvtPaths() {
    return state.rvtList.filter((path) => state.rvtChecked.has(path));
  }

  function updateButtons() {
    const disabled = state.running;
    btnRun.disabled = disabled || !state.selectedParams.length || !state.rvtList.length || !getCheckedRvtPaths().length;
    btnViewResult.disabled = disabled || !hasRunResults();
    btnExport.disabled = disabled || !hasRunResults();
    const modal = buildParamPickerModal;
    if (modal.addBtn) modal.addBtn.disabled = disabled;
    bulkApplyBtn.disabled = disabled || state.selectedParams.length < 2;
  }

  function hasRunResults() {
    if (state.logs.length) return true;
    const summary = state.summary || DEFAULT_SUMMARY;
    return Number(summary.okCount || 0) + Number(summary.failCount || 0) + Number(summary.skipCount || 0) > 0;
  }

  function formatParamGroup(value) {
    const match = state.paramGroups.find(g => g.key === value || g.Key === value || g.id === value || g.Id === value);
    if (match) return formatRevitParamGroupLabel(match.label || match.Label || match.key || match.Key || match.id || match.Id || value || '');
    return formatRevitParamGroupLabel(value || '');
  }

  function getParamGroupValue(group) {
    if (!group) return '';
    return group.key || group.Key || group.id || group.Id || '';
  }

  function getDefaultParamGroupValue() {
    const preferred = state.paramGroups.find((group) => {
      const key = String(group?.key || group?.Key || '').toLowerCase();
      const id = String(group?.id || group?.Id || '').toLowerCase();
      const label = String(group?.label || group?.Label || '').toLowerCase();
      return key === 'pg_data'
        || id === 'pg_data'
        || key.includes('parameter.group:data')
        || label === 'data';
    });
    return getParamGroupValue(preferred || state.paramGroups[0]) || 'PG_DATA';
  }

  function applyBulkSettings() {
    if (state.selectedParams.length < 2) return;
    const template = state.selectedParams[0]?.settings;
    if (!template) return;
    const cloned = cloneSettings(template);
    state.selectedParams.forEach((p, idx) => {
      if (idx === 0 || !p) return;
      p.settings = cloneSettings(cloned);
    });
    renderSelectedParams();
  }

  function cloneSettings(src) {
    return {
      isInstanceBinding: !!src.isInstanceBinding,
      paramGroup: src.paramGroup,
      allowVaryBetweenGroups: !!src.allowVaryBetweenGroups,
      categories: Array.isArray(src.categories) ? src.categories.map(c => ({
        idInt: c.idInt,
        name: c.name,
        path: c.path
      })) : []
    };
  }

  function buildSettingsModal() {
    const overlay = div('sharedparambatch-modal-overlay');
    const modal = div('sharedparambatch-modal');
    const header = div('sharedparambatch-modal__header');
    const title = div('sharedparambatch-modal__title');
    const closeBtn = actionBtn('닫기', closeSettingsModal, 'btn--ghost');
    header.append(title, closeBtn);

    const body = div('sharedparambatch-modal__body');
    const footer = div('sharedparambatch-modal__footer');
    const saveBtn = cardBtn('저장', saveSettings);
    footer.append(saveBtn);

    modal.append(header, body, footer);
    overlay.append(modal);

    overlay.addEventListener('click', (ev) => { if (ev.target === overlay) closeSettingsModal(); });

    buildSettingsModal.overlay = overlay;
    buildSettingsModal.titleEl = title;
    buildSettingsModal.body = body;
    buildSettingsModal.saveBtn = saveBtn;
    buildSettingsModal.currentIndex = -1;

    return overlay;
  }

  function buildSelectHeader(titleText, pickerButton, runButton, exportButton) {
    const header = div('spb-cardHeader');
    const left = div('spb-headerLeft');
    const title = document.createElement('h3');
    title.textContent = titleText;
    left.append(title, pickerButton);
    const right = div('spb-headerRight');
    right.append(runButton, exportButton);
    header.append(left, right);
    return header;
  }

  function buildParamPickerModal() {
    const overlay = div('sharedparambatch-modal-overlay');
    const modal = div('sharedparambatch-modal');
    const header = div('sharedparambatch-modal__header');
    const title = div('sharedparambatch-modal__title');
    title.textContent = '파라미터 선택';
    const closeBtn = actionBtn('닫기', closeParamPicker, 'btn--ghost');
    header.append(title, closeBtn);

    const body = div('sharedparambatch-modal__body');
    const footer = div('sharedparambatch-modal__footer');
    const addBtn = cardBtn('선택 추가', onAddSelectedDefs);
    const closeBtn2 = cardBtn('닫기', closeParamPicker, 'btn--secondary');
    footer.append(addBtn, closeBtn2);

    const groupSelect = document.createElement('select');
    groupSelect.className = 'sharedparambatch-select spb-groupSelect';
    groupSelect.addEventListener('change', () => {
      state.selectedGroup = groupSelect.value;
      renderDefinitionList();
    });

    const searchInput = document.createElement('input');
    searchInput.type = 'search';
    searchInput.className = 'sharedparambatch-input spb-search';
    searchInput.placeholder = '파라미터 검색…';
    searchInput.addEventListener('input', () => {
      state.defSearch = (searchInput.value || '').trim();
      renderDefinitionList();
    });

    const defHeader = div('spb-defHeader');
    defHeader.append(labelSpan('그룹'), groupSelect, searchInput);

    const defList = document.createElement('select');
    defList.className = 'sharedparambatch-def-list spb-defList';
    defList.multiple = true;
    defList.addEventListener('change', () => {
      const next = new Set();
      Array.from(defList.selectedOptions || []).forEach((opt) => {
        const guid = opt.dataset.guid;
        if (guid) next.add(guid);
      });
      state.defSelection = next;
    });

    const defBox = div('sharedparambatch-card');
    defBox.append(defHeader, defList);

    body.append(defBox);
    modal.append(header, body, footer);
    overlay.append(modal);

    overlay.addEventListener('click', (ev) => { if (ev.target === overlay) closeParamPicker(); });

    buildParamPickerModal.overlay = overlay;
    buildParamPickerModal.groupSelect = groupSelect;
    buildParamPickerModal.defList = defList;
    buildParamPickerModal.searchInput = searchInput;
    buildParamPickerModal.addBtn = addBtn;

    return overlay;
  }

  function closeParamPicker() {
    const modal = buildParamPickerModal;
    if (modal.overlay) modal.overlay.classList.remove('is-open');
  }

  function buildRvtModal() {
    const overlay = div('sharedparambatch-modal-overlay');
    const modal = div('sharedparambatch-modal spb-modalLarge');
    const header = div('sharedparambatch-modal__header');
    const title = div('sharedparambatch-modal__title');
    title.textContent = 'RVT 파일 목록';
    const closeBtn = actionBtn('닫기', closeRvtModal, 'btn--ghost');
    header.append(title, closeBtn);

    const body = div('sharedparambatch-modal__body');
    const actions = div('section-actions');
    actions.append(
      cardBtn('추가', () => post('sharedparambatch:browse-rvts', {})),
      cardBtn('폴더 선택', onBrowseFolder),
      cardBtn('선택 삭제', removeSelectedRvts, 'btn--secondary'),
      cardBtn('전체 삭제', clearRvts, 'btn--secondary')
    );
    const { table, tbody, master } = createRvtTable();
    const wrap = div('spb-rvtTableWrap rvt-drop-zone');
    wrap.append(table);
    body.append(actions, wrap);
    attachRvtDropZone(wrap, {
      onDropPaths: (paths) => {
        const added = appendDroppedRvts(paths);
        if (!added) {
          toast('이미 등록된 RVT입니다.', 'warn');
          return;
        }
        renderRvtList();
        toast(`${added}개 RVT를 추가했습니다.`, 'ok');
      },
      onInvalid: () => toast('RVT 파일만 드래그해서 추가할 수 있습니다.', 'warn')
    });

    modal.append(header, body);
    overlay.append(modal);

    overlay.addEventListener('click', (ev) => { if (ev.target === overlay) closeRvtModal(); });

    buildRvtModal.overlay = overlay;
    buildRvtModal.tbody = tbody;
    buildRvtModal.master = master;

    return overlay;
  }

  function buildResultModal() {
    const overlay = div('sharedparambatch-modal-overlay');
    const modal = div('sharedparambatch-modal spb-modalLarge');
    const header = div('sharedparambatch-modal__header');
    const title = div('sharedparambatch-modal__title');
    title.textContent = '결과 보기';
    const closeBtn = actionBtn('닫기', closeResultModal, 'btn--ghost');
    header.append(title, closeBtn);

    const body = div('sharedparambatch-modal__body');
    const footer = div('sharedparambatch-modal__footer');
    const spacer = document.createElement('span');
    const closeFooterBtn = cardBtn('닫기', closeResultModal, 'btn--secondary');
    footer.append(spacer, closeFooterBtn);

    modal.append(header, body, footer);
    overlay.append(modal);

    overlay.addEventListener('click', (ev) => { if (ev.target === overlay) closeResultModal(); });

    return { overlay, body };
  }

  function openResultModal() {
    if (!hasRunResults()) {
      toast('최근 실행 결과가 없습니다.', 'err');
      return;
    }
    if (resultModal.overlay) resultModal.overlay.classList.add('is-open');
  }

  function closeResultModal() {
    if (resultModal.overlay) resultModal.overlay.classList.remove('is-open');
  }

  function openRvtModal() {
    const modal = buildRvtModal;
    if (modal.overlay) modal.overlay.classList.add('is-open');
    renderRvtList();
  }

  function closeRvtModal() {
    const modal = buildRvtModal;
    if (modal.overlay) modal.overlay.classList.remove('is-open');
  }

  function openSettingsModal(index) {
    const modal = buildSettingsModal;
    if (!modal.body || !state.selectedParams[index]) return;

    const param = state.selectedParams[index];
    modal.currentIndex = index;
    modal.title.textContent = `${param.name} 설정`;
    modal.body.innerHTML = '';

    const bindingRow = div('sharedparambatch-modal-row');
    bindingRow.append(subTitle('바인딩'));
    const bindingOpts = div('sharedparambatch-binding');
    const instRadio = radio('spb-binding', '인스턴스', param.settings.isInstanceBinding, (checked) => {
      if (checked) param.settings.isInstanceBinding = true;
      toggleVary();
    });
    const typeRadio = radio('spb-binding', '타입', !param.settings.isInstanceBinding, (checked) => {
      if (checked) param.settings.isInstanceBinding = false;
      toggleVary();
    });
    bindingOpts.append(instRadio.wrap, typeRadio.wrap);
    bindingRow.append(bindingOpts);

    const groupRow = div('sharedparambatch-modal-row');
    groupRow.append(subTitle('파라미터 그룹'));
    const groupSelectEl = document.createElement('select');
    groupSelectEl.className = 'sharedparambatch-select';
    state.paramGroups.forEach((g) => {
      const opt = document.createElement('option');
      opt.value = getParamGroupValue(g);
      opt.textContent = formatParamGroup(opt.value);
      groupSelectEl.append(opt);
    });
    const currentGroup = state.paramGroups.find((g) => {
      const value = getParamGroupValue(g);
      return value === param.settings.paramGroup
        || g.id === param.settings.paramGroup
        || g.Id === param.settings.paramGroup
        || g.key === param.settings.paramGroup
        || g.Key === param.settings.paramGroup;
    });
    groupSelectEl.value = getParamGroupValue(currentGroup) || param.settings.paramGroup || getDefaultParamGroupValue();
    param.settings.paramGroup = groupSelectEl.value || getDefaultParamGroupValue();
    groupSelectEl.addEventListener('change', () => { param.settings.paramGroup = groupSelectEl.value; });
    groupRow.append(groupSelectEl);

    const varyRow = div('sharedparambatch-modal-row');
    const varyWrap = div('sharedparambatch-option');
    const varyChk = document.createElement('input');
    varyChk.type = 'checkbox';
    varyChk.checked = !!param.settings.allowVaryBetweenGroups;
    varyChk.id = 'spb-vary';
    varyChk.addEventListener('change', () => { param.settings.allowVaryBetweenGroups = !!varyChk.checked; });
    const varyLbl = document.createElement('label');
    varyLbl.setAttribute('for', 'spb-vary');
    varyLbl.textContent = '그룹별 값 다르게 허용';
    varyWrap.append(varyChk, varyLbl);
    varyRow.append(varyWrap);

    const categoryRow = div('sharedparambatch-modal-row');
    categoryRow.append(subTitle('카테고리'));
    const categoryActions = div('sharedparambatch-category-actions');
    const btnAll = cardBtn('전체 선택', () => selectAllCategories(param), 'btn--secondary');
    const btnClear = cardBtn('전체 해제', () => clearAllCategories(param), 'btn--secondary');
    const btnClearChildren = cardBtn('서브카테고리 해제', () => clearSubCategories(param), 'btn--secondary');
    categoryActions.append(btnAll, btnClear, btnClearChildren);
    categoryRow.append(categoryActions);

    const categoryFilter = document.createElement('input');
    categoryFilter.type = 'text';
    categoryFilter.className = 'sharedparambatch-input';
    categoryFilter.placeholder = '카테고리 검색…';
    categoryFilter.value = state.categoryFilterText || '';
    categoryFilter.addEventListener('input', () => {
      state.categoryFilterText = categoryFilter.value || '';
      const modalNow = buildSettingsModal;
      if (modalNow.body) renderCategoryTree(modalNow.body.querySelector('.sharedparambatch-category-tree'), param);
    });
    categoryRow.append(categoryFilter);

    const presetRow = div('sharedparambatch-preset-row');
    const presetSelect = document.createElement('select');
    presetSelect.className = 'sharedparambatch-select sharedparambatch-preset-select';
    const presets = loadCategoryPresets();
    const defaultOpt = document.createElement('option');
    defaultOpt.value = '';
    defaultOpt.textContent = '카테고리 프리셋 선택';
    presetSelect.append(defaultOpt);
    presets.map((p) => p.name).sort().forEach((name) => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      presetSelect.append(opt);
    });

    const presetNameInput = document.createElement('input');
    presetNameInput.type = 'text';
    presetNameInput.className = 'sharedparambatch-input sharedparambatch-preset-name';
    presetNameInput.placeholder = '프리셋 이름';

    const btnSavePreset = cardBtn('저장', () => {
      const nextName = (presetNameInput.value || presetSelect.value || '').trim();
      if (!nextName) {
        toast('저장할 프리셋 이름을 입력해 주세요.', 'err');
        return;
      }
      saveCategoryPreset(nextName, param.settings.categories || []);
      openSettingsModal(modal.currentIndex);
      toast(`프리셋 저장 완료: ${nextName}`, 'ok');
    }, 'btn--secondary');

    const btnLoadPreset = cardBtn('불러오기', () => {
      if (!presetSelect.value) {
        toast('불러올 프리셋을 선택해 주세요.', 'err');
        return;
      }
      applyCategoryPreset(param, presetSelect.value);
      const modalNow = buildSettingsModal;
      if (modalNow.body) renderCategoryTree(modalNow.body.querySelector('.sharedparambatch-category-tree'), param);
      toast(`프리셋 적용 완료: ${presetSelect.value}`, 'ok');
    }, 'btn--secondary');

    const btnDeletePreset = cardBtn('삭제', () => {
      const presetName = (presetSelect.value || '').trim();
      if (!presetName) {
        toast('삭제할 프리셋을 선택해 주세요.', 'err');
        return;
      }
      deleteCategoryPreset(presetName);
      openSettingsModal(modal.currentIndex);
      toast(`프리셋 삭제 완료: ${presetName}`, 'ok');
    }, 'btn--secondary');

    presetSelect.addEventListener('change', () => {
      presetNameInput.value = presetSelect.value || '';
    });

    presetRow.append(presetSelect, presetNameInput, btnSavePreset, btnLoadPreset, btnDeletePreset);
    categoryRow.append(presetRow);

    const treeWrap = div('sharedparambatch-category-tree');
    renderCategoryTree(treeWrap, param);
    categoryRow.append(treeWrap);

    modal.body.append(bindingRow, groupRow, varyRow, categoryRow);
    toggleVary();

    if (modal.overlay) modal.overlay.classList.add('is-open');

    function toggleVary() {
      varyChk.disabled = !param.settings.isInstanceBinding;
      if (!param.settings.isInstanceBinding) {
        param.settings.allowVaryBetweenGroups = false;
        varyChk.checked = false;
      }
    }
  }

  function closeSettingsModal() {
    const modal = buildSettingsModal;
    if (modal.overlay) modal.overlay.classList.remove('is-open');
    modal.currentIndex = -1;
  }

  function saveSettings() {
    closeSettingsModal();
    renderSelectedParams();
  }

  function renderCategoryTree(container, param) {
    container.innerHTML = '';
    if (!state.categoryTree.length) {
      container.textContent = '카테고리 정보가 없습니다.';
      return;
    }
    const selected = new Set(param.settings.categories.map(c => c.path || c.name));
    const keyword = (state.categoryFilterText || '').trim().toLowerCase();
    const list = document.createElement('ul');
    list.className = 'sharedparambatch-tree';
    state.categoryTree.forEach((node) => {
      const built = buildCategoryNode(node, keyword);
      if (built) list.append(built);
    });
    container.append(list);

    function buildCategoryNode(node, keyword) {
      const nameText = String(node.name || '').toLowerCase();
      const childMatches = [];
      if (node.children && node.children.length) {
        node.children.forEach((child) => {
          const cnode = buildCategoryNode(child, keyword);
          if (cnode) childMatches.push(cnode);
        });
      }
      const selfMatch = !keyword || nameText.includes(keyword);
      if (!selfMatch && childMatches.length === 0) return null;

      const li = document.createElement('li');
      const row = div('sharedparambatch-tree-row');
      if (node.isBindable) {
        const chk = document.createElement('input');
        chk.type = 'checkbox';
        chk.checked = selected.has(node.path || node.name);
        chk.addEventListener('change', () => {
          if (chk.checked) {
            addCategory(param, node);
          } else {
            removeCategory(param, node);
          }
        });
        row.append(chk);
      }
      const label = document.createElement('span');
      label.textContent = node.name || '';
      row.append(label);
      li.append(row);

      if (childMatches.length) {
        const childList = document.createElement('ul');
        childList.className = 'sharedparambatch-tree';
        childMatches.forEach((child) => childList.append(child));
        li.append(childList);
      }
      return li;
    }
  }

  function addCategory(param, node) {
    if (!param.settings.categories.find(c => c.path === node.path)) {
      param.settings.categories.push({
        idInt: node.idInt,
        name: node.name,
        path: node.path
      });
    }
  }

  function removeCategory(param, node) {
    param.settings.categories = param.settings.categories.filter(c => c.path !== node.path);
  }

  function selectAllCategories(param) {
    const all = [];
    collectBindable(state.categoryTree, all);
    param.settings.categories = all.map((n) => ({ idInt: n.idInt, name: n.name, path: n.path }));
    const modal = buildSettingsModal;
    if (modal.body) renderCategoryTree(modal.body.querySelector('.sharedparambatch-category-tree'), param);
  }

  function clearAllCategories(param) {
    param.settings.categories = [];
    const modal = buildSettingsModal;
    if (modal.body) renderCategoryTree(modal.body.querySelector('.sharedparambatch-category-tree'), param);
  }

  function clearSubCategories(param) {
    const targets = new Set();
    collectBindableWithDepth(state.categoryTree, targets, 0);
    param.settings.categories = param.settings.categories.filter(c => !targets.has(c.path || c.name));
    const modal = buildSettingsModal;
    if (modal.body) renderCategoryTree(modal.body.querySelector('.sharedparambatch-category-tree'), param);
  }

  function collectBindable(nodes, out) {
    nodes.forEach((n) => {
      if (n.isBindable) out.push(n);
      if (n.children && n.children.length) collectBindable(n.children, out);
    });
  }

  function collectBindableWithDepth(nodes, out, depth) {
    nodes.forEach((n) => {
      if (depth >= 1 && n.isBindable) out.add(n.path || n.name);
      if (n.children && n.children.length) collectBindableWithDepth(n.children, out, depth + 1);
    });
  }


  function loadCategoryPresets() {
    try {
      const raw = localStorage.getItem(CATEGORY_PRESET_KEY) || '[]';
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        return parsed
          .filter((row) => row && typeof row.name === 'string')
          .map((row) => ({ name: row.name, ids: Array.isArray(row.ids) ? row.ids : [] }));
      }
      if (parsed && typeof parsed === 'object') {
        return Object.keys(parsed).map((name) => {
          const old = parsed[name];
          if (Array.isArray(old)) {
            return {
              name,
              ids: old.map((item) => item?.idInt).filter((id) => Number.isFinite(Number(id)))
            };
          }
          return {
            name,
            ids: Array.isArray(old?.ids) ? old.ids : []
          };
        });
      }
      return [];
    } catch {
      return [];
    }
  }

  function saveCategoryPresets(presets) {
    const rows = Array.isArray(presets) ? presets : [];
    try { localStorage.setItem(CATEGORY_PRESET_KEY, JSON.stringify(rows)); } catch {}
  }

  function saveCategoryPreset(name, categories) {
    const key = String(name || '').trim();
    if (!key) return;
    const presets = loadCategoryPresets();
    const next = {
      name: key,
      ids: (categories || []).map((c) => c.idInt).filter((id) => Number.isFinite(Number(id)))
    };
    const idx = presets.findIndex((p) => p.name === key);
    if (idx >= 0) presets[idx] = next;
    else presets.push(next);
    saveCategoryPresets(presets);
  }

  function applyCategoryPreset(param, name) {
    const presets = loadCategoryPresets();
    const preset = presets.find((p) => p.name === name);
    const ids = Array.isArray(preset?.ids) ? preset.ids : [];
    const picked = [];
    const idSet = new Set(ids.map((id) => Number(id)));
    collectBindable(state.categoryTree, picked);
    param.settings.categories = picked
      .filter((cat) => idSet.has(Number(cat.idInt)))
      .map((cat) => ({ idInt: cat.idInt, name: cat.name, path: cat.path }));
  }

  function deleteCategoryPreset(name) {
    const key = String(name || '').trim();
    if (!key) return;
    const presets = loadCategoryPresets();
    saveCategoryPresets(presets.filter((p) => p.name !== key));
  }

  function sectionHeader(title, buttons) {
    const wrap = div('section-header');
    const h = document.createElement('h3');
    h.textContent = title;
    wrap.append(h);
    if (Array.isArray(buttons) && buttons.length) {
      const actions = div('section-actions');
      buttons.forEach(btn => actions.append(btn));
      wrap.append(actions);
    }
    return wrap;
  }

  function subTitle(text) {
    const el = document.createElement('div');
    el.className = 'sharedparambatch-subtitle';
    el.textContent = text;
    return el;
  }

  function summaryBadge(label, value) {
    const wrap = div('sharedparambatch-badge');
    const span = document.createElement('span');
    span.textContent = label;
    const val = document.createElement('strong');
    val.textContent = value;
    wrap.append(span, val);
    return { wrap, value: val };
  }

  function cardBtn(text, onClick, extraClass = '') {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `btn ${extraClass}`.trim();
    btn.textContent = text;
    if (typeof onClick === 'function') btn.addEventListener('click', onClick);
    return btn;
  }

  function actionBtn(text, onClick, extraClass = '') {
    return cardBtn(text, onClick, `btn--small ${extraClass}`.trim());
  }

  function tdText(value) {
    const td = document.createElement('td');
    td.textContent = value == null ? '' : String(value);
    return td;
  }

  function labelSpan(text) {
    const span = document.createElement('span');
    span.className = 'sharedparambatch-label';
    span.textContent = text;
    return span;
  }

  function radio(name, label, checked, onChange) {
    const wrap = div('sharedparambatch-radio');
    const input = document.createElement('input');
    input.type = 'radio';
    input.name = name;
    input.checked = checked;
    const span = document.createElement('span');
    span.textContent = label;
    input.addEventListener('change', () => { if (input.checked && onChange) onChange(true); });
    wrap.append(input, span);
    return { wrap, input };
  }
}
