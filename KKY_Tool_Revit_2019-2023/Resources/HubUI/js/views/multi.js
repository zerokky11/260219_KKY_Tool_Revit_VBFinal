import { clear, div, toast, setBusy, showExcelSavedDialog, chooseExcelMode } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { post, onHost } from '../core/bridge.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js';

const FEATURE_META = {
    connector: { label: '커넥터 파라미터 연속성 검토', desc: 'BQC 핵심 검토 · 연결 객체들의 파라미터 값 연속성을 확인합니다.', requiresSharedParams: false },
    guid: {
        label: '공유파라미터 GUID 검토', desc: '프로젝트/패밀리 내 공유 파라미터 GUID 검토', requiresSharedParams: true },
  familylink: { label: '패밀리 공유파라미터 연동 검토', desc: '복합 패밀리의 하위 패밀리 파라미터 연동 상태를 검토합니다.', requiresSharedParams: true },
  points: { label: 'Point 추출', desc: 'Project/Survey Point 좌표 추출', requiresSharedParams: false }
};
const FEATURE_KEYS = Object.keys(FEATURE_META);
const BQC_FEATURE_KEYS = ['connector'];
const UTILITY_FEATURE_KEYS = FEATURE_KEYS.filter((key) => !BQC_FEATURE_KEYS.includes(key));
const COMMON_OPTIONS_KEY = 'kky.hub.commonOptions';
const GROUP_FILTER_KEY = 'kky.hub.multiGroupFilter';
const MULTI_MODE_KEY = 'kky.hub.multiMode';
const GROUPS = [
  { id: 'all', label: '전체' },
  { id: 'bqc', label: '납품 시 BQC 검토' },
  { id: 'utility', label: '유틸리티' }
];

export function renderMulti(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);
  const top = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar');
  if (top) top.classList.add('hub-topbar');

  const multiMode = normalizeMultiMode(getMultiMode());

  const state = {
    rvtList: [],
    rvtChecked: new Set(),
    busy: false,
    common: createConfigState({
      extraParams: '',
      targetFilter: '',
      excludeEndDummy: false
    }),
    features: {
      connector: createFeatureState({ tol: 1.0, unit: 'inch', param: 'Comments', paramItems: ['Comments'] }),
      guid: createFeatureState({ includeFamily: false, includeAnnotation: false }),
      familylink: createFeatureState({ targetsText: '', selectedTargets: [], targets: [] }),
      points: createFeatureState({ unit: 'ft' })
    },
    results: {},
    sharedParamStatus: null,
    sharedParamItems: [],
    connectorParamItems: [],
    ui: {
      modalOpen: false,
      activeFeatureKey: '',
      activeFeatureTitle: '',
      panels: {},
      controls: {},
      lastProgressPct: 0,
      runCompleted: false,
      commonSummaryEl: null,
      sharedParamBanner: null,
      runSummaryTitle: null,
      runSummaryDetail: null,
      runSharedParamHint: null,
      runProgressText: null,
      runProgressDetail: null,
      runProgressFill: null,
      actionCommonSummaryEl: null,
      selectedTableBody: null,
      selectedRows: new Map(),
      groupFilter: 'all',
      multiMode: multiMode,
      isRvtListExpanded: false,
      reviewSummaryData: null,
      currentDocButtons: [],
      openMultiButtons: [],
      modalRunButtons: [],
      resultResetButtons: [],
      reviewSummaryByMode: { bqc: null, utility: null },
      lastRunAtByMode: { bqc: null, utility: null },
      bqcRecentCaption: null,
      bqcRecentHint: null,
      bqcRecentTableBody: null,
      bqcRecentEmpty: null,
      bqcRecentOpenBtn: null,
      bqcRecentExportBtn: null,
      modalFeatureList: null,
      modalFeatureCount: null,
      modalRecentCaption: null,
      modalRecentHint: null,
      modalRecentTableBody: null,
      modalRecentEmpty: null,
      lastRunAt: null
    }
  };

  FEATURE_KEYS.forEach((k) => {
    state.results[k] = { count: 0, stale: true, hasRun: false };
  });

  const page = div('feature-shell multi-page HubShell');
  const hasLocalCommonOptions = loadCommonOptionsFromStorage();
  const header = div('feature-header multi-header');
  header.innerHTML = buildHeaderHtml(state.ui.multiMode);
  page.append(header);

  const layout = div(`multi-layout HubBody ${state.ui.multiMode === 'utility' ? 'multi-layout--utility' : 'multi-layout--bqc'}`);
  const mainCol = div('multi-workspace multi-workspace--main');
  const sideCol = div('multi-workspace multi-workspace--sidebar');

  const group1 = buildGroupSection('납품 시 BQC 검토', '가장 많이 사용하는 핵심 검토를 먼저 선택하세요.', 'bqc');
  const group3 = buildGroupSection('유틸리티', 'PMS / GUID / 패밀리 연동 / Point 추출 / Project Parameter', 'utility');
  group3.section.id = 'utilities';

  buildGroup1Options();
  group1.section.append(buildToggleRow('connector', buildConnectorConfig()));
  group3.section.append(buildPmsWorkflowRow());
  group3.section.append(buildToggleRow('guid', buildGuidConfig()));
  group3.section.append(buildToggleRow('familylink', buildFamilyLinkConfig()));
  group3.section.append(buildToggleRow('points', buildPointsConfig()));
  group3.section.append(buildSharedParamBatchRow());
  if (state.ui.multiMode === 'bqc' && group1.section.querySelectorAll('.feature-row').length === 0) {
    const empty = div('feature-note');
    empty.textContent = '등록된 BQC 검토 기능이 없습니다.';
    group1.section.append(empty);
  }

  state.ui.groupFilter = state.ui.multiMode;
  saveGroupFilter(state.ui.multiMode);
  if (state.ui.multiMode === 'bqc') {
    group1.wrap.classList.add('multi-section--bqc-hero');
    const group1Head = group1.wrap.querySelector('.multi-section-title');
    if (group1Head) {
      group1Head.remove();
      group1.wrap.classList.add('multi-section--headerless');
    }
    const bqcSidebar = div('multi-sidebar-stack multi-sidebar-stack--sticky');
    mainCol.append(group1.wrap);
    bqcSidebar.append(
      buildExecutionActionPanel({ mode: 'bqc' }),
      buildSelectedFeaturesSection({
        title: '선택된 기능',
        showCurrentButton: false,
        sectionClass: 'selected-panel--sidebar'
      })
    );
    sideCol.append(bqcSidebar);
    layout.append(mainCol, sideCol);
  } else {
    group3.wrap.classList.add('multi-section--utility');
    const group3Head = group3.wrap.querySelector('.multi-section-title');
    if (group3Head) {
      group3Head.remove();
      group3.wrap.classList.add('multi-section--headerless');
    }
    const utilitySidebar = div('multi-sidebar-stack multi-sidebar-stack--sticky');
    utilitySidebar.append(
      buildExecutionActionPanel({ mode: 'utility' }),
      buildSelectedFeaturesSection({
        title: '선택된 기능',
        showCurrentButton: false,
        sectionClass: 'selected-panel--sidebar'
      })
    );
    mainCol.append(group3.wrap);
    sideCol.append(utilitySidebar);
    layout.append(mainCol, sideCol);
  }
  page.append(layout);
  page.append(buildSettingsModal());
  page.append(buildReviewSummaryModal());
  page.append(buildRvtExpandedModal());
  target.append(page);

  onHost('hub:rvt-picked', (payload) => {
    const paths = Array.isArray(payload?.paths) ? payload.paths : [];
    if (!paths.length) return;
    let changed = false;
    paths.forEach((p) => {
      if (!state.rvtList.includes(p)) {
        state.rvtList.push(p);
        state.rvtChecked.add(p);
        changed = true;
      }
    });
    if (changed) {
      markAllStale();
      renderRvtList();
    }
  });

  onHost('hub:multi-progress', (payload) => {
    const basePct = Number(payload?.percent);
    const altPct = Number(payload?.phaseProgress);
    const hasBase = Number.isFinite(basePct);
    const hasAlt = Number.isFinite(altPct);
    const pctValue = hasBase ? basePct : (hasAlt ? altPct * 100 : state.ui.lastProgressPct);
    const pct = Math.max(0, Math.min(100, pctValue));
    state.ui.lastProgressPct = pct;
    const phase = String(payload?.phase || payload?.Phase || '').toLowerCase();
    ProgressDialog.show(payload?.title || '납품시 BQC 검토', payload?.message || '');
    ProgressDialog.update(pct, payload?.message || '', payload?.detail || '');
    updateRunProgress(pct, payload?.message || '', payload?.detail || '');
    if (phase === 'done' || pct >= 100) {
      ProgressDialog.hide();
    }
  });

  onHost('hub:multi-done', (payload) => {
    setBusyState(false);
    ProgressDialog.update(100, '완료', '검토가 완료되었습니다.');
    ProgressDialog.hide();
    updateRunProgress(100, '완료', '검토가 완료되었습니다.');
    updateResultSummary(payload?.summary || {});
    state.ui.runCompleted = true;
    updateRunActionLabel();
  });

  onHost('multi:review-summary', (payload) => {
    ProgressDialog.hide();
    showReviewSummary(payload || {});
  });

  onHost('sharedparam:list', (payload) => {
    const ok = payload?.ok !== false;
    state.sharedParamItems = ok && Array.isArray(payload?.items) ? payload.items : [];
    if (buildFamilyLinkConfig.renderList) buildFamilyLinkConfig.renderList(payload);
  });

  onHost('connector:param-list:done', (payload) => {
    const ok = payload?.ok !== false;
    state.connectorParamItems = ok ? normalizeConnectorParamItems(payload) : [];
    if (buildConnectorConfig.renderList) buildConnectorConfig.renderList(payload);
  });

  onHost('hub:multi-error', (payload) => {
    setBusyState(false);
    ProgressDialog.hide();
    updateRunProgress(0, '오류 발생', payload?.message || '');
    toast(payload?.message || '배치 검토 중 오류가 발생했습니다.', 'err');
    state.ui.runCompleted = false;
    updateRunActionLabel();
  });

  onHost('hub:multi-exported', (payload) => {
    setBusyState(false);
    ProgressDialog.hide();
    state.ui.lastProgressPct = 0;
    const path = payload?.path;
    if (path) {
      requestAnimationFrame(() => {
        showExcelSavedDialog('엑셀 저장 완료', path, (p) => post('excel:open', { path: p }));
      });
    } else {
      toast(payload?.message || '엑셀 저장에 실패했습니다.', 'err');
    }
  });

  onHost('sharedparam:status', (payload) => {
    state.sharedParamStatus = payload || {};
    updateSharedParamBanner();
    updateRunSummary();
  });

  requestConnectorParamList('render');

  window.addEventListener('keydown', (ev) => {
    if (ev.key === 'Escape' && state.ui.isRvtListExpanded) {
      ev.preventDefault();
      closeExpandedRvtModal();
    }
  });

  function handleAddRvt() {
    post('hub:pick-rvt', {});
  }

  function handleRemoveSelected() {
    if (state.rvtChecked.size === 0) return;
    state.rvtList = state.rvtList.filter((p) => !state.rvtChecked.has(p));
    state.rvtChecked.clear();
    markAllStale();
    if (buildRvtSection.render) buildRvtSection.render();
  }

  function handleClearList() {
    if (state.rvtList.length === 0) return;
    const confirmed = window.confirm('RVT 목록을 모두 삭제할까요?');
    if (!confirmed) return;
    state.rvtList = [];
    state.rvtChecked.clear();
    markAllStale();
    if (buildRvtSection.render) buildRvtSection.render();
  }

  function requestSharedParamList(context) {
    post('sharedparam:list', { source: 'multi', context: context || '' });
  }

  function requestConnectorParamList(context) {
    post('connector:param-list', { source: 'multi', context: context || '' });
  }

  function normalizeMultiMode(value) {
    if (value === 'utility') return 'utility';
    return 'bqc';
  }

  function getMultiMode() {
    try {
      return localStorage.getItem(MULTI_MODE_KEY) || 'bqc';
    } catch {
      return 'bqc';
    }
  }

  function buildHeaderHtml(mode) {
    if (mode === 'utility') {
      return `
    <div class="feature-heading">
      <span class="feature-kicker">Utilities</span>
      <h2 class="feature-title">유틸리티</h2>
      <p class="feature-sub">납품시 BQC 검토 유틸리티 도구 모음입니다.</p>
    </div>`;
    }
    return `
    <div class="feature-heading">
      <span class="feature-kicker">Multi RVT Hub</span>
      <h2 class="feature-title">납품시 BQC 검토</h2>
      <p class="feature-sub">기능 영역을 눌러 선택한 뒤 검토 시작을 눌러주세요. 여러 기능을 한 번에 선택해 같은 RVT 목록으로 순차 검토할 수 있습니다.</p>
    </div>`;
  }

  function buildGroupSection(title, desc, groupId) {
    const wrap = div('multi-section');
    if (groupId) wrap.dataset.group = groupId;
    wrap.style.borderRadius = '20px';
    wrap.style.border = '1px solid var(--border-accent-soft)';
    wrap.style.background = 'var(--surface-note)';
    wrap.style.padding = '14px';
    const head = div('multi-section-title');
    head.innerHTML = `<h3>${title}</h3><span class="feature-note">${desc}</span>`;
    head.style.marginBottom = '12px';
    head.style.padding = '10px 12px';
    head.style.borderRadius = '14px';
    head.style.border = '1px solid var(--border-soft)';
    head.style.background = 'var(--surface-help)';
    wrap.append(head);
    return { wrap, section: wrap };
  }

  function buildFeatureChecklistGuide() {
    const guide = div('feature-row__summary');
    guide.style.display = 'grid';
    guide.style.gap = '8px';
    guide.style.marginBottom = '12px';
    guide.style.padding = '16px 18px';
    guide.style.borderRadius = '18px';
    guide.style.border = '1px solid var(--border-accent-soft)';
    guide.style.background = 'var(--surface-help)';
    guide.style.boxShadow = 'var(--shadow-accent-soft)';
    const kicker = document.createElement('span');
    kicker.className = 'chip chip--info';
    kicker.textContent = '선택 후 실행';
    kicker.style.width = 'fit-content';
    const title = document.createElement('strong');
    title.textContent = '아래 기능 영역을 눌러 실행할 검토를 선택하세요.';
    const sub = document.createElement('span');
    sub.textContent = '여러 기능을 한 번에 선택해 같은 RVT 목록으로 순차 검토할 수 있습니다. 별도 워크플로우 기능은 카드를 누르면 해당 화면으로 바로 전환됩니다.';
    guide.append(kicker, title, sub);
    return guide;
  }

  function buildGroupFilter() {
    const wrap = div('group-filter');
    const stored = getGroupFilter();
    state.ui.groupFilter = stored;
    const mode = state.ui.multiMode || 'bqc';
    const allowed = mode === 'utility' ? ['utility'] : ['bqc'];
    GROUPS.forEach((group) => {
      if (group.id !== 'all' && !allowed.includes(group.id)) return;
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'group-filter__btn';
      btn.textContent = group.label;
      btn.dataset.group = group.id;
      btn.classList.toggle('is-active', group.id === stored);
      btn.addEventListener('click', () => {
        state.ui.groupFilter = group.id;
        saveGroupFilter(group.id);
        wrap.querySelectorAll('.group-filter__btn').forEach((el) => {
          el.classList.toggle('is-active', el.dataset.group === group.id);
        });
        renderGroupVisibility();
      });
      wrap.append(btn);
    });
    return wrap;
  }

  onHost('commonoptions:loaded', (payload) => {
    if (hasLocalCommonOptions) return;
    applyCommonOptionsFromStorage(payload);
    persistCommonOptions(state.common.configCommitted, { skipHost: true });
  });

  function buildGroup1Options() {
    const panel = div('group-common-mini');
    panel.style.padding = '4px 8px';
    panel.style.borderRadius = '12px';
    panel.style.border = '1px solid var(--border-soft)';
    panel.style.background = 'var(--surface-subtle)';
    panel.style.marginBottom = '6px';
    const header = div('group-common-mini__header');
    header.style.display = 'flex';
    header.style.alignItems = 'center';
    header.style.justifyContent = 'space-between';
    header.style.gap = '10px';
    const title = document.createElement('h4');
    title.textContent = '공통 옵션';
    title.style.margin = '0';
    title.style.fontSize = '12px';
    const settingsBtn = document.createElement('button');
    settingsBtn.type = 'button';
    settingsBtn.className = 'btn btn--secondary';
    settingsBtn.textContent = '설정';
    settingsBtn.addEventListener('click', () => openSettings('common', '그룹 공통 옵션'));
    header.append(title, settingsBtn);

    const summary = div('group-common-mini__summary');
    state.ui.commonSummaryEl = summary;
    summary.textContent = buildCommonSummary();
    summary.style.fontSize = '11px';
    summary.style.lineHeight = '1.35';
    summary.style.color = 'var(--muted,#64748b)';
    panel.append(header, summary);

    const fields = div('multi-config is-open');
    const extra = makeField('추가 Parameter 값 추출', 'extra', 'PM1, PM2', 'textarea');
    const filter = makeField('검토 대상 필터', 'filter', 'ex) PM1=Value;PM2=Value2', 'text');
    const exclude = makeCheckboxField('End_ + Dummy 패밀리 제외');

    const draft = state.common.configDraft;
    extra.input.value = draft.extraParams;
    filter.input.value = draft.targetFilter;
    exclude.input.checked = draft.excludeEndDummy;

    extra.input.addEventListener('change', () => {
      state.common.configDraft.extraParams = extra.input.value;
      markCommonDirty();
      updateCommonSummary(summary);
    });
    filter.input.addEventListener('change', () => {
      state.common.configDraft.targetFilter = filter.input.value;
      markCommonDirty();
      updateCommonSummary(summary);
    });
    exclude.input.addEventListener('change', () => {
      state.common.configDraft.excludeEndDummy = exclude.input.checked;
      markCommonDirty();
      updateCommonSummary(summary);
    });

    fields.append(extra.field, filter.field, exclude.field);
    fields.append(buildFilterExamples());
    fields.classList.add('settings-panel', 'is-open');
    state.ui.panels.common = fields;
    state.ui.controls.common = { extra, filter, exclude };

    return panel;
  }

  function buildToggleRow(key, config) {
    const meta = FEATURE_META[key] || {};
    const row = div('feature-row');
    row.dataset.key = key;
    row.style.borderRadius = '16px';
    row.style.border = '1px solid var(--border-soft)';
    row.style.background = 'var(--surface-elevated)';
    row.style.boxShadow = 'var(--shadow-soft)';
    const header = div('feature-row__header');
    const toggle = document.createElement('input');
    toggle.type = 'checkbox';
    toggle.className = 'feature-toggle';
    const setFeatureEnabled = (enabled, shouldOpenSettings) => {
      const feature = state.features[key];
      const nextEnabled = !!enabled;
      feature.enabled = nextEnabled;
      toggle.checked = nextEnabled;
      if (!nextEnabled) {
        feature.applied = false;
        feature.dirty = false;
        resetDraftFromCommitted(key);
      } else if (!feature.enabled || shouldOpenSettings) {
        feature.applied = false;
        feature.dirty = false;
        if (shouldOpenSettings) openSettings(key, meta.label);
      }
      row.classList.toggle('is-active', nextEnabled);
      markStale(key);
      updateRunSummary();
      refreshConnectorFeatureSummary();
      renderSelectedFeatures();
    };
    toggle.addEventListener('change', () => {
      setFeatureEnabled(toggle.checked, toggle.checked);
    });

    const metaWrap = div('feature-row__left');
    const textWrap = div('feature-row__text');
    const metaTitle = document.createElement('strong');
    metaTitle.textContent = meta.label || key;
    const metaDesc = document.createElement('span');
    metaDesc.textContent = meta.desc || '';
    textWrap.append(metaTitle, metaDesc);
    metaWrap.append(toggle, textWrap);

    const right = div('feature-row__right');
    if (key === 'connector') {
      const badge = document.createElement('span');
      badge.className = 'chip chip--info';
      badge.textContent = 'BQC 핵심';
      right.append(badge);
      row.classList.add('feature-row--workflow', 'feature-row--connector-hero');
      row.style.border = '1px solid var(--border-accent-soft)';
      row.style.boxShadow = 'var(--shadow-accent-soft)';
      row.style.background = 'var(--surface-note)';
    }

    header.append(metaWrap, right);
    row.append(header);
    row.classList.add('is-clickable');
    row.addEventListener('click', (ev) => {
      if (ev.target.closest('button, input, select, textarea, a, label')) return;
      setFeatureEnabled(!state.features[key].enabled, !state.features[key].enabled);
    });

    if (key === 'connector') {
      const summary = div('feature-row__summary');
      summary.style.display = 'grid';
      summary.style.gap = '6px';
      summary.style.padding = '12px 14px';
      summary.style.borderRadius = '14px';
      summary.style.border = '1px solid var(--border-accent-soft)';
      summary.style.background = 'var(--surface-help)';
      const top = document.createElement('strong');
      const sub = document.createElement('span');
      top.textContent = '체크 후 옵션을 열어 공유 파라미터 목록에서 검토 대상을 선택하세요.';
      sub.textContent = '아직 검토 파라미터가 지정되지 않았습니다.';
      summary.append(top, sub);
      row.append(summary);
      state.ui.connectorHeroSummary = { row, top, sub };
      refreshConnectorFeatureSummary();
    }

    config.title = meta.label || key;
    config.key = key;
    config.panel.classList.add('settings-panel', 'is-open');
    state.ui.panels[key] = config.panel;
    state.ui.controls[key] = config.controls || {};
    row.classList.toggle('is-active', !!state.features[key]?.enabled);
    return row;
  }

  function buildPmsWorkflowRow() {
    return buildWorkflowLaunchRow({
      iconLabel: 'PMS',
      title: 'PMS 검토',
      desc: 'Segment ↔ PMS 매핑 및 사이즈 검토',
      summary: '추출 → PMS 등록 → 매핑 준비 → 비교 실행 → 결과 내보내기',
      route: 'segmentpms'
    });
  }

  function buildSharedParamBatchRow() {
    return buildWorkflowLaunchRow({
      iconLabel: 'SP',
      title: 'Project Parameter 추가 (Project/Shared)',
      desc: 'Project/Shared 파라미터를 여러 RVT에 일괄 추가/바인딩합니다.',
      summary: '파라미터 선택 → 바인딩 설정 → RVT 실행 → 로그/엑셀',
      route: 'sharedparambatch'
    });
  }

  function buildWorkflowLaunchRow(options = {}) {
    const row = div('feature-row feature-row--workflow');
    row.dataset.route = options.route || '';
    const header = div('feature-row__header');
    const left = div('feature-row__left');
    const icon = document.createElement('span');
    icon.className = 'feature-row__icon';
    icon.textContent = options.iconLabel || 'WF';
    const text = div('feature-row__text');
    const title = document.createElement('strong');
    title.textContent = options.title || '';
    const desc = document.createElement('span');
    desc.textContent = options.desc || '';
    text.append(title, desc);
    left.append(icon, text);

    const right = div('feature-row__right');
    const chip = document.createElement('span');
    chip.className = 'chip chip--info';
    chip.textContent = '별도 워크플로우';
    chip.title = '클릭하면 별도 워크플로우 화면으로 이동합니다.';
    right.append(chip);

    const summary = div('feature-row__summary');
    summary.textContent = options.summary || '';
    header.append(left, right);
    row.append(header, summary);
    row.classList.add('is-clickable');
    row.setAttribute('role', 'button');
    row.setAttribute('aria-label', `${options.title || '워크플로우'} 열기`);
    row.tabIndex = 0;
    const navigate = () => navigateToFeatureRoute(options.route || '');
    row.addEventListener('click', () => {
      navigate();
    });
    chip.addEventListener('click', (ev) => {
      ev.stopPropagation();
      navigate();
    });
    row.addEventListener('keydown', (ev) => {
      if (ev.key !== 'Enter' && ev.key !== ' ') return;
      ev.preventDefault();
      navigate();
    });
    return row;
  }

  function navigateToFeatureRoute(route) {
    const target = String(route || '').replace(/^#/, '');
    if (!target) return;
    const current = (location.hash || '').replace('#', '');
    if (current === target) {
      window.dispatchEvent(new Event('hashchange'));
      return;
    }
    location.hash = `#${target}`;
    window.setTimeout(() => {
      if ((location.hash || '').replace('#', '') === target) {
        window.dispatchEvent(new Event('hashchange'));
      }
    }, 0);
  }

  function buildConnectorConfig() {
    const panel = div('multi-config');
    panel.style.display = 'flex';
    panel.style.flexDirection = 'column';
    panel.style.alignItems = 'stretch';
    panel.style.justifyContent = 'flex-start';
    panel.style.alignContent = 'stretch';
    panel.style.gap = '10px';
    panel.style.width = '100%';
    panel.style.maxWidth = 'none';
    panel.style.minWidth = '0';
    panel.style.boxSizing = 'border-box';

    const tol = makeField('허용범위', 'tol', '', 'number');
    tol.input.value = state.features.connector.configDraft.tol;
    tol.input.min = '0';
    tol.input.step = '0.01';
    tol.input.style.fontWeight = '600';
    tol.input.addEventListener('change', () => {
      state.features.connector.configDraft.tol = parseFloat(tol.input.value || '1') || 1;
      markFeatureDirty('connector');
      refreshConnectorFeatureSummary();
    });

    const unit = makeSelectField('단위', [
      { value: 'inch', label: 'inch' },
      { value: 'mm', label: 'mm' }
    ]);
    unit.select.value = state.features.connector.configDraft.unit;
    unit.select.style.fontWeight = '600';
    unit.select.addEventListener('change', () => {
      state.features.connector.configDraft.unit = unit.select.value;
      markFeatureDirty('connector');
      refreshConnectorFeatureSummary();
    });

    const basicsCard = div('feature-row__summary');
    basicsCard.style.display = 'grid';
    basicsCard.style.gap = '10px';
    basicsCard.style.padding = '12px';
    basicsCard.style.borderRadius = '18px';
    basicsCard.style.border = '1px solid var(--border-accent-soft)';
    basicsCard.style.background = 'var(--surface-elevated)';
    basicsCard.style.boxShadow = 'var(--shadow-soft)';
    basicsCard.style.width = '100%';
    basicsCard.style.boxSizing = 'border-box';

    const basicsTitle = document.createElement('strong');
    basicsTitle.textContent = '기본 설정';
    basicsTitle.style.fontSize = '13px';
    basicsTitle.style.lineHeight = '1.3';

    const basics = div('multi-config');
    basics.style.display = 'grid';
    basics.style.gridTemplateColumns = 'repeat(2, minmax(0, 1fr))';
    basics.style.gap = '10px';
    basics.style.alignItems = 'stretch';

    const tolCard = div('feature-row__summary');
    tolCard.style.display = 'grid';
    tolCard.style.gap = '8px';
    tolCard.style.padding = '10px';
    tolCard.style.borderRadius = '14px';
    tolCard.style.border = '1px solid var(--border-accent-soft)';
    tolCard.style.background = 'var(--surface-control)';

    const unitCard = div('feature-row__summary');
    unitCard.style.display = 'grid';
    unitCard.style.gap = '8px';
    unitCard.style.padding = '10px';
    unitCard.style.borderRadius = '14px';
    unitCard.style.border = '1px solid var(--border-accent-soft)';
    unitCard.style.background = 'var(--surface-control)';

    tol.field.style.margin = '0';
    unit.field.style.margin = '0';
    tol.field.style.padding = '0';
    unit.field.style.padding = '0';
    tol.field.style.border = '0';
    unit.field.style.border = '0';
    tol.field.style.background = 'transparent';
    unit.field.style.background = 'transparent';
    tol.field.style.display = 'grid';
    unit.field.style.display = 'grid';
    tol.field.style.gap = '6px';
    unit.field.style.gap = '6px';

    tol.input.style.width = '100%';
    unit.select.style.width = '100%';
    tol.input.style.boxSizing = 'border-box';
    unit.select.style.boxSizing = 'border-box';
    tol.input.style.padding = '8px 10px';
    unit.select.style.padding = '8px 10px';
    tol.input.style.borderRadius = '12px';
    unit.select.style.borderRadius = '12px';
    tol.input.style.border = '1px solid var(--border-soft)';
    unit.select.style.border = '1px solid var(--border-soft)';
    tol.input.style.background = 'var(--surface-control)';
    unit.select.style.background = 'var(--surface-control)';

    tolCard.append(tol.field);
    unitCard.append(unit.field);
    basics.append(tolCard, unitCard);
    basicsCard.append(basicsTitle, basics);

    const selectedWrap = div('feature-row__summary');
    selectedWrap.style.display = 'grid';
    selectedWrap.style.gridTemplateRows = 'auto minmax(0, 1fr)';
    selectedWrap.style.gap = '8px';
    selectedWrap.style.marginTop = '0';
    selectedWrap.style.padding = '10px 12px';
    selectedWrap.style.borderRadius = '16px';
    selectedWrap.style.border = '1px solid var(--border-accent-soft)';
    selectedWrap.style.background = 'var(--surface-help)';
    selectedWrap.style.width = '100%';
    selectedWrap.style.boxSizing = 'border-box';
    selectedWrap.style.height = '92px';
    selectedWrap.style.minHeight = '92px';
    selectedWrap.style.overflow = 'hidden';

    const selectedHead = document.createElement('div');
    selectedHead.style.display = 'flex';
    selectedHead.style.justifyContent = 'space-between';
    selectedHead.style.alignItems = 'center';
    selectedHead.style.gap = '8px';
    const selectedTitle = document.createElement('strong');
    selectedTitle.textContent = '검토 파라미터';
    const selectedCount = document.createElement('span');
    selectedCount.className = 'chip chip--info';
    selectedHead.append(selectedTitle, selectedCount);

    const selectedChips = div('familylink-selected-chips');
    selectedChips.style.display = 'flex';
    selectedChips.style.flexWrap = 'wrap';
    selectedChips.style.gap = '8px';
    selectedChips.style.height = '42px';
    selectedChips.style.minHeight = '42px';
    selectedChips.style.alignContent = 'flex-start';
    selectedChips.style.alignItems = 'flex-start';
    selectedChips.style.padding = '2px 0 0';
    selectedChips.style.overflow = 'auto';
    selectedWrap.append(selectedHead, selectedChips);

    const picker = div('feature-row__summary');
    picker.style.display = 'grid';
    picker.style.gap = '8px';
    picker.style.padding = '10px 12px';
    picker.style.border = '1px solid var(--border-soft)';
    picker.style.borderRadius = '16px';
    picker.style.background = 'var(--surface-elevated)';
    picker.style.boxShadow = 'var(--shadow-soft)';
    picker.style.width = '100%';
    picker.style.boxSizing = 'border-box';

    const pickerHead = document.createElement('div');
    pickerHead.style.display = 'flex';
    pickerHead.style.justifyContent = 'space-between';
    pickerHead.style.alignItems = 'center';
    pickerHead.style.gap = '8px';
    const pickerTitle = document.createElement('strong');
    pickerTitle.textContent = '검토 파라미터 선택';
    const pickerBadge = document.createElement('span');
    pickerBadge.className = 'chip chip--info';
    pickerBadge.textContent = '선택 필요';
    pickerHead.append(pickerTitle, pickerBadge);

    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = '공유 파라미터 이름 / 그룹 검색';
    searchInput.style.width = '100%';
    searchInput.style.padding = '8px 10px';
    searchInput.style.borderRadius = '12px';
    searchInput.style.border = '1px solid var(--border-accent-soft)';
    searchInput.style.background = 'var(--surface-control)';
    searchInput.style.boxSizing = 'border-box';
    searchInput.style.outline = 'none';

    const searchMeta = document.createElement('div');
    searchMeta.style.display = 'flex';
    searchMeta.style.justifyContent = 'space-between';
    searchMeta.style.alignItems = 'center';
    searchMeta.style.gap = '10px';
    const searchInfo = document.createElement('span');
    searchInfo.style.color = 'var(--muted,#64748b)';
    searchInfo.style.fontSize = '12px';
    const refreshBtn = document.createElement('button');
    refreshBtn.type = 'button';
    refreshBtn.className = 'btn btn--secondary';
    refreshBtn.textContent = '목록 새로고침';
    refreshBtn.addEventListener('click', () => requestConnectorParamList('settings-refresh'));
    searchMeta.append(searchInfo, refreshBtn);

    const listWrap = div('familylink-target-list');
    listWrap.style.height = '132px';
    listWrap.style.minHeight = '132px';
    listWrap.style.overflow = 'auto';
    listWrap.style.border = '1px solid var(--border-soft)';
    listWrap.style.borderRadius = '12px';
    listWrap.style.padding = '6px';
    listWrap.style.background = 'var(--surface-control)';
    const empty = div('familylink-target-empty');
    empty.style.display = 'grid';
    empty.style.placeItems = 'center';
    empty.style.minHeight = '100%';
    empty.textContent = '목록을 불러오는 중입니다.';

    searchInput.addEventListener('input', () => renderConnectorList());

    picker.append(pickerHead, searchInput, searchMeta, listWrap);
    panel.append(basicsCard, selectedWrap, picker);

    function normalizeDraftParamsForDisplay(raw) {
      let next = normalizeConnectorParamNames(raw);
      if (next.length > 1) {
        next = next.filter((name) => String(name).toLowerCase() !== 'comments');
      }
      return next;
    }

    function getDraftParams() {
      const draft = state.features.connector.configDraft;
      const raw = draft.paramItems && draft.paramItems.length ? draft.paramItems : draft.param;
      draft.paramItems = normalizeDraftParamsForDisplay(raw);
      draft.param = draft.paramItems.join(',') || 'Comments';
      return draft.paramItems;
    }

    function commitParamSelection(next) {
      const draft = state.features.connector.configDraft;
      draft.paramItems = normalizeDraftParamsForDisplay(next);
      draft.param = draft.paramItems.join(',') || 'Comments';
      markFeatureDirty('connector');
      renderConnectorSelected();
      renderConnectorList();
      refreshConnectorFeatureSummary();
    }

    function renderConnectorSelected() {
      const selected = getDraftParams();
      selectedCount.textContent = selected.length ? `${selected.length}개 선택` : '미선택';
      pickerBadge.textContent = selected.length > 0 ? `선택 ${selected.length}` : '선택 필요';
      selectedChips.innerHTML = '';
      if (!selected.length) {
        const emptyState = document.createElement('div');
        emptyState.style.width = '100%';
        emptyState.style.padding = '8px 12px';
        emptyState.style.borderRadius = '12px';
        emptyState.style.border = '1px dashed var(--border-accent-soft)';
        emptyState.style.background = 'var(--surface-empty)';
        emptyState.style.color = 'var(--muted,#64748b)';
        emptyState.style.display = 'grid';
        emptyState.style.placeItems = 'center';
        emptyState.style.minHeight = '100%';
        emptyState.textContent = '선택된 파라미터가 없습니다.';
        selectedChips.append(emptyState);
        return;
      }
      selected.forEach((name) => {
        const chip = document.createElement('button');
        chip.type = 'button';
        chip.className = 'chip chip--info';
        chip.textContent = `${name} ×`;
        chip.style.padding = '6px 10px';
        chip.style.borderRadius = '999px';
        chip.style.border = '1px solid var(--border-accent-soft)';
        chip.style.background = 'var(--surface-control)';
        chip.addEventListener('click', () => {
          commitParamSelection(getDraftParams().filter((x) => x.toLowerCase() !== String(name).toLowerCase()));
        });
        selectedChips.append(chip);
      });
    }

    function renderConnectorList(payload) {
      if (payload) {
        state.connectorParamItems = payload?.ok !== false ? normalizeConnectorParamItems(payload) : [];
      }
      const query = String(searchInput.value || '').trim().toLowerCase();
      const items = Array.isArray(state.connectorParamItems) ? state.connectorParamItems : [];
      const selected = new Set(getDraftParams().map((x) => String(x).toLowerCase()));
      const filtered = items.filter((item) => {
        if (!query) return true;
        const hay = `${item.name || ''} ${item.groupName || ''}`.toLowerCase();
        return hay.includes(query);
      });
      searchInfo.textContent = items.length ? `공유 파라미터 ${filtered.length}/${items.length}개` : '공유 파라미터 목록이 없습니다.';
      listWrap.innerHTML = '';
      if (payload?.ok === false) {
        const err = div('familylink-target-empty');
        err.textContent = payload?.message || '공유 파라미터 목록을 불러오지 못했습니다.';
        listWrap.append(err);
        renderConnectorSelected();
        return;
      }
      if (!items.length) {
        listWrap.append(empty);
        renderConnectorSelected();
        return;
      }
      if (!filtered.length) {
        const nohit = div('familylink-target-empty');
        nohit.textContent = '검색 결과가 없습니다.';
        listWrap.append(nohit);
        renderConnectorSelected();
        return;
      }
      filtered.forEach((item) => {
        const row = document.createElement('button');
        row.type = 'button';
        row.className = 'familylink-target-row';
        row.style.display = 'grid';
        row.style.gridTemplateColumns = '1fr auto';
        row.style.alignItems = 'center';
        row.style.width = '100%';
        row.style.textAlign = 'left';
        row.style.border = '0';
        row.style.background = selected.has(String(item.name || '').toLowerCase()) ? 'var(--surface-note)' : 'transparent';
        row.style.borderRadius = '10px';
        row.style.padding = '8px 10px';
        row.style.cursor = 'pointer';

        const info = document.createElement('span');
        info.style.display = 'grid';
        info.style.gap = '2px';
        const name = document.createElement('strong');
        name.textContent = item.name || '';
        const sub = document.createElement('small');
        sub.textContent = item.groupName ? `${item.groupName}${item.guid ? ' · ' + item.guid.slice(0, 8) : ''}` : (item.guid ? item.guid.slice(0, 8) : '');
        info.append(name, sub);

        const action = document.createElement('span');
        action.className = selected.has(String(item.name || '').toLowerCase()) ? 'chip chip--ok' : 'chip chip--info';
        action.textContent = selected.has(String(item.name || '').toLowerCase()) ? '선택됨' : '추가';

        row.append(info, action);
        row.addEventListener('click', () => {
          const next = getDraftParams();
          const idx = next.findIndex((x) => String(x).toLowerCase() === String(item.name || '').toLowerCase());
          if (idx >= 0) next.splice(idx, 1);
          else next.push(item.name);
          commitParamSelection(next);
        });
        listWrap.append(row);
      });
      renderConnectorSelected();
    }

    buildConnectorConfig.renderList = renderConnectorList;
    renderConnectorList();
    return { panel, controls: { tol, unit, searchInput, listWrap, selectedWrap, renderConnectorList, renderConnectorSelected } };
  }

  function buildGuidConfig() {

    const panel = div('multi-config');
    const includeFamily = makeCheckboxField('패밀리 포함');
    includeFamily.input.checked = state.features.guid.configDraft.includeFamily;
    includeFamily.input.addEventListener('change', () => {
      state.features.guid.configDraft.includeFamily = includeFamily.input.checked;
      markFeatureDirty('guid');
    });
    const includeAnno = makeCheckboxField('Annotation 패밀리 포함');
    includeAnno.input.checked = state.features.guid.configDraft.includeAnnotation;
    includeAnno.input.addEventListener('change', () => {
      state.features.guid.configDraft.includeAnnotation = includeAnno.input.checked;
      markFeatureDirty('guid');
    });
    panel.append(includeFamily.field, includeAnno.field);
    return { panel, controls: { includeFamily, includeAnno } };
  }

  function buildPointsConfig() {
    const panel = div('multi-config');
    const unit = makeSelectField('단위', [
      { value: 'ft', label: 'Decimal Feet' },
      { value: 'm', label: 'Meters (m)' },
      { value: 'mm', label: 'Millimeters (mm)' }
    ]);
    unit.select.value = state.features.points.configDraft.unit;
    unit.select.addEventListener('change', () => {
      state.features.points.configDraft.unit = unit.select.value;
      markFeatureDirty('points');
    });
    panel.append(unit.field);
    return { panel, controls: { unit } };
  }

  function buildFamilyLinkConfig() {
    const panel = div('multi-config');
    const searchWrap = div('familylink-search');
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = '파라미터 검색...';
    searchWrap.append(searchInput);

    const listWrap = div('familylink-target-list');
    const listEmpty = div('familylink-target-empty');
    listEmpty.textContent = 'Shared Parameter 목록이 없습니다.';
    listWrap.append(listEmpty);

    const selectedWrap = div('familylink-selected');
    const selectedCount = document.createElement('strong');
    const selectedChips = div('familylink-selected-chips');
    selectedWrap.append(selectedCount, selectedChips);

    const advanced = makeField('고급 입력 (이름|GUID)', 'familylinkTargets', '예: ParamA|11111111-1111-1111-1111-111111111111', 'textarea');
    advanced.input.value = state.features.familylink.configDraft.targetsText;
    advanced.input.addEventListener('change', () => {
      state.features.familylink.configDraft.targetsText = advanced.input.value;
      state.features.familylink.configDraft.selectedTargets = parseFamilyLinkTargets(advanced.input.value);
      markFeatureDirty('familylink');
      renderFamilyLinkList();
    });

    searchInput.addEventListener('input', () => {
      renderFamilyLinkList();
    });

    panel.append(searchWrap, listWrap, selectedWrap, advanced.field);

    function renderFamilyLinkList(payload) {
      const ok = payload?.ok !== false;
      const items = ok ? state.sharedParamItems : [];
      const query = searchInput.value.trim().toLowerCase();
      listWrap.innerHTML = '';
      if (!ok) {
        const error = div('familylink-target-empty');
        error.textContent = payload?.message || 'Shared Parameter 목록을 불러오지 못했습니다.';
        listWrap.append(error);
      } else if (!items.length) {
        listWrap.append(listEmpty);
      } else {
        const filtered = items.filter((item) => {
          if (!query) return true;
          const hay = `${item.name || ''} ${item.groupName || ''} ${item.guid || ''}`.toLowerCase();
          return hay.includes(query);
        });
        if (!filtered.length) {
          listWrap.append(listEmpty);
        } else {
          filtered.forEach((item) => {
            const row = document.createElement('label');
            row.className = 'familylink-target-row';
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.checked = state.features.familylink.configDraft.selectedTargets.some((t) => t.guid === item.guid);
            checkbox.addEventListener('change', () => {
              const next = [...state.features.familylink.configDraft.selectedTargets];
              if (checkbox.checked) {
                next.push(item);
              } else {
                const idx = next.findIndex((t) => t.guid === item.guid);
                if (idx >= 0) next.splice(idx, 1);
              }
              state.features.familylink.configDraft.selectedTargets = dedupeTargets(next);
              state.features.familylink.configDraft.targetsText = buildTargetsText(state.features.familylink.configDraft.selectedTargets);
              advanced.input.value = state.features.familylink.configDraft.targetsText;
              markFeatureDirty('familylink');
              renderFamilyLinkSelected();
            });
            const label = document.createElement('span');
            const guidShort = (item.guid || '').slice(0, 8);
            label.textContent = `${item.name || ''} (${item.groupName || '-'}) ${guidShort}`;
            row.append(checkbox, label);
            listWrap.append(row);
          });
        }
      }
      renderFamilyLinkSelected();
    }

    function renderFamilyLinkSelected() {
      const selected = state.features.familylink.configDraft.selectedTargets || [];
      selectedCount.textContent = `선택됨 ${selected.length}개`;
      selectedChips.innerHTML = '';
      selected.forEach((item) => {
        const chip = document.createElement('span');
        chip.className = 'chip chip--info';
        chip.textContent = item.name || '';
        selectedChips.append(chip);
      });
      updateRunSummary();
    }

    buildFamilyLinkConfig.renderList = renderFamilyLinkList;
    renderFamilyLinkList();
    return { panel, controls: { searchInput, listWrap, selectedWrap, advanced } };
  }

  function buildRvtSection() {
    const section = div('multi-section rvt-panel HubLeftRvt');
    const head = div('rvt-panel-header');
    const title = document.createElement('div');
    title.className = 'rvt-panel-title';
    const badge = document.createElement('span');
    badge.className = 'chip chip--info';
    title.innerHTML = '<h3>RVT 리스트</h3>';
    title.append(badge);

    const controls = div('multi-rvt-controls');
    const btnAdd = cardBtn('RVT 추가', handleAddRvt, 'btn--primary');
    const btnRemove = cardBtn('선택 제거', handleRemoveSelected, 'btn--secondary');
    const btnClear = cardBtn('목록 지우기', handleClearList, 'btn--danger');
    controls.append(btnAdd, btnRemove, btnClear);

    head.append(title, controls);
    section.append(head);

    const body = div('rvt-panel-body');
    const tableWrap = div('rvt-table-wrap');
    const { table, tbody, master } = createRvtTable();
    const summary = div('multi-rvt-summary');
    const footer = div('rvt-list-footer');
    const footerRight = div('rvt-list-footer__right');
    const expandBtn = cardBtn('리스트 크게 보기', () => openExpandedRvtModal(), 'btn--secondary');
    const empty = div('rvt-empty');
    const emptyTitle = document.createElement('strong');
    emptyTitle.textContent = '등록된 RVT가 없습니다.';
    const emptySub = document.createElement('span');
    emptySub.textContent = 'RVT 추가로 파일을 등록하세요.';
    const emptyBtn = cardBtn('RVT 추가', handleAddRvt, 'btn--primary');
    empty.append(emptyTitle, emptySub, emptyBtn);

    tableWrap.append(table);
    footerRight.append(expandBtn);
    footer.append(summary, footerRight);
    body.append(tableWrap, empty, footer);
    section.append(body);

    function syncMaster() {
      const allChecked = state.rvtList.length > 0 && state.rvtList.every((p) => state.rvtChecked.has(p));
      master.checked = allChecked;
    }

    master.addEventListener('change', () => {
      if (master.checked) {
        state.rvtList.forEach((p) => state.rvtChecked.add(p));
      } else {
        state.rvtChecked.clear();
      }
      renderExpandedList();
      renderRvtList();
    });

    function renderRvtList() {
      const rows = state.rvtList.map((path, idx) => ({
        index: idx + 1,
        path,
        name: getRvtName(path),
        checked: state.rvtChecked.has(path),
        onToggle: (checked) => {
          if (checked) state.rvtChecked.add(path);
          else state.rvtChecked.delete(path);
          syncMaster();
        }
      }));
      const count = state.rvtList.length;
      tbody.innerHTML = '';
      if (count > 0) {
        renderRvtRows(tbody, rows);
      }
      summary.textContent = `총 파일 수: ${count}`;
      badge.textContent = `${count}개`;
      empty.style.display = count ? 'none' : 'flex';
      tableWrap.style.display = count ? 'block' : 'none';
      expandBtn.disabled = count === 0;
      syncMaster();
      btnRemove.disabled = state.rvtChecked.size === 0;
      btnClear.disabled = state.rvtList.length === 0;
      updateRunSummary();
      if (buildRvtExpandedModal.render) buildRvtExpandedModal.render();
    }

    buildRvtSection.render = renderRvtList;
    renderRvtList();
    return section;
  }

  function buildExecutionActionPanel(options = {}) {
    const mode = options.mode || state.ui.multiMode || 'bqc';
    state.ui.currentDocButtons = [];
    state.ui.openMultiButtons = [];

    const section = div('multi-section multi-action-card');
    const head = div('multi-action-card__head');
    const headTitle = document.createElement('strong');
    headTitle.textContent = '검토 실행';
    const headText = document.createElement('span');
    headText.textContent = '왼쪽 기능 영역을 눌러 선택한 뒤 검토를 시작하세요.';
    head.append(headTitle, headText);

    const summary = div('multi-action-card__summary');
    const sharedHint = div('run-summary__hint');
    sharedHint.style.display = 'none';
    summary.append(sharedHint);
    state.ui.actionCommonSummaryEl = null;

    const progress = div('multi-action-card__progress');
    const progressText = document.createElement('span');
    const progressDetail = document.createElement('small');
    const progressBar = document.createElement('div');
    progressBar.className = 'run-progress';
    const progressFill = document.createElement('div');
    progressFill.className = 'run-progress-fill';
    progressBar.append(progressFill);
    progress.append(progressText, progressDetail, progressBar);

    const actions = div('multi-action-card__actions');
    const actionRow = div('multi-action-card__buttons');
    const currentBtn = cardBtn('현재 파일 검토', handleRunCurrentFile, 'btn--primary');
    const recentBtn = cardBtn('최근 결과 보기', openRecentResultView, 'btn--primary');
    recentBtn.classList.add('btn--recent');
    currentBtn.classList.add('btn--action-main');
    recentBtn.classList.add('btn--action-main');
    actionRow.append(currentBtn, recentBtn);

    const multiBtn = cardBtn('RVT 여러개 검토', openExpandedRvtModal, 'btn--secondary');
    multiBtn.classList.add('btn--multi', 'btn--action-main');
    actionRow.append(multiBtn);

    actions.append(actionRow);
    if (mode === 'bqc') {
      const actionExtras = div('multi-action-card__extras');
      const commonBtn = cardBtn('공통 설정', () => openSettings('common', '그룹 공통 옵션'), 'btn--secondary');
      commonBtn.classList.add('btn--settings-inline');
      const commonSummary = document.createElement('div');
      commonSummary.className = 'multi-action-stack__note';
      commonSummary.textContent = `공통 설정: ${buildCommonSummary()}`;
      actionExtras.append(commonBtn, commonSummary);
      actions.append(actionExtras);
      state.ui.actionCommonSummaryEl = commonSummary;
    } else {
      state.ui.actionCommonSummaryEl = null;
    }
    state.ui.currentDocButtons.push(currentBtn);
    state.ui.openMultiButtons.push(multiBtn);
    state.ui.bqcRecentOpenBtn = recentBtn;

    section.append(head, progress, actions, summary);

    buildRunBar.summary = summary;
    buildRunBar.summaryTitle = null;
    buildRunBar.summaryDetail = null;
    buildRunBar.runSharedParamHint = sharedHint;
    buildRunBar.progressText = progressText;
    buildRunBar.progressDetail = progressDetail;
    buildRunBar.progressFill = progressFill;
    buildRunBar.startBtn = null;
    updateRunSummary();
    updateRunProgress(0, '대기 중', '');
    updateRunActionLabel();
    return section;
  }

  function buildBqcRecentResultPanel() {
    const section = div('multi-section multi-side-card multi-side-card--launcher');
    const top = div('multi-side-card__top');
    const head = div('multi-side-card__head');
    const title = document.createElement('h3');
    title.textContent = '최근 결과 보기';
    const caption = document.createElement('div');
    caption.className = 'multi-recent-caption';
    caption.textContent = '아직 실행 결과가 없습니다.';
    const hint = document.createElement('div');
    hint.className = 'multi-recent-hint';
    hint.textContent = '영역을 눌러 최근 결과 상세 창을 열고, 검토별 엑셀을 각각 내보낼 수 있습니다.';
    head.append(title, caption, hint);

    const preview = div('multi-recent-launcher');
    const previewTitle = document.createElement('strong');
    previewTitle.textContent = '상세 결과 창 열기';
    const previewText = document.createElement('span');
    previewText.textContent = '파일별 결과 표와 기능별 엑셀 내보내기를 별도 창에서 확인합니다.';
    preview.append(previewTitle, previewText);

    const actions = div('multi-side-card__actions');
    const openBtn = cardBtn('최근 결과 보기', openRecentResultView, 'btn--primary');
    const resetBtn = cardBtn('결과 초기화', resetRunResults, 'btn--secondary');
    actions.append(openBtn, resetBtn);
    top.append(head, preview, actions);
    section.append(top);
    section.classList.add('is-clickable');
    section.setAttribute('role', 'button');
    section.tabIndex = 0;
    section.addEventListener('click', (ev) => {
      if (ev.target.closest('button')) return;
      openRecentResultView();
    });
    section.addEventListener('keydown', (ev) => {
      if (ev.key !== 'Enter' && ev.key !== ' ') return;
      if (ev.target.closest('button')) return;
      ev.preventDefault();
      openRecentResultView();
    });

    state.ui.bqcRecentCaption = caption;
    state.ui.bqcRecentHint = hint;
    state.ui.bqcRecentOpenBtn = openBtn;
    state.ui.resultResetButtons = [resetBtn];
    updateBqcSidebar();
    return section;
  }

  function buildRvtExpandedModal() {
    const overlay = div('rvt-expand-overlay');
    overlay.classList.add('is-hidden');
    const modal = div('rvt-expand-modal');
    const toolbar = div('rvt-expand-toolbar');
    const titleWrap = div('rvt-expand-title');
    const title = document.createElement('h3');
    title.textContent = 'RVT 여러개 검토';
    const badge = document.createElement('span');
    badge.className = 'chip chip--info';
    titleWrap.append(title, badge);
    const toolbarActions = div('rvt-expand-actions');
    const btnAdd = cardBtn('RVT 추가', handleAddRvt, 'btn--primary');
    const btnRemove = cardBtn('선택 제거', handleRemoveSelected, 'btn--secondary');
    const btnClear = cardBtn('목록 지우기', handleClearList, 'btn--secondary');
    const btnClose = cardBtn('닫기', closeExpandedRvtModal, 'btn--secondary');
    toolbarActions.append(btnAdd, btnRemove, btnClear, btnClose);
    toolbar.append(titleWrap, toolbarActions);

    const body = div('rvt-expand-body');
    const panel = div('rvt-expand-panel');

    const listSection = div('rvt-expand-section rvt-expand-section--list');
    const listHead = div('rvt-expand-section__head');
    const listTitle = document.createElement('h4');
    listTitle.textContent = '선택한 RVT 목록';
    const listSub = document.createElement('span');
    listSub.textContent = '목록에서 체크된 파일만 제거할 수 있고, 검토는 전체 등록 목록 기준으로 실행됩니다.';
    listHead.append(listTitle, listSub);
    const tableWrap = div('rvt-expand-table');
    const { table, tbody, master } = createRvtTable();
    table.classList.add('rvt-expand-table__grid');
    const listEmpty = div('rvt-expand-list-empty');
    const listEmptyTitle = document.createElement('strong');
    listEmptyTitle.textContent = '등록된 RVT가 없습니다.';
    const listEmptyText = document.createElement('span');
    listEmptyText.textContent = 'RVT 추가 버튼으로 파일을 등록하면 여기에서 바로 목록을 확인할 수 있습니다.';
    const listEmptyBtn = cardBtn('RVT 추가', handleAddRvt, 'btn--primary');
    listEmpty.append(listEmptyTitle, listEmptyText, listEmptyBtn);
    tableWrap.append(table, listEmpty);
    listSection.append(listHead, tableWrap);

    const sideSection = div('rvt-expand-section rvt-expand-section--side');
    const featureSection = div('rvt-expand-subsection');
    const featureHead = div('rvt-expand-subsection__head');
    const featureTitle = document.createElement('h4');
    featureTitle.textContent = '선택된 기능';
    const featureCount = document.createElement('span');
    featureCount.className = 'chip chip--info';
    const featureList = div('rvt-expand-feature-list');
    featureHead.append(featureTitle, featureCount);
    featureSection.append(featureHead, featureList);

    const recentSection = div('rvt-expand-subsection rvt-expand-subsection--recent');
    const recentHead = div('rvt-expand-subsection__head');
    const recentTitle = document.createElement('h4');
    recentTitle.textContent = '최근 결과';
    const recentCaption = document.createElement('div');
    recentCaption.className = 'multi-recent-caption';
    const recentHint = document.createElement('div');
    recentHint.className = 'multi-recent-hint';
    recentHead.append(recentTitle, recentCaption, recentHint);
    const recentTable = div('multi-recent-table multi-recent-table--modal');
    const recentTableEl = document.createElement('table');
    recentTableEl.innerHTML = `
      <colgroup>
        <col style="width:52%">
        <col style="width:16%">
        <col style="width:16%">
        <col style="width:16%">
      </colgroup>
      <thead>
        <tr>
          <th>파일</th>
          <th>전체</th>
          <th>오류</th>
          <th>near</th>
        </tr>
      </thead>
      <tbody></tbody>`;
    const recentTbody = recentTableEl.querySelector('tbody');
    const recentEmpty = div('multi-recent-table__empty');
    recentEmpty.textContent = '최근 실행 결과가 없습니다.';
    recentTable.append(recentTableEl, recentEmpty);
    recentSection.append(recentHead, recentTable);

    sideSection.append(featureSection, recentSection);
    panel.append(listSection, sideSection);
    body.append(panel);

    const footer = div('rvt-expand-footer');
    const footerSummary = div('rvt-expand-footer__summary');
    const footerTitle = document.createElement('strong');
    footerTitle.textContent = '선택한 RVT 검토 준비';
    const footerSub = document.createElement('span');
    footerSummary.append(footerTitle, footerSub);
    const footerActions = div('rvt-expand-footer__actions');
    const runBtn = cardBtn('선택한 RVT 검토 시작', onRun, 'btn--primary');
    const closeBtn = cardBtn('닫기', closeExpandedRvtModal, 'btn--secondary');
    footerActions.append(runBtn, closeBtn);
    footer.append(footerSummary, footerActions);

    modal.append(toolbar, body, footer);
    overlay.append(modal);

    overlay.addEventListener('click', (ev) => {
      ev.stopPropagation();
    });
    modal.addEventListener('click', (ev) => {
      ev.stopPropagation();
    });

    master.addEventListener('change', () => {
      if (master.checked) {
        state.rvtList.forEach((p) => state.rvtChecked.add(p));
      } else {
        state.rvtChecked.clear();
      }
      renderExpandedList();
      renderRvtList();
    });

    function renderExpandedList() {
      const rows = state.rvtList.map((path, idx) => ({
        index: idx + 1,
        path,
        name: getRvtName(path),
        checked: state.rvtChecked.has(path),
        onToggle: (checked) => {
          if (checked) state.rvtChecked.add(path);
          else state.rvtChecked.delete(path);
          renderExpandedList();
          renderRvtList();
        }
      }));
      const count = state.rvtList.length;
      tbody.innerHTML = '';
      if (count > 0) renderRvtRows(tbody, rows);
      badge.textContent = `${count}개`;
      footerSub.textContent = count ? `${count}개 RVT가 등록되어 있습니다.` : 'RVT를 추가하면 다중 검토를 시작할 수 있습니다.';
      listEmpty.style.display = count ? 'none' : 'flex';
      table.style.display = count ? 'table' : 'none';
      master.checked = count > 0 && state.rvtList.every((p) => state.rvtChecked.has(p));
      btnRemove.disabled = state.rvtChecked.size === 0;
      btnClear.disabled = state.rvtList.length === 0;
      renderModalFeatureSummary();
      renderRecentResultTable([{
        caption: recentCaption,
        hint: recentHint,
        tbody: recentTbody,
        empty: recentEmpty
      }]);
      updateMultiRunBtnState();
    }

    state.ui.modalFeatureList = featureList;
    state.ui.modalFeatureCount = featureCount;
    state.ui.modalRecentCaption = recentCaption;
    state.ui.modalRecentHint = recentHint;
    state.ui.modalRecentTableBody = recentTbody;
    state.ui.modalRecentEmpty = recentEmpty;
    state.ui.modalRunButtons = [runBtn];
    buildRvtExpandedModal.overlay = overlay;
    buildRvtExpandedModal.badge = badge;
    buildRvtExpandedModal.render = renderExpandedList;
    return overlay;
  }

  function buildReviewSummaryModal() {
    const overlay = div('review-summary-backdrop is-hidden');
    const modal = div('review-summary-modal');
    const header = div('review-summary-header');
    const title = document.createElement('h3');
    title.textContent = '최근 결과 보기';
    header.append(title);

    const body = div('review-summary-body');
    const stats = div('review-summary-stats');

    const caption = document.createElement('div');
    caption.className = 'review-summary-caption';
    caption.textContent = '파일별 검토 결과';

    const tableWrap = div('review-summary-table');
    const table = document.createElement('table');
    table.innerHTML = `
      <thead>
        <tr>
          <th>파일명</th>
          <th>상태</th>
          <th>전체</th>
          <th>오류</th>
          <th>near</th>
          <th>비고</th>
        </tr>
      </thead>
      <tbody></tbody>`;
    const tbody = table.querySelector('tbody');
    tableWrap.append(table);

    const exportGuide = div('review-export-guide');
    exportGuide.textContent = '기능별 엑셀 버튼은 기존 저장 방식과 동일한 이벤트를 재사용합니다. 여러 기능을 실행한 경우 필요한 결과만 기능별로 각각 저장할 수 있습니다.';

    const featureList = div('review-feature-list');

    body.append(stats, caption, tableWrap, exportGuide, featureList);

    const footer = div('review-summary-footer');
    const resetBtn = document.createElement('button');
    resetBtn.type = 'button';
    resetBtn.className = 'btn btn--secondary';
    resetBtn.textContent = '결과 초기화';
    const confirmBtn = document.createElement('button');
    confirmBtn.type = 'button';
    confirmBtn.className = 'btn btn--primary';
    confirmBtn.textContent = '닫기';
    footer.append(resetBtn, confirmBtn);

    modal.append(header, body, footer);
    overlay.append(modal);

    overlay.addEventListener('click', (ev) => ev.stopPropagation());
    modal.addEventListener('click', (ev) => ev.stopPropagation());
    confirmBtn.addEventListener('click', () => {
      overlay.classList.add('is-hidden');
    });
    resetBtn.addEventListener('click', () => {
      resetRunResults();
      overlay.classList.add('is-hidden');
    });

    buildReviewSummaryModal.overlay = overlay;
    buildReviewSummaryModal.title = title;
    buildReviewSummaryModal.stats = stats;
    buildReviewSummaryModal.tbody = tbody;
    buildReviewSummaryModal.featureList = featureList;
    buildReviewSummaryModal.resetBtn = resetBtn;
    return overlay;
  }

  function showReviewSummary(payload) {
    if (!buildReviewSummaryModal.overlay) return;
    const mode = normalizeMultiMode(state.ui.multiMode || 'bqc');
    const featureSummaries = payload?.featureSummaries && typeof payload.featureSummaries === 'object'
      ? payload.featureSummaries
      : {};
    const items = Array.isArray(payload?.items) ? payload.items : [];
    const hasPayloadData = !!(
      items.length ||
      Object.keys(featureSummaries).length ||
      Number(payload?.total) ||
      Number(payload?.success) ||
      Number(payload?.skipped) ||
      Number(payload?.failed)
    );
    if (hasPayloadData) {
      state.ui.reviewSummaryData = payload;
      state.ui.reviewSummaryByMode[mode] = payload;
    }
    if (payload?.finishedAt) {
      state.ui.lastRunAt = payload.finishedAt;
      state.ui.lastRunAtByMode[mode] = payload.finishedAt;
    } else if (hasPayloadData && !state.ui.lastRunAt) {
      state.ui.lastRunAt = new Date().toISOString();
      state.ui.lastRunAtByMode[mode] = state.ui.lastRunAt;
    }
    const stats = buildReviewSummaryModal.stats;
    const tbody = buildReviewSummaryModal.tbody;
    const featureList = buildReviewSummaryModal.featureList;
    if (!stats || !tbody || !featureList) return;

    const total = Number(payload?.total) || 0;
    const success = Number(payload?.success) || 0;
    const skipped = Number(payload?.skipped) || 0;
    const failed = Number(payload?.failed) || 0;
    const rows = getReviewTableRows(payload);
    if (buildReviewSummaryModal.title) {
      buildReviewSummaryModal.title.textContent = '최근 결과 보기';
    }

    stats.innerHTML = '';
    stats.append(
      buildSummaryChip('전체', total, 'summary-chip'),
      buildSummaryChip('완료', success, 'summary-chip summary-chip--success'),
      buildSummaryChip('스킵', skipped, 'summary-chip summary-chip--skip'),
      buildSummaryChip('실패', failed, 'summary-chip summary-chip--fail')
    );

    tbody.innerHTML = '';
    if (!rows.length) {
      const row = document.createElement('tr');
      const cell = document.createElement('td');
      cell.colSpan = 6;
      cell.className = 'review-summary-empty';
      cell.textContent = '표시할 결과가 없습니다.';
      row.append(cell);
      tbody.append(row);
    } else {
      rows.forEach((item) => {
        const row = document.createElement('tr');
        const fileCell = document.createElement('td');
        fileCell.className = 'review-summary-file';
        fileCell.textContent = item.file || '-';

        const statusCell = document.createElement('td');
        const statusChip = document.createElement('span');
        statusChip.className = `summary-status summary-status--${item.status}`;
        statusChip.textContent = normalizeReviewStatus(item.status);
        statusCell.append(statusChip);

        const totalCell = document.createElement('td');
        totalCell.textContent = formatReviewMetric(item.total);
        const issueCell = document.createElement('td');
        issueCell.textContent = formatReviewMetric(item.issues);
        const nearCell = document.createElement('td');
        nearCell.textContent = formatReviewMetric(item.near);
        const noteCell = document.createElement('td');
        noteCell.textContent = item.reason || '';
        row.append(fileCell, statusCell, totalCell, issueCell, nearCell, noteCell);
        tbody.append(row);
      });
    }

    featureList.innerHTML = '';
    const featureEntryMap = new Map();
    const modeKeys = new Set(getModeFeatureKeys(mode));
    Object.entries(featureSummaries).forEach(([key, summary]) => {
      if (!modeKeys.has(key)) return;
      featureEntryMap.set(key, summary || {});
    });
    Array.from(modeKeys)
      .filter((key) => state.results[key]?.hasRun && !featureEntryMap.has(key))
      .forEach((key) => {
        featureEntryMap.set(key, {
          label: FEATURE_META[key]?.label || key,
          lines: state.results[key]?.count
            ? [`결과 건수: ${state.results[key].count}건`]
            : ['최근 실행 결과가 저장되어 있습니다.']
        });
      });
    const featureEntries = Array.from(featureEntryMap.entries());
    featureList.classList.toggle('is-empty', false);
    if (!featureEntries.length) {
      const emptyCard = div('review-feature-card review-feature-card--empty');
      const emptyTitle = document.createElement('strong');
      emptyTitle.textContent = '내보낼 최근 결과가 없습니다.';
      const emptyText = document.createElement('div');
      emptyText.className = 'review-feature-card__empty';
      emptyText.textContent = '검토를 실행하면 이 창에서 기능별 결과 확인과 엑셀 내보내기를 바로 진행할 수 있습니다.';
      emptyCard.append(emptyTitle, emptyText);
      featureList.append(emptyCard);
    }
    featureEntries.forEach(([key, summary]) => {
      const card = div('review-feature-card');
      const head = div('review-feature-card__head');
      const label = document.createElement('strong');
      label.textContent = summary?.label || FEATURE_META[key]?.label || key;
      const badge = document.createElement('span');
      badge.className = 'review-feature-card__badge';
      badge.textContent = state.results[key]?.hasRun ? '내보내기 가능' : '결과 요약';
      head.append(label, badge);

      const list = document.createElement('ul');
      list.className = 'review-feature-card__list';
      const lines = Array.isArray(summary?.lines) ? summary.lines : [];
      (lines.length ? lines : ['표시할 기능 요약이 없습니다.']).forEach((line) => {
        const item = document.createElement('li');
        item.textContent = line;
        list.append(item);
      });

      const action = cardBtn(getFeatureExportActionLabel(key), () => onExport(key), 'btn--primary');
      action.classList.add('review-feature-card__action');
      action.disabled = state.busy || !(state.results[key]?.hasRun);
      card.append(head, list, action);
      featureList.append(card);
    });

    updateBqcSidebar();
    buildReviewSummaryModal.overlay.classList.remove('is-hidden');
  }

  function buildSummaryChip(label, value, className) {
    const chip = document.createElement('div');
    chip.className = className;
    chip.textContent = `${label}: ${value}`;
    return chip;
  }

  function openRecentResultView() {
    showReviewSummary(getCurrentModeReviewSummary());
  }

  function getModeFeatureKeys(mode = state.ui.multiMode || 'bqc') {
    return normalizeMultiMode(mode) === 'utility' ? UTILITY_FEATURE_KEYS.slice() : BQC_FEATURE_KEYS.slice();
  }

  function getCurrentModeReviewSummary() {
    const mode = normalizeMultiMode(state.ui.multiMode || 'bqc');
    return state.ui.reviewSummaryByMode?.[mode] || {};
  }

  function getCurrentModeLastRunAt() {
    const mode = normalizeMultiMode(state.ui.multiMode || 'bqc');
    return state.ui.lastRunAtByMode?.[mode] || null;
  }

  function openExpandedRvtModal() {
    const blockReason = getOpenMultiBlockingReason();
    if (blockReason) {
      toast(blockReason, 'warn');
      return;
    }
    if (!buildRvtExpandedModal.overlay) return;
    state.ui.isRvtListExpanded = true;
    buildRvtExpandedModal.overlay.classList.remove('is-hidden');
    buildRvtExpandedModal.render();
  }

  function closeExpandedRvtModal() {
    if (!buildRvtExpandedModal.overlay) return;
    state.ui.isRvtListExpanded = false;
    buildRvtExpandedModal.overlay.classList.add('is-hidden');
  }

  function buildSelectedFeaturesSection(options = {}) {
    state.ui.currentDocButtons = state.ui.currentDocButtons || [];
    const section = div(`multi-section selected-panel ${options.sectionClass || ''}`.trim());
    const head = div('selected-panel__header');
    const title = document.createElement('h3');
    title.textContent = options.title || '선택된 기능 목록';
    const count = document.createElement('span');
    count.className = 'chip chip--info';
    const actions = div('selected-panel__actions');
    actions.append(count);
    if (options.showCurrentButton) {
      const currentBtn = document.createElement('button');
      currentBtn.type = 'button';
      currentBtn.className = 'btn btn--secondary selected-current-btn';
      currentBtn.textContent = '현재 파일 검토';
      currentBtn.addEventListener('click', handleRunCurrentFile);
      actions.append(currentBtn);
      state.ui.currentDocButtons.push(currentBtn);
    }
    head.append(title, actions);


    const table = document.createElement('table');
    table.className = 'selected-table';
    table.innerHTML = `
      <colgroup>
        <col style="width:auto">
        <col style="width:104px">
        <col style="width:84px">
      </colgroup>
      <thead>
        <tr>
          <th>기능</th>
          <th class="selected-status-col selected-action-col">상태</th>
          <th class="selected-action-col">설정</th>
        </tr>
      </thead>
      <tbody></tbody>`;
    const tbody = table.querySelector('tbody');
    section.append(head, table);

    state.ui.selectedTableBody = tbody;
    state.ui.selectedCount = count;
    state.ui.selectedSection = section;
    renderSelectedFeatures();
    return section;
  }

  function buildRunBar() {
    return buildExecutionActionPanel({ mode: state.ui.multiMode || 'bqc' });
  }

  function renderModalFeatureSummary() {
    const list = state.ui.modalFeatureList;
    const countEl = state.ui.modalFeatureCount;
    if (!list || !countEl) return;
    const enabledKeys = FEATURE_KEYS.filter((key) => state.features[key].enabled);
    countEl.textContent = `${enabledKeys.length}개`;
    list.innerHTML = '';

    if (!enabledKeys.length) {
      const empty = div('rvt-expand-feature-empty');
      empty.textContent = '선택된 기능이 없습니다.';
      list.append(empty);
      return;
    }

    enabledKeys.forEach((key) => {
      const item = div('rvt-expand-feature-item');
      const text = div('rvt-expand-feature-item__text');
      const meta = div('rvt-expand-feature-item__meta');
      const name = document.createElement('strong');
      name.textContent = FEATURE_META[key]?.label || key;
      const desc = document.createElement('span');
      const status = getSelectedFeatureStatus(key);
      desc.textContent = `${FEATURE_META[key]?.desc || ''} · ${status.label}`;
      const badge = document.createElement('span');
      badge.className = `chip status-chip ${status.className}`;
      badge.textContent = status.label;
      meta.append(name, badge);
      text.append(meta, desc);
      item.append(text);
      list.append(item);
    });
  }

  function getRecentResultRows() {
    const payload = getCurrentModeReviewSummary();
    const connectorRows = Array.isArray(payload?.featureSummaries?.connector?.fileSummaries)
      ? payload.featureSummaries.connector.fileSummaries
      : [];
    if (connectorRows.length) {
      return connectorRows.map((row) => ({
        file: row?.file || '',
        total: Number(row?.total) || 0,
        issues: Number(row?.issues) || 0,
        near: Number(row?.near) || 0,
        status: String(row?.status || 'pending'),
        reason: ''
      }));
    }

    const items = Array.isArray(payload?.items) ? payload.items : [];
    return items.map((item) => ({
      file: item?.file || '',
      total: '',
      issues: '',
      near: '',
      status: String(item?.status || 'pending'),
      reason: item?.reason || ''
    }));
  }

  function renderRecentResultTable(targets = []) {
    const rows = getRecentResultRows();
    const runAt = state.ui.lastRunAt ? new Date(state.ui.lastRunAt) : null;
    const captionText = runAt && !Number.isNaN(runAt.getTime())
      ? `마지막 실행 ${runAt.toLocaleString('ko-KR')}`
      : '아직 실행 결과가 없습니다.';
    const hintText = rows.length
      ? '파일별 건수를 빠르게 확인하고 결과 팝업을 다시 열 수 있습니다.'
      : '검토 실행 후 파일별 결과가 여기에 표시됩니다.';

    targets.forEach((target) => {
      if (!target) return;
      if (target.caption) target.caption.textContent = captionText;
      if (target.hint) target.hint.textContent = hintText;
      if (target.tbody) target.tbody.innerHTML = '';
      if (target.empty) target.empty.style.display = rows.length ? 'none' : 'block';
      if (!target.tbody) return;
      rows.forEach((item) => {
        const row = document.createElement('tr');
        const fileCell = document.createElement('td');
        fileCell.className = 'multi-recent-table__file';
        fileCell.textContent = item.file || '-';
        fileCell.title = item.file || '-';
        const totalCell = document.createElement('td');
        totalCell.textContent = formatReviewMetric(item.total);
        const issueCell = document.createElement('td');
        issueCell.textContent = formatReviewMetric(item.issues);
        const nearCell = document.createElement('td');
        nearCell.textContent = formatReviewMetric(item.near);
        row.append(fileCell, totalCell, issueCell, nearCell);
        target.tbody.append(row);
      });
    });
  }

  function getReviewTableRows(payload) {
    const result = [];
    const byFile = new Map();
    const items = Array.isArray(payload?.items) ? payload.items : [];
    items.forEach((item) => {
      const file = item?.file || '';
      if (!file) return;
      byFile.set(file, {
        file,
        status: String(item?.status || 'pending'),
        total: '',
        issues: '',
        near: '',
        reason: item?.reason || ''
      });
    });

    const connectorRows = Array.isArray(payload?.featureSummaries?.connector?.fileSummaries)
      ? payload.featureSummaries.connector.fileSummaries
      : [];
    connectorRows.forEach((item) => {
      const file = item?.file || '';
      if (!file) return;
      const existing = byFile.get(file) || {
        file,
        status: String(item?.status || 'pending'),
        total: '',
        issues: '',
        near: '',
        reason: ''
      };
      existing.total = Number(item?.total) || 0;
      existing.issues = Number(item?.issues) || 0;
      existing.near = Number(item?.near) || 0;
      existing.status = existing.status || String(item?.status || 'pending');
      byFile.set(file, existing);
    });

    byFile.forEach((value) => result.push(value));
    return result;
  }

  function normalizeReviewStatus(status) {
    if (status === 'success') return '완료';
    if (status === 'skipped') return '스킵';
    if (status === 'failed') return '실패';
    return '대기';
  }

  function formatReviewMetric(value) {
    return value === '' || value === null || value === undefined ? '-' : String(value);
  }

  function getFeatureExportActionLabel(key) {
    if (key === 'connector') return '커넥터 결과 엑셀';
    if (key === 'guid') return 'GUID 결과 엑셀';
    if (key === 'familylink') return '패밀리 연동 결과 엑셀';
    if (key === 'points') return 'Point 결과 엑셀';
    return '엑셀 내보내기';
  }

  function getDefaultRecentExportKey() {
    const keys = FEATURE_KEYS.filter((key) => state.results[key]?.hasRun);
    return keys.length === 1 ? keys[0] : '';
  }

  function buildSettingsModal() {
    const overlay = div('modal-overlay');
    const modal = div('modal');
    modal.style.width = 'min(1420px, 96vw)';
    modal.style.maxWidth = '1420px';
    const header = div('modal__header');
    const title = document.createElement('div');
    title.className = 'modal__title';
    const badge = document.createElement('span');
    badge.className = 'chip chip--warn';
    badge.style.display = 'none';
    header.append(title, badge);

    const body = div('modal__body');
    body.style.display = 'grid';
    body.style.gridTemplateColumns = 'minmax(0, 2.35fr) minmax(300px, 1fr)';
    body.style.gap = '14px';
    body.style.alignItems = 'start';
    body.style.width = '100%';
    body.style.justifyContent = 'stretch';
    const form = div('modal__form');
    form.style.display = 'grid';
    form.style.gridTemplateColumns = 'minmax(0, 1fr)';
    form.style.gridAutoFlow = 'row';
    form.style.gap = '14px';
    form.style.alignContent = 'start';
    form.style.justifyItems = 'stretch';
    form.style.minWidth = '0';
    form.style.width = '100%';
    form.style.maxWidth = 'none';
    const help = div('modal__help');
    help.style.minWidth = '0';
    help.style.width = '100%';
    help.style.maxWidth = 'none';
    const sharedBanner = buildSharedParamStatusBanner();
    sharedBanner.style.display = 'none';
    form.append(sharedBanner);
    body.append(form, help);

    const footer = div('modal__footer');
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.className = 'btn btn--ghost';
    cancelBtn.textContent = '취소';
    const applyBtn = document.createElement('button');
    applyBtn.type = 'button';
    applyBtn.className = 'btn btn--primary';
    applyBtn.textContent = '적용';
    footer.append(cancelBtn, applyBtn);

    modal.append(header, body, footer);
    overlay.append(modal);

    cancelBtn.addEventListener('click', cancelSettings);
    applyBtn.addEventListener('click', applySettings);

    buildSettingsModal.overlay = overlay;
    buildSettingsModal.modal = modal;
    buildSettingsModal.title = title;
    buildSettingsModal.badge = badge;
    buildSettingsModal.form = form;
    buildSettingsModal.body = body;
    buildSettingsModal.help = help;
    buildSettingsModal.sharedBanner = sharedBanner;
    return overlay;
  }

  function makeField(label, name, placeholder, type) {
    const field = div('field');
    const lab = document.createElement('label');
    lab.textContent = label;
    const input = type === 'textarea' ? document.createElement('textarea') : document.createElement('input');
    if (type !== 'textarea') input.type = type;
    input.placeholder = placeholder || '';
    input.name = name;
    field.append(lab, input);
    return { field, input };
  }

  function makeSelectField(label, options) {
    const field = div('field');
    const lab = document.createElement('label');
    lab.textContent = label;
    const select = document.createElement('select');
    options.forEach((opt) => {
      const option = document.createElement('option');
      option.value = opt.value;
      option.textContent = opt.label;
      select.append(option);
    });
    field.append(lab, select);
    return { field, select };
  }

  function makeCheckboxField(label) {
    const field = div('field');
    const wrapper = document.createElement('label');
    const input = document.createElement('input');
    input.type = 'checkbox';
    wrapper.append(input, document.createTextNode(` ${label}`));
    field.append(wrapper);
    return { field, input };
  }

  function cardBtn(label, onClick, variant = 'btn--secondary') {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `btn ${variant}`;
    btn.textContent = label;
    if (onClick) btn.addEventListener('click', onClick);
    return btn;
  }

  function buildSharedParamStatusBanner() {
    const banner = div('sharedparam-banner');
    const head = div('sharedparam-banner__head');
    const title = document.createElement('strong');
    title.textContent = 'Shared Parameter 상태';
    const badge = document.createElement('span');
    badge.className = 'sharedparam-banner__badge';
    const refresh = document.createElement('button');
    refresh.type = 'button';
    refresh.className = 'btn btn--ghost sharedparam-banner__refresh';
    refresh.textContent = '상태 새로고침';
    refresh.addEventListener('click', () => {
        requestSharedParamStatus('manual');
        requestSharedParamList('manual');
    });
    head.append(title, badge, refresh);

    const path = div('sharedparam-banner__path');
    const pathLabel = document.createElement('span');
    pathLabel.textContent = '경로';
    const pathValue = document.createElement('span');
    pathValue.className = 'sharedparam-banner__value';
    path.append(pathLabel, pathValue);

    const note = div('sharedparam-banner__note');

    banner.append(head, path, note);
    state.ui.sharedParamBanner = {
      root: banner,
      badge,
      pathValue,
      note
    };
    return banner;
  }

  function markStale(key) {
    state.results[key].stale = true;
    state.results[key].count = 0;
    state.results[key].hasRun = false;
    syncFeatureRow(key);
    updateFeatureSummary(key);
    post('hub:multi-clear', { key });
  }

  function markAllStale() {
    FEATURE_KEYS.forEach(markStale);
  }

  function syncFeatureRow(key) {
    updateSelectedFeatureRow(key);
    if (key === 'connector') refreshConnectorFeatureSummary();
  }

  function updateResultSummary(summary) {
    Object.keys(summary || {}).forEach((key) => {
      if (!state.results[key]) return;
      state.results[key].count = summary[key].rows || 0;
      state.results[key].stale = false;
      state.results[key].hasRun = true;
      syncFeatureRow(key);
    });
    state.ui.lastRunAt = new Date().toISOString();
    updateBqcSidebar();
  }

  function handleRunAction() {
    if (state.ui.runCompleted) {
      resetRunResults();
      return;
    }
    onRun();
  }

  function handleRunCurrentFile() {
    if (state.ui.runCompleted) {
      resetRunResults();
      return;
    }
    onRunCurrentFile();
  }

  function onRunCurrentFile() {
    state.ui.runCompleted = false;
    updateRunActionLabel();
    const blockReason = getRunBlockingReason({ requireRvt: false });
    if (blockReason) {
      toast(blockReason, 'warn');
      return;
    }
    setBusyState(true);
    ProgressDialog.show('현재 파일 검토', '준비 중...');
    ProgressDialog.update(0, '준비 중...', '');
    const payload = buildPayload();
    payload.useActiveDocument = true;
    payload.rvtPaths = [];
    post('hub:multi-run', payload);
  }

  function updateCurrentDocBtnState() {
    const blockReason = getRunBlockingReason({ requireRvt: false });
    (state.ui.currentDocButtons || []).forEach((btn) => {
      if (!btn) return;
      btn.disabled = state.busy || !!blockReason;
      btn.title = btn.disabled ? blockReason : '';
    });
  }


  function onRun() {
    state.ui.runCompleted = false;
    updateRunActionLabel();
    const blockReason = getRunBlockingReason({ requireRvt: true });
    if (blockReason) {
      toast(blockReason, 'warn');
      return;
    }
    setBusyState(true);
    ProgressDialog.show('납품시 BQC 검토', '준비 중...');
    ProgressDialog.update(0, '준비 중...', '');
    post('hub:multi-run', buildPayload());
  }

  function getRunBlockingReason(options = {}) {
    const silent = !!options.silent;
    const selected = FEATURE_KEYS.filter((k) => state.features[k].enabled);
    if (!selected.length) return '선택된 기능이 없습니다.';
    const needsShared = FEATURE_KEYS.some((key) => state.features[key].enabled && requiresSharedParams(key));
    const status = state.sharedParamStatus || {};
    if (needsShared && status.status !== 'ok') {
      if (!silent && !status.status) {
        requestSharedParamStatus('run');
        return 'Shared Parameter 상태를 확인 중입니다.';
      }
      if (!silent && status.status && status.status !== 'ok') {
        requestSharedParamStatus('run');
      }
      return status.warning || status.errorMessage || 'Shared Parameter 상태를 확인하세요.';
    }
    if (state.features.familylink.enabled) {
      const targets = state.features.familylink.configCommitted.selectedTargets || [];
      if (!targets.length) return '패밀리 공유파라미터 연동 검토 대상이 없습니다.';
    }
    if (options.requireRvt && !state.rvtList.length) return 'RVT 파일을 추가하세요.';
    return '';
  }

  function buildPayload() {
    return {
      rvtPaths: state.rvtList.slice(),
      commonOptions: state.common.configCommitted,
      features: {
        connector: buildCommittedFeature('connector'),
        guid: buildCommittedFeature('guid'),
        familylink: buildCommittedFeature('familylink'),
        points: buildCommittedFeature('points')
      }
    };
  }

  function onExport(key) {
    setBusyState(true);
    chooseExcelMode((mode) => {
      post('hub:multi-export', { key, excelMode: mode || 'fast' });
    });
  }

  function updateOpenMultiBtnState() {
    const blockReason = getOpenMultiBlockingReason();
    (state.ui.openMultiButtons || []).forEach((btn) => {
      if (!btn) return;
      btn.disabled = !!blockReason;
      btn.title = blockReason || '';
    });
  }

  function getOpenMultiBlockingReason() {
    if (state.busy) return '작업 진행 중입니다.';
    const enabledCount = FEATURE_KEYS.filter((k) => state.features[k].enabled).length;
    if (!enabledCount) return '선택된 기능이 있을 때만 RVT 여러개 검토를 열 수 있습니다.';
    return '';
  }

  function updateMultiRunBtnState() {
    const blockReason = getRunBlockingReason({ requireRvt: true, silent: true });
    (state.ui.modalRunButtons || []).forEach((btn) => {
      if (!btn) return;
      btn.disabled = state.busy || !!blockReason;
      btn.title = btn.disabled ? blockReason : '';
    });
  }

  function updateBqcSidebar() {
    const rows = getRecentResultRows();
    const lastRunAt = getCurrentModeLastRunAt();
    const runAt = lastRunAt ? new Date(lastRunAt) : null;
    const captionText = runAt && !Number.isNaN(runAt.getTime())
      ? `마지막 실행 ${runAt.toLocaleString('ko-KR')}`
      : '아직 실행 결과가 없습니다.';
    const hintText = rows.length
      ? '영역을 눌러 최근 결과 상세 창을 열고, 검토별 엑셀을 각각 내보낼 수 있습니다.'
      : '검토 실행 후 이 영역에서 최근 결과 상세 창을 열 수 있습니다.';

    renderRecentResultTable([
      {
        caption: state.ui.modalRecentCaption,
        hint: state.ui.modalRecentHint,
        tbody: state.ui.modalRecentTableBody,
        empty: state.ui.modalRecentEmpty
      }
    ]);

    if (state.ui.bqcRecentCaption) {
      state.ui.bqcRecentCaption.textContent = captionText;
    }

    if (state.ui.bqcRecentHint) {
      state.ui.bqcRecentHint.textContent = hintText;
    }

    if (state.ui.bqcRecentOpenBtn) {
      state.ui.bqcRecentOpenBtn.disabled = state.busy;
      state.ui.bqcRecentOpenBtn.title = rows.length ? '' : '현재 화면의 최근 결과를 확인합니다.';
    }

    (state.ui.resultResetButtons || []).forEach((btn) => {
      if (!btn) return;
      btn.disabled = state.busy || !state.ui.reviewSummaryData;
    });
  }

  function setBusyState(on) {
    state.busy = on;
    setBusy(on);
    FEATURE_KEYS.forEach(syncFeatureRow);
    const inputs = page.querySelectorAll('input, select, textarea, button');
    inputs.forEach((el) => {
      if (on) {
        el.disabled = true;
      } else {
        el.disabled = false;
      }
    });
    if (!on) renderRvtList();
    if (!on) updateSharedParamRunState();
    updateCurrentDocBtnState();
    updateOpenMultiBtnState();
    updateMultiRunBtnState();
    updateBqcSidebar();
  }

  function renderRvtList() {
    if (buildRvtSection.render) buildRvtSection.render();
  }

  function openSettings(key, title) {
    const config = key === 'common' ? state.common : state.features[key];
    if (!buildSettingsModal.form) return;
    state.ui.modalOpen = true;
    state.ui.activeFeatureKey = key;
    state.ui.activeFeatureTitle = title || '';
    if (buildSettingsModal.modal && buildSettingsModal.body && buildSettingsModal.help && buildSettingsModal.form) {
      if (key === 'connector') {
        buildSettingsModal.modal.style.width = 'min(1320px, 95vw)';
        buildSettingsModal.modal.style.maxWidth = '1320px';
        buildSettingsModal.body.style.display = 'grid';
        buildSettingsModal.body.style.gridTemplateColumns = 'minmax(0, 2.18fr) minmax(280px, 0.82fr)';
        buildSettingsModal.body.style.alignItems = 'start';
        buildSettingsModal.body.style.gap = '10px';
        buildSettingsModal.body.style.columnGap = '10px';
        buildSettingsModal.form.style.display = 'block';
        buildSettingsModal.form.style.gridTemplateColumns = '';
        buildSettingsModal.form.style.gridAutoFlow = '';
        buildSettingsModal.form.style.gap = '';
        buildSettingsModal.form.style.alignContent = 'start';
        buildSettingsModal.form.style.justifyItems = 'stretch';
        buildSettingsModal.form.style.flex = '';
        buildSettingsModal.form.style.minWidth = '0';
        buildSettingsModal.form.style.width = '100%';
        buildSettingsModal.form.style.maxWidth = 'none';
        buildSettingsModal.form.style.margin = '0';
        buildSettingsModal.form.style.justifySelf = 'stretch';
        buildSettingsModal.help.style.display = 'grid';
        buildSettingsModal.help.style.gap = '10px';
        buildSettingsModal.help.style.alignContent = 'start';
        buildSettingsModal.help.style.flex = '';
        buildSettingsModal.help.style.minWidth = '0';
        buildSettingsModal.help.style.width = '100%';
        buildSettingsModal.help.style.maxWidth = 'none';
        buildSettingsModal.help.style.margin = '0';
        buildSettingsModal.help.style.justifySelf = 'stretch';
        buildSettingsModal.help.style.alignSelf = 'start';
      } else if (key === 'common') {
        buildSettingsModal.modal.style.width = 'min(980px, 92vw)';
        buildSettingsModal.modal.style.maxWidth = '980px';
        buildSettingsModal.body.style.display = 'grid';
        buildSettingsModal.body.style.gridTemplateColumns = 'minmax(0, 1.12fr) minmax(300px, 0.88fr)';
        buildSettingsModal.body.style.alignItems = 'start';
        buildSettingsModal.body.style.gap = '14px';
        buildSettingsModal.body.style.columnGap = '14px';
        buildSettingsModal.form.style.display = 'grid';
        buildSettingsModal.form.style.gridTemplateColumns = 'minmax(0, 1fr)';
        buildSettingsModal.form.style.gridAutoFlow = 'row';
        buildSettingsModal.form.style.gap = '12px';
        buildSettingsModal.form.style.alignContent = 'start';
        buildSettingsModal.form.style.justifyItems = 'stretch';
        buildSettingsModal.form.style.flex = '';
        buildSettingsModal.form.style.minWidth = '0';
        buildSettingsModal.form.style.width = '100%';
        buildSettingsModal.form.style.maxWidth = '620px';
        buildSettingsModal.form.style.margin = '0';
        buildSettingsModal.form.style.justifySelf = 'start';
        buildSettingsModal.help.style.display = 'grid';
        buildSettingsModal.help.style.gap = '10px';
        buildSettingsModal.help.style.alignContent = 'start';
        buildSettingsModal.help.style.flex = '';
        buildSettingsModal.help.style.minWidth = '0';
        buildSettingsModal.help.style.width = '100%';
        buildSettingsModal.help.style.maxWidth = '380px';
        buildSettingsModal.help.style.margin = '0';
        buildSettingsModal.help.style.justifySelf = 'stretch';
        buildSettingsModal.help.style.alignSelf = 'start';
      } else {
        buildSettingsModal.modal.style.width = 'min(1220px, 96vw)';
        buildSettingsModal.modal.style.maxWidth = '1220px';
        buildSettingsModal.body.style.display = 'grid';
        buildSettingsModal.body.style.gridTemplateColumns = 'minmax(0, 2.35fr) minmax(300px, 1fr)';
        buildSettingsModal.body.style.alignItems = 'start';
        buildSettingsModal.body.style.gap = '14px';
        buildSettingsModal.form.style.flex = '';
        buildSettingsModal.form.style.minWidth = '0';
        buildSettingsModal.form.style.width = '100%';
        buildSettingsModal.form.style.maxWidth = 'none';
        buildSettingsModal.form.style.justifySelf = 'stretch';
        buildSettingsModal.help.style.flex = '';
        buildSettingsModal.help.style.minWidth = '0';
        buildSettingsModal.help.style.width = '100%';
        buildSettingsModal.help.style.maxWidth = 'none';
        buildSettingsModal.help.style.justifySelf = 'stretch';
        buildSettingsModal.help.style.alignSelf = 'start';
      }
      buildSettingsModal.body.style.width = '100%';
      buildSettingsModal.body.style.justifyContent = 'stretch';
    }
    buildSettingsModal.title.textContent = `${title || ''} 설정`;
    const readiness = key === 'common' ? { label: '설정', className: 'chip--ok' } : getFeatureReadiness(config);
    if (buildSettingsModal.badge) {
      buildSettingsModal.badge.textContent = readiness.label;
      buildSettingsModal.badge.className = `chip ${readiness.className}`;
      buildSettingsModal.badge.style.display = readiness.className === 'chip--warn' ? 'inline-flex' : 'none';
    }
    buildSettingsModal.form.innerHTML = '';
    buildSettingsModal.help.innerHTML = '';
    resetDraftFromCommitted(key);
    syncControlsFromDraft(key);
    renderSharedParamBanner(key);
    if (key === 'connector') requestConnectorParamList('settings-open');
    const panel = getFeaturePanel(key);
    if (panel) {
      if (key === 'connector') {
        panel.style.display = 'flex';
        panel.style.flexDirection = 'column';
        panel.style.alignItems = 'stretch';
        panel.style.justifyContent = 'flex-start';
        panel.style.width = '100%';
        panel.style.maxWidth = 'none';
        panel.style.minWidth = '0';
        panel.style.margin = '0';
        panel.style.boxSizing = 'border-box';
        Array.from(panel.children || []).forEach((child) => {
          if (!child || !child.style) return;
          child.style.width = '100%';
          child.style.maxWidth = 'none';
          child.style.minWidth = '0';
          child.style.margin = '0';
          child.style.boxSizing = 'border-box';
        });
      }
      buildSettingsModal.form.append(panel);
    }
    renderHelp(key, title);
    if (key === 'connector' && buildSettingsModal.help) {
      buildSettingsModal.help.style.width = '100%';
      buildSettingsModal.help.style.maxWidth = 'none';
      buildSettingsModal.help.style.minWidth = '0';
      buildSettingsModal.help.style.margin = '0';
    }
    buildSettingsModal.overlay.classList.add('is-open');
  }

  function closeSettings() {
    if (!buildSettingsModal.overlay) return;
    state.ui.modalOpen = false;
    buildSettingsModal.overlay.classList.remove('is-open');
  }

  function applySettings() {
    const key = state.ui.activeFeatureKey;
    if (!key) return;
    if (key === 'common') {
      commitConfig(state.common);
      updateCommonSummary();
      persistCommonOptions(state.common.configCommitted);
      emitCommonOptionsChanged();
      markStale('connector');
    } else {
      commitConfig(state.features[key]);
      markStale(key);
    }
    closeSettings();
  }

  function cancelSettings() {
    const key = state.ui.activeFeatureKey;
    if (!key) return;
    resetDraftFromCommitted(key);
    syncControlsFromDraft(key);
    closeSettings();
  }

  function getFeaturePanel(key) {
    return state.ui.panels[key] || null;
  }

  function updateRunSummary() {
    renderSelectedFeatures();
    renderModalFeatureSummary();
    updateActionSummaryVisibility();
    updateBqcSidebar();
    updateSharedParamRunState();
  }

  function updateRunProgress(percent, message, detail) {
    if (!buildRunBar.progressText) return;
    buildRunBar.progressText.textContent = message || '대기 중';
    buildRunBar.progressDetail.textContent = detail || '';
    buildRunBar.progressDetail.style.display = detail ? 'block' : 'none';
    if (buildRunBar.progressFill) {
      const pct = Math.max(0, Math.min(100, Number(percent) || 0));
      buildRunBar.progressFill.style.width = `${pct}%`;
    }
  }

  function updateActionSummaryVisibility() {
    if (!buildRunBar.summary) return;
    const hasHint = !!(
      buildRunBar.runSharedParamHint &&
      buildRunBar.runSharedParamHint.style.display !== 'none' &&
      String(buildRunBar.runSharedParamHint.textContent || '').trim()
    );
    buildRunBar.summary.style.display = hasHint ? 'flex' : 'none';
  }

  function updateRunActionLabel() {
    if (!buildRunBar.startBtn) return;
    buildRunBar.startBtn.textContent = state.ui.runCompleted ? '검토 결과 초기화' : '검토 시작';
  }

  function resetRunResults() {
    state.ui.runCompleted = false;
    state.ui.lastProgressPct = 0;
    state.ui.reviewSummaryData = null;
    state.ui.lastRunAt = null;
    state.ui.reviewSummaryByMode = { bqc: null, utility: null };
    state.ui.lastRunAtByMode = { bqc: null, utility: null };
    updateRunProgress(0, '대기 중', '');
    FEATURE_KEYS.forEach((key) => {
      if (state.results[key]) {
        state.results[key].count = 0;
        state.results[key].stale = true;
        state.results[key].hasRun = false;
      }
    });
    syncFeatureRow('connector');
    syncFeatureRow('guid');
    syncFeatureRow('familylink');
    syncFeatureRow('points');
    updateRunActionLabel();
    updateBqcSidebar();
    post('hub:multi-clear', {});
  }

  function getFeatureReadiness(feature) {
    if (!feature?.enabled) {
      return { label: 'OFF', className: 'chip--off' };
    }
    if (!feature.applied || feature.dirty) {
      return { label: '설정 필요', className: 'chip--warn' };
    }
    if (state.ui.activeFeatureKey === 'familylink') {
      const targets = feature.configCommitted.selectedTargets || [];
      const sharedOk = state.sharedParamStatus?.status === 'ok';
      if (!sharedOk || targets.length < 1) {
        return { label: '설정 필요', className: 'chip--warn' };
      }
    }
    return { label: '검토 준비됨', className: 'chip--ok' };
  }

  function updateDrawerBadge(key) {
    if (!state.ui.modalOpen || state.ui.activeFeatureKey !== key || !buildSettingsModal.badge) return;
    const readiness = getFeatureReadiness(state.features[key]);
    buildSettingsModal.badge.textContent = readiness.label;
    buildSettingsModal.badge.className = `chip ${readiness.className}`;
    buildSettingsModal.badge.style.display = readiness.className === 'chip--warn' ? 'inline-flex' : 'none';
  }

  function updateFeatureSummary(key) {
    updateSelectedFeatureRow(key);
    updateDrawerBadge(key);
    if (key === 'connector') refreshConnectorFeatureSummary();
  }

  function buildCommonSummary() {
    const committed = state.common.configCommitted;
    const extraCount = committed.extraParams ? committed.extraParams.split(',').filter((v) => v.trim()).length : 0;
    const filterText = committed.targetFilter ? committed.targetFilter : '필터 없음';
    const excludeText = committed.excludeEndDummy ? 'Dummy 제외' : 'Dummy 포함';
    return `extra=${extraCount} / ${filterText} / ${excludeText}`;
  }

  function updateCommonSummary(el) {
    const target = el || state.ui.commonSummaryEl;
    if (target) {
      target.textContent = buildCommonSummary();
    }
    if (state.ui.actionCommonSummaryEl) {
      state.ui.actionCommonSummaryEl.textContent = `공통 설정: ${buildCommonSummary()}`;
    }
    updateActionSummaryVisibility();
    updateFeatureSummary('connector');
  }

  function renderHelp(key, title) {
    const help = buildSettingsModal.help;
    if (!help) return;
    help.style.display = 'grid';
    help.style.gap = '10px';
    help.style.alignContent = 'start';
    help.style.alignItems = 'stretch';
    help.style.minWidth = '0';
    help.style.width = '100%';
    help.style.maxWidth = 'none';
    const helpTitle = document.createElement('strong');
    helpTitle.textContent = title || '설정 안내';
    helpTitle.style.display = 'block';
    helpTitle.style.width = '100%';
    const list = document.createElement('ul');
    list.className = 'help-list';
    list.style.margin = '0';
    list.style.padding = '0';
    list.style.listStyle = 'none';
    list.style.display = 'grid';
    list.style.gap = '10px';
    list.style.width = '100%';
    list.style.boxSizing = 'border-box';
    getHelpItems(key).forEach((text) => {
      const item = document.createElement('li');
      item.textContent = text;
      item.style.width = '100%';
      item.style.boxSizing = 'border-box';
      item.style.margin = '0';
      item.style.padding = '12px 14px';
      item.style.borderRadius = '14px';
      item.style.border = '1px solid var(--border-soft)';
      item.style.background = 'var(--surface-help)';
      list.append(item);
    });
    help.append(helpTitle, list);
  }

  function getHelpItems(key) {
    if (key === 'common') {
      return [
        '추가 Parameter 값은 콤마로 구분해 입력합니다.',
        '검토 대상 필터는 “PM1=Value” 형식으로 입력합니다.',
        'Dummy/End_ 패밀리 제외 여부를 설정합니다.'
      ];
    }
    if (key === 'connector') {
      return [
        '공유 파라미터 txt 목록에서 검토 대상을 검색해 선택합니다.',
        '여러 파라미터를 선택하면 같은 논리로 연속성 검토를 진행합니다.',
        '허용범위와 단위는 기존 커넥터 검토 로직 그대로 적용됩니다.'
      ];
    }
    if (key === 'guid') {
      return [
        '패밀리/Annotation 포함 여부를 선택합니다.',
        '공유 파라미터 GUID 일치 여부를 검토합니다.'
      ];
    }
    if (key === 'familylink') {
      return [
        '공유 파라미터 목록에서 검토 대상 파라미터를 선택합니다.',
        '고급 입력으로 “이름|GUID” 형식을 직접 입력할 수 있습니다.'
      ];
    }
    if (key === 'points') {
      return [
        '좌표 추출 단위를 선택합니다.',
        'Decimal Feet / Meter / Millimeter를 지원합니다.'
      ];
    }
    return [];
  }

  function createFeatureState(config) {
    return {
      enabled: false,
      applied: false,
      dirty: false,
      configCommitted: deepCopy(config),
      configDraft: deepCopy(config)
    };
  }

  function createConfigState(config) {
    return {
      applied: false,
      dirty: false,
      configCommitted: deepCopy(config),
      configDraft: deepCopy(config)
    };
  }

  function deepCopy(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  function parseFamilyLinkTargets(text) {
    const lines = String(text || '').split(/\r?\n/);
    const targets = [];
    lines.forEach((line) => {
      const trimmed = line.trim();
      if (!trimmed) return;
      const parts = trimmed.split('|');
      const name = (parts[0] || '').trim();
      const guid = (parts[1] || '').trim();
      if (!name || !guid) return;
      targets.push({ name, guid });
    });
    return targets;
  }

  function dedupeTargets(items) {
    const byGuid = new Map();
    items.forEach((item) => {
      if (!item || !item.guid) return;
      byGuid.set(item.guid, item);
    });
    return Array.from(byGuid.values());
  }

  function buildTargetsText(items) {
    return (items || [])
      .filter((item) => item && item.name && item.guid)
      .map((item) => `${item.name}|${item.guid}`)
      .join('\n');
  }

  function normalizeConnectorParamNames(raw) {
    const arr = Array.isArray(raw) ? raw : String(raw || '').split(',');
    const seen = new Set();
    const out = [];
    arr.forEach((value) => {
      const name = String(value && value.name ? value.name : value || '').trim();
      if (!name) return;
      const key = name.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      out.push(name);
    });
    return out;
  }

  function normalizeConnectorParamItems(payload) {
    const seen = new Set();
    const items = [];
    const push = (name, groupName, guid, source) => {
      const clean = String(name || '').trim();
      if (!clean) return;
      const key = clean.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      items.push({ name: clean, groupName: String(groupName || '').trim(), guid: String(guid || '').trim(), source: String(source || '').trim() });
    };
    const rawItems = Array.isArray(payload?.items) ? payload.items : [];
    rawItems.forEach((item) => push(item?.name, item?.groupName, item?.guid, item?.source));
    const rawParams = Array.isArray(payload?.params) ? payload.params : [];
    rawParams.forEach((name) => push(name, '', '', 'fallback'));
    items.sort((a, b) => a.name.localeCompare(b.name, 'ko'));
    return items;
  }

  function refreshConnectorFeatureSummary() {
    const target = state.ui.connectorHeroSummary;
    if (!target) return;
    const feature = state.features.connector;
    const selected = normalizeConnectorParamNames(
      feature.configDraft?.paramItems && feature.configDraft.paramItems.length
        ? feature.configDraft.paramItems
        : feature.configCommitted?.paramItems && feature.configCommitted.paramItems.length
          ? feature.configCommitted.paramItems
          : feature.configCommitted?.param
    );
    const unit = feature.configDraft?.unit || feature.configCommitted?.unit || 'inch';
    const tol = feature.configDraft?.tol || feature.configCommitted?.tol || 1;
    const selectedText = selected.length ? selected.join(', ') : '선택 없음';
    target.top.textContent = feature.enabled
      ? '체크됨 · 설정창에서 검토 파라미터를 선택하거나 수정할 수 있습니다.'
      : '체크 후 옵션을 열어 공유 파라미터 목록에서 검토 대상을 선택하세요.';
    target.sub.textContent = `선택 파라미터 ${selected.length}개 · ${selectedText} · 허용범위 ${tol} ${unit}`;
    target.row.classList.toggle('is-active', !!feature.enabled);
  }

  function renderSelectedFeatures() {
    if (!state.ui.selectedTableBody) return;
    const enabledKeys = FEATURE_KEYS.filter((key) => state.features[key].enabled);
    state.ui.selectedRows.clear();
    state.ui.selectedTableBody.innerHTML = '';
    if (state.ui.selectedSection) {
      state.ui.selectedSection.classList.toggle('selected-panel--has-selection', enabledKeys.length > 0);
    }

    if (enabledKeys.length === 0) {
      const row = document.createElement('tr');
      const cell = document.createElement('td');
      cell.colSpan = 3;
      cell.className = 'selected-empty';
      cell.textContent = '선택된 기능이 없습니다.';
      row.append(cell);
      state.ui.selectedTableBody.append(row);
    }

    enabledKeys.forEach((key) => {
      const row = document.createElement('tr');
      row.dataset.key = key;
      const nameCell = document.createElement('td');
      const nameWrap = div('selected-name');
      const nameMain = document.createElement('strong');
      nameMain.textContent = FEATURE_META[key]?.label || key;
      nameMain.title = FEATURE_META[key]?.label || key;
      const nameSub = document.createElement('span');
      nameSub.textContent = FEATURE_META[key]?.desc || '';
      nameSub.title = FEATURE_META[key]?.desc || '';
      nameWrap.append(nameMain, nameSub);
      nameCell.append(nameWrap);

      const statusCell = document.createElement('td');
      statusCell.className = 'selected-status-col selected-action-col';
      const statusChip = document.createElement('span');
      statusChip.className = 'chip status-chip';
      statusCell.append(statusChip);

      const settingsCell = document.createElement('td');
      settingsCell.className = 'selected-action-col';
      const settingsBtn = document.createElement('button');
      settingsBtn.type = 'button';
      settingsBtn.className = 'btn btn--secondary';
      settingsBtn.textContent = '설정';
      settingsBtn.addEventListener('click', () => openSettings(key, FEATURE_META[key]?.label));
      settingsCell.append(settingsBtn);

      row.append(nameCell, statusCell, settingsCell);
      state.ui.selectedTableBody.append(row);

      state.ui.selectedRows.set(key, { row, statusChip });
      updateSelectedFeatureRow(key);
    });

    if (state.ui.selectedCount) {
      state.ui.selectedCount.textContent = `${enabledKeys.length}개`;
      state.ui.selectedCount.className = `chip ${enabledKeys.length ? 'chip--ok' : 'chip--info'}`;
    }

    updateCurrentDocBtnState();
    updateMultiRunBtnState();
  }

  function updateSelectedFeatureRow(key) {
    const entry = state.ui.selectedRows.get(key);
    const status = getSelectedFeatureStatus(key);
    if (entry) {
      entry.statusChip.textContent = status.label;
      entry.statusChip.className = `chip status-chip ${status.className}`;
    }
    renderModalFeatureSummary();
    updateBqcSidebar();
  }

  function getSelectedFeatureStatus(key) {
    const feature = state.features[key];
    if (!feature) return { label: '검토 전', className: 'status-chip--idle' };
    if (requiresSharedParams(key) && state.sharedParamStatus?.status && state.sharedParamStatus.status !== 'ok') {
      return { label: 'Shared Param 확인 필요', className: 'status-chip--warn' };
    }
    if (key === 'familylink') {
      const targets = feature.configCommitted.selectedTargets || [];
      if (!targets.length) {
        return { label: '설정 필요', className: 'status-chip--warn' };
      }
    }
    if (state.busy) {
      return { label: '진행 중', className: 'status-chip--running' };
    }
    const res = state.results[key];
    if (res && res.hasRun && !res.stale) {
      return { label: '완료', className: 'status-chip--done' };
    }
    if (!feature.applied || feature.dirty) {
      return { label: '검토 전', className: 'status-chip--idle' };
    }
    return { label: '검토 준비됨', className: 'status-chip--ready' };
  }

  function requiresSharedParams(key) {
    return !!FEATURE_META[key]?.requiresSharedParams;
  }

  function getGroupFilter() {
    try {
      return localStorage.getItem(GROUP_FILTER_KEY) || 'all';
    } catch {
      return 'all';
    }
  }

  function saveGroupFilter(value) {
    try {
      localStorage.setItem(GROUP_FILTER_KEY, value);
    } catch {
    }
  }

  function renderGroupVisibility() {
    const filter = state.ui.groupFilter || 'all';
    const mode = state.ui.multiMode || 'bqc';
    const sections = page.querySelectorAll('.multi-section[data-group]');
    sections.forEach((section) => {
      const group = section.dataset.group || '';
      const allowGroup = mode === 'utility' ? group === 'utility' : group === 'bqc';
      const show = allowGroup && (filter === 'all' || group === filter || filter === mode);
      section.classList.toggle('is-hidden', !show);
    });
  }

  function requestSharedParamStatus(context) {
    post('sharedparam:status', { source: 'multi', context });
  }

  function updateSharedParamBanner() {
    const banner = buildSettingsModal.sharedBanner || state.ui.sharedParamBanner?.root;
    if (!banner || banner.style.display === 'none') return;
    const status = state.sharedParamStatus || {};
    const label = status.statusLabel || '조회 중';
    const badge = state.ui.sharedParamBanner?.badge;
    const pathValue = state.ui.sharedParamBanner?.pathValue;
    const note = state.ui.sharedParamBanner?.note;
    if (badge) {
      badge.textContent = label;
      badge.classList.remove('is-ok', 'is-warn', 'is-error');
      if (status.status === 'ok') badge.classList.add('is-ok');
      else if (status.status === 'warn' || status.status === 'unset' || status.status === 'missing') badge.classList.add('is-warn');
      else badge.classList.add('is-error');
    }
    if (pathValue) {
      const pathText = status.path || '미설정';
      pathValue.textContent = pathText;
      pathValue.title = pathText;
    }
    if (note) {
      const warning = status.warning || status.errorMessage || '';
      note.textContent = warning ? warning : '';
      note.style.display = warning ? 'block' : 'none';
    }
  }

  function renderSharedParamBanner(key) {
    const banner = buildSettingsModal.sharedBanner;
    if (!banner) return;
    if (!requiresSharedParams(key)) {
      banner.style.display = 'none';
      return;
    }
    banner.style.display = 'block';
    buildSettingsModal.form.append(banner);
    requestSharedParamStatus('settings');
    requestSharedParamList('settings');
    updateSharedParamBanner();
  }

  function updateSharedParamRunState() {
    const needsShared = FEATURE_KEYS.some((key) => state.features[key].enabled && requiresSharedParams(key));
    const ok = state.sharedParamStatus?.status === 'ok';
    const familyLinkTargets = state.features.familylink?.configCommitted?.selectedTargets || [];
    const familyLinkNeedsTargets = state.features.familylink?.enabled && familyLinkTargets.length < 1;
    if (buildRunBar.runSharedParamHint) {
      if (needsShared && !ok) {
        const warning = state.sharedParamStatus?.warning || 'Shared Parameter 미등록으로 실행이 제한됩니다.';
        buildRunBar.runSharedParamHint.textContent = warning;
        buildRunBar.runSharedParamHint.style.display = 'block';
      } else if (familyLinkNeedsTargets) {
        buildRunBar.runSharedParamHint.textContent = '패밀리 공유파라미터 검토 대상이 없습니다.';
        buildRunBar.runSharedParamHint.style.display = 'block';
      } else {
        buildRunBar.runSharedParamHint.style.display = 'none';
      }
    }
    updateActionSummaryVisibility();
    if (!state.busy) {
      updateCurrentDocBtnState();
      updateMultiRunBtnState();
    }
  }

  function canRunWithSharedParams() {
    const needsShared = FEATURE_KEYS.some((key) => state.features[key].enabled && requiresSharedParams(key));
    if (!needsShared) return true;
    const status = state.sharedParamStatus || {};
    if (status.status === 'ok') return true;
    if (!status.status) {
      requestSharedParamStatus('run');
      toast('Shared Parameter 상태를 확인 중입니다.', 'warn');
      return false;
    }
    requestSharedParamStatus('run');
    const msg = status.warning || status.errorMessage || 'Shared Parameter 상태를 확인하세요.';
    toast(msg, 'warn');
    return false;
  }

  function normalizeCommonOptions(raw) {
    return {
      extraParams: typeof raw?.extraParamsText === 'string' ? raw.extraParamsText : '',
      targetFilter: typeof raw?.targetFilterText === 'string' ? raw.targetFilterText : '',
      excludeEndDummy: !!raw?.excludeEndDummy
    };
  }

  function loadCommonOptionsFromStorage() {
    let stored = null;
    try {
      const raw = localStorage.getItem(COMMON_OPTIONS_KEY);
      if (raw) stored = JSON.parse(raw);
    } catch {
      stored = null;
    }
    if (stored) {
      applyCommonOptionsFromStorage(stored);
      return true;
    }
    post('commonoptions:get', { source: 'multi' });
    return false;
  }

  function applyCommonOptionsFromStorage(stored) {
    const normalized = normalizeCommonOptions(stored);
    state.common.configCommitted = deepCopy(normalized);
    state.common.configDraft = deepCopy(normalized);
    state.common.applied = true;
    state.common.dirty = false;
    syncControlsFromDraft('common');
    updateCommonSummary();
  }

  function persistCommonOptions(committed, options = {}) {
    const payload = {
      extraParamsText: committed.extraParams || '',
      targetFilterText: committed.targetFilter || '',
      excludeEndDummy: !!committed.excludeEndDummy
    };
    try {
      localStorage.setItem(COMMON_OPTIONS_KEY, JSON.stringify(payload));
    } catch {
    }
    if (!options.skipHost) {
      post('commonoptions:save', payload);
    }
  }

  function emitCommonOptionsChanged() {
    const detail = { ...state.common.configCommitted };
    window.dispatchEvent(new CustomEvent('commonOptions:changed', { detail }));
  }

  function buildCommittedFeature(key) {
    const feature = state.features[key];
    if (key === 'familylink') {
      return {
        enabled: feature.enabled,
        targets: feature.configCommitted.selectedTargets || []
      };
    }
    if (key === 'connector') {
      const committed = deepCopy(feature.configCommitted || {});
      committed.paramItems = normalizeConnectorParamNames(committed.paramItems && committed.paramItems.length ? committed.paramItems : committed.param);
      committed.param = committed.paramItems.join(',') || 'Comments';
      return {
        enabled: feature.enabled,
        ...committed
      };
    }
    return {
      enabled: feature.enabled,
      ...feature.configCommitted
    };
  }

  function commitConfig(target) {
    if (state.ui.activeFeatureKey === 'familylink') {
      const parsed = parseFamilyLinkTargets(target.configDraft.targetsText);
      target.configDraft.selectedTargets = dedupeTargets([...target.configDraft.selectedTargets, ...parsed]);
      target.configDraft.targets = target.configDraft.selectedTargets;
      target.configDraft.targetsText = buildTargetsText(target.configDraft.selectedTargets);
    }
    if (state.ui.activeFeatureKey === 'connector') {
      target.configDraft.paramItems = normalizeConnectorParamNames(target.configDraft.paramItems && target.configDraft.paramItems.length ? target.configDraft.paramItems : target.configDraft.param);
      target.configDraft.param = target.configDraft.paramItems.join(',') || 'Comments';
    }
    target.configCommitted = deepCopy(target.configDraft);
    target.applied = true;
    target.dirty = false;
    if (state.ui.activeFeatureKey !== 'common') {
      updateFeatureSummary(state.ui.activeFeatureKey);
    }
  }

  function resetDraftFromCommitted(key) {
    if (key === 'common') {
      state.common.configDraft = deepCopy(state.common.configCommitted);
      state.common.dirty = false;
      return;
    }
    const feature = state.features[key];
    if (!feature) return;
    feature.configDraft = deepCopy(feature.configCommitted);
    if (key === 'connector') {
      feature.configDraft.paramItems = normalizeConnectorParamNames(feature.configDraft.paramItems && feature.configDraft.paramItems.length ? feature.configDraft.paramItems : feature.configDraft.param);
      feature.configDraft.param = feature.configDraft.paramItems.join(',') || 'Comments';
    }
    if (key === 'familylink') {
      feature.configDraft.targetsText = buildTargetsText(feature.configDraft.selectedTargets);
    }
    feature.dirty = false;
  }

  function markFeatureDirty(key) {
    const feature = state.features[key];
    if (!feature) return;
    feature.dirty = true;
    feature.applied = false;
    markStale(key);
  }

  function markCommonDirty() {
    state.common.dirty = true;
    state.common.applied = false;
    markStale('connector');
  }

  function syncControlsFromDraft(key) {
    const controls = state.ui.controls[key];
    if (!controls) return;
    if (key === 'connector') {
      const draft = state.features.connector.configDraft;
      draft.paramItems = normalizeConnectorParamNames(draft.paramItems && draft.paramItems.length ? draft.paramItems : draft.param);
      draft.param = draft.paramItems.join(',') || 'Comments';
      controls.tol.input.value = draft.tol;
      controls.unit.select.value = draft.unit;
      if (controls.searchInput) controls.searchInput.value = '';
      if (controls.renderConnectorList) controls.renderConnectorList();
      if (controls.renderConnectorSelected) controls.renderConnectorSelected();
    } else if (key === 'guid') {
      const draft = state.features.guid.configDraft;
      controls.includeFamily.input.checked = draft.includeFamily;
      controls.includeAnno.input.checked = draft.includeAnnotation;
    } else if (key === 'familylink') {
      const draft = state.features.familylink.configDraft;
      controls.advanced.input.value = draft.targetsText;
      if (buildFamilyLinkConfig.renderList) buildFamilyLinkConfig.renderList();
    } else if (key === 'points') {
      const draft = state.features.points.configDraft;
      controls.unit.select.value = draft.unit;
    } else if (key === 'common') {
      const draft = state.common.configDraft;
      controls.extra.input.value = draft.extraParams;
      controls.filter.input.value = draft.targetFilter;
      controls.exclude.input.checked = draft.excludeEndDummy;
    }
  }

  function buildFilterExamples() {
    const wrap = div('filter-examples');
    const title = document.createElement('strong');
    title.textContent = '필터 예시';
    const note = document.createElement('p');
    note.textContent = '좌측 Param 토큰은 공백 없는 이름을 권장합니다. 구분자는 콤마(,) 또는 세미콜론(;)을 사용할 수 있습니다.';
    note.className = 'filter-examples__note';
    const list = document.createElement('ul');
    list.className = 'filter-examples__list';

    const examples = [
      "and(PM1='A',PM2='B')",
      "or(SYSTEM='DCW',SYSTEM='DHW')",
      "not(Family='End_Dummy')",
      "and(PM1='A',not(PM2='X'))",
      "PM1='A';PM2='B'"
    ];

    examples.forEach((text) => {
      const item = document.createElement('li');
      const code = document.createElement('code');
      code.textContent = text;
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'btn btn--ghost';
      btn.textContent = '복사';
      btn.addEventListener('click', () => copyToClipboard(text));
      item.append(code, btn);
      list.append(item);
    });

    wrap.append(title, note, list);
    return wrap;
  }

  function copyToClipboard(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => toast('복사되었습니다.', 'ok')).catch(() => toast('복사에 실패했습니다.', 'err'));
      return;
    }
    const temp = document.createElement('textarea');
    temp.value = text;
    temp.style.position = 'fixed';
    temp.style.opacity = '0';
    document.body.append(temp);
    temp.focus();
    temp.select();
    try {
      document.execCommand('copy');
      toast('복사되었습니다.', 'ok');
    } catch (e) {
      toast('복사에 실패했습니다.', 'err');
    } finally {
      temp.remove();
    }
  }
}
