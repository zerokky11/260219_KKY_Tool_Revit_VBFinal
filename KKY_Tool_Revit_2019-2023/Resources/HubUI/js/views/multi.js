import { clear, div, toast, setBusy, showExcelSavedDialog, chooseExcelMode } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { getLastExcelExportLocale } from '../core/dom.js';
import { refreshUiAfterHostDialog } from '../core/hostDialog.js';
import { attachRvtDropZone } from '../core/rvtDrop.js';
import { post, onHost, DEV } from '../core/bridge.js';
import {
  bindHubEntryContextMenu,
  consumePendingHubEntry,
  getFavoriteEntries,
  getHubEntry,
  getHubPanelSearch,
  HUB_QUICK_ACCESS_CHANGE_EVENT,
  openHubEntry,
  replaceHubFavorites,
  recordHubEntryUse,
  searchHubEntries,
  setHubPanelSearch
} from '../core/hubFavorites.js?v=20260504e';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js?v=20260504a';

const FEATURE_META = {
    connector: {
      label: '파라미터 연속성 검토',
      cardLabel: '파라미터 연속성 검토',
      desc: '연결된 MEP 객체의 파라미터 연속성을 검토',
      cardDesc: '연결된 MEP 객체의 파라미터 연속성을 검토',
      categoryLabel: '파라미터',
      categoryTitle: '파라미터 검토/연속성 확인 기능',
      requiresSharedParams: false
    },
    floorinfo: {
      label: '레벨 영역별 파라미터 검토',
      cardLabel: '레벨 영역별 파라미터 검토',
      desc: '선택한 레벨 영역을 기준으로 파라미터 일치 여부를 검토',
      cardDesc: '선택한 레벨 영역을 기준으로 파라미터 일치 여부를 검토',
      categoryLabel: '파라미터',
      categoryTitle: '파라미터 값 검토 기능',
      requiresSharedParams: false
    },
    familysuitability: {
      label: '패밀리 타입 적합성 검토',
      cardLabel: '패밀리 타입 적합성 검토',
      desc: '기준 엑셀의 카테고리/패밀리/타입 조합으로 실제 사용 타입 적합성을 검토',
      cardDesc: '기준 엑셀의 카테고리/패밀리/타입 조합과 실제 사용 객체를 비교해 검토 문구를 출력합니다.',
      categoryLabel: '패밀리',
      categoryTitle: '패밀리/타입 적합성 검토 기능',
      requiresSharedParams: false
    },
    tapalign: {
      label: '탭/분기 축 틀어짐 검토',
      cardLabel: '탭/분기 축 틀어짐 검토',
      desc: '탭/분기 피팅 축이 연결된 배관/덕트 중심축에서 벗어났는지 검토',
      cardDesc: '탭/분기 피팅 축이 연결된 배관/덕트 중심축에서 벗어났는지 검토',
      categoryLabel: 'MEP',
      categoryTitle: 'MEP 축 정렬 검토 기능',
      requiresSharedParams: false
    },
    dupclash: {
      label: '중복 / 자체 간섭 검토',
      cardLabel: '중복 / 자체 간섭 검토',
      desc: '여러 RVT를 대상으로 중복 검토 또는 자체 간섭 검토 중 하나를 선택해 약식 배치 검토',
      cardDesc: '기능 설정에서 검토 모드를 고르면 공통 설정의 포함/제외 대상 필터를 그대로 따라가며 선택한 검토만 실행합니다.',
      categoryLabel: 'MEP',
      categoryTitle: 'MEP 충돌/중복 검토 기능',
      requiresSharedParams: false
    },
    worksetassignment: {
      label: '웍셋 배정 검토',
      cardLabel: '웍셋 배정 검토',
      desc: '모델 객체의 웍셋 배정을 기본 웍셋(Workset1) 또는 입력한 특정 웍셋 기준으로 검토',
      cardDesc: '검토 대상 RVT의 모델 객체가 기본 웍셋(Workset1) 이외 웍셋에 있는지, 또는 입력한 특정 웍셋에 속해 있는지 검토합니다.',
      categoryLabel: 'MEP',
      categoryTitle: 'BQC 웍셋 배정 검토 기능',
      requiresSharedParams: false
    },
    parameterduplication: {
      label: '프로젝트 파라미터 중복 검토',
      cardLabel: '프로젝트 파라미터 중복 검토',
      desc: '추가된 프로젝트 파라미터 중 이름이 중복된 파라미터를 검토',
      cardDesc: '지정한 파라미터 또는 추가된 전체 프로젝트 파라미터를 대상으로 이름 중복 여부를 검토합니다.',
      categoryLabel: '파라미터',
      categoryTitle: 'BQC 파라미터 중복 검토 기능',
      requiresSharedParams: false
    },
    parametermissing: {
      label: '파라미터 누락 검토',
      cardLabel: '파라미터 누락 검토',
      desc: '조건별 객체 속성 추출과 같은 대상 기준으로 지정한 공유 텍스트 파라미터 누락 여부를 검토',
      cardDesc: '공유 텍스트 파라미터를 선택하고, 공통 검토대상 필터 기준에 맞춰 누락 예외 규칙(AND/OR)을 설정해 누락 여부를 검토합니다.',
      categoryLabel: '파라미터',
      categoryTitle: 'BQC 파라미터 누락 검토 기능',
      requiresSharedParams: true
    },
    guid: {
        label: '파라미터 GUID 검토 및 정리',
        cardLabel: '파라미터 GUID 검토 및 정리',
        desc: '프로젝트/패밀리 파라미터 GUID를 검토하고 삭제용 엑셀 기준으로 정리',
        cardDesc: '프로젝트/패밀리 파라미터 GUID를 검토하고 삭제용 엑셀 기준으로 정리합니다.',
        categoryLabel: '파라미터',
        categoryTitle: '파라미터 검토 기능',
        requiresSharedParams: true
    },
  familylink: {
    label: '패밀리 공유파라미터 연동 검토',
    cardLabel: '패밀리 공유파라미터 연동 검토',
    desc: '복합패밀리 대상으로 하위 패밀리와의 지정한 파라미터 연동 여부 검토',
    cardDesc: '복합패밀리 대상으로 하위 패밀리와의 지정한 파라미터 연동 여부 검토',
    categoryLabel: '패밀리',
    categoryTitle: '패밀리 연동/관리 기능',
    requiresSharedParams: true
  },
  points: {
    label: '기준점/북각 추출',
    cardLabel: '기준점/북각 추출',
    desc: 'RVT의 프로젝트 기준점, 측량 기준점, 북각 값을 추출',
    cardDesc: 'RVT의 프로젝트 기준점, 측량 기준점, 북각 값을 추출',
    categoryLabel: '데이터',
    categoryTitle: '데이터 추출/좌표 확인 기능',
    requiresSharedParams: false
  },
  linkworkset: {
    label: '링크 기본 웍셋 점검/적용',
    cardLabel: '링크 기본 웍셋 점검/적용',
    desc: '각 RVT의 Revit 링크를 점검하고 기본 웍셋(Workset1)만 열리도록 적용',
    cardDesc: '링크별 로드 상태와 열려 있는 웍셋 현황을 확인하고, 필요 시 기본 웍셋(Workset1)만 열리도록 재적용합니다.',
    categoryLabel: '링크',
    categoryTitle: '링크 로드/웍셋 관리 기능',
    requiresSharedParams: false
  }
};
const FEATURE_KEYS = Object.keys(FEATURE_META);
const BQC_FEATURE_KEYS = ['connector', 'floorinfo', 'familysuitability', 'tapalign', 'dupclash', 'worksetassignment', 'parameterduplication', 'parametermissing'];
const COMMON_SCOPE_DEPENDENT_FEATURE_KEYS = ['connector', 'floorinfo', 'tapalign', 'dupclash', 'worksetassignment', 'parametermissing'];
const COMMON_EXTRA_PARAM_DEPENDENT_FEATURE_KEYS = ['connector', 'floorinfo', 'tapalign', 'dupclash', 'worksetassignment', 'parametermissing'];
const COMMON_TAPALIGN_OPTION_DEPENDENT_FEATURE_KEYS = ['tapalign'];
const UTILITY_FEATURE_KEYS = FEATURE_KEYS.filter((key) => !BQC_FEATURE_KEYS.includes(key));
const COMMON_OPTIONS_KEY = 'kky.hub.commonOptions';
const GROUP_FILTER_KEY = 'kky.hub.multiGroupFilter';
const MULTI_MODE_KEY = 'kky.hub.multiMode';
const TAPALIGN_STORAGE_KEY = 'kky_tapalign_opts';
const FAMILY_SUITABILITY_STORAGE_KEY = 'kky.hub.familySuitabilityRecent';
const FAMILY_SUITABILITY_PRESET_KEY = 'kky.hub.familySuitabilityPresets';
const PARAMETER_DUPLICATION_PRESET_KEY = 'kky.hub.parameterDuplicationPresets';
const PARAMETER_DUPLICATION_RECENT_KEY = 'kky.hub.parameterDuplicationRecent';
const PARAMETER_DUPLICATION_RECENT_LIMIT = 8;
const PARAMETER_MISSING_STORAGE_KEY = 'kky.hub.parameterMissingConfig';
const PARAMETER_MISSING_RECENT_KEY = 'kky.hub.parameterMissingRecent';
const PARAMETER_MISSING_RECENT_LIMIT = 8;
const PARAMETER_MISSING_PRESET_KIND = 'kky.hub.parameterMissingPreset';
const PARAMETER_MISSING_PRESET_VERSION = 1;
const PARAMETER_MISSING_PRESET_EXTENSION = '.kkypm.json';
const EXCEL_PHASE_WEIGHT = { EXCEL_INIT: 0.05, EXCEL_WRITE: 0.85, EXCEL_SAVE: 0.08, AUTOFIT: 0.02, DONE: 1, ERROR: 1 };
const EXCEL_PHASE_ORDER = ['EXCEL_INIT', 'EXCEL_WRITE', 'EXCEL_SAVE', 'AUTOFIT', 'DONE'];
const PARAMETER_MISSING_FILTER_OPERATORS = [
  'Equals', 'NotEquals', 'Contains', 'NotContains', 'BeginsWith', 'NotBeginsWith',
  'EndsWith', 'NotEndsWith', 'Greater', 'GreaterOrEqual', 'Less', 'LessOrEqual',
  'HasValue', 'HasNoValue'
];
const PARAMETER_MISSING_VALUELESS_OPERATORS = new Set(['HasValue', 'HasNoValue']);
const PARAMETER_MISSING_TARGET_FILTER_MODES = [
  { value: 'include', label: '조건에 해당하는 객체만 검토' },
  { value: 'exclude', label: '조건에 해당하지 않는 객체만 검토' }
];
const FAVORITE_PRESET_KIND = 'kky.hub.favoritesPreset';
const FAVORITE_PRESET_VERSION = 2;
const FAVORITE_PRESET_EXTENSION = '.kkyfav.json';
const GROUPS = [
  { id: 'all', label: '전체' },
  { id: 'bqc', label: '납품 시 BQC 검토' },
  { id: 'utility', label: '유틸리티' }
];

export function renderMulti(root, options = {}) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);
  const top = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar');
  if (top) top.classList.add('hub-topbar');

  const multiMode = options.viewMode === 'favorites'
    ? 'favorites'
    : normalizeMultiMode(getMultiMode());
  const favoriteEntries = multiMode === 'favorites' ? getFavoriteEntries() : [];
  const persistedPanelSearch = multiMode === 'favorites'
    ? { query: '', entryId: '' }
    : getHubPanelSearch(multiMode);
  const tapAlignPersisted = loadTapAlignConfigFromStorage();
  const familySuitabilityPersisted = loadFamilySuitabilityConfigFromStorage();
  const parameterMissingPersisted = loadParameterMissingConfigFromStorage();

  const state = {
    rvtList: [],
    rvtChecked: new Set(),
    busy: false,
    common: createConfigState({
      extraParams: '',
      targetFilter: '',
      excludeTargetFilter: '',
      excludeEndDummy: false,
      includePointXY: false,
      includeLinearMetrics: false
    }),
    features: {
      connector: createFeatureState({
        tol: 1.0,
        unit: 'inch',
        param: 'Comments',
        paramItems: ['Comments'],
        excludeEndDummy: false,
        includePointXY: false,
        includeLinearMetrics: false
      }),
      floorinfo: createFeatureState({ parameterName: '', levelRules: [], documentTitle: '', warnings: [] }),
      familysuitability: createFeatureState(familySuitabilityPersisted.config),
      tapalign: createFeatureState(tapAlignPersisted.config),
      dupclash: createFeatureState({ mode: 'duplicate', tolFeet: 1 / 64 }),
      worksetassignment: createFeatureState({ expectedWorksetName: 'Workset1', flaggedWorksetName: '' }),
      parameterduplication: createFeatureState({
        scope: 'all',
        parameterNames: [],
        sharedParamSourcePath: '',
        sharedParamImportCount: 0
      }),
      parametermissing: createFeatureState(parameterMissingPersisted.config),
      guid: createFeatureState({
        includeFamily: false,
        includeAnnotation: false,
        closeAllWorksetsOnOpen: true,
        useSyncComment: false,
        syncComment: 'KKY Tools - 파라미터 GUID 정리'
      }),
      familylink: createFeatureState({ targetsText: '', selectedTargets: [], targets: [] }),
      points: createFeatureState({ unit: 'ft' }),
      linkworkset: createFeatureState({
        applyDefaultWorksetOnly: true,
        useSyncComment: false,
        syncComment: 'KKY Tools - 링크 기본 웍셋 적용'
      })
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
      lastExcelPct: 0,
      excelBatchStartPct: null,
      excelBatchEndPct: null,
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
      reviewSummaryByMode: { bqc: null, utility: null, favorites: null },
      lastRunAtByMode: { bqc: null, utility: null, favorites: null },
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
      lastRunAt: null,
      featureInfoPanels: {},
      searchQuery: String(persistedPanelSearch?.query || ''),
      searchEntryId: String(persistedPanelSearch?.entryId || ''),
      searchInput: null,
      searchHint: null,
      searchClearBtn: null,
      searchSection: null,
      searchMode: multiMode === 'favorites' ? '' : multiMode,
      searchEmptyEl: null,
      favoriteSectionEl: null,
      favoriteSourceSections: [],
      favoritePresetInfoEl: null,
      favoritePresetName: '',
      favoritePresetPath: '',
      favoritePresetSource: ''
    }
  };

  FEATURE_KEYS.forEach((k) => {
    state.results[k] = { count: 0, stale: true, hasRun: false };
  });

  if (tapAlignPersisted.hasStored) {
    state.features.tapalign.applied = true;
  }
  if (familySuitabilityPersisted.hasStored) {
    state.features.familysuitability.applied = true;
  }
  if (parameterMissingPersisted.hasStored) {
    state.features.parametermissing.applied = true;
  }

  const page = div('feature-shell multi-page HubShell');
  const hasLocalCommonOptions = loadCommonOptionsFromStorage();
  const header = div('feature-header multi-header');
  header.innerHTML = buildHeaderHtml(state.ui.multiMode);
  const headerHeading = header.querySelector('.feature-heading');
  const headerGuide = buildFeatureTypeGuide('header');
  if (headerHeading) {
    headerHeading.append(headerGuide);
  } else {
    header.append(headerGuide);
  }
  if (state.ui.multiMode !== 'favorites') {
    header.append(buildFeatureSearchPanel(state.ui.multiMode));
  }
  page.append(header);

  const layoutClass = state.ui.multiMode === 'favorites'
    ? 'multi-layout--favorites'
    : state.ui.multiMode === 'utility'
      ? 'multi-layout--utility'
      : 'multi-layout--bqc';
  const layout = div(`multi-layout HubBody ${layoutClass}`);
  const mainCol = div('multi-workspace multi-workspace--main');
  const sideCol = div('multi-workspace multi-workspace--sidebar');

  const group1 = buildGroupSection('납품 시 BQC 검토', '가장 많이 사용하는 핵심 검토를 먼저 선택해 주세요.', 'bqc');
  const group3 = buildGroupSection('유틸리티', '활성 문서 검토 및 보조 워크플로우를 실행합니다.', 'utility');
  group3.section.id = 'utilities';

  buildGroup1Options();
  group1.section.append(buildDeliveryCleanerWorkflowRow({
    chipLabel: '파일 정리',
    chipTitle: '파일 정리 워크플로우 화면으로 이동합니다.'
  }));
  group1.section.append(buildConditionExtractWorkflowRow());
  group1.section.append(buildToggleRow('connector', buildConnectorConfig()));
  group1.section.append(buildToggleRow('floorinfo', buildFloorInfoConfig()));
  group1.section.append(buildToggleRow('familysuitability', buildFamilySuitabilityConfig()));
  group1.section.append(buildToggleRow('tapalign', buildTapAlignConfig()));
  group1.section.append(buildToggleRow('dupclash', buildDupClashConfig()));
  group1.section.append(buildToggleRow('worksetassignment', buildWorksetAssignmentConfig()));
  group1.section.append(buildToggleRow('parameterduplication', buildParameterDuplicationConfig()));
  group1.section.append(buildToggleRow('parametermissing', buildParameterMissingConfig()));
  group1.section.append(buildDupWorkflowRow({
    chipLabel: 'BQC 보조',
    chipTitle: 'BQC 보조 검토 항목에서 중복 / 자체 간섭 검토를 바로 엽니다.'
  }));
  organizeFeatureRows(group1.section);
  group3.section.append(buildDupWorkflowRow());
  group3.section.append(buildParamPropWorkflowRow());
  group3.section.append(buildPmsWorkflowRow());
  group3.section.append(buildParamModifierWorkflowRow());
  group3.section.append(buildLinkPathWorkflowRow());
  group3.section.append(buildLateralNozzleWorkflowRow());
  group3.section.append(buildGuidWorkflowRow());
  group3.section.append(buildToggleRow('familylink', buildFamilyLinkConfig()));
  group3.section.append(buildToggleRow('points', buildPointsConfig()));
  group3.section.append(buildToggleRow('linkworkset', buildLinkWorksetConfig()));
  group3.section.append(buildSharedParamBatchRow());
  organizeFeatureRows(group3.section);
  if (state.ui.multiMode === 'bqc' && group1.section.querySelectorAll('.feature-row').length === 0) {
    const empty = div('feature-note');
    empty.textContent = '등록된 BQC 검토 기능이 없습니다.';
    group1.section.append(empty);
  }

  if (state.ui.multiMode === 'favorites') {
    const favoriteSection = buildGroupSection('즐겨찾기', '우클릭으로 추가한 기능만 따로 모아 실행합니다.', 'favorites');
    const favoriteSourceSections = [group3.section, group1.section];
    const favoriteRows = collectFavoriteRows(favoriteSourceSections, favoriteEntries);
    favoriteSection.section.id = 'favorites';
    state.ui.favoriteSectionEl = favoriteSection.section;
    state.ui.favoriteSourceSections = favoriteSourceSections;
    favoriteRows.forEach((row) => favoriteSection.section.append(row));
    organizeFeatureRows(favoriteSection.section);
    syncFavoritesSectionRows(favoriteSection.section, favoriteSourceSections);
    state.ui.groupFilter = 'favorites';
    primeFeatureInfoPanel('favorites', favoriteSection.section);
    favoriteSection.wrap.classList.add('multi-section--utility', 'multi-section--favorites');
    const favoriteHead = favoriteSection.wrap.querySelector('.multi-section-title');
    if (favoriteHead) {
      favoriteHead.remove();
      favoriteSection.wrap.classList.add('multi-section--headerless');
    }
    const favoriteSidebar = div('multi-sidebar-stack multi-sidebar-stack--sticky');
    favoriteSidebar.append(
      buildExecutionActionPanel({ mode: 'favorites' }),
      buildSelectedFeaturesSection({
        title: '선택된 기능',
        showCurrentButton: false,
        sectionClass: 'selected-panel--sidebar'
      })
    );
    mainCol.append(favoriteSection.wrap);
    sideCol.append(favoriteSidebar);
    layout.append(mainCol, sideCol);
    const onFavoritesChange = () => {
      if (!favoriteSection.wrap.isConnected) {
        window.removeEventListener(HUB_QUICK_ACCESS_CHANGE_EVENT, onFavoritesChange);
        return;
      }
      syncFavoritesSectionRows(favoriteSection.section, favoriteSourceSections);
    };
    window.addEventListener(HUB_QUICK_ACCESS_CHANGE_EVENT, onFavoritesChange);
  } else if (state.ui.multiMode === 'bqc') {
    state.ui.groupFilter = state.ui.multiMode;
    saveGroupFilter(state.ui.multiMode);
    state.ui.searchMode = 'bqc';
    state.ui.searchSection = group1.section;
    primeFeatureInfoPanel('bqc', group1.section);
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
    state.ui.groupFilter = state.ui.multiMode;
    saveGroupFilter(state.ui.multiMode);
    state.ui.searchMode = 'utility';
    state.ui.searchSection = group3.section;
    primeFeatureInfoPanel('utility', group3.section);
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
  if (state.ui.multiMode !== 'favorites') {
    applyFeatureSearchFilter();
  }
  applyPendingFeatureFocus(page, state.ui.multiMode === 'favorites' ? 'favorites' : 'multi');

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
      refreshUiAfterHostDialog(() => renderRvtList());
    }
  });

  onHost('hub:multi-progress', (payload) => {
    const rawPhase = String(payload?.phase || payload?.Phase || '').trim();
    const isExcelProgress = payload?.isExcel === true || String(payload?.isExcel || '').trim().toLowerCase() === 'true';
    if (isExcelProgress && isExcelProgressPhase(rawPhase)) {
      const phase = normalizeExcelPhase(rawPhase);
      updateExcelBatchWindow(phase, payload);
      const total = Number(payload?.total) || 0;
      const current = Number(payload?.current) || 0;
      const percent = computeExcelPercent(phase, current, total, payload?.phaseProgress, payload?.percent);
      const subtitle = buildExcelSubtitle(phase, current, total);
      const detail = formatExcelDetail(phase, payload?.message, payload?.detail);
      ProgressDialog.show('엑셀 내보내기', subtitle);
      ProgressDialog.update(percent, subtitle, detail);
      return;
    }
    const basePct = Number(payload?.percent);
    const altPct = Number(payload?.phaseProgress);
    const hasBase = Number.isFinite(basePct);
    const hasAlt = Number.isFinite(altPct);
    const pctValue = hasBase ? basePct : (hasAlt ? altPct * 100 : state.ui.lastProgressPct);
    const pct = Math.max(0, Math.min(100, pctValue));
    state.ui.lastProgressPct = pct;
    const phase = String(payload?.phase || payload?.Phase || '').toLowerCase();
    ProgressDialog.show(payload?.title || '납품 시 BQC 검토', payload?.message || '');
    ProgressDialog.update(pct, payload?.message || '', payload?.detail || '');
    updateRunProgress(pct, payload?.message || '', payload?.detail || '');
    if (phase === 'done') {
      ProgressDialog.hide();
    }
  });

  onHost('hub:multi-done', (payload) => {
    setBusyState(false);
    resetExcelProgressState();
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
    if (buildParameterDuplicationConfig.renderSharedParamList) buildParameterDuplicationConfig.renderSharedParamList(payload);
    if (buildParameterMissingConfig.renderSharedParamList) buildParameterMissingConfig.renderSharedParamList(payload);
  });

  onHost('connector:param-list:done', (payload) => {
    const ok = payload?.ok !== false;
    state.connectorParamItems = ok ? normalizeConnectorParamItems(payload) : [];
    if (buildConnectorConfig.renderList) buildConnectorConfig.renderList(payload);
  });

  onHost('floorinfo:config-loaded', (payload) => {
    if (buildFloorInfoConfig.applySnapshot) buildFloorInfoConfig.applySnapshot(payload);
  });
  onHost('familysuitability:criteria-picked', (payload) => {
    if (buildFamilySuitabilityConfig.applyCriteriaPicked) {
      refreshUiAfterHostDialog(() => buildFamilySuitabilityConfig.applyCriteriaPicked(payload));
    }
  });
  onHost('hub:multi-error', (payload) => {
    setBusyState(false);
    resetExcelProgressState();
    ProgressDialog.hide();
    updateRunProgress(0, '오류 발생', payload?.message || '');
    toast(payload?.message || '다중 검토 실행 중 오류가 발생했습니다. 선택한 기능, 대상 RVT, 공통 설정과 필터를 확인한 뒤 다시 실행해 주세요. 계속 실패하면 메시지를 관리자에게 전달해 주세요.', 'err');
    state.ui.runCompleted = false;
    updateRunActionLabel();
  });

  onHost('hub:multi-canceled', (payload) => {
    setBusyState(false);
    resetExcelProgressState();
    ProgressDialog.hide();
    updateRunProgress(0, '실행 취소', payload?.message || '');
    toast(payload?.message || '작업을 취소했습니다.', 'warn');
    state.ui.runCompleted = false;
    updateRunActionLabel();
  });

  onHost('favorites:preset-saved', (payload) => {
    if (!page.isConnected || state.ui.multiMode !== 'favorites') return;
    handleFavoritePresetSaved(payload);
  });

  onHost('favorites:preset-loaded', (payload) => {
    if (!page.isConnected || state.ui.multiMode !== 'favorites') return;
    handleFavoritePresetLoaded(payload);
  });

  onHost('favorites:preset-error', (payload) => {
    if (!page.isConnected || state.ui.multiMode !== 'favorites') return;
    toast(payload?.message || '즐겨찾기 프리셋을 처리하지 못했습니다. 프리셋 파일 형식과 현재 즐겨찾기 목록을 확인한 뒤 다시 실행해 주세요. 계속 실패하면 메시지를 관리자에게 전달해 주세요.', 'err');
  });

  onHost('hub:multi-exported', (payload) => {
    setBusyState(false);
    ProgressDialog.hide();
    state.ui.lastProgressPct = 0;
    resetExcelProgressState();
    const path = payload?.path;
    if (path) {
      const isFolder = payload?.kind === 'folder';
      requestAnimationFrame(() => {
        showExcelSavedDialog(
          payload?.message || '다중 검토 결과 엑셀 저장 완료',
          path,
          (p) => post('excel:open', { path: p }),
          isFolder
            ? { description: '저장한 폴더를 바로 여시겠습니까?', openLabel: '폴더 열기' }
            : undefined
        );
      });
    } else if (payload?.cancelled) {
      toast(payload?.message || '엑셀 저장이 취소되었습니다.', 'warn');
    } else {
      toast(payload?.message || '다중 검토 결과 엑셀 저장에 실패했습니다. 저장 경로 권한과 파일이 열려 있는지 확인해 주세요. 계속 실패하면 메시지를 관리자에게 전달해 주세요.', 'err');
    }
  });

  onHost('sharedparam:status', (payload) => {
    state.sharedParamStatus = payload || {};
    updateSharedParamBanner();
    updateRunSummary();
    if (buildParameterDuplicationConfig.renderSharedParamStatus) buildParameterDuplicationConfig.renderSharedParamStatus(payload);
    if (buildParameterMissingConfig.renderSharedParamStatus) buildParameterMissingConfig.renderSharedParamStatus(payload);
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

  function mergeRvtPaths(paths) {
    const list = Array.isArray(paths) ? paths : [];
    const existing = new Set(state.rvtList.map((path) => String(path || '').toLowerCase()));
    let added = 0;
    list.forEach((path) => {
      if (!path) return;
      const key = path.toLowerCase();
      if (!existing.has(key)) {
        state.rvtList.push(path);
        existing.add(key);
        added += 1;
      }
      state.rvtChecked.add(path);
    });
    return { added };
  }

  function handleDroppedRvts(droppedPaths) {
    const { added } = mergeRvtPaths(droppedPaths);
    if (!added) {
      toast('이미 등록된 RVT입니다.', 'warn');
      return;
    }
    markAllStale();
    renderRvtList();
    toast(`${added}개 RVT를 추가했습니다.`, 'ok');
  }

  function handleRemoveSelected() {
    if (state.rvtChecked.size === 0) return;
    state.rvtList = state.rvtList.filter((p) => !state.rvtChecked.has(p));
    state.rvtChecked.clear();
    markAllStale();
    renderRvtList();
  }

  function handleClearList() {
    if (state.rvtList.length === 0) return;
    const confirmed = window.confirm('RVT 목록을 모두 삭제할까요?');
    if (!confirmed) return;
    state.rvtList = [];
    state.rvtChecked.clear();
    markAllStale();
    renderRvtList();
  }

  function requestSharedParamList(context) {
    post('sharedparam:list', { source: 'multi', context: context || '' });
  }

  function requestConnectorParamList(context) {
    post('connector:param-list', { source: 'multi', context: context || '' });
  }

  function normalizeMultiMode(value) {
    if (value === 'favorites') return 'favorites';
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
    if (mode === 'favorites') {
      return `
    <div class="feature-heading">
      <span class="feature-kicker">즐겨찾기</span>
      <h2 class="feature-title">자주 사용하는 기능</h2>
      <p class="feature-sub">즐겨찾기로 모은 기능만 따로 열어 선택형은 함께 검토하고, 별도 프로세스는 바로 실행합니다.</p>
    </div>`;
    }
    if (mode === 'utility') {
      return `
    <div class="feature-heading">
      <span class="feature-kicker">유틸리티</span>
      <h2 class="feature-title">유틸리티</h2>
      <p class="feature-sub">보조 검토와 일괄 작업 기능을 실행합니다.</p>
    </div>`;
    }
      return `
    <div class="feature-heading">
      <span class="feature-kicker">BQC 검토</span>
      <h2 class="feature-title">납품 시 BQC 검토</h2>
      <p class="feature-sub">납품 검토에 필요한 기능을 선택해 실행합니다.</p>
    </div>`;
  }

  function buildFeatureSearchPanel(mode) {
    const wrap = div('multi-search-bar');
    const top = div('multi-search-bar__top');
    const title = document.createElement('strong');
    title.textContent = '기능 검색';
    const hint = document.createElement('span');
    hint.className = 'multi-search-bar__hint';
    top.append(title, hint);

    const field = div('multi-search-bar__field');
    const icon = document.createElement('span');
    icon.className = 'multi-search-bar__icon';
    icon.textContent = '⌕';
    const input = document.createElement('input');
    input.type = 'search';
    input.className = 'multi-search-bar__input';
    input.placeholder = mode === 'utility'
      ? '유틸리티 기능 검색'
      : 'BQC 기능 검색';
    input.value = state.ui.searchQuery || '';
    input.autocomplete = 'off';
    input.spellcheck = false;
    const clearBtn = document.createElement('button');
    clearBtn.type = 'button';
    clearBtn.className = 'btn btn--secondary multi-search-bar__clear';
    clearBtn.textContent = '지우기';

    field.append(icon, input, clearBtn);
    wrap.append(top, field);

    state.ui.searchInput = input;
    state.ui.searchHint = hint;
    state.ui.searchClearBtn = clearBtn;

    const commit = (nextQuery, nextEntryId = '') => {
      state.ui.searchQuery = String(nextQuery || '').trim();
      state.ui.searchEntryId = state.ui.searchQuery ? String(nextEntryId || '').trim() : '';
      setHubPanelSearch(mode, {
        query: state.ui.searchQuery,
        entryId: state.ui.searchEntryId
      });
      syncFeatureSearchUi();
      applyFeatureSearchFilter();
    };

    input.addEventListener('input', () => {
      commit(input.value, '');
    });
    input.addEventListener('search', () => {
      commit(input.value, '');
    });
    clearBtn.addEventListener('click', () => {
      input.value = '';
      commit('', '');
      input.focus();
    });

    syncFeatureSearchUi();
    return wrap;
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
    const body = div('multi-feature-list');
    if (groupId) body.dataset.group = groupId;
    const info = buildFeatureInfoPanel(groupId, title, desc);
    body.addEventListener('mouseover', (ev) => {
      const row = ev.target.closest('.feature-row');
      if (!row || !body.contains(row)) return;
      updateFeatureInfoPanel(groupId, row);
    });
    body.addEventListener('focusin', (ev) => {
      const row = ev.target.closest('.feature-row');
      if (!row || !body.contains(row)) return;
      updateFeatureInfoPanel(groupId, row);
    });
    body.addEventListener('mouseleave', () => {
      updateFeatureInfoPanel(groupId, null);
    });
    body.addEventListener('focusout', () => {
      window.setTimeout(() => {
        if (!body.contains(document.activeElement)) updateFeatureInfoPanel(groupId, null);
      }, 0);
    });
    wrap.append(head, body, info);
    return { wrap, section: body };
  }

  function buildFeatureTypeGuide(variant = 'default') {
    const guide = div(`multi-feature-guide multi-feature-guide--${variant}`.trim());
    const selectable = document.createElement('span');
    selectable.className = 'multi-feature-guide__chip multi-feature-guide__chip--selectable';
    selectable.textContent = '선택형 카드';
    const workflow = document.createElement('span');
    workflow.className = 'multi-feature-guide__chip multi-feature-guide__chip--workflow';
    workflow.textContent = '별도 워크플로우';
    const text = document.createElement('span');
    text.className = 'multi-feature-guide__text';
    text.textContent = variant === 'header'
      ? '선택형 카드는 함께 검토하고, 별도 워크플로우는 개별 화면에서 실행합니다.'
      : '선택형 카드는 함께 검토하고, 별도 워크플로우 카드는 개별 화면에서 실행합니다.';
    guide.append(selectable, workflow, text);
    return guide;
  }

  function buildFeatureGroupTitle(title, desc, variant) {
    const card = div(`multi-feature-group-title multi-feature-group-title--${variant || 'default'}`.trim());
    const heading = document.createElement('strong');
    heading.textContent = title || '';
    const body = document.createElement('span');
    body.textContent = desc || '';
    card.append(heading, body);
    return card;
  }

  function buildFeatureGroupBlock(title, desc, variant, cards) {
    const group = div(`multi-feature-group multi-feature-group--${variant || 'default'}`.trim());
    const grid = div('multi-feature-group-grid');
    group.append(buildFeatureGroupTitle(title, desc, variant), grid);
    cards.forEach((card) => grid.append(card));
    return group;
  }

  function organizeFeatureRows(section) {
    if (!section) return;
    const cards = Array.from(section.children).filter((el) => el.classList?.contains('feature-row'));
    if (!cards.length) return;
    const selectable = cards.filter((el) => !el.classList.contains('feature-row--launch'));
    const workflow = cards.filter((el) => el.classList.contains('feature-row--launch'));
    const others = Array.from(section.children).filter((el) => !el.classList?.contains('feature-row'));
    section.replaceChildren();
    if (selectable.length) {
      section.append(buildFeatureGroupBlock('선택형 카드', '눌러서 켠 기능을 한 번에 함께 검토합니다.', 'selectable', selectable));
    }
    if (workflow.length) {
      section.append(buildFeatureGroupBlock('별도 워크플로우', '전용 화면으로 이동해 개별 워크플로우를 실행합니다.', 'workflow', workflow));
    }
    others.forEach((node) => section.append(node));
  }

  function collectFavoriteRows(sections, entries) {
    const sourceSections = Array.isArray(sections) ? sections : [];
    const favoriteList = Array.isArray(entries) ? entries : [];
    const seen = new Set();
    const rows = [];

    favoriteList.forEach((entry) => {
      const entryId = String(entry?.id || '').trim();
      if (!entryId || seen.has(entryId)) return;
      for (const section of sourceSections) {
        const row = section?.querySelector?.(`[data-entry-id="${entryId}"]`);
        if (!row) continue;
        seen.add(entryId);
        rows.push(row);
        break;
      }
    });

    return rows;
  }

  function clearFavoriteFeatureSelection(entryId) {
    const key = String(entryId || '').trim();
    if (!FEATURE_KEYS.includes(key) || !state.features[key]) return false;
    const feature = state.features[key];
    if (!feature.enabled && !feature.applied && !feature.dirty) return false;
    feature.enabled = false;
    feature.applied = false;
    feature.dirty = false;
    resetDraftFromCommitted(key);
    markStale(key);
    syncFeatureRow(key);
    return true;
  }

  function syncFavoritesSectionRows(section, sourceSections = state.ui.favoriteSourceSections || []) {
    if (!section) return;

    const favoriteEntries = getFavoriteEntries();
    const favoriteIds = new Set(
      favoriteEntries
        .map((entry) => String(entry?.id || '').trim())
        .filter(Boolean)
    );

    let changedSelection = false;
    Array.from(section.querySelectorAll('.feature-row[data-entry-id]')).forEach((row) => {
      const entryId = String(row.dataset.entryId || '').trim();
      if (!entryId || favoriteIds.has(entryId)) return;
      changedSelection = clearFavoriteFeatureSelection(entryId) || changedSelection;
    });

    const rows = collectFavoriteRows([section, ...(Array.isArray(sourceSections) ? sourceSections : [])], favoriteEntries);
    section.replaceChildren();

    if (rows.length) {
      rows.forEach((row) => section.append(row));
      organizeFeatureRows(section);
    } else {
      const empty = div('feature-note');
      empty.dataset.favoritesEmpty = 'true';
      empty.textContent = '아직 즐겨찾기한 기능이 없습니다. 각 기능 카드에서 오른쪽 클릭으로 즐겨찾기를 추가해 주세요.';
      section.append(empty);
    }

    if (changedSelection) {
      updateRunSummary();
      updateRunActionLabel();
    }
    updateFeatureInfoPanel('favorites', section.querySelector('.feature-row[data-entry-id]'));
  }

  function buildFeatureInfoPanel(groupId, title, desc) {
    const panel = div('multi-feature-info');
    const kicker = document.createElement('span');
    kicker.className = 'multi-feature-info__kicker';
    kicker.textContent = groupId === 'bqc' ? 'BQC 안내' : '기능 안내';
    const heading = document.createElement('strong');
    heading.className = 'multi-feature-info__title';
    const body = document.createElement('span');
    body.className = 'multi-feature-info__body';
    panel.append(kicker, heading, body);
    state.ui.featureInfoPanels[groupId || 'default'] = {
      root: panel,
      heading,
      body,
      defaultTitle: title || '기능 안내',
      defaultDesc: desc || '카드에 마우스를 올리면 설명이 표시됩니다.'
    };
    updateFeatureInfoPanel(groupId, null);
    return panel;
  }

  function updateFeatureInfoPanel(groupId, row) {
    const panel = state.ui.featureInfoPanels[groupId || 'default'];
    if (!panel) return;
    const title = row?.dataset?.infoTitle || panel.defaultTitle;
    const desc = row?.dataset?.infoDesc || panel.defaultDesc;
    panel.heading.textContent = title;
    panel.body.textContent = desc;
  }

  function primeFeatureInfoPanel(groupId, section) {
    if (!section) return;
    updateFeatureInfoPanel(groupId, section.querySelector('.feature-row'));
  }

  function syncFeatureSearchUi() {
    const query = String(state.ui.searchQuery || '').trim();
    const entryId = String(state.ui.searchEntryId || '').trim();
    if (state.ui.searchInput && state.ui.searchInput.value !== query) {
      state.ui.searchInput.value = query;
    }
    if (state.ui.searchHint) {
      if (!query) {
        state.ui.searchHint.textContent = '검색어를 지우면 전체 기능이 다시 보입니다.';
      } else if (entryId) {
        state.ui.searchHint.textContent = `"${query}" 검색 결과에서 선택한 기능만 표시 중입니다.`;
      } else {
        state.ui.searchHint.textContent = `"${query}" 키워드와 일치하는 기능만 표시 중입니다.`;
      }
    }
    if (state.ui.searchClearBtn) {
      state.ui.searchClearBtn.disabled = !query;
    }
  }

  function applyFeatureSearchFilter() {
    const section = state.ui.searchSection;
    const mode = state.ui.searchMode || state.ui.multiMode || 'bqc';
    if (!section || !mode) return;

    const query = String(state.ui.searchQuery || '').trim();
    const entryId = query ? String(state.ui.searchEntryId || '').trim() : '';
    const matchedIds = query
      ? new Set(
        searchHubEntries(query)
          .map((entry) => String(entry?.id || '').trim())
          .filter(Boolean)
      )
      : null;

    let visibleCount = 0;
    Array.from(section.querySelectorAll('.feature-row[data-entry-id]')).forEach((row) => {
      const currentEntryId = String(row.dataset.entryId || '').trim();
      const show = !query
        ? true
        : entryId
          ? currentEntryId === entryId
          : matchedIds.has(currentEntryId);
      row.hidden = !show;
      if (show) visibleCount += 1;
    });

    Array.from(section.querySelectorAll('.multi-feature-group')).forEach((group) => {
      const hasVisibleRows = !!group.querySelector('.feature-row[data-entry-id]:not([hidden])');
      group.hidden = !hasVisibleRows;
    });

    if (!state.ui.searchEmptyEl) {
      const empty = div('multi-search-empty feature-note');
      empty.hidden = true;
      state.ui.searchEmptyEl = empty;
      section.append(empty);
    }
    if (state.ui.searchEmptyEl) {
      state.ui.searchEmptyEl.textContent = query
        ? `"${query}" 검색 결과가 없습니다. 검색어를 지우면 전체 기능이 다시 보입니다.`
        : '';
      state.ui.searchEmptyEl.hidden = !query || visibleCount > 0;
    }

    updateFeatureInfoPanel(mode, section.querySelector('.feature-row[data-entry-id]:not([hidden])'));
  }

  function buildFeatureTooltip(parts = []) {
    return parts
      .map((part) => String(part || '').trim())
      .filter(Boolean)
      .join(' · ');
  }

  function applyFeatureRowTooltip(row, parts = [], info = {}) {
    if (!row) return;
    const tooltip = buildFeatureTooltip(parts);
    row.classList.add('feature-row--tooltip');
    row.dataset.tooltip = tooltip;
    row.dataset.infoTitle = String(info.title || '').trim();
    row.dataset.infoDesc = String(info.desc || '').trim();
    row.removeAttribute('title');
    if (tooltip) {
      row.setAttribute('aria-description', tooltip);
    } else {
      row.removeAttribute('aria-description');
    }
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
    title.textContent = '아래 기능 영역을 눌러 실행할 검토를 선택해 주세요.';
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
    const extra = makeField('추가 파라미터 값 추출', 'extra', 'PM1, PM2', 'textarea');
    const filter = makeField('검토 대상 필터', 'filter', '예: PM1=값; PM2=값2', 'text');
    const excludeFilter = makeField('검토 제외 대상 필터', 'exclude_filter', '예: Family=Dummy; Comments=SKIP', 'text');

    const draft = state.common.configDraft;
    extra.input.value = draft.extraParams;
    filter.input.value = draft.targetFilter;
    excludeFilter.input.value = draft.excludeTargetFilter;

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
    excludeFilter.input.addEventListener('change', () => {
      state.common.configDraft.excludeTargetFilter = excludeFilter.input.value;
      markCommonDirty();
      updateCommonSummary(summary);
    });

    fields.append(extra.field, filter.field, excludeFilter.field);
    fields.append(buildFilterExamples());
    fields.classList.add('settings-panel', 'is-open');
    state.ui.panels.common = fields;
    state.ui.controls.common = { extra, filter, excludeFilter };

    return panel;
  }

  function buildToggleRow(key, config) {
    config = config || { panel: div('settings-panel'), controls: {} };
    const meta = FEATURE_META[key] || {};
    const cardLabel = meta.cardLabel || meta.label || key;
    const cardDesc = meta.cardDesc || meta.desc || '';
    const row = div('feature-row');
    row.classList.add('feature-row--selectable');
    row.dataset.key = key;
    row.dataset.entryId = key;
    const header = div('feature-row__header');
    const setFeatureEnabled = (enabled, shouldOpenSettings) => {
      const feature = state.features[key];
      const wasEnabled = !!feature.enabled;
      const nextEnabled = !!enabled;
      feature.enabled = nextEnabled;
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
      row.setAttribute('aria-pressed', nextEnabled ? 'true' : 'false');
      row.dataset.selected = nextEnabled ? 'true' : 'false';
      markStale(key);
      updateRunSummary();
      refreshConnectorFeatureSummary();
      refreshFloorInfoFeatureSummary();
      refreshFamilySuitabilityFeatureSummary();
      refreshTapAlignFeatureSummary();
      refreshDupClashFeatureSummary();
      refreshWorksetAssignmentFeatureSummary();
      renderSelectedFeatures();
      if (nextEnabled && !wasEnabled) {
        recordHubEntryUse(key);
      }
    };

    const metaWrap = div('feature-row__left');
    const textWrap = div('feature-row__text');
    const metaTitle = document.createElement('strong');
    metaTitle.textContent = cardLabel;
    textWrap.append(metaTitle);
    metaWrap.append(textWrap);
    applyFeatureRowTooltip(row, [meta.label || key, meta.desc || '', cardDesc], {
      title: meta.label || key,
      desc: meta.desc || ''
    });

    const right = div('feature-row__right');
    if (key === 'connector') {
      const badge = document.createElement('span');
      badge.className = 'chip chip--info';
      badge.textContent = 'BQC 핵심';
      badge.setAttribute('aria-label', 'BQC 핵심 검토 기능');
      right.append(badge);
    } else if (key === 'floorinfo' || key === 'familysuitability' || key === 'tapalign' || key === 'dupclash' || key === 'worksetassignment' || key === 'parameterduplication' || key === 'parametermissing') {
      const badge = document.createElement('span');
      badge.className = 'chip chip--info';
      badge.textContent = 'BQC 보조';
      badge.setAttribute('aria-label', 'BQC 보조 검토 기능');
      right.append(badge);
    } else if (meta.categoryLabel) {
      const badge = document.createElement('span');
      badge.className = 'chip chip--info';
      badge.textContent = meta.categoryLabel;
      badge.setAttribute('aria-label', meta.categoryTitle || `${meta.categoryLabel} 기능`);
      right.append(badge);
    }

    const metaDesc = div('feature-row__desc');
    metaDesc.textContent = cardDesc;

    header.append(metaWrap, right);
    row.append(header, metaDesc);
    row.classList.add('is-clickable');
    row.setAttribute('role', 'button');
    row.setAttribute('aria-label', `${meta.label || key} 선택`);
    row.setAttribute('aria-pressed', state.features[key]?.enabled ? 'true' : 'false');
    row.tabIndex = 0;
    row.addEventListener('click', (ev) => {
      if (ev.target.closest('button, input, select, textarea, a, label')) return;
      setFeatureEnabled(!state.features[key].enabled, !state.features[key].enabled);
    });
    row.addEventListener('keydown', (ev) => {
      if (ev.target.closest('button, input, select, textarea, a, label')) return;
      if (ev.key !== 'Enter' && ev.key !== ' ') return;
      ev.preventDefault();
      setFeatureEnabled(!state.features[key].enabled, !state.features[key].enabled);
    });
    bindHubEntryContextMenu(row, key);

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
      top.textContent = '필요할 때만 켠 뒤 설정 창에서 공유 파라미터 검토 대상을 선택해 주세요.';
      sub.textContent = '아직 검토 파라미터가 지정되지 않았습니다.';
      summary.append(top, sub);
      row.append(summary);
      state.ui.connectorHeroSummary = { row, top, sub };
      refreshConnectorFeatureSummary();
    } else if (key === 'floorinfo') {
      const summary = div('feature-row__summary');
      summary.style.display = 'grid';
      summary.style.gap = '4px';
      summary.style.padding = '10px 12px';
      summary.style.borderRadius = '14px';
      summary.style.border = '1px dashed var(--border-soft)';
      summary.style.background = 'var(--surface-help)';
      const top = document.createElement('strong');
      const sub = document.createElement('span');
      top.textContent = '레벨 모델 기준 Z와 기대 층정보 값을 설정하면 모델 객체를 구간별로 검토합니다.';
      sub.textContent = '아직 검토 파라미터와 레벨 규칙이 지정되지 않았습니다.';
      summary.append(top, sub);
      row.append(summary);
      state.ui.floorInfoSummary = { row, top, sub };
      refreshFloorInfoFeatureSummary();
    } else if (key === 'familysuitability') {
      const summary = div('feature-row__summary');
      summary.style.display = 'grid';
      summary.style.gap = '4px';
      summary.style.padding = '10px 12px';
      summary.style.borderRadius = '14px';
      summary.style.border = '1px dashed var(--border-soft)';
      summary.style.background = 'var(--surface-help)';
      const top = document.createElement('strong');
      const sub = document.createElement('span');
      top.textContent = '기준 엑셀의 카테고리/패밀리/타입 조합으로 실제 사용 타입 적합성을 검토합니다.';
      sub.textContent = '아직 기준 엑셀과 검토 문구가 지정되지 않았습니다.';
      summary.append(top, sub);
      row.append(summary);
      state.ui.familySuitabilitySummary = { row, top, sub };
      refreshFamilySuitabilityFeatureSummary();
    } else if (key === 'tapalign') {
      const summary = div('feature-row__summary');
      summary.style.display = 'grid';
      summary.style.gap = '4px';
      summary.style.padding = '10px 12px';
      summary.style.borderRadius = '14px';
      summary.style.border = '1px dashed var(--border-soft)';
      summary.style.background = 'var(--surface-help)';
      const top = document.createElement('strong');
      const sub = document.createElement('span');
      top.textContent = '허용 범위와 공통 옵션을 기준으로 탭/분기 피팅 축 이탈 여부를 검토합니다.';
      sub.textContent = '아직 허용 범위, 검토 범위, 엑셀 언어가 적용되지 않았습니다.';
      summary.append(top, sub);
      row.append(summary);
      state.ui.tapAlignSummary = { row, top, sub };
      refreshTapAlignFeatureSummary();
    } else if (key === 'dupclash') {
      const summary = div('feature-row__summary');
      summary.style.display = 'grid';
      summary.style.gap = '4px';
      summary.style.padding = '10px 12px';
      summary.style.borderRadius = '14px';
      summary.style.border = '1px dashed var(--border-soft)';
      summary.style.background = 'var(--surface-help)';
      const top = document.createElement('strong');
      const sub = document.createElement('span');
      top.textContent = '완전 중복과 자체 간섭을 같은 배치에서 같이 검토합니다.';
      sub.textContent = '개별 필터 없이 공통 설정의 포함/제외 대상 필터만 그대로 따릅니다.';
      summary.append(top, sub);
      row.append(summary);
      state.ui.dupClashSummary = { row, top, sub };
      refreshDupClashFeatureSummary();
    } else if (key === 'worksetassignment') {
      const summary = div('feature-row__summary');
      summary.style.display = 'grid';
      summary.style.gap = '4px';
      summary.style.padding = '10px 12px';
      summary.style.borderRadius = '14px';
      summary.style.border = '1px dashed var(--border-soft)';
      summary.style.background = 'var(--surface-help)';
      const top = document.createElement('strong');
      const sub = document.createElement('span');
      top.textContent = '모델 객체가 기본 웍셋(Workset1)에 배정되었는지 빠르게 확인합니다.';
      sub.textContent = '기본 웍셋(Workset1) 이외의 웍셋에 속한 객체만 오류 행으로 내보냅니다.';
      summary.append(top, sub);
      row.append(summary);
      state.ui.worksetAssignmentSummary = { row, top, sub };
      refreshWorksetAssignmentFeatureSummary();
    } else if (key === 'parameterduplication') {
      const summary = div('feature-row__summary');
      summary.style.display = 'grid';
      summary.style.gap = '4px';
      summary.style.padding = '10px 12px';
      summary.style.borderRadius = '14px';
      summary.style.border = '1px dashed var(--border-soft)';
      summary.style.background = 'var(--surface-help)';
      const top = document.createElement('strong');
      const sub = document.createElement('span');
      top.textContent = '추가된 프로젝트 파라미터 이름이 중복되는지 확인합니다.';
      sub.textContent = '전체 검토 또는 지정 파라미터만 검토하도록 설정할 수 있습니다.';
      summary.append(top, sub);
      row.append(summary);
      state.ui.parameterDuplicationSummary = { row, top, sub };
      refreshParameterDuplicationFeatureSummary();
    } else if (key === 'parametermissing') {
      const summary = div('feature-row__summary');
      summary.style.display = 'grid';
      summary.style.gap = '4px';
      summary.style.padding = '10px 12px';
      summary.style.borderRadius = '14px';
      summary.style.border = '1px dashed var(--border-soft)';
      summary.style.background = 'var(--surface-help)';
      const top = document.createElement('strong');
      const sub = document.createElement('span');
      top.textContent = '조건별 객체 속성 추출과 같은 대상 필터로 지정 파라미터 누락을 검토합니다.';
      sub.textContent = '아직 검토 파라미터와 누락 예외 규칙이 지정되지 않았습니다.';
      summary.append(top, sub);
      row.append(summary);
      state.ui.parameterMissingSummary = { row, top, sub };
      refreshParameterMissingFeatureSummary();
    }

    config.displayTitle = meta.label || key;
    config.key = key;
    config.panel.classList.add('settings-panel', 'is-open');
    state.ui.panels[key] = config.panel;
    state.ui.controls[key] = config.controls || {};
    row.classList.toggle('is-active', !!state.features[key]?.enabled);
    row.dataset.selected = state.features[key]?.enabled ? 'true' : 'false';
    return row;
  }

  function buildPmsWorkflowRow() {
    return buildWorkflowLaunchRow({
      iconLabel: 'PMS',
      title: 'Segment-PMS 비교 검토',
      cardDesc: 'PMS 양식을 입력받아 Segment와 OD, ID를 비교 검토',
      infoDesc: 'PMS 양식을 입력받아 Segment와 OD, ID를 비교 검토합니다.',
      categoryLabel: '데이터',
      categoryTitle: '데이터 매핑/검토 기능',
      summary: '추출 → PMS 등록 → 매핑 준비 → 비교 실행 → 결과 내보내기',
      route: 'segmentpms'
    });
  }

  function buildParamModifierWorkflowRow() {
    return buildWorkflowLaunchRow({
      iconLabel: 'PM',
      title: '파라미터 수정기',
      cardDesc: '입력 조건 기반 필터링 대상으로\n지정 파라미터에 지정 속성 일괄입력',
      infoDesc: '입력 조건 기반 필터링 대상으로 지정 파라미터에 지정 속성 일괄입력',
      categoryLabel: '파라미터',
      categoryTitle: '파라미터 입력/수정 기능',
      summary: '활성 문서 또는 다중 RVT 선택 → 필터/조건/입력값 설정 → 실행 → 결과 엑셀/로그 저장',
      route: 'parammodifier'
    });
  }

  function buildGuidWorkflowRow() {
    return buildWorkflowLaunchRow({
      iconLabel: 'GUID',
      title: '파라미터 GUID 검토 및 정리',
      cardDesc: '검토 결과를 삭제용 엑셀로 내보내고\n삭제 표시 후 다시 불러와 정리',
      infoDesc: '프로젝트/패밀리 파라미터 GUID를 검토한 뒤 삭제용 엑셀에서 삭제 대상을 표시하고 다시 불러와 정리합니다.',
      categoryLabel: '파라미터',
      categoryTitle: '파라미터 검토/정리 기능',
      summary: "GUID 검토 실행 → 삭제용 엑셀 내보내기 → 엑셀에서 '삭제여부=삭제' 입력 → 삭제용 엑셀 불러오기",
      route: 'guid'
    });
  }

  function buildLinkPathWorkflowRow() {
    return buildWorkflowLaunchRow({
      iconLabel: 'LINK',
      title: 'Revit 링크 경로 추출/재지정',
      cardDesc: '닫힌 RVT의 링크 경로를 추출하고\n엑셀 기준으로 대상 경로를 반영',
      infoDesc: '닫힌 RVT의 Revit 링크 경로를 추출하고 엑셀에서 지정한 TargetLinkPath 기준으로 반영합니다.',
      categoryLabel: '링크',
      categoryTitle: '링크 경로 관리 기능',
      summary: 'RVT 등록 → 링크 추출 → 엑셀에서 대상 경로 지정 → 불러오기 → 적용',
      route: 'linkpath'
    });
  }

  function buildLateralNozzleWorkflowRow() {
    return buildWorkflowLaunchRow({
      iconLabel: 'KTA',
      title: '노즐코드 KTA 단일화',
      cardDesc: '접수받은 KTA 양식을\n정해진 하나의 시트 양식으로 추출',
      infoDesc: '접수받은 KTA 양식을 정해진 하나의 시트 양식으로 추출합니다.',
      categoryLabel: '엑셀',
      categoryTitle: '엑셀 추출/정리 기능',
      summary: '엑셀 파일 추가 → 헤더 블록 자동 탐색 → 누락/형식 검토 → 결과 시트 저장',
      route: 'lateralnozzle'
    });
  }

function buildConditionExtractWorkflowRow() {
  return buildWorkflowLaunchRow({
    iconLabel: 'BQC',
    title: '조건별 객체 대상 속성 추출',
    cardDesc: '조건식으로 객체를 추려\n지정 속성 + 좌표/선형 정보를 함께 추출',
    infoDesc: '조건식으로 객체를 추려 지정 속성과 좌표/선형 정보를 함께 추출합니다.',
    categoryLabel: 'BQC 검토',
    categoryTitle: '조건별 속성 추출 기능',
    summary: '활성 문서 또는 다중 RVT 선택 → 조건/추출 파라미터 설정 → 좌표/선형 옵션 선택 → 결과 엑셀 저장',
    route: 'conditionextract'
  });
}

  function buildParamPropWorkflowRow() {
    return buildWorkflowLaunchRow({
      iconLabel: 'SP',
      title: '패밀리 공유파라미터 추가/연동',
      cardDesc: '복합 및 하위 패밀리에 지정 파라미터를 추가하고 연동',
      infoDesc: '복합 및 하위 패밀리에 지정 파라미터를 추가하고 연동합니다.',
      categoryLabel: '패밀리',
      categoryTitle: '패밀리 연동/공유 파라미터 기능',
      summary: '공유 파라미터 선택 → 대상 그룹/모드 설정 → 연동 실행 → 결과 확인/엑셀 내보내기',
      route: 'paramprop'
    });
  }

  function buildDupWorkflowRow(options = {}) {
    return buildWorkflowLaunchRow({
      iconLabel: 'DUP',
      title: '중복 / 자체 간섭 검토',
      cardDesc: '활성 문서에서 중복 객체와 파일 내 간섭을 검토',
      infoDesc: '활성 문서에서 중복 객체와 파일 내 간섭을 검토합니다.',
      categoryLabel: options.chipLabel || 'RVT 검토',
      categoryTitle: options.chipTitle || '모델 검토/검사 기능',
        summary: '중복 검토 ↔ 자체 간섭 전환 → 규칙 설정 → 검토 실행 → 결과 확인/엑셀 내보내기',
      route: 'dup'
    });
  }

  function buildTapAlignWorkflowRow(options = {}) {
    return buildWorkflowLaunchRow({
      iconLabel: 'TAP',
      title: '탭/분기 축 틀어짐 검토',
      cardDesc: '탭/분기 피팅의 삽입축이\n연결 라인 중심축을 통과하는지 검토',
      infoDesc: '탭/분기 피팅의 삽입축이 연결된 배관/덕트 중심축을 정확히 통과하는지 검토합니다.',
      categoryLabel: options.chipLabel || 'BQC 보조',
      categoryTitle: options.chipTitle || 'BQC 보조 검토 기능',
      summary: '활성 문서 기준 → 허용 범위/공종 선택 → 탭/분기 피팅 검토 → 결과 확인/엑셀 내보내기',
      route: 'tapalign'
    });
  }

  function buildDeliveryCleanerWorkflowRow(options = {}) {
    return buildWorkflowLaunchRow({
      iconLabel: 'RVT',
      title: 'RVT 정리 (납품용)',
      cardDesc: '납품 파일 작성을 위한 뷰 정리/불필요 항목 제거/속성 추출',
      infoDesc: '납품 파일 작성을 위한 뷰 정리, 불필요 항목 제거(Purge), 검토용 속성 추출',
      categoryLabel: options.chipLabel || 'RVT 검토',
      categoryTitle: options.chipTitle || '모델 검토/정리 기능',
      summary: 'RVT 선택 → 정리 옵션 설정 → 정리 실행 → 검토/속성 추출 → 불필요 항목 제거',
      route: 'deliverycleaner'
    });
  }

  function buildSharedParamBatchRow() {
    return buildWorkflowLaunchRow({
      iconLabel: 'SP',
      title: '프로젝트 파라미터 일괄 추가',
      cardDesc: '프로젝트 파일에 지정 파라미터를\n일괄 추가',
      infoDesc: '프로젝트 파일에 지정 파라미터를 일괄 추가합니다.',
      categoryLabel: '파라미터',
      categoryTitle: '파라미터 추가/바인딩 기능',
      summary: '파라미터 선택 → 바인딩 설정 → RVT 실행 → 로그/엑셀',
      route: 'sharedparambatch'
    });
  }

  function buildWorkflowLaunchRow(options = {}) {
    const row = div('feature-row feature-row--workflow feature-row--launch');
    const entryId = options.entryId || options.route || '';
    row.dataset.route = options.route || '';
    row.dataset.entryId = entryId;
    const header = div('feature-row__header');
    const left = div('feature-row__left');
    const text = div('feature-row__text');
    const title = document.createElement('strong');
    title.textContent = options.title || '';
    text.append(title);
    left.append(text);

    const right = div('feature-row__right');
    const chip = document.createElement('span');
    chip.className = 'chip chip--info';
    chip.textContent = options.categoryLabel || '기능';
    chip.setAttribute('aria-label', options.categoryTitle || `${options.categoryLabel || '기능'} 분류`);
    right.append(chip);

    const desc = div('feature-row__desc');
    desc.textContent = options.cardDesc || options.desc || options.infoDesc || '';
    const summary = div('feature-row__summary');
    summary.textContent = options.summary || '';
    header.append(left, right);
    row.append(header, desc, summary);
    applyFeatureRowTooltip(row, [options.title || '', options.infoDesc || options.desc || '', options.summary || ''], {
      title: options.title || '워크플로우',
      desc: options.infoDesc || options.desc || options.summary || ''
    });
    row.classList.add('is-clickable');
    row.setAttribute('role', 'button');
    row.setAttribute('aria-label', `${options.title || '워크플로우'} 열기`);
    row.tabIndex = 0;
    if (getHubEntry(entryId)) bindHubEntryContextMenu(row, entryId);
    const navigate = () => {
      if (getHubEntry(entryId)) openHubEntry(entryId);
      else navigateToFeatureRoute(options.route || '');
    };
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

  function applyPendingFeatureFocus(scopeRoot, routeName = 'multi') {
    const pending = consumePendingHubEntry(routeName);
    if (!pending?.entryId || !scopeRoot) return;

    window.setTimeout(() => {
      const row = scopeRoot.querySelector(`[data-entry-id="${pending.entryId}"]`);
      if (!row) return;
      row.classList.add('feature-row--spotlight');
      if (typeof row.scrollIntoView === 'function') {
        row.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
      if (typeof row.focus === 'function') {
        row.focus({ preventScroll: true });
      }
      window.setTimeout(() => {
        row.classList.remove('feature-row--spotlight');
      }, 1800);
    }, 120);
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

    const tol = makeField('허용 범위', 'tol', '', 'number');
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

    const pointXY = makeCheckboxField('좌표 X/Y 추출');
    pointXY.input.checked = !!state.features.connector.configDraft.includePointXY;
    pointXY.input.addEventListener('change', () => {
      state.features.connector.configDraft.includePointXY = pointXY.input.checked;
      markFeatureDirty('connector');
      refreshConnectorFeatureSummary();
    });

    const linearMetrics = makeCheckboxField('선형 길이 / 방향 벡터 추출');
    linearMetrics.input.checked = !!state.features.connector.configDraft.includeLinearMetrics;
    linearMetrics.input.addEventListener('change', () => {
      state.features.connector.configDraft.includeLinearMetrics = linearMetrics.input.checked;
      markFeatureDirty('connector');
      refreshConnectorFeatureSummary();
    });

    const excludeEndDummy = makeCheckboxField('End + Dummy 패밀리 제외');
    excludeEndDummy.input.checked = !!state.features.connector.configDraft.excludeEndDummy;
    excludeEndDummy.input.addEventListener('change', () => {
      state.features.connector.configDraft.excludeEndDummy = excludeEndDummy.input.checked;
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

    const extractOptionsCard = div('feature-row__summary');
    extractOptionsCard.style.display = 'grid';
    extractOptionsCard.style.gap = '8px';
    extractOptionsCard.style.padding = '10px';
    extractOptionsCard.style.borderRadius = '14px';
    extractOptionsCard.style.border = '1px solid var(--border-accent-soft)';
    extractOptionsCard.style.background = 'var(--surface-control)';

    const extractOptionsTitle = document.createElement('strong');
    extractOptionsTitle.textContent = '추가 추출 옵션';
    extractOptionsTitle.style.fontSize = '12px';
    extractOptionsTitle.style.lineHeight = '1.3';

    pointXY.field.style.margin = '0';
    linearMetrics.field.style.margin = '0';
    excludeEndDummy.field.style.margin = '0';
    pointXY.field.style.padding = '0';
    linearMetrics.field.style.padding = '0';
    excludeEndDummy.field.style.padding = '0';
    pointXY.field.style.border = '0';
    linearMetrics.field.style.border = '0';
    excludeEndDummy.field.style.border = '0';
    pointXY.field.style.background = 'transparent';
    linearMetrics.field.style.background = 'transparent';
    excludeEndDummy.field.style.background = 'transparent';

    extractOptionsCard.append(extractOptionsTitle, pointXY.field, linearMetrics.field, excludeEndDummy.field);
    basicsCard.append(extractOptionsCard);

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
      selectedCount.textContent = selected.length ? `${selected.length}개 선택` : '파라미터 선택 필요';
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
        action.textContent = selected.has(String(item.name || '').toLowerCase()) ? '선택 완료' : '추가';

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
    return {
      panel,
      controls: {
        tol,
        unit,
        pointXY,
        linearMetrics,
        excludeEndDummy,
        searchInput,
        listWrap,
        selectedWrap,
        renderConnectorList,
        renderConnectorSelected
      }
    };
  }

  function buildTapAlignConfig() {
    const panel = div('multi-config');
    panel.style.display = 'grid';
    panel.style.gap = '12px';
    panel.style.width = '100%';
    panel.style.maxWidth = 'none';
    panel.style.minWidth = '0';
    panel.style.boxSizing = 'border-box';

    const tol = makeField('허용 범위', 'tapalign_tol', '', 'number');
    tol.input.value = state.features.tapalign.configDraft.tol;
    tol.input.min = '0';
    tol.input.step = '0.01';
    tol.input.style.fontWeight = '600';
    const featureTargetFilter = makeField('기능 전용 필터', 'tapalign_feature_target_filter', '예: PM1=값; PM2=값2', 'text');
    featureTargetFilter.input.value = state.features.tapalign.configDraft.featureTargetFilter || '';

    const unit = makeSelectField('거리 단위', [
      { value: 'mm', label: 'mm' },
      { value: 'inch', label: 'inch' }
    ]);
    unit.select.value = state.features.tapalign.configDraft.unit;
    unit.select.style.fontWeight = '600';

    const domain = makeSelectField('검토 범위', [
      { value: 'all', label: '배관 + 덕트' },
      { value: 'pipe', label: '배관' },
      { value: 'duct', label: '덕트' }
    ]);
    domain.select.value = state.features.tapalign.configDraft.domain;

    const markDirty = () => {
      state.features.tapalign.configDraft.tol = Math.max(0, parseFloat(tol.input.value || '0.5') || 0.5);
      state.features.tapalign.configDraft.unit = normalizeTapAlignUnit(unit.select.value);
      state.features.tapalign.configDraft.domain = normalizeTapAlignDomain(domain.select.value);
      state.features.tapalign.configDraft.featureTargetFilter = String(featureTargetFilter.input.value || '').trim();
      markFeatureDirty('tapalign');
      refreshTapAlignFeatureSummary();
    };

    tol.input.addEventListener('change', markDirty);
    tol.input.addEventListener('blur', markDirty);
    unit.select.addEventListener('change', markDirty);
    domain.select.addEventListener('change', markDirty);
    featureTargetFilter.input.addEventListener('change', markDirty);
    featureTargetFilter.input.addEventListener('blur', markDirty);

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

    const fieldCards = [tol.field, unit.field, domain.field].map((field) => {
      const card = div('feature-row__summary');
      card.style.display = 'grid';
      card.style.gap = '8px';
      card.style.padding = '10px';
      card.style.borderRadius = '14px';
      card.style.border = '1px solid var(--border-accent-soft)';
      card.style.background = 'var(--surface-control)';

      field.style.margin = '0';
      field.style.padding = '0';
      field.style.border = '0';
      field.style.background = 'transparent';
      field.style.display = 'grid';
      field.style.gap = '6px';
      field.style.minWidth = '0';
      field.style.boxSizing = 'border-box';

      const control = field.querySelector('input, select');
      if (control) {
        control.style.width = '100%';
        control.style.boxSizing = 'border-box';
        control.style.padding = '8px 10px';
        control.style.borderRadius = '12px';
        control.style.border = '1px solid var(--border-soft)';
        control.style.background = 'var(--surface-control)';
        control.style.minWidth = '0';
      }

      card.append(field);
      return card;
    });

    basics.append(...fieldCards);
    basicsCard.append(basicsTitle, basics);

    const featureFilterCard = div('feature-row__summary');
    featureFilterCard.style.display = 'grid';
    featureFilterCard.style.gap = '8px';
    featureFilterCard.style.padding = '12px';
    featureFilterCard.style.borderRadius = '18px';
    featureFilterCard.style.border = '1px solid var(--border-accent-soft)';
    featureFilterCard.style.background = 'var(--surface-elevated)';
    featureFilterCard.style.width = '100%';
    featureFilterCard.style.boxSizing = 'border-box';

    const featureFilterTitle = document.createElement('strong');
    featureFilterTitle.textContent = '기능 전용 필터';

    featureTargetFilter.field.style.margin = '0';
    featureTargetFilter.field.style.padding = '0';
    featureTargetFilter.field.style.border = '0';
    featureTargetFilter.field.style.background = 'transparent';
    featureTargetFilter.field.style.display = 'grid';
    featureTargetFilter.field.style.gap = '6px';
    featureTargetFilter.field.style.minWidth = '0';
    featureTargetFilter.field.style.boxSizing = 'border-box';
    featureTargetFilter.input.style.width = '100%';
    featureTargetFilter.input.style.boxSizing = 'border-box';
    featureTargetFilter.input.style.padding = '8px 10px';
    featureTargetFilter.input.style.borderRadius = '12px';
    featureTargetFilter.input.style.border = '1px solid var(--border-soft)';
    featureTargetFilter.input.style.background = 'var(--surface-control)';
    featureTargetFilter.input.style.minWidth = '0';

    const featureFilterHint = document.createElement('div');
    featureFilterHint.className = 'feature-note';
    featureFilterHint.textContent = '공통 검토 대상 필터에 이어 추가 AND 조건으로 적용됩니다.';

    featureFilterCard.append(featureFilterTitle, featureTargetFilter.field, featureFilterHint);

    const commonCard = div('feature-row__summary');
    commonCard.style.display = 'grid';
    commonCard.style.gap = '8px';
    commonCard.style.padding = '12px';
    commonCard.style.borderRadius = '18px';
    commonCard.style.border = '1px solid var(--border-accent-soft)';
    commonCard.style.background = 'var(--surface-help)';
    commonCard.style.width = '100%';
    commonCard.style.boxSizing = 'border-box';

    const commonTitle = document.createElement('strong');
    commonTitle.textContent = '공통 옵션 반영';

    const buildCommonRow = (label) => {
      const row = document.createElement('div');
      row.style.display = 'grid';
      row.style.gap = '4px';
      row.style.padding = '10px 12px';
      row.style.borderRadius = '14px';
      row.style.border = '1px solid var(--border-soft)';
      row.style.background = 'var(--surface-elevated)';
      row.style.minWidth = '0';

      const name = document.createElement('strong');
      name.textContent = label;
      name.style.fontSize = '12px';
      name.style.lineHeight = '1.3';

      const value = document.createElement('span');
      value.style.minWidth = '0';
      value.style.overflowWrap = 'anywhere';
      value.style.wordBreak = 'break-word';
      value.style.lineHeight = '1.4';
      value.style.color = 'var(--muted,#64748b)';

      row.append(name, value);
      return { row, value };
    };

    const commonExtra = buildCommonRow('추가 추출 파라미터');
    const commonFilter = buildCommonRow('검토 대상 필터');
    const commonExcludeFilter = buildCommonRow('검토 제외 대상 필터');
    const commonOptions = buildCommonRow('추가 추출 옵션');
    const commonHint = document.createElement('div');
    commonHint.className = 'feature-note';
    commonHint.textContent = '공통 옵션은 BQC 공통 설정에서 관리한 값을 그대로 반영합니다.';

    const renderCommonSummary = () => {
      const committed = state.common.configCommitted || {};
      const extras = String(committed.extraParams || '')
        .split(',')
        .map((value) => value.trim())
        .filter(Boolean);
      const filterText = String(committed.targetFilter || '').trim();
      const excludeFilterText = String(committed.excludeTargetFilter || '').trim();
      const optionParts = [];
      if (committed.includePointXY) optionParts.push('좌표 X/Y');
      if (committed.includeLinearMetrics) optionParts.push('선형 길이/방향');

      const extraText = extras.length ? extras.join(', ') : '추가 파라미터가 없습니다.';
      const optionText = optionParts.length ? optionParts.join(', ') : '추가 추출 옵션이 없습니다.';

      commonExtra.value.textContent = extraText;
      commonExtra.value.setAttribute('aria-label', extraText);
      commonFilter.value.textContent = filterText || '필터가 없습니다.';
      commonFilter.value.setAttribute('aria-label', filterText || '필터가 없습니다.');
      commonExcludeFilter.value.textContent = excludeFilterText || '제외 필터가 없습니다.';
      commonExcludeFilter.value.setAttribute('aria-label', excludeFilterText || '제외 필터가 없습니다.');
      commonOptions.value.textContent = optionText;
      commonOptions.value.setAttribute('aria-label', optionText);
    };

    renderCommonSummary();
    commonCard.append(commonTitle, commonExtra.row, commonFilter.row, commonExcludeFilter.row, commonOptions.row, commonHint);

    panel.append(basicsCard, featureFilterCard, commonCard);
    return {
      panel,
      controls: {
        tol,
        unit,
        domain,
        featureTargetFilter,
        renderCommonSummary
      }
    };
  }

  function buildDupClashConfig() {
    const panel = div('multi-config');
    panel.style.display = 'grid';
    panel.style.gap = '12px';
    panel.style.width = '100%';
    panel.style.maxWidth = 'none';
    panel.style.minWidth = '0';
    panel.style.boxSizing = 'border-box';

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
    basicsTitle.textContent = '검토 모드';
    basicsTitle.style.fontSize = '13px';
    basicsTitle.style.lineHeight = '1.3';

    const modeGrid = div('multi-config');
    modeGrid.style.display = 'grid';
    modeGrid.style.gridTemplateColumns = 'repeat(2, minmax(0, 1fr))';
    modeGrid.style.gap = '10px';
    modeGrid.style.alignItems = 'stretch';

    const createModeCard = (value, titleText, descText) => {
      const label = document.createElement('label');
      label.style.display = 'grid';
      label.style.gap = '8px';
      label.style.padding = '12px';
      label.style.borderRadius = '14px';
      label.style.border = '1px solid var(--border-accent-soft)';
      label.style.background = 'var(--surface-control)';
      label.style.cursor = 'pointer';
      label.style.alignContent = 'start';

      const input = document.createElement('input');
      input.type = 'radio';
      input.name = 'dupclash-mode';
      input.value = value;

      const title = document.createElement('strong');
      title.textContent = titleText;

      const desc = document.createElement('span');
      desc.textContent = descText;
      desc.style.color = 'var(--muted,#64748b)';
      desc.style.lineHeight = '1.45';

      label.append(input, title, desc);
      return { label, input };
    };

    const duplicateMode = createModeCard('duplicate', '중복 검토', '완전 중복으로 판단되는 객체만 묶어서 결과에 남깁니다.');
    const clashMode = createModeCard('clash', '자체 간섭 검토', '연결되지 않은 묻힘/겹침 같은 자체 간섭 대상만 결과에 남깁니다.');

    const markDirty = () => {
      const nextMode = duplicateMode.input.checked ? 'duplicate' : 'clash';
      state.features.dupclash.configDraft.mode = normalizeDupClashMode(nextMode);
      state.features.dupclash.configDraft.tolFeet = Number(state.features.dupclash.configDraft.tolFeet) > 0
        ? Number(state.features.dupclash.configDraft.tolFeet)
        : 1 / 64;
      markFeatureDirty('dupclash');
      refreshDupClashFeatureSummary();
    };

    duplicateMode.input.addEventListener('change', markDirty);
    clashMode.input.addEventListener('change', markDirty);

    modeGrid.append(duplicateMode.label, clashMode.label);
    basicsCard.append(basicsTitle, modeGrid);

    const commonCard = div('feature-row__summary');
    commonCard.style.display = 'grid';
    commonCard.style.gap = '8px';
    commonCard.style.padding = '12px';
    commonCard.style.borderRadius = '18px';
    commonCard.style.border = '1px solid var(--border-accent-soft)';
    commonCard.style.background = 'var(--surface-help)';
    commonCard.style.width = '100%';
    commonCard.style.boxSizing = 'border-box';

    const commonTitle = document.createElement('strong');
    commonTitle.textContent = '공통 옵션 반영';

    const commonSummary = document.createElement('span');
    commonSummary.style.color = 'var(--muted,#64748b)';
    commonSummary.style.lineHeight = '1.5';

    const commonNote = div('feature-note');
    commonNote.textContent = '기존 중복 / 자체 간섭 화면의 세부 필터는 쓰지 않고, 공통 설정의 포함/제외 대상 필터를 그대로 적용합니다.';

    const renderCommonSummary = () => {
      const committed = state.common.configCommitted || {};
      const extras = String(committed.extraParams || '')
        .split(',')
        .map((value) => value.trim())
        .filter(Boolean);
      const filterText = String(committed.targetFilter || '').trim() || '필터가 없습니다.';
      const excludeText = String(committed.excludeTargetFilter || '').trim() || '제외 필터가 없습니다.';
      commonSummary.textContent = `공통 포함 필터 ${filterText} · 제외 필터 ${excludeText} · 추가 파라미터 ${extras.length}개`;
    };

    renderCommonSummary();
    commonCard.append(commonTitle, commonSummary, commonNote);

    panel.append(basicsCard, commonCard);
    return {
      panel,
      controls: {
        modeDuplicate: duplicateMode,
        modeClash: clashMode,
        renderCommonSummary
      }
    };
  }

  function buildFloorInfoConfig() {
    const panel = div('multi-config');
    const levelInputState = new Map();
    panel.classList.add('floorinfo-config');
    panel.style.display = 'flex';
    panel.style.flexDirection = 'column';
    panel.style.alignItems = 'stretch';
    panel.style.justifyContent = 'flex-start';
    panel.style.gap = '12px';
    panel.style.width = '100%';
    panel.style.maxWidth = 'none';
    panel.style.minWidth = '0';
    panel.style.padding = '0';
    panel.style.border = '0';
    panel.style.background = 'transparent';
    panel.style.boxShadow = 'none';
    panel.style.boxSizing = 'border-box';

    const param = makeField('층정보 파라미터명', 'floorinfo-param', '예: FloorInfo', 'text');
    param.field.style.margin = '0';
    param.field.style.alignSelf = 'stretch';
    param.field.style.width = '100%';
    param.field.style.minWidth = '0';
    param.input.value = state.features.floorinfo.configDraft.parameterName || '';
    param.input.addEventListener('input', () => {
      state.features.floorinfo.configDraft.parameterName = param.input.value || '';
      markFeatureDirty('floorinfo');
      refreshFloorInfoFeatureSummary();
    });

    const controlsCard = div('feature-row__summary floorinfo-config__controls');
    controlsCard.style.display = 'grid';
    controlsCard.style.gap = '12px';
    controlsCard.style.padding = '12px';
    controlsCard.style.borderRadius = '16px';
    controlsCard.style.border = '1px solid var(--border-accent-soft)';
    controlsCard.style.background = 'var(--surface-elevated)';
    controlsCard.style.boxShadow = 'var(--shadow-soft)';
    controlsCard.style.width = '100%';
    controlsCard.style.minWidth = '0';
    controlsCard.style.boxSizing = 'border-box';

    const controlsMeta = document.createElement('div');
    controlsMeta.className = 'floorinfo-config__control-note feature-note';
    controlsMeta.textContent = '층정보 영역을 구분할 레벨만 선택해 주세요. 선택하지 않은 중간 레벨은 구간 계산에서 무시되며, 관통 객체는 시작하는 가장 아래 구간의 기대 층정보 값으로 판정합니다.';

    const sourceCard = div('feature-row__summary floorinfo-config__rules');
    sourceCard.style.display = 'grid';
    sourceCard.style.gap = '8px';
    sourceCard.style.padding = '12px';
    sourceCard.style.borderRadius = '16px';
    sourceCard.style.border = '1px solid var(--border-accent-soft)';
    sourceCard.style.background = 'var(--surface-elevated)';
    sourceCard.style.boxShadow = 'var(--shadow-soft)';
    sourceCard.style.width = '100%';
    sourceCard.style.minWidth = '0';
    sourceCard.style.boxSizing = 'border-box';

    const sourceHead = document.createElement('div');
    sourceHead.className = 'floorinfo-config__header';
    sourceHead.style.display = 'flex';
    sourceHead.style.justifyContent = 'space-between';
    sourceHead.style.alignItems = 'center';
    sourceHead.style.gap = '10px';
    sourceHead.style.flexWrap = 'wrap';
    const sourceTitle = document.createElement('strong');
    sourceTitle.textContent = '활성 문서 레벨 기준';
    const refreshBtn = document.createElement('button');
    refreshBtn.type = 'button';
    refreshBtn.className = 'btn btn--secondary';
    refreshBtn.textContent = '레벨 새로고침';
    refreshBtn.addEventListener('click', () => requestFloorInfoConfig('settings-refresh'));
    sourceHead.append(sourceTitle, refreshBtn);

    const documentMeta = document.createElement('div');
    documentMeta.style.fontSize = '12px';
    documentMeta.style.color = 'var(--muted,#64748b)';

    const warningBox = div('feature-note');
    warningBox.classList.add('floorinfo-config__warning');
    warningBox.style.display = 'none';
    warningBox.style.whiteSpace = 'pre-wrap';

    const tableWrap = div('paramprop-table-wrap floorinfo-config__table');
    tableWrap.style.maxHeight = '420px';
    tableWrap.style.border = '1px solid var(--border-soft)';
    tableWrap.style.borderRadius = '14px';
    tableWrap.style.background = 'var(--surface-control)';
    tableWrap.style.overflow = 'auto';
    const table = document.createElement('table');
    table.className = 'paramprop-table';
    table.innerHTML = `
      <thead>
        <tr>
          <th>영역 기준</th>
          <th>레벨명</th>
          <th>모델 기준 Z(mm)</th>
          <th>기대 층정보 값</th>
        </tr>
      </thead>
      <tbody></tbody>`;
    const tbody = table.querySelector('tbody');
    tableWrap.append(table);

    controlsCard.append(param.field, controlsMeta);
    sourceCard.append(sourceHead, documentMeta, warningBox, tableWrap);
    panel.append(controlsCard, sourceCard);

    function updateWarnings() {
      const warnings = Array.isArray(state.features.floorinfo.configDraft.warnings)
        ? state.features.floorinfo.configDraft.warnings.filter(Boolean)
        : [];
      warningBox.style.display = warnings.length ? 'block' : 'none';
      warningBox.textContent = warnings.length ? warnings.join('\n') : '';
    }

    function renderRules() {
      const draft = state.features.floorinfo.configDraft;
      const rules = normalizeFloorInfoRules(draft.levelRules);
      levelInputState.clear();
      documentMeta.textContent = draft.documentTitle
        ? `활성 문서: ${draft.documentTitle}`
        : '활성 문서 기준으로 레벨과 모델 기준 Z를 불러옵니다.';
      updateWarnings();
      tbody.innerHTML = '';

      if (!rules.length) {
        const row = document.createElement('tr');
        const cell = document.createElement('td');
        cell.colSpan = 4;
        cell.className = 'paramprop-empty';
        cell.textContent = '레벨 목록이 없습니다. 활성 문서를 열고 레벨 새로고침을 눌러 주세요.';
        row.append(cell);
        tbody.append(row);
        return;
      }

      rules.forEach((rule) => {
        const row = document.createElement('tr');
        row.classList.toggle('is-selected', rule.useAsBoundary !== false);
        row.classList.toggle('is-inactive', rule.useAsBoundary === false);

        const boundaryCell = document.createElement('td');
        boundaryCell.style.textAlign = 'center';
        const toggle = document.createElement('input');
        toggle.type = 'checkbox';
        toggle.checked = rule.useAsBoundary !== false;
        toggle.setAttribute('aria-label', '이 레벨을 층정보 영역 경계로 사용');
        toggle.addEventListener('change', () => {
          draft.levelRules = rules.map((item) => item.levelName === rule.levelName
            ? { ...item, useAsBoundary: toggle.checked }
            : item);
          markFeatureDirty('floorinfo');
          renderRules();
          refreshFloorInfoFeatureSummary();
        });
        boundaryCell.append(toggle);

        const nameCell = document.createElement('td');
        nameCell.textContent = rule.levelName;

        const zCell = document.createElement('td');
        zCell.textContent = Number(rule.absoluteZMm || 0).toLocaleString('ko-KR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });

        const valueCell = document.createElement('td');
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'paramprop-search';
        input.style.width = '100%';
        input.style.minWidth = '160px';
        input.value = rule.expectedValue || '';
        input.placeholder = rule.useAsBoundary === false ? '영역 기준 아님' : '예: 1F';
        input.disabled = rule.useAsBoundary === false;
        input.addEventListener('input', () => {
          draft.levelRules = rules.map((item) => item.levelName === rule.levelName
            ? { ...item, expectedValue: input.value || '' }
            : item);
          markFeatureDirty('floorinfo');
          refreshFloorInfoFeatureSummary();
        });
        valueCell.append(input);
        levelInputState.set(rule.levelName.toLowerCase(), {
          levelName: rule.levelName,
          absoluteZFt: rule.absoluteZFt,
          absoluteZMm: rule.absoluteZMm,
          toggle,
          input
        });

        row.append(boundaryCell, nameCell, zCell, valueCell);
        tbody.append(row);
      });
    }

    function collectDraft() {
      const draft = state.features.floorinfo.configDraft;
      const currentRules = normalizeFloorInfoRules(draft.levelRules);
      const currentMap = new Map(currentRules.map((rule) => [rule.levelName.toLowerCase(), rule]));
      const nextRules = [];

      levelInputState.forEach((entry, key) => {
        const existing = currentMap.get(key);
        nextRules.push({
          levelName: entry.levelName,
          absoluteZFt: Number(entry.absoluteZFt ?? existing?.absoluteZFt) || 0,
          absoluteZMm: Number(entry.absoluteZMm ?? existing?.absoluteZMm) || 0,
          useAsBoundary: !!entry.toggle?.checked,
          expectedValue: String(entry.input?.value || '')
        });
      });

      draft.parameterName = String(param.input?.value || '');
      if (nextRules.length) {
        draft.levelRules = normalizeFloorInfoRules(nextRules);
      } else {
        draft.levelRules = normalizeFloorInfoRules(draft.levelRules);
      }
      return draft;
    }

    function applySnapshot(payload) {
      collectDraft();
      const draft = state.features.floorinfo.configDraft;
      if (payload?.ok === false) {
        draft.documentTitle = '';
        draft.levelRules = [];
        draft.warnings = [payload?.message || '레벨 정보를 불러오지 못했습니다.'];
        renderRules();
        return;
      }
      const levels = Array.isArray(payload?.levels) ? payload.levels : [];
      draft.documentTitle = payload?.documentTitle || '';
      draft.levelRules = mergeFloorInfoRules(draft.levelRules, levels);
      draft.warnings = Array.isArray(payload?.warnings) ? payload.warnings : [];
      renderRules();
      refreshFloorInfoFeatureSummary();
    }

    buildFloorInfoConfig.applySnapshot = applySnapshot;
    renderRules();
    return { panel, controls: { param, renderRules, updateWarnings, collectDraft } };
  }

  function buildFamilySuitabilityConfig() {
    const panel = div('multi-config');
    panel.classList.add('familysuitability-config');
    panel.style.width = '100%';
    panel.style.minWidth = '0';

    const presetCard = div('feature-row__summary familysuitability-card familysuitability-card--full');
    const presetTitle = document.createElement('strong');
    presetTitle.textContent = '설정 저장 / 불러오기';
    const presetNote = div('feature-note');
    presetNote.textContent = '최근 적용한 설정은 자동으로 기억합니다. 자주 쓰는 조합은 이름을 붙여 저장해 두고 다시 불러올 수 있습니다.';
    const presetSelect = makeSelectField('저장된 설정', [
      { value: '', label: '저장된 설정이 없습니다.' }
    ]);
    presetSelect.field.style.margin = '0';
    const presetName = makeField('설정 이름', 'familysuitabilityPresetName', '예: 기계 납품 기본', 'text');
    presetName.field.style.margin = '0';
    const presetActions = div('feature-row__actions');
    presetActions.style.display = 'flex';
    presetActions.style.flexWrap = 'wrap';
    presetActions.style.gap = '8px';

    const savePresetBtn = document.createElement('button');
    savePresetBtn.type = 'button';
    savePresetBtn.className = 'btn btn--secondary';
    savePresetBtn.textContent = '설정 저장';

    const loadPresetBtn = document.createElement('button');
    loadPresetBtn.type = 'button';
    loadPresetBtn.className = 'btn btn--secondary';
    loadPresetBtn.textContent = '설정 불러오기';

    const deletePresetBtn = document.createElement('button');
    deletePresetBtn.type = 'button';
    deletePresetBtn.className = 'btn btn--ghost';
    deletePresetBtn.textContent = '설정 삭제';

    presetActions.append(savePresetBtn, loadPresetBtn, deletePresetBtn);
    presetCard.append(presetTitle, presetNote, presetSelect.field, presetName.field, presetActions);

    const criteriaCard = div('feature-row__summary familysuitability-card');

    const criteriaHead = document.createElement('div');
    criteriaHead.style.display = 'flex';
    criteriaHead.style.justifyContent = 'space-between';
    criteriaHead.style.alignItems = 'center';
    criteriaHead.style.gap = '10px';
    criteriaHead.style.flexWrap = 'wrap';
    const criteriaTitle = document.createElement('strong');
    criteriaTitle.textContent = '기준 엑셀';
    const criteriaActions = div('feature-row__actions');
    criteriaActions.style.display = 'flex';
    criteriaActions.style.gap = '8px';
    const browseBtn = document.createElement('button');
    browseBtn.type = 'button';
    browseBtn.className = 'btn btn--secondary';
    browseBtn.textContent = '엑셀 선택';
    const clearBtn = document.createElement('button');
    clearBtn.type = 'button';
    clearBtn.className = 'btn btn--ghost';
    clearBtn.textContent = '비우기';
    criteriaActions.append(browseBtn, clearBtn);
    criteriaHead.append(criteriaTitle, criteriaActions);

    const criteriaPath = makeField('기준 파일 경로', 'familysuitability-criteria', '카테고리/패밀리/타입 헤더가 있는 엑셀', 'text');
    criteriaPath.input.readOnly = true;
    criteriaPath.input.placeholder = '기준 엑셀 파일을 선택해 주세요.';
    criteriaPath.field.style.margin = '0';
    const criteriaSummary = document.createElement('div');
    criteriaSummary.className = 'feature-note';

    criteriaCard.append(criteriaHead, criteriaPath.field, criteriaSummary);

    const reviewCard = div('feature-row__summary familysuitability-card');

    const reviewTitle = document.createElement('strong');
    reviewTitle.textContent = '기본 검토 문구';
    const matchReview = makeField('기준 일치 문구', 'familysuitability-match-review', '예: 기준 리스트와 일치', 'text');
    const mismatchReview = makeField('기준 미일치 문구', 'familysuitability-mismatch-review', '예: 기준 리스트에 없는 조합', 'text');
    reviewCard.append(reviewTitle, matchReview.field, mismatchReview.field);

    const filterCard = div('feature-row__summary familysuitability-card familysuitability-card--filter');

    const filterHead = document.createElement('div');
    filterHead.style.display = 'flex';
    filterHead.style.justifyContent = 'space-between';
    filterHead.style.alignItems = 'center';
    filterHead.style.gap = '10px';
    filterHead.style.flexWrap = 'wrap';
    const filterTitle = document.createElement('strong');
    filterTitle.textContent = '이름 포함 필터';
    const addFilterBtn = document.createElement('button');
    addFilterBtn.type = 'button';
    addFilterBtn.className = 'btn btn--secondary';
    addFilterBtn.textContent = '필터 추가';
    filterHead.append(filterTitle, addFilterBtn);

    const filterGuide = document.createElement('div');
    filterGuide.className = 'feature-note';
    filterGuide.textContent = '필터 규칙은 OR 조건으로 평가합니다. 하나라도 일치하면 필터 검토 문구를 우선 적용하고, 여러 개가 동시에 일치하면 위에서 아래 순서로 먼저 등록된 문구를 사용합니다.';

    const filterList = div('familysuitability-filter-list');

    filterCard.append(filterHead, filterGuide, filterList);
    panel.append(presetCard, criteriaCard, reviewCard, filterCard);

    function ensureFilterRules() {
      const draft = state.features.familysuitability.configDraft;
      draft.filterRules = normalizeFamilySuitabilityFilterRules(draft.filterRules, { keepEmpty: true });
      if (!draft.filterRules.length) {
        draft.filterRules = [createEmptyFamilySuitabilityFilterRule()];
      }
    }

    function renderCriteriaSummary() {
      const draft = state.features.familysuitability.configDraft;
      const comboCount = Number(draft.criteriaComboCount) || 0;
      const rowCount = Number(draft.criteriaRowCount) || 0;
      const sheetCount = Number(draft.criteriaSheetCount) || 0;
      const fileLabel = getFamilySuitabilityCriteriaLabel(draft.criteriaExcelPath);
      criteriaPath.input.value = draft.criteriaExcelPath || '';
      if (!draft.criteriaExcelPath) {
        criteriaSummary.textContent = '기준 엑셀을 선택하면 카테고리/패밀리/타입 헤더를 찾아 조합 수를 바로 확인합니다.';
        return;
      }
      const stats = [];
      if (sheetCount) stats.push(`시트 ${sheetCount}개`);
      if (rowCount) stats.push(`원본 행 ${rowCount}개`);
      if (comboCount) stats.push(`유효 조합 ${comboCount}개`);
      criteriaSummary.textContent = stats.length
        ? `${fileLabel} · ${stats.join(' · ')}`
        : `${fileLabel} · 기준 엑셀을 사용합니다.`;
    }

    function renderPresetOptions(selectedName = '') {
      const presets = loadFamilySuitabilityPresets();
      const nextSelected = String(selectedName || presetSelect.select.value || '').trim();
      presetSelect.select.innerHTML = '';

      const defaultOption = document.createElement('option');
      defaultOption.value = '';
      defaultOption.textContent = presets.length ? '저장된 설정 선택' : '저장된 설정이 없습니다.';
      presetSelect.select.append(defaultOption);

      presets.forEach((preset) => {
        const option = document.createElement('option');
        option.value = preset.name;
        option.textContent = preset.name;
        presetSelect.select.append(option);
      });

      if (nextSelected && presets.some((preset) => preset.name === nextSelected)) {
        presetSelect.select.value = nextSelected;
      } else {
        presetSelect.select.value = '';
      }
    }

    function renderFilterRules() {
      ensureFilterRules();
      const draft = state.features.familysuitability.configDraft;
      filterList.innerHTML = '';
      draft.filterRules.forEach((rule, index) => {
        const item = div('familysuitability-filter-item');

        const targetField = div('field familysuitability-target-field');
        targetField.style.margin = '0';
        targetField.style.minWidth = '0';
        const targetLabel = document.createElement('label');
        targetLabel.textContent = '대상';
        const targetOptions = div('familysuitability-target-options');
        const targetButtons = [];
        const targetValue = normalizeFamilySuitabilityFilterTarget(rule.target);
        [
          ['familyOrType', '패밀리/타입'],
          ['family', '패밀리'],
          ['type', '타입']
        ].forEach(([value, label]) => {
          const optionBtn = document.createElement('button');
          optionBtn.type = 'button';
          optionBtn.className = 'familysuitability-target-option';
          optionBtn.textContent = label;
          optionBtn.classList.toggle('is-active', value === targetValue);
          optionBtn.addEventListener('click', () => {
            targetButtons.forEach((entry) => entry.classList.toggle('is-active', entry === optionBtn));
            updateRule({ target: value });
          });
          targetButtons.push(optionBtn);
          targetOptions.append(optionBtn);
        });
        targetField.append(targetLabel, targetOptions);

        const keywordField = makeField('포함 키워드', `familysuitability-filter-keyword-${index}`, '', 'text');
        keywordField.input.value = rule.keyword || '';
        keywordField.field.style.margin = '0';
        keywordField.field.style.minWidth = '0';

        const reviewField = makeField('출력 검토 문구', `familysuitability-filter-review-${index}`, '예: 키워드 기준 별도 검토', 'text');
        reviewField.input.value = rule.reviewText || '';
        reviewField.field.style.margin = '0';
        reviewField.field.style.minWidth = '0';

        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'btn btn--ghost';
        removeBtn.textContent = '삭제';

        const updateRule = (patch) => {
          draft.filterRules = draft.filterRules.map((entry, entryIndex) => (
            entryIndex === index
              ? { ...entry, ...patch }
              : entry
          ));
          markFeatureDirty('familysuitability');
          refreshFamilySuitabilityFeatureSummary();
        };

        keywordField.input.addEventListener('input', () => {
          updateRule({ keyword: keywordField.input.value || '' });
        });
        reviewField.input.addEventListener('input', () => {
          updateRule({ reviewText: reviewField.input.value || '' });
        });
        removeBtn.addEventListener('click', () => {
          const nextRules = draft.filterRules.filter((_, entryIndex) => entryIndex !== index);
          draft.filterRules = nextRules.length ? nextRules : [createEmptyFamilySuitabilityFilterRule()];
          markFeatureDirty('familysuitability');
          renderFilterRules();
          refreshFamilySuitabilityFeatureSummary();
        });

        const headRow = div('familysuitability-filter-item__head');
        headRow.append(targetField, removeBtn);

        const bodyGrid = div('familysuitability-filter-item__body');
        bodyGrid.append(keywordField.field, reviewField.field);

        item.append(headRow, bodyGrid);
        filterList.append(item);
      });
    }

    function collectDraft() {
      const draft = state.features.familysuitability.configDraft;
      draft.criteriaExcelPath = String(criteriaPath.input?.value || '').trim();
      draft.matchReviewText = String(matchReview.input?.value || '').trim();
      draft.mismatchReviewText = String(mismatchReview.input?.value || '').trim();
      draft.filterRules = normalizeFamilySuitabilityFilterRules(draft.filterRules, { keepEmpty: true });
      if (!draft.filterRules.length) {
        draft.filterRules = [createEmptyFamilySuitabilityFilterRule()];
      }
      return draft;
    }

    function applyCriteriaPicked(payload) {
      if (payload?.ok === false) {
        toast(payload?.message || '기준 엑셀을 읽지 못했습니다. 선택한 파일이 열 수 있는 엑셀 파일인지 확인해 주세요.', 'err');
        return;
      }
      const draft = state.features.familysuitability.configDraft;
      draft.criteriaExcelPath = String(payload?.path || '').trim();
      draft.criteriaRowCount = Number(payload?.rowCount) || 0;
      draft.criteriaComboCount = Number(payload?.uniqueCount) || 0;
      draft.criteriaSheetCount = Number(payload?.sheetCount) || 0;
      markFeatureDirty('familysuitability');
      renderCriteriaSummary();
      refreshFamilySuitabilityFeatureSummary();
      toast('패밀리 타입 적합성 기준 엑셀을 불러왔습니다.', 'ok');
    }

    browseBtn.addEventListener('click', () => {
      post('familysuitability:pick-criteria', { source: 'multi' });
    });
    clearBtn.addEventListener('click', () => {
      const draft = state.features.familysuitability.configDraft;
      draft.criteriaExcelPath = '';
      draft.criteriaRowCount = 0;
      draft.criteriaComboCount = 0;
      draft.criteriaSheetCount = 0;
      markFeatureDirty('familysuitability');
      renderCriteriaSummary();
      refreshFamilySuitabilityFeatureSummary();
    });
    matchReview.input.addEventListener('input', () => {
      state.features.familysuitability.configDraft.matchReviewText = matchReview.input.value || '';
      markFeatureDirty('familysuitability');
      refreshFamilySuitabilityFeatureSummary();
    });
    mismatchReview.input.addEventListener('input', () => {
      state.features.familysuitability.configDraft.mismatchReviewText = mismatchReview.input.value || '';
      markFeatureDirty('familysuitability');
      refreshFamilySuitabilityFeatureSummary();
    });
    addFilterBtn.addEventListener('click', () => {
      const draft = state.features.familysuitability.configDraft;
      draft.filterRules = normalizeFamilySuitabilityFilterRules(draft.filterRules, { keepEmpty: true });
      draft.filterRules.push(createEmptyFamilySuitabilityFilterRule());
      markFeatureDirty('familysuitability');
      renderFilterRules();
      refreshFamilySuitabilityFeatureSummary();
    });

    presetSelect.select.addEventListener('change', () => {
      presetName.input.value = presetSelect.select.value || '';
    });

    savePresetBtn.addEventListener('click', () => {
      const presetLabel = String(presetName.input.value || presetSelect.select.value || '').trim();
      if (!presetLabel) {
        toast('저장할 설정 이름을 입력해 주세요.', 'warn');
        return;
      }
      const currentDraft = collectDraft();
      saveFamilySuitabilityPreset(presetLabel, currentDraft);
      renderPresetOptions(presetLabel);
      presetName.input.value = presetLabel;
      toast(`패밀리 타입 적합성 설정을 저장했습니다: ${presetLabel}`, 'ok');
    });

    loadPresetBtn.addEventListener('click', () => {
      const presetLabel = String(presetSelect.select.value || presetName.input.value || '').trim();
      if (!presetLabel) {
        toast('불러올 설정을 선택해 주세요.', 'warn');
        return;
      }
      const loaded = applyFamilySuitabilityPreset(state.features.familysuitability.configDraft, presetLabel);
      if (!loaded) {
        toast('선택한 설정을 찾지 못했습니다.', 'err');
        renderPresetOptions();
        return;
      }
      markFeatureDirty('familysuitability');
      syncControlsFromDraft('familysuitability');
      renderPresetOptions(presetLabel);
      presetName.input.value = presetLabel;
      refreshFamilySuitabilityFeatureSummary();
      toast(`패밀리 타입 적합성 설정을 불러왔습니다: ${presetLabel}`, 'ok');
    });

    deletePresetBtn.addEventListener('click', () => {
      const presetLabel = String(presetSelect.select.value || presetName.input.value || '').trim();
      if (!presetLabel) {
        toast('삭제할 설정을 선택해 주세요.', 'warn');
        return;
      }
      deleteFamilySuitabilityPreset(presetLabel);
      renderPresetOptions();
      if (presetName.input.value === presetLabel) presetName.input.value = '';
      toast(`패밀리 타입 적합성 설정을 삭제했습니다: ${presetLabel}`, 'ok');
    });

    buildFamilySuitabilityConfig.applyCriteriaPicked = applyCriteriaPicked;
    renderPresetOptions();
    renderCriteriaSummary();
    renderFilterRules();
    return {
      panel,
      controls: {
        presetSelect,
        presetName,
        criteriaPath,
        criteriaSummary,
        matchReview,
        mismatchReview,
        renderPresetOptions,
        renderCriteriaSummary,
        renderFilterRules,
        collectDraft
      }
    };
  }

  function buildGuidConfig() {
    const draft = state.features.guid.configDraft;
    const panel = div('multi-config multi-config--guid');
    const flow = div('guid-config-flow');

    [
      ['1', '검토/추출'],
      ['2', '삭제 표시'],
      ['3', '삭제용 엑셀 적용']
    ].forEach(([stepNo, label]) => {
      const step = document.createElement('span');
      step.className = 'guid-config-step';
      step.innerHTML = `<strong>${stepNo}</strong><span>${label}</span>`;
      flow.append(step);
    });

    const makeGuidCard = (kicker, title, desc) => {
      const card = div('guid-config-card');
      const head = div('guid-config-card__head');
      const kickerEl = document.createElement('span');
      kickerEl.className = 'guid-config-card__kicker';
      kickerEl.textContent = kicker;
      const titleEl = document.createElement('strong');
      titleEl.className = 'guid-config-card__title';
      titleEl.textContent = title;
      const descEl = document.createElement('p');
      descEl.className = 'guid-config-card__desc';
      descEl.textContent = desc;
      const body = div('guid-config-card__body');
      head.append(kickerEl, titleEl, descEl);
      card.append(head, body);
      return { card, body };
    };

    const scopeCard = makeGuidCard(
      '영역 1',
      '검토 범위 선택',
      '프로젝트 파라미터를 기본으로 보고, 필요할 때만 로드된 패밀리와 Annotation 패밀리까지 확장합니다.'
    );
    const includeFamily = makeCheckboxField('로드된 패밀리 파라미터까지 함께 검토');
    includeFamily.field.classList.add('guid-config-toggle');
    includeFamily.input.checked = !!draft.includeFamily;
    const includeAnno = makeCheckboxField('Annotation 패밀리 포함');
    includeAnno.field.classList.add('guid-config-toggle');
    includeAnno.input.checked = !!draft.includeAnnotation;
    scopeCard.body.append(includeFamily.field, includeAnno.field);

    const excelRuleCard = makeGuidCard(
      '영역 2',
      '삭제용 엑셀 입력 규칙',
      '검토 후 내보낸 동일한 엑셀을 수정해서 다시 불러오면 됩니다. 사용자가 입력할 값은 삭제 표시뿐입니다.'
    );
    const ruleList = document.createElement('ul');
    ruleList.className = 'guid-config-list';
    [
      "삭제할 행은 '삭제여부' 열에 '삭제'라고 입력합니다.",
      '삭제하지 않을 행은 비워 두면 됩니다.',
      '마지막 숨김 키 행은 건드리지 않고 그대로 유지해야 합니다.'
    ].forEach((text) => {
      const li = document.createElement('li');
      li.textContent = text;
      ruleList.append(li);
    });
    const deleteBadge = div('guid-config-code');
    deleteBadge.innerHTML = '<span>입력 예시</span><strong>삭제여부 = 삭제</strong>';
    excelRuleCard.body.append(deleteBadge, ruleList);

    const processCard = makeGuidCard(
      '영역 3',
      '적용 프로세스',
      '삭제용 엑셀 불러오기를 실행하면 엑셀의 숨김 키로 대상을 찾고, 파일별로 순서대로 정리합니다.'
    );
    const processList = document.createElement('div');
    processList.className = 'guid-config-process';
    [
      ['1', '검토 결과를 삭제용 엑셀로 내보냅니다.'],
      ['2', "엑셀에서 삭제할 행의 '삭제여부'에 '삭제'를 입력합니다."],
      ['3', "'삭제용 엑셀 불러오기'로 같은 파일을 다시 선택해 정리를 적용합니다."],
      ['4', '센트럴 파일은 항상 로컬로, 모든 웍셋을 닫은 상태로 열어 처리합니다.']
    ].forEach(([num, text]) => {
      const row = div('guid-config-process__row');
      const numEl = document.createElement('strong');
      numEl.className = 'guid-config-process__num';
      numEl.textContent = num;
      const textEl = document.createElement('span');
      textEl.textContent = text;
      row.append(numEl, textEl);
      processList.append(row);
    });
    processCard.body.append(processList);

    const cleanupCard = makeGuidCard(
      '영역 4',
      '동기화 기록',
      '정리 후 중앙 모델을 동기화할 때만 코멘트를 남기도록 선택합니다. 문서 열기 방식은 고정 규칙으로 자동 처리됩니다.'
    );
    const useSyncComment = makeCheckboxField('동기화 시 코멘트 작성');
    useSyncComment.field.classList.add('guid-config-toggle');
    useSyncComment.input.checked = !!draft.useSyncComment;
    const syncComment = makeField('동기화 코멘트', 'guidSyncComment', '예) KKY Tools - 파라미터 GUID 정리');
    syncComment.field.classList.add('guid-config-comment');
    syncComment.input.value = draft.syncComment || 'KKY Tools - 파라미터 GUID 정리';
    const cleanupNote = div('guid-config-card__note');
    cleanupNote.textContent = '코멘트를 끄면 빈 코멘트로 동기화하고, 센트럴/워크셰어링 문서는 항상 동일한 열기 정책으로 처리합니다.';
    cleanupCard.body.append(useSyncComment.field, syncComment.field, cleanupNote);

    const updateAnnotationState = () => {
      const disabled = !includeFamily.input.checked;
      includeAnno.input.disabled = disabled;
      includeAnno.field.classList.toggle('is-disabled', disabled);
    };

    const updateSyncCommentState = () => {
      const disabled = !useSyncComment.input.checked;
      syncComment.input.disabled = disabled;
      syncComment.field.classList.toggle('is-disabled', disabled);
    };

    includeFamily.input.addEventListener('change', () => {
      state.features.guid.configDraft.includeFamily = !!includeFamily.input.checked;
      updateAnnotationState();
      markFeatureDirty('guid');
    });

    includeAnno.input.addEventListener('change', () => {
      state.features.guid.configDraft.includeAnnotation = !!includeAnno.input.checked;
      markFeatureDirty('guid');
    });

    useSyncComment.input.addEventListener('change', () => {
      state.features.guid.configDraft.useSyncComment = !!useSyncComment.input.checked;
      updateSyncCommentState();
      markFeatureDirty('guid');
    });

    syncComment.input.addEventListener('input', () => {
      state.features.guid.configDraft.syncComment = syncComment.input.value || '';
      markFeatureDirty('guid');
    });

    updateAnnotationState();
    updateSyncCommentState();
    state.features.guid.configDraft.closeAllWorksetsOnOpen = true;

    panel.append(flow, scopeCard.card, excelRuleCard.card, processCard.card, cleanupCard.card);
    return {
      panel,
      controls: {
        includeFamily,
        includeAnno,
        useSyncComment,
        syncComment,
        updateAnnotationState,
        updateSyncCommentState
      }
    };
  }

  function buildPointsConfig() {
    const panel = div('multi-config');
    const unit = makeSelectField('단위', [
      { value: 'ft', label: '십진 피트' },
      { value: 'm', label: '미터 (m)' },
      { value: 'mm', label: '밀리미터 (mm)' }
    ]);
    unit.select.value = state.features.points.configDraft.unit;
    unit.select.addEventListener('change', () => {
      state.features.points.configDraft.unit = unit.select.value;
      markFeatureDirty('points');
    });
    panel.append(unit.field);
    return { panel, controls: { unit } };
  }

  function buildWorksetAssignmentConfig() {
    const panel = div('multi-config');
    panel.style.display = 'grid';
    panel.style.gap = '12px';
    const flaggedWorkset = makeField('오류로 볼 웍셋 이름', 'worksetassignmentFlaggedWorkset', '예: Workset2');
    flaggedWorkset.input.value = state.features.worksetassignment.configDraft.flaggedWorksetName || '';
    flaggedWorkset.input.addEventListener('input', () => {
      state.features.worksetassignment.configDraft.flaggedWorksetName = String(flaggedWorkset.input.value || '').trim();
      markFeatureDirty('worksetassignment');
      refreshWorksetAssignmentFeatureSummary();
    });
    const note = div('feature-note');
    note.textContent = '비워두면 기본 웍셋(Workset1) 이외의 웍셋을 모두 오류로 봅니다. 이름을 입력하면 그 웍셋에 속한 객체만 오류로 기록하고, 없으면 기본 웍셋(Workset1) 정상 배정 요약 1행만 출력합니다.';
    panel.append(flaggedWorkset.field, note);
    return { panel, controls: { flaggedWorkset } };
  }

  function buildParameterDuplicationConfig() {
    const panel = div('multi-config');
    panel.style.display = 'grid';
    panel.style.gap = '12px';

    const presetCard = div('feature-row__summary');
    presetCard.style.display = 'grid';
    presetCard.style.gap = '10px';
    presetCard.style.padding = '12px';
    presetCard.style.borderRadius = '16px';
    presetCard.style.border = '1px solid var(--border-accent-soft)';
    presetCard.style.background = 'var(--surface-elevated)';

    const presetTitle = document.createElement('strong');
    presetTitle.textContent = '파라미터 세트';

    const presetSelect = makeSelectField('저장된 세트', [
      { value: '', label: '저장된 세트가 없습니다.' }
    ]);
    const presetName = makeField('세트 이름', 'parameterduplicationPresetName', '예: 기계 납품 기본 세트', 'text');
    const presetActions = div('feature-row__actions');
    presetActions.style.display = 'flex';
    presetActions.style.flexWrap = 'wrap';
    presetActions.style.gap = '8px';

    const savePresetBtn = document.createElement('button');
    savePresetBtn.type = 'button';
    savePresetBtn.className = 'btn btn--secondary';
    savePresetBtn.textContent = '세트 저장';

    const loadPresetBtn = document.createElement('button');
    loadPresetBtn.type = 'button';
    loadPresetBtn.className = 'btn btn--secondary';
    loadPresetBtn.textContent = '세트 불러오기';

    const deletePresetBtn = document.createElement('button');
    deletePresetBtn.type = 'button';
    deletePresetBtn.className = 'btn btn--ghost';
    deletePresetBtn.textContent = '세트 삭제';

    presetActions.append(savePresetBtn, loadPresetBtn, deletePresetBtn);

    const recentCard = div('feature-row__summary');
    recentCard.style.display = 'grid';
    recentCard.style.gap = '10px';
    recentCard.style.padding = '12px';
    recentCard.style.borderRadius = '16px';
    recentCard.style.border = '1px solid var(--border-accent-soft)';
    recentCard.style.background = 'var(--surface-elevated)';

    const recentTitle = document.createElement('strong');
    recentTitle.textContent = '최근 검토 항목';

    const recentSelect = makeSelectField('자동 저장 항목', [
      { value: '', label: '최근 항목이 없습니다.' }
    ]);
    const recentActions = div('feature-row__actions');
    recentActions.style.display = 'flex';
    recentActions.style.flexWrap = 'wrap';
    recentActions.style.gap = '8px';

    const loadRecentBtn = document.createElement('button');
    loadRecentBtn.type = 'button';
    loadRecentBtn.className = 'btn btn--secondary';
    loadRecentBtn.textContent = '최근 항목 불러오기';

    const clearRecentBtn = document.createElement('button');
    clearRecentBtn.type = 'button';
    clearRecentBtn.className = 'btn btn--ghost';
    clearRecentBtn.textContent = '기록 비우기';

    recentActions.append(loadRecentBtn, clearRecentBtn);
    const recentSummary = div('feature-note');

    const scope = makeSelectField('검토 범위', [
      { value: 'all', label: '추가된 전체 프로젝트 파라미터' },
      { value: 'selected', label: '지정 파라미터만' }
    ]);
    const names = makeField('검토 파라미터명', 'parameterduplicationNames', '예: 설명(Comments), 마크(Mark), 타입 설명(Type Comments)', 'textarea');
    names.input.rows = 5;

    const importCard = div('feature-row__summary');
    importCard.style.display = 'grid';
    importCard.style.gap = '10px';
    importCard.style.padding = '12px';
    importCard.style.borderRadius = '16px';
    importCard.style.border = '1px solid var(--border-accent-soft)';
    importCard.style.background = 'var(--surface-elevated)';

    const importHead = document.createElement('div');
    importHead.style.display = 'flex';
    importHead.style.justifyContent = 'space-between';
    importHead.style.alignItems = 'center';
    importHead.style.gap = '10px';
    importHead.style.flexWrap = 'wrap';
    const importTitle = document.createElement('strong');
    importTitle.textContent = '공유파라미터 TXT 가져오기';
    const importActions = div('feature-row__actions');
    importActions.style.display = 'flex';
    importActions.style.flexWrap = 'wrap';
    importActions.style.gap = '8px';

    const importBtn = document.createElement('button');
    importBtn.type = 'button';
    importBtn.className = 'btn btn--secondary';
    importBtn.textContent = 'TXT 추가';

    const clearNamesBtn = document.createElement('button');
    clearNamesBtn.type = 'button';
    clearNamesBtn.className = 'btn btn--ghost';
    clearNamesBtn.textContent = '목록 비우기';

    importActions.append(importBtn, clearNamesBtn);
    importHead.append(importTitle, importActions);

    const sharedPath = makeField('가져온 TXT', 'parameterduplicationSharedParamTxt', '공유파라미터 TXT를 선택해 주세요.', 'text');
    sharedPath.input.readOnly = true;
    sharedPath.input.placeholder = '공유파라미터 TXT를 선택해 주세요.';
    const sharedSummary = div('feature-note');

    const note = div('feature-note');
    note.textContent = '지정 파라미터만 선택하면 입력한 이름과 일치하는 프로젝트 파라미터만 검토합니다. 구분자는 쉼표, 세미콜론, 줄바꿈을 모두 지원하며, 공유파라미터 TXT에서 이름 목록을 바로 추가할 수 있습니다.';

    const draft = state.features.parameterduplication.configDraft || {};
    scope.select.value = normalizeParameterDuplicationScope(draft.scope);
    names.input.value = buildParameterDuplicationNamesText(draft.parameterNames);

    function renderRecentSummary() {
      const recents = loadParameterDuplicationRecent();
      const selectedKey = String(recentSelect.select.value || '').trim();
      const current = recents.find((item) => item.key === selectedKey) || recents[0];
      if (!current) {
        recentSummary.textContent = '최근 적용한 검토 대상이 자동으로 기록됩니다.';
        return;
      }
      const scopeLabel = current.scope === 'all'
        ? '전체 프로젝트 파라미터'
        : `지정 ${current.parameterNames.length}개 · ${buildParameterDuplicationNamePreview(current.parameterNames, 4)}`;
      const sourcePath = String(current.sharedParamSourcePath || '').trim();
      const sourceLabel = sourcePath ? getPathLeafLabel(sourcePath, sourcePath) : '';
      const timeLabel = formatParameterDuplicationRecentTimestamp(current.updatedAt);
      recentSummary.textContent = [scopeLabel, sourceLabel, timeLabel].filter(Boolean).join(' · ');
    }

    function renderRecentOptions(selectedKey = '') {
      const recents = loadParameterDuplicationRecent();
      const nextSelected = String(selectedKey || recentSelect.select.value || '').trim();
      recentSelect.select.innerHTML = '';

      const defaultOption = document.createElement('option');
      defaultOption.value = '';
      defaultOption.textContent = recents.length ? '최근 항목 선택' : '최근 항목이 없습니다.';
      recentSelect.select.append(defaultOption);

      recents.forEach((item) => {
        const option = document.createElement('option');
        option.value = item.key;
        option.textContent = buildParameterDuplicationRecentOptionLabel(item);
        recentSelect.select.append(option);
      });

      if (nextSelected && recents.some((item) => item.key === nextSelected)) {
        recentSelect.select.value = nextSelected;
      } else {
        recentSelect.select.value = '';
      }

      renderRecentSummary();
    }

    function renderPresetOptions(selectedName = '') {
      const presets = loadParameterDuplicationPresets();
      const nextSelected = String(selectedName || presetSelect.select.value || '').trim();
      presetSelect.select.innerHTML = '';

      const defaultOption = document.createElement('option');
      defaultOption.value = '';
      defaultOption.textContent = presets.length ? '저장된 세트 선택' : '저장된 세트가 없습니다.';
      presetSelect.select.append(defaultOption);

      presets.forEach((preset) => {
        const option = document.createElement('option');
        option.value = preset.name;
        option.textContent = preset.name;
        presetSelect.select.append(option);
      });

      if (nextSelected && presets.some((preset) => preset.name === nextSelected)) {
        presetSelect.select.value = nextSelected;
      } else {
        presetSelect.select.value = '';
      }
    }

    function renderImportedSummary() {
      const currentDraft = state.features.parameterduplication.configDraft || {};
      const sourcePath = String(currentDraft.sharedParamSourcePath || '').trim();
      const importedCount = Number(currentDraft.sharedParamImportCount) || 0;
      sharedPath.input.value = sourcePath;
      if (!sourcePath) {
        sharedSummary.textContent = 'TXT를 가져오면 지정 파라미터 목록에 이름을 자동으로 추가합니다.';
        return;
      }
      const sourceLabel = getPathLeafLabel(sourcePath, sourcePath);
      const countLabel = importedCount > 0 ? `${importedCount}개 로드` : '로드 완료';
      sharedSummary.textContent = `${sourceLabel} · ${countLabel}`;
    }

    const updateNamesState = () => {
      const isSelectedScope = scope.select.value === 'selected';
      names.input.disabled = !isSelectedScope;
      names.field.classList.toggle('is-disabled', !isSelectedScope);
    };

    function applySharedParams(payload) {
      if (payload?.cancelled) return;
      if (payload?.ok === false) {
        toast(payload?.message || '공유파라미터 TXT를 읽지 못했습니다. Revit에 연결된 공유파라미터 파일과 파일 권한을 확인해 주세요.', 'err');
        return;
      }

      const importedNames = parseParameterDuplicationNames(Array.isArray(payload?.parameterNames)
        ? payload.parameterNames.join('\n')
        : payload?.parameterNames);
      if (!importedNames.length) {
        toast(payload?.message || 'TXT에서 사용할 파라미터 이름을 찾지 못했습니다. 공유파라미터 TXT에 파라미터 정의가 있는지 확인해 주세요.', 'warn');
        return;
      }

      const currentDraft = state.features.parameterduplication.configDraft || {};
      const existingNames = parseParameterDuplicationNames(Array.isArray(currentDraft.parameterNames)
        ? currentDraft.parameterNames.join('\n')
        : currentDraft.parameterNames);
      const mergedNames = parseParameterDuplicationNames([...existingNames, ...importedNames].join('\n'));
      const addedCount = Math.max(0, mergedNames.length - existingNames.length);

      currentDraft.scope = 'selected';
      currentDraft.parameterNames = mergedNames;
      currentDraft.sharedParamSourcePath = String(payload?.path || '').trim();
      currentDraft.sharedParamImportCount = importedNames.length;
      scope.select.value = 'selected';
      names.input.value = buildParameterDuplicationNamesText(currentDraft.parameterNames);
      markFeatureDirty('parameterduplication');
      updateNamesState();
      renderImportedSummary();
      refreshParameterDuplicationFeatureSummary();

      if (addedCount > 0) {
        toast(`공유파라미터 TXT에서 ${addedCount}개를 추가했습니다.`, 'ok');
      } else {
        toast('가져온 파라미터가 이미 모두 목록에 있습니다.', 'warn');
      }
    }

    scope.select.addEventListener('change', () => {
      state.features.parameterduplication.configDraft.scope = normalizeParameterDuplicationScope(scope.select.value);
      markFeatureDirty('parameterduplication');
      updateNamesState();
      refreshParameterDuplicationFeatureSummary();
    });

    names.input.addEventListener('input', () => {
      state.features.parameterduplication.configDraft.parameterNames = parseParameterDuplicationNames(names.input.value);
      markFeatureDirty('parameterduplication');
      refreshParameterDuplicationFeatureSummary();
    });

    importBtn.addEventListener('click', () => {
      post('parameterduplication:pick-sharedparams', { source: 'multi' });
    });

    clearNamesBtn.addEventListener('click', () => {
      const currentDraft = state.features.parameterduplication.configDraft || {};
      currentDraft.parameterNames = [];
      currentDraft.sharedParamSourcePath = '';
      currentDraft.sharedParamImportCount = 0;
      names.input.value = '';
      markFeatureDirty('parameterduplication');
      renderImportedSummary();
      refreshParameterDuplicationFeatureSummary();
    });

    presetSelect.select.addEventListener('change', () => {
      presetName.input.value = presetSelect.select.value || '';
    });

    recentSelect.select.addEventListener('change', () => {
      renderRecentSummary();
    });

    savePresetBtn.addEventListener('click', () => {
      const presetLabel = String(presetName.input.value || presetSelect.select.value || '').trim();
      if (!presetLabel) {
        toast('저장할 세트 이름을 입력해 주세요.', 'warn');
        return;
      }
      const currentDraft = state.features.parameterduplication.configDraft || {};
      const scopeValue = normalizeParameterDuplicationScope(currentDraft.scope);
      const parameterNames = parseParameterDuplicationNames(Array.isArray(currentDraft.parameterNames)
        ? currentDraft.parameterNames.join('\n')
        : currentDraft.parameterNames);
      if (scopeValue === 'selected' && !parameterNames.length) {
        toast('지정 파라미터 세트는 대상 이름이 1개 이상 있어야 저장할 수 있습니다.', 'warn');
        return;
      }
      saveParameterDuplicationPreset(presetLabel, {
        scope: scopeValue,
        parameterNames,
        sharedParamSourcePath: String(currentDraft.sharedParamSourcePath || '').trim(),
        sharedParamImportCount: Number(currentDraft.sharedParamImportCount) || 0
      });
      renderPresetOptions(presetLabel);
      presetName.input.value = presetLabel;
      toast(`파라미터 세트를 저장했습니다: ${presetLabel}`, 'ok');
    });

    loadPresetBtn.addEventListener('click', () => {
      const presetLabel = String(presetSelect.select.value || presetName.input.value || '').trim();
      if (!presetLabel) {
        toast('불러올 세트를 선택해 주세요.', 'warn');
        return;
      }
      const loaded = applyParameterDuplicationPreset(state.features.parameterduplication.configDraft, presetLabel);
      if (!loaded) {
        toast('선택한 세트를 찾지 못했습니다.', 'err');
        renderPresetOptions();
        return;
      }
      presetName.input.value = presetLabel;
      markFeatureDirty('parameterduplication');
      syncControlsFromDraft('parameterduplication');
      refreshParameterDuplicationFeatureSummary();
      toast(`파라미터 세트를 불러왔습니다: ${presetLabel}`, 'ok');
    });

    loadRecentBtn.addEventListener('click', () => {
      const recentKey = String(recentSelect.select.value || '').trim();
      if (!recentKey) {
        toast('불러올 최근 항목을 선택해 주세요.', 'warn');
        return;
      }
      const loaded = applyParameterDuplicationRecent(state.features.parameterduplication.configDraft, recentKey);
      if (!loaded) {
        toast('선택한 최근 항목을 찾지 못했습니다.', 'err');
        renderRecentOptions();
        return;
      }
      markFeatureDirty('parameterduplication');
      syncControlsFromDraft('parameterduplication');
      refreshParameterDuplicationFeatureSummary();
      toast('최근 검토 항목을 불러왔습니다.', 'ok');
    });

    clearRecentBtn.addEventListener('click', () => {
      clearParameterDuplicationRecent();
      renderRecentOptions();
      toast('최근 검토 항목 기록을 비웠습니다.', 'ok');
    });

    deletePresetBtn.addEventListener('click', () => {
      const presetLabel = String(presetSelect.select.value || presetName.input.value || '').trim();
      if (!presetLabel) {
        toast('삭제할 세트를 선택해 주세요.', 'warn');
        return;
      }
      deleteParameterDuplicationPreset(presetLabel);
      renderPresetOptions();
      if (presetName.input.value === presetLabel) presetName.input.value = '';
      toast(`파라미터 세트를 삭제했습니다: ${presetLabel}`, 'ok');
    });

    updateNamesState();
    renderRecentOptions();
    renderPresetOptions();
    renderImportedSummary();

    presetCard.append(presetTitle, presetSelect.field, presetName.field, presetActions);
    recentCard.append(recentTitle, recentSelect.field, recentActions, recentSummary);
    importCard.append(importHead, sharedPath.field, sharedSummary);
    panel.append(presetCard, recentCard, scope.field, names.field, importCard, note);
    buildParameterDuplicationConfig.applySharedParams = applySharedParams;
    return {
      panel,
      controls: {
        scope,
        names,
        sharedPath,
        sharedSummary,
        presetSelect,
        presetName,
        recentSelect,
        updateNamesState,
        renderImportedSummary,
        renderPresetOptions,
        renderRecentOptions
      }
    };
  }

  function buildParameterDuplicationConfig() {
    const panel = div('multi-config');
    panel.style.display = 'grid';
    panel.style.gridTemplateColumns = 'repeat(auto-fit, minmax(320px, 1fr))';
    panel.style.gap = '12px';
    panel.style.alignItems = 'start';
    panel.style.width = '100%';
    panel.style.maxWidth = 'none';
    panel.style.minWidth = '0';
    panel.style.boxSizing = 'border-box';

    const recentCard = div('feature-row__summary');
    recentCard.style.display = 'grid';
    recentCard.style.gap = '10px';
    recentCard.style.padding = '12px';
    recentCard.style.borderRadius = '16px';
    recentCard.style.border = '1px solid var(--border-accent-soft)';
    recentCard.style.background = 'var(--surface-elevated)';
    recentCard.style.gridColumn = '1 / -1';

    const recentTitle = document.createElement('strong');
    recentTitle.textContent = '최근 검토 항목';

    const recentSelect = makeSelectField('자동 저장 항목', [
      { value: '', label: '최근 항목이 없습니다.' }
    ]);
    const recentActions = div('feature-row__actions');
    recentActions.style.display = 'flex';
    recentActions.style.flexWrap = 'wrap';
    recentActions.style.gap = '8px';

    const loadRecentBtn = document.createElement('button');
    loadRecentBtn.type = 'button';
    loadRecentBtn.className = 'btn btn--secondary';
    loadRecentBtn.textContent = '최근 항목 불러오기';

    const clearRecentBtn = document.createElement('button');
    clearRecentBtn.type = 'button';
    clearRecentBtn.className = 'btn btn--ghost';
    clearRecentBtn.textContent = '기록 비우기';

    recentActions.append(loadRecentBtn, clearRecentBtn);
    const recentSummary = div('feature-note');

    const overviewCard = div('feature-row__summary');
    overviewCard.style.display = 'grid';
    overviewCard.style.gap = '10px';
    overviewCard.style.padding = '12px';
    overviewCard.style.borderRadius = '16px';
    overviewCard.style.border = '1px solid var(--border-accent-soft)';
    overviewCard.style.background = 'var(--surface-elevated)';

    const overviewTitle = document.createElement('strong');
    overviewTitle.textContent = '검토 설정';

    const scope = makeSelectField('검토 범위', [
      { value: 'all', label: '추가된 전체 프로젝트 파라미터' },
      { value: 'selected', label: '지정 파라미터만' }
    ]);
    const sourceInfo = div('feature-note');

    const selectedWrap = div('feature-row__summary');
    selectedWrap.style.display = 'grid';
    selectedWrap.style.gridTemplateRows = 'auto minmax(0, 1fr)';
    selectedWrap.style.gap = '8px';
    selectedWrap.style.padding = '10px 12px';
    selectedWrap.style.borderRadius = '16px';
    selectedWrap.style.border = '1px solid var(--border-accent-soft)';
    selectedWrap.style.background = 'var(--surface-help)';
    selectedWrap.style.minHeight = '160px';

    const selectedHead = document.createElement('div');
    selectedHead.style.display = 'flex';
    selectedHead.style.justifyContent = 'space-between';
    selectedHead.style.alignItems = 'center';
    selectedHead.style.gap = '8px';

    const selectedTitle = document.createElement('strong');
    selectedTitle.textContent = '선택된 검토 파라미터';
    const selectedCount = document.createElement('span');
    selectedCount.className = 'chip chip--info';
    selectedHead.append(selectedTitle, selectedCount);

    const selectedChips = div('familylink-selected-chips');
    selectedChips.style.display = 'flex';
    selectedChips.style.flexWrap = 'wrap';
    selectedChips.style.gap = '8px';
    selectedChips.style.alignContent = 'flex-start';
    selectedChips.style.alignItems = 'flex-start';
    selectedChips.style.minHeight = '92px';
    selectedChips.style.maxHeight = '120px';
    selectedChips.style.overflow = 'auto';
    selectedWrap.append(selectedHead, selectedChips);

    const pickerCard = div('feature-row__summary');
    pickerCard.style.display = 'grid';
    pickerCard.style.gap = '8px';
    pickerCard.style.padding = '12px';
    pickerCard.style.borderRadius = '16px';
    pickerCard.style.border = '1px solid var(--border-accent-soft)';
    pickerCard.style.background = 'var(--surface-elevated)';

    const pickerHead = document.createElement('div');
    pickerHead.style.display = 'flex';
    pickerHead.style.justifyContent = 'space-between';
    pickerHead.style.alignItems = 'center';
    pickerHead.style.gap = '8px';
    const pickerTitle = document.createElement('strong');
    pickerTitle.textContent = '공유파라미터 검색';
    const pickerBadge = document.createElement('span');
    pickerBadge.className = 'chip chip--info';
    pickerHead.append(pickerTitle, pickerBadge);

    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = '공유 파라미터 이름 / 그룹 / GUID 검색';
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
    searchMeta.style.flexWrap = 'wrap';
    const searchInfo = document.createElement('span');
    searchInfo.style.color = 'var(--muted,#64748b)';
    searchInfo.style.fontSize = '12px';
    const refreshBtn = document.createElement('button');
    refreshBtn.type = 'button';
    refreshBtn.className = 'btn btn--secondary';
    refreshBtn.textContent = '목록 새로고침';
    searchMeta.append(searchInfo, refreshBtn);

    const listWrap = div('familylink-target-list');
    listWrap.style.height = '180px';
    listWrap.style.minHeight = '180px';
    listWrap.style.overflow = 'auto';
    listWrap.style.border = '1px solid var(--border-soft)';
    listWrap.style.borderRadius = '12px';
    listWrap.style.padding = '6px';
    listWrap.style.background = 'var(--surface-control)';

    const note = div('feature-note');
    note.textContent = '현재 Revit에 연결된 공유파라미터 파일에서만 목록을 불러옵니다. 오른쪽 목록을 클릭하면 지정 파라미터 검토로 자동 전환되고, 최근 적용한 조합은 최근 검토 항목에 자동으로 기록됩니다.';
    note.style.gridColumn = '1 / -1';

    const draft = state.features.parameterduplication.configDraft || {};
    draft.scope = normalizeParameterDuplicationScope(draft.scope);
    draft.parameterNames = parseParameterDuplicationNames(Array.isArray(draft.parameterNames)
      ? draft.parameterNames.join('\n')
      : draft.parameterNames);
    scope.select.value = draft.scope;

    function syncSharedParamSourcePath() {
      const currentDraft = state.features.parameterduplication.configDraft || {};
      const statusPath = String(state.sharedParamStatus?.path || '').trim();
      currentDraft.sharedParamSourcePath = statusPath || String(currentDraft.sharedParamSourcePath || '').trim();
      currentDraft.sharedParamImportCount = 0;
      return currentDraft.sharedParamSourcePath;
    }

    function getDraftParameterNames() {
      const currentDraft = state.features.parameterduplication.configDraft || {};
      currentDraft.parameterNames = parseParameterDuplicationNames(Array.isArray(currentDraft.parameterNames)
        ? currentDraft.parameterNames.join('\n')
        : currentDraft.parameterNames);
      return currentDraft.parameterNames;
    }

    function getAvailableSharedParamItems() {
      const seen = new Set();
      return (Array.isArray(state.sharedParamItems) ? state.sharedParamItems : [])
        .map((item) => ({
          name: String(item?.name || '').trim(),
          groupName: String(item?.groupName || '').trim(),
          guid: String(item?.guid || '').trim(),
          dataTypeToken: String(item?.dataTypeToken || item?.paramType || '').trim()
        }))
        .filter((item) => {
          if (!item.name) return false;
          const key = item.name.toLowerCase();
          if (seen.has(key)) return false;
          seen.add(key);
          return true;
        })
        .sort((a, b) => a.name.localeCompare(b.name, 'ko'));
    }

    function renderRecentSummary() {
      const recents = loadParameterDuplicationRecent();
      const selectedKey = String(recentSelect.select.value || '').trim();
      const current = recents.find((item) => item.key === selectedKey) || recents[0];
      if (!current) {
        recentSummary.textContent = '최근 적용한 검토 대상이 자동으로 기록됩니다.';
        return;
      }
      const scopeLabel = current.scope === 'all'
        ? '전체 프로젝트 파라미터'
        : `지정 ${current.parameterNames.length}개 · ${buildParameterDuplicationNamePreview(current.parameterNames, 4)}`;
      const sourcePath = String(current.sharedParamSourcePath || '').trim();
      const sourceLabel = sourcePath ? getPathLeafLabel(sourcePath, sourcePath) : '';
      const timeLabel = formatParameterDuplicationRecentTimestamp(current.updatedAt);
      recentSummary.textContent = [scopeLabel, sourceLabel, timeLabel].filter(Boolean).join(' · ');
    }

    function renderRecentOptions(selectedKey = '') {
      const recents = loadParameterDuplicationRecent();
      const nextSelected = String(selectedKey || recentSelect.select.value || '').trim();
      recentSelect.select.innerHTML = '';

      const defaultOption = document.createElement('option');
      defaultOption.value = '';
      defaultOption.textContent = recents.length ? '최근 항목 선택' : '최근 항목이 없습니다.';
      recentSelect.select.append(defaultOption);

      recents.forEach((item) => {
        const option = document.createElement('option');
        option.value = item.key;
        option.textContent = buildParameterDuplicationRecentOptionLabel(item);
        recentSelect.select.append(option);
      });

      if (nextSelected && recents.some((item) => item.key === nextSelected)) {
        recentSelect.select.value = nextSelected;
      } else {
        recentSelect.select.value = '';
      }

      renderRecentSummary();
    }

    function updateSelectionState() {
      const isSelectedScope = normalizeParameterDuplicationScope(scope.select.value) === 'selected';
      pickerBadge.textContent = isSelectedScope ? '지정 검토' : '전체 검토';
      selectedWrap.style.opacity = isSelectedScope ? '1' : '0.78';
      searchInput.placeholder = isSelectedScope
        ? '공유 파라미터 이름 / 그룹 / GUID 검색'
        : '목록에서 선택하면 지정 검토로 전환됩니다.';
    }

    function renderSharedParamStatus() {
      const currentDraft = state.features.parameterduplication.configDraft || {};
      const sharedPath = syncSharedParamSourcePath();
      const status = state.sharedParamStatus || {};
      const pathLabel = sharedPath ? getPathLeafLabel(sharedPath, sharedPath) : '미설정';
      const warning = String(status.warning || status.errorMessage || '').trim();
      if (warning) {
        sourceInfo.textContent = sharedPath
          ? `현재 공유파라미터 파일: ${pathLabel} · ${warning}`
          : warning;
      } else if (sharedPath) {
        sourceInfo.textContent = `현재 공유파라미터 파일: ${pathLabel}`;
      } else if (normalizeParameterDuplicationScope(currentDraft.scope) === 'selected') {
        sourceInfo.textContent = '현재 Revit에 연결된 공유파라미터 파일을 찾지 못했습니다.';
      } else {
        sourceInfo.textContent = '전체 프로젝트 파라미터 검토는 공유파라미터 선택 없이도 실행할 수 있습니다.';
      }
    }

    function renderParameterDuplicationSelected() {
      const currentDraft = state.features.parameterduplication.configDraft || {};
      const selected = getDraftParameterNames();
      const isSelectedScope = normalizeParameterDuplicationScope(currentDraft.scope) === 'selected';
      selectedCount.textContent = isSelectedScope ? `${selected.length}개 선택` : '전체 검토';
      selectedChips.innerHTML = '';

      if (!selected.length) {
        const emptyState = document.createElement('div');
        emptyState.style.width = '100%';
        emptyState.style.padding = '10px 12px';
        emptyState.style.borderRadius = '12px';
        emptyState.style.border = '1px dashed var(--border-accent-soft)';
        emptyState.style.background = 'var(--surface-empty)';
        emptyState.style.color = 'var(--muted,#64748b)';
        emptyState.style.display = 'grid';
        emptyState.style.placeItems = 'center';
        emptyState.style.minHeight = '92px';
        emptyState.textContent = isSelectedScope
          ? '오른쪽 목록에서 검토할 파라미터를 선택해 주세요.'
          : '현재는 추가된 전체 프로젝트 파라미터를 검토하도록 설정되어 있습니다.';
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
          const currentDraft = state.features.parameterduplication.configDraft || {};
          currentDraft.parameterNames = getDraftParameterNames().filter((item) => item.toLowerCase() !== String(name).toLowerCase());
          currentDraft.sharedParamSourcePath = syncSharedParamSourcePath();
          currentDraft.sharedParamImportCount = 0;
          markFeatureDirty('parameterduplication');
          updateSelectionState();
          renderParameterDuplicationSelected();
          renderSharedParamList();
          renderSharedParamStatus();
          refreshParameterDuplicationFeatureSummary();
        });
        selectedChips.append(chip);
      });
    }

    function commitParameterSelection(nextNames) {
      const currentDraft = state.features.parameterduplication.configDraft || {};
      currentDraft.scope = 'selected';
      currentDraft.parameterNames = parseParameterDuplicationNames(Array.isArray(nextNames) ? nextNames.join('\n') : nextNames);
      currentDraft.sharedParamSourcePath = syncSharedParamSourcePath();
      currentDraft.sharedParamImportCount = 0;
      scope.select.value = 'selected';
      markFeatureDirty('parameterduplication');
      updateSelectionState();
      renderSharedParamStatus();
      renderParameterDuplicationSelected();
      renderSharedParamList();
      refreshParameterDuplicationFeatureSummary();
    }

    function renderSharedParamList(payload) {
      const currentDraft = state.features.parameterduplication.configDraft || {};
      const isSelectedScope = normalizeParameterDuplicationScope(currentDraft.scope) === 'selected';
      const query = String(searchInput.value || '').trim().toLowerCase();
      const items = getAvailableSharedParamItems();
      const selected = new Set(getDraftParameterNames().map((item) => item.toLowerCase()));
      const filtered = items.filter((item) => {
        if (!query) return true;
        const hay = `${item.name || ''} ${item.groupName || ''} ${item.guid || ''} ${item.dataTypeToken || ''}`.toLowerCase();
        return hay.includes(query);
      });

      searchInfo.textContent = items.length
        ? `공유파라미터 ${filtered.length}/${items.length}개`
        : '현재 연결된 공유파라미터 목록이 없습니다.';
      listWrap.innerHTML = '';

      if (payload?.ok === false) {
        const err = div('familylink-target-empty');
        err.textContent = payload?.message || '공유파라미터 목록을 불러오지 못했습니다.';
        listWrap.append(err);
        renderParameterDuplicationSelected();
        return;
      }

      if (!items.length) {
        const empty = div('familylink-target-empty');
        empty.textContent = state.sharedParamStatus?.warning || '현재 Revit에 연결된 공유파라미터 파일에서 표시할 정의가 없습니다.';
        listWrap.append(empty);
        renderParameterDuplicationSelected();
        return;
      }

      if (!filtered.length) {
        const nohit = div('familylink-target-empty');
        nohit.textContent = '검색 결과가 없습니다.';
        listWrap.append(nohit);
        renderParameterDuplicationSelected();
        return;
      }

      filtered.forEach((item) => {
        const selectedNow = selected.has(String(item.name || '').toLowerCase());
        const row = document.createElement('button');
        row.type = 'button';
        row.className = 'familylink-target-row';
        row.style.display = 'grid';
        row.style.gridTemplateColumns = '1fr auto';
        row.style.alignItems = 'center';
        row.style.width = '100%';
        row.style.textAlign = 'left';
        row.style.border = '0';
        row.style.background = selectedNow ? 'var(--surface-note)' : 'transparent';
        row.style.borderRadius = '10px';
        row.style.padding = '8px 10px';
        row.style.cursor = 'pointer';

        const info = document.createElement('span');
        info.style.display = 'grid';
        info.style.gap = '2px';
        const name = document.createElement('strong');
        name.textContent = item.name || '';
        const sub = document.createElement('small');
        const parts = [];
        if (item.groupName) parts.push(item.groupName);
        if (item.dataTypeToken) parts.push(item.dataTypeToken);
        if (item.guid) parts.push(item.guid.slice(0, 8));
        sub.textContent = parts.join(' · ');
        info.append(name, sub);

        const action = document.createElement('span');
        action.className = selectedNow ? 'chip chip--ok' : 'chip chip--info';
        action.textContent = selectedNow ? '선택 완료' : (isSelectedScope ? '추가' : '선택');

        row.append(info, action);
        row.addEventListener('click', () => {
          const next = getDraftParameterNames();
          const key = String(item.name || '').toLowerCase();
          const index = next.findIndex((value) => String(value).toLowerCase() === key);
          if (index >= 0) next.splice(index, 1);
          else next.push(item.name);
          commitParameterSelection(next);
        });
        listWrap.append(row);
      });

      renderParameterDuplicationSelected();
    }

    scope.select.addEventListener('change', () => {
      const currentDraft = state.features.parameterduplication.configDraft || {};
      currentDraft.scope = normalizeParameterDuplicationScope(scope.select.value);
      currentDraft.sharedParamSourcePath = syncSharedParamSourcePath();
      currentDraft.sharedParamImportCount = 0;
      markFeatureDirty('parameterduplication');
      updateSelectionState();
      renderSharedParamStatus();
      renderParameterDuplicationSelected();
      renderSharedParamList();
      refreshParameterDuplicationFeatureSummary();
    });

    recentSelect.select.addEventListener('change', () => {
      renderRecentSummary();
    });

    loadRecentBtn.addEventListener('click', () => {
      const recentKey = String(recentSelect.select.value || '').trim();
      if (!recentKey) {
        toast('불러올 최근 항목을 선택해 주세요.', 'warn');
        return;
      }
      const loaded = applyParameterDuplicationRecent(state.features.parameterduplication.configDraft, recentKey);
      if (!loaded) {
        toast('선택한 최근 항목을 찾지 못했습니다.', 'err');
        renderRecentOptions();
        return;
      }
      markFeatureDirty('parameterduplication');
      syncControlsFromDraft('parameterduplication');
      refreshParameterDuplicationFeatureSummary();
      toast('최근 검토 항목을 불러왔습니다.', 'ok');
    });

    clearRecentBtn.addEventListener('click', () => {
      clearParameterDuplicationRecent();
      renderRecentOptions();
      toast('최근 검토 항목 기록을 비웠습니다.', 'ok');
    });

    searchInput.addEventListener('input', () => {
      renderSharedParamList();
    });

    refreshBtn.addEventListener('click', () => {
      requestSharedParamStatus('parameterduplication-refresh');
      requestSharedParamList('parameterduplication-refresh');
    });

    updateSelectionState();
    renderRecentOptions();
    renderSharedParamStatus();
    renderParameterDuplicationSelected();
    renderSharedParamList();

    recentCard.append(recentTitle, recentSelect.field, recentActions, recentSummary);
    overviewCard.append(overviewTitle, scope.field, sourceInfo, selectedWrap);
    pickerCard.append(pickerHead, searchInput, searchMeta, listWrap);
    panel.append(recentCard, overviewCard, pickerCard, note);

    buildParameterDuplicationConfig.renderSharedParamStatus = renderSharedParamStatus;
    buildParameterDuplicationConfig.renderSharedParamList = renderSharedParamList;
    return {
      panel,
      controls: {
        scope,
        selectedWrap,
        searchInput,
        listWrap,
        recentSelect,
        updateSelectionState,
        renderSharedParamStatus,
        renderSharedParamList,
        renderParameterDuplicationSelected,
        renderRecentOptions
      }
    };
  }

  function buildLinkWorksetConfig() {
    const panel = div('multi-config');
    const applyDefault = makeCheckboxField('실행 시 기본 웍셋(Workset1)만 열리도록 자동 적용');
    applyDefault.input.checked = state.features.linkworkset.configDraft.applyDefaultWorksetOnly !== false;
    applyDefault.input.addEventListener('change', () => {
      state.features.linkworkset.configDraft.applyDefaultWorksetOnly = applyDefault.input.checked;
      markFeatureDirty('linkworkset');
    });

    const useSyncComment = makeCheckboxField('동기화 시 코멘트 적용');
    useSyncComment.input.checked = !!state.features.linkworkset.configDraft.useSyncComment;
    const syncComment = makeField('동기화 코멘트', 'linkworksetSyncComment', '예) KKY Tools - 링크 기본 웍셋 적용');
    syncComment.input.value = state.features.linkworkset.configDraft.syncComment || 'KKY Tools - 링크 기본 웍셋 적용';
    syncComment.input.addEventListener('input', () => {
      state.features.linkworkset.configDraft.syncComment = syncComment.input.value;
      markFeatureDirty('linkworkset');
    });
    const updateSyncCommentState = () => {
      syncComment.input.disabled = !useSyncComment.input.checked;
    };
    useSyncComment.input.addEventListener('change', () => {
      state.features.linkworkset.configDraft.useSyncComment = useSyncComment.input.checked;
      markFeatureDirty('linkworkset');
      updateSyncCommentState();
    });
    updateSyncCommentState();

    const note = div('feature-note');
    note.textContent = '최상위 Revit 링크를 대상으로 현재 로드 상태와 열려 있는 사용자 웍셋 현황을 추출하고, 필요 시 기본 웍셋(Workset1)만 열리도록 재로드합니다. 코멘트 적용을 켜면 동기화 시 입력한 문구를 함께 기록합니다.';

    panel.append(applyDefault.field, useSyncComment.field, syncComment.field, note);
    return { panel, controls: { applyDefault, useSyncComment, syncComment } };
  }

  function buildParameterMissingConfig() {
    const panel = div('multi-config');
    panel.style.display = 'grid';
    panel.style.gridTemplateColumns = 'repeat(auto-fit, minmax(320px, 1fr))';
    panel.style.gap = '12px';
    panel.style.alignItems = 'start';
    panel.style.width = '100%';
    panel.style.maxWidth = 'none';
    panel.style.minWidth = '0';

    const sharedParamListId = `parameter-missing-shared-${Math.random().toString(36).slice(2, 8)}`;
    const sharedParamList = document.createElement('datalist');
    sharedParamList.id = sharedParamListId;

    function makeCard(titleText, options = {}) {
      const card = div('feature-row__summary');
      card.style.display = 'grid';
      card.style.gap = '10px';
      card.style.padding = '12px';
      card.style.borderRadius = '16px';
      card.style.border = '1px solid var(--border-accent-soft)';
      card.style.background = 'var(--surface-elevated)';
      if (options.fullWidth) card.style.gridColumn = '1 / -1';
      const title = document.createElement('strong');
      title.textContent = titleText;
      card.append(title);
      return { card, title };
    }

    function styleConditionInput(input) {
      input.style.width = '100%';
      input.style.padding = '8px 10px';
      input.style.borderRadius = '10px';
      input.style.border = '1px solid var(--border-soft)';
      input.style.background = 'var(--surface-control)';
      input.style.boxSizing = 'border-box';
    }

    function getDraft() {
      const draft = state.features.parametermissing.configDraft || {};
      const snapshot = createParameterMissingConfigSnapshot(draft);
      Object.assign(draft, snapshot);
      state.features.parametermissing.configDraft = draft;
      return draft;
    }

    function markParameterMissingDirty() {
      markFeatureDirty('parametermissing');
      refreshParameterMissingFeatureSummary();
      updateRunSummary();
    }

    function renderSharedParamNameOptions() {
      sharedParamList.innerHTML = '';
      getParameterMissingSharedParamItems().forEach((item) => {
        const option = document.createElement('option');
        option.value = item.name;
        option.label = [item.groupName, item.guid].filter(Boolean).join(' · ');
        sharedParamList.append(option);
      });
    }

    const picker = makeCard('누락 검토 파라미터');
    const pickerHead = document.createElement('div');
    pickerHead.style.display = 'flex';
    pickerHead.style.justifyContent = 'space-between';
    pickerHead.style.alignItems = 'center';
    pickerHead.style.gap = '8px';
    pickerHead.style.flexWrap = 'wrap';
    const pickerBadge = document.createElement('span');
    pickerBadge.className = 'chip chip--info';
    pickerHead.append(picker.title, pickerBadge);
    picker.card.innerHTML = '';
    picker.card.append(pickerHead);

    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = '공유 텍스트 파라미터 이름 / 그룹 / GUID 검색';
    styleConditionInput(searchInput);

    const pickerMeta = document.createElement('div');
    pickerMeta.style.display = 'flex';
    pickerMeta.style.justifyContent = 'space-between';
    pickerMeta.style.alignItems = 'center';
    pickerMeta.style.gap = '8px';
    pickerMeta.style.flexWrap = 'wrap';
    const pickerInfo = div('feature-note');
    pickerInfo.style.margin = '0';
    const refreshBtn = document.createElement('button');
    refreshBtn.type = 'button';
    refreshBtn.className = 'btn btn--secondary';
    refreshBtn.textContent = '목록 새로고침';
    pickerMeta.append(pickerInfo, refreshBtn);

    const listWrap = div('familylink-target-list');
    listWrap.style.height = '168px';
    listWrap.style.minHeight = '168px';
    listWrap.style.maxHeight = '168px';
    listWrap.style.overflow = 'auto';
    listWrap.style.border = '1px solid var(--border-soft)';
    listWrap.style.borderRadius = '12px';
    listWrap.style.padding = '6px';
    listWrap.style.background = 'var(--surface-control)';

    const selectedCard = makeCard('선택된 검토 파라미터');
    const selectedHead = document.createElement('div');
    selectedHead.style.display = 'flex';
    selectedHead.style.justifyContent = 'space-between';
    selectedHead.style.alignItems = 'center';
    selectedHead.style.gap = '8px';
    selectedHead.style.flexWrap = 'wrap';
    const selectedBadge = document.createElement('span');
    selectedBadge.className = 'chip chip--info';
    selectedHead.append(selectedCard.title, selectedBadge);
    selectedCard.card.innerHTML = '';
    selectedCard.card.append(selectedHead);
    const selectedChips = div('familylink-selected-chips');
    selectedChips.style.display = 'flex';
    selectedChips.style.flexWrap = 'wrap';
    selectedChips.style.gap = '8px';
    selectedChips.style.alignItems = 'flex-start';
    selectedChips.style.alignContent = 'flex-start';
    selectedChips.style.minHeight = '40px';
    selectedChips.style.maxHeight = '88px';
    selectedChips.style.overflow = 'auto';
    const selectedSummary = div('feature-note');
    selectedSummary.style.margin = '0';
    const recentSelect = makeSelectField('최근 설정', [
      { value: '', label: '최근 설정이 없습니다.' }
    ]);
    recentSelect.field.style.margin = '0';
    const recentActions = div('feature-row__actions');
    recentActions.style.display = 'flex';
    recentActions.style.flexWrap = 'wrap';
    recentActions.style.gap = '8px';
    const loadRecentBtn = document.createElement('button');
    loadRecentBtn.type = 'button';
    loadRecentBtn.className = 'btn btn--secondary';
    loadRecentBtn.textContent = '최근 불러오기';
    const clearRecentBtn = document.createElement('button');
    clearRecentBtn.type = 'button';
    clearRecentBtn.className = 'btn btn--ghost';
    clearRecentBtn.textContent = '기록 비우기';
    const saveFileBtn = document.createElement('button');
    saveFileBtn.type = 'button';
    saveFileBtn.className = 'btn btn--secondary';
    saveFileBtn.textContent = '파일 저장';
    const loadFileBtn = document.createElement('button');
    loadFileBtn.type = 'button';
    loadFileBtn.className = 'btn btn--secondary';
    loadFileBtn.textContent = '파일 불러오기';
    recentActions.append(loadRecentBtn, clearRecentBtn, saveFileBtn, loadFileBtn);
    const recentSummary = div('feature-note');
    recentSummary.style.margin = '0';

    const exceptionCard = makeCard('누락 예외 필터', { fullWidth: true });
    const exceptionDesc = div('feature-note');
    exceptionDesc.textContent = '선택한 파라미터별로 누락을 무시할 예외 규칙을 설정합니다. 예외가 없으면 비어 있는 값은 모두 오류로 봅니다.';
    exceptionDesc.style.margin = '0';
    const exceptionWrap = div('multi-config');
    exceptionWrap.style.display = 'grid';
    exceptionWrap.style.gridTemplateColumns = 'repeat(auto-fit, minmax(320px, 1fr))';
    exceptionWrap.style.gap = '12px';

    const targetFilterCard = makeCard('검토 대상 필터링', { fullWidth: true });
    const targetFilterDesc = div('feature-note');
    targetFilterDesc.textContent = '이 기능에서만 추가로 적용할 객체 필터입니다. 공통 검토대상 필터 이후에 한 번 더 범위를 좁히거나 제외할 수 있습니다.';
    targetFilterDesc.style.margin = '0';
    const targetFilterActions = div('feature-row__actions');
    targetFilterActions.style.display = 'flex';
    targetFilterActions.style.flexWrap = 'wrap';
    targetFilterActions.style.gap = '8px';
    const targetFilterToggleBtn = document.createElement('button');
    targetFilterToggleBtn.type = 'button';
    targetFilterToggleBtn.className = 'btn btn--secondary';
    targetFilterToggleBtn.textContent = '객체 필터 설정';
    const targetFilterClearBtn = document.createElement('button');
    targetFilterClearBtn.type = 'button';
    targetFilterClearBtn.className = 'btn btn--ghost';
    targetFilterClearBtn.textContent = '필터 비우기';
    targetFilterActions.append(targetFilterToggleBtn, targetFilterClearBtn);

    const targetFilterBody = div('multi-config');
    targetFilterBody.style.display = 'none';
    targetFilterBody.style.gridTemplateColumns = 'repeat(auto-fit, minmax(280px, 1fr))';
    targetFilterBody.style.gap = '12px';
    targetFilterBody.style.alignItems = 'start';
    const targetFilterMode = makeSelectField('검토 방식', PARAMETER_MISSING_TARGET_FILTER_MODES);
    const targetFilterCombination = makeSelectField('조건 결합', [
      { value: 'And', label: 'AND (모두 만족)' },
      { value: 'Or', label: 'OR (하나라도 만족)' }
    ]);
    const targetFilterRows = div('multi-config');
    targetFilterRows.style.gridColumn = '1 / -1';
    targetFilterRows.style.display = 'grid';
    targetFilterRows.style.gap = '8px';
    const targetFilterAddBtn = document.createElement('button');
    targetFilterAddBtn.type = 'button';
    targetFilterAddBtn.className = 'btn btn--secondary';
    targetFilterAddBtn.textContent = '필터 조건 추가';
    const targetFilterSummary = div('feature-note');
    targetFilterSummary.style.margin = '0';
    targetFilterBody.append(targetFilterMode.field, targetFilterCombination.field, targetFilterRows, targetFilterAddBtn);

    function getTargetFilterDraft() {
      const draft = getDraft();
      draft.targetFilter = createParameterMissingTargetFilterState(draft.targetFilter);
      return draft.targetFilter;
    }

    function renderTargetFilterSummary(filterOverride = null) {
      const filter = filterOverride || getTargetFilterDraft();
      const configuredCount = countParameterMissingConfiguredConditions(filter.conditions);
      if (!configuredCount) {
        targetFilterSummary.textContent = '전용 객체 필터가 없습니다. 공통 BQC 검토 대상 필터만 적용됩니다.';
        targetFilterClearBtn.disabled = true;
        return;
      }
      const modeLabel = resolveParameterMissingTargetFilterModeLabel(filter.mode);
      const comboLabel = normalizeParameterMissingCombinationMode(filter.combinationMode, 'And') === 'Or' ? 'OR' : 'AND';
      targetFilterSummary.textContent = `${modeLabel} · 조건 ${configuredCount}개 · ${comboLabel} · ${buildParameterMissingConditionSummary(filter.conditions, filter.combinationMode)}`;
      targetFilterClearBtn.disabled = false;
    }

    function renderTargetFilterRows() {
      const filter = getTargetFilterDraft();
      filter.conditions = normalizeParameterMissingConditionRows(filter.conditions, { keepEmpty: true });
      targetFilterMode.select.value = normalizeParameterMissingTargetFilterMode(filter.mode);
      targetFilterCombination.select.value = normalizeParameterMissingCombinationMode(filter.combinationMode, 'And');
      targetFilterRows.innerHTML = '';

      filter.conditions.forEach((row, conditionIndex) => {
        const line = div('multi-config');
        line.style.display = 'grid';
        line.style.gridTemplateColumns = 'minmax(180px, 1.4fr) minmax(150px, 0.9fr) minmax(160px, 1fr) auto';
        line.style.gap = '8px';
        line.style.alignItems = 'center';
        line.dataset.pmTargetFilterRow = 'true';

        const paramInput = createConditionInput(row.parameterName);
        const operatorSelect = createOperatorSelect(row.operatorName);
        const valueInput = createValueInput(row.value);
        paramInput.dataset.pmTargetFilterField = 'parameter';
        operatorSelect.dataset.pmTargetFilterField = 'operator';
        valueInput.dataset.pmTargetFilterField = 'value';
        syncValueInputState(valueInput, row);

        paramInput.addEventListener('input', () => {
          filter.conditions[conditionIndex].parameterName = paramInput.value;
          renderTargetFilterSummary();
          markParameterMissingDirty();
        });
        operatorSelect.addEventListener('change', () => {
          filter.conditions[conditionIndex].operatorName = operatorSelect.value;
          syncValueInputState(valueInput, filter.conditions[conditionIndex]);
          renderTargetFilterSummary();
          markParameterMissingDirty();
        });
        valueInput.addEventListener('input', () => {
          filter.conditions[conditionIndex].value = valueInput.value;
          renderTargetFilterSummary();
          markParameterMissingDirty();
        });

        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'btn btn--ghost';
        removeBtn.textContent = '삭제';
        removeBtn.addEventListener('click', () => {
          const currentFilter = syncTargetFilterDraftFromControls();
          if (currentFilter.conditions.length <= 1) {
            currentFilter.conditions = [createEmptyParameterMissingCondition()];
          } else {
            currentFilter.conditions.splice(conditionIndex, 1);
          }
          renderTargetFilterRows();
          markParameterMissingDirty();
        });

        line.append(paramInput, operatorSelect, valueInput, removeBtn);
        targetFilterRows.append(line);
      });

      renderTargetFilterSummary();
    }

    function readConditionFromElement(rowEl, fieldAttributeName) {
      const findField = (name) => rowEl.querySelector(`[${fieldAttributeName}="${name}"]`);
      const operatorName = String(findField('operator')?.value || 'Equals').trim() || 'Equals';
      return {
        enabled: true,
        parameterName: String(findField('parameter')?.value || '').trim(),
        operatorName,
        value: isParameterMissingConditionValueless(operatorName)
          ? ''
          : String(findField('value')?.value || '')
      };
    }

    function syncTargetFilterDraftFromControls() {
      const filter = getTargetFilterDraft();
      filter.mode = normalizeParameterMissingTargetFilterMode(targetFilterMode.select.value);
      filter.combinationMode = normalizeParameterMissingCombinationMode(targetFilterCombination.select.value, 'And');
      const rows = Array.from(targetFilterRows.querySelectorAll('[data-pm-target-filter-row="true"]'))
        .map((rowEl) => readConditionFromElement(rowEl, 'data-pm-target-filter-field'));
      filter.conditions = normalizeParameterMissingConditionRows(rows, { keepEmpty: true });
      renderTargetFilterSummary(filter);
      return filter;
    }

    targetFilterToggleBtn.addEventListener('click', () => {
      const willOpen = targetFilterBody.style.display === 'none';
      if (!willOpen) {
        syncTargetFilterDraftFromControls();
      }
      targetFilterBody.style.display = willOpen ? 'grid' : 'none';
      targetFilterToggleBtn.textContent = willOpen ? '객체 필터 접기' : '객체 필터 설정';
      if (willOpen) renderTargetFilterRows();
    });
    targetFilterMode.select.addEventListener('change', () => {
      getTargetFilterDraft().mode = normalizeParameterMissingTargetFilterMode(targetFilterMode.select.value);
      renderTargetFilterSummary();
      markParameterMissingDirty();
    });
    targetFilterCombination.select.addEventListener('change', () => {
      getTargetFilterDraft().combinationMode = normalizeParameterMissingCombinationMode(targetFilterCombination.select.value, 'And');
      renderTargetFilterSummary();
      markParameterMissingDirty();
    });
    targetFilterAddBtn.addEventListener('click', () => {
      const filter = syncTargetFilterDraftFromControls();
      filter.conditions = normalizeParameterMissingConditionRows(filter.conditions, { keepEmpty: true });
      filter.conditions.push(createEmptyParameterMissingCondition());
      renderTargetFilterRows();
      markParameterMissingDirty();
    });
    targetFilterClearBtn.addEventListener('click', () => {
      const filter = getTargetFilterDraft();
      filter.mode = 'include';
      filter.combinationMode = 'And';
      filter.conditions = [createEmptyParameterMissingCondition()];
      renderTargetFilterRows();
      renderTargetFilterSummary();
      markParameterMissingDirty();
    });

    function renderParameterMissingSelected() {
      const draft = getDraft();
      const names = Array.isArray(draft.parameterNames) ? draft.parameterNames : [];
      selectedBadge.textContent = `${names.length}개`;
      selectedChips.innerHTML = '';

      if (!names.length) {
        const empty = div('familylink-target-empty');
        empty.textContent = '공유 텍스트 파라미터를 선택하면 여기 표시됩니다.';
        selectedChips.append(empty);
      } else {
        names.forEach((name) => {
          const chip = document.createElement('button');
          chip.type = 'button';
          chip.className = 'chip chip--ok';
          chip.textContent = `${name} ×`;
          chip.addEventListener('click', () => {
            const current = getDraft();
            current.parameterNames = current.parameterNames.filter((item) => String(item || '').toLowerCase() !== String(name || '').toLowerCase());
            current.exceptionRules = normalizeParameterMissingExceptionRules(current.exceptionRules, current.parameterNames, { keepEmpty: true });
            renderParameterMissingSelected();
            renderSharedParamList();
            renderExceptionRules();
            markParameterMissingDirty();
          });
          selectedChips.append(chip);
        });
      }

      selectedSummary.textContent = names.length
        ? `검토 대상 ${names.length}개 · ${buildParameterMissingSelectionPreview(names, 6)}`
        : '아직 누락 검토 파라미터가 선택되지 않았습니다.';
    }

    function renderRecentSummary() {
      const recents = loadParameterMissingRecent();
      const selectedKey = String(recentSelect.select.value || '').trim();
      const current = recents.find((item) => item.key === selectedKey);
      if (!current) {
        recentSummary.textContent = '최근 적용한 설정은 자동으로 기억합니다. 필요할 때 여기서 다시 불러오거나 파일로 저장할 수 있습니다.';
        return;
      }
      const snapshot = createParameterMissingConfigSnapshot(current);
      const exceptionRuleCount = countParameterMissingConfiguredRules(snapshot);
      const timeLabel = formatParameterDuplicationRecentTimestamp(current.updatedAt);
      recentSummary.textContent = [
        timeLabel,
        `파라미터 ${snapshot.parameterNames.length}개`,
        exceptionRuleCount ? `예외 ${exceptionRuleCount}개` : '예외가 없습니다.',
        buildParameterMissingSelectionPreview(snapshot.parameterNames, 4)
      ].filter(Boolean).join(' · ');
    }

    function renderRecentOptions(selectedKey = '') {
      const recents = loadParameterMissingRecent();
      const nextSelected = String(selectedKey || recentSelect.select.value || '').trim();
      recentSelect.select.innerHTML = '';

      const defaultOption = document.createElement('option');
      defaultOption.value = '';
      defaultOption.textContent = recents.length ? '최근 설정 선택' : '최근 설정이 없습니다.';
      recentSelect.select.append(defaultOption);

      recents.forEach((item) => {
        const option = document.createElement('option');
        option.value = item.key;
        option.textContent = buildParameterMissingRecentOptionLabel(item);
        recentSelect.select.append(option);
      });

      if (nextSelected && recents.some((item) => item.key === nextSelected)) {
        recentSelect.select.value = nextSelected;
      } else {
        recentSelect.select.value = '';
      }

      renderRecentSummary();
    }

    function renderSharedParamStatus() {
      renderSharedParamNameOptions();
      const availableCount = getParameterMissingSharedParamItems().length;
      pickerBadge.textContent = `Text ${availableCount}개`;
      const status = state.sharedParamStatus || {};
      if (status.status === 'ok') {
        const sourceText = status.path ? ` · ${getPathLeafLabel(status.path, status.path)}` : '';
        pickerInfo.textContent = `연결된 공유파라미터 TXT 기준 Text 정의 ${availableCount}개${sourceText}`;
      } else {
        pickerInfo.textContent = status.warning || status.errorMessage || '공유파라미터 상태를 확인하는 중입니다.';
      }
    }

    function renderSharedParamList(payload) {
      const ok = payload?.ok !== false;
      const draft = getDraft();
      const selected = new Set((draft.parameterNames || []).map((name) => String(name || '').toLowerCase()));
      const query = String(searchInput.value || '').trim().toLowerCase();
      const items = ok ? getParameterMissingSharedParamItems() : [];
      const filtered = items.filter((item) => {
        if (!query) return true;
        return [item.name, item.groupName, item.guid]
          .filter(Boolean)
          .some((text) => String(text).toLowerCase().includes(query));
      });

      listWrap.innerHTML = '';
      if (!ok) {
        const error = div('familylink-target-empty');
        error.textContent = payload?.message || '공유 파라미터 목록을 불러오지 못했습니다.';
        listWrap.append(error);
        renderParameterMissingSelected();
        return;
      }

      if (!items.length) {
        const empty = div('familylink-target-empty');
        empty.textContent = state.sharedParamStatus?.warning || '현재 Revit에 연결된 공유파라미터 파일에서 Text 정의를 찾지 못했습니다.';
        listWrap.append(empty);
        renderParameterMissingSelected();
        return;
      }

      if (!filtered.length) {
        const nohit = div('familylink-target-empty');
        nohit.textContent = '검색 결과가 없습니다.';
        listWrap.append(nohit);
        renderParameterMissingSelected();
        return;
      }

      filtered.forEach((item) => {
        const selectedNow = selected.has(String(item.name || '').toLowerCase());
        const row = document.createElement('button');
        row.type = 'button';
        row.className = 'familylink-target-row';
        row.style.display = 'grid';
        row.style.gridTemplateColumns = '1fr auto';
        row.style.alignItems = 'center';
        row.style.width = '100%';
        row.style.textAlign = 'left';
        row.style.border = '0';
        row.style.background = selectedNow ? 'var(--surface-note)' : 'transparent';
        row.style.borderRadius = '10px';
        row.style.padding = '8px 10px';
        row.style.cursor = 'pointer';

        const info = document.createElement('span');
        info.style.display = 'grid';
        info.style.gap = '2px';
        const name = document.createElement('strong');
        name.textContent = item.name || '';
        const sub = document.createElement('small');
        const parts = [];
        if (item.groupName) parts.push(item.groupName);
        if (item.guid) parts.push(item.guid.slice(0, 8));
        sub.textContent = parts.join(' · ');
        info.append(name, sub);

        const action = document.createElement('span');
        action.className = selectedNow ? 'chip chip--ok' : 'chip chip--info';
        action.textContent = selectedNow ? '선택 완료' : '추가';

        row.append(info, action);
        row.addEventListener('click', () => {
          const current = getDraft();
          const next = parseParameterMissingNames(current.parameterNames.join('\n'));
          const key = String(item.name || '').toLowerCase();
          const index = next.findIndex((value) => String(value || '').toLowerCase() === key);
          if (index >= 0) next.splice(index, 1);
          else next.push(item.name);
          current.parameterNames = next;
          current.exceptionRules = normalizeParameterMissingExceptionRules(current.exceptionRules, current.parameterNames, { keepEmpty: true });
          renderParameterMissingSelected();
          renderSharedParamList();
          renderExceptionRules();
          markParameterMissingDirty();
        });
        listWrap.append(row);
      });

      renderParameterMissingSelected();
    }

    function createConditionInput(value) {
      const input = document.createElement('input');
      input.type = 'text';
      input.value = value || '';
      input.placeholder = '파라미터명';
      input.setAttribute('list', sharedParamListId);
      styleConditionInput(input);
      return input;
    }

    function createOperatorSelect(value) {
      const select = document.createElement('select');
      PARAMETER_MISSING_FILTER_OPERATORS.forEach((operatorName) => {
        const option = document.createElement('option');
        option.value = operatorName;
        option.textContent = operatorName;
        select.append(option);
      });
      select.value = PARAMETER_MISSING_FILTER_OPERATORS.includes(value) ? value : 'Equals';
      styleConditionInput(select);
      return select;
    }

    function createValueInput(value) {
      const input = document.createElement('input');
      input.type = 'text';
      input.value = value || '';
      input.placeholder = '값';
      styleConditionInput(input);
      return input;
    }

    function syncValueInputState(input, row) {
      const valueless = isParameterMissingConditionValueless(row?.operatorName);
      input.disabled = valueless;
      input.placeholder = valueless ? '값 입력 불필요' : '값';
      if (valueless) {
        input.value = '';
        row.value = '';
      }
    }

    function renderExceptionRules() {
      const draft = getDraft();
      draft.exceptionRules = normalizeParameterMissingExceptionRules(draft.exceptionRules, draft.parameterNames, { keepEmpty: true });
      exceptionWrap.innerHTML = '';

      if (!draft.parameterNames.length) {
        const empty = div('familylink-target-empty');
        empty.textContent = '먼저 누락 검토 파라미터를 선택해 주세요.';
        exceptionWrap.append(empty);
        return;
      }

      draft.exceptionRules.forEach((rule, ruleIndex) => {
        const ruleCard = div('feature-row__summary');
        ruleCard.style.display = 'grid';
        ruleCard.style.gap = '10px';
        ruleCard.style.padding = '12px';
        ruleCard.style.borderRadius = '14px';
        ruleCard.style.border = '1px solid var(--border-soft)';
        ruleCard.style.background = 'var(--surface-help)';
        ruleCard.dataset.pmExceptionRule = 'true';
        ruleCard.dataset.pmExceptionParameter = rule.parameterName || '';

        const head = document.createElement('strong');
        head.textContent = `${rule.parameterName} 누락 예외`;
        const combo = makeSelectField('예외 조건 결합', [
          { value: 'Or', label: 'OR (하나라도 만족하면 제외)' },
          { value: 'And', label: 'AND (모두 만족해야 제외)' }
        ]);
        combo.select.value = normalizeParameterMissingCombinationMode(rule.combinationMode, 'Or');
        combo.select.dataset.pmExceptionField = 'combination';
        combo.select.addEventListener('change', () => {
          syncExceptionRulesFromControls();
          if (draft.exceptionRules[ruleIndex]) {
            draft.exceptionRules[ruleIndex].combinationMode = combo.select.value;
          }
          renderExceptionRules();
          markParameterMissingDirty();
        });

        const table = document.createElement('table');
        table.className = 'data-table';
        table.style.width = '100%';
        const thead = document.createElement('thead');
        thead.innerHTML = '<tr><th>연관 파라미터</th><th>조건</th><th>값</th><th></th></tr>';
        const tbody = document.createElement('tbody');
        table.append(thead, tbody);

        const ruleRows = normalizeParameterMissingConditionRows(rule.conditions, { keepEmpty: true });
        draft.exceptionRules[ruleIndex].conditions = ruleRows;

        const ruleSummary = div('feature-note');
        ruleSummary.style.margin = '0';

        ruleRows.forEach((row, conditionIndex) => {
          const tr = document.createElement('tr');
          tr.dataset.pmExceptionCondition = 'true';
          const paramTd = document.createElement('td');
          const operatorTd = document.createElement('td');
          const valueTd = document.createElement('td');
          const actionTd = document.createElement('td');

          const paramInput = createConditionInput(row.parameterName);
          paramInput.dataset.pmExceptionField = 'parameter';
          paramInput.addEventListener('input', () => {
            draft.exceptionRules[ruleIndex].conditions[conditionIndex].parameterName = paramInput.value;
            renderRuleSummary();
            markParameterMissingDirty();
          });

          const operatorSelect = createOperatorSelect(row.operatorName);
          operatorSelect.dataset.pmExceptionField = 'operator';
          operatorSelect.addEventListener('change', () => {
            draft.exceptionRules[ruleIndex].conditions[conditionIndex].operatorName = operatorSelect.value;
            syncValueInputState(valueInput, draft.exceptionRules[ruleIndex].conditions[conditionIndex]);
            renderRuleSummary();
            markParameterMissingDirty();
          });

          const valueInput = createValueInput(row.value);
          valueInput.dataset.pmExceptionField = 'value';
          syncValueInputState(valueInput, row);
          valueInput.addEventListener('input', () => {
            draft.exceptionRules[ruleIndex].conditions[conditionIndex].value = valueInput.value;
            renderRuleSummary();
            markParameterMissingDirty();
          });

          const removeBtn = document.createElement('button');
          removeBtn.type = 'button';
          removeBtn.className = 'btn btn--ghost';
          removeBtn.textContent = '삭제';
          removeBtn.addEventListener('click', () => {
            syncExceptionRulesFromControls();
            const currentRows = draft.exceptionRules[ruleIndex]?.conditions || [];
            if (currentRows.length <= 1) {
              if (draft.exceptionRules[ruleIndex]) {
                draft.exceptionRules[ruleIndex].conditions = [createEmptyParameterMissingCondition()];
              }
            } else {
              currentRows.splice(conditionIndex, 1);
            }
            renderExceptionRules();
            markParameterMissingDirty();
          });

          paramTd.append(paramInput);
          operatorTd.append(operatorSelect);
          valueTd.append(valueInput);
          actionTd.append(removeBtn);
          tr.append(paramTd, operatorTd, valueTd, actionTd);
          tbody.append(tr);
        });

        const addRuleBtn = document.createElement('button');
        addRuleBtn.type = 'button';
        addRuleBtn.className = 'btn btn--secondary';
        addRuleBtn.textContent = '예외 조건 추가';
        addRuleBtn.addEventListener('click', () => {
          syncExceptionRulesFromControls();
          if (!draft.exceptionRules[ruleIndex]) return;
          draft.exceptionRules[ruleIndex].conditions = normalizeParameterMissingConditionRows(draft.exceptionRules[ruleIndex].conditions, { keepEmpty: true });
          draft.exceptionRules[ruleIndex].conditions.push(createEmptyParameterMissingCondition());
          renderExceptionRules();
          markParameterMissingDirty();
        });

        function renderRuleSummary() {
          const configuredCount = countParameterMissingConfiguredConditions(draft.exceptionRules[ruleIndex].conditions);
          if (!configuredCount) {
            ruleSummary.textContent = `${rule.parameterName} 값이 비어 있으면 그대로 오류로 판단합니다.`;
            return;
          }
          const comboLabel = normalizeParameterMissingCombinationMode(draft.exceptionRules[ruleIndex].combinationMode, 'Or') === 'Or' ? 'OR' : 'AND';
          ruleSummary.textContent = `예외 ${configuredCount}개 · ${comboLabel} · ${buildParameterMissingConditionSummary(draft.exceptionRules[ruleIndex].conditions, draft.exceptionRules[ruleIndex].combinationMode)}`;
        }

        renderRuleSummary();
        ruleCard.append(head, combo.field, table, addRuleBtn, ruleSummary);
        exceptionWrap.append(ruleCard);
      });
    }

    function syncExceptionRulesFromControls() {
      const draft = getDraft();
      const ruleElements = Array.from(exceptionWrap.querySelectorAll('[data-pm-exception-rule="true"]'));
      if (!ruleElements.length) {
        draft.exceptionRules = normalizeParameterMissingExceptionRules(draft.exceptionRules, draft.parameterNames, { keepEmpty: true });
        return draft.exceptionRules;
      }
      const rules = ruleElements.map((ruleEl) => {
        const parameterName = String(ruleEl.dataset.pmExceptionParameter || '').trim();
        const combinationMode = normalizeParameterMissingCombinationMode(
          ruleEl.querySelector('[data-pm-exception-field="combination"]')?.value,
          'Or'
        );
        const conditions = Array.from(ruleEl.querySelectorAll('[data-pm-exception-condition="true"]'))
          .map((rowEl) => readConditionFromElement(rowEl, 'data-pm-exception-field'));
        return {
          enabled: true,
          parameterName,
          combinationMode,
          conditions: normalizeParameterMissingConditionRows(conditions, { keepEmpty: true })
        };
      });
      draft.exceptionRules = normalizeParameterMissingExceptionRules(rules, draft.parameterNames, { keepEmpty: true });
      return draft.exceptionRules;
    }

    function syncParameterMissingDraftFromControls() {
      const draft = getDraft();
      syncTargetFilterDraftFromControls();
      syncExceptionRulesFromControls();
      Object.assign(draft, createParameterMissingConfigSnapshot(draft));
      return draft;
    }

    searchInput.addEventListener('input', () => {
      renderSharedParamList();
    });

    refreshBtn.addEventListener('click', () => {
      requestSharedParamStatus('parametermissing-settings');
      requestSharedParamList('parametermissing-settings');
    });

    recentSelect.select.addEventListener('change', () => {
      renderRecentSummary();
    });

    loadRecentBtn.addEventListener('click', () => {
      const recentKey = String(recentSelect.select.value || '').trim();
      if (!recentKey) {
        toast('불러올 최근 설정을 선택해 주세요.', 'warn');
        return;
      }
      const loaded = applyParameterMissingRecent(state.features.parametermissing.configDraft, recentKey);
      if (!loaded) {
        toast('선택한 최근 설정을 찾지 못했습니다.', 'err');
        renderRecentOptions();
        return;
      }
      markFeatureDirty('parametermissing');
      syncControlsFromDraft('parametermissing');
      refreshParameterMissingFeatureSummary();
      toast('최근 설정을 불러왔습니다.', 'ok');
    });

    clearRecentBtn.addEventListener('click', () => {
      clearParameterMissingRecent();
      renderRecentOptions();
      toast('최근 설정 기록을 비웠습니다.', 'ok');
    });

    saveFileBtn.addEventListener('click', () => {
      const currentDraft = syncParameterMissingDraftFromControls();
      if (!currentDraft.parameterNames.length) {
        toast('저장할 누락 검토 파라미터를 1개 이상 선택해 주세요.', 'warn');
        return;
      }
      if (hasIncompleteParameterMissingConfig(currentDraft)) {
        toast('객체 필터 또는 누락 예외 조건을 먼저 완성한 뒤 저장해 주세요.', 'warn');
        return;
      }
      const json = buildParameterMissingPresetJson(currentDraft);
      const fileName = buildParameterMissingPresetDefaultName(currentDraft);
      downloadParameterMissingPresetInBrowser(json, fileName);
      toast(`파라미터 누락 검토 설정 파일을 저장했습니다: ${fileName}`, 'ok');
    });

    loadFileBtn.addEventListener('click', () => {
      openParameterMissingPresetFileInBrowser((payload) => {
        let snapshot = null;
        try {
          snapshot = parseParameterMissingPresetSnapshot(payload?.json || '');
        } catch (error) {
          toast(error?.message || '파라미터 누락 검토 설정 파일을 읽지 못했습니다.', 'err');
          return;
        }
        Object.assign(state.features.parametermissing.configDraft, snapshot);
        markFeatureDirty('parametermissing');
        syncControlsFromDraft('parametermissing');
        refreshParameterMissingFeatureSummary();
        const fileLabel = String(payload?.fileName || '').trim();
        toast(fileLabel ? `파라미터 누락 검토 설정 파일을 불러왔습니다: ${fileLabel}` : '파라미터 누락 검토 설정 파일을 불러왔습니다.', 'ok');
      });
    });

    const pickerNote = div('feature-note');
    pickerNote.textContent = '검토할 파라미터를 고르면 바로 아래 누락 예외 필터에서 예외 조건을 설정할 수 있습니다.';
    pickerNote.style.margin = '0';

    picker.card.append(searchInput, pickerMeta, listWrap, pickerNote);
    selectedCard.card.append(selectedChips, selectedSummary, recentSelect.field, recentActions, recentSummary);
    const scopeNote = div('feature-note');
    scopeNote.textContent = '검토 대상은 BQC 공통 옵션의 검토 대상 필터 / 검토 제외 대상 필터를 그대로 사용합니다.';
    scopeNote.style.margin = '0';
    targetFilterCard.card.append(targetFilterDesc, targetFilterActions, targetFilterBody, targetFilterSummary);
    exceptionCard.card.append(exceptionDesc, scopeNote, exceptionWrap);

    panel.append(picker.card, selectedCard.card, targetFilterCard.card, exceptionCard.card, sharedParamList);

    renderSharedParamStatus();
    renderParameterMissingSelected();
    renderSharedParamList();
    renderTargetFilterRows();
    renderExceptionRules();
    renderRecentOptions();

    buildParameterMissingConfig.renderSharedParamList = renderSharedParamList;
    buildParameterMissingConfig.renderSharedParamStatus = renderSharedParamStatus;
    return {
      panel,
      controls: {
        searchInput,
        renderSharedParamList,
        renderSharedParamStatus,
        renderParameterMissingSelected,
        renderTargetFilterRows,
        renderExceptionRules,
        renderRecentOptions,
        syncDraftFromControls: syncParameterMissingDraftFromControls
      }
    };
  }

  function buildFamilyLinkConfig() {
    const panel = div('multi-config');
    const searchWrap = div('familylink-search');
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = '파라미터 검색';
    searchWrap.append(searchInput);

    const listWrap = div('familylink-target-list');
    const listEmpty = div('familylink-target-empty');
    listEmpty.textContent = '공유파라미터 목록이 없습니다.';
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
        error.textContent = payload?.message || '공유파라미터 목록을 불러오지 못했습니다.';
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
      selectedCount.textContent = `${selected.length}개 선택 완료`;
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
    const dropHint = div('rvt-drop-hint');
    dropHint.textContent = 'RVT 추가 버튼을 누르거나 탐색기에서 .rvt 파일을 이 목록으로 끌어다 놓으면 바로 등록됩니다.';
    section.append(dropHint);

    const body = div('rvt-panel-body');
    const tableWrap = div('rvt-table-wrap rvt-drop-zone');
    const { table, tbody, master } = createRvtTable();
    const summary = div('multi-rvt-summary');
    const footer = div('rvt-list-footer');
    const footerRight = div('rvt-list-footer__right');
    const expandBtn = cardBtn('리스트 크게 보기', () => openExpandedRvtModal(), 'btn--secondary');
    const empty = div('rvt-empty');
    const emptyTitle = document.createElement('strong');
    emptyTitle.textContent = '등록된 RVT가 없습니다.';
    const emptySub = document.createElement('span');
    emptySub.textContent = 'RVT 추가 또는 드래그 앤 드롭으로 파일을 등록해 주세요.';
    const emptyBtn = cardBtn('RVT 추가', handleAddRvt, 'btn--primary');
    empty.append(emptyTitle, emptySub, emptyBtn);

    tableWrap.append(table);
    footerRight.append(expandBtn);
    footer.append(summary, footerRight);
    body.append(tableWrap, empty, footer);
    section.append(body);
    attachRvtDropZone(tableWrap, {
      onDropPaths: handleDroppedRvts,
      onInvalid: () => toast('RVT 파일만 드래그해서 추가할 수 있습니다.', 'warn')
    });

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
          btnRemove.disabled = state.rvtChecked.size === 0;
          updateRunSummary();
          if (buildRvtExpandedModal.render) buildRvtExpandedModal.render();
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
    headText.textContent = mode === 'favorites'
      ? '왼쪽 즐겨찾기 영역에서 선택형 기능을 켠 뒤 검토를 시작해 주세요.'
      : '왼쪽 기능 영역을 눌러 선택한 뒤 검토를 시작해 주세요.';
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

    const multiBtn = cardBtn('RVT 여러 개 검토', openExpandedRvtModal, 'btn--secondary');
    multiBtn.classList.add('btn--multi', 'btn--action-main');
    actionRow.append(multiBtn);

    actions.append(actionRow);
    if (mode === 'bqc' || mode === 'favorites') {
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
    if (mode === 'favorites') {
      const presetExtras = div('multi-action-card__extras multi-action-card__extras--preset');
      const presetButtons = div('multi-action-card__buttons multi-action-card__buttons--preset');
      const savePresetBtn = cardBtn('프리셋 저장', requestFavoritePresetSave, 'btn--secondary');
      const loadPresetBtn = cardBtn('프리셋 불러오기', requestFavoritePresetLoad, 'btn--secondary');
      savePresetBtn.classList.add('btn--preset');
      loadPresetBtn.classList.add('btn--preset');
      presetButtons.append(savePresetBtn, loadPresetBtn);
      const presetNote = document.createElement('div');
      presetNote.className = 'multi-action-stack__note multi-action-stack__note--preset';
      presetExtras.append(presetButtons, presetNote);
      actions.append(presetExtras);
      state.ui.favoritePresetInfoEl = presetNote;
      updateFavoritePresetSummary();
    } else {
      state.ui.favoritePresetInfoEl = null;
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

  function collectCurrentFavoriteEntryIds() {
    const seen = new Set();
    return getFavoriteEntries()
      .map((entry) => String(entry?.id || '').trim())
      .filter(Boolean)
      .filter((key) => {
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
  }

  function collectCurrentFavoriteKeys() {
    return collectCurrentFavoriteEntryIds().filter((key) => FEATURE_KEYS.includes(key));
  }

  function buildFavoriteFeaturePresetConfig(key, config) {
    if (key === 'parametermissing') {
      return createParameterMissingSerializableConfig(config);
    }
    return deepCopy(config || {});
  }

  function buildFavoritePresetSnapshot() {
    const favoriteEntryIds = collectCurrentFavoriteEntryIds();
    const favoriteKeys = collectCurrentFavoriteKeys();
    const features = {};
    favoriteKeys.forEach((key) => {
      const feature = state.features[key];
      if (!feature) return;
      features[key] = {
        enabled: !!feature.enabled,
        config: buildFavoriteFeaturePresetConfig(key, feature.configCommitted || {})
      };
    });
    return {
      kind: FAVORITE_PRESET_KIND,
      version: FAVORITE_PRESET_VERSION,
      mode: 'favorites',
      createdAt: new Date().toISOString(),
      favoriteEntryIds,
      favoriteKeys,
      commonOptions: deepCopy(state.common.configCommitted || {}),
      features
    };
  }

  function sanitizeFavoritePresetFileLabel(value) {
    return String(value || '')
      .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function buildFavoritePresetDefaultName(snapshot) {
    const dateToken = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const featureMap = snapshot && typeof snapshot.features === 'object' ? snapshot.features : {};
    const enabledKeys = Object.keys(featureMap).filter((key) => featureMap[key]?.enabled);
    const favoriteEntryIds = Array.isArray(snapshot?.favoriteEntryIds) ? snapshot.favoriteEntryIds : [];
    let baseName = '즐겨찾기 프리셋';
    if (enabledKeys.length === 1) {
      baseName = FEATURE_META[enabledKeys[0]]?.label || enabledKeys[0];
    } else if (enabledKeys.length > 1) {
      baseName = `즐겨찾기 ${enabledKeys.length}개 기능`;
    } else if (favoriteEntryIds.length > 1) {
      baseName = `즐겨찾기 ${favoriteEntryIds.length}개`;
    }
    return `${dateToken}_${sanitizeFavoritePresetFileLabel(baseName) || '즐겨찾기 프리셋'}${FAVORITE_PRESET_EXTENSION}`;
  }

  function requestFavoritePresetSave() {
    if (state.ui.multiMode !== 'favorites') return;
    const snapshot = buildFavoritePresetSnapshot();
    const favoriteEntryIds = Array.isArray(snapshot?.favoriteEntryIds) ? snapshot.favoriteEntryIds : [];
    if (!favoriteEntryIds.length) {
      toast('현재 즐겨찾기 항목이 없어 프리셋을 저장할 수 없습니다.', 'warn');
      return;
    }
    const defaultName = buildFavoritePresetDefaultName(snapshot);
    const json = JSON.stringify(snapshot, null, 2);
    if (DEV) {
      downloadFavoritePresetInBrowser(json, defaultName);
      handleFavoritePresetSaved({ fileName: defaultName, path: defaultName });
      return;
    }
    post('favorites:preset-save', { json, defaultName });
  }

  function requestFavoritePresetLoad() {
    if (state.ui.multiMode !== 'favorites') return;
    if (DEV) {
      openFavoritePresetFileInBrowser();
      return;
    }
    post('favorites:preset-load', {});
  }

  function downloadFavoritePresetInBrowser(json, fileName) {
    const blob = new Blob([json], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.append(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  function openFavoritePresetFileInBrowser() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json,.kkyfav.json,application/json';
    input.addEventListener('change', () => {
      const file = input.files && input.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        handleFavoritePresetLoaded({
          fileName: file.name,
          path: file.name,
          json: typeof reader.result === 'string' ? reader.result : ''
        });
      };
      reader.onerror = () => {
        toast('다중 검토 즐겨찾기 프리셋 파일을 읽지 못했습니다.', 'err');
      };
      reader.readAsText(file, 'utf-8');
    }, { once: true });
    input.click();
  }

  function handleFavoritePresetSaved(payload) {
    const path = String(payload?.path || '').trim();
    const name = String(payload?.fileName || '').trim() || getPathLeafLabel(path, '즐겨찾기 프리셋');
    state.ui.favoritePresetSource = 'saved';
    state.ui.favoritePresetName = name;
    state.ui.favoritePresetPath = path || name;
    updateFavoritePresetSummary();
    toast(`다중 검토 즐겨찾기 프리셋을 저장했습니다. (${name})`, 'ok');
  }

  function handleFavoritePresetLoaded(payload) {
    const json = String(payload?.json || '').replace(/^\uFEFF/, '').trim();
    if (!json) {
      toast('불러온 다중 검토 즐겨찾기 프리셋 파일이 비어 있습니다.', 'err');
      return;
    }
    let snapshot = null;
    try {
      snapshot = parseFavoritePresetSnapshot(json);
    } catch (error) {
      toast(error?.message || '다중 검토 즐겨찾기 프리셋 형식을 읽지 못했습니다.', 'err');
      return;
    }
    applyFavoritePresetSnapshot(snapshot, {
      fileName: String(payload?.fileName || '').trim(),
      path: String(payload?.path || '').trim()
    });
  }

  function parseFavoritePresetSnapshot(json) {
    const snapshot = JSON.parse(json);
    if (!snapshot || typeof snapshot !== 'object' || Array.isArray(snapshot)) {
      throw new Error('다중 검토 즐겨찾기 프리셋 형식이 올바르지 않습니다.');
    }
    const kind = String(snapshot.kind || '').trim();
    if (kind && kind !== FAVORITE_PRESET_KIND) {
      throw new Error('KKY 다중 검토 즐겨찾기 프리셋 파일이 아닙니다.');
    }
    const version = Number(snapshot.version || FAVORITE_PRESET_VERSION);
    if (version > FAVORITE_PRESET_VERSION) {
      throw new Error('현재 버전보다 최신 다중 검토 즐겨찾기 프리셋이라 불러올 수 없습니다.');
    }
    if (!snapshot.features || typeof snapshot.features !== 'object' || Array.isArray(snapshot.features)) {
      throw new Error('다중 검토 즐겨찾기 프리셋에 기능 정보가 없습니다.');
    }
    return snapshot;
  }

  function getFavoriteEntryIdsFromSnapshot(snapshot) {
    const hasExplicitList = Array.isArray(snapshot?.favoriteEntryIds);
    const rawIds = hasExplicitList
      ? snapshot.favoriteEntryIds
      : Array.isArray(snapshot?.favoriteKeys)
        ? snapshot.favoriteKeys
        : [];
    const seen = new Set();
    const ids = rawIds
      .map((entryId) => String(entryId || '').trim())
      .filter((entryId) => getHubEntry(entryId))
      .filter((entryId) => {
        if (seen.has(entryId)) return false;
        seen.add(entryId);
        return true;
      });
    return { ids, hasExplicitList };
  }

  function applyFavoritePresetSnapshot(snapshot, meta = {}) {
    const featureMap = snapshot && typeof snapshot.features === 'object' ? snapshot.features : {};
    const snapshotKeys = Object.keys(featureMap).filter((key) => FEATURE_KEYS.includes(key));
    const favoritePreset = getFavoriteEntryIdsFromSnapshot(snapshot);
    const targetFavoriteIds = favoritePreset.hasExplicitList ? favoritePreset.ids : snapshotKeys;
    if (!targetFavoriteIds.length && !snapshotKeys.length) {
      toast('적용할 수 있는 즐겨찾기 정보가 없는 다중 검토 프리셋입니다.', 'warn');
      return false;
    }

    const restoredFavoriteIds = replaceHubFavorites(targetFavoriteIds, { toast: false });
    const currentFavoriteKeys = restoredFavoriteIds.filter((key) => FEATURE_KEYS.includes(key));

    applyFavoriteCommonPreset(snapshot.commonOptions);

    currentFavoriteKeys.forEach((key) => {
      if (Object.prototype.hasOwnProperty.call(featureMap, key)) {
        applyFavoriteFeaturePreset(key, featureMap[key]);
      } else {
        clearFavoriteFeatureSelection(key);
      }
    });

    const skippedKeys = snapshotKeys.filter((key) => !currentFavoriteKeys.includes(key));
    const loadedName = String(meta.fileName || '').trim() || getPathLeafLabel(meta.path, '즐겨찾기 프리셋');
    state.ui.favoritePresetSource = 'loaded';
    state.ui.favoritePresetName = loadedName;
    state.ui.favoritePresetPath = String(meta.path || '').trim() || loadedName;
    updateFavoritePresetSummary();
    updateRunSummary();
    updateRunActionLabel();

    const enabledCount = currentFavoriteKeys.filter((key) => state.features[key]?.enabled).length;
    const summaryParts = [`즐겨찾기 ${restoredFavoriteIds.length}개 복원`];
    if (currentFavoriteKeys.length) {
      summaryParts.push(`설정 적용 ${enabledCount}/${currentFavoriteKeys.length}개`);
    }
    const skipMessage = skippedKeys.length
      ? ` 일부 기능 ${skippedKeys.length}개는 현재 등록할 수 없어 제외했습니다.`
      : '';
    toast(`다중 검토 즐겨찾기 프리셋을 불러왔습니다. (${summaryParts.join(', ')})${skipMessage}`, skippedKeys.length ? 'warn' : 'ok');
    return true;
  }

  function mergeFavoritePresetConfig(baseConfig, incomingConfig) {
    const base = deepCopy(baseConfig || {});
    const source = incomingConfig && typeof incomingConfig === 'object' && !Array.isArray(incomingConfig)
      ? deepCopy(incomingConfig)
      : {};
    delete source.enabled;
    delete source.config;
    return Object.assign(base, source);
  }

  function buildFavoriteFeatureDraftFromPreset(key, baseConfig, incomingConfig) {
    if (key === 'parametermissing') {
      return createParameterMissingConfigSnapshot(incomingConfig || {});
    }
    return mergeFavoritePresetConfig(baseConfig, incomingConfig);
  }

  function applyFavoriteFeaturePreset(key, featureSnapshot) {
    const feature = state.features[key];
    if (!feature) return false;
    const raw = featureSnapshot && typeof featureSnapshot === 'object' ? featureSnapshot : {};
    const desiredEnabled = !!raw.enabled;
    const configSource = raw.config && typeof raw.config === 'object' ? raw.config : raw;
    const previousActiveKey = state.ui.activeFeatureKey;
    state.ui.activeFeatureKey = key;
    feature.enabled = desiredEnabled;
    feature.configDraft = buildFavoriteFeatureDraftFromPreset(key, feature.configCommitted, configSource);
    commitConfig(feature);
    state.ui.activeFeatureKey = previousActiveKey;
    if (!desiredEnabled) {
      feature.applied = false;
      feature.dirty = false;
      resetDraftFromCommitted(key);
    }
    markStale(key);
    syncFeatureRow(key);
    updateFeatureSummary(key);
    return true;
  }

  function applyFavoriteCommonPreset(commonOptions) {
    if (!commonOptions || typeof commonOptions !== 'object' || Array.isArray(commonOptions)) return false;
    const previousActiveKey = state.ui.activeFeatureKey;
    const previousCommitted = deepCopy(state.common.configCommitted || {});
    state.ui.activeFeatureKey = 'common';
    state.common.configDraft = mergeFavoritePresetConfig(state.common.configCommitted, commonOptions);
    commitConfig(state.common);
    state.ui.activeFeatureKey = previousActiveKey;
    updateCommonSummary();
    persistCommonOptions(state.common.configCommitted);
    emitCommonOptionsChanged();
    markCommonDependentFeaturesStale(previousCommitted, state.common.configCommitted || {});
    return true;
  }

  function updateFavoritePresetSummary() {
    const note = state.ui.favoritePresetInfoEl;
    if (!note) return;
    const name = String(state.ui.favoritePresetName || '').trim();
    const path = String(state.ui.favoritePresetPath || '').trim();
    const source = String(state.ui.favoritePresetSource || '').trim();
    if (!name) {
      note.textContent = '즐겨찾기 등록 목록과 기능별 설정을 파일로 저장하고, 나중에 그대로 다시 불러올 수 있습니다.';
      note.removeAttribute('title');
      return;
    }
    note.textContent = `${source === 'loaded' ? '불러온 프리셋' : '마지막 저장 프리셋'}: ${name}`;
    if (path) {
      note.setAttribute('aria-label', path);
    } else {
      note.removeAttribute('aria-label');
    }
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
    title.textContent = 'RVT 여러 개 검토';
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
    listSub.textContent = '목록에서 선택한 파일만 제거할 수 있고, 탐색기에서 .rvt 파일을 드래그해 바로 추가할 수 있습니다.';
    listHead.append(listTitle, listSub);
    const tableWrap = div('rvt-expand-table rvt-drop-zone');
    const { table, tbody, master } = createRvtTable();
    table.classList.add('rvt-expand-table__grid');
    const listEmpty = div('rvt-expand-list-empty');
    const listEmptyTitle = document.createElement('strong');
    listEmptyTitle.textContent = '등록된 RVT가 없습니다.';
    const listEmptyText = document.createElement('span');
    listEmptyText.textContent = 'RVT 추가 버튼을 누르거나 탐색기에서 .rvt 파일을 이 영역으로 끌어오면 바로 목록에 추가됩니다.';
    const listEmptyBtn = cardBtn('RVT 추가', handleAddRvt, 'btn--primary');
    listEmpty.append(listEmptyTitle, listEmptyText, listEmptyBtn);
    tableWrap.append(table, listEmpty);
    listSection.append(listHead, tableWrap);
    attachRvtDropZone(tableWrap, {
      onDropPaths: handleDroppedRvts,
      onInvalid: () => toast('RVT 파일만 드래그해서 추가할 수 있습니다.', 'warn')
    });

    const sideSection = div('rvt-expand-section rvt-expand-section--side');
    const featureSection = div('rvt-expand-subsection rvt-expand-subsection--features');
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
    buildReviewSummaryModal.body = body;
    buildReviewSummaryModal.titleEl = title;
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
    if (buildReviewSummaryModal.titleEl) {
      buildReviewSummaryModal.titleEl.textContent = '최근 결과 보기';
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
    if (buildReviewSummaryModal.body) {
      buildReviewSummaryModal.body.scrollTop = 0;
    }
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
    const normalized = normalizeMultiMode(mode);
    if (normalized === 'favorites') {
      return getFavoriteEntries()
        .map((entry) => entry.id)
        .filter((id) => FEATURE_KEYS.includes(id));
    }
    return normalized === 'utility' ? UTILITY_FEATURE_KEYS.slice() : BQC_FEATURE_KEYS.slice();
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
    const tableWrap = div('selected-table-wrap');
    tableWrap.append(table);
    section.append(head, tableWrap);

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
      const status = getSelectedFeatureStatus(key);
      const badge = document.createElement('span');
      badge.className = `chip status-chip ${status.className}`;
      badge.textContent = status.label;
      meta.append(name, badge);
      text.append(meta);
      item.append(text);
      list.append(item);
    });
  }

  function getReviewItemFile(item) {
    return String(item?.file || item?.File || '').trim();
  }

  function getReviewItemStatus(item) {
    return String(item?.status || item?.Status || 'pending');
  }

  function getReviewItemReason(item) {
    return String(item?.reason || item?.Reason || '');
  }

  function getRecentResultRows() {
    const payload = getCurrentModeReviewSummary();
    const modeKeys = getModeFeatureKeys();
    const summaryRows = [];
    modeKeys.forEach((key) => {
      const fileSummaries = Array.isArray(payload?.featureSummaries?.[key]?.fileSummaries)
        ? payload.featureSummaries[key].fileSummaries
        : [];
      fileSummaries.forEach((row) => summaryRows.push(row));
    });
    if (summaryRows.length) {
      const merged = new Map();
      summaryRows.forEach((row) => {
        const file = getReviewItemFile(row);
        if (!file) return;
        const current = merged.get(file) || {
          file,
          total: 0,
          issues: 0,
          near: 0,
          status: getReviewItemStatus(row),
          reason: getReviewItemReason(row)
        };
        current.total = Math.max(Number(current.total) || 0, Number(row?.total) || 0);
        current.issues += Number(row?.issues) || 0;
        current.near += Number(row?.near) || 0;
        current.status = current.status || getReviewItemStatus(row);
        if (!current.reason && getReviewItemReason(row)) current.reason = getReviewItemReason(row);
        merged.set(file, current);
      });
      return Array.from(merged.values()).map((row) => ({
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
      file: getReviewItemFile(item),
      total: '',
      issues: '',
      near: '',
      status: getReviewItemStatus(item),
      reason: getReviewItemReason(item)
    })).filter((item) => !!item.file);
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
        fileCell.setAttribute('aria-label', item.file || '-');
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
      const file = getReviewItemFile(item);
      if (!file) return;
      byFile.set(file, {
        file,
        status: getReviewItemStatus(item),
        total: '',
        issues: '',
        near: '',
        reason: getReviewItemReason(item)
      });
    });

    getModeFeatureKeys().forEach((key) => {
      const fileSummaries = Array.isArray(payload?.featureSummaries?.[key]?.fileSummaries)
        ? payload.featureSummaries[key].fileSummaries
        : [];
      fileSummaries.forEach((item) => {
        const file = getReviewItemFile(item);
        if (!file) return;
        const existing = byFile.get(file) || {
          file,
          status: getReviewItemStatus(item),
          total: '',
          issues: '',
          near: '',
          reason: ''
        };
        existing.total = existing.total === '' ? (Number(item?.total) || 0) : Math.max(Number(existing.total) || 0, Number(item?.total) || 0);
        existing.issues = (Number(existing.issues) || 0) + (Number(item?.issues) || 0);
        existing.near = (Number(existing.near) || 0) + (Number(item?.near) || 0);
        existing.status = existing.status || getReviewItemStatus(item);
        if (!existing.reason && getReviewItemReason(item)) existing.reason = getReviewItemReason(item);
        byFile.set(file, existing);
      });
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
    if (key === 'floorinfo') return '층정보 결과 엑셀';
    if (key === 'familysuitability') return '패밀리 타입 적합성 결과 엑셀';
    if (key === 'tapalign') return '탭/분기 축 결과 엑셀';
    if (key === 'dupclash') {
      return normalizeDupClashMode(state.features.dupclash?.configCommitted?.mode) === 'clash'
        ? '자체 간섭 결과 엑셀'
        : '중복 결과 엑셀';
    }
    if (key === 'worksetassignment') return '웍셋 배정 결과 엑셀';
    if (key === 'parameterduplication') return '파라미터 중복 결과 엑셀';
    if (key === 'parametermissing') return '파라미터 누락 결과 엑셀';
    if (key === 'guid') return 'GUID 결과 엑셀';
    if (key === 'familylink') return '패밀리 연동 결과 엑셀';
    if (key === 'points') return '기준점/북각 결과 엑셀';
    if (key === 'linkworkset') return '링크 웍셋 결과 엑셀';
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
    buildSettingsModal.titleEl = title;
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
    title.textContent = '공유파라미터 상태';
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
    if (key === 'floorinfo') refreshFloorInfoFeatureSummary();
    if (key === 'familysuitability') refreshFamilySuitabilityFeatureSummary();
    if (key === 'tapalign') refreshTapAlignFeatureSummary();
    if (key === 'dupclash') refreshDupClashFeatureSummary();
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
    resetExcelProgressState();
    state.ui.lastProgressPct = 0;
    updateRunActionLabel();
    const blockReason = getRunBlockingReason({ requireRvt: false });
    if (blockReason) {
      toast(blockReason, 'warn');
      return;
    }
    setBusyState(true);
    ProgressDialog.show('현재 파일 검토', '검토 구성을 확인하는 중입니다.');
    ProgressDialog.update(0, '검토 구성을 확인하는 중입니다.', '선택한 기능과 활성 문서 상태를 정리하는 중입니다.');
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
      if (btn.disabled && blockReason) btn.setAttribute('aria-label', blockReason);
      else btn.removeAttribute('aria-label');
    });
  }


  function onRun() {
    state.ui.runCompleted = false;
    resetExcelProgressState();
    state.ui.lastProgressPct = 0;
    updateRunActionLabel();
    const blockReason = getRunBlockingReason({ requireRvt: true });
    if (blockReason) {
      toast(blockReason, 'warn');
      return;
    }
    if (state.ui.isRvtListExpanded) {
      closeExpandedRvtModal();
    }
    setBusyState(true);
    ProgressDialog.show('납품 시 BQC 검토', '검토 구성을 확인하는 중입니다.');
    ProgressDialog.update(0, '검토 구성을 확인하는 중입니다.', '선택한 기능과 RVT 목록을 정리하는 중입니다.');
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
        return '공유파라미터 상태를 확인 중입니다.';
      }
      if (!silent && status.status && status.status !== 'ok') {
        requestSharedParamStatus('run');
      }
      return status.warning || status.errorMessage || '공유파라미터 상태를 확인해 주세요.';
    }
    if (state.features.familylink.enabled) {
      const targets = state.features.familylink.configCommitted.selectedTargets || [];
      if (!targets.length) return '패밀리 공유파라미터 연동 검토 대상이 없습니다.';
    }
    if (state.features.floorinfo.enabled) {
      const config = state.features.floorinfo.configCommitted || {};
      const rules = normalizeFloorInfoRules(config.levelRules);
      const selectedRules = rules.filter((rule) => rule.useAsBoundary !== false);
      const configuredCount = selectedRules.filter((rule) => String(rule.expectedValue || '').trim()).length;
      if (!String(config.parameterName || '').trim()) return '층정보 검토 파라미터명을 입력해 주세요.';
      if (!rules.length) return '활성 문서에서 레벨 목록을 불러와 층정보 검토 규칙을 설정해 주세요.';
      if (!selectedRules.length) return '층정보 영역을 구분할 레벨을 최소 1개 이상 선택해 주세요.';
      if (configuredCount < selectedRules.length) return '선택한 영역 기준 레벨마다 기대 층정보 값을 입력해 주세요.';
    }
    if (state.features.familysuitability.enabled) {
      const config = state.features.familysuitability.configCommitted || {};
      const filterRules = normalizeFamilySuitabilityFilterRules(config.filterRules, { keepEmpty: true });
      const invalidFilter = filterRules.find((rule) => {
        const keyword = String(rule.keyword || '').trim();
        const reviewText = String(rule.reviewText || '').trim();
        return (!!keyword || !!reviewText) && (!keyword || !reviewText);
      });
      if (!String(config.criteriaExcelPath || '').trim()) return '패밀리 타입 적합성 기준 엑셀 파일을 선택해 주세요.';
      if (!String(config.matchReviewText || '').trim()) return '기준 일치 검토 문구를 입력해 주세요.';
      if (!String(config.mismatchReviewText || '').trim()) return '기준 미일치 검토 문구를 입력해 주세요.';
      if (invalidFilter) return '패밀리 타입 적합성 필터 규칙은 키워드와 검토 문구를 함께 입력해 주세요.';
    }
    if (state.features.parameterduplication.enabled) {
      const config = state.features.parameterduplication.configCommitted || {};
      const scope = normalizeParameterDuplicationScope(config.scope);
      const names = Array.isArray(config.parameterNames) ? config.parameterNames : [];
      if (scope === 'selected' && !names.length) return '중복 검토 대상 파라미터명을 1개 이상 입력해 주세요.';
    }
    if (state.features.parametermissing.enabled) {
      const config = createParameterMissingConfigSnapshot(state.features.parametermissing.configCommitted || {});
      if (!config.parameterNames.length) return '파라미터 누락 검토 대상 파라미터를 1개 이상 선택해 주세요.';
      if (hasIncompleteParameterMissingConfig(config)) return '파라미터 누락 검토의 객체 필터 또는 누락 예외 조건을 모두 입력해 주세요.';
    }
    if (options.requireRvt && !state.rvtList.length) return 'RVT 파일을 추가해 주세요.';
    if (options.requireRvt && !getCheckedRvtPaths().length) return '검토할 RVT를 1개 이상 선택해 주세요.';
    return '';
  }

  function getCheckedRvtPaths() {
    return state.rvtList.filter((path) => state.rvtChecked.has(path));
  }

  function buildPayload() {
    return {
      rvtPaths: getCheckedRvtPaths(),
      commonOptions: state.common.configCommitted,
      features: {
        connector: buildCommittedFeature('connector'),
        floorinfo: buildCommittedFeature('floorinfo'),
        familysuitability: buildCommittedFeature('familysuitability'),
        tapalign: buildCommittedFeature('tapalign'),
        dupclash: buildCommittedFeature('dupclash'),
        worksetassignment: buildCommittedFeature('worksetassignment'),
        parameterduplication: buildCommittedFeature('parameterduplication'),
        parametermissing: buildCommittedFeature('parametermissing'),
        guid: buildCommittedFeature('guid'),
        familylink: buildCommittedFeature('familylink'),
        points: buildCommittedFeature('points'),
        linkworkset: buildCommittedFeature('linkworkset')
      }
    };
  }

  function onExport(key) {
    chooseExcelMode((choice) => {
      const mode = typeof choice === 'object' ? choice?.mode : choice;
      const splitByFile = !!(choice && typeof choice === 'object' && choice.splitByFile);
      resetExcelProgressState();
      setBusyState(true);
      ProgressDialog.show('엑셀 내보내기', '엑셀 내보내기를 준비하는 중입니다.');
      ProgressDialog.update(
        0,
        splitByFile ? '엑셀 저장 폴더를 준비하는 중입니다.' : '엑셀 내보내기를 준비하는 중입니다.',
        splitByFile
          ? '저장 폴더를 선택한 뒤 파일별 개별 엑셀을 생성합니다.'
          : '결과 시트와 저장 옵션을 정리하는 중입니다.'
      );
      post('hub:multi-export', {
        key,
        excelMode: mode || 'fast',
        locale: getLastExcelExportLocale(),
        splitByFile
      });
    }, { allowSplit: true });
  }

  function updateOpenMultiBtnState() {
    const blockReason = getOpenMultiBlockingReason();
    (state.ui.openMultiButtons || []).forEach((btn) => {
      if (!btn) return;
      btn.disabled = !!blockReason;
      if (blockReason) btn.setAttribute('aria-label', blockReason);
      else btn.removeAttribute('aria-label');
    });
  }

  function getOpenMultiBlockingReason() {
    if (state.busy) return '작업을 진행하는 중입니다.';
    const enabledCount = FEATURE_KEYS.filter((k) => state.features[k].enabled).length;
    if (!enabledCount) return '선택된 기능이 있을 때만 RVT 여러 개 검토를 열 수 있습니다.';
    return '';
  }

  function updateMultiRunBtnState() {
    const blockReason = getRunBlockingReason({ requireRvt: true, silent: true });
    (state.ui.modalRunButtons || []).forEach((btn) => {
      if (!btn) return;
      btn.disabled = state.busy || !!blockReason;
      if (btn.disabled && blockReason) btn.setAttribute('aria-label', blockReason);
      else btn.removeAttribute('aria-label');
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
      if (rows.length) state.ui.bqcRecentOpenBtn.removeAttribute('aria-label');
      else state.ui.bqcRecentOpenBtn.setAttribute('aria-label', '현재 화면의 최근 결과를 확인합니다.');
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
    if (buildRvtExpandedModal.render) buildRvtExpandedModal.render();
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
      } else if (key === 'floorinfo') {
        buildSettingsModal.modal.style.width = 'min(1280px, 95vw)';
        buildSettingsModal.modal.style.maxWidth = '1280px';
        buildSettingsModal.body.style.display = 'grid';
        buildSettingsModal.body.style.gridTemplateColumns = 'minmax(0, 2.05fr) minmax(300px, 0.82fr)';
        buildSettingsModal.body.style.alignItems = 'start';
        buildSettingsModal.body.style.gap = '12px';
        buildSettingsModal.body.style.columnGap = '12px';
        buildSettingsModal.form.style.display = 'grid';
        buildSettingsModal.form.style.gridTemplateColumns = 'minmax(0, 1fr)';
        buildSettingsModal.form.style.gridAutoFlow = 'row';
        buildSettingsModal.form.style.gap = '12px';
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
        buildSettingsModal.help.style.maxWidth = '380px';
        buildSettingsModal.help.style.margin = '0';
        buildSettingsModal.help.style.justifySelf = 'stretch';
        buildSettingsModal.help.style.alignSelf = 'start';
      } else if (key === 'familysuitability') {
        buildSettingsModal.modal.style.width = 'min(1380px, 96vw)';
        buildSettingsModal.modal.style.maxWidth = '1380px';
        buildSettingsModal.body.style.display = 'grid';
        buildSettingsModal.body.style.gridTemplateColumns = 'minmax(0, 2.16fr) minmax(280px, 0.74fr)';
        buildSettingsModal.body.style.alignItems = 'start';
        buildSettingsModal.body.style.gap = '12px';
        buildSettingsModal.body.style.columnGap = '12px';
        buildSettingsModal.form.style.display = 'grid';
        buildSettingsModal.form.style.gridTemplateColumns = 'minmax(0, 1fr)';
        buildSettingsModal.form.style.gridAutoFlow = 'row';
        buildSettingsModal.form.style.gap = '12px';
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
        buildSettingsModal.help.style.maxWidth = '340px';
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
      } else if (key === 'guid') {
        buildSettingsModal.modal.style.width = 'min(1080px, 94vw)';
        buildSettingsModal.modal.style.maxWidth = '1080px';
        buildSettingsModal.body.style.display = 'grid';
        buildSettingsModal.body.style.gridTemplateColumns = 'minmax(0, 1.46fr) minmax(300px, 0.9fr)';
        buildSettingsModal.body.style.alignItems = 'start';
        buildSettingsModal.body.style.gap = '12px';
        buildSettingsModal.body.style.columnGap = '12px';
        buildSettingsModal.form.style.display = 'grid';
        buildSettingsModal.form.style.gridTemplateColumns = 'minmax(0, 1fr)';
        buildSettingsModal.form.style.gridAutoFlow = 'row';
        buildSettingsModal.form.style.gap = '12px';
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
        buildSettingsModal.help.style.maxWidth = '360px';
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
    buildSettingsModal.titleEl.textContent = `${title || ''} 설정`;
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
    if (key === 'floorinfo') requestFloorInfoConfig('settings-open');
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
      } else if (key === 'floorinfo' || key === 'guid' || key === 'tapalign') {
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
      } else if (key === 'dupclash') {
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
      const previousCommitted = deepCopy(state.common.configCommitted || {});
      commitConfig(state.common);
      updateCommonSummary();
      persistCommonOptions(state.common.configCommitted);
      emitCommonOptionsChanged();
      markCommonDependentFeaturesStale(previousCommitted, state.common.configCommitted || {});
    } else {
      if (key === 'floorinfo' && state.ui.controls.floorinfo?.collectDraft) {
        state.ui.controls.floorinfo.collectDraft();
      }
      if (key === 'familysuitability' && state.ui.controls.familysuitability?.collectDraft) {
        state.ui.controls.familysuitability.collectDraft();
      }
      commitConfig(state.features[key]);
      if (key === 'familysuitability') {
        persistFamilySuitabilityConfig(state.features.familysuitability.configCommitted);
      }
      if (key === 'tapalign') {
        persistTapAlignConfig(state.features.tapalign.configCommitted);
      }
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

  function resetExcelProgressState() {
    state.ui.lastExcelPct = 0;
    state.ui.excelBatchStartPct = null;
    state.ui.excelBatchEndPct = null;
  }

  function updateExcelBatchWindow(phase, payload) {
    const message = String(payload?.message || '');
    const total = Number(payload?.total) || 0;
    const current = Math.max(0, Number(payload?.current) || 0);
    const explicitStart = Number(payload?.batchStartPercent);
    const explicitEnd = Number(payload?.batchEndPercent);
    const hasExplicitWindow = phase === 'EXCEL_INIT'
      && Number.isFinite(explicitStart)
      && Number.isFinite(explicitEnd)
      && explicitEnd > explicitStart;
    const isSplitBatchInit = phase === 'EXCEL_INIT'
      && message.includes('파일별 엑셀 저장')
      && total > 0;

    if (hasExplicitWindow || isSplitBatchInit) {
      const start = hasExplicitWindow
        ? Math.max(0, Math.min(100, explicitStart * 100))
        : Math.max(0, Math.min(100, (current / total) * 100));
      const end = hasExplicitWindow
        ? Math.max(start, Math.min(100, explicitEnd * 100))
        : Math.max(start, Math.min(100, ((current + 1) / total) * 100));
      state.ui.excelBatchStartPct = start;
      state.ui.excelBatchEndPct = end;
      state.ui.lastExcelPct = start;
      return;
    }

    if (phase === 'EXCEL_INIT') {
      state.ui.excelBatchStartPct = null;
      state.ui.excelBatchEndPct = null;
    }
  }

  function getExcelBatchWindow() {
    const start = Number(state.ui.excelBatchStartPct);
    const end = Number(state.ui.excelBatchEndPct);
    if (!Number.isFinite(start) || !Number.isFinite(end) || end <= start) return null;
    return { start, end };
  }

  function isExcelProgressPhase(phase) {
    if (!String(phase || '').trim()) return false;
    const normalized = normalizeExcelPhase(phase);
    return normalized === 'EXCEL_INIT'
      || normalized === 'EXCEL_WRITE'
      || normalized === 'EXCEL_SAVE'
      || normalized === 'AUTOFIT'
      || normalized === 'DONE'
      || normalized === 'ERROR';
  }

  function normalizeExcelPhase(phase) {
    return String(phase || '').trim().toUpperCase();
  }

  function clamp01(value) {
    const n = Number(value);
    return Number.isFinite(n) ? Math.max(0, Math.min(1, n)) : 0;
  }

  function computeExcelPercent(phase, current, total, phaseProgress, percentOverride) {
    const normalized = normalizeExcelPhase(phase);
    const batch = getExcelBatchWindow();
    if (normalized === 'DONE') {
      const donePct = batch ? batch.end : 100;
      state.ui.lastExcelPct = donePct;
      return donePct;
    }
    if (normalized === 'ERROR') return state.ui.lastExcelPct;

    if (typeof percentOverride === 'number' && Number.isFinite(percentOverride) && percentOverride > 0 && percentOverride <= 1) {
      if (normalized === 'EXCEL_INIT' && batch) {
        state.ui.lastExcelPct = batch.start;
        return batch.start;
      }
      state.ui.lastExcelPct = Math.max(state.ui.lastExcelPct, percentOverride * 100);
      return state.ui.lastExcelPct;
    }

    const completed = EXCEL_PHASE_ORDER.reduce((acc, key) => {
      if (key === normalized) return acc;
      return acc + (EXCEL_PHASE_WEIGHT[key] || 0);
    }, 0);
    const weight = EXCEL_PHASE_WEIGHT[normalized] || 0;
    const ratio = total > 0 ? Math.max(0, Math.min(1, current / total)) : 0;
    const staged = Math.max(ratio, clamp01(phaseProgress));
    let pct = (completed + weight * staged) * 100;
    if (batch) {
      pct = batch.start + (Math.max(0, Math.min(100, pct)) / 100) * (batch.end - batch.start);
    }
    state.ui.lastExcelPct = Math.max(state.ui.lastExcelPct, Math.min(100, pct));
    return state.ui.lastExcelPct;
  }

  function buildExcelSubtitle(phase, current, total) {
    switch (normalizeExcelPhase(phase)) {
      case 'EXCEL_INIT': return '다중 검토 결과 엑셀 워크북을 준비하는 중입니다.';
      case 'EXCEL_WRITE': return `다중 검토 결과 엑셀 데이터를 작성하는 중입니다. (${current}/${Math.max(total, current || 1)})`;
      case 'EXCEL_SAVE': return '다중 검토 결과 엑셀을 저장하는 중입니다.';
      case 'AUTOFIT': return '다중 검토 결과 엑셀 열 너비를 자동으로 맞추는 중입니다.';
      case 'DONE': return '다중 검토 결과 엑셀 내보내기 완료';
      case 'ERROR': return '다중 검토 결과 엑셀 내보내기 오류';
      default: return '다중 검토 결과 엑셀 내보내기를 진행하는 중입니다.';
    }
  }

  function formatExcelDetail(phase, message, detail) {
    if (detail) return detail;
    if (message) return message;
    return normalizeExcelPhase(phase) === 'DONE' ? '다중 검토 결과 엑셀 내보내기가 완료되었습니다.' : '';
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
    resetExcelProgressState();
    state.ui.lastProgressPct = 0;
    state.ui.reviewSummaryData = null;
    state.ui.lastRunAt = null;
    state.ui.reviewSummaryByMode = { bqc: null, utility: null, favorites: null };
    state.ui.lastRunAtByMode = { bqc: null, utility: null, favorites: null };
    updateRunProgress(0, '대기 중', '');
    FEATURE_KEYS.forEach((key) => {
      if (state.results[key]) {
        state.results[key].count = 0;
        state.results[key].stale = true;
        state.results[key].hasRun = false;
      }
    });
    FEATURE_KEYS.forEach((key) => syncFeatureRow(key));
    updateRunActionLabel();
    updateBqcSidebar();
    post('hub:multi-clear', {});
  }

  function getFeatureReadiness(feature) {
    if (!feature?.enabled) {
      return { label: 'OFF', className: 'chip--off' };
    }
    if (state.ui.activeFeatureKey === 'parameterduplication') {
      const committed = feature.configCommitted || {};
      const scope = normalizeParameterDuplicationScope(committed.scope);
      const names = Array.isArray(committed.parameterNames) ? committed.parameterNames : [];
      if (scope === 'selected' && !names.length) {
        return { label: '설정 필요', className: 'chip--warn' };
      }
    }
    if (state.ui.activeFeatureKey === 'parametermissing') {
      const committed = createParameterMissingConfigSnapshot(feature.configCommitted || {});
      if (!committed.parameterNames.length || hasIncompleteParameterMissingConfig(committed)) {
        return { label: '설정 필요', className: 'chip--warn' };
      }
    }
    if (state.ui.activeFeatureKey === 'floorinfo') {
      const committed = feature.configCommitted || {};
      const rules = normalizeFloorInfoRules(committed.levelRules);
      const selectedRules = rules.filter((rule) => rule.useAsBoundary !== false);
      const configuredCount = selectedRules.filter((rule) => String(rule.expectedValue || '').trim()).length;
      if (!String(committed.parameterName || '').trim() || !selectedRules.length || configuredCount < selectedRules.length) {
        return { label: '설정 필요', className: 'chip--warn' };
      }
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
    return { label: '검토 준비 완료', className: 'chip--ok' };
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
    if (key === 'floorinfo') refreshFloorInfoFeatureSummary();
    if (key === 'familysuitability') refreshFamilySuitabilityFeatureSummary();
    if (key === 'tapalign') refreshTapAlignFeatureSummary();
    if (key === 'dupclash') refreshDupClashFeatureSummary();
    if (key === 'worksetassignment') refreshWorksetAssignmentFeatureSummary();
    if (key === 'parameterduplication') refreshParameterDuplicationFeatureSummary();
    if (key === 'parametermissing') refreshParameterMissingFeatureSummary();
  }

  function buildCommonSummary() {
    const committed = state.common.configCommitted;
    const extraCount = committed.extraParams ? committed.extraParams.split(',').filter((v) => v.trim()).length : 0;
    const filterText = committed.targetFilter ? `포함 필터 ${committed.targetFilter}` : '포함 필터가 없습니다.';
    const excludeText = committed.excludeTargetFilter ? `제외 필터 ${committed.excludeTargetFilter}` : '제외 필터가 없습니다.';
    return `추가 파라미터 ${extraCount}개 / ${filterText} / ${excludeText}`;
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
    updateFeatureSummary('tapalign');
    updateFeatureSummary('dupclash');
    if (state.ui.controls.tapalign?.renderCommonSummary) state.ui.controls.tapalign.renderCommonSummary();
    if (state.ui.controls.dupclash?.renderCommonSummary) state.ui.controls.dupclash.renderCommonSummary();
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
        '추가 파라미터 값은 콤마로 구분해 입력합니다.',
        '검토 대상 필터는 필터에 맞는 객체만 검토합니다.',
        '검토 제외 대상 필터는 필터에 맞는 객체를 제외하고 검토합니다.'
      ];
    }
    if (key === 'connector') {
      return [
        '공유 파라미터 txt 목록에서 검토 대상을 검색해 선택합니다.',
        '여러 파라미터를 선택하면 같은 논리로 연속성 검토를 진행합니다.',
        '추가 추출 파라미터는 인스턴스 우선, 없으면 타입 파라미터에서도 값을 찾습니다.',
        '좌표 X/Y 옵션을 켜면 결과 엑셀에 좌표 열이 자동으로 추가됩니다.',
        '선형 길이/방향 벡터 옵션을 켜면 결과 엑셀에 선형 길이(Curve Length)와 방향 벡터(Direction X/Y/Z) 열이 함께 추가됩니다.',
        '허용 범위와 단위는 기존 커넥터 검토 로직 그대로 적용됩니다.'
      ];
    }
    if (key === 'floorinfo') {
      return [
        '활성 문서의 레벨 중 층정보 영역을 구분할 레벨만 선택하고, 그 구간별 기대 층정보 값을 설정합니다.',
        '선택하지 않은 중간 레벨은 무시되므로 1F, 1.5F, 2F 중 1F와 2F만 선택해 1F~2F 구간을 1F로 판정할 수 있습니다.',
        '객체가 여러 레벨 구간을 관통하면 가장 아래 구간의 층정보를 기대값으로 사용합니다.',
        '공통 옵션의 검토 대상/검토 제외 필터로 평가 대상을 제한하고, 추가 파라미터 값은 결과 엑셀 열로 함께 저장합니다.',
        'BQC 보조 기능이라 필요한 프로젝트에서만 선택해 추가 검토로 활용하면 됩니다.'
      ];
    }
    if (key === 'familysuitability') {
      return [
        '기준 엑셀에서 카테고리, 패밀리, 타입 헤더를 찾아 실제 사용 객체의 조합과 비교합니다.',
        '시스템 타입과 로더블 패밀리를 모두 집계하지만, 타입 정의만 있고 실제 배치되지 않은 항목은 제외합니다.',
        '기준 일치, 기준 미일치, 필터 일치의 검토 문구를 각각 다르게 입력할 수 있습니다.',
        '필터 규칙은 OR 조건으로 평가되며, 하나라도 일치하면 필터 검토 문구를 적용합니다. 여러 개가 동시에 일치하면 위에서 아래 순서로 먼저 등록한 문구를 사용합니다.',
        '최근 적용한 설정은 자동으로 기억하고, 자주 쓰는 조합은 이름을 붙여 저장/불러오기로 재사용할 수 있습니다.'
      ];
    }
    if (key === 'tapalign') {
      return [
        '탭 또는 분기 피팅의 연결 커넥터 축이 연결된 배관/덕트 중심축을 통과하는지 확인합니다.',
        '허용 범위 이내의 중심축 이탈은 오류로 보지 않으며, 초과한 경우만 결과에 포함합니다.',
        '거리 단위는 설정값을 따르고, 결과 내용 언어는 엑셀 내보내기 시점에 선택합니다.',
        '추가 추출 파라미터, 포함/제외 대상 필터와 추가 추출 옵션은 BQC 공통 설정 값을 사용합니다.'
      ];
    }
    if (key === 'dupclash') {
      return [
        '기능 설정에서 중복 검토 또는 자체 간섭 검토 중 하나를 선택합니다.',
        '기존 중복 / 자체 간섭 화면의 개별 범위, 제외 키워드 같은 세부 필터는 사용하지 않습니다.',
        '공통 설정의 포함/제외 대상 필터만 그대로 따라가며, 추가 파라미터는 결과 엑셀 열로 저장합니다.',
        '실행 결과 요약과 엑셀 내보내기도 선택한 모드 결과만 표시합니다.'
      ];
    }
    if (key === 'guid') {
      return [
        '먼저 GUID 검토 결과를 삭제용 엑셀로 내보내고, 같은 엑셀을 수정한 뒤 다시 불러오는 흐름으로 사용합니다.',
        "삭제할 행은 엑셀의 '삭제여부' 열에 '삭제'라고 입력하면 됩니다.",
        '삭제용 엑셀 불러오기는 마지막 숨김 키 행을 기준으로 프로젝트/로드 패밀리 파라미터를 찾아 적용합니다.',
        '센트럴 파일과 워크셰어링 파일은 항상 로컬로, 모든 웍셋을 닫은 상태로 열어 처리합니다.'
      ];
    }
    if (key === 'worksetassignment') {
      return [
        '모델 카테고리 객체를 순회하면서 현재 배정된 웍셋 이름을 확인합니다.',
        '오류 대상 웍셋 이름을 입력하면 그 웍셋에 속한 객체만 오류로 출력하고, 비워두면 기본 웍셋(Workset1) 이외의 웍셋을 모두 오류로 봅니다.',
        '공통 옵션의 검토 대상/검토 제외 필터로 검토 범위를 제한하고, 추가 파라미터 값은 오류 결과 행의 엑셀 열로 함께 저장합니다.',
        '오류가 하나도 없을 때는 검토한 객체 수와 함께 기본 웍셋(Workset1) 정상 배정 요약 1행만 내보냅니다.'
      ];
    }
    if (key === 'parameterduplication') {
      return [
        '추가된 프로젝트 파라미터를 수집한 뒤 이름이 중복되는 파라미터를 찾습니다.',
        '검토 범위를 전체 프로젝트 파라미터 또는 지정한 파라미터 이름만으로 제한할 수 있습니다.',
        '지정 파라미터는 현재 Revit에 연결된 공유파라미터 파일에서 검색해 선택합니다.',
        '전체 검토 상태에서 목록을 클릭하면 지정 파라미터 검토로 자동 전환됩니다.',
        '최근 적용한 검토 대상은 자동으로 기억되며, 최근 검토 항목에서 바로 다시 불러올 수 있습니다.',
        '엑셀은 BQC 결과 포맷으로 내보내고, 중복이면 오류(Error), 아니면 정상(OK)과 성공 메시지를 기록합니다.'
      ];
    }
    if (key === 'parametermissing') {
      return [
        '누락 검토 파라미터는 현재 Revit에 연결된 공유파라미터 TXT의 Text 정의만 검색해 선택합니다.',
        '검토 대상은 BQC 공통 옵션의 검토 대상/검토 제외 대상 필터를 그대로 사용합니다.',
        '공통 옵션의 추가 파라미터 값은 누락 오류 결과 엑셀 열로 함께 저장합니다.',
        '누락 예외 필터는 선택한 파라미터마다 따로 설정하며, 조건이 맞으면 비어 있어도 누락으로 보지 않습니다.',
        '최근 적용한 설정은 자동으로 기억하며, 현재 설정은 파일로 저장하거나 다시 불러올 수 있습니다.',
        '값이 필요한 조건은 파라미터명과 값을 함께 입력하고, 값 있음 / 값 없음 조건은 값 없이 사용할 수 있습니다.'
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
        '십진 피트, 미터, 밀리미터를 지원합니다.'
      ];
    }
    if (key === 'linkworkset') {
      return [
        '최상위 Revit 링크를 순회하면서 현재 로드 상태와 열려 있는 사용자 웍셋 현황을 점검합니다.',
        '자동 적용을 켜면 기본 웍셋(Workset1)만 열리도록 링크를 재로드합니다.',
        '활성 문서와 다중 RVT 배치 실행 모두 같은 결과 포맷으로 요약과 엑셀 내보내기를 지원합니다.'
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

  function getPathLeafLabel(path, fallback = '') {
    const text = String(path || '').trim();
    if (!text) return fallback;
    const parts = text.split(/[\\/]/).filter(Boolean);
    return parts.length ? parts[parts.length - 1] : text;
  }

  function normalizeParameterDuplicationScope(value) {
    return String(value || '').trim().toLowerCase() === 'selected' ? 'selected' : 'all';
  }

  function parseParameterDuplicationNames(value) {
    const seen = new Set();
    return String(value || '')
      .split(/[\n,;\r]+/)
      .map((item) => item.trim())
      .filter((item) => {
        if (!item) return false;
        const key = item.toLowerCase();
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
  }

  function buildParameterDuplicationNamesText(items) {
    return (Array.isArray(items) ? items : [])
      .map((item) => String(item || '').trim())
      .filter(Boolean)
      .join('\n');
  }

  function buildParameterDuplicationNamePreview(items, maxCount = 6) {
    const names = (Array.isArray(items) ? items : [])
      .map((item) => String(item || '').trim())
      .filter(Boolean);
    if (!names.length) return '입력한 파라미터가 없습니다.';
    if (names.length <= maxCount) return names.join(', ');
    return `${names.slice(0, maxCount).join(', ')} 외 ${names.length - maxCount}개`;
  }

  function createParameterDuplicationConfigSnapshot(config) {
    return {
      scope: normalizeParameterDuplicationScope(config?.scope),
      parameterNames: parseParameterDuplicationNames(Array.isArray(config?.parameterNames) ? config.parameterNames.join('\n') : config?.parameterNames),
      sharedParamSourcePath: String(config?.sharedParamSourcePath || state.sharedParamStatus?.path || '').trim(),
      sharedParamImportCount: Number(config?.sharedParamImportCount) || 0
    };
  }

  function parseParameterMissingNames(value) {
    return parseParameterDuplicationNames(value);
  }

  function normalizeParameterMissingCombinationMode(value, fallback = 'And') {
    const normalized = String(value || '').trim().toLowerCase();
    if (normalized === 'or') return 'Or';
    return String(fallback || '').trim().toLowerCase() === 'or' ? 'Or' : 'And';
  }

  function normalizeParameterMissingTargetFilterMode(value) {
    const normalized = String(value || '').trim().toLowerCase();
    return normalized === 'exclude' || normalized === 'except' || normalized === 'not' ? 'exclude' : 'include';
  }

  function resolveParameterMissingTargetFilterModeLabel(value) {
    return normalizeParameterMissingTargetFilterMode(value) === 'exclude'
      ? '조건에 해당하지 않는 객체만 검토'
      : '조건에 해당하는 객체만 검토';
  }

  function createEmptyParameterMissingCondition() {
    return {
      enabled: true,
      parameterName: '',
      operatorName: 'Equals',
      value: ''
    };
  }

  function normalizeParameterMissingConditionRows(raw, options = {}) {
    const keepEmpty = !!options.keepEmpty;
    const rows = (Array.isArray(raw) ? raw : [])
      .map((row) => ({
        enabled: row?.enabled !== false,
        parameterName: String(row?.parameterName || '').trim(),
        operatorName: PARAMETER_MISSING_FILTER_OPERATORS.includes(String(row?.operatorName || '').trim())
          ? String(row?.operatorName || '').trim()
          : 'Equals',
        value: String(row?.value || '')
      }))
      .filter((row) => keepEmpty || row.parameterName || row.value);
    if (keepEmpty && !rows.length) rows.push(createEmptyParameterMissingCondition());
    return rows;
  }

  function createParameterMissingTargetFilterState(raw) {
    return {
      enabled: true,
      mode: normalizeParameterMissingTargetFilterMode(raw?.mode || raw?.targetMode || raw?.filterMode),
      combinationMode: normalizeParameterMissingCombinationMode(raw?.combinationMode, 'And'),
      conditions: normalizeParameterMissingConditionRows(raw?.conditions, { keepEmpty: true }),
      assignments: []
    };
  }

  function createEmptyParameterMissingRule(parameterName = '') {
    return {
      enabled: true,
      parameterName: String(parameterName || '').trim(),
      combinationMode: 'Or',
      conditions: [createEmptyParameterMissingCondition()]
    };
  }

  function normalizeParameterMissingExceptionRules(raw, parameterNames, options = {}) {
    const keepEmpty = !!options.keepEmpty;
    const selectedNames = parseParameterMissingNames(Array.isArray(parameterNames) ? parameterNames.join('\n') : parameterNames);
    const mapped = new Map();

    (Array.isArray(raw) ? raw : []).forEach((rule) => {
      const parameterName = String(rule?.parameterName || '').trim();
      if (!parameterName) return;
      if (!selectedNames.some((name) => name.toLowerCase() === parameterName.toLowerCase())) return;
      mapped.set(parameterName.toLowerCase(), {
        enabled: rule?.enabled !== false,
        parameterName,
        combinationMode: normalizeParameterMissingCombinationMode(rule?.combinationMode, 'Or'),
        conditions: normalizeParameterMissingConditionRows(rule?.conditions, { keepEmpty })
      });
    });

    return selectedNames.map((parameterName) => {
      const existing = mapped.get(parameterName.toLowerCase());
      if (existing) return existing;
      return keepEmpty ? createEmptyParameterMissingRule(parameterName) : {
        enabled: true,
        parameterName,
        combinationMode: 'Or',
        conditions: []
      };
    });
  }

  function createParameterMissingConfigSnapshot(config) {
    const parameterNames = parseParameterMissingNames(
      Array.isArray(config?.parameterNames) ? config.parameterNames.join('\n') : config?.parameterNames
    );
    return {
      parameterNames,
      targetFilter: createParameterMissingTargetFilterState(config?.targetFilter),
      exceptionRules: normalizeParameterMissingExceptionRules(config?.exceptionRules, parameterNames, { keepEmpty: true })
    };
  }

  function isTextSharedParameterItem(item) {
    const token = String(item?.dataTypeToken || '').trim().toLowerCase();
    if (!token) return true;
    return token.includes('text') || token.includes('string');
  }

  function getParameterMissingSharedParamItems() {
    const seen = new Set();
    return (Array.isArray(state.sharedParamItems) ? state.sharedParamItems : [])
      .map((item) => ({
        name: String(item?.name || '').trim(),
        groupName: String(item?.groupName || '').trim(),
        guid: String(item?.guid || '').trim(),
        dataTypeToken: String(item?.dataTypeToken || '').trim()
      }))
      .filter((item) => item.name && isTextSharedParameterItem(item))
      .filter((item) => {
        const key = item.name.toLowerCase();
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      })
      .sort((a, b) => a.name.localeCompare(b.name, 'ko'));
  }

  function isParameterMissingConditionValueless(operatorName) {
    return PARAMETER_MISSING_VALUELESS_OPERATORS.has(String(operatorName || '').trim());
  }

  function isParameterMissingConditionConfigured(row) {
    const parameterName = String(row?.parameterName || '').trim();
    const value = String(row?.value || '').trim();
    const operatorName = String(row?.operatorName || 'Equals').trim();
    if (!parameterName) return false;
    return isParameterMissingConditionValueless(operatorName) || !!value;
  }

  function isParameterMissingConditionIncomplete(row) {
    const parameterName = String(row?.parameterName || '').trim();
    const value = String(row?.value || '').trim();
    const operatorName = String(row?.operatorName || 'Equals').trim();
    if (!parameterName && !value) return false;
    if (!parameterName) return true;
    return !isParameterMissingConditionValueless(operatorName) && !value;
  }

  function countParameterMissingConfiguredConditions(rows) {
    return (Array.isArray(rows) ? rows : []).filter((row) => isParameterMissingConditionConfigured(row)).length;
  }

  function hasIncompleteParameterMissingConfig(config) {
    if (createParameterMissingTargetFilterState(config?.targetFilter).conditions.some((row) => isParameterMissingConditionIncomplete(row))) {
      return true;
    }
    return (Array.isArray(config?.exceptionRules) ? config.exceptionRules : []).some((rule) =>
      (Array.isArray(rule?.conditions) ? rule.conditions : []).some((row) => isParameterMissingConditionIncomplete(row))
    );
  }

  function buildParameterMissingConditionSummary(rows, combinationMode) {
    const configured = (Array.isArray(rows) ? rows : []).filter((row) => isParameterMissingConditionConfigured(row));
    if (!configured.length) return '조건이 없습니다.';
    const joiner = normalizeParameterMissingCombinationMode(combinationMode, 'And') === 'Or' ? ' OR ' : ' AND ';
    return configured
      .map((row) => {
        const operatorName = String(row.operatorName || 'Equals').trim();
        if (isParameterMissingConditionValueless(operatorName)) {
          return `${row.parameterName} ${operatorName}`;
        }
        return `${row.parameterName} ${operatorName} ${String(row.value || '').trim()}`;
      })
      .join(joiner);
  }

  function buildParameterMissingSelectionPreview(items, maxCount = 4) {
    const names = (Array.isArray(items) ? items : []).map((item) => String(item || '').trim()).filter(Boolean);
    if (!names.length) return '선택한 파라미터가 없습니다.';
    if (names.length <= maxCount) return names.join(', ');
    return `${names.slice(0, maxCount).join(', ')} 외 ${names.length - maxCount}개`;
  }

  function buildParameterDuplicationConfigKey(config) {
    const snapshot = createParameterDuplicationConfigSnapshot(config);
    if (snapshot.scope === 'all') return 'all';
    return `selected:${snapshot.parameterNames.map((name) => name.toLowerCase()).join('|')}`;
  }

  function formatParameterDuplicationRecentTimestamp(value) {
    const date = value ? new Date(value) : null;
    if (!date || Number.isNaN(date.getTime())) return '';
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hour = String(date.getHours()).padStart(2, '0');
    const minute = String(date.getMinutes()).padStart(2, '0');
    return `${month}-${day} ${hour}:${minute}`;
  }

  function buildParameterDuplicationRecentOptionLabel(item) {
    const snapshot = createParameterDuplicationConfigSnapshot(item);
    const baseLabel = snapshot.scope === 'all'
      ? '전체 프로젝트 파라미터'
      : `지정 ${snapshot.parameterNames.length}개 · ${buildParameterDuplicationNamePreview(snapshot.parameterNames, 3)}`;
    const sourcePath = String(snapshot.sharedParamSourcePath || '').trim();
    const sourceLabel = sourcePath ? getPathLeafLabel(sourcePath, sourcePath) : '';
    const timeLabel = formatParameterDuplicationRecentTimestamp(item?.updatedAt);
    return [timeLabel, baseLabel, sourceLabel].filter(Boolean).join(' · ');
  }

  function loadFamilySuitabilityPresets() {
    try {
      const raw = localStorage.getItem(FAMILY_SUITABILITY_PRESET_KEY) || '[]';
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];
      return parsed
        .filter((row) => row && typeof row.name === 'string')
        .map((row) => ({
          name: String(row.name || '').trim(),
          ...normalizeFamilySuitabilityConfig(row)
        }))
        .filter((row) => !!row.name)
        .sort((a, b) => a.name.localeCompare(b.name, 'ko'));
    } catch {
      return [];
    }
  }

  function saveFamilySuitabilityPresets(presets) {
    try {
      localStorage.setItem(FAMILY_SUITABILITY_PRESET_KEY, JSON.stringify(Array.isArray(presets) ? presets : []));
    } catch {
    }
  }

  function saveFamilySuitabilityPreset(name, config) {
    const key = String(name || '').trim();
    if (!key) return;
    const presets = loadFamilySuitabilityPresets();
    const next = {
      name: key,
      ...normalizeFamilySuitabilityConfig(config)
    };
    const index = presets.findIndex((preset) => preset.name === key);
    if (index >= 0) presets[index] = next;
    else presets.push(next);
    saveFamilySuitabilityPresets(presets);
  }

  function applyFamilySuitabilityPreset(targetConfig, name) {
    const key = String(name || '').trim();
    if (!targetConfig || !key) return false;
    const preset = loadFamilySuitabilityPresets().find((item) => item.name === key);
    if (!preset) return false;
    Object.assign(targetConfig, normalizeFamilySuitabilityConfig(preset, { keepEmptyFilters: true }));
    if (!targetConfig.filterRules.length) {
      targetConfig.filterRules = [createEmptyFamilySuitabilityFilterRule()];
    }
    return true;
  }

  function deleteFamilySuitabilityPreset(name) {
    const key = String(name || '').trim();
    if (!key) return;
    saveFamilySuitabilityPresets(loadFamilySuitabilityPresets().filter((preset) => preset.name !== key));
  }

  function loadParameterDuplicationPresets() {
    try {
      const raw = localStorage.getItem(PARAMETER_DUPLICATION_PRESET_KEY) || '[]';
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];
      return parsed
        .filter((row) => row && typeof row.name === 'string')
        .map((row) => ({
          name: String(row.name || '').trim(),
          ...createParameterDuplicationConfigSnapshot(row)
        }))
        .filter((row) => !!row.name)
        .sort((a, b) => a.name.localeCompare(b.name, 'ko'));
    } catch {
      return [];
    }
  }

  function saveParameterDuplicationPresets(presets) {
    try {
      localStorage.setItem(PARAMETER_DUPLICATION_PRESET_KEY, JSON.stringify(Array.isArray(presets) ? presets : []));
    } catch {
    }
  }

  function saveParameterDuplicationPreset(name, config) {
    const key = String(name || '').trim();
    if (!key) return;
    const presets = loadParameterDuplicationPresets();
    const next = {
      name: key,
      ...createParameterDuplicationConfigSnapshot(config)
    };
    const index = presets.findIndex((preset) => preset.name === key);
    if (index >= 0) presets[index] = next;
    else presets.push(next);
    saveParameterDuplicationPresets(presets);
  }

  function applyParameterDuplicationPreset(targetConfig, name) {
    const key = String(name || '').trim();
    if (!targetConfig || !key) return false;
    const preset = loadParameterDuplicationPresets().find((item) => item.name === key);
    if (!preset) return false;
    Object.assign(targetConfig, createParameterDuplicationConfigSnapshot(preset));
    return true;
  }

  function deleteParameterDuplicationPreset(name) {
    const key = String(name || '').trim();
    if (!key) return;
    saveParameterDuplicationPresets(loadParameterDuplicationPresets().filter((preset) => preset.name !== key));
  }

  function loadParameterDuplicationRecent() {
    try {
      const raw = localStorage.getItem(PARAMETER_DUPLICATION_RECENT_KEY) || '[]';
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];
      return parsed
        .filter((row) => !!row)
        .map((row) => {
          const snapshot = createParameterDuplicationConfigSnapshot(row);
          return {
            key: String(row.key || buildParameterDuplicationConfigKey(snapshot)).trim(),
            updatedAt: String(row.updatedAt || '').trim(),
            ...snapshot
          };
        })
        .filter((row) => row.scope === 'all' || row.parameterNames.length > 0)
        .sort((a, b) => String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')));
    } catch {
      return [];
    }
  }

  function saveParameterDuplicationRecent(items) {
    try {
      localStorage.setItem(PARAMETER_DUPLICATION_RECENT_KEY, JSON.stringify(Array.isArray(items) ? items : []));
    } catch {
    }
  }

  function rememberParameterDuplicationRecent(config) {
    const snapshot = createParameterDuplicationConfigSnapshot(config);
    if (snapshot.scope === 'selected' && !snapshot.parameterNames.length) return;
    const key = buildParameterDuplicationConfigKey(snapshot);
    const next = {
      key,
      updatedAt: new Date().toISOString(),
      ...snapshot
    };
    const recents = loadParameterDuplicationRecent().filter((item) => item.key !== key);
    recents.unshift(next);
    saveParameterDuplicationRecent(recents.slice(0, PARAMETER_DUPLICATION_RECENT_LIMIT));
  }

  function applyParameterDuplicationRecent(targetConfig, key) {
    const recentKey = String(key || '').trim();
    if (!targetConfig || !recentKey) return false;
    const recent = loadParameterDuplicationRecent().find((item) => item.key === recentKey);
    if (!recent) return false;
    Object.assign(targetConfig, createParameterDuplicationConfigSnapshot(recent));
    return true;
  }

  function clearParameterDuplicationRecent() {
    saveParameterDuplicationRecent([]);
  }

  function createParameterMissingSerializableConfig(config) {
    const snapshot = createParameterMissingConfigSnapshot(config);
    const targetFilter = createParameterMissingTargetFilterState(snapshot.targetFilter);
    const targetFilterConditions = (Array.isArray(targetFilter.conditions) ? targetFilter.conditions : [])
      .filter((row) => isParameterMissingConditionConfigured(row))
      .map((row) => {
        const operatorName = String(row?.operatorName || 'Equals').trim() || 'Equals';
        return {
          enabled: row?.enabled !== false,
          parameterName: String(row?.parameterName || '').trim(),
          operatorName,
          value: isParameterMissingConditionValueless(operatorName) ? '' : String(row?.value || '').trim()
        };
      });
    return {
      parameterNames: Array.isArray(snapshot.parameterNames) ? [...snapshot.parameterNames] : [],
      targetFilter: {
        enabled: targetFilterConditions.length > 0,
        mode: normalizeParameterMissingTargetFilterMode(targetFilter.mode),
        combinationMode: normalizeParameterMissingCombinationMode(targetFilter.combinationMode, 'And'),
        conditions: targetFilterConditions,
        assignments: []
      },
      exceptionRules: (Array.isArray(snapshot.exceptionRules) ? snapshot.exceptionRules : [])
        .map((rule) => ({
          parameterName: String(rule?.parameterName || '').trim(),
          combinationMode: normalizeParameterMissingCombinationMode(rule?.combinationMode, 'Or'),
          conditions: (Array.isArray(rule?.conditions) ? rule.conditions : [])
            .filter((row) => isParameterMissingConditionConfigured(row))
            .map((row) => {
              const operatorName = String(row?.operatorName || 'Equals').trim() || 'Equals';
              return {
                parameterName: String(row?.parameterName || '').trim(),
                operatorName,
                value: isParameterMissingConditionValueless(operatorName) ? '' : String(row?.value || '').trim()
              };
            })
        }))
        .filter((rule) => rule.parameterName && rule.conditions.length > 0)
    };
  }

  function buildParameterMissingConfigKey(config) {
    const snapshot = createParameterMissingSerializableConfig(config);
    return JSON.stringify({
      parameterNames: snapshot.parameterNames.map((name) => String(name || '').trim().toLowerCase()),
      targetFilter: {
        mode: normalizeParameterMissingTargetFilterMode(snapshot.targetFilter?.mode),
        combinationMode: normalizeParameterMissingCombinationMode(snapshot.targetFilter?.combinationMode, 'And'),
        conditions: (Array.isArray(snapshot.targetFilter?.conditions) ? snapshot.targetFilter.conditions : []).map((row) => ({
          parameterName: String(row.parameterName || '').trim().toLowerCase(),
          operatorName: String(row.operatorName || 'Equals').trim() || 'Equals',
          value: String(row.value || '').trim()
        }))
      },
      exceptionRules: snapshot.exceptionRules.map((rule) => ({
        parameterName: String(rule.parameterName || '').trim().toLowerCase(),
        combinationMode: normalizeParameterMissingCombinationMode(rule.combinationMode, 'Or'),
        conditions: rule.conditions.map((row) => ({
          parameterName: String(row.parameterName || '').trim().toLowerCase(),
          operatorName: String(row.operatorName || 'Equals').trim() || 'Equals',
          value: String(row.value || '').trim()
        }))
      }))
    });
  }

  function countParameterMissingConfiguredRules(config) {
    const snapshot = createParameterMissingConfigSnapshot(config);
    return (Array.isArray(snapshot.exceptionRules) ? snapshot.exceptionRules : [])
      .filter((rule) => countParameterMissingConfiguredConditions(rule.conditions) > 0)
      .length;
  }

  function countParameterMissingTargetFilterConditions(config) {
    return countParameterMissingConfiguredConditions(createParameterMissingConfigSnapshot(config).targetFilter.conditions);
  }

  function buildParameterMissingRecentOptionLabel(item) {
    const snapshot = createParameterMissingConfigSnapshot(item);
    const timeLabel = formatParameterDuplicationRecentTimestamp(item?.updatedAt);
    const exceptionRuleCount = countParameterMissingConfiguredRules(snapshot);
    return [
      timeLabel,
      `파라미터 ${snapshot.parameterNames.length}개`,
      countParameterMissingTargetFilterConditions(snapshot) ? '객체 필터 적용' : '',
      exceptionRuleCount ? `예외 ${exceptionRuleCount}개` : '예외가 없습니다.',
      buildParameterMissingSelectionPreview(snapshot.parameterNames, 3)
    ].filter(Boolean).join(' · ');
  }

  function loadParameterMissingRecent() {
    try {
      const raw = localStorage.getItem(PARAMETER_MISSING_RECENT_KEY) || '[]';
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];
      return parsed
        .filter((row) => !!row)
        .map((row) => {
          const snapshot = createParameterMissingConfigSnapshot(row);
          return {
            key: String(row.key || buildParameterMissingConfigKey(snapshot)).trim(),
            updatedAt: String(row.updatedAt || '').trim(),
            ...snapshot
          };
        })
        .filter((row) => row.parameterNames.length > 0)
        .sort((a, b) => String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')));
    } catch {
      return [];
    }
  }

  function saveParameterMissingRecent(items) {
    try {
      localStorage.setItem(PARAMETER_MISSING_RECENT_KEY, JSON.stringify(Array.isArray(items) ? items : []));
    } catch {
    }
  }

  function rememberParameterMissingRecent(config) {
    const snapshot = createParameterMissingSerializableConfig(config);
    if (!snapshot.parameterNames.length) return;
    const key = buildParameterMissingConfigKey(snapshot);
    const next = {
      key,
      updatedAt: new Date().toISOString(),
      ...snapshot
    };
    const recents = loadParameterMissingRecent().filter((item) => item.key !== key);
    recents.unshift(next);
    saveParameterMissingRecent(recents.slice(0, PARAMETER_MISSING_RECENT_LIMIT));
  }

  function applyParameterMissingRecent(targetConfig, key) {
    const recentKey = String(key || '').trim();
    if (!targetConfig || !recentKey) return false;
    const recent = loadParameterMissingRecent().find((item) => item.key === recentKey);
    if (!recent) return false;
    Object.assign(targetConfig, createParameterMissingConfigSnapshot(recent));
    return true;
  }

  function clearParameterMissingRecent() {
    saveParameterMissingRecent([]);
  }

  function buildParameterMissingPresetSnapshot(config) {
    return {
      kind: PARAMETER_MISSING_PRESET_KIND,
      version: PARAMETER_MISSING_PRESET_VERSION,
      feature: 'parametermissing',
      savedAt: new Date().toISOString(),
      config: createParameterMissingSerializableConfig(config)
    };
  }

  function buildParameterMissingPresetJson(config) {
    return JSON.stringify(buildParameterMissingPresetSnapshot(config), null, 2);
  }

  function buildParameterMissingPresetDefaultName(config) {
    const snapshot = createParameterMissingConfigSnapshot(config);
    const dateToken = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const countLabel = snapshot.parameterNames.length ? `${snapshot.parameterNames.length}params` : 'config';
    return `${dateToken}_${sanitizeFavoritePresetFileLabel(`parameter-missing-${countLabel}`)}${PARAMETER_MISSING_PRESET_EXTENSION}`;
  }

  function parseParameterMissingPresetSnapshot(json) {
    let parsed = null;
    try {
      parsed = JSON.parse(String(json || ''));
    } catch {
      throw new Error('파라미터 누락 검토 설정 파일 JSON을 읽지 못했습니다.');
    }

    if (!parsed || typeof parsed !== 'object') {
      throw new Error('올바른 파라미터 누락 검토 설정 파일이 아닙니다.');
    }

    if (parsed.kind && parsed.kind !== PARAMETER_MISSING_PRESET_KIND) {
      throw new Error('파라미터 누락 검토 설정 파일이 아닙니다.');
    }

    const version = Number(parsed.version || PARAMETER_MISSING_PRESET_VERSION);
    if (version > PARAMETER_MISSING_PRESET_VERSION) {
      throw new Error('현재 버전보다 최신 파라미터 누락 검토 설정 파일입니다.');
    }

    const source = parsed.config && typeof parsed.config === 'object' ? parsed.config : parsed;
    const snapshot = createParameterMissingConfigSnapshot(source);
    if (!snapshot.parameterNames.length) {
      throw new Error('파라미터 누락 검토 설정 파일에 검토 파라미터가 없습니다.');
    }
    return snapshot;
  }

  function downloadParameterMissingPresetInBrowser(json, fileName) {
    const blob = new Blob([json], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.append(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  function openParameterMissingPresetFileInBrowser(onLoaded) {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = `.json,${PARAMETER_MISSING_PRESET_EXTENSION},application/json`;
    input.addEventListener('change', () => {
      const file = input.files && input.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        if (typeof onLoaded === 'function') {
          onLoaded({
            fileName: file.name,
            path: file.name,
            json: typeof reader.result === 'string' ? reader.result : ''
          });
        }
      };
      reader.onerror = () => {
        toast('파라미터 누락 검토 설정 파일을 읽지 못했습니다.', 'err');
      };
      reader.readAsText(file, 'utf-8');
    }, { once: true });
    input.click();
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

  function normalizeFloorInfoRules(raw) {
    const seen = new Set();
    const rules = [];
    const items = Array.isArray(raw) ? raw : [];
    items.forEach((item) => {
      const levelName = String(item?.levelName || '').trim();
      if (!levelName) return;
      const key = levelName.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      rules.push({
        levelName,
        absoluteZFt: Number(item?.absoluteZFt) || 0,
        absoluteZMm: Number(item?.absoluteZMm) || 0,
        useAsBoundary: item?.useAsBoundary !== false,
        expectedValue: String(item?.expectedValue || '')
      });
    });
    return rules.sort((a, b) => a.absoluteZFt - b.absoluteZFt || a.levelName.localeCompare(b.levelName, 'ko'));
  }

  function mergeFloorInfoRules(existingRules, snapshotLevels) {
    const existingMap = new Map(normalizeFloorInfoRules(existingRules).map((rule) => [rule.levelName.toLowerCase(), rule]));
    return normalizeFloorInfoRules(snapshotLevels).map((level) => ({
      levelName: level.levelName,
      absoluteZFt: Number(level.absoluteZFt) || 0,
      absoluteZMm: Number(level.absoluteZMm) || 0,
      useAsBoundary: existingMap.get(level.levelName.toLowerCase())?.useAsBoundary !== false,
      expectedValue: existingMap.get(level.levelName.toLowerCase())?.expectedValue || ''
    }));
  }

  function createEmptyFamilySuitabilityFilterRule() {
    return {
      target: 'familyOrType',
      keyword: '',
      reviewText: ''
    };
  }

  function normalizeFamilySuitabilityFilterTarget(value) {
    const normalized = String(value || 'familyOrType').trim().toLowerCase();
    if (normalized === 'family') return 'family';
    if (normalized === 'type' || normalized === 'typename') return 'type';
    return 'familyOrType';
  }

  function normalizeFamilySuitabilityFilterRules(raw, options = {}) {
    const keepEmpty = !!options.keepEmpty;
    const items = Array.isArray(raw) ? raw : [];
    const rules = [];
    items.forEach((item) => {
      const target = normalizeFamilySuitabilityFilterTarget(item?.target);
      const keyword = String(item?.keyword || '').trim();
      const reviewText = String(item?.reviewText || '').trim();
      if (!keepEmpty && !keyword && !reviewText) return;
      rules.push({ target, keyword, reviewText });
    });
    return rules;
  }

  function normalizeFamilySuitabilityConfig(raw, options = {}) {
    const keepEmptyFilters = !!options.keepEmptyFilters;
    return {
      criteriaExcelPath: String(raw?.criteriaExcelPath || '').trim(),
      criteriaRowCount: Number(raw?.criteriaRowCount) || 0,
      criteriaComboCount: Number(raw?.criteriaComboCount) || 0,
      criteriaSheetCount: Number(raw?.criteriaSheetCount) || 0,
      matchReviewText: String(raw?.matchReviewText || '').trim(),
      mismatchReviewText: String(raw?.mismatchReviewText || '').trim(),
      filterRules: normalizeFamilySuitabilityFilterRules(raw?.filterRules, { keepEmpty: keepEmptyFilters })
    };
  }

  function getFamilySuitabilityCriteriaLabel(path) {
    const text = String(path || '').trim();
    if (!text) return '기준 엑셀이 선택되지 않았습니다.';
    const parts = text.split(/[\\/]/).filter(Boolean);
    return parts.length ? parts[parts.length - 1] : text;
  }

  function requestFloorInfoConfig(context = 'manual') {
    post('floorinfo:config-load', { source: 'multi', context });
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
    const excludeEndDummy = !!(feature.configDraft?.excludeEndDummy ?? feature.configCommitted?.excludeEndDummy);
    const includePointXY = !!(feature.configDraft?.includePointXY ?? feature.configCommitted?.includePointXY);
    const includeLinearMetrics = !!(feature.configDraft?.includeLinearMetrics ?? feature.configCommitted?.includeLinearMetrics);
    const selectedText = selected.length ? selected.join(', ') : '선택한 파라미터가 없습니다.';
    const optionParts = [];
    if (excludeEndDummy) optionParts.push('End+Dummy 제외');
    if (includePointXY) optionParts.push('좌표 X/Y');
    if (includeLinearMetrics) optionParts.push('선형 길이/방향');
    const optionText = optionParts.length ? ` · 옵션 ${optionParts.join(', ')}` : '';
    target.top.textContent = feature.enabled
      ? '선택 완료 · 설정 창에서 검토 파라미터를 선택하거나 수정할 수 있습니다.'
      : '필요할 때만 켠 뒤 설정 창에서 공유 파라미터 검토 대상을 선택해 주세요.';
    target.sub.textContent = `선택 파라미터 ${selected.length}개 · ${selectedText} · 허용 범위 ${tol} ${unit}${optionText}`;
    target.row.classList.toggle('is-active', !!feature.enabled);
    applyFeatureRowTooltip(target.row, [
      FEATURE_META.connector?.label || '파라미터 값 연속성 검토',
      FEATURE_META.connector?.desc || '',
      target.top.textContent,
      target.sub.textContent
    ], {
      title: FEATURE_META.connector?.label || '파라미터 값 연속성 검토',
      desc: `${FEATURE_META.connector?.desc || ''} ${target.sub.textContent}`.trim()
    });
  }

  function refreshFloorInfoFeatureSummary() {
    const target = state.ui.floorInfoSummary;
    if (!target) return;
    const feature = state.features.floorinfo;
    const committedRules = normalizeFloorInfoRules(feature.configCommitted?.levelRules);
    const draftRules = normalizeFloorInfoRules(feature.configDraft?.levelRules);
    const rules = draftRules.length ? draftRules : committedRules;
    const selectedRules = rules.filter((rule) => rule.useAsBoundary !== false);
    const configuredCount = selectedRules.filter((rule) => String(rule.expectedValue || '').trim()).length;
    const parameterName = feature.configDraft?.parameterName || feature.configCommitted?.parameterName || '입력 필요';
    const common = state.common.configCommitted || {};
    const commonExtraCount = String(common.extraParams || '')
      .split(',')
      .map((value) => value.trim())
      .filter(Boolean)
      .length;
    const hasCommonScope = !!String(common.targetFilter || '').trim() || !!String(common.excludeTargetFilter || '').trim();
    target.top.textContent = feature.enabled
      ? '선택 완료 · 설정 창에서 영역 기준 레벨과 기대 층정보 값을 관리할 수 있습니다.'
      : '보조 기능입니다. 필요할 때만 켜서 층정보 파라미터의 레벨/Z 일치 여부를 검토해 주세요.';
    target.sub.textContent = `파라미터 ${parameterName} · 영역 기준 ${selectedRules.length || 0}개 · 규칙 ${configuredCount}/${selectedRules.length || 0}개 · ${hasCommonScope ? '공통 필터 적용됨' : '공통 필터가 없습니다.'} · 추가 추출 ${commonExtraCount}개`;
    target.row.classList.toggle('is-active', !!feature.enabled);
    applyFeatureRowTooltip(target.row, [
      FEATURE_META.floorinfo?.label || '층/레벨 정보 검토',
      FEATURE_META.floorinfo?.desc || '',
      target.top.textContent,
      target.sub.textContent
    ], {
      title: FEATURE_META.floorinfo?.label || '층/레벨 정보 검토',
      desc: `${FEATURE_META.floorinfo?.desc || ''} ${target.sub.textContent}`.trim()
    });
  }

  function refreshFamilySuitabilityFeatureSummary() {
    const target = state.ui.familySuitabilitySummary;
    if (!target) return;
    const feature = state.features.familysuitability;
    const draft = feature.configDraft || {};
    const committed = feature.configCommitted || {};
    const criteriaPath = String(draft.criteriaExcelPath || committed.criteriaExcelPath || '').trim();
    const matchReview = String(draft.matchReviewText || committed.matchReviewText || '').trim();
    const mismatchReview = String(draft.mismatchReviewText || committed.mismatchReviewText || '').trim();
    const comboCount = Number(draft.criteriaComboCount ?? committed.criteriaComboCount) || 0;
    const rowCount = Number(draft.criteriaRowCount ?? committed.criteriaRowCount) || 0;
    const filterRules = normalizeFamilySuitabilityFilterRules(
      draft.filterRules && draft.filterRules.length ? draft.filterRules : committed.filterRules,
      { keepEmpty: true }
    );
    const activeFilterCount = filterRules.filter((rule) => String(rule.keyword || '').trim() && String(rule.reviewText || '').trim()).length;
    const invalidFilter = filterRules.some((rule) => {
      const keyword = String(rule.keyword || '').trim();
      const reviewText = String(rule.reviewText || '').trim();
      return (!!keyword || !!reviewText) && (!keyword || !reviewText);
    });
    const criteriaLabel = getFamilySuitabilityCriteriaLabel(criteriaPath);
    if (!feature.enabled) {
      target.top.textContent = '보조 기능입니다. 필요할 때만 켜서 실제 사용 패밀리/타입 조합의 기준 적합성을 검토해 주세요.';
    } else if (!criteriaPath || !matchReview || !mismatchReview || invalidFilter) {
      target.top.textContent = '선택 완료 · 기준 엑셀, 검토 문구, 필터 규칙을 마저 확인하면 바로 실행할 수 있습니다.';
    } else {
      target.top.textContent = '선택 완료 · 실제 사용된 객체만 집계해 기준 일치, 미일치, 필터 우선 검토 문구를 출력합니다.';
    }

    const reviewState = `${matchReview ? '일치 문구 설정' : '일치 문구 필요'} / ${mismatchReview ? '미일치 문구 설정' : '미일치 문구 필요'}`;
    const filterState = invalidFilter
      ? `필터 ${activeFilterCount}개 + 미완성 규칙 있음`
      : `필터 ${activeFilterCount}개`;
    const basisState = comboCount
      ? `기준 ${comboCount}조합`
      : rowCount
        ? `원본 ${rowCount}행`
        : '기준 정보가 없습니다.';
    target.sub.textContent = `${criteriaLabel} · ${basisState} · ${filterState} · ${reviewState}`;
    target.row.classList.toggle('is-active', !!feature.enabled);
    applyFeatureRowTooltip(target.row, [
      FEATURE_META.familysuitability?.label || '패밀리 타입 적합성 검토',
      FEATURE_META.familysuitability?.desc || '',
      target.top.textContent,
      target.sub.textContent
    ], {
      title: FEATURE_META.familysuitability?.label || '패밀리 타입 적합성 검토',
      desc: `${FEATURE_META.familysuitability?.desc || ''} ${target.sub.textContent}`.trim()
    });
  }

  function refreshTapAlignFeatureSummary() {
    const target = state.ui.tapAlignSummary;
    if (!target) return;
    const feature = state.features.tapalign;
    const draft = feature.configDraft || {};
    const committed = feature.configCommitted || {};
    const tolRaw = draft.tol ?? committed.tol ?? 0.5;
    const tol = Number.isFinite(Number(tolRaw)) ? Number(tolRaw) : 0.5;
    const unit = normalizeTapAlignUnit(draft.unit || committed.unit || 'mm');
    const domain = normalizeTapAlignDomain(draft.domain || committed.domain || 'all');
    const featureFilterText = String(draft.featureTargetFilter || committed.featureTargetFilter || '').trim();
    const common = state.common.configCommitted || {};
    const extraCount = String(common.extraParams || '')
      .split(',')
      .map((value) => value.trim())
      .filter(Boolean)
      .length;
    const optionParts = [];
    if (common.includePointXY) optionParts.push('좌표 X/Y');
    if (common.includeLinearMetrics) optionParts.push('선형 길이/방향');
    const optionText = optionParts.length ? ` · 공통 옵션 ${optionParts.join(', ')}` : '';
    const featureFilterNote = featureFilterText ? ' · 기능 필터 적용' : '';
    target.top.textContent = feature.enabled
      ? '선택 완료 · 연결 라인 중심축과 탭/분기 피팅 축의 정렬 상태를 검토합니다.'
      : '보조 기능입니다. 필요할 때만 켜서 탭/분기 피팅 축 이탈 여부를 검토해 주세요.';
    target.sub.textContent = `허용 범위 ${tol} ${unit} · 범위 ${resolveTapAlignDomainLabel(domain)} · 추가 추출 ${extraCount}개${optionText}${featureFilterNote}`;
    target.row.classList.toggle('is-active', !!feature.enabled);
    applyFeatureRowTooltip(target.row, [
      FEATURE_META.tapalign?.label || '탭/분기 축 틀어짐 검토',
      FEATURE_META.tapalign?.desc || '',
      target.top.textContent,
      target.sub.textContent
    ], {
      title: FEATURE_META.tapalign?.label || '탭/분기 축 틀어짐 검토',
      desc: `${FEATURE_META.tapalign?.desc || ''} ${target.sub.textContent}`.trim()
    });
  }

  function refreshDupClashFeatureSummary() {
    const target = state.ui.dupClashSummary;
    if (!target) return;
    const feature = state.features.dupclash;
    const draft = feature.configDraft || {};
    const committed = feature.configCommitted || {};
    const mode = normalizeDupClashMode(draft.mode || committed.mode || 'duplicate');
    const modeLabel = resolveDupClashModeLabel(mode);
    const common = state.common.configCommitted || {};
    const extraCount = String(common.extraParams || '')
      .split(',')
      .map((value) => value.trim())
      .filter(Boolean)
      .length;
    const filterText = String(common.targetFilter || '').trim() || '필터가 없습니다.';
    const excludeText = String(common.excludeTargetFilter || '').trim() || '제외 필터가 없습니다.';

    target.top.textContent = feature.enabled
      ? '선택 완료 · 설정 창에서 중복 검토와 자체 간섭 검토 중 하나를 선택해 실행합니다.'
      : '보조 기능입니다. 필요할 때만 켠 뒤 설정 창에서 검토 모드를 선택해 주세요.';
    target.sub.textContent = `선택 모드 ${modeLabel} · 포함 필터 ${filterText} · 제외 필터 ${excludeText} · 추가 파라미터 ${extraCount}개`;
    target.row.classList.toggle('is-active', !!feature.enabled);
    applyFeatureRowTooltip(target.row, [
      FEATURE_META.dupclash?.label || '중복 / 자체 간섭 검토',
      FEATURE_META.dupclash?.desc || '',
      target.top.textContent,
      target.sub.textContent
    ], {
      title: FEATURE_META.dupclash?.label || '중복 / 자체 간섭 검토',
      desc: `${FEATURE_META.dupclash?.desc || ''} ${target.sub.textContent}`.trim()
    });
  }

  function refreshWorksetAssignmentFeatureSummary() {
    const target = state.ui.worksetAssignmentSummary;
    if (!target) return;
    const feature = state.features.worksetassignment;
    const draft = feature.configDraft || {};
    const committed = feature.configCommitted || {};
    const expectedWorksetName = String(draft.expectedWorksetName || committed.expectedWorksetName || 'Workset1').trim() || 'Workset1';
    const flaggedWorksetName = String(draft.flaggedWorksetName || committed.flaggedWorksetName || '').trim();
    const common = state.common.configCommitted || {};
    const commonExtraCount = String(common.extraParams || '')
      .split(',')
      .map((value) => value.trim())
      .filter(Boolean)
      .length;
    const hasCommonScope = !!String(common.targetFilter || '').trim() || !!String(common.excludeTargetFilter || '').trim();

    if (!feature.enabled) {
      target.top.textContent = '보조 기능입니다. 필요할 때만 켜서 기본 웍셋(Workset1) 기준 또는 특정 웍셋 존재 여부를 검토해 주세요.';
    } else if (flaggedWorksetName) {
      target.top.textContent = '선택 완료 · 입력한 웍셋에 속한 객체만 오류로 검토합니다.';
    } else {
      target.top.textContent = '선택 완료 · 기본 웍셋(Workset1) 이외의 웍셋에 속한 객체를 오류로 검토합니다.';
    }

    target.sub.textContent = flaggedWorksetName
      ? `기준 ${expectedWorksetName} · 오류 대상 ${flaggedWorksetName} · ${hasCommonScope ? '공통 필터 적용됨' : '공통 필터가 없습니다.'} · 추가 추출 ${commonExtraCount}개`
      : `기준 ${expectedWorksetName} · 입력 대상이 없습니다. 기본 웍셋 외 전체를 오류로 검토합니다. · ${hasCommonScope ? '공통 필터 적용됨' : '공통 필터가 없습니다.'} · 추가 추출 ${commonExtraCount}개`;
    target.row.classList.toggle('is-active', !!feature.enabled);
    applyFeatureRowTooltip(target.row, [
      FEATURE_META.worksetassignment?.label || '웍셋 배정 검토',
      FEATURE_META.worksetassignment?.desc || '',
      target.top.textContent,
      target.sub.textContent
    ], {
      title: FEATURE_META.worksetassignment?.label || '웍셋 배정 검토',
      desc: `${FEATURE_META.worksetassignment?.desc || ''} ${target.sub.textContent}`.trim()
    });
  }

  function refreshParameterDuplicationFeatureSummary() {
    const target = state.ui.parameterDuplicationSummary;
    if (!target) return;
    const feature = state.features.parameterduplication;
    const draft = feature.configDraft || {};
    const committed = feature.configCommitted || {};
    const scope = normalizeParameterDuplicationScope(draft.scope || committed.scope || 'all');
    const names = Array.isArray(draft.parameterNames) && draft.parameterNames.length
      ? draft.parameterNames
      : Array.isArray(committed.parameterNames)
        ? committed.parameterNames
        : [];
    const sourcePath = String(draft.sharedParamSourcePath || committed.sharedParamSourcePath || '').trim();
    const scopeLabel = scope === 'selected' ? '지정 파라미터만' : '추가된 전체 프로젝트 파라미터';
    const selectedText = buildParameterDuplicationNamePreview(names);
    const sourceText = sourcePath ? ` · 공유파라미터 ${getPathLeafLabel(sourcePath, sourcePath)}` : '';

    if (!feature.enabled) {
      target.top.textContent = '보조 기능입니다. 필요할 때만 켜서 프로젝트 파라미터 이름 중복 여부를 검토해 주세요.';
    } else if (scope === 'selected' && !names.length) {
      target.top.textContent = '선택 완료 · 지정 파라미터만 검토하려면 공유파라미터 목록에서 대상 파라미터를 1개 이상 선택해 주세요.';
    } else {
      target.top.textContent = '선택 완료 · 프로젝트 파라미터 이름 중복 여부를 BQC 포맷으로 정리합니다.';
    }

    target.sub.textContent = scope === 'selected'
      ? `검토 범위 ${scopeLabel} · 대상 ${names.length}개${sourceText} · ${selectedText}`
      : `검토 범위 ${scopeLabel} · 문서에 추가된 모든 프로젝트 파라미터를 검토합니다.`;
    target.row.classList.toggle('is-active', !!feature.enabled);
    applyFeatureRowTooltip(target.row, [
      FEATURE_META.parameterduplication?.label || '프로젝트 파라미터 중복 검토',
      FEATURE_META.parameterduplication?.desc || '',
      target.top.textContent,
      target.sub.textContent
    ], {
      title: FEATURE_META.parameterduplication?.label || '프로젝트 파라미터 중복 검토',
      desc: `${FEATURE_META.parameterduplication?.desc || ''} ${target.sub.textContent}`.trim()
    });
  }

  function refreshParameterMissingFeatureSummary() {
    const target = state.ui.parameterMissingSummary;
    if (!target) return;
    const feature = state.features.parametermissing;
    const config = createParameterMissingConfigSnapshot(feature.configDraft || feature.configCommitted || {});
    const parameterCount = config.parameterNames.length;
    const exceptionRuleCount = (Array.isArray(config.exceptionRules) ? config.exceptionRules : [])
      .filter((rule) => countParameterMissingConfiguredConditions(rule.conditions) > 0)
      .length;
    const targetFilterCount = countParameterMissingTargetFilterConditions(config);
    const hasIncomplete = hasIncompleteParameterMissingConfig(config);
    const common = state.common.configCommitted || {};
    const commonExtraCount = String(common.extraParams || '')
      .split(',')
      .map((value) => value.trim())
      .filter(Boolean)
      .length;
    const hasCommonScope = !!String(common.targetFilter || '').trim() || !!String(common.excludeTargetFilter || '').trim();

    if (!feature.enabled) {
      target.top.textContent = '보조 기능입니다. 필요한 경우에만 켜서 지정한 공유 텍스트 파라미터의 누락을 검토해 주세요.';
    } else if (!parameterCount) {
      target.top.textContent = '선택 완료 · 공유파라미터 목록에서 누락 검토할 파라미터를 1개 이상 선택해 주세요.';
    } else if (hasIncomplete) {
      target.top.textContent = '선택 완료 · 객체 필터 또는 누락 예외 규칙의 미완성 항목을 확인해 주세요.';
    } else {
      target.top.textContent = targetFilterCount
        ? '선택 완료 · 공통 필터와 기능 전용 객체 필터 기준으로 지정 파라미터의 누락 여부를 BQC 포맷으로 정리합니다.'
        : '선택 완료 · 공통 검토대상 필터 기준으로 지정 파라미터의 누락 여부를 BQC 포맷으로 정리합니다.';
    }

    target.sub.textContent = `파라미터 ${parameterCount}개 · 객체 필터 ${targetFilterCount ? `${targetFilterCount}개` : '없음'} · 누락 예외 ${exceptionRuleCount}개 · ${hasCommonScope ? '공통 필터 적용됨' : '공통 필터가 없습니다.'} · 추가 추출 ${commonExtraCount}개 · ${buildParameterMissingSelectionPreview(config.parameterNames, 5)}`;
    target.row.classList.toggle('is-active', !!feature.enabled);
    applyFeatureRowTooltip(target.row, [
      FEATURE_META.parametermissing?.label || '파라미터 누락 검토',
      FEATURE_META.parametermissing?.desc || '',
      target.top.textContent,
      target.sub.textContent
    ], {
      title: FEATURE_META.parametermissing?.label || '파라미터 누락 검토',
      desc: `${FEATURE_META.parametermissing?.desc || ''} ${target.sub.textContent}`.trim()
    });
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
      nameMain.setAttribute('aria-label', FEATURE_META[key]?.label || key);
      const nameSub = document.createElement('span');
      nameSub.textContent = FEATURE_META[key]?.desc || '';
      if (FEATURE_META[key]?.desc) {
        nameSub.setAttribute('aria-label', FEATURE_META[key].desc);
      }
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
    updateCurrentDocBtnState();
    updateMultiRunBtnState();
  }

  function getSelectedFeatureStatus(key) {
    const feature = state.features[key];
    if (!feature) return { label: '검토 전', className: 'status-chip--idle' };
    if (requiresSharedParams(key) && state.sharedParamStatus?.status && state.sharedParamStatus.status !== 'ok') {
      return { label: '공유파라미터 확인 필요', className: 'status-chip--warn' };
    }
    if (key === 'familylink') {
      const targets = feature.configCommitted.selectedTargets || [];
      if (!targets.length) {
        return { label: '설정 필요', className: 'status-chip--warn' };
      }
    }
    if (key === 'floorinfo') {
      const rules = normalizeFloorInfoRules(feature.configCommitted.levelRules);
      const selectedRules = rules.filter((rule) => rule.useAsBoundary !== false);
      const configuredCount = selectedRules.filter((rule) => String(rule.expectedValue || '').trim()).length;
      if (!String(feature.configCommitted.parameterName || '').trim() || !selectedRules.length || configuredCount < selectedRules.length) {
        return { label: '설정 필요', className: 'status-chip--warn' };
      }
    }
    if (key === 'familysuitability') {
      const config = feature.configCommitted || {};
      const filterRules = normalizeFamilySuitabilityFilterRules(config.filterRules, { keepEmpty: true });
      const invalidFilter = filterRules.find((rule) => {
        const keyword = String(rule.keyword || '').trim();
        const reviewText = String(rule.reviewText || '').trim();
        return (!!keyword || !!reviewText) && (!keyword || !reviewText);
      });
      if (!String(config.criteriaExcelPath || '').trim()
          || !String(config.matchReviewText || '').trim()
          || !String(config.mismatchReviewText || '').trim()
          || invalidFilter) {
        return { label: '설정 필요', className: 'status-chip--warn' };
      }
    }
    if (key === 'parameterduplication') {
      const config = feature.configCommitted || {};
      const scope = normalizeParameterDuplicationScope(config.scope);
      const names = Array.isArray(config.parameterNames) ? config.parameterNames : [];
      if (scope === 'selected' && !names.length) {
        return { label: '설정 필요', className: 'status-chip--warn' };
      }
    }
    if (key === 'parametermissing') {
      const config = createParameterMissingConfigSnapshot(feature.configCommitted || {});
      if (!config.parameterNames.length || hasIncompleteParameterMissingConfig(config)) {
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
    return { label: '검토 준비 완료', className: 'status-chip--ready' };
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
      pathValue.setAttribute('aria-label', pathText);
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
      if (key === 'parameterduplication') {
        requestSharedParamStatus('parameterduplication-settings');
        requestSharedParamList('parameterduplication-settings');
      }
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
        const warning = state.sharedParamStatus?.warning || '공유파라미터 미등록으로 실행이 제한됩니다.';
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
      toast('공유파라미터 상태를 확인 중입니다.', 'warn');
      return false;
    }
    requestSharedParamStatus('run');
    const msg = status.warning || status.errorMessage || '공유파라미터 상태를 확인해 주세요.';
    toast(msg, 'warn');
    return false;
  }

  function normalizeCommonOptions(raw) {
    return {
      extraParams: typeof raw?.extraParamsText === 'string'
        ? raw.extraParamsText
        : (typeof raw?.extraParams === 'string' ? raw.extraParams : ''),
      targetFilter: typeof raw?.targetFilterText === 'string'
        ? raw.targetFilterText
        : (typeof raw?.targetFilter === 'string' ? raw.targetFilter : ''),
      excludeTargetFilter: typeof raw?.excludeTargetFilterText === 'string'
        ? raw.excludeTargetFilterText
        : (typeof raw?.excludeTargetFilter === 'string' ? raw.excludeTargetFilter : ''),
      excludeEndDummy: false,
      includePointXY: !!raw?.includePointXY,
      includeLinearMetrics: !!raw?.includeLinearMetrics
    };
  }

  function normalizeTapAlignUnit(value) {
    const normalized = String(value || 'mm').trim().toLowerCase();
    return normalized === 'inch' || normalized === 'in' || normalized === 'inches' ? 'inch' : 'mm';
  }

  function normalizeDupClashMode(value) {
    const normalized = String(value || 'duplicate').trim().toLowerCase();
    return normalized === 'clash' || normalized === 'selfclash' || normalized === 'self-clash'
      ? 'clash'
      : 'duplicate';
  }

  function resolveDupClashModeLabel(value) {
    return normalizeDupClashMode(value) === 'clash' ? '자체 간섭 검토' : '중복 검토';
  }

  function normalizeTapAlignDomain(value) {
    const normalized = String(value || 'all').trim().toLowerCase();
    if (normalized === 'pipe' || normalized === 'piping') return 'pipe';
    if (normalized === 'duct' || normalized === 'hvac') return 'duct';
    return 'all';
  }

  function normalizeTapAlignExportLocale(value) {
    const normalized = String(value || 'ko').trim().toLowerCase();
    return normalized === 'en' || normalized === 'eng' || normalized === 'english' ? 'en' : 'ko';
  }

  function resolveTapAlignDomainLabel(value) {
    const normalized = normalizeTapAlignDomain(value);
    if (normalized === 'pipe') return '배관';
    if (normalized === 'duct') return '덕트';
    return '배관 + 덕트';
  }

  function resolveTapAlignLocaleLabel(value) {
    return normalizeTapAlignExportLocale(value) === 'en' ? '영문' : '한글';
  }

  function loadTapAlignConfigFromStorage() {
    const defaults = {
      tol: 0.5,
      unit: 'mm',
      domain: 'all',
      featureTargetFilter: ''
    };

    let stored = null;
    try {
      const raw = localStorage.getItem(TAPALIGN_STORAGE_KEY);
      if (raw) stored = JSON.parse(raw);
    } catch {
      stored = null;
    }

    const config = {
      tol: Math.max(0, parseFloat(stored?.tol) || defaults.tol),
      unit: normalizeTapAlignUnit(stored?.unit || defaults.unit),
      domain: normalizeTapAlignDomain(stored?.domain || defaults.domain),
      featureTargetFilter: typeof stored?.featureTargetFilter === 'string' ? stored.featureTargetFilter : defaults.featureTargetFilter
    };

    return {
      config,
      hasStored: !!stored
    };
  }

  function persistTapAlignConfig(committed) {
    const normalized = {
      tol: Math.max(0, parseFloat(committed?.tol) || 0.5),
      unit: normalizeTapAlignUnit(committed?.unit || 'mm'),
      domain: normalizeTapAlignDomain(committed?.domain || 'all'),
      featureTargetFilter: String(committed?.featureTargetFilter || '').trim()
    };

    let payload = {};
    try {
      const raw = localStorage.getItem(TAPALIGN_STORAGE_KEY);
      if (raw) payload = JSON.parse(raw) || {};
    } catch {
      payload = {};
    }

    try {
      localStorage.setItem(TAPALIGN_STORAGE_KEY, JSON.stringify({ ...payload, ...normalized }));
    } catch {
    }
  }

  function loadFamilySuitabilityConfigFromStorage() {
    const defaults = {
      criteriaExcelPath: '',
      criteriaRowCount: 0,
      criteriaComboCount: 0,
      criteriaSheetCount: 0,
      matchReviewText: '',
      mismatchReviewText: '',
      filterRules: []
    };

    let stored = null;
    try {
      const raw = localStorage.getItem(FAMILY_SUITABILITY_STORAGE_KEY);
      if (raw) stored = JSON.parse(raw);
    } catch {
      stored = null;
    }

    const config = normalizeFamilySuitabilityConfig(stored || defaults, { keepEmptyFilters: true });
    if (!config.filterRules.length) {
      config.filterRules = [createEmptyFamilySuitabilityFilterRule()];
    }

    return {
      config,
      hasStored: !!stored
    };
  }

  function persistFamilySuitabilityConfig(committed) {
    try {
      localStorage.setItem(
        FAMILY_SUITABILITY_STORAGE_KEY,
        JSON.stringify(normalizeFamilySuitabilityConfig(committed))
      );
    } catch {
    }
  }

  function loadParameterMissingConfigFromStorage() {
    let stored = null;
    try {
      const raw = localStorage.getItem(PARAMETER_MISSING_STORAGE_KEY);
      if (raw) stored = JSON.parse(raw);
    } catch {
      stored = null;
    }

    const config = createParameterMissingConfigSnapshot(stored || {});
    return {
      config,
      hasStored: !!stored && config.parameterNames.length > 0
    };
  }

  function persistParameterMissingConfig(committed) {
    try {
      localStorage.setItem(
        PARAMETER_MISSING_STORAGE_KEY,
        JSON.stringify(createParameterMissingSerializableConfig(committed))
      );
    } catch {
    }
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
      excludeTargetFilterText: committed.excludeTargetFilter || '',
      excludeEndDummy: false,
      includePointXY: !!committed.includePointXY,
      includeLinearMetrics: !!committed.includeLinearMetrics
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
    if (key === 'floorinfo') {
      const committed = deepCopy(feature.configCommitted || {});
      committed.levelRules = normalizeFloorInfoRules(committed.levelRules);
      committed.parameterName = String(committed.parameterName || '').trim();
      delete committed.baseLevelName;
      return {
        enabled: feature.enabled,
        ...committed
      };
    }
    if (key === 'familysuitability') {
      const committed = deepCopy(feature.configCommitted || {});
      committed.criteriaExcelPath = String(committed.criteriaExcelPath || '').trim();
      committed.criteriaRowCount = Number(committed.criteriaRowCount) || 0;
      committed.criteriaComboCount = Number(committed.criteriaComboCount) || 0;
      committed.criteriaSheetCount = Number(committed.criteriaSheetCount) || 0;
      committed.matchReviewText = String(committed.matchReviewText || '').trim();
      committed.mismatchReviewText = String(committed.mismatchReviewText || '').trim();
      committed.filterRules = normalizeFamilySuitabilityFilterRules(committed.filterRules);
      return {
        enabled: feature.enabled,
        ...committed
      };
    }
    if (key === 'parameterduplication') {
      const committed = deepCopy(feature.configCommitted || {});
      committed.scope = normalizeParameterDuplicationScope(committed.scope);
      committed.parameterNames = parseParameterDuplicationNames(Array.isArray(committed.parameterNames)
        ? committed.parameterNames.join('\n')
        : committed.parameterNames);
      committed.sharedParamSourcePath = String(committed.sharedParamSourcePath || '').trim();
      committed.sharedParamImportCount = Number(committed.sharedParamImportCount) || 0;
      return {
        enabled: feature.enabled,
        ...committed
      };
    }
    if (key === 'parametermissing') {
      return {
        enabled: feature.enabled,
        ...createParameterMissingConfigSnapshot(feature.configCommitted || {})
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
    if (key === 'tapalign') {
      const committed = deepCopy(feature.configCommitted || {});
      committed.tol = Math.max(0, parseFloat(committed.tol) || 0.5);
      committed.unit = normalizeTapAlignUnit(committed.unit);
      committed.domain = normalizeTapAlignDomain(committed.domain);
      committed.featureTargetFilter = String(committed.featureTargetFilter || '').trim();
      return {
        enabled: feature.enabled,
        ...committed
      };
    }
    if (key === 'dupclash') {
      const committed = deepCopy(feature.configCommitted || {});
      committed.mode = normalizeDupClashMode(committed.mode);
      committed.tolFeet = Number(committed.tolFeet) > 0 ? Number(committed.tolFeet) : 1 / 64;
      return {
        enabled: feature.enabled,
        ...committed
      };
    }
    if (key === 'worksetassignment') {
      const committed = deepCopy(feature.configCommitted || {});
      committed.expectedWorksetName = String(committed.expectedWorksetName || 'Workset1').trim() || 'Workset1';
      committed.flaggedWorksetName = String(committed.flaggedWorksetName || '').trim();
      return {
        enabled: feature.enabled,
        ...committed
      };
    }
    if (key === 'guid') {
      const committed = deepCopy(feature.configCommitted || {});
      committed.closeAllWorksetsOnOpen = true;
      committed.syncComment = String(committed.syncComment || '');
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
    if (state.ui.activeFeatureKey === 'tapalign') {
      target.configDraft.tol = Math.max(0, parseFloat(target.configDraft.tol) || 0.5);
      target.configDraft.unit = normalizeTapAlignUnit(target.configDraft.unit);
      target.configDraft.domain = normalizeTapAlignDomain(target.configDraft.domain);
    }
    if (state.ui.activeFeatureKey === 'dupclash') {
      target.configDraft.mode = normalizeDupClashMode(target.configDraft.mode);
      target.configDraft.tolFeet = Number(target.configDraft.tolFeet) > 0 ? Number(target.configDraft.tolFeet) : 1 / 64;
    }
    if (state.ui.activeFeatureKey === 'floorinfo') {
      target.configDraft.parameterName = String(target.configDraft.parameterName || '').trim();
      target.configDraft.levelRules = normalizeFloorInfoRules(target.configDraft.levelRules);
      delete target.configDraft.baseLevelName;
    }
    if (state.ui.activeFeatureKey === 'familysuitability') {
      target.configDraft.criteriaExcelPath = String(target.configDraft.criteriaExcelPath || '').trim();
      target.configDraft.criteriaRowCount = Number(target.configDraft.criteriaRowCount) || 0;
      target.configDraft.criteriaComboCount = Number(target.configDraft.criteriaComboCount) || 0;
      target.configDraft.criteriaSheetCount = Number(target.configDraft.criteriaSheetCount) || 0;
      target.configDraft.matchReviewText = String(target.configDraft.matchReviewText || '').trim();
      target.configDraft.mismatchReviewText = String(target.configDraft.mismatchReviewText || '').trim();
      target.configDraft.filterRules = normalizeFamilySuitabilityFilterRules(target.configDraft.filterRules);
    }
    if (state.ui.activeFeatureKey === 'parameterduplication') {
      Object.assign(target.configDraft, createParameterDuplicationConfigSnapshot(target.configDraft));
    }
    if (state.ui.activeFeatureKey === 'parametermissing') {
      if (state.ui.modalOpen && typeof state.ui.controls.parametermissing?.syncDraftFromControls === 'function') {
        state.ui.controls.parametermissing.syncDraftFromControls();
      }
      Object.assign(target.configDraft, createParameterMissingConfigSnapshot(target.configDraft));
    }
    if (state.ui.activeFeatureKey === 'worksetassignment') {
      target.configDraft.expectedWorksetName = String(target.configDraft.expectedWorksetName || 'Workset1').trim() || 'Workset1';
      target.configDraft.flaggedWorksetName = String(target.configDraft.flaggedWorksetName || '').trim();
    }
    if (state.ui.activeFeatureKey === 'guid') {
      target.configDraft.closeAllWorksetsOnOpen = true;
      target.configDraft.syncComment = String(target.configDraft.syncComment || '').trim();
    }
    target.configCommitted = deepCopy(target.configDraft);
    if (state.ui.activeFeatureKey === 'parameterduplication') {
      rememberParameterDuplicationRecent(target.configCommitted);
    }
    if (state.ui.activeFeatureKey === 'parametermissing') {
      persistParameterMissingConfig(target.configCommitted);
      rememberParameterMissingRecent(target.configCommitted);
    }
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
    if (key === 'floorinfo') {
      feature.configDraft.levelRules = normalizeFloorInfoRules(feature.configDraft.levelRules);
      feature.configDraft.parameterName = String(feature.configDraft.parameterName || '').trim();
      delete feature.configDraft.baseLevelName;
    }
    if (key === 'familysuitability') {
      feature.configDraft.criteriaExcelPath = String(feature.configDraft.criteriaExcelPath || '').trim();
      feature.configDraft.criteriaRowCount = Number(feature.configDraft.criteriaRowCount) || 0;
      feature.configDraft.criteriaComboCount = Number(feature.configDraft.criteriaComboCount) || 0;
      feature.configDraft.criteriaSheetCount = Number(feature.configDraft.criteriaSheetCount) || 0;
      feature.configDraft.matchReviewText = String(feature.configDraft.matchReviewText || '').trim();
      feature.configDraft.mismatchReviewText = String(feature.configDraft.mismatchReviewText || '').trim();
      feature.configDraft.filterRules = normalizeFamilySuitabilityFilterRules(feature.configDraft.filterRules, { keepEmpty: true });
      if (!feature.configDraft.filterRules.length) {
        feature.configDraft.filterRules = [createEmptyFamilySuitabilityFilterRule()];
      }
    }
    if (key === 'parameterduplication') {
      feature.configDraft.scope = normalizeParameterDuplicationScope(feature.configDraft.scope);
      feature.configDraft.parameterNames = parseParameterDuplicationNames(Array.isArray(feature.configDraft.parameterNames)
        ? feature.configDraft.parameterNames.join('\n')
        : feature.configDraft.parameterNames);
      feature.configDraft.sharedParamSourcePath = String(feature.configDraft.sharedParamSourcePath || '').trim();
      feature.configDraft.sharedParamImportCount = Number(feature.configDraft.sharedParamImportCount) || 0;
    }
    if (key === 'parametermissing') {
      Object.assign(feature.configDraft, createParameterMissingConfigSnapshot(feature.configDraft));
    }
    if (key === 'worksetassignment') {
      feature.configDraft.expectedWorksetName = String(feature.configDraft.expectedWorksetName || 'Workset1').trim() || 'Workset1';
      feature.configDraft.flaggedWorksetName = String(feature.configDraft.flaggedWorksetName || '').trim();
    }
    if (key === 'dupclash') {
      feature.configDraft.mode = normalizeDupClashMode(feature.configDraft.mode);
      feature.configDraft.tolFeet = Number(feature.configDraft.tolFeet) > 0 ? Number(feature.configDraft.tolFeet) : 1 / 64;
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

  function normalizeCommonExtraParamsSignature(value) {
    return String(value || '')
      .split(',')
      .map((item) => item.trim())
      .filter(Boolean)
      .join('|');
  }

  function collectCommonDependentFeatureKeys(previousCommon = {}, nextCommon = {}) {
    const keys = new Set();
    const previousTargetFilter = String(previousCommon.targetFilter || '').trim();
    const nextTargetFilter = String(nextCommon.targetFilter || '').trim();
    const previousExcludeTargetFilter = String(previousCommon.excludeTargetFilter || '').trim();
    const nextExcludeTargetFilter = String(nextCommon.excludeTargetFilter || '').trim();
    const previousExtraParams = normalizeCommonExtraParamsSignature(previousCommon.extraParams);
    const nextExtraParams = normalizeCommonExtraParamsSignature(nextCommon.extraParams);
    const previousPointXY = !!previousCommon.includePointXY;
    const nextPointXY = !!nextCommon.includePointXY;
    const previousLinearMetrics = !!previousCommon.includeLinearMetrics;
    const nextLinearMetrics = !!nextCommon.includeLinearMetrics;

    if (previousTargetFilter !== nextTargetFilter || previousExcludeTargetFilter !== nextExcludeTargetFilter) {
      COMMON_SCOPE_DEPENDENT_FEATURE_KEYS.forEach((key) => keys.add(key));
    }
    if (previousExtraParams !== nextExtraParams) {
      COMMON_EXTRA_PARAM_DEPENDENT_FEATURE_KEYS.forEach((key) => keys.add(key));
    }
    if (previousPointXY !== nextPointXY || previousLinearMetrics !== nextLinearMetrics) {
      COMMON_TAPALIGN_OPTION_DEPENDENT_FEATURE_KEYS.forEach((key) => keys.add(key));
    }

    return Array.from(keys);
  }

  function markCommonDependentFeaturesStale(previousCommon = {}, nextCommon = {}) {
    collectCommonDependentFeatureKeys(previousCommon, nextCommon).forEach(markStale);
  }

  function markCommonDirty() {
    state.common.dirty = true;
    state.common.applied = false;
    markCommonDependentFeaturesStale(state.common.configCommitted || {}, state.common.configDraft || {});
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
      if (controls.excludeEndDummy?.input) controls.excludeEndDummy.input.checked = !!draft.excludeEndDummy;
      if (controls.pointXY?.input) controls.pointXY.input.checked = !!draft.includePointXY;
      if (controls.linearMetrics?.input) controls.linearMetrics.input.checked = !!draft.includeLinearMetrics;
      if (controls.searchInput) controls.searchInput.value = '';
      if (controls.renderConnectorList) controls.renderConnectorList();
      if (controls.renderConnectorSelected) controls.renderConnectorSelected();
    } else if (key === 'tapalign') {
      const draft = state.features.tapalign.configDraft;
      controls.tol.input.value = draft.tol;
      controls.unit.select.value = normalizeTapAlignUnit(draft.unit);
      controls.domain.select.value = normalizeTapAlignDomain(draft.domain);
      if (controls.featureTargetFilter?.input) controls.featureTargetFilter.input.value = draft.featureTargetFilter || '';
      if (controls.renderCommonSummary) controls.renderCommonSummary();
    } else if (key === 'dupclash') {
      const draft = state.features.dupclash.configDraft;
      const mode = normalizeDupClashMode(draft.mode);
      if (controls.modeDuplicate?.input) controls.modeDuplicate.input.checked = mode === 'duplicate';
      if (controls.modeClash?.input) controls.modeClash.input.checked = mode === 'clash';
      if (controls.renderCommonSummary) controls.renderCommonSummary();
    } else if (key === 'floorinfo') {
      const draft = state.features.floorinfo.configDraft;
      draft.levelRules = normalizeFloorInfoRules(draft.levelRules);
      if (controls.param?.input) controls.param.input.value = draft.parameterName || '';
      if (controls.renderRules) controls.renderRules();
    } else if (key === 'familysuitability') {
      const draft = state.features.familysuitability.configDraft;
      draft.criteriaExcelPath = String(draft.criteriaExcelPath || '').trim();
      draft.criteriaRowCount = Number(draft.criteriaRowCount) || 0;
      draft.criteriaComboCount = Number(draft.criteriaComboCount) || 0;
      draft.criteriaSheetCount = Number(draft.criteriaSheetCount) || 0;
      draft.matchReviewText = String(draft.matchReviewText || '').trim();
      draft.mismatchReviewText = String(draft.mismatchReviewText || '').trim();
      draft.filterRules = normalizeFamilySuitabilityFilterRules(draft.filterRules, { keepEmpty: true });
      if (!draft.filterRules.length) {
        draft.filterRules = [createEmptyFamilySuitabilityFilterRule()];
      }
      if (controls.criteriaPath?.input) controls.criteriaPath.input.value = draft.criteriaExcelPath || '';
      if (controls.matchReview?.input) controls.matchReview.input.value = draft.matchReviewText || '';
      if (controls.mismatchReview?.input) controls.mismatchReview.input.value = draft.mismatchReviewText || '';
      if (typeof controls.renderPresetOptions === 'function') controls.renderPresetOptions();
      if (controls.renderCriteriaSummary) controls.renderCriteriaSummary();
      if (controls.renderFilterRules) controls.renderFilterRules();
    } else if (key === 'parameterduplication') {
      const draft = state.features.parameterduplication.configDraft || {};
      if (controls.scope?.select) controls.scope.select.value = normalizeParameterDuplicationScope(draft.scope);
      if (typeof controls.updateSelectionState === 'function') controls.updateSelectionState();
      if (typeof controls.renderSharedParamStatus === 'function') controls.renderSharedParamStatus();
      if (typeof controls.renderParameterDuplicationSelected === 'function') controls.renderParameterDuplicationSelected();
      if (typeof controls.renderSharedParamList === 'function') controls.renderSharedParamList();
      if (typeof controls.renderRecentOptions === 'function') controls.renderRecentOptions();
    } else if (key === 'parametermissing') {
      const draft = state.features.parametermissing.configDraft || {};
      if (typeof controls.renderSharedParamStatus === 'function') controls.renderSharedParamStatus();
      if (typeof controls.renderParameterMissingSelected === 'function') controls.renderParameterMissingSelected();
      if (typeof controls.renderSharedParamList === 'function') controls.renderSharedParamList();
      if (typeof controls.renderTargetFilterRows === 'function') controls.renderTargetFilterRows();
      if (typeof controls.renderExceptionRules === 'function') controls.renderExceptionRules();
      if (typeof controls.renderRecentOptions === 'function') controls.renderRecentOptions();
    } else if (key === 'worksetassignment') {
      const draft = state.features.worksetassignment.configDraft || {};
      if (controls.flaggedWorkset?.input) controls.flaggedWorkset.input.value = String(draft.flaggedWorksetName || '').trim();
    } else if (key === 'guid') {
      const draft = state.features.guid.configDraft;
      controls.includeFamily.input.checked = !!draft.includeFamily;
      controls.includeAnno.input.checked = !!draft.includeAnnotation;
      if (controls.useSyncComment?.input) controls.useSyncComment.input.checked = !!draft.useSyncComment;
      if (controls.syncComment?.input) controls.syncComment.input.value = draft.syncComment || 'KKY Tools - 파라미터 GUID 정리';
      if (controls.updateAnnotationState) controls.updateAnnotationState();
      if (controls.updateSyncCommentState) controls.updateSyncCommentState();
    } else if (key === 'familylink') {
      const draft = state.features.familylink.configDraft;
      controls.advanced.input.value = draft.targetsText;
      if (buildFamilyLinkConfig.renderList) buildFamilyLinkConfig.renderList();
    } else if (key === 'points') {
      const draft = state.features.points.configDraft;
      controls.unit.select.value = draft.unit;
    } else if (key === 'linkworkset') {
      const draft = state.features.linkworkset.configDraft;
      controls.applyDefault.input.checked = draft.applyDefaultWorksetOnly !== false;
      controls.useSyncComment.input.checked = !!draft.useSyncComment;
      controls.syncComment.input.value = draft.syncComment || 'KKY Tools - 링크 기본 웍셋 적용';
      controls.syncComment.input.disabled = !controls.useSyncComment.input.checked;
    } else if (key === 'common') {
      const draft = state.common.configDraft;
      controls.extra.input.value = draft.extraParams;
      controls.filter.input.value = draft.targetFilter;
      controls.excludeFilter.input.value = draft.excludeTargetFilter;
    }
  }

  function buildFilterExamples() {
    const wrap = div('filter-examples');
    const title = document.createElement('strong');
    title.textContent = '필터 예시';
    const note = document.createElement('p');
    note.textContent = '검토 대상 필터와 검토 제외 대상 필터는 같은 구문을 사용합니다. 좌측 파라미터 토큰은 공백 없는 이름을 권장하고, 구분자는 콤마(,) 또는 세미콜론(;)을 사용할 수 있습니다.';
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
      navigator.clipboard.writeText(text)
        .then(() => toast('예시 문구를 클립보드에 복사했습니다.', 'ok'))
        .catch(() => toast('브라우저 보안 설정 때문에 예시 문구를 복사하지 못했습니다. 직접 선택해 복사해 주세요.', 'err'));
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
      toast('예시 문구를 클립보드에 복사했습니다.', 'ok');
    } catch (e) {
      toast('예시 문구를 복사하지 못했습니다. 직접 선택해 복사해 주세요.', 'err');
    } finally {
      temp.remove();
    }
  }
}
