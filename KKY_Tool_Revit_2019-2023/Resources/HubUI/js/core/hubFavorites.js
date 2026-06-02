import { toast } from './dom.js';

const STORAGE_KEY = 'kky.hub.quickAccess.v1';
const PENDING_KEY = 'kky.hub.quickAccess.pending';
const MULTI_MODE_KEY = 'kky.hub.multiMode';
const PANEL_SEARCH_KEY = 'kky.hub.panelSearch.v1';
export const HUB_QUICK_ACCESS_CHANGE_EVENT = 'hub:quickaccess-changed';

const BQC_GROUP_LABEL = '납품 시 BQC 검토';
const UTILITY_GROUP_LABEL = '유틸리티';

const HUB_ENTRY_CATALOG = Object.freeze({
  connector: {
    id: 'connector',
    label: '파라미터 연속성 검토',
    desc: '연결된 MEP 객체 사이의 파라미터 연속성을 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: BQC_GROUP_LABEL
  },
  floorinfo: {
    id: 'floorinfo',
    label: '레벨 영역별 파라미터 검토',
    desc: '선택한 레벨 영역을 기준으로 파라미터 일치 여부를 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: BQC_GROUP_LABEL
  },
  familysuitability: {
    id: 'familysuitability',
    label: '패밀리 타입 적합성 검토',
    desc: '기준 엑셀의 카테고리/패밀리/타입 조합으로 실제 사용 타입의 적합성을 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: BQC_GROUP_LABEL
  },
  tapalign: {
    id: 'tapalign',
    label: '탭/분기 축 틀어짐 검토',
    desc: '탭/분기 피팅 축이 연결된 배관 또는 덕트 중심축에서 벗어났는지 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: BQC_GROUP_LABEL
  },
  tapdepth: {
    id: 'tapdepth',
    label: 'Tap, Saddle 모델링 검토 (묻힘)',
    desc: '실제 묻힘 깊이를 Takeoff Length Projection / Takeoff Length 기준과 비교합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: BQC_GROUP_LABEL
  },
  dupclash: {
    id: 'dupclash',
    label: '중복 / 자체 간섭 검토',
    desc: '여러 RVT를 대상으로 중복 검토 또는 자체 간섭 검토를 선택해 배치 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: BQC_GROUP_LABEL
  },
  worksetassignment: {
    id: 'worksetassignment',
    label: '웍셋 배정 검토',
    desc: '모델 객체의 웍셋 배정을 기본 웍셋(Workset1) 또는 지정한 웍셋 기준으로 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: BQC_GROUP_LABEL
  },
  parameterduplication: {
    id: 'parameterduplication',
    label: '프로젝트 파라미터 중복 검토',
    desc: '추가된 프로젝트 파라미터 중 이름이 중복된 항목을 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: BQC_GROUP_LABEL
  },
  parametermissing: {
    id: 'parametermissing',
    label: '파라미터 누락 검토',
    desc: '공유 텍스트 파라미터의 누락 여부를 공통 대상 필터와 예외 규칙 기준으로 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: BQC_GROUP_LABEL
  },
  deliverycleaner: {
    id: 'deliverycleaner',
    label: 'RVT 정리 (납품용)',
    desc: '납품 파일 작성을 위한 뷰 정리, 불필요 항목 제거(Purge), 검토용 속성 추출을 진행합니다.',
    route: 'deliverycleaner',
    groupLabel: BQC_GROUP_LABEL
  },
  conditionextract: {
    id: 'conditionextract',
    label: '조건별 객체 대상 속성 추출',
    desc: '조건식으로 객체를 추려 지정 속성, 좌표, 선형 정보를 함께 추출합니다.',
    route: 'conditionextract',
    groupLabel: BQC_GROUP_LABEL
  },
  dup: {
    id: 'dup',
    label: '중복 / 자체 간섭 검토',
    desc: '활성 문서에서 중복 객체 또는 자체 간섭을 검토합니다.',
    route: 'dup',
    groupLabel: UTILITY_GROUP_LABEL
  },
  paramprop: {
    id: 'paramprop',
    label: '패밀리 공유파라미터 추가/연동',
    desc: '복합 및 하위 패밀리에 지정한 파라미터를 추가하고 연동합니다.',
    route: 'paramprop',
    groupLabel: UTILITY_GROUP_LABEL
  },
  segmentpms: {
    id: 'segmentpms',
    label: 'Segment-PMS 비교 검토',
    desc: 'PMS 양식을 입력받아 Segment와 OD, ID 값을 비교 검토합니다.',
    route: 'segmentpms',
    groupLabel: UTILITY_GROUP_LABEL
  },
  parammodifier: {
    id: 'parammodifier',
    label: '파라미터 수정기',
    desc: '입력 조건 기반 필터링 대상의 지정 파라미터에 값을 일괄 입력합니다.',
    route: 'parammodifier',
    groupLabel: UTILITY_GROUP_LABEL
  },
  linkpath: {
    id: 'linkpath',
    label: 'Revit 링크 경로 추출/재지정',
    desc: '열린 RVT의 Revit 링크 경로를 추출하고 엑셀 기준으로 대상 경로를 반영합니다.',
    route: 'linkpath',
    groupLabel: UTILITY_GROUP_LABEL
  },
  lateralnozzle: {
    id: 'lateralnozzle',
    label: '노즐코드 KTA 단일화',
    desc: '접수받은 KTA 양식을 정해진 하나의 시트 양식으로 추출합니다.',
    route: 'lateralnozzle',
    groupLabel: UTILITY_GROUP_LABEL
  },
  guid: {
    id: 'guid',
    label: '파라미터 GUID 검토 및 정리',
    desc: '프로젝트와 패밀리 파라미터 GUID를 검토하고 정리 기준으로 정리합니다.',
    route: 'guid',
    groupLabel: UTILITY_GROUP_LABEL
  },
  tapdepthutility: {
    id: 'tapdepthutility',
    featureKey: 'tapdepth',
    label: 'Tap, Saddle 모델링 검토 (묻힘)',
    desc: '활성/선택 RVT에서 Tap/Saddle 묻힘 깊이를 두 Takeoff Length 기준으로 검토합니다.',
    route: 'multi',
    multiMode: 'utility',
    groupLabel: UTILITY_GROUP_LABEL
  },
  familylink: {
    id: 'familylink',
    label: '패밀리 공유파라미터 연동 검토',
    desc: '복합 패밀리를 대상으로 하위 패밀리와의 파라미터 연동 여부를 검토합니다.',
    route: 'multi',
    multiMode: 'utility',
    groupLabel: UTILITY_GROUP_LABEL
  },
  points: {
    id: 'points',
    label: '기준점/북각 추출',
    desc: 'RVT의 프로젝트 기준점, 측량 기준점, 북각 값을 추출합니다.',
    route: 'multi',
    multiMode: 'utility',
    groupLabel: UTILITY_GROUP_LABEL
  },
  linkworkset: {
    id: 'linkworkset',
    label: '링크 기본 웍셋 점검/적용',
    desc: '링크별 로드 상태와 열려 있는 웍셋 현황을 확인하고 기본 웍셋(Workset1)만 열리도록 재적용합니다.',
    route: 'multi',
    multiMode: 'utility',
    groupLabel: UTILITY_GROUP_LABEL
  },
  sharedparambatch: {
    id: 'sharedparambatch',
    label: '프로젝트 파라미터 일괄 추가',
    desc: '프로젝트 파일에 지정한 파라미터를 일괄 추가합니다.',
    route: 'sharedparambatch',
    groupLabel: UTILITY_GROUP_LABEL
  }
});

const boundEntries = new Set();
let memoryState = createState();
let memoryPending = null;
let memoryPanelSearch = createPanelSearchState();
let contextMenuEl = null;
let contextMenuCleanup = null;

function createState() {
  return {
    favorites: [],
    usage: {}
  };
}

function createPanelSearchState() {
  return {
    bqc: { query: '', entryId: '' },
    utility: { query: '', entryId: '' }
  };
}

function normalizeFavoriteEntryIds(entryIds) {
  const list = Array.isArray(entryIds) ? entryIds : [];
  return Array.from(new Set(
    list
      .map((id) => String(id || '').trim())
      .filter((id) => HUB_ENTRY_CATALOG[id])
  ));
}

function isSameFavoriteList(left, right) {
  const a = normalizeFavoriteEntryIds(left);
  const b = normalizeFavoriteEntryIds(right);
  if (a.length !== b.length) return false;
  return a.every((id, index) => id === b[index]);
}

function normalizeState(raw) {
  const next = createState();
  const favorites = Array.isArray(raw?.favorites) ? raw.favorites : [];
  next.favorites = Array.from(new Set(favorites.filter((id) => HUB_ENTRY_CATALOG[id])));

  const usage = raw && typeof raw.usage === 'object' ? raw.usage : {};
  Object.keys(usage).forEach((id) => {
    if (!HUB_ENTRY_CATALOG[id]) return;
    const item = usage[id] || {};
    const count = Math.max(0, Number(item.count) || 0);
    const lastUsedAt = Math.max(0, Number(item.lastUsedAt) || 0);
    if (count > 0 || lastUsedAt > 0) {
      next.usage[id] = { count, lastUsedAt };
    }
  });

  return next;
}

function readState() {
  try {
    const raw = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
    memoryState = normalizeState(raw);
  } catch {
    // Use in-memory fallback when localStorage is unavailable.
  }
  return memoryState;
}

function writeState(next) {
  memoryState = normalizeState(next);
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(memoryState));
  } catch {
    // Keep in-memory state only.
  }
  return memoryState;
}

function normalizePending(raw) {
  if (!raw || !HUB_ENTRY_CATALOG[raw.entryId]) return null;
  const createdAt = Math.max(0, Number(raw.createdAt) || Date.now());
  if (Date.now() - createdAt > 5 * 60 * 1000) return null;
  return {
    entryId: raw.entryId,
    route: String(raw.route || ''),
    createdAt
  };
}

function readPending() {
  try {
    const raw = JSON.parse(localStorage.getItem(PENDING_KEY) || 'null');
    memoryPending = normalizePending(raw);
    if (!memoryPending) localStorage.removeItem(PENDING_KEY);
  } catch {
    memoryPending = normalizePending(memoryPending);
  }
  return memoryPending;
}

function writePending(next) {
  memoryPending = normalizePending(next);
  try {
    if (memoryPending) localStorage.setItem(PENDING_KEY, JSON.stringify(memoryPending));
    else localStorage.removeItem(PENDING_KEY);
  } catch {
    // Keep in-memory fallback only.
  }
}

function normalizePanelSearchMode(mode) {
  return mode === 'utility' ? 'utility' : 'bqc';
}

function normalizePanelSearchState(raw) {
  const next = createPanelSearchState();
  ['bqc', 'utility'].forEach((mode) => {
    const item = raw && typeof raw === 'object' ? raw[mode] : null;
    const query = String(item?.query || '').trim();
    const entryId = HUB_ENTRY_CATALOG[item?.entryId] ? String(item.entryId) : '';
    next[mode] = {
      query,
      entryId: query ? entryId : ''
    };
  });
  return next;
}

function readPanelSearchState() {
  try {
    const raw = JSON.parse(localStorage.getItem(PANEL_SEARCH_KEY) || 'null');
    memoryPanelSearch = normalizePanelSearchState(raw);
  } catch {
    memoryPanelSearch = normalizePanelSearchState(memoryPanelSearch);
  }
  return memoryPanelSearch;
}

function writePanelSearchState(next) {
  memoryPanelSearch = normalizePanelSearchState(next);
  try {
    localStorage.setItem(PANEL_SEARCH_KEY, JSON.stringify(memoryPanelSearch));
  } catch {
    // Keep in-memory fallback only.
  }
  return memoryPanelSearch;
}

function emitChange() {
  syncHubEntryBindings();
  try {
    window.dispatchEvent(new CustomEvent(HUB_QUICK_ACCESS_CHANGE_EVENT));
  } catch {
    // Ignore custom event failures in embedded hosts.
  }
}

function compareQuickAccess(a, b) {
  if (a.favorite !== b.favorite) return a.favorite ? -1 : 1;
  if (a.count !== b.count) return b.count - a.count;
  if (a.lastUsedAt !== b.lastUsedAt) return b.lastUsedAt - a.lastUsedAt;
  return a.label.localeCompare(b.label, 'ko');
}

function setMultiMode(mode) {
  try {
    localStorage.setItem(MULTI_MODE_KEY, mode);
  } catch {
    // Ignore localStorage failures in embedded hosts.
  }
}

function navigateToRoute(route) {
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

function cleanupDisconnectedBindings() {
  for (const binding of Array.from(boundEntries)) {
    if (!binding.element || !binding.element.isConnected) {
      boundEntries.delete(binding);
    }
  }
}

function syncBinding(binding) {
  if (!binding?.element || !binding.entryId) return;
  binding.element.classList.add('hub-entry');
  binding.element.dataset.entryId = binding.entryId;
  if (isHubEntryFavorite(binding.entryId)) binding.element.dataset.favorite = 'true';
  else delete binding.element.dataset.favorite;
}

function hideContextMenu() {
  if (typeof contextMenuCleanup === 'function') {
    contextMenuCleanup();
  }
  contextMenuCleanup = null;

  if (contextMenuEl) {
    contextMenuEl.remove();
    contextMenuEl = null;
  }
}

function placeContextMenu(menu, x, y) {
  const pad = 12;
  const rect = menu.getBoundingClientRect();
  const left = Math.max(pad, Math.min(x, window.innerWidth - rect.width - pad));
  const top = Math.max(pad, Math.min(y, window.innerHeight - rect.height - pad));
  menu.style.left = `${left}px`;
  menu.style.top = `${top}px`;
}

function showContextMenu(x, y, entryId) {
  const entry = getHubEntry(entryId);
  if (!entry) return;

  hideContextMenu();

  const favorite = isHubEntryFavorite(entryId);
  const menu = document.createElement('div');
  menu.className = 'hub-context-menu';

  const meta = document.createElement('div');
  meta.className = 'hub-context-menu__meta';
  meta.textContent = entry.label;

  const action = document.createElement('button');
  action.type = 'button';
  action.className = 'hub-context-menu__action';
  action.textContent = favorite ? '즐겨찾기에서 제거' : '즐겨찾기에 추가';
  action.addEventListener('click', () => {
    setHubEntryFavorite(entryId, !favorite);
    hideContextMenu();
  });

  menu.append(meta, action);
  document.body.append(menu);
  placeContextMenu(menu, x, y);
  contextMenuEl = menu;

  const closeOnPointer = (ev) => {
    if (contextMenuEl && contextMenuEl.contains(ev.target)) return;
    hideContextMenu();
  };
  const closeOnKey = (ev) => {
    if (ev.key === 'Escape') hideContextMenu();
  };

  document.addEventListener('mousedown', closeOnPointer, true);
  document.addEventListener('contextmenu', closeOnPointer, true);
  document.addEventListener('keydown', closeOnKey, true);
  window.addEventListener('resize', hideContextMenu);
  window.addEventListener('blur', hideContextMenu);
  window.addEventListener('scroll', hideContextMenu, true);
  window.addEventListener('hashchange', hideContextMenu);

  contextMenuCleanup = () => {
    document.removeEventListener('mousedown', closeOnPointer, true);
    document.removeEventListener('contextmenu', closeOnPointer, true);
    document.removeEventListener('keydown', closeOnKey, true);
    window.removeEventListener('resize', hideContextMenu);
    window.removeEventListener('blur', hideContextMenu);
    window.removeEventListener('scroll', hideContextMenu, true);
    window.removeEventListener('hashchange', hideContextMenu);
  };
}

export function getHubEntry(entryId) {
  return HUB_ENTRY_CATALOG[entryId] || null;
}

export function getHubEntries(options = {}) {
  const groupLabel = String(options?.groupLabel || '').trim();
  const multiMode = String(options?.multiMode || '').trim();
  return Object.values(HUB_ENTRY_CATALOG)
    .filter((entry) => !groupLabel || entry.groupLabel === groupLabel)
    .filter((entry) => !multiMode || entry.multiMode === multiMode)
    .map((entry) => ({ ...entry }));
}

export function isHubEntryFavorite(entryId) {
  return readState().favorites.includes(entryId);
}

export function getHubEntryUsage(entryId) {
  const usage = readState().usage[entryId] || {};
  return {
    count: Math.max(0, Number(usage.count) || 0),
    lastUsedAt: Math.max(0, Number(usage.lastUsedAt) || 0)
  };
}

export function setHubEntryFavorite(entryId, nextFavorite) {
  const entry = getHubEntry(entryId);
  if (!entry) return false;

  const current = readState();
  const favorites = new Set(current.favorites);
  const shouldFavorite = typeof nextFavorite === 'boolean' ? nextFavorite : !favorites.has(entryId);

  if (shouldFavorite) favorites.add(entryId);
  else favorites.delete(entryId);

  writeState({
    ...current,
    favorites: Array.from(favorites)
  });
  emitChange();

  toast(
    shouldFavorite
      ? `'${entry.label}' 기능을 즐겨찾기에 추가했습니다.`
      : `'${entry.label}' 기능을 즐겨찾기에서 제거했습니다.`,
    shouldFavorite ? 'ok' : 'info'
  );

  return shouldFavorite;
}

export function replaceHubFavorites(entryIds, options = {}) {
  const current = readState();
  const favorites = normalizeFavoriteEntryIds(entryIds);
  if (isSameFavoriteList(current.favorites, favorites)) {
    return favorites.slice();
  }

  writeState({
    ...current,
    favorites
  });
  emitChange();

  if (options?.toast !== false) {
    const sourceLabel = String(options?.sourceLabel || '').trim();
    const prefix = sourceLabel ? `${sourceLabel}에서 ` : '';
    toast(
      favorites.length
        ? `${prefix}즐겨찾기 ${favorites.length}개를 복원했습니다.`
        : `${prefix}즐겨찾기 목록이 비어 있습니다.`,
      favorites.length ? 'ok' : 'info'
    );
  }

  return favorites.slice();
}

export function recordHubEntryUse(entryId) {
  const entry = getHubEntry(entryId);
  if (!entry) return;

  const current = readState();
  const usage = { ...(current.usage || {}) };
  const prev = usage[entryId] || { count: 0, lastUsedAt: 0 };
  usage[entryId] = {
    count: Math.max(0, Number(prev.count) || 0) + 1,
    lastUsedAt: Date.now()
  };

  writeState({
    ...current,
    usage
  });
  emitChange();
}

export function getHubQuickAccessEntries(limit = 6) {
  const state = readState();
  const items = Object.values(HUB_ENTRY_CATALOG)
    .map((entry) => {
      const usage = state.usage[entry.id] || {};
      return {
        ...entry,
        favorite: state.favorites.includes(entry.id),
        count: Math.max(0, Number(usage.count) || 0),
        lastUsedAt: Math.max(0, Number(usage.lastUsedAt) || 0)
      };
    })
    .filter((entry) => entry.favorite || entry.count > 0)
    .sort(compareQuickAccess);

  return typeof limit === 'number' && limit > 0 ? items.slice(0, limit) : items;
}

export function getFavoriteEntries(limit = 0) {
  const state = readState();
  const items = state.favorites
    .map((id) => {
      const entry = HUB_ENTRY_CATALOG[id];
      if (!entry) return null;
      const usage = state.usage[id] || {};
      return {
        ...entry,
        favorite: true,
        count: Math.max(0, Number(usage.count) || 0),
        lastUsedAt: Math.max(0, Number(usage.lastUsedAt) || 0)
      };
    })
    .filter(Boolean)
    .sort(compareQuickAccess);

  return typeof limit === 'number' && limit > 0 ? items.slice(0, limit) : items;
}

function normalizeSearchValue(value) {
  return String(value || '').toLowerCase().trim();
}

function compactSearchValue(value) {
  return normalizeSearchValue(value).replace(/\s+/g, '');
}

function buildSearchRank(entry, queryNorm, queryCompact) {
  if (!entry || !queryNorm) return 0;

  const labelNorm = normalizeSearchValue(entry.label);
  const labelCompact = compactSearchValue(entry.label);
  const descNorm = normalizeSearchValue(entry.desc);
  const descCompact = compactSearchValue(entry.desc);
  const groupNorm = normalizeSearchValue(entry.groupLabel);
  const groupCompact = compactSearchValue(entry.groupLabel);

  let score = 0;

  if (labelNorm === queryNorm || labelCompact === queryCompact) score += 120;
  else if (labelNorm.startsWith(queryNorm) || labelCompact.startsWith(queryCompact)) score += 90;
  else if (labelNorm.includes(queryNorm) || labelCompact.includes(queryCompact)) score += 70;

  if (groupNorm.includes(queryNorm) || groupCompact.includes(queryCompact)) score += 24;
  if (descNorm.includes(queryNorm) || descCompact.includes(queryCompact)) score += 12;

  return score;
}

export function searchHubEntries(query, limit = 0) {
  const queryNorm = normalizeSearchValue(query);
  const queryCompact = queryNorm.replace(/\s+/g, '');
  if (!queryNorm) return [];

  const state = readState();
  const items = Object.values(HUB_ENTRY_CATALOG)
    .map((entry) => {
      const usage = state.usage[entry.id] || {};
      return {
        ...entry,
        favorite: state.favorites.includes(entry.id),
        count: Math.max(0, Number(usage.count) || 0),
        lastUsedAt: Math.max(0, Number(usage.lastUsedAt) || 0),
        searchRank: buildSearchRank(entry, queryNorm, queryCompact)
      };
    })
    .filter((entry) => entry.searchRank > 0)
    .sort((a, b) => b.searchRank - a.searchRank || compareQuickAccess(a, b))
    .map(({ searchRank, ...entry }) => entry);

  return typeof limit === 'number' && limit > 0 ? items.slice(0, limit) : items;
}

export function getHubPanelSearch(mode = 'bqc') {
  const key = normalizePanelSearchMode(mode);
  const state = readPanelSearchState();
  const item = state[key] || {};
  return {
    query: String(item.query || ''),
    entryId: String(item.entryId || '')
  };
}

export function setHubPanelSearch(mode = 'bqc', nextValue = {}) {
  const key = normalizePanelSearchMode(mode);
  const current = readPanelSearchState();
  current[key] = {
    query: String(nextValue?.query || '').trim(),
    entryId: String(nextValue?.entryId || '').trim()
  };
  const next = writePanelSearchState(current);
  return {
    query: String(next[key]?.query || ''),
    entryId: String(next[key]?.entryId || '')
  };
}

export function bindHubEntryContextMenu(element, entryId) {
  if (!element || !getHubEntry(entryId)) return element;

  const binding = { element, entryId };
  boundEntries.add(binding);
  syncBinding(binding);

  if (!element.dataset.hubFavoriteBound) {
    element.dataset.hubFavoriteBound = 'true';
    element.addEventListener('contextmenu', (ev) => {
      ev.preventDefault();
      ev.stopPropagation();
      showContextMenu(ev.clientX, ev.clientY, entryId);
    });
  }

  return element;
}

export function syncHubEntryBindings() {
  cleanupDisconnectedBindings();
  boundEntries.forEach(syncBinding);
}

export function openHubEntry(entryId, options = {}) {
  const entry = getHubEntry(entryId);
  if (!entry) return;

  if (options.recordUsage !== false) {
    recordHubEntryUse(entryId);
  }

  const panelRoute = String(options.panelRoute || '').replace(/^#/, '');
  const route = String(panelRoute || entry.route || '').replace(/^#/, '');
  if (!route) return;

  if (panelRoute) {
    writePending({
      entryId,
      route,
      createdAt: Date.now()
    });
  } else if (route === 'multi') {
    writePending({
      entryId,
      route,
      createdAt: Date.now()
    });
    setMultiMode(entry.multiMode || 'bqc');
  } else {
    writePending(null);
  }

  navigateToRoute(route);
}

export function consumePendingHubEntry(route = '') {
  const pending = readPending();
  if (!pending) return null;
  if (route && pending.route !== route) return null;
  writePending(null);
  return pending;
}
