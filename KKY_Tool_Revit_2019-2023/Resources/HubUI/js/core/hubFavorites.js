import { toast } from './dom.js';

const STORAGE_KEY = 'kky.hub.quickAccess.v1';
const PENDING_KEY = 'kky.hub.quickAccess.pending';
const MULTI_MODE_KEY = 'kky.hub.multiMode';
const PANEL_SEARCH_KEY = 'kky.hub.panelSearch.v1';
export const HUB_QUICK_ACCESS_CHANGE_EVENT = 'hub:quickaccess-changed';

const HUB_ENTRY_CATALOG = Object.freeze({
  connector: {
    id: 'connector',
    label: '파라미터 연속성 검토',
    desc: '연결된 MEP 객체 사이의 파라미터 연속성을 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: '납품 시 BQC 검토'
  },
  floorinfo: {
    id: 'floorinfo',
    label: '층 정보 파라미터 검토',
    desc: '선택한 층 정보 기준으로 파라미터 값 일치 여부를 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: '납품 시 BQC 검토'
  },
  familysuitability: {
    id: 'familysuitability',
    label: 'Family 적합성 검토',
    desc: '기준 엑셀의 Category / Family / Type 조합으로 실제 사용 타입 적합성을 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: '납품 시 BQC 검토'
  },
  tapalign: {
    id: 'tapalign',
    label: '탭/분기 축 틀어짐 검토',
    desc: '탭/분기 피팅 축이 연결된 배관 또는 덕트 중심축에서 벗어났는지 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: '납품 시 BQC 검토'
  },
  dupclash: {
    id: 'dupclash',
    label: '중복 / 자체간섭 검토',
    desc: '여러 RVT를 대상으로 중복검토 또는 자체간섭검토를 실행합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: '납품 시 BQC 검토'
  },
  worksetassignment: {
    id: 'worksetassignment',
    label: '웍셋 배정 검토',
    desc: '모델 객체가 기본 Workset1 기준에 맞게 배정되었는지 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: '납품 시 BQC 검토'
  },
  parameterduplication: {
    id: 'parameterduplication',
    label: 'Project Parameter 중복 검토',
    desc: '추가된 Project Parameter 중 이름이 중복된 항목을 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: '납품 시 BQC 검토'
  },
  parametermissing: {
    id: 'parametermissing',
    label: '파라미터 누락 검토',
    desc: '공유 Text 파라미터의 누락 여부를 공통 대상 필터와 예외 규칙 기준으로 검토합니다.',
    route: 'multi',
    multiMode: 'bqc',
    groupLabel: '납품 시 BQC 검토'
  },
  deliverycleaner: {
    id: 'deliverycleaner',
    label: 'RVT 정리 (납품용)',
    desc: '납품 파일 작성을 위한 뷰 정리, Purge, 검토용 속성 추출을 진행합니다.',
    route: 'deliverycleaner',
    groupLabel: '납품 시 BQC 검토'
  },
  conditionextract: {
    id: 'conditionextract',
    label: '조건별 객체 대상 속성 추출',
    desc: '조건식으로 객체를 추려 지정한 속성값과 좌표, 수량 정보를 함께 추출합니다.',
    route: 'conditionextract',
    groupLabel: '납품 시 BQC 검토'
  },
  dup: {
    id: 'dup',
    label: '중복 / 자체간섭 검토',
    desc: '활성 문서에서 중복 객체 또는 자체 간섭을 검토합니다.',
    route: 'dup',
    groupLabel: '유틸리티'
  },
  paramprop: {
    id: 'paramprop',
    label: '패밀리 공유파라미터 추가/연동',
    desc: '복합 및 하위 패밀리에 지정한 파라미터를 추가하고 연동합니다.',
    route: 'paramprop',
    groupLabel: '유틸리티'
  },
  segmentpms: {
    id: 'segmentpms',
    label: 'Segment와 PMS 비교 검토',
    desc: 'PMS 양식을 기준으로 Segment의 OD, ID 값을 비교 검토합니다.',
    route: 'segmentpms',
    groupLabel: '유틸리티'
  },
  parammodifier: {
    id: 'parammodifier',
    label: '파라미터 수정기',
    desc: '입력 조건 기반 필터만 대상으로 지정한 파라미터 값을 일괄 입력합니다.',
    route: 'parammodifier',
    groupLabel: '유틸리티'
  },
  linkpath: {
    id: 'linkpath',
    label: 'Revit 링크 경로 추출/적용',
    desc: '여러 RVT의 링크 경로를 추출하고 수정 기준으로 반영합니다.',
    route: 'linkpath',
    groupLabel: '유틸리티'
  },
  lateralnozzle: {
    id: 'lateralnozzle',
    label: '인접코드 KTA 추출',
    desc: '정리된 KTA 양식에 맞춰 필요한 시트 형식으로 출력합니다.',
    route: 'lateralnozzle',
    groupLabel: '유틸리티'
  },
  guid: {
    id: 'guid',
    label: '파라미터 GUID 검토 및 정리',
    desc: '프로젝트와 패밀리 파라미터 GUID를 검토하고 정리 기준으로 정리합니다.',
    route: 'guid',
    groupLabel: '유틸리티'
  },
  familylink: {
    id: 'familylink',
    label: '패밀리 공유파라미터 연동 검토',
    desc: '복합 패밀리를 대상으로 하위 패밀리와의 파라미터 연동 여부를 검토합니다.',
    route: 'multi',
    multiMode: 'utility',
    groupLabel: '유틸리티'
  },
  points: {
    id: 'points',
    label: '프로젝트 좌표 Point 추출',
    desc: '지정한 RVT 파일의 Project/Survey 기준 좌표를 추출합니다.',
    route: 'multi',
    multiMode: 'utility',
    groupLabel: '유틸리티'
  },
  linkworkset: {
    id: 'linkworkset',
    label: '링크 기본 웍셋 닫기/적용',
    desc: '각 RVT의 Revit 링크를 닫고 기본 Workset1만 열리도록 적용합니다.',
    route: 'multi',
    multiMode: 'utility',
    groupLabel: '유틸리티'
  },
  sharedparambatch: {
    id: 'sharedparambatch',
    label: 'Project 파라미터 일괄 추가',
    desc: '프로젝트 파일에 지정한 파라미터를 일괄 추가합니다.',
    route: 'sharedparambatch',
    groupLabel: '유틸리티'
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
  action.textContent = favorite ? '利먭꺼李얘린 ?댁젣' : '利먭꺼李얘린 異붽?';
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
      ? `'${entry.label}'???먯＜ ?ъ슜?섎뒗 湲곕뒫??異붽??덉뒿?덈떎.`
      : `'${entry.label}'???먯＜ ?ъ슜?섎뒗 湲곕뒫?먯꽌 ?쒓굅?덉뒿?덈떎.`,
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
    const prefix = sourceLabel ? `${sourceLabel}?먯꽌 ` : '';
    toast(
      favorites.length
        ? `${prefix}${favorites.length}媛?利먭꺼李얘린瑜?蹂듭썝?덉뒿?덈떎.`
        : `${prefix}利먭꺼李얘린 紐⑸줉??鍮꾩썱?듬땲??`,
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
