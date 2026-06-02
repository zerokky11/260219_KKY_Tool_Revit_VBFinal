// Resources/HubUI/js/views/home.js
import { clear, div } from '../core/dom.js';
import {
  bindHubEntryContextMenu,
  getFavoriteEntries,
  getHubEntries,
  HUB_QUICK_ACCESS_CHANGE_EVENT,
  openHubEntry,
  searchHubEntries,
  setHubPanelSearch
} from '../core/hubFavorites.js?v=20260504e';

const MULTI_MODE_KEY = 'kky.hub.multiMode';
const BQC_GROUP_LABEL = '납품 시 BQC 검토';
const UTILITY_GROUP_LABEL = '유틸리티';

export function renderHome(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  const view = div('home-choice');
  const hero = div('home-choice-hero');
  hero.innerHTML = `
    <p class="home-choice-kicker">작업 시작</p>
    <h2>검토 방식을 선택해 주세요</h2>
    <p>납품 시 BQC 검토와 유틸리티 기능을 작업 흐름에 맞춰 바로 시작할 수 있습니다.</p>`;

  const bqcEntries = getHomeEntries(BQC_GROUP_LABEL);
  const utilityEntries = getHomeEntries(UTILITY_GROUP_LABEL);
  const searchSection = buildSearchSection();
  const favorites = getFavoriteEntries(4);
  const favoriteItems = favorites.length
    ? favorites.map((entry) => entry.label)
    : ['기능 카드 우클릭으로 즐겨찾기 추가', '선택한 검토만 빠르게 실행', '즐겨찾기 전용 화면에서 관리'];

  const grid = div('home-choice-grid');
  grid.append(
    buildCard(
      '납품 시 BQC 검토',
      '납품 전 검토에 필요한 기능을 선택해 실행하고 결과를 정리합니다.',
      'multi',
      bqcEntries.map((entry) => entry.label),
      'bqc',
      null,
      `현재 기능 ${bqcEntries.length}개`
    ),
    buildCard(
      '유틸리티',
      '링크, 파라미터, 패밀리처럼 반복되는 보조 작업을 실행합니다.',
      'multi',
      utilityEntries.map((entry) => entry.label),
      'utility',
      'utilities',
      `현재 기능 ${utilityEntries.length}개`
    ),
    buildCard(
      '자주 사용하는 기능',
      '즐겨찾기로 모은 기능만 따로 열어 자주 쓰는 작업을 바로 실행합니다.',
      'favorites',
      favoriteItems
    )
  );

  view.append(hero, searchSection, grid);
  target.append(view);

  function getHomeEntries(groupLabel) {
    return getHubEntries({ groupLabel })
      .sort((left, right) => {
        const a = getEntrySortRank(left);
        const b = getEntrySortRank(right);
        if (a !== b) return a - b;
        return String(left.label || '').localeCompare(String(right.label || ''), 'ko');
      });
  }

  function getEntrySortRank(entry) {
    const order = [
      'deliverycleaner',
      'conditionextract',
      'connector',
      'floorinfo',
      'familysuitability',
      'tapalign',
      'tapdepth',
      'dupclash',
      'worksetassignment',
      'parameterduplication',
      'parametermissing',
      'dup',
      'paramprop',
      'segmentpms',
      'parammodifier',
      'linkpath',
      'lateralnozzle',
      'guid',
      'tapdepthutility',
      'familylink',
      'points',
      'linkworkset',
      'sharedparambatch'
    ];
    const index = order.indexOf(entry?.id);
    return index >= 0 ? index : order.length;
  }

  function buildSearchSection() {
    const section = div('home-search');
    const top = div('home-search__top');
    top.innerHTML = `
      <p class="home-choice-kicker">검색</p>
      <h3>기능 검색</h3>
      <p>기능명이나 업무 키워드를 검색하면 해당 화면으로 바로 이동할 수 있습니다.</p>`;

    const field = div('home-search__field');
    const icon = document.createElement('span');
    icon.className = 'home-search__icon';
    icon.textContent = '⌕';

    const input = document.createElement('input');
    input.type = 'search';
    input.className = 'home-search__input';
    input.placeholder = '예: 중복, GUID, 링크, 파라미터';
    input.autocomplete = 'off';
    input.spellcheck = false;

    field.append(icon, input);

    const hint = div('home-search__hint');
    hint.textContent = '기능명, 검토 항목, 업무 키워드를 입력해 주세요.';

    const results = div('home-search__results');
    const empty = div('home-search__empty');
    empty.textContent = '검색 결과가 없습니다. 다른 기능명이나 업무 키워드를 입력해 주세요.';

    const renderSearchResults = () => {
      const query = String(input.value || '').trim();
      clear(results);

      if (!query) {
        section.classList.remove('is-searching');
        hint.textContent = '기능명, 검토 항목, 업무 키워드를 입력해 주세요.';
        return;
      }

      const entries = searchHubEntries(query, 8);
      section.classList.add('is-searching');

      if (!entries.length) {
        hint.textContent = `"${query}"에 대한 검색 결과가 없습니다. 다른 기능명이나 업무 키워드를 입력해 주세요.`;
        results.append(empty);
        return;
      }

      hint.textContent = `"${query}" 검색 결과 ${entries.length}개`;
      entries.forEach((entry) => {
        const card = div('home-search-result');
        card.innerHTML = `
          <div class="home-search-result__meta">
            <span>${escapeHtml(entry.groupLabel || '기능')}</span>
            ${entry.favorite ? '<span class="home-search-result__badge">즐겨찾기</span>' : ''}
          </div>
          <strong class="home-search-result__title">${escapeHtml(entry.label)}</strong>
          <span class="home-search-result__desc">${escapeHtml(entry.desc)}</span>`;

        const actions = div('home-search-result__actions');
        const pathLabel = document.createElement('span');
        pathLabel.className = 'home-search-result__path-label';
        pathLabel.textContent = `화면: ${getEntryPathLabel(entry)}`;
        const pathBtn = document.createElement('button');
        pathBtn.type = 'button';
        pathBtn.className = 'btn btn--secondary home-search-result__path';
        pathBtn.textContent = getEntryPathLabel(entry);
        pathBtn.addEventListener('click', () => {
          openSearchPath(entry, query);
        });
        actions.append(pathLabel, pathBtn);
        card.append(actions);
        bindHubEntryContextMenu(card, entry.id);
        results.append(card);
      });
    };

    input.addEventListener('input', renderSearchResults);

    const onFavoritesChange = () => {
      if (!section.isConnected) {
        window.removeEventListener(HUB_QUICK_ACCESS_CHANGE_EVENT, onFavoritesChange);
        return;
      }
      if (String(input.value || '').trim()) renderSearchResults();
    };
    window.addEventListener(HUB_QUICK_ACCESS_CHANGE_EVENT, onFavoritesChange);

    section.append(top, field, hint, results);
    return section;
  }

  function getEntryMode(entry) {
    if (entry?.multiMode === 'utility') return 'utility';
    if (entry?.multiMode === 'bqc') return 'bqc';
    return String(entry?.groupLabel || '').trim() === UTILITY_GROUP_LABEL ? 'utility' : 'bqc';
  }

  function getEntryPathLabel(entry) {
    return getEntryMode(entry) === 'utility' ? UTILITY_GROUP_LABEL : BQC_GROUP_LABEL;
  }

  function openSearchPath(entry, query) {
    const mode = getEntryMode(entry);
    setMultiMode(mode);
    setHubPanelSearch(mode, {
      query,
      entryId: entry?.id || ''
    });
    openHubEntry(entry.id, { panelRoute: 'multi', recordUsage: false });
  }

  function buildCard(title, desc, hash, items, multiMode, anchorId, metaLabel = '대표 기능') {
    const card = document.createElement('button');
    card.type = 'button';
    card.className = 'home-choice-card';
    const previewItems = Array.isArray(items) ? items.slice(0, 3) : [];
    const remainingCount = Array.isArray(items) ? Math.max(0, items.length - previewItems.length) : 0;
    const previewText = previewItems.length
      ? `${metaLabel}: ${previewItems.join(', ')}${remainingCount > 0 ? ` 외 ${remainingCount}개` : ''}`
      : '';
    const listHtml = previewText
      ? `<p class="home-choice-meta">${escapeHtml(previewText)}</p>`
      : '';
    card.innerHTML = `
      <div class="home-choice-card__body">
        <div>
          <h3>${escapeHtml(title)}</h3>
          <p>${escapeHtml(desc)}</p>
          ${listHtml}
        </div>
        <span class="home-choice-card__icon">+</span>
      </div>
      <span class="home-choice-cta btn btn--primary">바로가기</span>`;
    card.addEventListener('click', () => {
      if (multiMode) setMultiMode(multiMode);
      location.hash = `#${hash}`;
      if (anchorId) {
        setTimeout(() => {
          const el = document.getElementById(anchorId);
          if (el && el.scrollIntoView) {
            el.scrollIntoView({ behavior: 'smooth', block: 'start' });
          }
        }, 240);
      }
    });
    return card;
  }

  function setMultiMode(mode) {
    try {
      localStorage.setItem(MULTI_MODE_KEY, mode);
    } catch {
      // Ignore localStorage failures in embedded hosts.
    }
  }

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
}
