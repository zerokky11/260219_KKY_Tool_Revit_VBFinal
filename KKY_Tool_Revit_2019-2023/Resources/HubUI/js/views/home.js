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
} from '../core/hubFavorites.js?v=20260427a';

const MULTI_MODE_KEY = 'kky.hub.multiMode';
const BQC_GROUP_LABEL = '납품 시 BQC 검토';
const UTILITY_GROUP_LABEL = '유틸리티';

export function renderHome(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  const view = div('home-choice');
  const hero = div('home-choice-hero');
  hero.innerHTML = `
    <p class="home-choice-kicker">KKY Tool Hub</p>
    <h2>검토 방식을 선택하세요</h2>
    <p>납품 시 BQC 검토와 유틸리티 기능카드에서 원하는 기능을 빠르게 시작할 수 있습니다.</p>`;

  const bqcEntries = getHomeEntries(BQC_GROUP_LABEL);
  const utilityEntries = getHomeEntries(UTILITY_GROUP_LABEL);
  const searchSection = buildSearchSection();
  const favorites = getFavoriteEntries(4);
  const favoriteItems = favorites.length
    ? favorites.map((entry) => entry.label)
    : ['우클릭으로 즐겨찾기 추가', '선택형 / 별도 프로세스 분리', '즐겨찾기 전용 실행 화면'];

  const grid = div('home-choice-grid');
  grid.append(
    buildCard(
      '납품 시 BQC 검토',
      '납품 검토에 필요한 선택형 검토와 별도 워크플로우를 실행합니다.',
      'multi',
      bqcEntries.map((entry) => entry.label),
      'bqc',
      null,
      `현재 기능 ${bqcEntries.length}개`
    ),
    buildCard(
      '유틸리티',
      '보조 검토와 일괄 작업 기능을 실행합니다.',
      'multi',
      utilityEntries.map((entry) => entry.label),
      'utility',
      'utilities',
      `현재 기능 ${utilityEntries.length}개`
    ),
    buildCard(
      '자주 사용하는 기능',
      '즐겨찾기로 모은 기능만 따로 열어 선택형 검토와 별도 프로세스 실행을 바로 할 수 있습니다.',
      'favorites',
      favoriteItems
    )
  );

  const featureListSection = buildFeatureListSection(bqcEntries, utilityEntries);
  view.append(hero, searchSection, grid, featureListSection);
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
      'familylink',
      'points',
      'linkworkset',
      'sharedparambatch'
    ];
    const index = order.indexOf(entry?.id);
    return index >= 0 ? index : order.length;
  }

  function buildFeatureListSection(bqcItems, utilityItems) {
    const section = div('home-manual');
    const header = div('home-manual__header');
    header.innerHTML = `
      <p class="home-choice-kicker">Manual</p>
      <h3>현재 기능 목록</h3>
      <p>홈 검색과 즐겨찾기에 연결된 기능 기준입니다. 항목을 누르면 해당 화면으로 바로 이동합니다.</p>`;

    const grid = div('home-manual-grid');
    grid.append(
      buildFeatureListGroup(BQC_GROUP_LABEL, '선택형 검토 + 별도 워크플로우', bqcItems),
      buildFeatureListGroup(UTILITY_GROUP_LABEL, '보조 검토 + 일괄 작업', utilityItems)
    );
    section.append(header, grid);
    return section;
  }

  function buildFeatureListGroup(title, desc, entries) {
    const group = div('home-feature-group');
    const head = div('home-feature-group__head');
    const count = Array.isArray(entries) ? entries.length : 0;
    head.innerHTML = `
      <div>
        <strong>${title}</strong>
        <span>${desc}</span>
      </div>
      <em>${count}개</em>`;

    const list = div('home-feature-list');
    entries.forEach((entry) => {
      const item = document.createElement('button');
      item.type = 'button';
      item.className = 'home-feature-item';
      item.dataset.entryId = entry.id || '';

      const titleEl = document.createElement('strong');
      titleEl.className = 'home-feature-item__title';
      titleEl.textContent = entry.label || '';

      const descEl = document.createElement('span');
      descEl.className = 'home-feature-item__desc';
      descEl.textContent = entry.desc || '';

      item.append(titleEl, descEl);
      item.addEventListener('click', () => openHomeEntry(entry));
      bindHubEntryContextMenu(item, entry.id);
      list.append(item);
    });

    group.append(head, list);
    return group;
  }

  function openHomeEntry(entry) {
    if (!entry?.id) return;
    openHubEntry(entry.id);
  }

  function buildSearchSection() {
    const section = div('home-search');
    const top = div('home-search__top');
    top.innerHTML = `
      <p class="home-choice-kicker">Search</p>
      <h3>기능 검색</h3>
      <p>검색 결과에서 경로를 누르면 BQC 또는 유틸리티 화면에서 같은 검색어로 바로 이어집니다.</p>`;

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
    hint.textContent = '기능명이나 키워드를 입력하세요.';

    const results = div('home-search__results');
    const empty = div('home-search__empty');
    empty.textContent = '검색 결과가 없습니다.';

    const renderSearchResults = () => {
      const query = String(input.value || '').trim();
      clear(results);

      if (!query) {
        section.classList.remove('is-searching');
        hint.textContent = '기능명이나 키워드를 입력하세요.';
        return;
      }

      const entries = searchHubEntries(query, 8);
      section.classList.add('is-searching');

      if (!entries.length) {
        hint.textContent = `"${query}" 검색 결과가 없습니다.`;
        results.append(empty);
        return;
      }

      hint.textContent = `"${query}" 검색 결과 ${entries.length}개`;
      entries.forEach((entry) => {
        const card = div('home-search-result');
        card.innerHTML = `
          <div class="home-search-result__meta">
            <span>${entry.groupLabel || '기능'}</span>
            ${entry.favorite ? '<span class="home-search-result__badge">즐겨찾기</span>' : ''}
          </div>
          <strong class="home-search-result__title">${entry.label}</strong>
          <span class="home-search-result__desc">${entry.desc}</span>`;

        const actions = div('home-search-result__actions');
        const pathLabel = document.createElement('span');
        pathLabel.className = 'home-search-result__path-label';
        pathLabel.textContent = `경로: ${getEntryPathLabel(entry)}`;
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

  function buildFavoritePreviewSection() {
    const section = div('home-frequent');
    const header = div('home-frequent__header');
    header.innerHTML = `
      <div>
        <p class="home-choice-kicker">Favorites</p>
        <h3>자주 사용하는 기능</h3>
        <p>홈에서는 이름만 간단히 보고, 전체 열기에서 바로 실행하세요.</p>
      </div>`;

    const openAllBtn = document.createElement('button');
    openAllBtn.type = 'button';
    openAllBtn.className = 'btn btn--primary';
    openAllBtn.textContent = '전체 열기';
    openAllBtn.addEventListener('click', () => {
      location.hash = '#favorites';
    });
    header.append(openAllBtn);

    const quickGrid = div('home-frequent-grid');
    section.append(header, quickGrid);

    const renderFavorites = () => {
      clear(quickGrid);
      const entries = getFavoriteEntries(10);

      if (!entries.length) {
        const empty = div('home-frequent-empty');
        empty.textContent = '기능 카드에서 오른쪽 클릭으로 즐겨찾기를 추가하면 여기에서 바로 열 수 있습니다.';
        quickGrid.append(empty);
        return;
      }

      entries.forEach((entry) => {
        const chip = document.createElement('button');
        chip.type = 'button';
        chip.className = 'home-frequent-chip';
        chip.textContent = entry.label;
        chip.title = `${entry.label} 열기`;
        chip.addEventListener('click', () => {
          openHubEntry(entry.id, { panelRoute: 'favorites', recordUsage: false });
        });
        bindHubEntryContextMenu(chip, entry.id);
        quickGrid.append(chip);
      });
    };

    renderFavorites();

    const onFavoritesChange = () => {
      if (!section.isConnected) {
        window.removeEventListener(HUB_QUICK_ACCESS_CHANGE_EVENT, onFavoritesChange);
        return;
      }
      renderFavorites();
    };
    window.addEventListener(HUB_QUICK_ACCESS_CHANGE_EVENT, onFavoritesChange);

    return section;
  }

  function getEntryMode(entry) {
    if (entry?.multiMode === 'utility') return 'utility';
    if (entry?.multiMode === 'bqc') return 'bqc';
    return String(entry?.groupLabel || '').trim() === '유틸리티' ? 'utility' : 'bqc';
  }

  function getEntryPathLabel(entry) {
    return getEntryMode(entry) === 'utility' ? '유틸리티' : '납품 시 BQC 검토';
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
      ? `<p class="home-choice-meta">${previewText}</p>`
      : '';
    card.innerHTML = `
      <div class="home-choice-card__body">
        <div>
          <h3>${title}</h3>
          <p>${desc}</p>
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
}
