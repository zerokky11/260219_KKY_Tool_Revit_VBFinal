// Resources/HubUI/js/views/home.js
import { clear, div } from '../core/dom.js';

const MULTI_MODE_KEY = 'kky.hub.multiMode';

export function renderHome(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  const view = div('home-choice');
  const hero = div('home-choice-hero');
  hero.innerHTML = `
    <p class="home-choice-kicker">KKY Tool Hub</p>
    <h2>검토 방식을 선택하세요</h2>
    <p>납품 시 BQC 검토와 유틸리티 기능카드에서 원하는 기능을 빠르게 시작할 수 있습니다.</p>`;

  const grid = div('home-choice-grid');
  grid.append(
    buildCard(
      '납품 시 BQC 검토',
      '납품 검토에 필요한 핵심 기능을 선택해 실행합니다.',
      'multi',
      [
        '파라미터 연속성 검토',
        '레벨 영역별 파라미터 검토',
        'RVT 정리 (납품용)',
        '중복 / 자체간섭 검토'
      ],
      'bqc'
    ),
    buildCard(
      '유틸리티',
      '보조 검토와 일괄 작업 기능을 실행합니다.',
      'multi',
      [
        '파라미터 GUID 검토',
        '패밀리 공유파라미터 연동 검토',
        '프로젝트대상 Point 좌표 추출',
        '중복 / 자체간섭 검토',
        '패밀리 공유파라미터 추가/연동',
        'Segment↔PMS 비교 검토',
        '노즐코드 KTA 단일화',
        '파라미터 수정기',
        'Project 파라미터 일괄 추가'
      ],
      'utility',
      'utilities'
    )
  );

  view.append(hero, grid);
  target.append(view);

  function buildCard(title, desc, hash, items, multiMode, anchorId) {
    const card = document.createElement('button');
    card.type = 'button';
    card.className = 'home-choice-card';
    const previewItems = Array.isArray(items) ? items.slice(0, 2) : [];
    const remainingCount = Array.isArray(items) ? Math.max(0, items.length - previewItems.length) : 0;
    const previewText = previewItems.length
      ? `대표 기능: ${previewItems.join(', ')}${remainingCount > 0 ? ` 외 ${remainingCount}개` : ''}`
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
