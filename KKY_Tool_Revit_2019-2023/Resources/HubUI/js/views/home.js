// Resources/HubUI/js/views/home.js
import { clear, div } from '../core/dom.js';

const MULTI_MODE_KEY = 'kky.hub.multiMode';

const HOME_FEATURE_GROUPS = [
    {
        title: '납품 시 BQC 검토',
        desc: '선택형 검토 + 별도 워크플로우',
        route: 'multi',
        multiMode: 'bqc',
        count: 10,
        items: [
            ['RVT 정리 (납품용)', '납품파일 작성을 위한 뷰정리, Purge, 검토용 속성 추출'],
            ['조건별 객체 대상 속성 추출', '조건식으로 객체를 추려 지정 속성 + 좌표/선형 정보 추출'],
            ['파라미터 연속성 검토', '연결된 MEP 객체의 파라미터 연속성 검토'],
            ['레벨 영역별 파라미터 검토', '선택한 레벨 영역 기준 파라미터 일치 여부 검토'],
            ['Family 적합성 검토', '기준 엑셀의 Category / Family / Type 조합 비교'],
            ['탭/분기 축 틀어짐 검토', '배관/덕트 중심축 기준 탭/분기 피팅 축 검토'],
            ['중복 / 자체간섭 검토', '여러 RVT 대상 중복검토 또는 자체간섭검토'],
            ['웍셋 배정 검토', 'Workset1 또는 입력한 특정 workset 기준 검토'],
            ['Project Parameter 중복 검토', '추가된 Project Parameter 이름 중복 검토'],
            ['파라미터 누락 검토', '공유 Text 파라미터 누락 여부와 예외 규칙 검토']
        ]
    },
    {
        title: '유틸리티',
        desc: '보조 검토 + 일괄 작업',
        route: 'multi',
        multiMode: 'utility',
        anchorId: 'utilities',
        count: 11,
        items: [
            ['중복 / 자체간섭 검토', '활성 문서에서 중복 객체 또는 자체 간섭 검토'],
            ['패밀리 공유파라미터 추가/연동', '복합 및 하위 패밀리에 지정 파라미터 추가/연동'],
            ['Segment↔PMS 비교 검토', 'PMS 양식 기준 Segment OD, ID 비교 검토'],
            ['파라미터 수정기', '조건 기반 필터링 대상으로 파라미터 값 일괄 입력'],
            ['Revit 링크 경로 추출/재지정', '닫힌 RVT의 링크 경로 추출 및 엑셀 기준 반영'],
            ['노즐코드 KTA 단일화', '접수받은 KTA 양식을 하나의 시트 양식으로 추출'],
            ['파라미터 GUID 검토 및 정리', '프로젝트/패밀리 파라미터 GUID 검토 및 삭제 기준 정리'],
            ['패밀리 공유파라미터 연동 검토', '복합 패밀리와 하위 패밀리 파라미터 연동 여부 검토'],
            ['프로젝트대상 Point 좌표 추출', 'RVT 파일의 Project/Survey 북각 좌표 추출'],
            ['링크 기본 웍셋 점검/적용', 'Revit 링크 open workset 현황 점검 및 기본 Workset1 적용'],
            ['Project 파라미터 일괄 추가', '프로젝트 파일에 지정 파라미터 일괄 추가']
        ]
    }
];

export function renderHome(root) {
    const target = root || document.getElementById('view-root') || document.getElementById('app');
    clear(target);

    const view = div('home-choice');
    const hero = div('home-choice-hero');
    hero.innerHTML = `
        <p class="home-choice-kicker">KKY Tool Hub</p>
        <h2>검토 방식을 선택하세요</h2>
        <p>활성 문서 기반 검토 또는 다중 RVT 배치 검토를 시작할 수 있습니다.</p>`;

    const grid = div('home-choice-grid');
    grid.append(
        buildCard(
            '활성 문서 기능',
            '현재 열려있는 Revit 문서를 대상으로 빠르게 검토를 수행합니다.',
            'active-menu',
            [
                '중복 객체 검토: 현재 열린 문서에서 중복 요소/패밀리 점검',
                '복합 패밀리 공유파라미터 추가 및 연동: 공유 파라미터 추가/연동 수행'
            ]
        ),
        buildCard(
            '납품시 BQC 검토',
            '납품 검토에 필요한 선택형 검토와 별도 워크플로우를 실행합니다.',
            'multi',
            HOME_FEATURE_GROUPS[0].items.slice(0, 4).map(([title]) => title),
            'bqc'
        ),
        buildCard(
            '유틸리티',
            '보조 검토와 일괄 작업 기능을 실행합니다.',
            'multi',
            HOME_FEATURE_GROUPS[1].items.slice(0, 5).map(([title]) => title),
            'utility',
            'utilities'
        )
    );

    view.append(hero, grid, buildFeatureListSection());
    target.append(view);

    function buildFeatureListSection() {
        const section = div('home-manual');
        const header = div('home-manual__header');
        header.innerHTML = `
            <p class="home-choice-kicker">Manual</p>
            <h3>현재 기능 목록</h3>
            <p>현재 허브 기준 기능 목록입니다. 각 그룹을 누르면 해당 화면으로 이동합니다.</p>`;

        const list = div('home-manual-grid');
        HOME_FEATURE_GROUPS.forEach((group) => list.append(buildFeatureGroup(group)));
        section.append(header, list);
        return section;
    }

    function buildFeatureGroup(group) {
        const card = document.createElement('button');
        card.type = 'button';
        card.className = 'home-feature-group';
        card.innerHTML = `
            <div class="home-feature-group__head">
              <div>
                <strong>${group.title}</strong>
                <span>${group.desc}</span>
              </div>
              <em>${group.count}개</em>
            </div>
            <ul class="home-feature-list">
              ${group.items.map(([title, desc]) => `
                <li>
                  <strong>${title}</strong>
                  <span>${desc}</span>
                </li>`).join('')}
            </ul>`;
        card.addEventListener('click', () => {
            if (group.multiMode) setMultiMode(group.multiMode);
            location.hash = `#${group.route}`;
            if (group.anchorId) {
                setTimeout(() => {
                    const el = document.getElementById(group.anchorId);
                    if (el && el.scrollIntoView) {
                        el.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    }
                }, 240);
            }
        });
        return card;
    }

    function buildCard(title, desc, hash, items, multiMode, anchorId) {
        const card = document.createElement('button');
        card.type = 'button';
        card.className = 'home-choice-card';
        const listHtml = Array.isArray(items) && items.length
          ? `<ul class="home-choice-list">${items.map((item) => `<li>${item}</li>`).join('')}</ul>`
          : '';
        card.innerHTML = `
            <div class="home-choice-card__body">
              <div>
                <h3>${title}</h3>
                <p>${desc}</p>
                ${listHtml}
              </div>
              <span class="home-choice-card__icon">→</span>
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
        }
    }
}
