import { clear, div } from '../core/dom.js';
import { onHost, post } from '../core/bridge.js?v=20260506a';

const MANAGEMENT_URL = 'https://update.zerokky.com/index.html#kky-users';

let state = {
  enabled: true,
  currentUser: '',
  allowed: false,
  message: '외부 사용자는 사용할 수 없습니다.',
  source: '',
  sourceUrl: 'https://update.zerokky.com/kky-tool/user-access.json',
  cachePath: '',
  allowedProfileKeywords: ['KCIM'],
  allowedUsers: []
};

export function renderUserManagement(root, options = {}) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  const view = div('user-management-view');
  const header = div('user-management-header');
  header.innerHTML = `
    <div>
      <p class="user-management-kicker">KKY TOOL</p>
      <h2>사용자 관리</h2>
    </div>
    <div class="user-management-status">${state.allowed ? '허용됨' : escapeHtml(state.message)}</div>`;

  const body = div('user-management-grid');
  body.append(buildInfoPanel(), buildReadonlyPanel());
  view.append(header, body);
  target.append(view);

  if (options.internal === true) return;

  onHost('usermanagement:state', (payload) => {
    state = normalizeState(payload);
    renderUserManagement(target, { internal: true });
  });

  try { post('usermanagement:init'); } catch { }
}

function buildInfoPanel() {
  const panel = div('user-management-panel');
  panel.innerHTML = `
    <div class="user-management-panel__head">
      <h3>중앙 관리 방식</h3>
      <span>${escapeHtml(state.source || '확인 중')}</span>
    </div>`;

  const current = div('user-management-current');
  current.innerHTML = `
    <span>현재 Revit 사용자</span>
    <strong>${escapeHtml(state.currentUser || '-')}</strong>
    <small>${state.allowed ? '허용됨' : escapeHtml(state.message || '외부 사용자는 사용할 수 없습니다.')}</small>`;

  const note = document.createElement('p');
  note.className = 'user-management-note';
  note.textContent = '사용 허용 키워드는 각 PC에서 수정하지 않습니다. 업데이트 홈페이지의 KKY Tool Users 탭에서 중앙 JSON으로 관리합니다.';

  const link = document.createElement('a');
  link.className = 'btn btn--primary';
  link.href = MANAGEMENT_URL;
  link.target = '_blank';
  link.rel = 'noopener noreferrer';
  link.textContent = '홈페이지 사용자 관리 열기';

  panel.append(current, note, link);
  return panel;
}

function buildReadonlyPanel() {
  const panel = div('user-management-panel user-management-panel--wide');
  panel.innerHTML = `
    <div class="user-management-panel__head">
      <h3>현재 적용된 중앙 정책</h3>
      <span>${state.enabled ? '사용 제한 켜짐' : '사용 제한 꺼짐'}</span>
    </div>`;

  panel.append(
    readonlyField('정책 파일', state.sourceUrl || 'https://update.zerokky.com/kky-tool/user-access.json'),
    readonlyField('허용 프로필 키워드', listToText(state.allowedProfileKeywords)),
    readonlyField('예외 허용 사용자', listToText(state.allowedUsers)),
    readonlyField('차단 메시지', state.message || '외부 사용자는 사용할 수 없습니다.'),
    readonlyField('캐시 파일', state.cachePath || '-')
  );

  return panel;
}

function readonlyField(label, value) {
  const wrap = div('user-management-field');
  const text = document.createElement('label');
  text.textContent = label;
  const control = document.createElement('textarea');
  control.className = 'user-management-input user-management-textarea';
  control.value = value || '-';
  control.readOnly = true;
  wrap.append(text, control);
  return wrap;
}

function normalizeState(payload) {
  const next = { ...state };
  if (payload && typeof payload === 'object') {
    next.enabled = payload.enabled !== false;
    next.currentUser = String(payload.currentUser || '');
    next.allowed = !!payload.allowed;
    next.message = String(payload.message || '외부 사용자는 사용할 수 없습니다.');
    next.source = String(payload.source || '');
    next.sourceUrl = String(payload.sourceUrl || payload.configPath || next.sourceUrl);
    next.cachePath = String(payload.cachePath || '');
    next.allowedProfileKeywords = Array.isArray(payload.allowedProfileKeywords) ? payload.allowedProfileKeywords : ['KCIM'];
    next.allowedUsers = Array.isArray(payload.allowedUsers) ? payload.allowedUsers : [];
  }
  return next;
}

function listToText(list) {
  return (Array.isArray(list) && list.length > 0) ? list.join('\n') : '-';
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
