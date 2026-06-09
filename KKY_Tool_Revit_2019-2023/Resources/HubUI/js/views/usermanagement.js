import { clear, div, toast } from '../core/dom.js';
import { onHost, post } from '../core/bridge.js?v=20260506a';

let state = {
  authenticated: false,
  enabled: true,
  currentUser: '',
  allowed: false,
  message: '외부 사용자는 사용할 수 없습니다.',
  requirePasswordChange: false,
  configPath: '',
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
  body.append(buildLoginPanel(target), buildSettingsPanel());
  view.append(header, body);
  target.append(view);

  if (options.internal === true) return;

  onHost('usermanagement:state', (payload) => {
    state = normalizeState(payload, false);
    renderUserManagement(target, { internal: true });
  });

  onHost('usermanagement:login-result', (payload) => {
    if (!payload?.authenticated) {
      toast(payload?.message || '관리자 비밀번호가 맞지 않습니다.', 'err');
      return;
    }
    state = normalizeState(payload, true);
    toast('관리자 모드로 전환했습니다.', 'ok');
    renderUserManagement(target, { internal: true });
  });

  onHost('usermanagement:saved', (payload) => {
    state = normalizeState(payload, true);
    toast('사용자 관리 설정을 저장했습니다.', 'ok');
    renderUserManagement(target, { internal: true });
  });

  onHost('usermanagement:error', (payload) => {
    toast(payload?.message || '사용자 관리 설정 저장에 실패했습니다.', 'err');
  });

  try { post('usermanagement:init'); } catch { }
}

function buildLoginPanel(target) {
  const panel = div('user-management-panel');
  panel.innerHTML = `
    <div class="user-management-panel__head">
      <h3>관리자 인증</h3>
      <span>${state.authenticated ? '인증됨' : '잠김'}</span>
    </div>`;

  const current = div('user-management-current');
  current.innerHTML = `
    <span>현재 Revit 사용자</span>
    <strong>${escapeHtml(state.currentUser || '-')}</strong>
    <small>${state.allowed ? '허용됨' : escapeHtml(state.message || '외부 사용자는 사용할 수 없습니다.')}</small>`;

  const password = document.createElement('input');
  password.type = 'password';
  password.autocomplete = 'current-password';
  password.placeholder = state.requirePasswordChange ? '초기 비밀번호 KKYTOOL' : '관리자 비밀번호';
  password.className = 'user-management-input';

  const loginBtn = document.createElement('button');
  loginBtn.type = 'button';
  loginBtn.className = 'btn btn--primary';
  loginBtn.textContent = '인증';
  loginBtn.addEventListener('click', () => post('usermanagement:login', { password: password.value }));

  const note = document.createElement('p');
  note.className = 'user-management-note';
  note.textContent = state.requirePasswordChange
    ? '초기 비밀번호로 접속한 뒤 새 비밀번호를 저장하세요.'
    : '설정 저장 시에도 현재 관리자 비밀번호가 필요합니다.';

  panel.append(current, password, loginBtn, note);
  return panel;
}

function buildSettingsPanel() {
  const panel = div('user-management-panel user-management-panel--wide');
  panel.innerHTML = `
    <div class="user-management-panel__head">
      <h3>접근 정책</h3>
      <span>${state.enabled ? '켜짐' : '꺼짐'}</span>
    </div>`;

  const enabledRow = labelWrap('사용 제한');
  const enabled = document.createElement('input');
  enabled.type = 'checkbox';
  enabled.checked = !!state.enabled;
  enabled.disabled = !state.authenticated;
  enabledRow.append(enabled);

  const keywords = textarea('허용 프로필 키워드', listToLines(state.allowedProfileKeywords), 'KCIM');
  const users = textarea('예외 허용 사용자', listToLines(state.allowedUsers), 'Autodesk/Revit 사용자명');
  const message = input('차단 메시지', state.message || '외부 사용자는 사용할 수 없습니다.');
  const password = input('현재 관리자 비밀번호', '', 'password');
  const newPassword = input('새 관리자 비밀번호', '', 'password');

  [keywords, users, message, password, newPassword].forEach((field) => {
    field.control.disabled = !state.authenticated;
  });

  const path = div('user-management-path');
  path.textContent = state.configPath ? `설정 파일: ${state.configPath}` : '';

  const save = document.createElement('button');
  save.type = 'button';
  save.className = 'btn btn--primary';
  save.textContent = '저장';
  save.disabled = !state.authenticated;
  save.addEventListener('click', () => {
    post('usermanagement:save', {
      enabled: enabled.checked,
      allowedProfileKeywords: linesToList(keywords.control.value),
      allowedUsers: linesToList(users.control.value),
      blockMessage: message.control.value || '외부 사용자는 사용할 수 없습니다.',
      password: password.control.value,
      newPassword: newPassword.control.value
    });
  });

  panel.append(enabledRow, keywords.wrap, users.wrap, message.wrap, password.wrap, newPassword.wrap, path, save);
  return panel;
}

function normalizeState(payload, authenticatedFallback) {
  const next = { ...state };
  if (payload && typeof payload === 'object') {
    next.authenticated = !!payload.authenticated || !!authenticatedFallback;
    next.enabled = payload.enabled !== false;
    next.currentUser = String(payload.currentUser || '');
    next.allowed = !!payload.allowed;
    next.message = String(payload.message || '외부 사용자는 사용할 수 없습니다.');
    next.requirePasswordChange = !!payload.requirePasswordChange;
    next.configPath = String(payload.configPath || '');
    next.allowedProfileKeywords = Array.isArray(payload.allowedProfileKeywords) ? payload.allowedProfileKeywords : ['KCIM'];
    next.allowedUsers = Array.isArray(payload.allowedUsers) ? payload.allowedUsers : [];
  }
  return next;
}

function labelWrap(label) {
  const wrap = div('user-management-field user-management-field--inline');
  const text = document.createElement('span');
  text.textContent = label;
  wrap.append(text);
  return wrap;
}

function input(label, value, type = 'text') {
  const wrap = div('user-management-field');
  const text = document.createElement('label');
  text.textContent = label;
  const control = document.createElement('input');
  control.type = type;
  control.className = 'user-management-input';
  control.value = value || '';
  wrap.append(text, control);
  return { wrap, control };
}

function textarea(label, value, placeholder) {
  const wrap = div('user-management-field');
  const text = document.createElement('label');
  text.textContent = label;
  const control = document.createElement('textarea');
  control.className = 'user-management-input user-management-textarea';
  control.value = value || '';
  control.placeholder = placeholder || '';
  wrap.append(text, control);
  return { wrap, control };
}

function listToLines(list) {
  return (Array.isArray(list) ? list : []).join('\n');
}

function linesToList(value) {
  return String(value || '')
    .split(/[\n,;]+/g)
    .map((x) => x.trim())
    .filter(Boolean);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
