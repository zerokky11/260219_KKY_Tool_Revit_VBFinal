// Resources/HubUI/js/core/topbar.js
import { div, toast } from './dom.js';
import { toggleTheme } from './theme.js';
import { setConn, ping, post } from './bridge.js';

const APP_VERSION_FALLBACK = 'v2.14';
const STARTUP_NOTICE_DURATION_MS = 4800;
const REQUESTS_PAGE_URL = 'https://update.zerokky.com/requests.html';

const TEXT = {
    tagline: '\u0052\u0065\u0076\u0069\u0074 \uc6cc\ud06c\ud50c\ub85c\uc6b0\ub97c \ud558\ub098\uc758 \ud5c8\ube0c\uc5d0\uc11c \uad00\ub9ac\ud558\uc138\uc694.',
    home: '\ud5c8\ube0c \ud648\uc73c\ub85c',
    back: '\ub4a4\ub85c\uac00\uae30',
    subtitle: '\u0052\u0065\u0076\u0069\u0074 \uc791\uc5c5 \ubcf4\uc870 \ud1b5\ud569 \ub3c4\uad6c',
    connected: '\uc5f0\uacb0\ub428',
    connectedNone: '\uc5f0\uacb0 \uc548\ub428',
    pin: '\ud56d\uc0c1 \uc704',
    theme: '\ud14c\ub9c8',
    request: '\uc694\uccad\ud558\uae30',
    updateCheck: '\u0054\u006f\u006f\u006c \ubc84\uc804 \uccb4\ud06c',
    updateChecking: '\u0054\u006f\u006f\u006c \ubc84\uc804 \ud655\uc778 \uc911',
    updateHint: '\u0054\u006f\u006f\u006c \ubc84\uc804 \uccb4\ud06c\ub97c \ub20c\ub7ec \ucd5c\uc2e0 \ubc84\uc804\uc744 \uc124\uce58\ud558\uc138\uc694.',
    startupToast: '\uc0c8 \u0054\u006f\u006f\u006c \ubc84\uc804\uc774 \uc788\uc2b5\ub2c8\ub2e4. \u0054\u006f\u006f\u006c \ubc84\uc804 \uccb4\ud06c\ub97c \ub20c\ub7ec \ud655\uc778\ud574 \uc8fc\uc138\uc694.',
    currentVersion: '\ud604\uc7ac \ubc84\uc804',
    latestVersion: '\ucd5c\uc2e0 \ubc84\uc804',
    feedUrl: '\uc5c5\ub370\uc774\ud2b8 \ud655\uc778 \uc8fc\uc18c',
    publishedAt: '\ubc30\ud3ec\uc77c',
    releaseNotes: '\ubcc0\uacbd \uc0ac\ud56d',
    dialogTitle: '\u0054\u006f\u006f\u006c \ubc84\uc804 \ud655\uc778',
    dialogErrorTitle: '\u0054\u006f\u006f\u006c \ubc84\uc804 \ud655\uc778 \uc2e4\ud328',
    dialogUpdateTitle: '\uc0c8 \u0054\u006f\u006f\u006c \ubc84\uc804\uc774 \uc788\uc2b5\ub2c8\ub2e4',
    dialogLatestTitle: '\ud604\uc7ac \ucd5c\uc2e0 \ubc84\uc804\uc785\ub2c8\ub2e4',
    dialogDownloadTitle: '\ucd5c\uc2e0 \uc124\uce58 \ud30c\uc77c \ub2e4\uc6b4\ub85c\ub4dc \uc911',
    dialogReadyTitle: '\uc124\uce58 \uc900\ube44 \uc644\ub8cc',
    statusDone: '\ud655\uc778 \uc644\ub8cc',
    statusError: '\ud655\uc778 \uc2e4\ud328',
    statusLatest: '\ucd5c\uc2e0 \uc0c1\ud0dc',
    statusUpdate: '\uc5c5\ub370\uc774\ud2b8 \ud544\uc694',
    statusDownloading: '\ub2e4\uc6b4\ub85c\ub4dc \uc9c4\ud589 \uc911',
    statusReady: '\uc124\uce58 \uc900\ube44\ub428',
    currentLatestSummary: '\ud604\uc7ac \ubc84\uc804 {current}\uc774 \ucd5c\uc2e0 \ubc84\uc804\uc785\ub2c8\ub2e4.',
    updateSummary: '\ud604\uc7ac \ubc84\uc804 {current}\uc5d0\uc11c \ucd5c\uc2e0 \ubc84\uc804 {latest}\ub85c \uc5c5\ub370\uc774\ud2b8\ud560 \uc218 \uc788\uc2b5\ub2c8\ub2e4.',
    downloadSummary: '\ucd5c\uc2e0 \uc5c5\ub370\uc774\ud2b8 \ud328\ud0a4\uc9c0\ub97c \uc784\uc2dc \ud3f4\ub354\ub85c \ub2e4\uc6b4\ub85c\ub4dc\ud558\uace0 \uc788\uc2b5\ub2c8\ub2e4.',
    readySummary: '\uc5c5\ub370\uc774\ud2b8 \ud328\ud0a4\uc9c0 \ub2e4\uc6b4\ub85c\ub4dc\uac00 \uc644\ub8cc\ub418\uc5c8\uc2b5\ub2c8\ub2e4. \ud604\uc7ac \uc2e4\ud589 \uc911\uc778 \u0052\u0065\u0076\u0069\u0074\uc744 \ubaa8\ub450 \uc885\ub8cc\ud558\uba74 \uc5c5\ub370\uc774\ud2b8\uac00 \uc801\uc6a9\ub429\ub2c8\ub2e4. \uc801\uc6a9 \ud6c4 \u0052\u0065\u0076\u0069\u0074\uc744 \ub2e4\uc2dc \uc2e4\ud589\ud574 \uc8fc\uc138\uc694.',
    preparingDownload: '\uc5c5\ub370\uc774\ud2b8 \ud328\ud0a4\uc9c0 \ub2e4\uc6b4\ub85c\ub4dc\ub97c \uc900\ube44\ud558\ub294 \uc911\uc785\ub2c8\ub2e4.',
    installing: '\uc124\uce58\ud558\uae30',
    close: '\ub2eb\uae30',
    shortcutsTitle: '\ub3c4\uc6c0\ub9d0 - \u004b\u004b\u0059 \u0054\u006f\u006f\u006c \u0048\u0075\u0062',
    shortcutsHtml: `
        <div><strong>\ub2e8\ucd95\ud0a4</strong></div>
        <ul>
            <li><code>/</code> \uac80\uc0c9 \ud3ec\ucee4\uc2a4</li>
            <li>\uce74\ub4dc \uc120\ud0dd \ud6c4 <code>Enter</code> \uc2e4\ud589, <code>F</code> \uc990\uaca8\ucc3e\uae30</li>
            <li><code>Ctrl</code>+<code>Shift</code>+<code>L</code> \ud14c\ub9c8 \uc804\ud658</li>
            <li><code>Ctrl</code>+<code>Shift</code>+<code>T</code> \ud56d\uc0c1 \uc704</li>
            <li><code>F1</code> \ub610\ub294 <code>?</code> \ub3c4\uc6c0\ub9d0</li>
        </ul>
    `
};

let _docList = [];
let _activeDoc = { name: '', path: '' };
let _topbarEl = null;
let _backBtn = null;
let _backHandler = null;
let _navBackBtn = null;
let _navWrap = null;
let _progressWrap = null;
let _progressFill = null;
let _progressText = null;
let _progressPct = null;
let _versionEl = null;
let _updateBtn = null;
let _updateHintEl = null;
let _updateDialogBackdrop = null;
let _updateStartupNoticeShown = false;
let _updateNeedsAttention = false;
let _updateDismissTimer = 0;
let _lastReadyNoticeKey = '';
let _updateState = {
    currentVersion: APP_VERSION_FALLBACK.replace(/^v/i, ''),
    currentVersionDisplay: APP_VERSION_FALLBACK.replace(/^v/i, ''),
    latestVersion: '',
    hasUpdate: false,
    canInstall: false,
    isConfigured: false,
    busy: false,
    message: '',
    kind: 'info',
    configPath: '',
    feedUrl: '',
    notes: '',
    publishedAt: ''
};

function readUpdateField(payload, camelName, pascalName = '') {
    if (!payload || typeof payload !== 'object') return undefined;
    if (Object.prototype.hasOwnProperty.call(payload, camelName)) return payload[camelName];
    if (pascalName && Object.prototype.hasOwnProperty.call(payload, pascalName)) return payload[pascalName];
    return undefined;
}

export function renderTopbar(root, withBack = false, onBack = null, canGoBack = false, onNavBack = null) {
    const host = document.getElementById('topbar-root') || root;
    if (!host) return;

    if (_topbarEl) {
        if (!_topbarEl.parentElement) host.append(_topbarEl);
    } else {
        _topbarEl = div('topbar');
        const left = div('topbar-left');
        const center = div('topbar-center');
        const right = div('topbar-right');

        const tagline = document.createElement('p');
        tagline.className = 'topbar-tagline';
        tagline.textContent = TEXT.tagline;
        center.append(tagline);

        _topbarEl.append(left, center, right);
        host.append(_topbarEl);

        _navWrap = div('topbar-nav');
        left.append(_navWrap);
        buildBrand(left);
        renderTopbarChips();

        _progressWrap = div('topbar-progress hidden');
        const progRow = div('topbar-progress-row');
        _progressText = div('topbar-progress-text');
        _progressPct = div('topbar-progress-pct');
        progRow.append(_progressText, _progressPct);

        const progBar = div('topbar-progress-bar');
        _progressFill = div('topbar-progress-fill');
        _progressFill.style.width = '0%';
        progBar.append(_progressFill);

        _progressWrap.append(progRow, progBar);
        _topbarEl.append(_progressWrap);
    }

    configureBackButton(withBack, onBack, canGoBack, onNavBack);
    setConn(true);
    applyActiveDocumentState();
    applyUpdateVisualState();
}

function configureBackButton(withBack, onBack, canGoBack, onNavBack) {
    const left = _topbarEl?.querySelector('.topbar-left');
    if (!left) return;

    if (!_navWrap) {
        _navWrap = div('topbar-nav');
        left.prepend(_navWrap);
    }

    if (!_backBtn) {
        _backBtn = createNavButton(TEXT.home, {
            iconSvg: homeIconSvg()
        });
        _navWrap.append(_backBtn);
    }

    if (!_navBackBtn) {
        _navBackBtn = createNavButton(TEXT.back, {
            iconSvg: backIconSvg()
        });
        _navWrap.append(_navBackBtn);
    }

    _backHandler = onBack;

    const smartGoHome = () => {
        try { window.dispatchEvent(new CustomEvent('kkyt:go-home')); } catch (_) { }
        const before = location.href;
        try { history.back(); } catch (_) { }
        setTimeout(() => {
            if (location.href === before) {
                const url = new URL(location.href);
                const parts = url.pathname.split('/');
                parts[parts.length - 1] = 'index.html';
                url.pathname = parts.join('/');
                location.href = url.toString();
            }
        }, 80);
    };

    _backBtn.onclick = () => {
        if (typeof _backHandler === 'function') {
            try { _backHandler(); } catch (_) { smartGoHome(); }
        } else {
            smartGoHome();
        }
    };
    _backBtn.classList.toggle('hidden', !withBack);

    _navBackBtn.disabled = !canGoBack;
    _navBackBtn.classList.toggle('is-disabled', !canGoBack);
    _navBackBtn.onclick = () => {
        if (!canGoBack) return;
        if (typeof onNavBack === 'function') onNavBack();
    };
}

function createNavButton(labelText, { iconSrc = '', iconAlt = '', iconSvg = '' } = {}) {
    const btn = document.createElement('button');
    btn.className = 'btn btn-ghost';
    btn.type = 'button';

    let icon = null;
    if (iconSvg) {
        icon = document.createElement('span');
        icon.className = 'back-btn-glyph';
        icon.setAttribute('aria-hidden', 'true');
        icon.innerHTML = iconSvg;
    } else {
        icon = document.createElement('img');
        icon.className = 'back-btn-icon';
        icon.src = iconSrc;
        icon.alt = iconAlt;
    }

    const label = document.createElement('span');
    label.className = 'back-btn-label';
    label.textContent = labelText;

    btn.append(icon, label);
    return btn;
}

function homeIconSvg() {
    return `
        <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M4.75 11.25L12 5.25L19.25 11.25V18.25C19.25 18.8023 18.8023 19.25 18.25 19.25H5.75C5.19772 19.25 4.75 18.8023 4.75 18.25V11.25Z" stroke="currentColor" stroke-width="2.1" stroke-linejoin="round"/>
            <path d="M9.25 19.25V13.9C9.25 13.3477 9.69772 12.9 10.25 12.9H13.75C14.3023 12.9 14.75 13.3477 14.75 13.9V19.25" stroke="currentColor" stroke-width="2.1" stroke-linejoin="round"/>
            <path d="M7.2 10.2H8.8" stroke="currentColor" stroke-width="2.1" stroke-linecap="round"/>
            <circle cx="17.6" cy="7.2" r="1.6" fill="currentColor" opacity="0.22"/>
            <circle cx="17.6" cy="7.2" r="0.85" fill="currentColor"/>
        </svg>
    `;
}

function backIconSvg() {
    return `
        <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M10 6L4 12L10 18" stroke="currentColor" stroke-width="2.25" stroke-linecap="round" stroke-linejoin="round"/>
            <path d="M5 12H20" stroke="currentColor" stroke-width="2.25" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
    `;
}

function buildBrand(host) {
    const wrap = div('topbar-brand');

    const logo = document.createElement('span');
    logo.className = 'topbar-logo';
    const logoImg = document.createElement('img');
    logoImg.src = 'assets/icons/KKY_Tool_48.png';
    logoImg.alt = 'KKY Tool';
    logo.append(logoImg);

    const text = document.createElement('div');
    text.className = 'topbar-brand-text';
    text.innerHTML = `<strong>KKY Tool Hub</strong><span>${TEXT.subtitle}</span>`;

    wrap.append(logo, text);
    host.append(wrap);
}

export function renderTopbarChips() {
    const actions = document.querySelector('.topbar-right');
    if (!actions) return;

    actions.innerHTML = '';

    const conn = createControlButton({
        id: 'connChip',
        label: TEXT.connected,
        icon: 'plug',
        classes: 'chip-toggle chip-connection',
        statusDot: true
    });
    conn.addEventListener('click', ping);
    actions.append(conn);

    const chipRow = document.createElement('div');
    chipRow.className = 'chip-row';

    const versionChip = document.createElement('span');
    versionChip.className = 'topbar-version topbar-version--inline';
    versionChip.textContent = APP_VERSION_FALLBACK;
    _versionEl = versionChip;
    chipRow.append(versionChip);

    const updateWrap = div('update-chip-wrap');
    const updateBtn = createControlButton({
        id: 'updateChip',
        label: TEXT.updateCheck,
        icon: 'update',
        classes: 'chip-btn update-chip'
    });
    updateBtn.addEventListener('click', onUpdateButtonClick);
    _updateBtn = updateBtn;

    const updateHint = div('update-chip-hint hidden');
    updateHint.textContent = TEXT.updateHint;
    _updateHintEl = updateHint;
    updateWrap.append(updateBtn, updateHint);
    chipRow.append(updateWrap);

    const themeBtn = createControlButton({
        label: TEXT.theme,
        icon: 'theme',
        classes: 'chip-btn theme-chip'
    });
    themeBtn.setAttribute('aria-pressed', 'false');
    const applyThemeState = () => {
        const cur = document.documentElement.dataset.theme || 'dark';
        themeBtn.classList.toggle('is-dark', cur === 'dark');
        themeBtn.classList.toggle('is-active', cur === 'dark');
        themeBtn.setAttribute('aria-pressed', cur === 'dark' ? 'true' : 'false');
    };
    themeBtn.onclick = () => {
        toggleTheme();
        applyThemeState();
    };
    applyThemeState();
    chipRow.append(themeBtn);

    const pin = createControlButton({
        id: 'pinChip',
        label: TEXT.pin,
        icon: 'pin',
        classes: 'chip-toggle pin-chip'
    });
    pin.setAttribute('aria-pressed', 'false');
    pin.classList.add('is-off');
    pin.onclick = () => {
        try { post('ui:toggle-topmost'); } catch (e) { console.error(e); }
    };
    chipRow.append(pin);

    const requestBtn = createControlButton({
        label: TEXT.request,
        icon: 'request',
        classes: 'chip-btn request-chip'
    });
    requestBtn.addEventListener('click', openRequestsPage);
    chipRow.append(requestBtn);

    actions.append(chipRow);

    applyActiveDocumentState();
    applyUpdateVisualState();
}

export function updateTopMost(on) {
    const pin = document.querySelector('#pinChip');
    if (!pin) return;
    const active = !!on;
    pin.classList.toggle('is-active', active);
    pin.classList.toggle('is-off', !active);
    pin.setAttribute('aria-pressed', active ? 'true' : 'false');
    const label = pin.querySelector('.chip-text');
    if (label) label.textContent = TEXT.pin;
}

export function setTopbarProgress(state) {
    if (!_progressWrap || !_progressFill || !_progressText || !_progressPct) return;
    if (!state) {
        _progressWrap.classList.add('hidden');
        _progressFill.style.width = '0%';
        _progressText.textContent = '';
        _progressPct.textContent = '';
        return;
    }

    const total = Math.max(0, Number(state.total) || 0);
    const idx = Math.max(0, Number(state.index) || 0);
    const pct = Math.max(0, Math.min(100, Number(state.percent) || 0));
    const file = state.file || '';
    const msg = state.message || '';
    _progressWrap.classList.remove('hidden');
    _progressFill.style.width = `${pct}%`;
    _progressText.textContent = `(${idx}/${total || '?'}) ${file} - ${msg}`;
    _progressPct.textContent = `${Math.round(pct)}%`;
}

function applyActiveDocumentState() {
    const label = document.querySelector('#connChip .chip-text');
    if (!label) return;
    label.textContent = _activeDoc.name ? `${TEXT.connected} \u00b7 ${_activeDoc.name}` : TEXT.connectedNone;
}

export function setActiveDocument(doc = {}) {
    _activeDoc = {
        name: doc?.name || '',
        path: doc?.path || ''
    };
    applyActiveDocumentState();
}

export function setDocList(payload) {
    let list = [];
    if (Array.isArray(payload?.docs)) list = payload.docs;
    if (Array.isArray(payload)) list = payload;
    _docList = list;
}

export function setUpdateInfo(payload = {}) {
    const incomingHasUpdate = !!readUpdateField(payload, 'hasUpdate', 'HasUpdate');
    const currentVersion = readUpdateField(payload, 'currentVersion', 'CurrentVersion');
    const currentVersionDisplay = readUpdateField(payload, 'currentVersionDisplay', 'CurrentVersionDisplay');
    const latestVersion = readUpdateField(payload, 'latestVersion', 'LatestVersion');
    const canInstall = readUpdateField(payload, 'canInstall', 'CanInstall');
    const isConfigured = readUpdateField(payload, 'isConfigured', 'IsConfigured');
    const configPath = readUpdateField(payload, 'configPath', 'ConfigPath');
    const feedUrl = readUpdateField(payload, 'feedUrl', 'FeedUrl');
    const notes = readUpdateField(payload, 'notes', 'Notes');
    const publishedAt = readUpdateField(payload, 'publishedAt', 'PublishedAt');
    const message = readUpdateField(payload, 'message', 'Message');

    _updateState = {
        ..._updateState,
        currentVersion: currentVersion || _updateState.currentVersion,
        currentVersionDisplay: currentVersionDisplay || currentVersion || _updateState.currentVersionDisplay,
        latestVersion: latestVersion || '',
        hasUpdate: incomingHasUpdate,
        canInstall: !!canInstall,
        isConfigured: !!isConfigured,
        configPath: configPath || '',
        feedUrl: feedUrl || '',
        notes: notes || '',
        publishedAt: publishedAt || '',
        message: message || _updateState.message
    };

    if (!incomingHasUpdate) {
        _updateStartupNoticeShown = false;
        _updateNeedsAttention = false;
        clearUpdateDismissTimer();
    } else {
        _updateNeedsAttention = true;
    }

    applyUpdateVisualState();
}

export function setUpdateState(payload = {}) {
    const hasUpdate = readUpdateField(payload, 'hasUpdate', 'HasUpdate');
    const latestVersion = readUpdateField(payload, 'latestVersion', 'LatestVersion');
    const currentVersionDisplay = readUpdateField(payload, 'currentVersionDisplay', 'CurrentVersionDisplay');
    const message = readUpdateField(payload, 'message', 'Message');
    const kind = readUpdateField(payload, 'kind', 'Kind');
    const busy = readUpdateField(payload, 'busy', 'Busy');

    _updateState = {
        ..._updateState,
        busy: !!busy,
        hasUpdate: typeof hasUpdate === 'boolean' ? hasUpdate : _updateState.hasUpdate,
        latestVersion: latestVersion || _updateState.latestVersion,
        currentVersionDisplay: currentVersionDisplay || _updateState.currentVersionDisplay,
        message: message || _updateState.message,
        kind: kind || _updateState.kind
    };

    if (_updateState.hasUpdate && !_updateNeedsAttention) {
        _updateNeedsAttention = true;
    }

    if (payload?.startupNotice && _updateState.hasUpdate) {
        showUpdateStartupNotice();
    }

    applyUpdateVisualState();

    if (payload?.phase === 'download' || payload?.phase === 'ready') {
        if (payload?.phase === 'ready') {
            showReadyRestartNotice(payload);
        }
        showUpdateResultDialog(payload);
        return;
    }

    if (payload?.showToast && payload?.message) {
        showUpdateResultDialog(payload);
    }
}

function showReadyRestartNotice(payload = {}) {
    const key = [
        payload?.phase || '',
        payload?.installerPath || '',
        payload?.scriptPath || '',
        _updateState.latestVersion || '',
        _updateState.currentVersionDisplay || ''
    ].join('|');

    if (key && key === _lastReadyNoticeKey) return;
    _lastReadyNoticeKey = key;

    const message = payload?.message || TEXT.readySummary;
    toast(message, 'warn', 5600);
}

function toggleHelpPanel(trigger) {
    const existing = document.querySelector('.settings-backdrop');
    if (existing) {
        if (existing._escListener) {
            document.removeEventListener('keydown', existing._escListener);
        }
        existing.remove();
        if (trigger) trigger.setAttribute('aria-expanded', 'false');
        return;
    }

    const backdrop = document.createElement('div');
    backdrop.className = 'settings-backdrop';

    const panel = document.createElement('section');
    panel.className = 'settings-panel';

    const header = document.createElement('header');
    const headerText = document.createElement('span');
    headerText.textContent = TEXT.shortcutsTitle;
    header.append(headerText);

    const body = document.createElement('div');
    body.className = 'body';
    body.innerHTML = TEXT.shortcutsHtml;

    const actions = document.createElement('div');
    actions.className = 'actions';
    const closeBtn = document.createElement('button');
    closeBtn.className = 'btn';
    closeBtn.type = 'button';
    closeBtn.textContent = TEXT.close;

    const closePanel = () => {
        backdrop.remove();
        if (trigger) trigger.setAttribute('aria-expanded', 'false');
        document.removeEventListener('keydown', escListener);
    };

    const escListener = (e) => {
        if (e.key === 'Escape') closePanel();
    };

    document.addEventListener('keydown', escListener);
    backdrop._escListener = escListener;
    closeBtn.onclick = closePanel;
    actions.append(closeBtn);

    panel.append(header, body, actions);
    backdrop.append(panel);
    document.body.append(backdrop);
    trigger?.setAttribute('aria-expanded', 'true');

    backdrop.addEventListener('click', (e) => {
        if (e.target === backdrop) closePanel();
    });
}

function openRequestsPage() {
    try {
        post('ui:open-external', { url: REQUESTS_PAGE_URL });
    } catch (_) {
        try {
            window.open(REQUESTS_PAGE_URL, '_blank', 'noopener,noreferrer');
        } catch (_) { }
    }
}

function createControlButton({ id, label, icon, classes = '', statusDot = false }) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `control-chip ${classes}`.trim();
    if (id) btn.id = id;

    if (statusDot) {
        btn.classList.add('has-status-dot');
        const status = document.createElement('span');
        status.className = 'status-dot';
        status.setAttribute('aria-hidden', 'true');
        btn.append(status);
    }

    const glyph = document.createElement('span');
    glyph.className = 'chip-glyph';
    glyph.setAttribute('aria-hidden', 'true');
    glyph.innerHTML = iconSvg(icon);

    const text = document.createElement('span');
    text.className = 'chip-text';
    text.textContent = label;

    btn.append(glyph, text);
    return btn;
}

function iconSvg(name) {
    switch (name) {
        case 'plug':
            return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M8 3v5m8-5v5" stroke-linecap="round"/><path d="M6 8h12v5a6 6 0 1 1-12 0Z" stroke-linejoin="round"/><path d="M12 18v3" stroke-linecap="round"/></svg>';
        case 'pin':
            return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M8 4h8l-1 5 3 3-6 6-6-6 3-3z" fill="none"/><path d="M12 18v4" stroke-linecap="round"/></svg>';
        case 'theme':
            return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path class="half-dark" d="M12 3a9 9 0 0 1 0 18V3Z" fill="currentColor"/><path class="half-light" d="M12 21a9 9 0 0 1 0-18v18Z" fill="currentColor"/><circle cx="12" cy="12" r="8.5" fill="none"/></svg>';
        case 'update':
            return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M12 4v6m0 0 2.5-2.5M12 10 9.5 7.5" stroke-linecap="round" stroke-linejoin="round"/><path d="M7 12a5 5 0 1 0 1.46-3.54" fill="none" stroke-linecap="round" stroke-linejoin="round"/><path d="M12 20v-2" stroke-linecap="round"/></svg>';
        case 'gear':
            return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M20 13.5v-3l-2.1-.6a6.1 6.1 0 0 0-.6-1.4l1.2-1.8-2.1-2.1-1.8 1.2a6.1 6.1 0 0 0-1.4-.6L13.5 2h-3l-.6 2.1a6.1 6.1 0 0 0-1.4.6L6.7 3.5 4.6 5.6l1.2 1.8c-.26.44-.47.91-.6 1.4L3 10.5v3l2.1.6c.13.49.34.96.6 1.4l-1.2 1.8 2.1 2.1 1.8-1.2c.44.26.91.47 1.4.6l.6 2.1h3l.6-2.1c.49-.13.96-.34 1.4-.6l1.8 1.2 2.1-2.1-1.2-1.8c.26-.44.47-.91.6-1.4Z" fill="none"/><circle cx="12" cy="12" r="3.2" fill="none"/></svg>';
        case 'request':
            return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M6 7.25A2.25 2.25 0 0 1 8.25 5h7.5A2.25 2.25 0 0 1 18 7.25v4.5A2.25 2.25 0 0 1 15.75 14h-4.1l-3.15 2.6V14.9A2.25 2.25 0 0 1 6 12.75v-5.5Z" stroke-linejoin="round"/><path d="M9 8.75h6M9 11.25h4" stroke-linecap="round"/></svg>';
        default:
            return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><circle cx="12" cy="12" r="8"/></svg>';
    }
}

function onUpdateButtonClick() {
    if (_updateState.busy) return;
    _updateNeedsAttention = false;
    clearUpdateDismissTimer();
    applyUpdateVisualState();

    if (_updateState.hasUpdate) {
        showUpdateResultDialog({
            kind: 'warn',
            message: _updateState.message
        });
        return;
    }

    post('update:check');
}

function buildUpdateTooltip() {
    const lines = [`${TEXT.currentVersion}: ${formatVersionText(_updateState.currentVersionDisplay)}`];
    if (_updateState.latestVersion) {
        lines.push(`${TEXT.latestVersion}: ${formatVersionText(_updateState.latestVersion)}`);
    }
    if (_updateState.hasUpdate && _updateState.latestVersion) {
        lines.push(`${formatVersionText(_updateState.latestVersion)} 업데이트가 있습니다.`);
    } else if (!_updateState.hasUpdate) {
        lines.push(TEXT.currentLatestSummary.replace('{current}', formatVersionText(_updateState.currentVersionDisplay)));
    }
    return lines.join('\n');
}

function showUpdateResultDialog(payload = {}) {
    closeUpdateResultDialog();

    const current = formatVersionText(_updateState.currentVersionDisplay);
    const latest = _updateState.latestVersion ? formatVersionText(_updateState.latestVersion) : '-';
    const hasUpdate = !!_updateState.hasUpdate;
    const installReady = !!payload?.installerPath || !!payload?.scriptPath;
    const isError = payload?.kind === 'err';
    const phase = payload?.phase || '';
    const isDownloading = phase === 'download';
    const progressPercent = Math.max(0, Math.min(100, Number(payload?.progressPercent) || 0));
    const progressMessage = TEXT.preparingDownload;

    let tone = 'info';
    let title = TEXT.dialogTitle;
    let summary = payload?.message || '';
    let statusText = TEXT.statusDone;

    if (isError) {
        tone = 'err';
        title = TEXT.dialogErrorTitle;
        statusText = TEXT.statusError;
    } else if (isDownloading) {
        tone = 'info';
        title = TEXT.dialogDownloadTitle;
        statusText = TEXT.statusDownloading;
        summary = TEXT.downloadSummary;
    } else if (installReady) {
        tone = 'ok';
        title = TEXT.dialogReadyTitle;
        statusText = TEXT.statusReady;
        summary = TEXT.readySummary;
    } else if (hasUpdate) {
        tone = 'warn';
        title = TEXT.dialogUpdateTitle;
        statusText = TEXT.statusUpdate;
        summary = TEXT.updateSummary.replace('{current}', current).replace('{latest}', latest);
    } else {
        tone = 'ok';
        title = TEXT.dialogLatestTitle;
        statusText = TEXT.statusLatest;
        summary = TEXT.currentLatestSummary.replace('{current}', current);
    }

    const backdrop = document.createElement('div');
    backdrop.className = 'update-result-backdrop';

    const dialog = document.createElement('section');
    dialog.className = `update-result-dialog ${tone}`.trim();
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    dialog.setAttribute('aria-label', title);

    const header = document.createElement('header');
    header.className = 'update-result-header';
    header.innerHTML = `<div class="update-result-badge ${tone}">${statusText}</div><h3>${title}</h3>`;

    const body = document.createElement('div');
    body.className = 'update-result-body';

    const desc = document.createElement('p');
    desc.className = 'update-result-desc';
    desc.textContent = summary;
    body.append(desc);

    const grid = document.createElement('div');
    grid.className = 'update-result-grid';
    grid.innerHTML = `
        <div class="update-result-item">
            <span class="update-result-item-label">${TEXT.currentVersion}</span>
            <strong class="update-result-item-value">${current}</strong>
        </div>
        <div class="update-result-item">
            <span class="update-result-item-label">${TEXT.latestVersion}</span>
            <strong class="update-result-item-value">${latest}</strong>
        </div>
    `;
    body.append(grid);

    if (isDownloading) {
        const progress = document.createElement('div');
        progress.className = 'update-result-progress';
        progress.innerHTML = `
            <div class="update-result-progress-meta">
                <span>${progressMessage}</span>
                <strong>${progressPercent}%</strong>
            </div>
            <div class="update-result-progress-bar">
                <div class="update-result-progress-fill" style="width:${progressPercent}%"></div>
            </div>
        `;
        body.append(progress);
    }

    const notes = [];
    if (_updateState.publishedAt) notes.push(`${TEXT.publishedAt}: ${_updateState.publishedAt}`);
    if (_updateState.notes) notes.push(`${TEXT.releaseNotes}: ${_updateState.notes}`);
    if (_updateState.feedUrl) notes.push(`${TEXT.feedUrl}: ${_updateState.feedUrl}`);
    notes.forEach((noteText) => {
        const note = document.createElement('div');
        note.className = 'update-result-note';
        note.textContent = noteText;
        body.append(note);
    });

    const footer = document.createElement('div');
    footer.className = 'update-result-footer';

    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'btn btn-ghost';
    closeBtn.textContent = TEXT.close;
    closeBtn.addEventListener('click', closeUpdateResultDialog);
    footer.append(closeBtn);

    if (!isError && hasUpdate && _updateState.canInstall && !installReady && !isDownloading) {
        const installBtn = document.createElement('button');
        installBtn.type = 'button';
        installBtn.className = 'btn';
        installBtn.textContent = TEXT.installing;
        installBtn.addEventListener('click', () => {
            showUpdateResultDialog({
                kind: 'info',
                phase: 'download',
                progressPercent: 0,
                progressMessage: TEXT.preparingDownload,
                message: TEXT.downloadSummary
            });
            post('update:install', { version: _updateState.latestVersion });
        });
        footer.append(installBtn);
    }

    dialog.append(header, body, footer);
    backdrop.append(dialog);
    backdrop.addEventListener('click', (e) => {
        if (e.target === backdrop) closeUpdateResultDialog();
    });

    const escHandler = (e) => {
        if (e.key === 'Escape') closeUpdateResultDialog();
    };
    backdrop._escHandler = escHandler;
    document.addEventListener('keydown', escHandler);

    document.body.append(backdrop);
    _updateDialogBackdrop = backdrop;
}

function closeUpdateResultDialog() {
    if (!_updateDialogBackdrop) return;
    if (_updateDialogBackdrop._escHandler) {
        document.removeEventListener('keydown', _updateDialogBackdrop._escHandler);
    }
    _updateDialogBackdrop.remove();
    _updateDialogBackdrop = null;
}

function clearUpdateDismissTimer() {
    if (_updateDismissTimer) {
        window.clearTimeout(_updateDismissTimer);
        _updateDismissTimer = 0;
    }
}

function applyUpdateVisualState() {
    if (_versionEl) {
        _versionEl.textContent = formatVersionText(_updateState.currentVersionDisplay);
    }

    if (_updateHintEl) {
        const hintVisible = _updateState.hasUpdate && !_updateState.busy && _updateNeedsAttention;
        _updateHintEl.classList.toggle('hidden', !hintVisible);
        if (hintVisible) {
            const latest = _updateState.latestVersion ? formatVersionText(_updateState.latestVersion) : '';
            _updateHintEl.textContent = latest
                ? `${latest} 업데이트가 있습니다. Tool 버전 체크를 눌러 확인하세요.`
                : TEXT.updateHint;
        }
    }

    if (!_updateBtn) return;

    const text = _updateBtn.querySelector('.chip-text');
    if (text) {
        text.textContent = _updateState.busy ? TEXT.updateChecking : TEXT.updateCheck;
    }

    _updateBtn.disabled = !!_updateState.busy;
    _updateBtn.classList.toggle('is-busy', !!_updateState.busy);
    _updateBtn.classList.toggle('has-update', !!_updateState.hasUpdate && !_updateState.busy);
    _updateBtn.classList.toggle('needs-attention', !!_updateNeedsAttention && !!_updateState.hasUpdate && !_updateState.busy);
    _updateBtn.title = buildUpdateTooltip();
}

function showUpdateStartupNotice() {
    if (_updateStartupNoticeShown) return;
    _updateStartupNoticeShown = true;
    _updateNeedsAttention = true;
    applyUpdateVisualState();
    clearUpdateDismissTimer();
    _updateDismissTimer = window.setTimeout(() => {
        _updateNeedsAttention = true;
        applyUpdateVisualState();
    }, STARTUP_NOTICE_DURATION_MS);
}

function formatVersionText(versionText) {
    const value = String(versionText || '').trim();
    if (!value) return APP_VERSION_FALLBACK;
    return value.toLowerCase().startsWith('v') ? value : `v${value}`;
}














