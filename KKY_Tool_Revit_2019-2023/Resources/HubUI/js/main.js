import { initTheme } from './core/theme.js?v=20260506a';
import { beginHostListenerScope, clearHostListenerScope, endHostListenerScope, onHost, post } from './core/bridge.js?v=20260506a';
import { updateTopMost, setActiveDocument, setDocList, setDocumentVisualAidSettings, setUpdateInfo, setUpdateState, renderTopbar } from './core/topbar.js?v=20260511a';
import { initLogConsole, toggleLogConsole, log } from './core/dom.js?v=20260506a';
import { renderHome } from './views/home.js?v=20260505d';
import { renderActiveMenu } from './views/activeMenu.js?v=20260504d';
import { renderDup } from './views/dup.js?v=20260505p';
import { renderConn } from './views/conn.js?v=20260505o';
import { renderExport } from './views/export.js?v=20260505e';
import { renderParamProp } from './views/paramprop.js?v=20260505c';
import { renderSharedParamBatch } from './views/sharedparambatch.js?v=20260511a';
import { renderSegmentPms } from './views/segmentpms.js?v=20260505f';
import { renderGuid } from './views/guid.js?v=20260505k';
import { renderFamilyLink } from './views/familylink.js?v=20260505f';
import { renderLinkPath } from './views/linkpath.js?v=20260505h';
import { renderMulti } from './views/multi.js?v=20260511g';
import { renderDeliveryCleaner } from './views/deliverycleaner.js?v=20260508a';
import { renderParamModifier } from './views/parammodifier.js?v=20260505j';
import { renderConditionExtract } from './views/conditionextract.js?v=20260505j';
import { renderLateralNozzle } from './views/lateralnozzle.js?v=20260505e';
import { renderTapAlign } from './views/tapalign.js?v=20260505i';

initTheme();

let _lastTop = null;
let _viewRoot = null;
let _topbarRoot = null;
let _lastHash = null;
let _historyStack = [];
let _suppressHistory = false;
let _activeViewScope = null;

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();

function boot() {
  const bootEl = document.getElementById('boot');
  if (bootEl) bootEl.remove();

  const app = document.getElementById('app');
  if (app) app.hidden = false;

  _viewRoot = document.getElementById('view-root') || app;
  _topbarRoot = document.getElementById('topbar-root') || app;
  renderTopbar(_topbarRoot, false, null);

  initLogConsole();

  window.addEventListener('keydown', (ev) => {
    const tag = ev.target && ev.target.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA') return;

    if (ev.key === 'F12') {
      ev.preventDefault();
      toggleLogConsole();
      log('Log console toggled via F12');
    }
  });

  try { post('ui:query-topmost'); } catch { }
  try { post('documentvisualaid:query-settings'); } catch { }
  try { post('update:query'); } catch { }
  window.setTimeout(() => {
    try { post('update:check', { silent: true, startup: true }); } catch { }
  }, 900);

  route();
  window.addEventListener('hashchange', route);

  onHost((msg) => {
    try {
      if (!msg || !msg.ev) return;

      try { console.log('[host] <-', msg.ev, msg.payload); } catch { }

      switch (msg.ev) {
        case 'host:topmost': {
          const on = (msg && typeof msg.payload === 'object') ? !!msg.payload.on : !!msg.payload;
          if (_lastTop === on) return;
          _lastTop = on;
          updateTopMost(on);
          break;
        }
        case 'host:doc-changed':
          setActiveDocument(msg.payload || {});
          break;
        case 'host:doc-list':
          setDocList(msg.payload);
          break;
        case 'host:update-info':
          setUpdateInfo(msg.payload || {});
          break;
        case 'host:update-state':
          setUpdateState(msg.payload || {});
          break;
        case 'host:document-visual-aid-settings':
          setDocumentVisualAidSettings(msg.payload || {});
          break;
        default:
          break;
      }
    } catch (e) {
      console.error('[main] onHost dispatch error:', e);
    }
  });

  requestHostContextSync();
  window.addEventListener('focus', requestHostContextSync);
  document.addEventListener('visibilitychange', () => {
    if (!document.hidden) requestHostContextSync();
  });
}

function requestHostContextSync() {
  try { post('ui:sync-context'); } catch { }
}

function route() {
  const hash = (location.hash || '').replace('#', '');
  try { post('ui:route-changed', { route: hash }); } catch { }

  if (!_suppressHistory && _lastHash !== null && _lastHash !== hash) {
    _historyStack.push(_lastHash);
  }
  if (_suppressHistory) _suppressHistory = false;
  if (hash === '') _historyStack = [];
  _lastHash = hash;

  const onBack = () => {
    _historyStack = [];
    location.hash = '';
  };
  const onNavBack = () => {
    if (_historyStack.length === 0) return;
    const prev = _historyStack.pop();
    _suppressHistory = true;
    location.hash = prev ? `#${prev}` : '';
  };

  const withBack = hash !== '';
  renderTopbar(_topbarRoot, withBack, hash === '' ? null : onBack, _historyStack.length > 0, onNavBack);

  if (_activeViewScope) clearHostListenerScope(_activeViewScope);
  const nextScope = `route:${hash || 'home'}`;
  _activeViewScope = nextScope;

  if (_viewRoot) _viewRoot.innerHTML = '';
  const targetRoot = _viewRoot || document.getElementById('app');

  let renderView = renderHome;
  let renderOptions = null;
  switch (hash) {
    case 'dup': renderView = renderDup; break;
    case 'conn': renderView = renderConn; break;
    case 'export': renderView = renderExport; break;
    case 'paramprop': renderView = renderParamProp; break;
    case 'sharedparambatch': renderView = renderSharedParamBatch; break;
    case 'segmentpms': renderView = renderSegmentPms; break;
    case 'guid': renderView = renderGuid; break;
    case 'familylink': renderView = renderFamilyLink; break;
    case 'linkpath': renderView = renderLinkPath; break;
    case 'multi': renderView = renderMulti; break;
    case 'favorites':
      renderView = renderMulti;
      renderOptions = { viewMode: 'favorites' };
      break;
    case 'deliverycleaner': renderView = renderDeliveryCleaner; break;
    case 'parammodifier': renderView = renderParamModifier; break;
    case 'conditionextract': renderView = renderConditionExtract; break;
    case 'lateralnozzle': renderView = renderLateralNozzle; break;
    case 'tapalign': renderView = renderTapAlign; break;
    case 'active-menu': renderView = renderActiveMenu; break;
    default: renderView = renderHome; break;
  }

  beginHostListenerScope(nextScope);
  try {
    if (renderOptions) return renderView(targetRoot, renderOptions);
    return renderView(targetRoot);
  } finally {
    endHostListenerScope(nextScope);
  }
}
