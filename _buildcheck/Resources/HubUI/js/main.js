import { initTheme } from './core/theme.js';
import { onHost, post } from './core/bridge.js';
import { updateTopMost, setActiveDocument, setDocList, setUpdateInfo, setUpdateState, renderTopbar } from './core/topbar.js';
import { initLogConsole, toggleLogConsole, log } from './core/dom.js';
import { renderHome } from './views/home.js';
import { renderActiveMenu } from './views/activeMenu.js';
import { renderDup } from './views/dup.js';
import { renderConn } from './views/conn.js';
import { renderExport } from './views/export.js';
import { renderParamProp } from './views/paramprop.js';
import { renderSharedParamBatch } from './views/sharedparambatch.js';
import { renderSegmentPms } from './views/segmentpms.js';
import { renderGuid } from './views/guid.js';
import { renderFamilyLink } from './views/familylink.js';
import { renderMulti } from './views/multi.js?v=20260323j';
import { renderDeliveryCleaner } from './views/deliverycleaner.js';
import { renderParamModifier } from './views/parammodifier.js';

initTheme();

let _lastTop = null;
let _viewRoot = null;
let _topbarRoot = null;
let _lastHash = null;
let _historyStack = [];
let _suppressHistory = false;

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

  if (_viewRoot) _viewRoot.innerHTML = '';
  const targetRoot = _viewRoot || document.getElementById('app');

  switch (hash) {
    case 'dup': return renderDup(targetRoot);
    case 'conn': return renderConn(targetRoot);
    case 'export': return renderExport(targetRoot);
    case 'paramprop': return renderParamProp(targetRoot);
    case 'sharedparambatch': return renderSharedParamBatch(targetRoot);
    case 'segmentpms': return renderSegmentPms(targetRoot);
    case 'guid': return renderGuid(targetRoot);
    case 'familylink': return renderFamilyLink(targetRoot);
    case 'multi': return renderMulti(targetRoot);
    case 'deliverycleaner': return renderDeliveryCleaner(targetRoot);
    case 'parammodifier': return renderParamModifier(targetRoot);
    case 'active-menu': return renderActiveMenu(targetRoot);
    default: return renderHome(targetRoot);
  }
}
