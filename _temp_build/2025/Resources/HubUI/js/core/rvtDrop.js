import { post, onHost } from './bridge.js';
import { log } from './dom.js';

export function normalizeDroppedRvtPath(value) {
  let path = String(value || '').trim();
  if (!path) return '';

  path = path.replace(/^"+|"+$/g, '');
  if (!path) return '';

  if (/^file:/i.test(path)) {
    try {
      const url = new URL(path);
      path = decodeURIComponent(url.pathname || '');
    } catch {
      path = path.replace(/^file:\/*/i, '');
    }
    if (/^\/[A-Za-z]:/.test(path)) path = path.slice(1);
    path = path.replace(/\//g, '\\');
  } else if (/^[A-Za-z]:\//.test(path)) {
    path = path.replace(/\//g, '\\');
  }

  if (!/\.rvt$/i.test(path)) return '';
  return path;
}

export function collectDroppedRvtPaths(dataTransfer) {
  if (!dataTransfer) return [];

  const picked = [];
  const seen = new Set();

  const pushPath = (value) => {
    const path = normalizeDroppedRvtPath(value);
    if (!path) return;
    const key = path.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    picked.push(path);
  };

  const files = Array.from(dataTransfer.files || []);
  files.forEach((file) => {
    pushPath(file?.path);
  });

  const items = Array.from(dataTransfer.items || []);
  items.forEach((item) => {
    try {
      const file = item?.getAsFile?.();
      pushPath(file?.path);
    } catch { }
  });

  const uriList = String(dataTransfer.getData?.('text/uri-list') || '').trim();
  if (uriList) {
    uriList
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter((line) => line && !line.startsWith('#'))
      .forEach(pushPath);
  }

  const text = String(dataTransfer.getData?.('text/plain') || '').trim();
  if (text) {
    text
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
      .forEach(pushPath);
  }

  return picked;
}

export function attachRvtDropZone(target, options = {}) {
  if (!target) return () => {};

  const activeClass = options.activeClass || 'is-drop-active';
  const onDropPaths = typeof options.onDropPaths === 'function' ? options.onDropPaths : () => {};
  const onInvalid = typeof options.onInvalid === 'function' ? options.onInvalid : null;
  let dragDepth = 0;

  const syncNativeOverlay = (active) => {
    try { post('ui:rvt-drop-overlay', { active: !!active }); } catch { }
  };

  const debugDropPayload = (event, resolvedPaths) => {
    try {
      const dt = event?.dataTransfer;
      const info = {
        types: Array.from(dt?.types || []),
        filesLength: Array.from(dt?.files || []).length,
        fileNames: Array.from(dt?.files || []).map((file) => ({
          name: file?.name || '',
          path: file?.path || '',
          type: file?.type || '',
          size: Number(file?.size || 0)
        })),
        items: Array.from(dt?.items || []).map((item) => ({
          kind: item?.kind || '',
          type: item?.type || ''
        })),
        textPlain: String(dt?.getData?.('text/plain') || ''),
        uriList: String(dt?.getData?.('text/uri-list') || ''),
        resolvedPaths: Array.isArray(resolvedPaths) ? resolvedPaths : []
      };
      log(`[rvtDrop] drop payload ${JSON.stringify(info)}`);
    } catch (err) {
      try { log(`[rvtDrop] drop payload logging failed: ${err?.message || err}`); } catch {}
    }
  };

  const hasFiles = (event) => {
    const types = Array.from(event?.dataTransfer?.types || []);
    return types.includes('Files');
  };

  const setActive = (active) => {
    target.classList.toggle(activeClass, !!active);
  };

  const isTargetVisible = () => {
    try {
      return !!target && target.isConnected && target.getClientRects().length > 0;
    } catch {
      return false;
    }
  };

  const handleDragEnter = (event) => {
    if (!hasFiles(event)) return;
    event.preventDefault();
    dragDepth += 1;
    setActive(true);
    syncNativeOverlay(true);
  };

  const handleDragOver = (event) => {
    if (!hasFiles(event)) return;
    event.preventDefault();
    if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy';
    setActive(true);
    syncNativeOverlay(true);
  };

  const handleDragLeave = (event) => {
    if (!hasFiles(event)) return;
    event.preventDefault();
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) {
      setActive(false);
      syncNativeOverlay(false);
    }
  };

  const handleDrop = (event) => {
    if (!hasFiles(event)) return;
    event.preventDefault();
    dragDepth = 0;
    setActive(false);
    syncNativeOverlay(false);
    const paths = collectDroppedRvtPaths(event.dataTransfer);
    debugDropPayload(event, paths);
    if (!paths.length) {
      try {
        post('ui:rvt-drop-commit', {
          filesLength: Array.from(event?.dataTransfer?.files || []).length,
          fileNames: Array.from(event?.dataTransfer?.files || []).map((file) => file?.name || '')
        });
      } catch {}
      return;
    }
    onDropPaths(paths);
  };

  target.addEventListener('dragenter', handleDragEnter);
  target.addEventListener('dragover', handleDragOver);
  target.addEventListener('dragleave', handleDragLeave);
  target.addEventListener('drop', handleDrop);
  const detachHostOverlay = onHost('host:rvt-drop-overlay', (payload) => {
    if (payload?.active) return;
    dragDepth = 0;
    setActive(false);
  });
  const detachInvalidDrop = onHost('host:rvt-drop-invalid', (payload) => {
    if (typeof onInvalid !== 'function') return;
    if (!isTargetVisible()) return;
    onInvalid(payload);
  });

  return () => {
    syncNativeOverlay(false);
    setActive(false);
    target.removeEventListener('dragenter', handleDragEnter);
    target.removeEventListener('dragover', handleDragOver);
    target.removeEventListener('dragleave', handleDragLeave);
    target.removeEventListener('drop', handleDrop);
    detachHostOverlay();
    detachInvalidDrop();
  };
}
