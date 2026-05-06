export const $ = (selector, root = document) => root.querySelector(selector);
export const $$ = (selector, root = document) => Array.from(root.querySelectorAll(selector));
export const clear = (el) => {
  while (el.firstChild) el.removeChild(el.firstChild);
};
export const div = (cls) => {
  const el = document.createElement('div');
  el.className = cls || '';
  return el;
};
export const tdText = (value) => {
  const td = document.createElement('td');
  td.textContent = value == null ? '' : String(value);
  return td;
};
export const debounce = (fn, delay = 200) => {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), delay);
  };
};

const EXCEL_EXPORT_LOCALE_KEY = 'kky.hub.excelExportLocale';

function normalizeExcelExportLocale(value) {
  const normalized = String(value || 'ko').trim().toLowerCase();
  return normalized === 'en' || normalized === 'eng' || normalized === 'english' ? 'en' : 'ko';
}

function readLastExcelExportLocale() {
  try {
    return normalizeExcelExportLocale(localStorage.getItem(EXCEL_EXPORT_LOCALE_KEY));
  } catch {
    return 'ko';
  }
}

let lastExcelExportLocale = readLastExcelExportLocale();

function setLastExcelExportLocale(value) {
  lastExcelExportLocale = normalizeExcelExportLocale(value);
  try {
    localStorage.setItem(EXCEL_EXPORT_LOCALE_KEY, lastExcelExportLocale);
  } catch {
  }
}

export function getLastExcelExportLocale() {
  return lastExcelExportLocale;
}

export function refreshUiAfterHostDialog(render, delay = 120) {
  if (typeof render !== 'function') return;

  const run = () => {
    try { render(); } catch { }
  };

  run();

  if (typeof window === 'undefined') return;

  let released = false;
  let raf1 = 0;
  let raf2 = 0;
  let timerRun = 0;
  let timerFinalize = 0;

  const cleanup = () => {
    if (released) return;
    released = true;
    window.removeEventListener('focus', onFocus, true);
    if (typeof document !== 'undefined') {
      document.removeEventListener('visibilitychange', onVisible, true);
    }
    if (raf1 && typeof window.cancelAnimationFrame === 'function') window.cancelAnimationFrame(raf1);
    if (raf2 && typeof window.cancelAnimationFrame === 'function') window.cancelAnimationFrame(raf2);
    if (timerRun) window.clearTimeout(timerRun);
    if (timerFinalize) window.clearTimeout(timerFinalize);
  };

  const rerender = () => {
    cleanup();
    run();
  };

  const onFocus = () => rerender();
  const onVisible = () => {
    if (document.visibilityState === 'visible') rerender();
  };

  window.addEventListener('focus', onFocus, true);
  if (typeof document !== 'undefined') {
    document.addEventListener('visibilitychange', onVisible, true);
  }

  if (typeof window.requestAnimationFrame === 'function') {
    raf1 = window.requestAnimationFrame(() => {
      run();
      raf2 = window.requestAnimationFrame(run);
    });
  } else {
    timerRun = window.setTimeout(run, 0);
  }

  timerFinalize = window.setTimeout(rerender, Math.max(0, Number(delay) || 0));
}

let logConsoleBody = null;
let logConsoleRoot = null;

export function initLogConsole() {
  logConsoleRoot = document.querySelector('[data-log-console]');
  logConsoleBody = document.querySelector('[data-log-body]');

  const closeBtn = logConsoleRoot?.querySelector('.log-console-close');
  if (closeBtn) closeBtn.addEventListener('click', hideLogConsole);

  hideLogConsole();
}

export function showLogConsole() {
  if (!logConsoleRoot) return;
  logConsoleRoot.classList.remove('is-hidden');
}

export function hideLogConsole() {
  if (!logConsoleRoot) return;
  logConsoleRoot.classList.add('is-hidden');
}

export function toggleLogConsole() {
  if (!logConsoleRoot) return;
  if (logConsoleRoot.classList.contains('is-hidden')) showLogConsole();
  else hideLogConsole();
}

export function log(message, payload) {
  try {
    if (payload !== undefined) console.log(message, payload);
    else console.log(message);
  } catch { }

  if (!logConsoleBody) return;

  const now = new Date();
  const hh = String(now.getHours()).padStart(2, '0');
  const mm = String(now.getMinutes()).padStart(2, '0');
  const ss = String(now.getSeconds()).padStart(2, '0');

  const line = document.createElement('div');
  line.className = 'log-line';

  let text = `[${hh}:${mm}:${ss}] ${message}`;
  if (payload !== undefined) {
    try {
      text += ` ${JSON.stringify(payload)}`;
    } catch { }
  }

  line.textContent = text;
  logConsoleBody.appendChild(line);
  logConsoleBody.scrollTop = logConsoleBody.scrollHeight;
}

let busyEl = null;
export function setBusy(on, text = '작업 중입니다.') {
  if (on) {
    if (busyEl) {
      const spinner = busyEl.querySelector('.spinner');
      if (spinner) spinner.textContent = text;
      return;
    }
    busyEl = document.createElement('div');
    busyEl.className = 'busy';
    const spinner = document.createElement('div');
    spinner.className = 'spinner';
    spinner.textContent = text;
    busyEl.append(spinner);
    document.body.append(busyEl);
    return;
  }

  if (busyEl) {
    busyEl.remove();
    busyEl = null;
  }
}

export function toast(msg, kind = 'info', ms = 2600) {
  const placement = kind === 'warn' ? 'top-center' : 'bottom-right';
  const selector = `.toast-wrap[data-placement="${placement}"]`;
  let wrap = $(selector);
  if (!wrap) {
    wrap = document.createElement('div');
    wrap.className = 'toast-wrap';
    wrap.dataset.placement = placement;
    document.body.append(wrap);
  }

  const toastEl = document.createElement('div');
  toastEl.className = `toast${
    kind === 'ok' ? ' ok' :
    kind === 'err' ? ' err' :
    kind === 'warn' ? ' warn' : ''
  }`;
  toastEl.textContent = msg;
  wrap.append(toastEl);

  const duration = kind === 'warn' && ms === 2600 ? 4200 : ms;
  setTimeout(() => {
    toastEl.remove();
    if (!wrap.children.length) wrap.remove();
  }, duration);
}

export function showExcelSavedDialog(message, filePath, onOpen, options = {}) {
  const {
    description = '저장한 파일을 바로 여시겠습니까?',
    openLabel = '파일 열기'
  } = options || {};

  toast(message || '엑셀 파일을 저장했습니다.', 'ok');

  const existing = document.querySelector('.excel-dialog-backdrop');
  if (existing) existing.remove();

  const backdrop = document.createElement('div');
  backdrop.className = 'excel-dialog-backdrop';

  const dialog = document.createElement('div');
  dialog.className = 'excel-dialog';

  const title = document.createElement('div');
  title.className = 'excel-dialog-title';
  title.textContent = message || '엑셀 파일을 저장했습니다.';

  const desc = document.createElement('div');
  desc.className = 'excel-dialog-desc';
  desc.textContent = description;

  const path = document.createElement('div');
  path.className = 'excel-dialog-path';
  path.textContent = filePath || '';

  const actions = document.createElement('div');
  actions.className = 'excel-dialog-actions';

  const btnOpen = document.createElement('button');
  btnOpen.type = 'button';
  btnOpen.className = 'btn btn-primary';
  btnOpen.textContent = openLabel;

  const btnClose = document.createElement('button');
  btnClose.type = 'button';
  btnClose.className = 'btn';
  btnClose.textContent = '닫기';

  const close = () => backdrop.remove();

  btnOpen.addEventListener('click', () => {
    close();
    if (typeof onOpen === 'function') onOpen(filePath);
  });
  btnClose.addEventListener('click', close);
  backdrop.addEventListener('click', (e) => {
    if (e.target === backdrop) close();
  });

  actions.append(btnOpen, btnClose);
  dialog.append(title, desc, path, actions);
  backdrop.append(dialog);
  document.body.append(backdrop);
  btnOpen.focus();
}

function showGlobalRuntimeError(detail) {
  try { console.error('[hub] runtime error', detail); } catch(e) {}
  toast('화면 처리 중 오류가 발생했습니다. 화면을 새로 열어 다시 실행해 주세요. 계속 실패하면 메시지를 관리자에게 전달해 주세요.', 'err', 4200);
}

window.addEventListener('error', (e) => showGlobalRuntimeError(e && (e.error || e.message || e)));
window.addEventListener('unhandledrejection', (e) => showGlobalRuntimeError(e && (e.reason || e)));

export function chooseExcelMode(onSelect, options = {}) {
  return new Promise((resolve) => {
    const allowSplit = !!(options && options.allowSplit);
    const existing = document.querySelector('.excelmode-backdrop');
    if (existing) existing.remove();

    const backdrop = document.createElement('div');
    backdrop.className = 'excelmode-backdrop';

    const dialog = document.createElement('div');
    dialog.className = 'excelmode-dialog';

    const title = document.createElement('div');
    title.className = 'excelmode-title';
    title.textContent = '엑셀 내보내기 옵션을 선택해 주세요.';

    const desc = document.createElement('div');
    desc.className = 'excelmode-desc';
    desc.textContent = '빠른 추출은 열 너비 맞춤을 건너뛰고, 일반 추출은 열 너비 자동 맞춤까지 수행합니다.';

    const localeSection = document.createElement('div');
    localeSection.className = 'excelmode-locale';

    const localeLabel = document.createElement('div');
    localeLabel.className = 'excelmode-locale-label';
    localeLabel.textContent = '엑셀 출력 양식';

    const localeHint = document.createElement('div');
    localeHint.className = 'excelmode-locale-hint';
    localeHint.textContent = '한글 선택 시 한글 헤더/문구가 적용되고, 미정의 항목은 원본(raw) 기준으로 유지됩니다. 영문 선택 시 영문 헤더와 문구가 적용됩니다.';

    const localeOptions = document.createElement('div');
    localeOptions.className = 'excelmode-locale-options';

    let selectedLocale = getLastExcelExportLocale();

    const buildLocaleOption = (value, titleText, descText) => {
      const label = document.createElement('label');
      label.className = 'excelmode-locale-option';

      const input = document.createElement('input');
      input.type = 'radio';
      input.name = 'excel-export-locale';
      input.value = value;
      input.checked = selectedLocale === value;
      input.addEventListener('change', () => {
        if (input.checked) selectedLocale = value;
      });

      const textWrap = document.createElement('span');
      textWrap.className = 'excelmode-locale-option__text';

      const title = document.createElement('strong');
      title.textContent = titleText;

      const description = document.createElement('span');
      description.textContent = descText;

      textWrap.append(title, description);
      label.append(input, textWrap);
      return label;
    };

    localeOptions.append(
      buildLocaleOption('ko', '한글', '한글 헤더와 검토 문구를 저장합니다.'),
      buildLocaleOption('en', '영문', '헤더와 검토 문구를 영문으로 저장합니다.')
    );
    localeSection.append(localeLabel, localeHint, localeOptions);

    let selectedScope = 'single';
    let scopeSection = null;
    if (allowSplit) {
      scopeSection = document.createElement('div');
      scopeSection.className = 'excelmode-scope';

      const scopeLabel = document.createElement('div');
      scopeLabel.className = 'excelmode-scope-label';
      scopeLabel.textContent = '저장 방식';

      const scopeHint = document.createElement('div');
      scopeHint.className = 'excelmode-scope-hint';
      scopeHint.textContent = '여러 RVT 검토 결과를 한 엑셀 파일로 묶거나, 폴더를 선택해 파일별로 따로 저장할 수 있습니다.';

      const scopeOptions = document.createElement('div');
      scopeOptions.className = 'excelmode-scope-options';

      const buildScopeOption = (value, titleText, descText) => {
        const label = document.createElement('label');
        label.className = 'excelmode-scope-option';

        const input = document.createElement('input');
        input.type = 'radio';
        input.name = 'excel-export-scope';
        input.value = value;
        input.checked = selectedScope === value;
        input.addEventListener('change', () => {
          if (input.checked) selectedScope = value;
        });

        const textWrap = document.createElement('span');
        textWrap.className = 'excelmode-scope-option__text';

        const title = document.createElement('strong');
        title.textContent = titleText;

        const description = document.createElement('span');
        description.textContent = descText;

        textWrap.append(title, description);
        label.append(input, textWrap);
        return label;
      };

      scopeOptions.append(
        buildScopeOption('single', '한 파일로 저장', '기존처럼 시트 또는 시트 묶음으로 저장합니다.'),
        buildScopeOption('split', '파일별로 저장', '폴더를 선택한 뒤 RVT별 개별 엑셀 파일로 저장합니다.')
      );
      scopeSection.append(scopeLabel, scopeHint, scopeOptions);
    }

    const actions = document.createElement('div');
    actions.className = 'excelmode-actions';

    const close = (mode) => {
      backdrop.remove();
      const picked = mode || null;
      if (!picked) {
        resolve(null);
        return;
      }

      setLastExcelExportLocale(selectedLocale);
      const result = allowSplit
        ? { mode: picked, splitByFile: selectedScope === 'split' }
        : picked;
      if (typeof onSelect === 'function') onSelect(result);
      resolve(result);
    };

    const btnFast = document.createElement('button');
    btnFast.type = 'button';
    btnFast.className = 'btn btn-primary';
    btnFast.textContent = '빠른 추출(기본)';
    btnFast.addEventListener('click', () => close('fast'));

    const btnNormal = document.createElement('button');
    btnNormal.type = 'button';
    btnNormal.className = 'btn';
    btnNormal.textContent = '일반 추출(열 너비 자동 맞춤)';
    btnNormal.addEventListener('click', () => close('normal'));

    const btnCancel = document.createElement('button');
    btnCancel.type = 'button';
    btnCancel.className = 'btn';
    btnCancel.textContent = '취소';
    btnCancel.addEventListener('click', () => close(null));

    actions.append(btnFast, btnNormal, btnCancel);
    dialog.append(title, desc, localeSection);
    if (scopeSection) dialog.append(scopeSection);
    dialog.append(actions);
    backdrop.append(dialog);
    document.body.append(backdrop);

    backdrop.addEventListener('click', (e) => {
      if (e.target === backdrop) close(null);
    });
  });
}

export function closeCompletionSummaryDialog() {
  const existing = document.querySelector('.completion-summary-backdrop');
  if (existing) existing.remove();
}

export function showCompletionSummaryDialog(options = {}) {
  closeCompletionSummaryDialog();

  const {
    title = '검토 완료',
    message = '검토가 완료되었습니다.',
    summaryItems = [],
    notes = [],
    actions = [],
    dialogClassName = '',
    exportLabel = '엑셀 내보내기',
    confirmLabel = '확인',
    showExport = true,
    exportDisabled = false,
    onExport,
    onClose
  } = options || {};

  const backdrop = document.createElement('div');
  backdrop.className = 'completion-summary-backdrop';

  const dialog = document.createElement('div');
  dialog.className = 'completion-summary-dialog';
  if (dialogClassName) {
    String(dialogClassName)
      .split(/\s+/)
      .map((name) => name.trim())
      .filter(Boolean)
      .forEach((name) => dialog.classList.add(name));
  }

  const header = document.createElement('div');
  header.className = 'completion-summary-header';
  const titleEl = document.createElement('h3');
  titleEl.textContent = title || '검토 완료';
  header.append(titleEl);

  const body = document.createElement('div');
  body.className = 'completion-summary-body';

  const messageEl = document.createElement('p');
  messageEl.className = 'completion-summary-message';
  messageEl.textContent = message || '검토가 완료되었습니다.';
  body.append(messageEl);

  const normalizedItems = Array.isArray(summaryItems) ? summaryItems.filter(Boolean) : [];
  if (normalizedItems.length) {
    const grid = document.createElement('div');
    grid.className = 'completion-summary-grid';

    normalizedItems.forEach((item) => {
      const row = document.createElement('div');
      row.className = 'completion-summary-item';

      const label = document.createElement('div');
      label.className = 'completion-summary-item-label';
      label.textContent = item.label == null ? '' : String(item.label);

      const value = document.createElement('div');
      value.className = 'completion-summary-item-value';
      value.textContent = item.value == null ? '' : String(item.value);

      row.append(label, value);
      grid.append(row);
    });

    body.append(grid);
  }

  const normalizedNotes = Array.isArray(notes)
    ? notes.map((note) => (note == null ? '' : String(note).trim())).filter(Boolean)
    : [];
  if (normalizedNotes.length) {
    const noteWrap = document.createElement('div');
    noteWrap.className = 'completion-summary-notes';
    normalizedNotes.forEach((note) => {
      const noteEl = document.createElement('div');
      noteEl.className = 'completion-summary-note';
      noteEl.textContent = note;
      noteWrap.append(noteEl);
    });
    body.append(noteWrap);
  }

  const footer = document.createElement('div');
  footer.className = 'completion-summary-footer';

  const confirmBtn = document.createElement('button');
  confirmBtn.type = 'button';
  confirmBtn.className = 'btn btn--primary';
  confirmBtn.textContent = confirmLabel || '확인';

  const close = () => {
    backdrop.remove();
    if (typeof onClose === 'function') onClose();
  };

  const normalizedActions = Array.isArray(actions) ? actions.filter(Boolean) : [];
  normalizedActions.forEach((action) => {
    const actionBtn = document.createElement('button');
    actionBtn.type = 'button';
    actionBtn.className = `btn ${action.variant === 'primary' ? 'btn--primary' : (action.variant === 'danger' ? 'btn--danger' : 'btn--secondary')}`;
    actionBtn.textContent = action.label == null ? '' : String(action.label);
    actionBtn.disabled = !!action.disabled;
    actionBtn.addEventListener('click', () => {
      if (actionBtn.disabled) return;
      if (action.closeOnClick !== false) close();
      if (typeof action.onClick === 'function') action.onClick();
    });
    footer.append(actionBtn);
  });

  if (showExport) {
    const exportBtn = document.createElement('button');
    exportBtn.type = 'button';
    exportBtn.className = 'btn';
    exportBtn.textContent = exportLabel || '엑셀 내보내기';
    exportBtn.disabled = !!exportDisabled;
    exportBtn.addEventListener('click', () => {
      if (exportBtn.disabled) return;
      close();
      if (typeof onExport === 'function') onExport();
    });
    footer.append(exportBtn);
    if (!exportBtn.disabled) {
      requestAnimationFrame(() => exportBtn.focus());
    }
  }

  confirmBtn.addEventListener('click', close);
  footer.append(confirmBtn);

  backdrop.addEventListener('click', (e) => {
    if (e.target === backdrop) close();
  });
  dialog.addEventListener('click', (e) => e.stopPropagation());

  dialog.append(header, body, footer);
  backdrop.append(dialog);
  document.body.append(backdrop);

  if (!showExport || exportDisabled) {
    confirmBtn.focus();
  }

  return { close };
}
