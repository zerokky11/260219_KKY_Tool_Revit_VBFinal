// Resources/HubUI/js/views/dup.js
import { clear, div, toast, showExcelSavedDialog, chooseExcelMode } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { onHost, post } from '../core/bridge.js';

const RESP_ROWS_EVENTS = ['dup:list', 'dup:rows', 'duplicate:list'];
const EV_DELETE_REQ = 'duplicate:delete';
const EV_RESTORE_REQ = 'duplicate:restore';
const EV_SELECT_REQ  = 'duplicate:select';
const EV_EXPORT_REQ  = 'duplicate:export';

const EV_DELETED_ONE   = 'dup:deleted';
const EV_RESTORED_ONE  = 'dup:restored';
const EV_DELETED_MULTI = 'duplicate:delete';
const EV_RESTORED_MULTI= 'duplicate:restore';
const EV_EXPORTED_A    = 'duplicate:export';
const EV_EXPORTED_B    = 'dup:exported';

const DUP_MODE_KEY = 'kky_dup_mode';         // "duplicate" | "clash"
const DUP_TOL_MM_KEY = 'kky_dup_tol_mm';
const DUP_TOL_MM_DEFAULT = 4.7625;           // 1/64 ft ≈ 4.7625mm

export function renderDup(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);

  if (!document.getElementById('dup-style-override')) {
    const st = document.createElement('style');
    st.id = 'dup-style-override';
    st.textContent = `
      .dup-row.is-deleted .cell { text-decoration: none !important; opacity: .55; }
      .dup-row .row-actions .table-action-btn.restore {
        background: color-mix(in oklab, var(--accent, #4c6fff) 85%, #ffffff 15%);
        color:#fff;
      }

      .dup-modebar { display:flex; align-items:center; gap:8px; }
      .dup-modebar .chip-btn.is-active{
        background: color-mix(in oklab, var(--accent, #4c6fff) 18%, transparent 82%);
        border-color: color-mix(in oklab, var(--accent, #4c6fff) 55%, transparent 45%);
        font-weight: 600;
      }

      .dup-tol {
        display:flex;
        align-items:center;
        gap:8px;
        padding: 6px 10px;
        border-radius: 999px;
        border: 1px solid color-mix(in oklab, var(--border, #d7dbe7) 70%, transparent 30%);
        background: color-mix(in oklab, var(--panel, #ffffff) 92%, #000000 8%);
      }
      .dup-tol .dup-tol-label { font-size: 12px; opacity: .85; white-space: nowrap; }
      .dup-tol .dup-tol-input {
        width: 92px;
        padding: 4px 8px;
        border-radius: 10px;
        border: 1px solid color-mix(in oklab, var(--border, #d7dbe7) 70%, transparent 30%);
        background: transparent;
        color: inherit;
        outline: none;
      }
      .dup-tol .dup-tol-input:disabled { opacity: .6; }

      .dup-info {
        margin-bottom: 10px;
        padding: 12px 14px;
        border-radius: 14px;
        border: 1px solid color-mix(in oklab, var(--border, #d7dbe7) 70%, transparent 30%);
        background: color-mix(in oklab, var(--panel, #ffffff) 92%, #000000 8%);
      }
      .dup-info .t { font-weight: 600; margin-bottom: 6px; }
      .dup-info .s { opacity: .85; font-size: 12px; line-height: 1.35; }
    `;
    document.head.appendChild(st);
  }

  const topbarEl = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar');
  if (topbarEl) topbarEl.classList.add('hub-topbar');

  const page    = div('dup-page feature-shell');
  const header  = div('feature-header dup-toolbar');
  const heading = div('feature-heading');

  let mode = readMode();
  let tolInputEl = null;
  let modeBtns = [];
  let activeModeForView = mode;

  const runBtn    = cardBtn('검토 시작', onRun);
  const exportBtn = cardBtn('엑셀 내보내기', onExport);
  exportBtn.disabled = true;

  const modeBar = buildModeBar();
  const tolCtl  = buildTolControl();

  const actions = div('feature-actions');
  actions.append(runBtn, modeBar, tolCtl, exportBtn);

  header.append(heading, actions);
  page.append(header);

  const summaryBar = div('dup-summarybar sticky hidden');
  page.append(summaryBar);

  const body = div('dup-body');
  page.append(body);

  target.append(page);

  const EXCEL_PHASE_WEIGHT = { EXCEL_INIT: 0.05, EXCEL_WRITE: 0.85, EXCEL_SAVE: 0.08, AUTOFIT: 0.02, DONE: 1, ERROR: 1 };
  const EXCEL_PHASE_ORDER = ['EXCEL_INIT', 'EXCEL_WRITE', 'EXCEL_SAVE', 'AUTOFIT', 'DONE'];

  let rows = [];
  let groups = [];
  let deleted = new Set();
  let expanded = new Set();
  let waitTimer = null;
  let busy = false;
  let exporting = false;
  let lastExcelPct = 0;

  // 결과 요약(Host)
  let lastResult = null;
  let lastTruncToastKey = '';

  applyHeadingByMode(mode);
  renderIntro(body, mode);

  onHost('revit:error', ({ message }) => {
    setLoading(false);
    ProgressDialog.hide();
    lastExcelPct = 0;
    exporting = false;
    exportBtn.disabled = rows.length === 0;
    toast(message || 'Revit 오류가 발생했습니다.', 'err', 3200);
  });

  onHost('host:error', ({ message }) => {
    setLoading(false);
    ProgressDialog.hide();
    lastExcelPct = 0;
    exporting = false;
    toast(message || '호스트 오류가 발생했습니다.', 'err', 3200);
  });

  onHost('host:warn', ({ message }) => {
    setLoading(false);
    if (message) toast(message, 'warn', 3600);
  });

  onHost(({ ev, payload }) => {
    if (RESP_ROWS_EVENTS.includes(ev)) {
      setLoading(false);
      const list = payload?.rows ?? payload?.data ?? payload ?? [];
      handleRows(list);
      return;
    }

    if (ev === 'dup:result') {
      setLoading(false);
      lastResult = payload || null;

      const m = String(payload?.mode ?? '').trim();
      if (m === 'duplicate' || m === 'clash') activeModeForView = m;

      if (payload?.truncated) {
        const k = `${payload?.mode}|${payload?.shown}|${payload?.total}`;
        if (k !== lastTruncToastKey) {
          lastTruncToastKey = k;
          toast(`결과가 많아 상위 ${payload?.shown ?? ''}건만 표시합니다. 전체는 엑셀 내보내기에서 확인하세요.`, 'warn', 4200);
        }
      }

      const groupsN = Number(payload?.groups ?? 0) || 0;
      const candN   = Number(payload?.candidates ?? 0) || 0;

      // dup:list가 안 오거나(전송 실패/과다) 0건일 때도 상태를 확실히 표시
      if ((groupsN === 0 && candN === 0) && rows.length === 0 && !busy) {
        handleRows([]);
      } else {
        paintGroups();
        refreshSummary();
      }
      return;
    }

    if (ev === EV_DELETED_ONE) {
      const id = String(payload?.id ?? '');
      if (id) { deleted.add(id); updateRowStates(); refreshSummary(); }
      return;
    }

    if (ev === EV_RESTORED_ONE) {
      const id = String(payload?.id ?? '');
      if (id) { deleted.delete(id); updateRowStates(); refreshSummary(); }
      return;
    }

    if (ev === EV_DELETED_MULTI) {
      (toIdArray(payload?.ids)).forEach(id => deleted.add(id));
      updateRowStates(); refreshSummary();
      return;
    }

    if (ev === EV_RESTORED_MULTI) {
      (toIdArray(payload?.ids)).forEach(id => deleted.delete(id));
      updateRowStates(); refreshSummary();
      return;
    }

    if (ev === 'dup:progress') {
      handleExcelProgress(payload || {});
      return;
    }

    if (ev === EV_EXPORTED_A || ev === EV_EXPORTED_B) {
      lastExcelPct = 0;
      ProgressDialog.hide();

      const path = payload?.path || '';
      if (payload?.ok || path) {
        showExcelSavedDialog('엑셀로 내보냈습니다.', path, (p) => {
          if (p) post('excel:open', { path: p });
        });
      } else {
        toast(payload?.message || '엑셀 내보내기 실패', 'err');
      }

      exporting = false;
      exportBtn.disabled = rows.length === 0;
      return;
    }
  });

  function setLoading(on) {
    busy = on;
    runBtn.disabled = on;
    runBtn.textContent = on ? '검토 중…' : '검토 시작';
    modeBtns.forEach(b => { try { b.disabled = on; } catch {} });
    if (tolInputEl) tolInputEl.disabled = on;

    if (!on && waitTimer) { clearTimeout(waitTimer); waitTimer = null; }
  }

  function onRun() {
    setLoading(true);
    exportBtn.disabled = true;
    deleted.clear();
    rows = [];
    groups = [];
    lastResult = null;

    body.innerHTML = '';
    body.append(buildSkeleton(6));

    waitTimer = setTimeout(() => {
      setLoading(false);
      toast('응답이 없습니다.\nAdd-in 이벤트명을 확인하세요 (예: dup:list).', 'err');
      body.innerHTML = '';
      renderIntro(body, mode);
    }, 10000);

    const tolFeet = getTolFeet();
    activeModeForView = mode;
    post('dup:run', { mode, tolFeet });
  }

  function onExport() {
    if (exporting) return;
    exporting = true;
    exportBtn.disabled = true;
    chooseExcelMode((excelMode) => post(EV_EXPORT_REQ, { excelMode: excelMode || 'fast' }));
  }

  function handleExcelProgress(payload) {
    if (!payload) {
      ProgressDialog.hide();
      exporting = false;
      exportBtn.disabled = rows.length === 0;
      lastExcelPct = 0;
      return;
    }

    const phase = normalizeExcelPhase(payload.phase);
    const total = Number(payload.total) || 0;
    const current = Number(payload.current) || 0;

    const percent = computeExcelPercent(phase, current, total, payload.phaseProgress);
    const subtitle = buildExcelSubtitle(phase, current, total);
    const detail = payload.message || '';

    exporting = phase !== 'DONE' && phase !== 'ERROR';
    exportBtn.disabled = exporting || rows.length === 0;

    ProgressDialog.show(`${modeTitle(activeModeForView)} 엑셀 내보내기`, subtitle);
    ProgressDialog.update(percent, subtitle, detail);

    if (!exporting) {
      setTimeout(() => { ProgressDialog.hide(); lastExcelPct = 0; }, 280);
    }
  }

  function normalizeExcelPhase(phase) {
    return String(phase || '').trim().toUpperCase() || 'EXCEL_WRITE';
  }

  function computeExcelPercent(phase, current, total, phaseProgress) {
    const norm = normalizeExcelPhase(phase);
    if (norm === 'DONE') { lastExcelPct = 100; return 100; }
    if (norm === 'ERROR') return lastExcelPct;

    const completed = EXCEL_PHASE_ORDER.reduce((acc, key) => {
      if (key === norm) return acc;
      return acc + (EXCEL_PHASE_WEIGHT[key] || 0);
    }, 0);

    const weight = EXCEL_PHASE_WEIGHT[norm] || 0;
    const ratio = total > 0 ? Math.min(1, Math.max(0, current / total)) : 0;
    const staged = Math.max(ratio, clamp01(phaseProgress));

    const pct = Math.min(100, Math.max(lastExcelPct, (completed + weight * staged) * 100));
    lastExcelPct = pct;
    return pct;
  }

  function buildExcelSubtitle(phase, current, total) {
    const labelMap = {
      EXCEL_INIT: '엑셀 준비',
      EXCEL_WRITE: '엑셀 작성',
      EXCEL_SAVE: '파일 저장',
      AUTOFIT: 'AutoFit',
      DONE: '완료',
      ERROR: '오류'
    };
    const label = labelMap[normalizeExcelPhase(phase)] || '엑셀 진행';
    const count = total > 0 ? ` (${Math.max(current, 0)}/${total})` : '';
    return `${label}${count}`;
  }

  function clamp01(v) {
    const n = Number(v);
    if (!Number.isFinite(n)) return 0;
    return Math.max(0, Math.min(1, n));
  }

  function handleRows(listLike) {
    const list = Array.isArray(listLike) ? listLike : [];
    rows = list.map(normalizeRow);

    const rowMode = rows.find(r => r.mode)?.mode;
    if (rowMode === 'duplicate' || rowMode === 'clash') activeModeForView = rowMode;

    groups = buildGroups(rows);
    exportBtn.disabled = rows.length === 0;
    setLoading(false);

    expanded = new Set(groups.map(g => g.key));
    paintGroups();

    if (!rows.length) {
      body.innerHTML = '';

      const isClash = activeModeForView === 'clash';
      const title = isClash ? '간섭이 없습니다' : '중복이 없습니다';

      const scan = Number(lastResult?.scan ?? 0) || 0;
      const gN = Number(lastResult?.groups ?? 0) || 0;
      const cN = Number(lastResult?.candidates ?? 0) || 0;
      const shown = Number(lastResult?.shown ?? 0) || 0;
      const total = Number(lastResult?.total ?? 0) || 0;

      const empty = div('dup-emptycard');
      empty.innerHTML = `
        <div class="empty-emoji">✅</div>
        <h3 class="empty-title">${title}</h3>
        <p class="empty-sub">${(gN === 0 && cN === 0) ? '검토 결과가 0건입니다.' : '표시할 결과가 없습니다.'}</p>
      `;
      body.append(empty);

      const info = div('dup-info');
      info.innerHTML = `
        <div class="t">검토 상태</div>
        <div class="s">스캔: ${scan.toLocaleString()}개 · 그룹: ${gN.toLocaleString()}개 · 결과행: ${cN.toLocaleString()}개</div>
        ${lastResult?.truncated ? `<div class="s">표시 제한: ${shown.toLocaleString()} / ${total.toLocaleString()} (전체는 엑셀 내보내기)</div>` : ``}
      `;
      body.append(info);
    }

    refreshSummary();
  }

  function paintGroups() {
    body.innerHTML = '';

    if (lastResult?.truncated) {
      const shown = Number(lastResult?.shown ?? 0) || 0;
      const total = Number(lastResult?.total ?? 0) || 0;
      const info = div('dup-info');
      info.innerHTML = `
        <div class="t">표시 제한</div>
        <div class="s">결과가 많아 상위 ${shown.toLocaleString()}건만 표시합니다. 전체(${total.toLocaleString()}건)는 엑셀 내보내기에서 확인하세요.</div>
      `;
      body.append(info);
    }

    const isClash = activeModeForView === 'clash';
    const grpPrefix = isClash ? '간섭 그룹' : '중복 그룹';

    groups.forEach((g, idx) => {
      const card = div('dup-grp');
      card.classList.add(g.rows.length >= 2 ? 'accent-danger' : 'accent-info');

      const h = div('grp-h');
      const left = div('grp-txt');

      const meta = buildGroupMeta(g);
      left.innerHTML = `
        <div class="grp-title">
          <span class="grp-badge">${grpPrefix} ${idx + 1}</span>
          <span class="grp-meta">${esc(meta)}</span>
        </div>
        <div class="grp-count">${g.rows.length}개</div>
      `;

      const right = div('grp-actions');
      const toggle = kbtn(expanded.has(g.key) ? '접기' : '펼치기', 'subtle', () => toggleGroup(g.key));
      right.append(toggle);

      h.append(left, right);
      card.append(h);

      const tbl = div('grp-body');

      const sh = div('dup-subhead');
      sh.append(
        cell('', 'ck'),
        cell('Element ID', 'th'),
        cell('Category', 'th'),
        cell('Family', 'th'),
        cell('Type', 'th'),
        cell('작업', 'th right')
      );
      tbl.append(sh);

      if (expanded.has(g.key)) g.rows.forEach(r => tbl.append(renderRow(r)));

      card.append(tbl);
      body.append(card);
    });

    updateRowStates();
  }

  function buildGroupMeta(g) {
    const cats = uniq(g.rows.map(r => r.category || '—'));
    const fams = uniq(g.rows.map(r => (r.family || (r.category ? `${r.category} Type` : '—')) || '—'));
    const types = uniq(g.rows.map(r => r.type || '—'));

    const catOut = cats.length === 1 ? cats[0] : `혼합(${cats.length})`;
    const famOut = fams.length === 1 ? fams[0] : `혼합(${fams.length})`;
    const typOut = types.length === 1 ? types[0] : `혼합(${types.length})`;

    return `${catOut} · ${famOut} · ${typOut}`;
  }

  function renderRow(r) {
    const row = div('dup-row');
    row.dataset.id = r.id;

    const ckCell = cell(null, 'ck');
    const ck = document.createElement('input');
    ck.type = 'checkbox';
    ck.className = 'ckbox';
    ck.onchange = () => row.classList.toggle('is-selected', ck.checked);
    ckCell.append(ck);

    row.append(ckCell);
    row.append(cell(r.id ?? '-', 'td mono right'));
    row.append(cell(r.category || '—', 'td'));

    const famOut = r.family ? r.family : (r.category ? `${r.category} Type` : '—');
    row.append(cell(famOut, 'td ell'));
    row.append(cell(r.type || '—', 'td ell'));

    const act = div('row-actions');

    const viewBtn = tableBtn('선택/줌', '', () => post(EV_SELECT_REQ, { id: r.id, zoom: true, mode: 'zoom' }));

    const delBtn = tableBtn(
      r.deleted ? '되돌리기' : '삭제',
      r.deleted ? 'restore' : 'table-action-btn--danger',
      () => {
        const ids = [r.id];
        if (delBtn.dataset.mode === 'restore') post(EV_RESTORE_REQ, { id: r.id, ids });
        else post(EV_DELETE_REQ, { id: r.id, ids });
      }
    );
    delBtn.dataset.mode = r.deleted ? 'restore' : 'delete';

    act.append(viewBtn, delBtn);
    row.append(cell(act, 'td right'));

    return row;
  }

  function toggleGroup(key) {
    if (expanded.has(key)) expanded.delete(key);
    else expanded.add(key);
    paintGroups();
  }

  function updateRowStates() {
    [...document.querySelectorAll('.dup-row')].forEach(rowEl => {
      const id = rowEl.dataset.id;
      if (!id) return;

      const isDel = deleted.has(String(id));
      rowEl.classList.toggle('is-deleted', isDel);

      const delBtn = rowEl.querySelector('.row-actions .table-action-btn:last-child');
      if (delBtn) {
        delBtn.textContent = isDel ? '되돌리기' : '삭제';
        delBtn.className = 'table-action-btn ' + (isDel ? 'restore' : 'table-action-btn--danger');
        delBtn.dataset.mode = isDel ? 'restore' : 'delete';
      }

      const ck = rowEl.querySelector('input.ckbox');
      if (ck) rowEl.classList.toggle('is-selected', !!ck.checked);
    });
  }

  function refreshSummary() {
    const totals = computeSummary(groups);
    summaryBar.innerHTML = '';
    summaryBar.classList.toggle('hidden', totals.totalCount === 0 && !busy);

    const isClash = activeModeForView === 'clash';
    const gLabel = isClash ? '간섭 그룹' : '중복 그룹';
    [chip(`${gLabel} ${totals.groupCount}`), chip(`요소 ${totals.totalCount}`)]
      .forEach(c => summaryBar.append(c));
  }

  function buildModeBar() {
    const wrap = div('dup-modebar');
    const bDup = kbtn('중복', 'subtle', () => setMode('duplicate'));
    const bClash = kbtn('자체간섭', 'subtle', () => setMode('clash'));
    modeBtns = [bDup, bClash];
    syncModeButtons();
    wrap.append(bDup, bClash);
    return wrap;
  }

  function setMode(next) {
    if (next !== 'duplicate' && next !== 'clash') return;
    mode = next;
    try { localStorage.setItem(DUP_MODE_KEY, mode); } catch {}
    syncModeButtons();

    rows = [];
    groups = [];
    deleted.clear();
    exportBtn.disabled = true;
    lastResult = null;

    body.innerHTML = '';
    applyHeadingByMode(mode);
    renderIntro(body, mode);
    refreshSummary();
  }

  function syncModeButtons() {
    if (!modeBtns || modeBtns.length < 2) return;
    const [bDup, bClash] = modeBtns;
    bDup.classList.toggle('is-active', mode === 'duplicate');
    bClash.classList.toggle('is-active', mode === 'clash');
  }

  function buildTolControl() {
    const wrap = div('dup-tol');
    wrap.title = '허용오차(mm). 중복=좌표/끝점 양자화, 자체간섭=여유/정밀판정 허용치로 사용됩니다.';

    const label = document.createElement('span');
    label.className = 'dup-tol-label';
    label.textContent = '허용오차(mm)';

    const input = document.createElement('input');
    input.className = 'dup-tol-input';
    input.type = 'number';
    input.min = '0.01';
    input.step = '0.1';

    const initMm = readTolMm();
    input.value = fmtTolMm(initMm);

    input.addEventListener('change', () => {
      const mm = sanitizeTolMm(Number(input.value));
      input.value = fmtTolMm(mm);
      try { localStorage.setItem(DUP_TOL_MM_KEY, String(mm)); } catch {}
    });

    wrap.append(label, input);
    tolInputEl = input;
    return wrap;
  }

  function applyHeadingByMode(m) {
    const title = modeTitle(m);
    const sub = (m === 'clash')
      ? '같은 파일 내 자체간섭 후보를 그룹으로 묶어 보여줍니다. (결과 과다 시 표시 제한될 수 있음)'
      : '중복 요소 후보를 그룹별로 확인하고 삭제/되돌리기를 관리합니다. (결과 과다 시 표시 제한될 수 있음)';

    heading.innerHTML = `
      <span class="feature-kicker">Duplicate Inspector</span>
      <h2 class="feature-title">${title}</h2>
      <p class="feature-sub">${sub}</p>
    `;
  }

  function modeTitle(m) { return m === 'clash' ? '자체간섭 검토' : '중복검토'; }

  function renderIntro(container, m) {
    const hero = div('dup-hero');
    const isClash = m === 'clash';
    hero.innerHTML = `
      <h3 class="hero-title">${isClash ? '자체간섭 검토를 시작해 보세요' : '중복검토를 시작해 보세요'}</h3>
      <p class="hero-sub">${isClash ? '같은 파일 내에서 자체간섭 후보를 그룹으로 묶어 보여줍니다.' : '모델의 중복 요소를 그룹으로 묶어 보여줍니다.'}</p>
      <ul class="hero-list">
        <li>상단 토글로 <b>중복</b>/<b>자체간섭</b> 모드를 전환할 수 있습니다.</li>
        <li>결과가 0건이면 <b>0건 안내</b>가 표시됩니다.</li>
        <li>결과가 너무 많으면 <b>표시 제한</b> 안내가 표시됩니다(전체는 엑셀).</li>
      </ul>`;
    container.append(hero);
  }

  function normalizeRow(r) {
    const id = safeId(r.elementId ?? r.ElementId ?? r.id ?? r.Id);
    const category = val(r.category ?? r.Category);
    const family   = val(r.family ?? r.Family);
    const type     = val(r.type ?? r.Type);
    const deletedFlag = !!(r.deleted ?? r.isDeleted ?? r.Deleted);
    const groupKey = val(r.groupKey ?? r.GroupKey);
    const rm = val(r.mode ?? r.Mode);
    const connectedIdsRaw = r.connectedIds ?? r.ConnectedIds ?? [];
    const connectedIds = Array.isArray(connectedIdsRaw)
      ? connectedIdsRaw.map(String)
      : (typeof connectedIdsRaw === 'string' && connectedIdsRaw.length
          ? connectedIdsRaw.split(/[,\s]+/).filter(Boolean)
          : []);
    return { id: id || '-', category, family, type, deleted: deletedFlag, groupKey, mode: rm, connectedIds };
  }

  function buildGroups(rs) {
    const hasGroupKey = rs.some(x => !!x.groupKey);
    if (hasGroupKey) {
      const map = new Map();
      for (const r of rs) {
        const key = r.groupKey || '_';
        let g = map.get(key);
        if (!g) { g = { key, rows: [] }; map.set(key, g); }
        g.rows.push(r);
      }
      return [...map.values()];
    }

    const map = new Map();
    for (const r of rs) {
      const cluster = [String(r.id), ...r.connectedIds.map(String)]
        .filter(Boolean)
        .map(x => x.trim())
        .sort((a, b) => Number(a) - Number(b))
        .join(',');
      const key = [r.category || '', r.family || '', r.type || '', cluster].join('|');
      let g = map.get(key);
      if (!g) { g = { key, rows: [] }; map.set(key, g); }
      g.rows.push(r);
    }
    return [...map.values()];
  }

  function computeSummary(groups) {
    let total = 0;
    groups.forEach(g => { total += g.rows.length; });
    return { groupCount: groups.length, totalCount: total };
  }

  function cardBtn(label, handler) {
    const b = document.createElement('button');
    b.className = 'card-action-btn';
    b.type = 'button';
    b.textContent = label;
    b.onclick = handler;
    return b;
  }

  function tableBtn(label, tone, handler) {
    const b = document.createElement('button');
    b.className = 'table-action-btn ' + (tone || '');
    b.type = 'button';
    b.textContent = label;
    b.onclick = handler;
    return b;
  }

  function kbtn(label, tone, handler) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'control-chip chip-btn ' + (tone || '');
    b.textContent = label;
    b.onclick = handler;
    return b;
  }

  function cell(content, cls) {
    const c = document.createElement('div');
    c.className = 'cell ' + (cls || '');
    if (content instanceof Node) c.append(content);
    else if (content != null) c.textContent = content;
    return c;
  }

  function chip(text, tone) {
    const b = div('chip ' + (tone || ''));
    b.textContent = text;
    return b;
  }

  function buildSkeleton(n = 6) {
    const wrap = div('dup-skeleton');
    for (let i = 0; i < n; i++) {
      const line = div('sk-row');
      line.append(div('sk-chip'), div('sk-id'), div('sk-wide'), div('sk-wide'), div('sk-act'));
      wrap.append(line);
    }
    return wrap;
  }

  function toIdArray(v) {
    if (!v) return [];
    if (Array.isArray(v)) return v.map(String);
    return [String(v)];
  }

  function uniq(arr) {
    const set = new Set();
    arr.forEach(x => set.add(String(x)));
    return [...set.values()];
  }

  function esc(s) {
    return String(s ?? '').replace(/[&<>"']/g, m => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[m]));
  }

  function safeId(v) { if (v === 0) return 0; if (v == null) return ''; return String(v); }
  function val(v) { return v == null || v === '' ? '' : String(v); }

  function readMode() {
    try {
      const m = String(localStorage.getItem(DUP_MODE_KEY) || '').trim();
      if (m === 'duplicate' || m === 'clash') return m;
    } catch {}
    return 'duplicate';
  }

  function readTolMm() {
    try {
      const raw = localStorage.getItem(DUP_TOL_MM_KEY);
      const n = Number(String(raw || '').trim());
      if (Number.isFinite(n) && n > 0) return sanitizeTolMm(n);
    } catch {}
    return DUP_TOL_MM_DEFAULT;
  }

  function sanitizeTolMm(n) {
    if (!Number.isFinite(n)) return DUP_TOL_MM_DEFAULT;
    return Math.max(0.01, Math.min(1000, n));
  }

  function fmtTolMm(mm) {
    const n = Number(mm);
    if (!Number.isFinite(n)) return String(DUP_TOL_MM_DEFAULT);
    return (Math.round(n * 1000) / 1000).toString();
  }

  function getTolFeet() {
    const mm = (tolInputEl ? sanitizeTolMm(Number(tolInputEl.value)) : readTolMm());
    const feet = mm / 304.8;
    return Math.max(0.000001, Number.isFinite(feet) ? feet : (DUP_TOL_MM_DEFAULT / 304.8));
  }
}
