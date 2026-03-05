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

const DUP_META_KEY = 'kky_dup_meta_v1';
const DUP_RULECFG_KEY = 'kky_dup_ruleset_v1';


const DUP_SCOPE_KEY = 'kky_dup_scope_mode';     // 'all' | 'scope' | 'exclude'
const DUP_EXCL_KW_KEY = 'kky_dup_excl_kw';      // comma separated keywords


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

      /* Rule/Set modal (floating window) */
      .dup-rulemodal{ position: fixed; inset: 0; z-index: 5000; display:none; }
      .dup-rulemodal.is-open{ display:block; }
      .dup-rulemodal .rm-backdrop{ position:absolute; inset:0; background: rgba(0,0,0,.22); }
      .dup-rulemodal .rm-window{
  position:absolute;
  inset: 10px;
  width: auto;
  height: auto;
  border-radius: 18px;
  border: 1px solid color-mix(in oklab, var(--border, #d7dbe7) 70%, transparent 30%);
  background: color-mix(in oklab, var(--panel, #ffffff) 94%, #000000 6%);
  box-shadow: 0 20px 60px rgba(0,0,0,.28);
  overflow:hidden;
  display:flex;
  flex-direction: column;
}
@media (max-width: 980px){
  .dup-rulemodal .rm-window{ inset: 8px; border-radius: 14px; }
}
      .dup-rulemodal .rm-head{ display:flex; align-items:center; justify-content:space-between; gap:12px; padding: 14px 16px; border-bottom: 1px solid color-mix(in oklab, var(--border, #d7dbe7) 70%, transparent 30%);
        background: linear-gradient(180deg, color-mix(in oklab, var(--panel, #ffffff) 92%, #000000 8%), color-mix(in oklab, var(--panel, #ffffff) 96%, #000000 4%)); }
      .dup-rulemodal .rm-title{ font-weight: 700; }
      .dup-rulemodal .rm-sub{ font-size:12px; opacity:.75; margin-top:2px; }
      .dup-rulemodal .rm-head-left{ display:flex; flex-direction:column; }
      .dup-rulemodal .rm-actions{ display:flex; gap:8px; flex-wrap:wrap; align-items:center; }
      .dup-rulemodal .rm-body{ padding: 14px 16px; overflow:auto; ; padding-bottom: 70px }
      .dup-rulemodal .rp-sec{ margin: 10px 0 14px 0; padding: 12px 12px; border-radius: 14px;
        border: 1px solid color-mix(in oklab, var(--border, #d7dbe7) 70%, transparent 30%);
        background: color-mix(in oklab, var(--panel, #ffffff) 92%, #000000 8%);
      }
      .dup-rulemodal .rp-sec-title{ font-weight: 700; margin-bottom: 8px; display:flex; align-items:center; gap:8px; }
      .dup-rulemodal .rp-sec-title .rp-pill{ font-size:11px; padding: 2px 8px; border-radius:999px;
        border:1px solid color-mix(in oklab, var(--border, #d7dbe7) 70%, transparent 30%);
        opacity:.8; }
      .dup-rulemodal .rp-grid{ display:grid; grid-template-columns: 160px 1fr; gap: 10px 12px; align-items:center; }
      .dup-rulemodal .rp-row{ display: contents; }
      .dup-rulemodal .rp-label{ font-size:12px; opacity:.82; }
      .dup-rulemodal .rp-hint{ font-size: 12px; opacity: .8; margin-top: 8px; line-height: 1.35; }
      .dup-rulemodal .rp-help{ font-size: 12px; opacity: .9; line-height: 1.4; }
      .dup-rulemodal .rp-help b{ font-weight: 700; }
      .dup-rulemodal .rp-examples{ margin-top: 8px; display:flex; flex-direction:column; gap:6px; opacity:.9; }
      .dup-rulemodal .rp-pair .rp-chk{ white-space:nowrap; font-size:12px; opacity:.85; }
      .dup-rulemodal .rp-pair .rp-chk input{ transform: translateY(1px); }
      .dup-rulemodal .rp-foot{ position: sticky; bottom: 0; background: linear-gradient(180deg, transparent, color-mix(in oklab, var(--panel, #ffffff) 92%, #000000 8%)); padding-top:10px; }


      .dup-summarybar{ z-index: 15; }

      .dup-rulemodal{ position: fixed; top: 74px; right: -420px; width: 400px; height: calc(100vh - 86px); z-index: 120; border-radius: 16px 0 0 16px; border: 1px solid color-mix(in oklab, var(--border, #d7dbe7) 70%, transparent 30%); background: color-mix(in oklab, var(--panel, #ffffff) 92%, #000 8%); box-shadow: 0 14px 40px rgba(0,0,0,.18); transition: right .22s ease; overflow: hidden; }
      .dup-rulemodal.is-open{ right: 10px; }
      .dup-rulemodal .rp-head{ display:flex; align-items:center; justify-content:space-between; gap:10px; padding:12px 12px; border-bottom:1px solid color-mix(in oklab, var(--border,#d7dbe7) 70%, transparent 30%); }
      .dup-rulemodal .rp-title{ font-weight:700; }
      .dup-rulemodal .rp-actions{ display:flex; gap:6px; flex-wrap:wrap; justify-content:flex-end; }
      .dup-rulemodal .rp-btn{ padding:6px 10px; border-radius:10px; border:1px solid color-mix(in oklab, var(--border,#d7dbe7) 70%, transparent 30%); background: transparent; color: inherit; cursor:pointer; }
      .dup-rulemodal .rp-btn--ghost{ opacity:.85; }
      .dup-rulemodal .rp-btn--add{ margin-top:8px; width:100%; }
      .dup-rulemodal .rp-btn--tiny{ padding:5px 8px; border-radius:9px; }
      .dup-rulemodal .rp-btn--primary{ background: color-mix(in oklab, var(--accent,#4c6fff) 18%, transparent 82%); border-color: color-mix(in oklab, var(--accent,#4c6fff) 55%, transparent 45%); font-weight:700; }
      .dup-rulemodal .rp-body{ height: calc(100% - 52px); overflow:auto; padding:12px; }
      .dup-rulemodal .rp-sec{ padding:10px; border-radius:14px; border:1px solid color-mix(in oklab, var(--border,#d7dbe7) 70%, transparent 30%); margin-bottom:10px; }
      .dup-rulemodal .rp-sec-title{ font-weight:700; margin-bottom:8px; }
      .dup-rulemodal .rp-row{ display:flex; gap:8px; align-items:center; margin-bottom:8px; }
      .dup-rulemodal .rp-label{ width:88px; font-size:12px; opacity:.85; }
      .dup-rulemodal .rp-input, .dup-rulemodal .rp-select{ flex:1; padding:6px 8px; border-radius:10px; border:1px solid color-mix(in oklab, var(--border,#d7dbe7) 70%, transparent 30%); background: transparent; color: inherit; }
      .dup-rulemodal .rp-input--sm{ padding:5px 8px; }
      .dup-rulemodal .rp-select--sm{ padding:5px 8px; }
      .dup-rulemodal .rp-hint{ font-size:12px; opacity:.8; line-height:1.35; }
      .dup-rulemodal .rp-card{ padding:10px; border-radius:14px; border:1px solid color-mix(in oklab, var(--border,#d7dbe7) 70%, transparent 30%); margin-bottom:8px; }
      .dup-rulemodal .rp-card-head{ display:grid; grid-template-columns: 1fr auto; gap:8px; align-items:center; }
      .dup-rulemodal .rp-card-sub{ grid-column: 1 / -1; font-size:11px; opacity:.7; margin-top:-4px; }
      .dup-rulemodal .rp-x{ border:none; background: transparent; color: inherit; font-size:18px; line-height:1; cursor:pointer; opacity:.7; }
      .dup-rulemodal .rp-group{ padding:8px; border-radius:12px; border:1px dashed color-mix(in oklab, var(--border,#d7dbe7) 60%, transparent 40%); margin-top:8px; }
      .dup-rulemodal .rp-group-head{ display:flex; justify-content:space-between; align-items:center; margin-bottom:6px; }
      .dup-rulemodal .rp-group-title{ font-weight:600; font-size:12px; opacity:.9; }
      .dup-rulemodal .rp-clause{ display:grid; grid-template-columns: 1fr 1fr 1.1fr 1.4fr auto; gap:6px; align-items:center; margin-bottom:6px; }
      .dup-rulemodal .rp-param{ min-width: 120px; }
      .dup-rulemodal .rp-pair, .dup-rulemodal .rp-exrow{ display:flex; gap:6px; align-items:center; margin-bottom:8px; }
      .dup-rulemodal .rp-vs{ opacity:.75; font-size:12px; }
      .dup-rulemodal .rp-chk{ display:flex; gap:6px; align-items:center; font-size:12px; opacity:.85; }
      .dup-rulemodal .rp-textarea{ width:100%; min-height:140px; resize:vertical; padding:8px 10px; border-radius:12px; border:1px solid color-mix(in oklab, var(--border,#d7dbe7) 70%, transparent 30%); background: transparent; color: inherit; }
      .dup-rulemodal .rp-foot{ position: sticky; bottom: 0; background: color-mix(in oklab, var(--panel,#fff) 92%, #000 8%); padding-top:8px; }

      .conn-cell{ display:inline-block; min-width: 18px; padding:2px 6px; border-radius: 999px; border: 1px solid color-mix(in oklab, var(--border, #d7dbe7) 70%, transparent 30%); }

    

      /* rulepanel alias */
      .dup-rulepanel{ display:none; }


/* v28 rule window override: make settings window truly large (like large list viewer) */
.dup-rulemodal{ z-index: 9000; }
.dup-rulemodal .rm-window{
  position: fixed !important;
  inset: 22px !important;
  width: auto !important;
  height: auto !important;
  max-width: none !important;
  max-height: none !important;
}
@media (max-width: 980px){
  .dup-rulemodal .rm-window{ inset: 10px !important; }
}
.dup-rulemodal .rm-head{ position: sticky; top: 0; z-index: 2; }
.dup-rulemodal .rp-foot{ position: sticky; bottom: 0; z-index: 2; }

/* make 'Exclude Families' stand out */
.dup-rulemodal .rp-sec--fam{
  border-color: color-mix(in oklab, var(--accent, #4c6fff) 38%, var(--border, #d7dbe7) 62%);
  background: linear-gradient(180deg,
    color-mix(in oklab, var(--accent, #4c6fff) 8%, var(--panel, #ffffff) 92%),
    color-mix(in oklab, var(--panel, #ffffff) 94%, #000000 6%));
}
.dup-rulemodal .rp-sec--fam .rp-sec-title{ justify-content: space-between; }
.dup-rulemodal .rp-sec--fam .rp-pill{
  border-color: color-mix(in oklab, var(--accent, #4c6fff) 55%, var(--border, #d7dbe7) 45%);
}
.dup-rulemodal .fam-row{
  display:flex; gap:10px; flex-wrap:wrap; align-items:center;
}
.dup-rulemodal .fam-row .rp-select{ min-width: 320px; flex: 1 1 320px; }
.dup-rulemodal .fam-chips{ display:flex; gap:8px; flex-wrap:wrap; margin-top: 10px; }
.dup-rulemodal .fam-chip{
  padding: 6px 10px; border-radius: 999px;
  border: 1px solid color-mix(in oklab, var(--border, #d7dbe7) 70%, transparent 30%);
  background: color-mix(in oklab, var(--panel, #ffffff) 94%, #000000 6%);
  cursor: pointer;
}
.dup-rulemodal .fam-chip:hover{
  border-color: color-mix(in oklab, var(--accent, #4c6fff) 40%, var(--border, #d7dbe7) 60%);
}



/* v33 results table layout: prevent vertical '작업', keep buttons horizontal, reduce awkward whitespace */
.dup-grp{
  max-width: 1120px;
  margin-left: auto;
  margin-right: auto;
}

.dup-subhead, .dup-row{
  display: grid;
  grid-template-columns: 220px 110px 150px minmax(220px, 1fr) minmax(240px, 1.2fr) 72px;
  align-items: center;
  column-gap: 12px;
}
@media (max-width: 1180px){
  .dup-subhead, .dup-row{
    grid-template-columns: 200px 104px 140px minmax(180px, 1fr) minmax(200px, 1fr) 64px;
  }
}
@media (max-width: 980px){
  .dup-subhead, .dup-row{
    grid-template-columns: 180px 96px 128px minmax(160px, 1fr) minmax(160px, 1fr) 56px;
    column-gap: 10px;
  }
}

.dup-subhead .cell, .dup-row .cell{
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  word-break: keep-all;
}
.dup-subhead .cell{
  font-weight: 650;
}

/* 작업(액션) 컬럼 */
.dup-row .row-actions{
  display: flex;
  flex-direction: row;
  align-items: center;
  justify-content: flex-start;
  gap: 10px;
  flex-wrap: nowrap;
}
.dup-row .row-actions .table-action-btn{
  white-space: nowrap;
  min-width: 88px;
  height: 32px;
  line-height: 32px;
  padding: 0 12px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
}

/* 연결 컬럼: 가운데 정렬 + 불필요한 빈 공간 감소 */
.dup-subhead .cell.conn, .dup-row .cell.conn{
  justify-self: center;
  text-align: center;
}


/* v33 conn fallback: if no .conn class is present, assume last column is connection */
.dup-subhead > .cell:last-child, .dup-row > .cell:last-child{
  justify-self: center;
  text-align: center;
}

      /* 카드 내부 패딩/행 높이 살짝 정리 */
.dup-subhead .cell{ padding-top: 10px; padding-bottom: 10px; }
.dup-row .cell{ padding-top: 10px; padding-bottom: 10px; }



/* v38 exclude lists */
.exfam-wrap{
  display:grid;
  grid-template-columns: 1fr 1fr;
  gap: 12px;
  margin-top: 10px;
}
@media (max-width: 980px){
  .exfam-wrap{ grid-template-columns: 1fr; }
}
.exfam-col{
  border: 1px solid color-mix(in oklab, var(--border, #d7dbe7) 70%, transparent 30%);
  border-radius: 14px;
  background: color-mix(in oklab, var(--panel, #ffffff) 96%, #000000 4%);
  overflow:hidden;
}
.exfam-title{
  padding: 10px 12px;
  font-weight: 700;
  border-bottom: 1px solid color-mix(in oklab, var(--border, #d7dbe7) 70%, transparent 30%);
  background: color-mix(in oklab, var(--panel, #ffffff) 90%, var(--accent, #4c6fff) 10%);
}
.exfam-list{ padding: 10px 12px; max-height: 260px; overflow:auto; }
.exfam-item{ display:flex; gap:10px; align-items:center; padding: 6px 4px; border-radius: 10px; }
.exfam-item:hover{ background: color-mix(in oklab, var(--accent, #4c6fff) 8%, transparent 92%); }
.exfam-text{ flex:1; min-width:0; overflow:hidden; text-overflow: ellipsis; white-space: nowrap; }
.exfam-empty{ opacity:.75; font-size: 12px; padding: 6px 0; }
.exfam-actions{ display:flex; gap:8px; margin-top: 10px; justify-content:flex-end; }

/* section background tints (lighter than gray) */
.dup-rulemodal .rp-sec--help{ background: color-mix(in oklab, var(--panel, #ffffff) 92%, var(--accent, #4c6fff) 8%); }
.dup-rulemodal .rp-sec--global{ background: color-mix(in oklab, var(--panel, #ffffff) 92%, #34c759 8%); }
.dup-rulemodal .rp-sec--fam{ background: color-mix(in oklab, var(--panel, #ffffff) 92%, #ff9f0a 8%); }
.dup-rulemodal .rp-sec--sets{ background: color-mix(in oklab, var(--panel, #ffffff) 92%, #af52de 8%); }
.dup-rulemodal .rp-sec--pairs{ background: color-mix(in oklab, var(--panel, #ffffff) 92%, #ff3b30 6%); }
.dup-rulemodal .rp-sec--ex{ background: color-mix(in oklab, var(--panel, #ffffff) 92%, #ff3b30 8%); }
.dup-rulemodal .rp-sec--io{ background: color-mix(in oklab, var(--panel, #ffffff) 92%, #8e8e93 8%); }



/* v39 settings contrast: make modal background distinct from sections */
.dup-rulemodal .rm-window{
  background: color-mix(in oklab, var(--panel, #ffffff) 98%, #000000 2%);
}
.dup-rulemodal .rm-body{
  background: color-mix(in oklab, var(--panel, #ffffff) 94%, #000000 6%);
}
.dup-rulemodal .rp-sec{
  border-color: color-mix(in oklab, var(--border, #d7dbe7) 80%, transparent 20%);
  background: color-mix(in oklab, var(--panel, #ffffff) 99%, transparent 1%);
  box-shadow: 0 1px 0 rgba(0,0,0,.04);
}

/* Exclude picker modal shares same styling */
.dup-expicker{ position: fixed; inset: 0; z-index: 9100; display:none; }
.dup-expicker.is-open{ display:block; }
.dup-expicker .rm-backdrop{ position:absolute; inset:0; background: rgba(0,0,0,.22); }
.dup-expicker .rm-window{ position: fixed; inset: 22px; border-radius: 18px; overflow:hidden;
  border: 1px solid color-mix(in oklab, var(--border, #d7dbe7) 70%, transparent 30%);
  background: color-mix(in oklab, var(--panel, #ffffff) 98%, #000000 2%);
  box-shadow: 0 20px 60px rgba(0,0,0,.28);
  display:flex; flex-direction:column;
}
.dup-expicker .rm-head{ position: sticky; top: 0; z-index:2; }
.dup-expicker .rm-body{ padding: 14px 16px 64px 16px; overflow:auto; background: color-mix(in oklab, var(--panel, #ffffff) 94%, #000000 6%); }



/* v40 button contrast */
.rp-btn{
  background: color-mix(in oklab, var(--panel, #ffffff) 96%, #000000 4%);
  border-color: color-mix(in oklab, var(--border, #d7dbe7) 78%, transparent 22%);
  color: inherit;
}
.rp-btn--ghost{
  background: transparent;
}
.rp-btn--add{
  background: color-mix(in oklab, var(--panel, #ffffff) 92%, var(--accent, #4c6fff) 8%);
  border-color: color-mix(in oklab, var(--accent, #4c6fff) 34%, var(--border, #d7dbe7) 66%);
}
.rp-btn--primary{
  color: #fff;
}

/* pair table */
.pair-table{
  margin: 10px 12px 14px 12px;
  border: 1px solid color-mix(in oklab, var(--border, #d7dbe7) 75%, transparent 25%);
  border-radius: 14px;
  overflow:hidden;
  background: color-mix(in oklab, var(--panel, #ffffff) 98%, #000000 2%);
}
.pair-subhead, .pair-row{
  display:grid;
  grid-template-columns: 120px minmax(220px, 1fr) 120px minmax(220px, 1fr);
  column-gap: 12px;
  align-items:center;
}
.pair-subhead{
  background: color-mix(in oklab, var(--accent, #4c6fff) 8%, var(--panel, #ffffff) 92%);
  font-weight: 650;
}
.pair-subhead .cell, .pair-row .cell{
  padding: 10px 12px;
  min-width:0;
  overflow:hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.pair-row{
  border-top: 1px solid color-mix(in oklab, var(--border, #d7dbe7) 75%, transparent 25%);
}



/* v41 header stabilize: prevent actions jumping between duplicate/clash */
.feature-header.dup-toolbar{
  display:flex;
  align-items:flex-start;
  justify-content: space-between;
  gap: 16px;
  flex-wrap: nowrap;
}
.feature-heading{ min-width: 0; }
.feature-sub{
  min-height: 40px; /* keep same height across modes */
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
}
.feature-actions{
  flex: 0 0 auto;
  display:flex;
  align-items:center;
  gap: 10px;
  flex-wrap: nowrap;
}

/* v41 stronger button contrast in settings window */
.dup-rulemodal .rp-btn{
  background: color-mix(in oklab, var(--panel, #ffffff) 99%, #000000 1%) !important;
  border-color: color-mix(in oklab, var(--border, #d7dbe7) 82%, transparent 18%) !important;
}
.dup-rulemodal .rp-btn--add{
  background: color-mix(in oklab, var(--accent, #4c6fff) 14%, var(--panel, #ffffff) 86%) !important;
  border-color: color-mix(in oklab, var(--accent, #4c6fff) 42%, var(--border, #d7dbe7) 58%) !important;
}
.dup-rulemodal .rp-btn--primary{
  background: color-mix(in oklab, var(--accent, #4c6fff) 92%, #000000 8%) !important;
  border-color: color-mix(in oklab, var(--accent, #4c6fff) 78%, transparent 22%) !important;
}
.dup-rulemodal .rp-btn--ghost{
  background: transparent !important;
}



.dup-expicker .rp-btn{
  background: color-mix(in oklab, var(--panel, #ffffff) 99%, #000000 1%) !important;
  border-color: color-mix(in oklab, var(--border, #d7dbe7) 82%, transparent 18%) !important;
}
.dup-expicker .rp-btn--primary{
  background: color-mix(in oklab, var(--accent, #4c6fff) 92%, #000000 8%) !important;
  border-color: color-mix(in oklab, var(--accent, #4c6fff) 78%, transparent 22%) !important;
  color:#fff;
}
`;
    document.head.appendChild(st);
  }

  const topbarEl = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar');
  if (topbarEl) topbarEl.classList.add('hub-topbar');

  const page    = div('dup-page feature-shell');
  const header  = div('feature-header dup-toolbar');
  const heading = div('feature-heading');

  let mode = readMode();
  let scopeMode = readScopeMode();
  let exclKwInputEl = null;

  let tolInputEl = null;
  let rulePanelEl = null;
  let rulePanelOpen = false;
  let ruleCfg = loadRuleConfig();

  let modeBtns = [];
  let activeModeForView = mode;

  const runBtn    = cardBtn('검토 시작', onRun);
  const exportBtn = cardBtn('엑셀 내보내기', onExport);
  exportBtn.disabled = true;

  const modeBar = buildModeBar();



const settingsBtn = kbtn('규칙/Set', 'subtle', () => {
  onOpenRulePanel();
  try { syncSettingsBtn(); } catch {}
});
settingsBtn.title = '규칙/Set 창 열기/닫기';

  const actions = div('feature-actions');
  actions.append(runBtn, modeBar, settingsBtn, exportBtn);

  header.append(heading, actions);
  page.append(header);

  rulePanelEl = buildRulePanel();
  page.append(rulePanelEl);

  // Exclude picker(별도 창)
  try { exPickerEl = buildExcludePicker(); page.append(exPickerEl); } catch {}

  try { syncSettingsBtn(); } catch {}


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
  let lastPairs = [];
  let metaModelFamilies = [];
  let metaSystemCategories = [];
  let lastTruncToastKey = '';

  applyHeadingByMode(mode);
  renderIntro(body, mode);

onHost('dup:meta', (payload) => {
  try {
    const meta = payload || {};
    try { localStorage.setItem(DUP_META_KEY, JSON.stringify(meta)); } catch {}
    toast('목록을 갱신했습니다.', 'ok', 1800);
    try { renderRulePanelMeta(); } catch {}
  } catch {}
});

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

    if (ev === 'dup:pairs') {
      lastPairs = Array.isArray(payload) ? payload : [];
      try { paintGroups(); } catch {}
      try { refreshSummary(); } catch {}
      return;
    }

    if (ev === 'dup:meta') {
      const p = payload || {};
      metaModelFamilies = Array.isArray(p.modelFamilies) ? p.modelFamilies : [];
      metaSystemCategories = Array.isArray(p.systemCategories) ? p.systemCategories : [];
      try { renderRulePanel(); } catch {}
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
    if (exclKwInputEl) exclKwInputEl.disabled = on;

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
    const excludeKeywords = getExcludeKeywordList();
    const ruleConfig = loadRuleConfig();
    const famKw = Array.isArray(ruleConfig.excludeFamilies) ? ruleConfig.excludeFamilies : [];
    const mergedExcludeKeywords = ensureListUnique([ ...(excludeKeywords||[]), ...famKw ]);
    post('dup:run', { mode, tolFeet, scopeMode, excludeKeywords: mergedExcludeKeywords, ruleConfig });
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

  

function paintPairGroups() {
  const pairs = Array.isArray(lastPairs) ? lastPairs : [];
  if (!pairs.length) return;

  // group by groupKey
  const map = new Map();
  for (const p of pairs) {
    const gk = String(p.groupKey || 'C0000');
    if (!map.has(gk)) map.set(gk, []);
    map.get(gk).push(p);
  }

  const sorted = Array.from(map.entries()).sort((a,b)=>b[1].length-a[1].length);

  for (let gi=0; gi<sorted.length; gi++) {
    const [gk, items] = sorted[gi];

    const card = div('dup-grp');
    const head = div('grp-h');

    const left = div('grp-txt');
    const title = div('grp-title');
    title.textContent = `간섭 그룹 ${gi+1} (${gk})`;

    const count = div('grp-count');
    count.textContent = `${items.length} 쌍`;

    left.append(title, count);
    head.append(left);
    card.append(head);

    const table = div('pair-table');
    const sub = div('pair-subhead');
    sub.innerHTML = `
      <div class="cell">A ID</div>
      <div class="cell">A 정보</div>
      <div class="cell">B ID</div>
      <div class="cell">B 정보</div>
    `;
    table.append(sub);

    for (const p of items) {
      const row = div('pair-row');

      const aBtn = document.createElement('button');
      aBtn.className = 'table-action-btn';
      aBtn.textContent = String(p.aId || '—');
      aBtn.onclick = () => post(EV_SELECT_REQ, { id: Number(p.aId) });

      const bBtn = document.createElement('button');
      bBtn.className = 'table-action-btn';
      bBtn.textContent = String(p.bId || '—');
      bBtn.onclick = () => post(EV_SELECT_REQ, { id: Number(p.bId) });

      const aInfo = `${p.aCategory || ''} · ${p.aFamily || ''}${p.aType ? ' : ' + p.aType : ''}`.trim();
      const bInfo = `${p.bCategory || ''} · ${p.bFamily || ''}${p.bType ? ' : ' + p.bType : ''}`.trim();

      row.append(
        cell(aBtn, 'cell'),
        cell(aInfo || '—', 'cell'),
        cell(bBtn, 'cell'),
        cell(bInfo || '—', 'cell'),
      );
      table.append(row);
    }

    card.append(table);
    out.append(card);
  }
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
        cell('연결', 'th'),
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

    // 연결(간섭 상대) 표시: count + hover ids
    const conn = document.createElement('div');
    conn.className = 'conn-cell';
    const ids = (r.connectedIds || []).map(String);
    const cnt = ids.length;
    conn.textContent = cnt ? String(cnt) : '0';
    if (cnt) conn.title = ids.join(', ');
    row.append(cell(conn, 'td mono right'));

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


function buildScopeBar() {
  const wrap = div('dup-modebar');
  wrap.title = 'Navisworks Set처럼: 선택집합을 범위로 사용하거나(선택만), 선택집합을 제외할 수 있습니다.';
  const bAll = kbtn('전체', 'subtle', () => setScopeMode('all'));
  const bScope = kbtn('선택만', 'subtle', () => setScopeMode('scope'));
  const bExcl = kbtn('선택 제외', 'subtle', () => setScopeMode('exclude'));
  const btns = [bAll, bScope, bExcl];

  function sync() {
    bAll.classList.toggle('is-active', scopeMode === 'all');
    bScope.classList.toggle('is-active', scopeMode === 'scope');
    bExcl.classList.toggle('is-active', scopeMode === 'exclude');
  }

  function setScopeMode(m) {
    if (!['all','scope','exclude'].includes(m)) return;
    scopeMode = m;
    try { localStorage.setItem(DUP_SCOPE_KEY, scopeMode); } catch {}
    sync();
  }

  sync();
  wrap.append(bAll, bScope, bExcl);
  return wrap;

  try { paintGroups(); } catch {}
}

function buildExcludeKeywordControl() {
  const wrap = div('dup-tol');
  wrap.title = '제외 키워드(콤마 구분). Family/Type/Category/Name에 포함되면 결과에서 제외됩니다.';
  const label = document.createElement('span');
  label.className = 'dup-tol-label';
  label.textContent = '제외 키워드';

  const input = document.createElement('input');
  input.className = 'dup-tol-input';
  input.type = 'text';
  input.placeholder = '예: Dummy, Temp';
  input.value = readExcludeKeywords();

  input.addEventListener('change', () => {
    try { localStorage.setItem(DUP_EXCL_KW_KEY, String(input.value || '').trim()); } catch {}
  });

  exclKwInputEl = input;
  wrap.append(label, input);
  return wrap;
}


function defaultRuleConfig() {
  return { version: 1, sets: [], pairs: [], excludeSetIds: [], excludeFamilies: [], excludeCategories: [] };
}
function normalizeRuleConfig(cfg) {
  const c = cfg && typeof cfg === 'object' ? cfg : {};
  const out = {
    version: 1,
    sets: Array.isArray(c.sets) ? c.sets : [],
    pairs: Array.isArray(c.pairs) ? c.pairs : [],
    excludeSetIds: Array.isArray(c.excludeSetIds) ? c.excludeSetIds : [],
    excludeFamilies: Array.isArray(c.excludeFamilies) ? c.excludeFamilies : [],
    excludeCategories: Array.isArray(c.excludeCategories) ? c.excludeCategories : []
  };

  // normalize sets
  out.sets = out.sets.map((s, i) => {
    const ss = s && typeof s === 'object' ? s : {};
    return {
      id: String(ss.id || `S${i + 1}`),
      name: String(ss.name || ss.id || `Set ${i + 1}`),
      logic: ss.logic === 'and' || ss.logic === 'or' ? ss.logic : 'or', // group logic (OR default)
      groups: Array.isArray(ss.groups) ? ss.groups : []               // groups = OR list; each group has clauses(AND)
    };
  });

  // normalize groups/clauses
  out.sets.forEach(s => {
    s.groups = s.groups.map(g => {
      const gg = g && typeof g === 'object' ? g : {};
      return {
        clauses: Array.isArray(gg.clauses) ? gg.clauses : []
      };
    });
    s.groups.forEach(g => {
      g.clauses = g.clauses.map(cl => {
        const cc = cl && typeof cl === 'object' ? cl : {};
        return {
          field: String(cc.field || 'category'), // category|family|type|name|param
          op: String(cc.op || 'contains'),       // contains|equal|...
          value: String(cc.value || ''),
          param: String(cc.param || '')
        };
      });
    });
  });

  // normalize pairs
  out.pairs = out.pairs.map(p => {
    const pp = p && typeof p === 'object' ? p : {};
    return { a: String(pp.a || ''), b: String(pp.b || '') };
  }).filter(p => p.a && p.b);

  // exclude ids
  out.excludeSetIds = out.excludeSetIds.map(String).filter(Boolean);

  // exclude families
  out.excludeFamilies = out.excludeFamilies.map(String).filter(Boolean);

  return out;
}

function loadRuleConfig() {
  try {
    const raw = localStorage.getItem(DUP_RULECFG_KEY);
    if (!raw) return defaultRuleConfig();
    const cfg = JSON.parse(raw);
    return normalizeRuleConfig(cfg);
  } catch {
    return defaultRuleConfig();
  }
}

function saveRuleConfig(cfg) {
  const norm = normalizeRuleConfig(cfg);
  try { localStorage.setItem(DUP_RULECFG_KEY, JSON.stringify(norm)); } catch {}
  return norm;
}

function loadMeta() {
  try {
    const raw = localStorage.getItem(DUP_META_KEY);
    if (!raw) return {};
    const meta = JSON.parse(raw);
    return meta && typeof meta === 'object' ? meta : {};
  } catch {
    return {};
  }
}

function readScopeMode() {
  try {
    const m = String(localStorage.getItem(DUP_SCOPE_KEY) || '').trim();
    if (m === 'all' || m === 'scope' || m === 'exclude') return m;
  } catch {}
  return 'all';
}

function readExcludeKeywords() {
  try { return String(localStorage.getItem(DUP_EXCL_KW_KEY) || '').trim(); } catch {}
  return '';
}

function getExcludeKeywordList() {
  const raw = String(exclKwInputEl ? exclKwInputEl.value : readExcludeKeywords());
  return raw.split(',').map(s => s.trim()).filter(Boolean);
}

function onOpenRulePanel(forceOpen) {
  if (typeof forceOpen === 'boolean') rulePanelOpen = forceOpen;
  else rulePanelOpen = !rulePanelOpen;

  if (rulePanelEl) rulePanelEl.classList.toggle('is-open', !!rulePanelOpen);

  if (rulePanelOpen) {
    try { renderRulePanel(); } catch {}
    try { rulePanelEl.setAttribute('tabindex','-1'); rulePanelEl.focus(); } catch {}
  }
  try { syncSettingsBtn(); } catch {}
}

function syncSettingsBtn() {
  try {
    if (typeof settingsBtn !== 'undefined' && settingsBtn) {
      settingsBtn.classList.toggle('is-active', !!rulePanelOpen);
    }
  } catch {}
}

function buildExcludePicker() {
  const wrap = div('dup-expicker');
  wrap.innerHTML = `
    <div class="rm-backdrop" data-act="close"></div>
    <div class="rm-window" role="dialog" aria-label="Exclude 목록 선택">
      <div class="rm-head">
        <div class="rm-head-left">
          <div class="rm-title">Exclude 목록 선택</div>
          <div class="rm-sub">모델에 존재하는 패밀리 / 시스템 카테고리를 검색해서 제외 목록으로 추가합니다.</div>
        </div>
        <div class="rm-actions">
          <button class="rp-btn rp-btn--ghost" data-act="refresh">목록 새로고침</button>
          <button class="rp-btn" data-act="close">닫기</button>
        </div>
      </div>

      <div class="rm-body">
        <div class="rp-sec rp-sec--global">
          <div class="rp-sec-title">검색 <span class="rp-pill">Filter</span></div>
          <div class="rp-grid">
            <div class="rp-label">키워드</div>
            <input class="rp-input" type="text" placeholder="검색…" data-bind="q"/>
          </div>
          <div class="rp-hint">체크 후 아래 “적용”을 누르면 설정창의 Exclude 목록에 반영됩니다.</div>
        </div>

        <div class="exfam-wrap">
          <div class="exfam-col">
            <div class="exfam-title">모델 패밀리</div>
            <div class="exfam-list" data-slot="fam"></div>
          </div>
          <div class="exfam-col">
            <div class="exfam-title">시스템(카테고리)</div>
            <div class="exfam-list" data-slot="cat"></div>
          </div>
        </div>

        <div class="exfam-actions">
          <button class="rp-btn rp-btn--ghost" data-act="all">전체 체크</button>
          <button class="rp-btn rp-btn--ghost" data-act="none">전체 해제</button>
        </div>

        <div class="rp-foot" style="margin-top:12px;">
          <button class="rp-btn rp-btn--primary" data-act="apply">적용</button>
        </div>
      </div>
    </div>
  `;

  wrap.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-act]');
    if (!btn) return;
    e.preventDefault();
    const act = btn.dataset.act;

    if (act === 'close') { closeExcludePicker(); return; }
    if (act === 'refresh') { post('dup:run', { mode, metaOnly: true }); return; }
    if (act === 'all') { for (const cb of wrap.querySelectorAll('input[type="checkbox"]')) cb.checked = true; return; }
    if (act === 'none') { for (const cb of wrap.querySelectorAll('input[type="checkbox"]')) cb.checked = false; return; }
    if (act === 'apply') { applyExcludePicker(); return; }
  });

  wrap.addEventListener('input', (e) => {
    if (e.target && e.target.matches('[data-bind="q"]')) renderExcludePicker();
  });

  return wrap;
}

let exPickerEl = null;
let exPickerOpen = false;

function openExcludePicker() {
  exPickerOpen = true;
  if (exPickerEl) exPickerEl.classList.add('is-open');
  // auto refresh if empty
  if ((!metaModelFamilies || !metaModelFamilies.length) && (!metaSystemCategories || !metaSystemCategories.length)) {
    post('dup:run', { mode, metaOnly: true });
  }
  renderExcludePicker();
}

function closeExcludePicker() {
  exPickerOpen = false;
  if (exPickerEl) exPickerEl.classList.remove('is-open');
}

function renderExcludePicker() {
  if (!exPickerEl) return;
  const cfg = loadRuleConfig();
  const famSel = new Set((cfg.excludeFamilies || []).map(String));
  const catSel = new Set((cfg.excludeCategories || []).map(String));

  const qEl = exPickerEl.querySelector('[data-bind="q"]');
  const q = (qEl ? qEl.value : '').trim().toLowerCase();

  const famBox = exPickerEl.querySelector('[data-slot="fam"]');
  const catBox = exPickerEl.querySelector('[data-slot="cat"]');
  if (!famBox || !catBox) return;

  const fams = (Array.isArray(metaModelFamilies) ? metaModelFamilies : []).map(String).filter(Boolean);
  const cats = (Array.isArray(metaSystemCategories) ? metaSystemCategories : []).map(String).filter(Boolean);

  function mk(kind, val, checked) {
    const safe = String(val).replace(/</g,'&lt;').replace(/>/g,'&gt;');
    return `<label class="exfam-item"><input type="checkbox" data-kind="${kind}" value="${safe}" ${checked?'checked':''}/><span class="exfam-text">${safe}</span></label>`;
  }

  const famHtml = [];
  for (const f of fams) {
    if (q && !f.toLowerCase().includes(q)) continue;
    famHtml.push(mk('fam', f, famSel.has(f)));
  }
  famBox.innerHTML = famHtml.length ? famHtml.join('') : `<div class="exfam-empty">목록이 비어있습니다. “목록 새로고침”을 누르세요.</div>`;

  const catHtml = [];
  for (const c of cats) {
    if (q && !c.toLowerCase().includes(q)) continue;
    catHtml.push(mk('cat', c, catSel.has(c)));
  }
  catBox.innerHTML = catHtml.length ? catHtml.join('') : `<div class="exfam-empty">목록이 비어있습니다. “목록 새로고침”을 누르세요.</div>`;
}

function applyExcludePicker() {
  if (!exPickerEl) return;
  const cfg = loadRuleConfig();

  cfg.excludeFamilies = Array.from(exPickerEl.querySelectorAll('input[data-kind="fam"]:checked')).map(x => x.value);
  cfg.excludeCategories = Array.from(exPickerEl.querySelectorAll('input[data-kind="cat"]:checked')).map(x => x.value);

  saveRuleConfig(cfg);
  try { toast('Exclude 목록이 적용되었습니다.', 'ok', 1200); } catch {}
  renderRulePanel();
  closeExcludePicker();
}

function buildRulePanel() {
  const wrap = div('dup-rulemodal');
  wrap.innerHTML = `
    <div class="rm-backdrop" data-act="close"></div>
    <div class="rm-window" role="dialog" aria-label="규칙/Set 설정">
      <div class="rm-head">
        <div class="rm-head-left">
          <div class="rm-title">규칙 / Set 설정</div>
          <div class="rm-sub">자체간섭(클래시) 상위호환 · Set/Rule/허용오차를 한 곳에서 관리합니다.</div>
        </div>
        <div class="rm-actions">
          <button class="rp-btn rp-btn--ghost" data-act="refresh">목록 새로고침</button>
          <button class="rp-btn rp-btn--ghost" data-act="export">Export</button>
          <button class="rp-btn rp-btn--ghost" data-act="import">Import</button>
          <button class="rp-btn" data-act="close">닫기</button>
        </div>
      </div>

      <div class="rm-body">
        <section class="rp-sec rp-sec--help">
          <div class="rp-sec-title">작동 방식 <span class="rp-pill">순서</span></div>
          <div class="rp-help">
            <ol>
              <li><b>범위(Selection)</b>로 검사 대상 풀(U)을 먼저 정합니다. <span class="rp-muted">(전체 / 선택만 / 선택 제외)</span></li>
              <li><b>제외(키워드/Exclude Sets)</b>로 U에서 1차 제거합니다.</li>
              <li><b>Pair(Set vs Set)</b>가 있으면 Pair만 간섭을 계산합니다. (없으면 U 전체끼리 계산)</li>
            </ol>
            <div class="rp-examples">
              <div><b>예1)</b> Dummy 무시: Dummy 선택 → <b>범위=선택 제외</b> → 검토 시작</div>
              <div><b>예2)</b> 장비만 검사: 장비 선택 → <b>범위=선택만</b> → 검토 시작</div>
            </div>
          </div>
        </section>

        <section class="rp-sec rp-sec--fam">
          <div class="rp-sec-title">공통 설정 <span class="rp-pill">Global</span></div>
          <div class="rp-grid">
            <div class="rp-label">허용오차(mm)</div>
            <input class="rp-input" type="number" step="0.1" min="0.01" data-bind="tolMm"/>

            <div class="rp-label">범위(Selection)</div>
            <select class="rp-select" data-bind="scopeMode">
              <option value="all">전체</option>
              <option value="scope">선택한 요소만 검사</option>
              <option value="exclude">선택한 요소는 제외</option>
            </select>

            <div class="rp-label">제외 키워드</div>
            <input class="rp-input" type="text" placeholder="예: Dummy, Temp, _TEST (콤마 구분)" data-bind="excludeKeywords"/>
          </div>
          <div class="rp-hint">※ 범위(Selection)는 Revit에서 요소를 선택한 뒤 <b>검토 시작</b>을 누르면 적용됩니다. Set/Rule은 이 범위 안에서만 동작합니다.</div>
        </section>

        <section class="rp-sec">
          <div class="rp-sec-title">Set 정의 <span class="rp-pill">Selection Set</span></div>
          <div class="rp-hint">Set은 <b>(그룹 OR) - (조건 AND)</b> 구조입니다. “Parameter” 조건은 목록 새로고침 후 선택할 수 있습니다.</div>
          <div class="rp-sets" data-slot="sets"></div>
          <button class="rp-btn rp-btn--add" data-act="add-set">+ Set 추가</button>
        </section>

        <section class="rp-sec">
          <div class="rp-sec-title">Pair (Set vs Set) <span class="rp-pill">Rules</span></div>
          <div class="rp-help">
            <div>• 체크된 Pair만 적용됩니다. (체크 해제 = 비활성)</div>
            <div>• <b>A vs B</b>: A에 속한 요소 ↔ B에 속한 요소 간 간섭만 계산</div>
            <div>• <b>A vs A</b>: A 내부 요소들끼리(같은 Set 내부) 간섭을 계산 <span class="rp-muted">(제외가 아님)</span></div>
            <div>• <b>__ALL__</b>: Set ↔ 전체(범위 U) 의미</div>
          </div>
          <div class="rp-pairs" data-slot="pairs"></div>
          <button class="rp-btn rp-btn--add" data-act="add-pair">+ Pair 추가</button>
        </section>

        <section class="rp-sec">
          <div class="rp-sec-title">Exclude Sets <span class="rp-pill">Ignore</span></div>
          <div class="rp-hint">등록된 Set에 포함되는 요소는 모든 간섭 검토에서 제외됩니다. (Dummy 제외 등에 사용)</div>
          <div class="rp-ex" data-slot="ex"></div>
          <button class="rp-btn rp-btn--add" data-act="add-ex">+ Exclude Set 추가</button>
        </section>

        
<section class="rp-sec rp-sec--fam">
  <div class="rp-sec-title">Exclude 목록 <span class="rp-pill">Ignore</span></div>
  <div class="rp-hint">
    체크된 항목(패밀리/시스템 카테고리)에 속한 요소는 간섭 검토에서 제외됩니다. (하위/중첩 포함 목적)
  </div>

  <div class="rp-grid">
    <div class="rp-label">선택</div>
    <div style="display:flex; gap:10px; flex-wrap:wrap; align-items:center;">
      <button class="rp-btn rp-btn--add" data-act="open-expicker">패밀리/시스템 목록 열기…</button>
      <button class="rp-btn rp-btn--ghost" data-act="refresh">목록 새로고침</button>
    </div>

    <div class="rp-label">현재 제외</div>
    <div class="rp-hint" data-slot="exsum">—</div>
  </div>

  <div class="rp-hint" style="margin-top:8px;">
    ※ 목록이 비어있으면 “목록 새로고침”을 누르세요. 목록이 매우 길면 별도 창에서 검색/체크 후 적용합니다.
  </div>
</section>

<section class="rp-sec">
          <div class="rp-sec-title">Import / Export <span class="rp-pill">JSON</span></div>
          <textarea class="rp-textarea" data-bind="json"></textarea>
          <div class="rp-hint">Export는 현재 설정을 JSON으로 생성합니다. Import는 JSON을 붙여넣고 Import 버튼을 누르면 적용됩니다.</div>
        </section>

        <div class="rp-foot">
          <button class="rp-btn rp-btn--primary" data-act="apply">적용</button>
        </div>
      </div>
    </div>
  `;

  // click actions (backdrop 포함)
  wrap.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-act]');
    if (!btn) return;
    e.preventDefault();
    const act = btn.dataset.act;

    if (act === 'close') { onOpenRulePanel(false); return; }
    if (act === 'refresh') { post('dup:run', { mode, metaOnly: true }); return; }
    if (act === 'add-set') { addSet(); renderRulePanel(); return; }
    if (act === 'add-pair') { addPair(); renderRulePanel(); return; }
    if (act === 'add-ex') { addExclude(); renderRulePanel(); return; }
    if (act === 'open-expicker') { openExcludePicker(); return; }
    if (act === 'add-exfam') { /* deprecated */ addExcludeFamilyFromPick(); renderRulePanel(); return; }
    if (act === 'export') { exportRuleJson(); return; }
    if (act === 'import') { importRuleJson(); return; }
    if (act === 'apply') { applyRulePanel(); return; }
  });

  // ESC to close
  wrap.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { onOpenRulePanel(false); }
  });

  return wrap;
}
function renderRulePanelMeta() {
  if (!rulePanelEl || !rulePanelOpen) return;
  renderRulePanel();
}

function ensureId(prefix='S') {
  return prefix + Math.random().toString(16).slice(2, 10);
}

function addSet() {
  ruleCfg = ruleCfg || { sets: [], pairs: [], excludeSetIds: [] };
  ruleCfg.sets = Array.isArray(ruleCfg.sets) ? ruleCfg.sets : [];
  ruleCfg.sets.push({ id: ensureId('S'), name: 'Set', groups: [{ clauses: [{ field: 'category', op: 'contains', value: '', param: '' }] }] });
  saveRuleConfig(ruleCfg);
  try { toast('Set가 추가되었습니다.', 'ok', 1200); } catch {}
}

function addPair() {
  ruleCfg = ruleCfg || { sets: [], pairs: [], excludeSetIds: [] };
  ruleCfg.pairs = Array.isArray(ruleCfg.pairs) ? ruleCfg.pairs : [];
  const a = (ruleCfg.sets[0]?.id) || '__ALL__';
  const b = (ruleCfg.sets[0]?.id) || '__ALL__';
  ruleCfg.pairs.push({ a, b, enabled: true });
  saveRuleConfig(ruleCfg);
  try { toast('Pair가 추가되었습니다.', 'ok', 1200); } catch {}
}

function addExclude() {
  if (!ruleCfg.sets || !ruleCfg.sets.length) {
    try { toast('Set이 없습니다. 먼저 Set을 추가하세요.', 'warn', 1600); } catch {}
    return;
  }
  ruleCfg.excludeSetIds = Array.isArray(ruleCfg.excludeSetIds) ? ruleCfg.excludeSetIds : [];
  const a = (ruleCfg.sets[0]?.id) || '';
  if (a) ruleCfg.excludeSetIds.push(a);
  saveRuleConfig(ruleCfg);
  try { toast('Exclude Set이 추가되었습니다.', 'ok', 1200); } catch {}
}

function ensureListUnique(arr) {
  const out = [];
  const seen = new Set();
  (arr || []).forEach(v => {
    const s = String(v || '').trim();
    if (!s) return;
    const k = s.toLowerCase();
    if (seen.has(k)) return;
    seen.add(k);
    out.push(s);
  });
  return out;
}

function getFamiliesFromMeta(meta) {
  const m = meta || {};
  const a = Array.isArray(m.modelFamilies) ? m.modelFamilies
          : Array.isArray(m.familiesInModel) ? m.familiesInModel
          : Array.isArray(m.families) ? m.families
          : [];
  return ensureListUnique(a).sort((x,y)=>x.localeCompare(y));
}

function getFamiliesFromRows() {
  const a = ensureListUnique((rows || []).map(r => r.family || '').filter(Boolean));
  return a.sort((x,y)=>x.localeCompare(y));
}

function addExcludeFamilyFromPick() {
  ruleCfg = ruleCfg || {};
  ruleCfg.excludeFamilies = Array.isArray(ruleCfg.excludeFamilies) ? ruleCfg.excludeFamilies : [];

  const sel = rulePanelEl && rulePanelEl.querySelector('[data-bind="exFamPick"]');
  const name = sel ? String(sel.value || '').trim() : '';
  if (!name) { toast('추가할 패밀리를 선택하세요.', 'warn', 1800); return; }

  ruleCfg.excludeFamilies = ensureListUnique([ ...ruleCfg.excludeFamilies, name ]);
  saveRuleConfig(ruleCfg);
  try { toast('패밀리 제외 목록에 추가되었습니다.', 'ok', 1400); } catch {}
}

function removeExcludeFamily(name) {
  ruleCfg = ruleCfg || {};
  const cur = Array.isArray(ruleCfg.excludeFamilies) ? ruleCfg.excludeFamilies : [];
  const key = String(name || '').toLowerCase();
  ruleCfg.excludeFamilies = cur.filter(x => String(x).toLowerCase() !== key);
  saveRuleConfig(ruleCfg);
}


function exportRuleJson() {
  const ta = rulePanelEl.querySelector('[data-bind="json"]');
  if (!ta) return;
  ta.value = JSON.stringify(loadRuleConfig(), null, 2);
  try { navigator.clipboard.writeText(ta.value); } catch {}
  toast('Export JSON 생성(클립보드 복사 시도)', 'ok', 1800);
}

function importRuleJson() {
  const ta = rulePanelEl.querySelector('[data-bind="json"]');
  if (!ta) return;
  let cfg = null;
  try { cfg = JSON.parse(String(ta.value || '').trim()); } catch { cfg = null; }
  if (!cfg || typeof cfg !== 'object') { toast('Import JSON 형식이 올바르지 않습니다.', 'err', 2600); return; }
  cfg.sets = Array.isArray(cfg.sets) ? cfg.sets : [];
  cfg.pairs = Array.isArray(cfg.pairs) ? cfg.pairs : [];
  cfg.excludeSetIds = Array.isArray(cfg.excludeSetIds) ? cfg.excludeSetIds : [];
  saveRuleConfig(cfg);
  renderRulePanel();
  toast('Import 적용 완료', 'ok', 1600);
}

function applyRulePanel() {
  const tolEl = rulePanelEl.querySelector('[data-bind="tolMm"]');
  const smEl = rulePanelEl.querySelector('[data-bind="scopeMode"]');
  const kwEl = rulePanelEl.querySelector('[data-bind="excludeKeywords"]');
  // 헤더 입력 컨트롤을 쓰지 않는 버전: 패널 입력을 소스로 사용
  if (!exclKwInputEl && kwEl) exclKwInputEl = kwEl;

  // 기존 컨트롤과 동기화
  try {
    const mm = Number(tolEl?.value);
    if (tolInputEl && Number.isFinite(mm)) tolInputEl.value = fmtTolMm(sanitizeTolMm(mm));
    try { localStorage.setItem(DUP_TOL_MM_KEY, String(sanitizeTolMm(mm))); } catch {}
  } catch {}

  try {
    const sm = String(smEl?.value || 'all');
    scopeMode = sm;
    try { localStorage.setItem(DUP_SCOPE_KEY, sm); } catch {}
  } catch {}

  try {
    const kw = String(kwEl?.value || '').trim();
    if (exclKwInputEl) exclKwInputEl.value = kw;
    try { localStorage.setItem(DUP_EXCL_KW_KEY, kw); } catch {}
  } catch {}

  toast('설정 적용 완료', 'ok', 1600);
}

function renderRulePanel() {
  if (!rulePanelEl) return;

  // bind inputs
  const tol = readTolMm();
  const tolEl = rulePanelEl.querySelector('[data-bind="tolMm"]');
  if (tolEl) tolEl.value = fmtTolMm(tol);

  const smEl = rulePanelEl.querySelector('[data-bind="scopeMode"]');
  if (smEl) smEl.value = scopeMode;

  const kwEl = rulePanelEl.querySelector('[data-bind="excludeKeywords"]');
  // 헤더 입력 컨트롤을 쓰지 않는 버전: 패널 입력을 소스로 사용
  if (!exclKwInputEl && kwEl) exclKwInputEl = kwEl;
  if (kwEl) kwEl.value = readExcludeKeywords();

  const meta = loadMeta();
  const params = Array.isArray(meta.parameters) ? meta.parameters : [];

  const fams = getFamiliesFromMeta(meta);
  const famFallback = fams.length ? fams : getFamiliesFromRows();


  // sets
  const setsSlot = rulePanelEl.querySelector('[data-slot="sets"]');
  if (setsSlot) {
    setsSlot.innerHTML = '';
    (ruleCfg.sets || []).forEach((s, si) => {
      const card = document.createElement('div');
      card.className = 'rp-card';
      card.innerHTML = `
        <div class="rp-card-head">
          <input class="rp-input rp-input--sm" data-set-name="${si}" value="${esc(s.name || 'Set')}" />
          <button class="rp-x" data-del-set="${si}">×</button>
          <div class="rp-card-sub">id: ${esc(s.id || '')}</div>
        </div>
        <div class="rp-groups" data-groups="${si}"></div>
        <button class="rp-btn rp-btn--tiny" data-add-group="${si}">+ OR 그룹</button>
      `;
      setsSlot.append(card);

      const groupsWrap = card.querySelector(`[data-groups="${si}"]`);
      (s.groups || []).forEach((g, gi) => {
        const gEl = document.createElement('div');
        gEl.className = 'rp-group';
        gEl.innerHTML = `
          <div class="rp-group-head">
            <div class="rp-group-title">그룹 ${gi+1} (AND)</div>
            <button class="rp-x" data-del-group="${si}:${gi}">×</button>
          </div>
          <div class="rp-clauses" data-clauses="${si}:${gi}"></div>
          <button class="rp-btn rp-btn--tiny" data-add-clause="${si}:${gi}">+ 조건</button>
        `;
        groupsWrap.append(gEl);

        const cWrap = gEl.querySelector(`[data-clauses="${si}:${gi}"]`);
        (g.clauses || []).forEach((c, ci) => {
          const row = document.createElement('div');
          row.className = 'rp-clause';
          row.innerHTML = `
            <select class="rp-select rp-select--sm" data-c-field="${si}:${gi}:${ci}">
              <option value="category">Category</option>
              <option value="family">Family</option>
              <option value="type">Type</option>
              <option value="name">Name</option>
              <option value="param">Parameter</option>
            </select>
            <select class="rp-select rp-select--sm" data-c-op="${si}:${gi}:${ci}">
              <option value="contains">Contains</option>
              <option value="equals">Equal</option>
              <option value="startswith">StartsWith</option>
              <option value="endswith">EndsWith</option>
              <option value="notcontains">NotContains</option>
              <option value="notequals">NotEqual</option>
            </select>
            <select class="rp-select rp-select--sm rp-param" data-c-param="${si}:${gi}:${ci}"></select>
            <input class="rp-input rp-input--sm" data-c-val="${si}:${gi}:${ci}" value="${esc(c.value || '')}" />
            <button class="rp-x" data-del-clause="${si}:${gi}:${ci}">×</button>
          `;
          cWrap.append(row);

          const fSel = row.querySelector(`[data-c-field="${si}:${gi}:${ci}"]`);
          const opSel = row.querySelector(`[data-c-op="${si}:${gi}:${ci}"]`);
          const pSel = row.querySelector(`[data-c-param="${si}:${gi}:${ci}"]`);
          const vInp = row.querySelector(`[data-c-val="${si}:${gi}:${ci}"]`);

          fSel.value = (c.field || 'category');
          opSel.value = (c.op || 'contains');

          pSel.innerHTML = `<option value="">(param)</option>` + params.map(p => `<option value="${esc(p)}">${esc(p)}</option>`).join('');
          pSel.value = c.param || '';
          pSel.style.display = (fSel.value === 'param') ? '' : 'none';

          const commit = () => {
            updateClause(si, gi, ci, { field: fSel.value, op: opSel.value, param: pSel.value, value: vInp.value });
          };

          fSel.addEventListener('change', () => { pSel.style.display = (fSel.value === 'param') ? '' : 'none'; commit(); });
          opSel.addEventListener('change', commit);
          pSel.addEventListener('change', commit);
          vInp.addEventListener('change', commit);
        });
      });
    });

    setsSlot.onclick = (e) => {
      const delSet = e.target.closest('[data-del-set]');
      if (delSet) { const i = Number(delSet.dataset.delSet); ruleCfg.sets.splice(i, 1); saveRuleConfig(ruleCfg); renderRulePanel(); return; }
      const addG = e.target.closest('[data-add-group]');
      if (addG) { const i = Number(addG.dataset.addGroup); ruleCfg.sets[i].groups.push({ clauses: [{ field:'category', op:'contains', value:'', param:'' }] }); saveRuleConfig(ruleCfg); renderRulePanel(); return; }
      const delG = e.target.closest('[data-del-group]');
      if (delG) { const [i,j] = delG.dataset.delGroup.split(':').map(Number); ruleCfg.sets[i].groups.splice(j,1); saveRuleConfig(ruleCfg); renderRulePanel(); return; }
      const addC = e.target.closest('[data-add-clause]');
      if (addC) { const [i,j] = addC.dataset.addClause.split(':').map(Number); ruleCfg.sets[i].groups[j].clauses.push({ field:'category', op:'contains', value:'', param:'' }); saveRuleConfig(ruleCfg); renderRulePanel(); return; }
      const delC = e.target.closest('[data-del-clause]');
      if (delC) { const [i,j,k] = delC.dataset.delClause.split(':').map(Number); ruleCfg.sets[i].groups[j].clauses.splice(k,1); saveRuleConfig(ruleCfg); renderRulePanel(); return; }
    };

    setsSlot.oninput = (e) => {
      const inp = e.target.closest('input[data-set-name]');
      if (!inp) return;
      const i = Number(inp.dataset.setName);
      ruleCfg.sets[i].name = inp.value;
      saveRuleConfig(ruleCfg);
    };
  }

  // pairs
  const pairsSlot = rulePanelEl.querySelector('[data-slot="pairs"]');
  if (pairsSlot) {
    pairsSlot.innerHTML = '';
    const opts = [`<option value="__ALL__">__ALL__ (전체)</option>`].concat((ruleCfg.sets||[]).map(s => `<option value="${esc(s.id)}">${esc(s.name||s.id)}</option>`));
    (ruleCfg.pairs || []).forEach((p, pi) => {
      const row = document.createElement('div');
      row.className = 'rp-pair';
      row.innerHTML = `
        <select class="rp-select rp-select--sm" data-p-a="${pi}">${opts.join('')}</select>
        <span class="rp-vs">vs</span>
        <select class="rp-select rp-select--sm" data-p-b="${pi}">${opts.join('')}</select>
        <label class="rp-chk"><input type="checkbox" data-p-en="${pi}" ${p.enabled!==false?'checked':''}/> 사용</label>
        <button class="rp-x" data-del-pair="${pi}">×</button>
      `;
      pairsSlot.append(row);
      row.querySelector(`[data-p-a="${pi}"]`).value = p.a || '__ALL__';
      row.querySelector(`[data-p-b="${pi}"]`).value = p.b || '__ALL__';
    });

    pairsSlot.onclick = (e) => {
      const del = e.target.closest('[data-del-pair]');
      if (del) { const i = Number(del.dataset.delPair); ruleCfg.pairs.splice(i,1); saveRuleConfig(ruleCfg); renderRulePanel(); }
    };
    pairsSlot.onchange = (e) => {
      const a = e.target.closest('select[data-p-a]');
      const b = e.target.closest('select[data-p-b]');
      const en = e.target.closest('input[data-p-en]');
      if (a) { const i = Number(a.dataset.pA); ruleCfg.pairs[i].a = a.value; saveRuleConfig(ruleCfg); }
      if (b) { const i = Number(b.dataset.pB); ruleCfg.pairs[i].b = b.value; saveRuleConfig(ruleCfg); }
      if (en) { const i = Number(en.dataset.pEn); ruleCfg.pairs[i].enabled = !!en.checked; saveRuleConfig(ruleCfg); }
    };
  }

  // exclude sets
  const exSlot = rulePanelEl.querySelector('[data-slot="ex"]');
  if (exSlot) {
    exSlot.innerHTML = '';
    const opts = (ruleCfg.sets||[]).map(s => `<option value="${esc(s.id)}">${esc(s.name||s.id)}</option>`);
    (ruleCfg.excludeSetIds || []).forEach((sid, ei) => {
      const row = document.createElement('div');
      row.className = 'rp-exrow';
      row.innerHTML = `
        <select class="rp-select rp-select--sm" data-ex="${ei}">${opts.join('')}</select>
        <button class="rp-x" data-del-ex="${ei}">×</button>
      `;
      exSlot.append(row);
      const sel = row.querySelector(`[data-ex="${ei}"]`);
      if (sel) sel.value = sid || '';
    });

    exSlot.onclick = (e) => {
      const del = e.target.closest('[data-del-ex]');
      if (del) { const i = Number(del.dataset.delEx); ruleCfg.excludeSetIds.splice(i,1); saveRuleConfig(ruleCfg); renderRulePanel(); }
    };
    exSlot.onchange = (e) => {
      const sel = e.target.closest('select[data-ex]');
      if (!sel) return;
      const i = Number(sel.dataset.ex);
      ruleCfg.excludeSetIds[i] = sel.value;
      saveRuleConfig(ruleCfg);
    };
  }
// exclude families
const famPick = rulePanelEl.querySelector('[data-bind="exFamPick"]');
if (famPick) {
  famPick.innerHTML = '';
  const list = famFallback || [];
  if (!list.length) {
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = '목록이 없습니다 (목록 새로고침)';
    famPick.append(opt);
  } else {
    const opt0 = document.createElement('option');
    opt0.value = '';
    opt0.textContent = '선택...';
    famPick.append(opt0);
    list.forEach(n => {
      const o = document.createElement('option');
      o.value = n;
      o.textContent = n;
      famPick.append(o);
    });
  }
}

const exFamSlot = rulePanelEl.querySelector('[data-slot="exfam"]');
if (exFamSlot) {
  exFamSlot.innerHTML = '';
  const cur = Array.isArray(ruleCfg.excludeFamilies) ? ruleCfg.excludeFamilies : [];
  if (!cur.length) {
    const t = document.createElement('div');
    t.className = 'rp-hint';
    t.textContent = '등록된 제외 패밀리가 없습니다.';
    exFamSlot.append(t);
  } else {
    const wrap = document.createElement('div');
    wrap.style.display = 'flex';
    wrap.style.flexWrap = 'wrap';
    wrap.style.gap = '8px';
    cur.forEach(n => {
      const chip = document.createElement('span');
      chip.className = 'control-chip is-active';
      chip.style.cursor = 'pointer';
      chip.title = '클릭: 제거';
      chip.textContent = n;
      chip.addEventListener('click', () => { removeExcludeFamily(n); renderRulePanel(); });
      wrap.append(chip);
    });
    exFamSlot.append(wrap);
  }
}


}

function updateClause(si, gi, ci, next) {
  const s = ruleCfg.sets[si];
  const g = s.groups[gi];
  const c = g.clauses[ci];
  g.clauses[ci] = { ...c, ...next };
  saveRuleConfig(ruleCfg);
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

// Exclude 목록(체크박스) 렌더링
// exclude summary
const exSum = rulePanelEl.querySelector('[data-slot="exsum"]');
if (exSum) {
  const f = Array.isArray(ruleCfg.excludeFamilies) ? ruleCfg.excludeFamilies : [];
  const c = Array.isArray(ruleCfg.excludeCategories) ? ruleCfg.excludeCategories : [];
  exSum.textContent = `패밀리 ${f.length}개, 시스템 ${c.length}개`;
}


// Exclude Sets UI
try {
  const exBox = rulePanelEl.querySelector('[data-slot="ex"]');
  if (exBox) {
    const sets = Array.isArray(ruleCfg.sets) ? ruleCfg.sets : [];
    const ex = Array.isArray(ruleCfg.excludeSetIds) ? ruleCfg.excludeSetIds : [];
    const rows = [];

    for (let i=0; i<ex.length; i++) {
      const sel = ex[i];
      const opts = sets.map((s, idx) => `<option value="${idx}" ${idx===sel?'selected':''}>Set ${idx+1}${s && s.name ? ' - ' + s.name : ''}</option>`).join('');
      rows.push(`
        <div class="rp-card">
          <div class="rp-card-head">
            <div style="font-weight:650;">Exclude Set ${i+1}</div>
            <button class="rp-x" data-act="rm-ex" data-i="${i}" title="삭제">×</button>
          </div>
          <select class="rp-select" data-act="chg-ex" data-i="${i}">${opts}</select>
        </div>
      `);
    }

    exBox.innerHTML = rows.length ? rows.join('') : `<div class="rp-hint">등록된 Exclude Set이 없습니다.</div>`;
  }
} catch {}

}
